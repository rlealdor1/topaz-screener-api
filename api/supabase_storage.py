"""Supabase Storage helper — uses the REST API directly via `requests`.

We avoid the `supabase-py` client because it transitively depends on
`pyiceberg`, which needs a C++ toolchain to build on Python 3.14.
Supabase's Storage REST endpoints are simple enough to hit directly.

Writes use the service-role key (bypasses RLS). Reads go through 1-hour
signed URLs so the Next.js frontend can link to them without further auth.
"""
from __future__ import annotations
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import requests


BUCKET = "research-deliverables"

# MIME types per deliverable extension
_MIME = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".png":  "image/png",
    ".txt":  "text/plain",
    ".pdf":  "application/pdf",
    ".json": "application/json",
}


def _base_url() -> str:
    url = os.environ.get("SUPABASE_URL")
    if not url:
        raise RuntimeError("SUPABASE_URL not set")
    return url.rstrip("/")


def _service_key() -> str:
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not key:
        raise RuntimeError("SUPABASE_SERVICE_ROLE_KEY not set")
    return key


def _content_type(path: Path) -> str:
    return _MIME.get(path.suffix.lower(), "application/octet-stream")


def _upload_file(remote_path: str, local_path: Path) -> None:
    """PUT a single file to the Supabase Storage bucket (upserts)."""
    url = f"{_base_url()}/storage/v1/object/{BUCKET}/{remote_path}"
    headers = {
        "Authorization": f"Bearer {_service_key()}",
        "apikey": _service_key(),
        "Content-Type": _content_type(local_path),
        # x-upsert=true so re-uploads (same path) don't 409
        "x-upsert": "true",
    }
    with open(local_path, "rb") as fh:
        resp = requests.post(url, headers=headers, data=fh.read(), timeout=60)
    if resp.status_code not in (200, 201):
        raise RuntimeError(
            f"Supabase upload failed ({resp.status_code}): {resp.text[:300]}"
        )


def _create_signed_url(remote_path: str, expires_in: int = 3600) -> str:
    """POST /storage/v1/object/sign/{bucket}/{path} returns {signedURL: ...}."""
    url = f"{_base_url()}/storage/v1/object/sign/{BUCKET}/{remote_path}"
    headers = {
        "Authorization": f"Bearer {_service_key()}",
        "apikey": _service_key(),
        "Content-Type": "application/json",
    }
    resp = requests.post(url, headers=headers,
                         json={"expiresIn": expires_in}, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(
            f"Supabase sign failed ({resp.status_code}): {resp.text[:300]}"
        )
    data = resp.json()
    # API returns `signedURL` as a relative path. Two cases observed across
    # Supabase versions:
    #   newer: "/storage/v1/object/sign/<bucket>/<path>?token=..."
    #   older: "/object/sign/<bucket>/<path>?token=..."
    # We normalize both to always include /storage/v1/ in the final URL.
    relative = data.get("signedURL") or data.get("signed_url") or data.get("signedUrl")
    if not relative:
        raise RuntimeError(f"No signedURL in response: {data}")
    if relative.startswith("http"):
        return relative
    if not relative.startswith("/"):
        relative = "/" + relative
    if not relative.startswith("/storage/v1/"):
        relative = "/storage/v1" + relative
    return f"{_base_url()}{relative}"


def upload_deliverables(
    ticker: str,
    local_files: Dict[str, Path],
    ts: Optional[str] = None,
) -> Dict[str, str]:
    """Upload a dict of {kind: local_path} to Supabase Storage.

    Returns a dict of {kind: signed_url} valid for 1 hour.
    Path layout: {TICKER}/{YYYY-MM-DD-HHmmss}/{filename}
    """
    if ts is None:
        ts = datetime.utcnow().strftime("%Y-%m-%d-%H%M%S")
    urls: Dict[str, str] = {}
    for kind, local_path in local_files.items():
        if not local_path.exists():
            continue
        remote_path = f"{ticker.upper()}/{ts}/{local_path.name}"
        _upload_file(remote_path, local_path)
        urls[kind] = _create_signed_url(remote_path)
    return urls
