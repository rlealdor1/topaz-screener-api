"""Shared-secret header authentication.

The Next.js API routes (deployed on Vercel) proxy requests here and forward
an X-Api-Secret header whose value matches INTERNAL_API_SECRET. This is the
only thing stopping random internet traffic from hitting this backend.
"""
from __future__ import annotations
import os
import secrets
from fastapi import Header, HTTPException, status


def require_secret(x_api_secret: str | None = Header(default=None)) -> None:
    expected = os.environ.get("INTERNAL_API_SECRET")
    if not expected:
        # Fail closed: if the server is misconfigured, refuse all requests.
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Server misconfigured: INTERNAL_API_SECRET not set",
        )
    if x_api_secret is None or not secrets.compare_digest(x_api_secret, expected):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or missing X-Api-Secret header",
        )
