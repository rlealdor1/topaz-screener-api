"""FastAPI entrypoint — Topaz Screener API.

Run locally:
    cd "5.0 Claude Screening"
    uvicorn api.main:app --reload --port 8000

Production (Docker):
    CMD ["uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8000"]
"""
from __future__ import annotations
import os
import re
from uuid import uuid4

from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Path, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from .auth import require_secret
from .jobs import JOBS, run_job


app = FastAPI(
    title="Topaz Screener API",
    description="Generates 5 equity-research deliverables per US stock ticker.",
    version="1.0.0",
)

# CORS: the Vercel frontend + local dev. Override with env var if needed.
_default_origins = [
    "https://internal-platform-pearl.vercel.app",
    "http://localhost:3000",
    "http://localhost:3001",
]
_extra = os.environ.get("CORS_EXTRA_ORIGINS", "")
origins = _default_origins + [o.strip() for o in _extra.split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type", "X-Api-Secret"],
    allow_credentials=False,
)


# ---------- Schemas ----------

class JobRequest(BaseModel):
    ticker: str


class JobCreatedResponse(BaseModel):
    job_id: str
    status: str
    ticker: str


# ---------- Routes ----------

@app.get("/health")
def health() -> dict:
    """Liveness probe — no auth required."""
    return {"status": "ok", "service": "topaz-screener-api"}


@app.post("/jobs", status_code=status.HTTP_201_CREATED,
          response_model=JobCreatedResponse)
def create_job(
    body: JobRequest,
    tasks: BackgroundTasks,
    _auth: None = Depends(require_secret),
) -> JobCreatedResponse:
    """Start a new research job. Returns immediately with a job_id."""
    ticker = body.ticker.upper().strip()
    if not re.fullmatch(r"[A-Z]{1,5}", ticker):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Ticker must be 1-5 uppercase letters",
        )

    job_id = str(uuid4())
    JOBS[job_id] = {
        "job_id": job_id,
        "ticker": ticker,
        "status": "queued",
        "step": "Queued",
    }
    tasks.add_task(run_job, job_id, ticker)
    return JobCreatedResponse(job_id=job_id, status="queued", ticker=ticker)


@app.get("/jobs/{job_id}")
def get_job(
    job_id: str = Path(..., min_length=32, max_length=36),
    _auth: None = Depends(require_secret),
) -> dict:
    """Poll job status. Returns the full job dict including URLs when complete."""
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Job {job_id} not found",
        )
    # Don't expose the full traceback to the frontend — keep a summary.
    return {k: v for k, v in job.items() if k not in ("traceback",)}


@app.get("/jobs")
def list_jobs(
    _auth: None = Depends(require_secret),
    limit: int = 20,
) -> list:
    """Admin: list recent jobs (most recent last). Useful for debugging."""
    # Sort by started_at if available, else by insertion
    items = list(JOBS.values())
    items.sort(key=lambda j: j.get("started_at") or "", reverse=True)
    return items[:limit]
