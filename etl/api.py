#!/usr/bin/env python3
"""
ETL upload API for Jutem Fund (company 22).

Wraps etl/monthly_update.py behind an authenticated HTTP endpoint so the
Jutem team can upload their Excel workbook from the React UI instead of
running Python on the server.

Run locally:
    pip install fastapi uvicorn python-multipart openpyxl
    export ETL_API_TOKEN='shared-secret'
    export DIRECTUS_TOKEN='directus-admin-token'
    uvicorn etl.api:app --host 127.0.0.1 --port 8001

Endpoints:
    GET  /jutem-upload/health
    POST /jutem-upload/sheets         - list sheet names in an uploaded workbook
    POST /jutem-upload/monthly-update - run monthly_update.py against an uploaded workbook
"""

import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

import openpyxl
from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

ETL_DIR = Path(__file__).resolve().parent
SCRIPT = ETL_DIR / "monthly_update.py"
ALLOWED_STEPS = {"loans", "interest", "expenses", "all"}

# Same shared secret the rest of the stack uses (set via env var on the server).
SHARED_SECRET = os.environ.get("ETL_API_TOKEN")
DIRECTUS_TOKEN = os.environ.get("DIRECTUS_TOKEN")

# Origins that may call this API (the React frontend).
ALLOWED_ORIGINS = [
    o.strip()
    for o in os.environ.get(
        "ETL_API_ORIGINS",
        "https://zedlenders.pickmesms.com,http://localhost:3000",
    ).split(",")
    if o.strip()
]

app = FastAPI(title="ZedLenders ETL API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["Authorization", "Content-Type"],
)


def _check_auth(authorization: str | None):
    if not SHARED_SECRET:
        raise HTTPException(500, "ETL_API_TOKEN is not configured on the server")
    expected = f"Bearer {SHARED_SECRET}"
    if authorization != expected:
        raise HTTPException(401, "unauthorized")


def _save_upload(file: UploadFile) -> Path:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "file must be an .xlsx workbook")
    job_id = uuid.uuid4().hex[:8]
    target = Path(tempfile.gettempdir()) / f"jutem-upload-{job_id}.xlsx"
    with target.open("wb") as f:
        shutil.copyfileobj(file.file, f)
    return target


@app.get("/jutem-upload/health")
def health():
    return {"ok": True, "script": str(SCRIPT)}


@app.post("/jutem-upload/sheets")
async def list_sheets(
    file: UploadFile = File(...),
    authorization: str | None = Header(None),
):
    """Return the list of sheet names in an uploaded workbook so the UI can offer
    a dropdown rather than asking the user to type 'APRIL 2026' verbatim."""
    _check_auth(authorization)
    workbook = _save_upload(file)
    try:
        wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
        return {"sheets": list(wb.sheetnames)}
    finally:
        workbook.unlink(missing_ok=True)


@app.post("/jutem-upload/monthly-update")
async def monthly_update(
    file: UploadFile = File(...),
    sheet: str = Form(...),
    execute: bool = Form(False),
    step: str = Form("all"),
    authorization: str | None = Header(None),
):
    """Run etl/monthly_update.py against an uploaded workbook. Streams the
    script's stdout/stderr back to the client line by line."""
    _check_auth(authorization)

    if step not in ALLOWED_STEPS:
        raise HTTPException(400, f"step must be one of {sorted(ALLOWED_STEPS)}")
    if not DIRECTUS_TOKEN:
        raise HTTPException(500, "DIRECTUS_TOKEN is not configured on the server")

    workbook = _save_upload(file)

    cmd = [
        sys.executable,
        str(SCRIPT),
        "--sheet", sheet,
        "--workbook", str(workbook),
    ]
    if execute:
        cmd.append("--execute")
    if step != "all":
        cmd.extend(["--step", step])

    env = os.environ.copy()
    env["DIRECTUS_TOKEN"] = DIRECTUS_TOKEN
    env["PYTHONUNBUFFERED"] = "1"

    def stream():
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=env,
            cwd=str(ETL_DIR.parent),
        )
        try:
            assert proc.stdout is not None
            for line in proc.stdout:
                yield line
            proc.wait()
            yield f"\n[exit {proc.returncode}]\n"
        finally:
            workbook.unlink(missing_ok=True)

    return StreamingResponse(stream(), media_type="text/plain")
