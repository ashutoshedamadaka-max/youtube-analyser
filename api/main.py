import os
import uuid
import threading
import subprocess
import sys
from pathlib import Path

from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

app = FastAPI(title="YouTube Analyser API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

jobs: dict = {}
_lock = threading.Lock()
_busy = False


class AnalyzeRequest(BaseModel):
    topic: str


def run_pipeline(job_id: str, topic: str):
    global _busy
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}

    steps = [
        ("Fetching videos",      [sys.executable, "tools/fetch_videos.py", "--topic", topic]),
        ("Fetching channels",    [sys.executable, "tools/fetch_channels.py"]),
        ("Fetching transcripts", [sys.executable, "tools/fetch_transcripts.py"]),
        ("Fetching thumbnails",  [sys.executable, "tools/fetch_thumbnails.py"]),
        ("Analysing content",    [sys.executable, "tools/analyze_content.py"]),
        ("Generating charts",    [sys.executable, "tools/generate_charts.py"]),
        ("Generating slides",    ["node", "tools/generate_slides.js"]),
    ]

    try:
        for msg, cmd in steps:
            jobs[job_id]["message"] = msg
            result = subprocess.run(cmd, capture_output=True, text=True, env=env)
            if result.returncode != 0:
                raise Exception(f"{msg} failed:\n{result.stderr[-1000:]}")

        output_files = sorted(
            Path("output").glob("*.pptx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not output_files:
            raise Exception("Pipeline completed but no PPTX was generated")

        jobs[job_id]["status"] = "complete"
        jobs[job_id]["message"] = "Done"
        jobs[job_id]["file"] = str(output_files[0])

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["message"] = str(e)

    finally:
        _busy = False


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/analyze")
async def analyze(req: AnalyzeRequest, background_tasks: BackgroundTasks):
    global _busy
    with _lock:
        if _busy:
            raise HTTPException(
                status_code=429,
                detail="A job is already running. Please wait and try again.",
            )
        _busy = True

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "running", "message": "Starting...", "file": None}
    background_tasks.add_task(run_pipeline, job_id, req.topic)
    return {"job_id": job_id}


@app.get("/status/{job_id}")
async def status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return job


@app.get("/download/{job_id}")
async def download(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job["status"] != "complete":
        raise HTTPException(status_code=400, detail="Report not ready yet")
    return FileResponse(
        job["file"],
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=Path(job["file"]).name,
    )
