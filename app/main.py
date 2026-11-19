import json
import os
import sys
import threading
import webbrowser
from typing import List, Optional

import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from app.dr2_device_stuck import (
    MAX_SCREENSHOT_BYTES,
    QUICK_COMMENTS,
    generate_device_stuck_mail,
    parse_stuck_rows,
)
from app.logger import setup_logger
from app.outlook_bridge import fetch_mail_by_subject_from_outlook

logger = setup_logger()
app = FastAPI(title="L2 Automation // Created By Adithya")


@app.middleware("http")
async def security_headers_middleware(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self' 'unsafe-inline' 'unsafe-eval' https://cdn.tailwindcss.com data: blob:; "
        "img-src 'self' data: blob:; connect-src 'self';"
    )
    return response


class ParseRequest(BaseModel):
    raw_text: str = Field(default="", max_length=25000)
    soti_env: str = Field(default="Production", max_length=20)


@app.get("/")
def serve_ui():
    ui_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "ui", "index.html"))
    if not os.path.exists(ui_path):
        raise HTTPException(status_code=404, detail="UI file not found.")
    return FileResponse(ui_path, media_type="text/html")


@app.get("/api/dr2/device-stuck/quick-comments")
def get_quick_comments():
    return {"comments": QUICK_COMMENTS}


@app.get("/api/dr2/device-stuck/fetch-mail")
def fetch_mail_api(subject: str = ""):
    return fetch_mail_by_subject_from_outlook(subject_query=subject)


@app.post("/api/dr2/device-stuck/parse")
def parse_rows_api(payload: ParseRequest):
    return {"rows": parse_stuck_rows(payload.raw_text, soti_env=payload.soti_env)}


@app.post("/api/dr2/device-stuck/generate-mail")
async def generate_mail_api(
    sender: str = Form("adithya.rajeev@accenture.com"),
    to_addr: str = Form("Presto.PDS.OpsL2Team@accenture.com"),
    cc_addr: str = Form(""),
    subject: str = Form(""),
    regards_name: str = Form("Adithya"),
    entry_id: str = Form(""),
    store_id: str = Form(""),
    rows_json: str = Form("[]"),
    screenshot_labels: Optional[List[str]] = Form(default=None),
    screenshots: Optional[List[UploadFile]] = File(default=None),
):
    try:
        parsed_rows_raw = json.loads(rows_json or "[]")
        if not isinstance(parsed_rows_raw, list):
            raise ValueError("rows_json must be a list.")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid rows payload: {exc}")

    rows: List[dict] = []
    for item in parsed_rows_raw[:200]:
        if isinstance(item, dict):
            rows.append(
                {
                    "device_name": str(item.get("device_name", "")),
                    "android_device_status": str(item.get("android_device_status", "")),
                    "online_duration": str(item.get("online_duration", "")),
                    "comments": str(item.get("comments", "")),
                }
            )

    raw_shots: List[dict] = []
    if screenshots:
        labels = screenshot_labels or []
        for idx, upload in enumerate(screenshots[:40]):
            content = await upload.read()
            if len(content) > MAX_SCREENSHOT_BYTES:
                raise HTTPException(status_code=413, detail=f"{upload.filename} exceeds 5 MB.")
            label = labels[idx] if idx < len(labels) and labels[idx] else (upload.filename or f"Device_{idx+1}")
            raw_shots.append(
                {
                    "device_label": label,
                    "filename": upload.filename or f"proof_{idx+1}.png",
                    "content": content,
                }
            )

    try:
        result = generate_device_stuck_mail(
            sender=sender,
            to_addr=to_addr,
            cc_addr=cc_addr,
            subject=subject,
            regards_name=regards_name,
            rows=rows,
            entry_id=entry_id,
            store_id=store_id,
            raw_screenshots=raw_shots,
        )
        logger.info(f"Device Stuck mail generated ({result.get('mode')})")
        return JSONResponse(content=result)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as exc:
        logger.error(f"Failed generating Device Stuck mail: {exc}")
        raise HTTPException(status_code=500, detail="Failed to generate mail.")


def open_browser():
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    logger.info("L2 Automation // Created By Adithya — starting on http://127.0.0.1:8000")
    threading.Timer(1.2, open_browser).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")