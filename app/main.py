import os
import sys
import webbrowser
import threading
import uvicorn
from fastapi import FastAPI
from fastapi.responses import FileResponse

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from app.logger import setup_logger

logger = setup_logger()
app = FastAPI(title="L2 Automation // PRESTO Ops Command")

@app.get("/")
def serve_ui():
    ui_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "ui", "index.html"))
    return FileResponse(ui_path)

def open_browser():
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    logger.info("==========================================")
    logger.info("L2 AUTOMATION — System Architect: Adithya")
    logger.info("Starting server on http://127.0.0.1:8000")
    logger.info("==========================================")
    threading.Timer(1.2, open_browser).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")