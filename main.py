"""
PPTX Slide Description, Link Checker & Font Analyzer (pptx-parser)

A FastAPI web application that enables users to upload `.pptx` files and extract:
    - Image descriptions from slide XML metadata
    - All internal and external links, checking for validity
    - Fonts used per slide (with layout/master names)
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional
import asyncio
from datetime import datetime
import time
import webbrowser
import threading
import socket
import io
import logging

import uvicorn
from fastapi import FastAPI, UploadFile, File, Request, WebSocket, Form
from fastapi.responses import HTMLResponse, StreamingResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# OOP extractors
from pptx_parser_ext.extractors.descriptions import DescriptionExtractor
from pptx_parser_ext.extractors.links import LinkCheckExtractor
from pptx_parser_ext.extractors.fonts import FontUsageExtractor

# --- Global State for Last Report Data ---
last_report_data = {
    "filename": None,
    "descriptions": None,
    "links": None,
    "fonts": None,
}

processing_state = {
    "in_progress": False,
    "error": None,
}

base_dir = Path(__file__).resolve().parent
log_file = "Parser.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, mode="a", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# --- FastAPI App and Template Config ---
app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Remember last selected modes so checkboxes stay checked after reload (no globals)
app.state.selected_modes = []


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Serves the homepage with the file upload form."""
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "descriptions": last_report_data.get("descriptions"),
            "links": last_report_data.get("links"),
            "fonts": last_report_data.get("fonts"),
            "error": processing_state.get("error"),
            "processing": processing_state.get("in_progress"),
            "context": {"title": "FastAPI Streaming Log Viewer", "log_file": log_file},
            "selected_mode": app.state.selected_modes,  # keep boxes checked
        },
    )


@app.post("/upload-form", response_class=HTMLResponse)
async def upload_form(
    request: Request, file: UploadFile = File(...), mode: Optional[List[str]] = Form(None)
):
    logger.info(f"Received file upload: {file.filename}")
    if not file.filename.endswith(".pptx"):
        logger.warning(f"Rejected file (invalid extension): {file.filename}")
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": "Only .pptx files are supported.",
                "processing": False,
                "descriptions": None,
                "links": None,
                "fonts": None,
                "selected_mode": app.state.selected_modes,  # keep whatever was last used
            },
        )

    content = await file.read()
    if not mode or len(mode) == 0:
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": "Please select at least one mode before submitting.",
                "descriptions": None,
                "links": None,
                "fonts": None,
                "processing": False,
                "selected_mode": [],  # explicitly none
            },
        )

    # Reset processing state
    processing_state["in_progress"] = True
    processing_state["error"] = None

    # Remember selection so it persists after results load
    app.state.selected_modes = mode[:] if mode else []

    # Instantiate extractors (OOP)
    desc_extractor = DescriptionExtractor()
    link_extractor = LinkCheckExtractor()
    font_extractor = FontUsageExtractor()

    def run_processing():
        try:
            # Extract Descriptions if requested
            if "extract_description" in mode:
                last_report_data["descriptions"] = desc_extractor.extract(content)
            else:
                last_report_data["descriptions"] = None

            # Check Links if requested
            if "check_links" in mode:
                last_report_data["links"] = link_extractor.extract(content)
            else:
                last_report_data["links"] = None

            # Analyze Fonts if requested
            if "analyze_fonts" in mode:
                last_report_data["fonts"] = font_extractor.extract(content)
            else:
                last_report_data["fonts"] = None

            last_report_data["filename"] = file.filename
        except Exception as e:
            logger.error(f"Failed to parse file {file.filename}: {str(e)}")
            processing_state["error"] = f"Error processing file: {str(e)}"
        finally:
            processing_state["in_progress"] = False

    # Start background thread for processing
    threading.Thread(target=run_processing, daemon=True).start()

    # Immediately return the page, showing “processing…” and logs
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "descriptions": None,
            "links": None,
            "fonts": None,
            "processing": True,
            "selected_mode": mode,  # reflect current selection during processing
            "error": None,
        },
    )


@app.post("/reset-ui", response_class=HTMLResponse)
async def reset_ui(request: Request):
    """
    Clear server-side state so the page looks like a fresh load.
    Note: this does NOT cancel an in-progress background parse.
    """
    last_report_data.update({"filename": None, "descriptions": None, "links": None, "fonts": None})
    processing_state["error"] = None
    processing_state["in_progress"] = False
    app.state.selected_modes = []  # so default checkbox state applies
    logger.info("UI state reset by user.")
    # PRG pattern: redirect so a reload won't resubmit the form
    return RedirectResponse(url="/", status_code=303)


@app.get("/status")
def status():
    return {"in_progress": processing_state["in_progress"], "error": processing_state["error"]}


@app.get("/download-report")
def download_report():
    """
    Generates and returns a unified downloadable TXT report of the last operation
    (descriptions, links, fonts — or any combination).
    """
    filename = last_report_data.get("filename") or "presentation"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_lines = [f"\U0001F4C4 Report for: {filename}", f"\U0001F551 Generated: {timestamp}", ""]

    descriptions = last_report_data.get("descriptions")
    links = last_report_data.get("links")
    fonts = last_report_data.get("fonts")
    sections = 0

    # Descriptions
    if descriptions:
        report_lines.append("=== Extracted Descriptions by Slide ===\n")
        for slide in descriptions:
            report_lines.append(f"Slide {slide['slide']}:")
            for desc in slide["descriptions"]:
                report_lines.append(f"  - {desc}")
            report_lines.append("")
        sections += 1

    # Links
    if links:
        if sections:
            report_lines.append("\n")
        report_lines.append("=== Checked Links in PPTX ===\n")
        report_lines.append("Slide | Type           | Status         | Code | Link")
        report_lines.append("------|----------------|----------------|------|-----")
        for link in links:
            line = f"{link['slide']:>5} | {link['type']:<14} | {link['status']:<14} | {str(link['code']):<4} | {link['link']}"
            report_lines.append(line)
        sections += 1

    # Fonts
    if fonts:
        if sections:
            report_lines.append("\n")
        report_lines.append("=== Fonts by Slide ===\n")
        report_lines.append("Slide | Master | Layout | Fonts")
        report_lines.append("----- | ------ | ------ | -----")
        for row in fonts:
            fonts_str = ", ".join(row.get("fonts") or [])
            report_lines.append(
                f"{row.get('slide','—'):>5} | "
                f"{(row.get('master') or '—'):<25} | "
                f"{(row.get('layout') or '—'):<28} | "
                f"{fonts_str}"
            )
        sections += 1

    if not descriptions and not links and not fonts:
        return HTMLResponse(content="No report available. Please upload and process a file first.", status_code=400)

    report_content = "\n".join(report_lines)
    file_like = io.StringIO(report_content)
    report_filename = f"report_{filename}.txt"
    return StreamingResponse(
        file_like,
        media_type="text/plain",
        headers={"Content-Disposition": f"attachment; filename={report_filename}"},
    )


async def log_reader(n: int = 5):
    """Reads the last N lines of the server log for display in the frontend."""
    log_lines = []
    with open(f"{base_dir}/{log_file}", "r", encoding="utf-8", errors="replace") as file:
        for line in file.readlines()[-n:]:
            if "ERROR" in line:
                log_lines.append(f'<span class="text-red-400">{line}</span><br/>')
            elif "WARNING" in line:
                log_lines.append(f'<span class="text-orange-300">{line}</span><br/>')
            else:
                log_lines.append(f"{line}<br/>")
    return log_lines


@app.websocket("/ws/log")
async def websocket_endpoint_log(websocket: WebSocket):
    """Streams server log entries to the frontend over a WebSocket connection."""
    await websocket.accept()
    try:
        while True:
            await asyncio.sleep(0)
            logs = await log_reader(3)
            await websocket.send_text("".join(logs))
    except Exception as e:
        print(e)
    # Do not forcibly close, allow reconnect


if __name__ == "__main__":
    # Dynamically bind to a free local port for development/testing
    sock = socket.socket()
    sock.bind(("127.0.0.1", 0))
    address, port = sock.getsockname()
    print(f"Will start server on http://{address}:{port}")

    # Enable CORS for the correct dynamic origin (for local multi-port flexibility)
    allowed_origin = f"http://{address}:{port}"
    app.add_middleware(
        CORSMiddleware,
        allow_origins=[allowed_origin],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Start the server on the chosen port, open in browser automatically
    config = uvicorn.Config(app=app, reload=True)
    server = uvicorn.Server(config=config)
    thread = threading.Thread(target=server.run, kwargs={"sockets": [sock]})
    thread.start()
    while not server.started:
        time.sleep(0.001)
    print(f"HTTP server is now running on {allowed_origin}")
    webbrowser.open(allowed_origin, new=1)
