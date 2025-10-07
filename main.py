"""
PPTX/DOCX/XLSX Description, Link Checker & (PPTX) Font Analyzer

Upload an Office file and extract:
  - Image descriptions (where supported)
  - Internal & external links (HTTP checks for external)
  - Fonts per slide (PPTX only, with layout/master names)
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, List
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

# ---- PPTX extractors (existing) ----
from pptx_parser_ext.extractors.descriptions import DescriptionExtractor as PptxDescriptionExtractor
from pptx_parser_ext.extractors.links import LinkCheckExtractor as PptxLinkExtractor
from pptx_parser_ext.extractors.fonts import FontUsageExtractor as PptxFontExtractor

# ---- DOCX/XLSX extractors (optional modules; handled if missing) ----
try:
    from docx_ext.extractors.descriptions import DocxDescriptionExtractor
    from docx_ext.extractors.links import DocxLinkExtractor
    from docx_ext.extractors.fonts import DocxFontExtractor
except Exception as _e:
    DocxDescriptionExtractor = None  # type: ignore
    DocxLinkExtractor = None  # type: ignore
    DocxFontExtractor = None  # type: ignore

try:
    from xlsx_ext.extractors.links import XlsxLinkExtractor
    from xlsx_ext.extractors.descriptions import XlsxDescriptionExtractor
except Exception as _e:
    XlsxLinkExtractor = None  # type: ignore
    XlsxDescriptionExtractor = None  # type: ignore

# -------------------- Global state --------------------

last_report_data = {
    "filename": None,
    "filetype": None,        # "pptx" | "docx" | "xlsx"
    "descriptions": None,    # list
    "links": None,           # list
    "fonts": None,           # list (PPTX only)
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

# -------------------- FastAPI setup --------------------

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Remember last selected modes so checkboxes stay checked after reload
app.state.selected_modes = []

# -------------------- Helpers --------------------

def _ext_to_type(filename: str) -> Optional[str]:
    name_lower = (filename or "").lower()
    if name_lower.endswith(".pptx"):
        return "pptx"
    if name_lower.endswith(".docx"):
        return "docx"
    if name_lower.endswith(".xlsx"):
        return "xlsx"
    return None

# -------------------- Routes --------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Serve the homepage with the file upload form."""
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "descriptions": last_report_data.get("descriptions"),
            "links": last_report_data.get("links"),
            "fonts": last_report_data.get("fonts"),
            "filetype": last_report_data.get("filetype"),  # tell template what to render
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
    fname = file.filename or ""
    logger.info(f"Received file upload: {fname}")
    ftype = _ext_to_type(fname)
    if not ftype:
        logger.warning(f"Rejected file (invalid extension): {fname}")
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": "Only .pptx, .docx, or .xlsx files are supported.",
                "processing": False,
                "descriptions": None,
                "links": None,
                "fonts": None,
                "filetype": None,
                "selected_mode": app.state.selected_modes,
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
                "filetype": None,
                "processing": False,
                "selected_mode": [],
            },
        )

    # Reset processing state & remember mode
    processing_state["in_progress"] = True
    processing_state["error"] = None
    app.state.selected_modes = mode[:] if mode else []

    def run_processing():
        try:
            # Set filename & filetype for the UI and report
            last_report_data["filename"] = fname
            last_report_data["filetype"] = ftype

            # Reset outputs first
            last_report_data["descriptions"] = None
            last_report_data["links"] = None
            last_report_data["fonts"] = None

            # ----- PPTX -----
            if ftype == "pptx":
                if "extract_description" in mode:
                    desc = PptxDescriptionExtractor().extract(content)
                    last_report_data["descriptions"] = desc
                if "check_links" in mode:
                    links = PptxLinkExtractor().extract(content)
                    last_report_data["links"] = links
                if "analyze_fonts" in mode:
                    fonts = PptxFontExtractor().extract(content)
                    last_report_data["fonts"] = fonts

            # ----- DOCX -----
            elif ftype == "docx":
                if "extract_description" in mode and DocxDescriptionExtractor:
                    try:
                        last_report_data["descriptions"] = DocxDescriptionExtractor().extract(content)
                    except Exception as e:
                        logger.warning(f"DOCX descriptions failed: {e}")
                if "check_links" in mode and DocxLinkExtractor:
                    last_report_data["links"] = DocxLinkExtractor().extract(content)
                if "analyze_fonts" in mode:
                    if DocxFontExtractor:
                        try:
                            last_report_data["fonts"] = DocxFontExtractor().extract(content)
                        except Exception as e:
                            logger.warning(f"DOCX fonts failed: {e}")
                    else:
                        logger.info("DocxFontExtractor not available; skipping fonts.")

            # ----- XLSX -----
            else:  # ftype == "xlsx"
                if "extract_description" in mode:
                    if XlsxDescriptionExtractor:
                        try:
                            last_report_data["descriptions"] = XlsxDescriptionExtractor().extract(content)
                        except Exception as e:
                            logger.warning(f"XLSX descriptions failed: {e}")
                    else:
                        logger.info("XlsxDescriptionExtractor not available; skipping descriptions.")

                if "check_links" in mode:
                    if XlsxLinkExtractor:
                        try:
                            links = XlsxLinkExtractor().extract(content)
                            last_report_data["links"] = links
                        except Exception as e:
                            logger.warning(f"XLSX links failed: {e}")
                    else:
                        logger.info("XlsxLinkExtractor not available; skipping links.")

                if "analyze_fonts" in mode:
                    logger.info("Fonts analysis is not implemented for XLSX. Skipping.")

        except Exception as e:
            logger.error(f"Failed to parse file {fname}: {str(e)}")
            processing_state["error"] = f"Error processing file: {str(e)}"
        finally:
            processing_state["in_progress"] = False

    # Start background thread for processing (non-blocking)
    threading.Thread(target=run_processing, daemon=True).start()

    # Immediately return the page, showing “processing…” and logs
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "descriptions": None,
            "links": None,
            "fonts": None,
            "filetype": ftype,
            "processing": True,
            "selected_mode": mode,
            "error": None,
        },
    )


@app.post("/reset-ui", response_class=HTMLResponse)
async def reset_ui(request: Request):
    """
    Clear server-side state so the page looks like a fresh load.
    Note: this does NOT cancel an in-progress background parse.
    """
    last_report_data.update({
        "filename": None,
        "filetype": None,
        "descriptions": None,
        "links": None,
        "fonts": None
    })
    processing_state["error"] = None
    processing_state["in_progress"] = False
    app.state.selected_modes = []  # default checkbox state on next render
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
    for the current file type.
    """
    filename = (last_report_data.get("filename") or "document").replace("\n", " ")
    filetype = last_report_data.get("filetype")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_lines = [
        f"\U0001F4C4 Report for: {filename}",
        f"\U0001F551 Generated: {timestamp}",
        f"Type: {filetype or 'unknown'}",
        ""
    ]

    descriptions = last_report_data.get("descriptions")
    links = last_report_data.get("links")
    fonts = last_report_data.get("fonts")
    sections = 0

    # ----- Descriptions -----
    if descriptions:
        if filetype == "pptx":
            report_lines.append("=== Extracted Descriptions by Slide ===\n")
            for slide in descriptions:
                report_lines.append(f"Slide {slide.get('slide')}:")
                for desc in slide.get("descriptions", []):
                    report_lines.append(f"  - {desc}")
                report_lines.append("")
        elif filetype == "docx":
            report_lines.append("=== Extracted Descriptions by Part ===\n")
            for item in descriptions:
                where = item.get("part") or item.get("section") or item.get("location") or "document"
                report_lines.append(f"Part {where}:")
                for desc in item.get("descriptions", []):
                    report_lines.append(f"  - {desc}")
                report_lines.append("")
        elif filetype == "xlsx":
            report_lines.append("=== Extracted Descriptions ===\n")
            for item in descriptions:
                where = item.get("sheet") or item.get("location") or "workbook"
                report_lines.append(f"{where}:")
                for desc in item.get("descriptions", []):
                    report_lines.append(f"  - {desc}")
                report_lines.append("")
        sections += 1

    # ----- Links -----
    if links:
        if sections:
            report_lines.append("\n")
        if filetype == "pptx":
            head = ("Slide", 5)
        elif filetype == "docx":
            head = ("Part", 6)
        else:
            head = ("Sheet", 6)

        report_lines.append(f"=== Checked Links in {filetype.upper() if filetype else 'FILE'} ===\n")
        report_lines.append(f"{head[0]:<{head[1]}} | Type           | Status         | Code | Link")
        report_lines.append(f"{'-'*head[1]}-|----------------|----------------|------|-----")
        for link in links:
            where = link.get("slide") or link.get("part") or link.get("section") or link.get("sheet") or link.get("location") or "—"
            line = f"{str(where):<{head[1]}} | {link.get('type',''):<14} | {link.get('status',''):<14} | {str(link.get('code','')):<4} | {link.get('link','')}"
            report_lines.append(line)
        sections += 1

    # ----- Fonts -----
    if fonts:
        if sections:
            report_lines.append("\n")
        if filetype == "pptx":
            report_lines.append("=== Fonts by Slide (PPTX) ===\n")
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
        elif filetype == "docx":
            report_lines.append("=== Fonts (DOCX) ===\n")
            for row in fonts:
                fonts_str = ", ".join(row.get("fonts") or [])
                part = row.get("part") or "document"
                report_lines.append(f"{part}: {fonts_str}")
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


# -------------------- Live logs --------------------

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


# -------------------- Dev entrypoint --------------------

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
