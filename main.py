"""
PPTX Slide Description & Link Checker (pptx-parser)

A FastAPI web application that enables users to upload `.pptx` files and extract:
    - Image descriptions from slide XML metadata
    - All internal and external links, checking for validity

Features:
    - Upload PowerPoint (.pptx) files via web form (with drag & drop support)
    - Select one or both features: "Extract Descriptions" and/or "Check Links"
    - Results are displayed by section on the results page
    - Download a unified report (TXT) containing all results
    - Live server log view via WebSocket (for troubleshooting)
    - CORS support with dynamic local port

Dependencies:
    - fastapi
    - uvicorn
    - lxml
    - Jinja2
    - requests

Developed by Dr. Buhlmeier Consulting Enterprise IT Intelligence.
"""

from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Request, WebSocket, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional
import uvicorn
import asyncio
from datetime import datetime
import time
import webbrowser
import threading
import socket
import zipfile
import io
import requests
from lxml import etree
import logging
import posixpath

# --- Global State for Last Report Data ---
last_report_data = {
    "filename": None,
    "descriptions": None,
    "links": None,
}

base_dir = Path(__file__).resolve().parent
log_file = "Parser.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, mode='a', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- FastAPI App and Template Config ---
app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


def extract_picture_descriptions(pptx_bytes):
    """
    Extracts image descriptions from a .pptx file's slide XML (p:cNvPr tags).

    Args:
        pptx_bytes (bytes): Binary content of uploaded PPTX file.

    Returns:
        list: List of dicts, each with slide number and extracted image descriptions.
    Raises:
        Exception: If parsing fails or pptx structure is invalid.
    """
    slides_output = []
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as pptx_zip:
            slide_files = sorted(
                [f for f in pptx_zip.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')],
                key=lambda x: int(''.join(filter(str.isdigit, x)))
            )
            logger.info(f"Found {len(slide_files)} slide(s) to scan for descriptions")

            for index, slide_file in enumerate(slide_files, start=1):
                slide_descriptions = []
                with pptx_zip.open(slide_file) as file:
                    tree = etree.parse(file)
                    for pic in tree.xpath('//p:cNvPr', namespaces={
                        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    }):
                        descr = pic.get('descr')
                        desc = descr if descr else "(No description)"
                        slide_descriptions.append(desc)

                slides_output.append({
                    "slide": index,
                    "descriptions": slide_descriptions
                })
        return slides_output

    except Exception as e:
        logger.exception("Error occurred while extracting descriptions")
        raise

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """
    Serves the homepage with the file upload form.

    Args:
        request (Request): FastAPI request object.

    Returns:
        TemplateResponse: Rendered HTML template for index page.
    """
    return templates.TemplateResponse("index.html", {
        "request": request,
        "descriptions": None,
        "links": None,
        "error": None,
        "context": {"title": "FastAPI Streaming Log Viewer", "log_file": log_file}
    })

@app.post("/upload-form", response_class=HTMLResponse)
async def upload_form(request: Request, file: UploadFile = File(...), mode: Optional[List[str]] = Form(None)):
    """
    Handles .pptx file upload and extraction of descriptions/links
    according to the selected mode(s) (checkboxes).

    Args:
        request (Request): FastAPI request object.
        file (UploadFile): Uploaded .pptx file.
        mode (List[str]): List of selected features ("extract_description", "check_links")

    Returns:
        TemplateResponse: HTML page with extracted descriptions and/or checked links.
    """
    logger.info(f"Received file upload: {file.filename}")
    if not file.filename.endswith(".pptx"):
        logger.warning(f"Rejected file (invalid extension): {file.filename}")
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": "Only .pptx files are supported.",
            "descriptions": None,
            "links": None,
        })

    content = await file.read()
    if not mode or len(mode) == 0:
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": "Please select at least one mode before submitting.",
                "descriptions": None,
                "links": None,
                "selected_mode": [],
            }
    )
    descriptions = links = None
    try:
        # Extract Descriptions if requested
        if "extract_description" in mode:
            descriptions = extract_picture_descriptions(content)
            last_report_data["descriptions"] = descriptions
        else:
            last_report_data["descriptions"] = None

        # Check Links if requested
        if "check_links" in mode:
            links = check_links_in_pptx(content)
            last_report_data["links"] = links
        else:
            last_report_data["links"] = None

        last_report_data["filename"] = file.filename

        return templates.TemplateResponse("index.html", {
            "request": request,
            "descriptions": descriptions,
            "links": links,
            "selected_mode": mode
        })
    except Exception as e:
        logger.error(f"Failed to parse file {file.filename}: {str(e)}")
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": f"Error processing file: {str(e)}",
            "descriptions": None,
            "links": None,
            "selected_mode": mode
        })


def get_relationships(z, rels_path):
    """
    Extracts relationships from a .rels XML file within a pptx archive.

    Args:
        z (zipfile.ZipFile): Open pptx archive.
        rels_path (str): Path to the relationships XML file.

    Returns:
        dict: Mapping from relationship ID to dict with 'target', 'type', 'target_mode'
    """
    rels = {}
    if rels_path in z.namelist():
        tree = etree.fromstring(z.read(rels_path))
        for rel in tree.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rel_id = rel.get("Id")
            target = rel.get("Target")
            rel_type = rel.get("Type")
            target_mode = rel.get("TargetMode")  # "External" or "Internal" (or None)
            rels[rel_id] = {"target": target, "type": rel_type, "target_mode": target_mode}
    return rels

def http_code_meaning(code):
    """
    Converts HTTP status code to a human-readable meaning.
    """
    if code is None:
        return "No response"
    try:
        code = int(code)
    except Exception:
        return "Unknown"
    if 200 <= code < 300:
        return "OK"
    elif 300 <= code < 400:
        return "Redirect"
    elif 400 <= code < 500:
        return "Client Error"
    elif 500 <= code < 600:
        return "Server Error"
    else:
        return "Other"

def check_url(url):
    """
    Checks the status of an external (HTTP) URL.

    Returns:
        tuple: (status_code, status_text, reason/description)
    """
    try:
        resp = requests.head(url, allow_redirects=True, timeout=5)
        return resp.status_code, http_code_meaning(resp.status_code), resp.reason
    except Exception as e:
        return None, "Bad link", str(e)

def check_links_in_pptx(pptx_bytes):
    """
    Extracts and checks all links in a PPTX file.

    - Checks external HTTP(S) links (status code, reachable).
    - Checks internal links (to other slides, images, embedded files) by verifying the target file exists in the archive.

    Args:
        pptx_bytes (bytes): Binary PPTX file content.

    Returns:
        list: List of dicts, each describing a found link and its status.
    """
    results = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        all_files = set(z.namelist())
        slide_files = sorted(
            [n for n in all_files if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
            key=lambda x: int(''.join(filter(str.isdigit, x)))
        )

        for slide_idx, slide_file in enumerate(slide_files, start=1):
            rels_path = slide_file.replace("slides/", "slides/_rels/") + ".rels"
            rels = get_relationships(z, rels_path)
            tree = etree.fromstring(z.read(slide_file))
            for elem in tree.iter():
                r_id = elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if r_id and r_id in rels:
                    rel = rels[r_id]
                    target = rel['target']
                    rel_type = rel['type']
                    target_mode = rel.get('target_mode', None)
                    # Check for external web links
                    if (target_mode == "External" and target.startswith("http")) or target.startswith("http"):
                        result = {
                            "slide": slide_idx,
                            "link": target,
                            "type": "External",
                        }
                        code, status, desc = check_url(target)
                        result.update({
                            "status": status,
                            "code": code,
                            "description": desc,
                        })
                        results.append(result)
                    else:
                        # Check internal file references by normalizing the path
                        slide_dir = posixpath.dirname(slide_file)
                        normalized_path = posixpath.normpath(posixpath.join(slide_dir, target))
                        exists = normalized_path in all_files
                        # Categorize type for reporting
                        if "/slides/" in normalized_path:
                            link_type = "Internal Slide"
                        elif "/media/" in normalized_path or "/embeddings/" in normalized_path:
                            link_type = "Internal File"
                        else:
                            link_type = "Internal"

                        result = {
                            "slide": slide_idx,
                            "link": target,
                            "type": link_type,
                        }
                        if exists:
                            result.update({
                                "status": "OK",
                                "code": "",
                                "description": f"Target exists: {normalized_path}"
                            })
                        else:
                            result.update({
                                "status": "Broken/Missing",
                                "code": "",
                                "description": f"Target missing: {normalized_path}"
                            })
                        results.append(result)
    return results


@app.get("/download-report")
def download_report():
    """
    Generates and returns a unified downloadable TXT report of the last operation (descriptions, links, or both).

    Returns:
        StreamingResponse: Text file containing the full report.
    """
    filename = last_report_data.get("filename") or "presentation"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_lines = [
        f"\U0001F4C4 Report for: {filename}",
        f"\U0001F551 Generated: {timestamp}",
        ""
    ]

    descriptions = last_report_data.get("descriptions")
    links = last_report_data.get("links")
    sections = 0

    # Add descriptions section if present
    if descriptions:
        report_lines.append("=== Extracted Descriptions by Slide ===\n")
        for slide in descriptions:
            report_lines.append(f"Slide {slide['slide']}:")
            for desc in slide['descriptions']:
                report_lines.append(f"  - {desc}")
            report_lines.append("")
        sections += 1

    # Add links section if present
    if links:
        if sections:
            report_lines.append("\n")  # Extra space if both
        report_lines.append("=== Checked Links in PPTX ===\n")
        report_lines.append("Slide | Type           | Status         | Code | Link")
        report_lines.append("------|----------------|----------------|------|-----")
        for link in links:
            line = f"{link['slide']:>5} | {link['type']:<14} | {link['status']:<14} | {str(link['code']):<4} | {link['link']}"
            report_lines.append(line)

    if not descriptions and not links:
        # No data found
        return HTMLResponse(content="No report available. Please upload and process a file first.", status_code=400)

    report_content = "\n".join(report_lines)
    file_like = io.StringIO(report_content)
    report_filename = f"report_{filename}.txt"
    return StreamingResponse(file_like,
                             media_type="text/plain",
                             headers={"Content-Disposition": f"attachment; filename={report_filename}"})

async def log_reader(n=5):
    """
    Reads the last N lines of the server log for display in the frontend.

    Args:
        n (int): Number of lines.

    Returns:
        list: List of formatted log lines as HTML strings.
    """
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
    """
    Streams server log entries to the frontend over a WebSocket connection.

    Args:
        websocket (WebSocket): WebSocket connection to the client.
    """
    await websocket.accept()
    try:
        while True:
            await asyncio.sleep(1)
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
