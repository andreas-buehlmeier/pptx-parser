"""Slide Description Extractor for PPTX Files (pptx-parser)

This FastAPI-based web application allows users to upload `.pptx` files
and extract image descriptions embedded in slide XML metadata (`p:cNvPr` tags).

Features:
    - Upload `.pptx` files via a web form
    - Parse slide metadata from XML
    - Display extracted descriptions per slide
    - Download a text-based report of the results
    - Monitor live server logs via WebSocket

Dependencies:
    - fastapi
    - uvicorn
    - lxml
    - Jinja2

Developed by Dr. Buhlmeier Consulting Enterprise IT Intelligence.
"""

from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Request, WebSocket
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import uvicorn
import asyncio
from datetime import datetime
import time
import webbrowser
import threading
import socket
import zipfile
import io
from lxml import etree
import logging

last_report_data = {
    "filename": None,
    "descriptions": None
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

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

def extract_picture_descriptions(pptx_bytes):
    """Extracts image descriptions from a .pptx file's slide XML.

    Args:
        pptx_bytes (bytes): The binary content of the uploaded PPTX file.

    Returns:
        list: A list of dictionaries containing slide numbers and image descriptions.

    Raises:
        Exception: If parsing fails or the .pptx structure is invalid.
    """
    slides_output = []
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as pptx_zip:
            slide_files = sorted(
                [f for f in pptx_zip.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')],
                key=lambda x: int(''.join(filter(str.isdigit, x)))
            )
            logger.info(f"Found {len(slide_files)} slide(s) to scan")

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
    """Serves the homepage with the file upload form.

    Args:
        request (Request): FastAPI request object.

    Returns:
        TemplateResponse: Rendered HTML template.
    """
    return templates.TemplateResponse("index.html", {
        "request": request,
        "descriptions": None,
        "error": None,
        "context": {"title": "FastAPI Streaming Log Viewer", "log_file": log_file}
    })

@app.post("/upload-form", response_class=HTMLResponse)
async def upload_form(request: Request, file: UploadFile = File(...)):
    """Handles .pptx file upload and description extraction.

    Args:
        request (Request): FastAPI request object.
        file (UploadFile): Uploaded .pptx file.

    Returns:
        TemplateResponse: HTML page with extracted descriptions or error.
    """
    logger.info(f"Received file upload: {file.filename}")
    if not file.filename.endswith(".pptx"):
        logger.warning(f"Rejected file (invalid extension): {file.filename}")
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": "Only .pptx files are supported.",
            "descriptions": None
        })

    content = await file.read()
    try:
        descriptions = extract_picture_descriptions(content)
        logger.info(f"Extracted picture descriptions from {file.filename}")
        last_report_data["filename"] = file.filename
        last_report_data["descriptions"] = descriptions
        return templates.TemplateResponse("index.html", {
            "request": request,
            "descriptions": descriptions
        })
    except Exception as e:
        logger.error(f"Failed to parse file {file.filename}: {str(e)}")
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": f"Error processing file: {str(e)}",
            "descriptions": None
        })

@app.get("/download-report")
def download_report():
    """Generates and returns a downloadable report of extracted descriptions.

    Returns:
        StreamingResponse: Text file containing slide-wise descriptions.
    """
    if not last_report_data["descriptions"]:
        return HTMLResponse(content="No report available. Please upload a file first.", status_code=400)

    filename = last_report_data["filename"]
    descriptions = last_report_data["descriptions"]
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    report_lines = [f"\U0001F4C4 Report for: {filename}", f"\U0001F551 Generated: {timestamp}", ""]
    for slide in descriptions:
        report_lines.append(f"Slide {slide['slide']}:")
        for desc in slide['descriptions']:
            report_lines.append(f"  - {desc}")
        report_lines.append("")

    content = "\n".join(report_lines)
    file_like = io.StringIO(content)

    return StreamingResponse(file_like,
                             media_type="text/plain",
                             headers={"Content-Disposition": f"attachment; filename=report_{filename}.txt"})

async def log_reader(n=5):
    """Reads the last N lines of the log file.

    Args:
        n (int): Number of recent lines to read.

    Returns:
        list: A list of HTML-formatted log lines.
    """
    log_lines = []
    with open(f"{base_dir}/{log_file}", "r", encoding="utf-8", errors="replace") as file:
        for line in file.readlines()[-n:]:
            if "ERROR" in line:
                log_lines.append(f'<span class="text-red-400">{line}</span><br/>' )
            elif "WARNING" in line:
                log_lines.append(f'<span class="text-orange-300">{line}</span><br/>' )
            else:
                log_lines.append(f"{line}<br/>")
    return log_lines

@app.websocket("/ws/log")
async def websocket_endpoint_log(websocket: WebSocket):
    """Streams log entries over a WebSocket connection.

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
    finally:
        await websocket.close()

if __name__ == "__main__":
    """Entry point for launching the application with browser auto-open."""
    config = uvicorn.Config(app=app, reload=True)
    server = uvicorn.Server(config=config)
    (sock := socket.socket()).bind(("127.0.0.1", 0))
    thread = threading.Thread(target=server.run, kwargs={"sockets": [sock]})
    thread.start()
    while not server.started:
        time.sleep(0.001)
    address, port = sock.getsockname()
    print(f"HTTP server is now running on http://{address}:{port}")
    webbrowser.open(f'http://{address}:{port}', new=1)
