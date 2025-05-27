# Import required FastAPI and utility modules
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Request, WebSocket
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from datetime import datetime
import asyncio

# Standard libraries for zip file handling and XML parsing
import zipfile
import io
from lxml import etree

# Logging libraries
import logging
import os


#--------------------------------
# Store last uploaded filename and descriptions
#--------------------------------
last_report_data = {
    "filename" : None,
    "descriptions" : None
}


# -------------------------------
# Logging Setup
# -------------------------------
base_dir = Path(__file__).resolve().parent
log_file = "Parser.log"
logging.basicConfig(
    level=logging.INFO,  # Set default log level (INFO)
    format="%(asctime)s [%(levelname)s] %(message)s",  # Format for log entries
    handlers=[
        logging.FileHandler(log_file, mode='a', encoding='utf-8'),  # Logs to file
        logging.StreamHandler()  # Logs to console (optional)
    ]
)
logger = logging.getLogger(__name__)

# -------------------------------
# FastAPI App Setup
# -------------------------------
app = FastAPI()

# Mount the /static path to serve static files (e.g., logo image)
app.mount("/static", StaticFiles(directory="static"), name="static")

# Set the templates directory for rendering HTML pages
templates = Jinja2Templates(directory="templates")

# -------------------------------
# Function: Extract Picture Descriptions from PPTX
# -------------------------------
def extract_picture_descriptions(pptx_bytes):
    slides_output = []

    try:
        # Open the .pptx file as a zip archive in memory
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as pptx_zip:

            # Find and sort all slide XML files (e.g., ppt/slides/slide1.xml)
            slide_files = sorted(
                [f for f in pptx_zip.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')],
                key=lambda x: int(''.join(filter(str.isdigit, x)))  # Sort by slide number
            )

            logger.info(f"Found {len(slide_files)} slide(s) to scan")

            # Loop through each slide file
            for index, slide_file in enumerate(slide_files, start=1):
                logger.debug(f"Parsing slide {index}: {slide_file}")
                slide_descriptions = []

                # Parse the XML for the current slide
                with pptx_zip.open(slide_file) as file:
                    tree = etree.parse(file)

                    # Find all picture elements and get their descriptions
                    for pic in tree.xpath('//p:cNvPr', namespaces={
                        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    }):
                        descr = pic.get('descr')
                        desc = descr if descr else "(No description)"
                        slide_descriptions.append(desc)
                        logger.debug(f"Slide {index} - Picture: {desc}")

                # Append slide info to result list
                slides_output.append({
                    "slide": index,
                    "descriptions": slide_descriptions
                })

        return slides_output

    except Exception as e:
        logger.exception("Error occurred while extracting descriptions")
        raise  # Rethrow exception for upstream handling

# -------------------------------
# Route: Homepage (GET)
# Serves the HTML form for uploading files
# -------------------------------
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    context = {
        "request": request,
        "descriptions": None,
        "error": None,
        "context": {"title": "FastAPI Streaming Log Viewer", "log_file": log_file}
    }
    return templates.TemplateResponse("index.html", context)

# -------------------------------
# Route: Handle Upload (POST)
# Processes the uploaded .pptx file
# -------------------------------
@app.post("/upload-form", response_class=HTMLResponse)
async def upload_form(request: Request, file: UploadFile = File(...)):
    logger.info(f"Received file upload: {file.filename}")

    # Validate file type
    if not file.filename.endswith(".pptx"):
        logger.warning(f"Rejected file (invalid extension): {file.filename}")
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": "Only .pptx files are supported.",
            "descriptions": None
        })

    # Read file contents
    content = await file.read()

    try:
        # Extract picture descriptions
        descriptions = extract_picture_descriptions(content)
        logger.info(f"Extracted picture descriptions from {file.filename}")

        # Save data for report
        last_report_data["filename"] = file.filename
        last_report_data["descriptions"] = descriptions

        # Return results to HTML page
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
    
#-------------------------------------
# Endpoint for the report.txt download
#-------------------------------------

@app.get("/download-report")
def download_report():
    if not last_report_data["descriptions"]:
        return HTMLResponse(content="No report available. Please upload a file first.", status_code=400)

    filename = last_report_data["filename"]
    descriptions = last_report_data["descriptions"]
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Build report content
    report_lines = [f"📄 Report for: {filename}", f"🕒 Generated: {timestamp}", ""]
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


    #------------------------------
    # Functon: Read the last N lines of the fog file
    # used by WebSocket to stream Logs
    #------------------------------
    
async def log_reader(n=5):
    log_lines = []
    with open(f"{base_dir}/{log_file}", "r", encoding="utf-8", errors="replace") as file:
        for line in file.readlines()[-n:]:
            if line.__contains__("ERROR"):
                log_lines.append(f'<span class="text-red-400">{line}</span><br/>')
            elif line.__contains__("WARNING"):
                log_lines.append(f'<span class="text-orange-300">{line}</span><br/>')
            else:
                log_lines.append(f"{line}<br/>")
        return log_lines

#-------------------------------
# WebSocket: /ws/log
# Streams Logs to the frontend in real-time
#-------------------------------

@app.websocket("/ws/log")
async def websocket_endpoint_log(websocket: WebSocket):
    await websocket.accept()

    try:
        while True:
            await asyncio.sleep(1) # Delay to reducde load
            logs = await log_reader(3)
            await websocket.send_text("".join(logs))
    except Exception as e:
        print(e)
    finally:
        await websocket.close()



# -------------------------------
# Allow launching directly via VSCode "Run Python File"
# Checks if the server is running before starting the browser (time limit 10 seconds)
# -------------------------------
import uvicorn
import webbrowser
import threading
import time
import socket

def wait_for_server(host, port, timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.create_connection((host, port), timeout=1):
                return True
        except OSError:
            time.sleep(0.5)
    return False

def open_browser_when_ready():
    if wait_for_server("127.0.0.1", 8000):
        webbrowser.open("http://127.0.0.1:8000")
    else:
        print("Server didn't start in time.")

if __name__ == "__main__":
    threading.Thread(target=open_browser_when_ready).start()
    uvicorn.run(
        "main:app", 
        host="127.0.0.1", 
        port=8000
        )
    
