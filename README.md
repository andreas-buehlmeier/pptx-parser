# pptx-parser

**pptx-parser** is a prototype web application built with **FastAPI** that allows users to upload `.pptx` files and extract metadata by parsing the internal XML structure. Specifically, it targets `p:cNvPr` tags and presents the extracted descriptions in a user-friendly format.

![Firm Logo](static/logo.png)

##  Features

- Upload `.pptx` files through a simple web interface.
- Parse and extract metadata from the internal XML (`p:cNvPr` tags).
- View parsed results directly in the browser.
- Download a detailed report of the results.
- Monitor server activity in real time via WebSocket log stream.

##  Technology Stack

- **Backend**: [FastAPI](https://fastapi.tiangolo.com/)
- **Templating**: Jinja2
- **XML Parsing**: lxml
- **WebSocket**: FastAPI WebSocket
- **Server**: Uvicorn

##  Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-org/pptx-parser.git
   cd pptx-parser
   ```

2. **Install dependencies**
   It’s recommended to use a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install fastapi uvicorn lxml
   ```

3. **Run the application by clicking Run Python File or run command**
   ```bash
   uvicorn main:app --reload
   ```

4. **Open in browser**
>   Note: Once the application starts, it will automatically open in your default web browser.
   The server runs on a randomly assigned free local port (e.g. 'http:127.0.0.1:49215'), which is detected and printed to the terminal.

##  Usage

- Upload a `.pptx` file.
- Click **Process File**.
- View the extracted object data.
- Download the analysis report.
- Watch the log stream in real-time.

>  Note: If WebSocket logs do not appear, check if any anti-tracking software or browser extensions are blocking the connection.

##  Configuration

No special configuration is required beyond installing dependencies.

##  Notes

- Designed as an internal tool/concept demonstration.
- Future improvements may include packaging it as a desktop app or deploying on a cloud platform.
- This program was tested on Python versions 3.12 and 3.13

##  License

This project is proprietary and currently not under any open-source license.

##  About

This project was developed by **Dr. Buhlmeier Consulting Enterprise IT Intelligence** as part of an internal exploration into automated PowerPoint metadata analysis.

## Links

- Youtube Video: [https://youtu.be/NRIaaqDFLOw](https://youtu.be/NRIaaqDFLOw)
- Blog Post: [https://www.buhlmeier.com/blog/](https://www.buhlmeier.com/blog/)