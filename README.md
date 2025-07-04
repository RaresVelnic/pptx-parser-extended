# pptx-parser (Enhanced Multi-Feature Version)

**pptx-parser** is an advanced, browser-based PowerPoint analysis tool built with **FastAPI**. It lets you upload `.pptx` files and analyze them for image descriptions and internal/external links—right from a modern web interface.

This program expands the original pptx-parser by offering selectable processing modes (extract descriptions, check links, or both at once), improved error handling, a more responsive UI, and combined reporting.

![Firm Logo](static/logo.png)

---

## Features

- **Flexible Mode Selection**  
  Select “Extract Descriptions”, “Check Links”, or both, before processing each `.pptx` file.

- **User-Friendly Web Interface**  
  Drag & drop `.pptx` files or use the file selector.

- **Image Description Extraction**  
  Parses slide XML (`p:cNvPr` tags) to extract all image descriptions.

- **Comprehensive Link Checker**  
  Validates all hyperlinks and embedded references, displaying status for external and internal links.

- **Downloadable Reports**  
  Download a combined, clearly organized text report with results from all selected analyses.

- **Live Log Stream**  
  Monitor server logs in real-time in your browser.

- **Automatic Port Selection**  
  Runs on a free local port and opens your default browser automatically.

---

## Technology Stack

- **Backend**: [FastAPI](https://fastapi.tiangolo.com/)
- **Templating**: Jinja2
- **XML Parsing**: lxml
- **WebSocket Logging**: FastAPI WebSocket
- **Server**: Uvicorn

---

## Installation

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

3. **Run the application**
   ```bash
   python main.py
   ```

4. **Open in browser**  
   The application will open automatically in your default browser on a random free local port.

---

## Usage

1. **Select at least one mode** (“Extract Descriptions”, “Check Links”, or both).
2. **Upload a `.pptx` file** (drag & drop or use the file dialog).
3. **Click “Process File”**.
4. **View results**: Descriptions and/or checked links are displayed separately and clearly.
5. **Download report**: Click the button to get a combined text report.
6. **Live logs**: Watch server logs in real time at the bottom of the page.

> If you submit without any mode selected, you’ll be prompted to select at least one.

---

## Configuration

No special configuration required.  
Tested on Python 3.12+.

---

## Notes

- Designed for internal/local use only (no user authentication or persistent file storage).
- All uploaded files are processed in-memory and not saved to disk.
- Future improvements may include packaging as a desktop app or enabling batch processing.

---

## License

[License](license.txt)

---

## About

This project was developed by **Dr. Buhlmeier Consulting Enterprise IT Intelligence** as part of an internal exploration into automated PowerPoint metadata and link analysis.

---

## Links

- Youtube Video: [https://youtu.be/NRIaaqDFLOw](https://youtu.be/NRIaaqDFLOw)
- Blog Post: [https://www.buhlmeier.com/blog/](https://www.buhlmeier.com/blog/)

---
