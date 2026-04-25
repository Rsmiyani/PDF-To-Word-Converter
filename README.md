# PDF and Word Batch Converter

This project is a Flask web app that converts and merges documents from a modern UI dashboard.

It supports multiple actions:

- PDF to Word (.docx)
- Word (.doc/.docx) to PDF
- Merge multiple PDFs into a single PDF
- Merge multiple Word files (.docx) into a single .docx

## Features

- Upload and convert up to 30 files at once (supports 15+ as requested).
- Batch PDF to Word conversion.
- Batch Word to PDF conversion.
- PDF merge (2+ files).
- Word merge for .docx files (2+ files).
- Downloads converted files as a ZIP archive.
- Downloads merged output as a single file.
- Adds a `conversion_report.txt` in the ZIP if any file fails conversion.
- Live status bar showing cumulative conversion and merge activity.
- Tool filters (All / Conversion / Merge) and drag-drop uploads.
- Production-ready setup with health endpoint and deployment files.

## Tech Stack

- Python
- Flask
- pdf2docx
- docx2pdf
- pypdf
- python-docx
- docxcompose
- gunicorn

## Run Locally

1. Create and activate a virtual environment:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. Install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. Start the app:

   ```powershell
   python app.py
   ```

4. Open in browser:

   ```
   http://127.0.0.1:5000
   ```

## Deploy / Host

### Option 1: Procfile Hosts (Render, Railway, Heroku-style)

1. Set environment variables:

   - `FLASK_SECRET_KEY` (required in production)
   - `PORT` (usually provided by host)
   - Optional: `MAX_FILES_PER_REQUEST`, `MAX_UPLOAD_SIZE_MB`, `SESSION_COOKIE_SECURE`

2. Use the included `Procfile`:

   ```
   web: gunicorn wsgi:application --bind 0.0.0.0:${PORT:-8000} --workers 2 --timeout 300
   ```

3. Health check endpoint:

   ```
   /healthz
   ```

### Option 2: Docker Hosts

This repository includes a `Dockerfile` that installs LibreOffice for hosted Word-to-PDF fallback.

Build and run:

```powershell
docker build -t doc-converter .
docker run -e FLASK_SECRET_KEY=replace_this -e PORT=8000 -p 8000:8000 doc-converter
```

Open:

```
http://127.0.0.1:8000
```

## Notes

- Upload limit is 250MB total per request.
- PDF to Word accepts `.pdf` files.
- Word to PDF accepts `.doc` and `.docx` files.
- Word to PDF uses `docx2pdf` first. If unavailable on your host, it falls back to LibreOffice (`soffice`) when installed.
- Word merge accepts `.docx` files only.
- On Linux hosts, `pywin32` is skipped automatically through environment markers in `requirements.txt`.
- Some complex PDFs (scanned/image-only PDFs) may convert with limited formatting depending on source quality.