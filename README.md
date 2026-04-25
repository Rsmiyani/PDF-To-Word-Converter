---
title: PDF To Word Converter
emoji: 📄
colorFrom: blue
colorTo: purple
sdk: docker
pinned: false
---

# 📄 PDF & Word Converter

> A fast, browser-based document conversion tool built with Flask. Convert, merge, and download documents — all in one place, with no sign-up required.

![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.0-lightgrey?logo=flask)
![Docker](https://img.shields.io/badge/Docker-Ready-2496ED?logo=docker&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

---

## ✨ What It Does

| Feature | Description |
|---|---|
| 📄 PDF → Word | Convert PDF files to editable `.docx` |
| 📝 Word → PDF | Convert `.doc` / `.docx` files to PDF |
| 🗂️ Merge PDFs | Combine multiple PDFs into one |
| 📋 Merge Word | Combine multiple `.docx` files into one |

- 🔁 **Batch processing** — upload up to 30 files at once
- 📦 **ZIP download** — get all converted files in a single archive
- 🖱️ **Drag & drop** file upload support
- 📊 **Live status bar** tracking conversions in real time
- 🐳 **Docker-ready** with LibreOffice for Linux/cloud hosting

---

## 🛠️ Tech Stack

`Python` · `Flask` · `pdf2docx` · `docx2pdf` · `pypdf` · `python-docx` · `gunicorn` · `Docker`

---

## 🚀 Run Locally

```powershell
# 1. Create virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the app
python app.py
```

Open → `http://127.0.0.1:5000`

---

## ☁️ Deploy with Docker

```bash
docker build -t pdf-converter .
docker run -e FLASK_SECRET_KEY=your_secret -p 7860:7860 pdf-converter
```

> Includes LibreOffice as fallback for Word → PDF on Linux hosts (e.g. Hugging Face Spaces).

---

## 📌 Notes

- Max upload: **250MB** per request
- Word → PDF uses `docx2pdf` on Windows, falls back to **LibreOffice** on Linux
- Scanned/image-only PDFs may have limited formatting after conversion