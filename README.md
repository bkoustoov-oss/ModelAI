# Agentic Studio

A desktop workspace for analyzing documents and chatting with local Ollama models. Agentic Studio combines document extraction, map-reduce analysis, persistent chat history, optional speech input, and an SSH terminal with SFTP file transfer.

## Features

- Analyze PDF, DOCX, CSV/Excel, text, and image files.
- Keep document-focused analysis separate from general chat.
- Select inference and map-reduce models from the Ollama models available locally.
- Save chat history and application settings in a local SQLite database.
- Use optional speech input, document export, and SSH/SFTP tools.

## Requirements

- Python 3.10 or newer
- Ollama installed and running, with at least one compatible model pulled
- Python packages used by the application (including CustomTkinter, Ollama, LARA, Paramiko, and the document/OCR packages listed in [AGENT.md](AGENT.md))
- Tesseract OCR for image text extraction; OCR support is optional

Install the packages in your Python environment, then start the application:

```powershell
python laya_agentic_studio_v15.py
```

The application creates `agentic_memory.db` in its working directory to store local history and settings. Keep this file private; it is excluded from Git by default.

## Project Files

- `laya_agentic_studio_v15.py` is the current application entry point.
- `AGENT.md` describes the architecture, workflows, and extension points.
- Older Python versions and local backups are personal copies and are not included in this repository.
