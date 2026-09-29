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
- Tesseract OCR installed separately for image text extraction; OCR support is optional

## Install

From PowerShell in the project directory, create an environment and install the Python dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install customtkinter python-docx laya ollama Pillow pytesseract pdfplumber pandas openpyxl SpeechRecognition paramiko
```

Optional microphone input requires PyAudio:

```powershell
python -m pip install PyAudio
```

Install [Ollama](https://ollama.com/download) separately, then download a model, for example:

```powershell
ollama pull llama3.2
```

Install Tesseract OCR separately if image OCR is needed. The application checks the standard Windows install path `C:\Program Files\Tesseract-OCR\tesseract.exe`.

Start the application:

```powershell
python laya_agentic_studio_v15.py
```

## Ollama Setup

In the app, open **Models → Ollama Setup** (or select **Ollama Setup** beside the model refresh button). Use **Check & Refresh Models** to connect to the local Ollama service, or enter a local model name and select **Pull** to download it. New users can select **Create Ollama Account**; for cloud models, select **Ollama Cloud Sign-in**, complete Ollama's sign-in in the new terminal window, then refresh the model list. Agentic Studio does not create a separate user account or store Ollama credentials.

Ollama-hosted `gpt-oss` models are not the same service as OpenAI's hosted GPT models. This app currently connects to Ollama; it does not configure an OpenAI API account.

The application creates `agentic_memory.db` in its working directory to store local history and settings. Keep this file private; it is excluded from Git by default.

## Project Files

- `laya_agentic_studio_v15.py` is the current application entry point.
- `AGENT.md` describes the architecture, workflows, and extension points.
- Older Python versions and local backups are personal copies and are not included in this repository.
