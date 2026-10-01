# Agentic Studio

A desktop workspace for analyzing documents and chatting with local Ollama models. Agentic Studio combines document extraction, map-reduce analysis, persistent chat history, optional speech input, and an SSH terminal with SFTP file transfer.

## Features

- Analyze PDF, DOCX, CSV/Excel, text, and image files.
- Create letters and other content directly from Agent Directives without uploading a source file.
- Keep document-focused analysis separate from general chat.
- Select inference and map-reduce models from local/remote Ollama or OpenAI-compatible endpoints.
- Save chat history and application settings in a local SQLite database.
- Use wake-word voice commands, interruptible read-aloud, DOCX/PDF/TXT/Markdown/HTML export, and SSH/SFTP tools.

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
python -m pip install -r requirements.txt
```

Install [Ollama](https://ollama.com/download) separately if using an Ollama server. For local Ollama, download a model, for example:

```powershell
ollama pull llama3.2
```

Install Tesseract OCR separately if image OCR is needed. The application checks the standard Windows install path `C:\Program Files\Tesseract-OCR\tesseract.exe`.

Start the application:

```powershell
python laya_agentic_studio_v15.py
```

## Model Providers and Voice

Open **Models → Provider / Endpoint** to configure local Ollama, remote Ollama, or an OpenAI-compatible local/remote API. Connect to list available models, then choose separate inference and map-reduce models. API keys are stored in the operating system credential vault. **Ollama Setup** can check a configured Ollama endpoint and pull models.

Voice settings include the “Ok Chacha” wake phrase, routing to Agent Directives or Workspace Chat, command execution, speech rate, and female/male read-aloud voice selection. Wake-word and dictation transcription use Google's online speech recognition service when enabled. Read-aloud playback can be stopped from the title bar, voice panel, or by saying “stop speaking.”

With **Output / Export** enabled, the app directly saves its completed response as DOCX, PDF, TXT, Markdown, or HTML.

The application creates `agentic_memory.db` in its working directory to store local history and settings. Keep this file private; it is excluded from Git by default.

## Project Files

- `laya_agentic_studio_v15.py` is the current application entry point.
- `AGENT.md` describes the architecture, workflows, and extension points.
- Older Python versions and local backups are personal copies and are not included in this repository.
