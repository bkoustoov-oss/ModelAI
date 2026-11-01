# Agentic Studio

A desktop workspace for analyzing documents and chatting with local Ollama models. Agentic Studio combines document extraction, map-reduce analysis, persistent chat history, optional speech input, and an SSH terminal with SFTP file transfer.

## Features

- Analyze PDF, DOCX, CSV/Excel, text, and image files.
- Create letters and other content directly from Agent Directives without uploading a source file.
- Keep document-focused analysis separate from general chat.
- Select inference and map-reduce models from local/remote Ollama or OpenAI-compatible endpoints.
- Save chat history and application settings in a local SQLite database.
- Use wake-word voice commands, interruptible read-aloud, DOCX/PDF/TXT/Markdown/HTML export, and SSH/SFTP tools.

## Requirements and installation

LAYA can connect to a local or remote Ollama server or an OpenAI-compatible
endpoint. Model servers and model weights are not bundled with the desktop
application. Install Ollama only if you intend to use an Ollama server.
Tesseract OCR is optional and is needed only for OCR of images or scanned
documents. Microphone dictation and the wake phrase use Google's online speech
recognition service when enabled, so those voice features require internet
access.

### Install a packaged desktop build

Download the artifact for your operating system from the repository's
[Releases page](https://github.com/bkoustoov-oss/ModelAI/releases). Installers
are built per operating system; do not try to run one OS's package on another.

#### Windows

1. Download `LAYA-Windows-Setup.exe` and run it. The installer is per-user and
   should not require administrator privileges.
2. Launch **LAYA Agentic Studio** from the Start menu.
3. If using a portable build, download `LAYA-Windows-x64.zip`, extract the
   entire archive to a writable folder, and run `LAYAAgenticStudio.exe` from
   the extracted `LAYAAgenticStudio` folder. Do not copy the EXE out of its
   companion `_internal` folder.
4. Optionally install [Ollama for Windows](https://ollama.com/download) and
   pull a model in PowerShell, for example `ollama pull llama3.2`.
5. Optionally install Tesseract OCR for image/scanned-document extraction.
   The app checks the standard install path
   `C:\Program Files\Tesseract-OCR\tesseract.exe`.

#### macOS

1. Download the `LAYA-macOS-*.dmg` artifact that matches your Mac's
   architecture (`arm64` for Apple silicon; `x64` for Intel).
2. Open the DMG and run the included `install_macos.command`. It installs the
   application in `~/Applications` and launches it.
3. The current packages are not Developer ID signed or notarized. Do not
   bypass macOS security controls for an untrusted build; use a signed and
   notarized release when one is available or run from source.
4. Optionally install [Ollama for macOS](https://ollama.com/download), start
   its server, and pull a model using `ollama pull llama3.2`.
5. Install Tesseract separately only if OCR is needed.

#### Linux

1. Download `LAYA-Linux-x64.tar.gz` and extract it, for example:

   ```bash
   mkdir -p ~/Downloads/laya-release
   tar -xzf LAYA-Linux-x64.tar.gz -C ~/Downloads/laya-release
   cd ~/Downloads/laya-release
   ```

2. Install for the current user:

   ```bash
   bash install_linux.sh
   ```

3. Launch with `~/.local/bin/laya-agentic-studio`, or use the desktop entry.
   Ensure `~/.local/bin` is on your `PATH` if launching from a terminal.
4. Optionally install [Ollama for Linux](https://ollama.com/download), start
   the server, and pull a model using `ollama pull llama3.2`.
5. Install Tesseract using your distribution's package manager only if OCR is
   needed.

Linux packages are built on Ubuntu 24.04 and may require compatible system
libraries on older distributions. The release is x64; ARM64 Linux is not
currently published.

### Run from source (Windows, macOS, or Linux)

Prerequisites:

- Python 3.10 or newer; Python 3.13 is used by release CI.
- Tcl/Tk support in the Python installation (required by the GUI).
- Native audio/PortAudio development libraries if installing PyAudio fails
  or microphone input is needed. On macOS, install with `brew install
  portaudio`; on Ubuntu/Debian, install `portaudio19-dev` and `python3-tk`.
- Git, if cloning the source repository.

Clone the repository and create an isolated environment:

```bash
git clone https://github.com/bkoustoov-oss/ModelAI.git
cd ModelAI
python -m venv .venv
```

Activate the environment and install dependencies:

```powershell
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

```bash
# macOS or Linux
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Run the app from the repository root:

```bash
python laya_agentic_studio_v15.py
```

If PowerShell blocks environment activation, use the environment's interpreter
directly: `.\.venv\Scripts\python.exe laya_agentic_studio_v15.py`.

### First launch and model setup

1. Open **Models → Provider / Endpoint**.
2. Select local Ollama, remote Ollama, or OpenAI-compatible. For a local
   Ollama server, start Ollama and use its endpoint (usually
   `http://localhost:11434`). For a remote server, enter its reachable URL and
   credentials if required.
3. Connect to discover models, then select the inference and map-reduce
   models. For local Ollama, pull at least one model first, for example:

   ```bash
   ollama pull llama3.2
   ```

4. Add files if analyzing existing content, or enter Agent Directives to
   create content from scratch. Enable **Output / Export** if you want to save
   the answer as DOCX, PDF, TXT, Markdown, or HTML.
5. Choose **Manual** or **LAYA** strategy, review the run plan, and start the
   operation. Generated Python execution is optional and disabled by default.

The app stores settings and its SQLite database in the current user's
OS-specific data directory, not beside the executable. Do not put personal
documents, databases, API keys, or model weights into the source repository.

## Model Providers and Voice

Open **Models → Provider / Endpoint** to configure local Ollama, remote Ollama, or an OpenAI-compatible local/remote API. Connect to list available models, then choose separate inference and map-reduce models. API keys are stored in the operating system credential vault. **Ollama Setup** can check a configured Ollama endpoint and pull models.

Voice settings include the “Ok Chacha” wake phrase, routing to Agent Directives or Workspace Chat, command execution, speech rate, and female/male read-aloud voice selection. Wake-word and dictation transcription use Google's online speech recognition service when enabled. Read-aloud playback can be stopped from the title bar, voice panel, or by saying “stop speaking.”

With **Output / Export** enabled, the app directly saves its completed response as DOCX, PDF, TXT, Markdown, or HTML.

The application stores its database, settings, and downloaded update packages in the OS-specific per-user data directory. A database in the old working directory is migrated on first launch.

## Desktop Releases and Updates

Tagged releases build a Windows per-user installer, a macOS DMG, and a Linux per-user tarball installer. PyInstaller bundles the application/runtime; Ollama and Tesseract remain separately installed services. The release workflow also signs a manifest containing each update archive's SHA-256 digest with Ed25519. Packaged apps verify both the signature and digest before offering an update from **Help → Check for Updates**. OTA updates replace the complete app bundle rather than applying binary deltas.

Before publishing signed releases, create or reuse the local signing key with
`python tools/prepare_update_signing_key.py`. Update the
`UPDATE_PUBLIC_KEY_B64` constant in `update_service.py` to the matching public
key printed by the tool, then commit that change. Add the printed private-key
value as the repository's GitHub Actions secret
`LAYA_UPDATE_SIGNING_KEY_B64`. Never commit the private key or paste it into
source code. Verify that the public/private pair matches before publishing.

```bash
python tools/prepare_update_signing_key.py
```

Keep the PEM file at the per-user configuration path printed by the tool
backed up privately. Never commit or upload it. The workflow will fail closed
without the Actions secret. Push a new version tag such as `v0.1.1` to build
and publish a release. macOS builds are not Apple-signed/notarized yet;
production distribution will need an Apple Developer signing identity.
PyInstaller bundles should be treated as software packages, not as a guarantee
that Python implementation details are impossible to inspect.

Builds run on GitHub-hosted Windows, macOS, and Linux runners; this Windows workstation cannot locally validate the macOS/Linux installers.

## Project Files

- `laya_agentic_studio_v15.py` is the current application entry point.
- `AGENT.md` describes the architecture, workflows, and extension points.
- Older Python versions and local backups are personal copies and are not included in this repository.
