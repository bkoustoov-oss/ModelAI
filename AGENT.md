# AGENT.md

## Overview

**Agentic Studio** is a self‑healing, multimodal AI‑assistant built on top of the **LARA Decision Engine** and **Ollama** language models.  It provides a graphical user interface (GUI) that allows users to:

1. **Load a variety of data sources** – PDF, DOCX, CSV/Excel, plain‑text, and images (via OCR).
2. **Extract dual‑stream representations** – a *spatial* (semantic) and a *raw* text stream for each document.
3. **Perform context‑aware analysis** using a large language model (LLM) that can:
   - Synthesize insights from many files.
   - Generate Python scripts (e.g., `python‑docx` report generation) wrapped in Markdown code fences.
   - Execute those scripts in a sandboxed subprocess and automatically repair failures by feeding the traceback back to the LLM (self‑healing).
4. **Interactively chat** with the agent, ask follow‑up questions, use voice input (via microphone), and retrieve updated results. Includes a dedicated **General Chat** for interactions independent of loaded documents.
5. **Persist session history & telemetry** in a local SQLite database (`agentic_memory.db`) capturing chat history, app settings, file access logs, model usage, and session metadata.
6. **Customize the environment** – featuring a dynamic Theme System with multiple themes, a Model Manager for selecting distinct inference and map-reduce models, and extensive UI improvements like a system terminal and zoom capabilities.

The application is designed for **agentic development** – i.e., creating autonomous AI agents that can reason, act, and correct themselves without human intervention beyond high‑level directives.

---

## Repository Structure

```
.
├── laya_agentic_studio_v15.py      # Main application (GUI + pipeline)
├── AGENT.md                        # 📄 Documentation you are reading now
├── requirements.txt                # Python dependencies (customtkinter, ollama, laya, speech_recognition, etc.)
├── README.md                       # Quick start guide (outside the scope of this file)
└── ... (potential assets, models, caches)
```

### Key Modules in `laya_agentic_studio_v15.py`

| Module | Purpose |
|--------|---------|
| **DatabaseManager** | Extended SQLite wrapper that logs user/agent messages per session, file accesses, model usage statistics, queries, and stores application settings. |
| **ModelManager** | Connects to local/remote Ollama or OpenAI-compatible endpoints, lists models, normalizes streaming responses, and stores API keys in the OS credential vault. |
| **extract_dual_stream_from_file** | Handles file‑type specific extraction (OCR for images, `pdfplumber` for PDFs, `python-docx` for DOCX, `pandas` for CSV/Excel, plain‑text). Returns `{spatial, raw}` streams. |
| **chunk_text** | Utility to split large text into safe‑size chunks (default 80 000 characters) to stay within LLM context windows. |
| **AgenticStudioApp** (CTk subclass) | Main GUI class – builds the fixed layout left panel (file list, inference control, models) and the multi-tab right panel (Workspace, General Chat, Terminal, Preview). Handles user actions, voice inputs, orchestrates the pipeline, manages themes, and displays streaming output. |
| **execute_agent_code** | Detects ```python``` blocks in LLM responses, writes them to a temporary file, runs them via `subprocess.run`, captures stdout/stderr, and if execution fails, automatically prompts the LLM for a corrected script (self‑healing). |
| **run_initial_agent** | Core pipeline for the first analysis: optionally extracts uploaded files, performs map-reduce chunking if needed, consults the decision engine, synthesizes and streams the response, optionally exports it directly, and triggers generated code execution. Directives can create content without source files. |
| **_process_followup** / **_process_general_chat** | Handles subsequent user queries in the document context workspace or general chat, streams the agent’s answer, logs it, and again runs any code blocks. |

---

## Core Concepts for Agentic Development

1. **Dual‑Stream Extraction**
   - *Spatial*: Intended for semantic understanding (layout‑aware extraction for PDFs, OCR for images).
   - *Raw*: The exact text as read from the source, useful for precise value extraction or debugging.
   - Both streams are concatenated into a single `master_raw` string that serves as the context for the LLM.

2. **Map‑Reduce Chunking**
   - When the aggregated raw data exceeds a safe limit (`SAFE_LIMIT = 80 000` chars), the system splits the data into chunks, runs a *map* step on each chunk (using the chosen Map-Reduce model), then aggregates the findings. This protects the model’s context window while still processing arbitrarily large corpora.

3. **LARA Decision Engine**
   - A lightweight rule‑based/ML decision engine (`laya.load("convaiinnovations/laya")`). It receives the user directive and suggests a high‑level strategy (e.g., *code_execution*, *data_extraction*, *comparison*). The result guides the final synthesis prompt.

4. **Self‑Healing Execution**
   - After the LLM returns a response, any Python code blocks are extracted.
   - The code is saved to a temporary script and executed in a subprocess.
   - If the script fails, the traceback is sent back to the LLM with a repair prompt. The LLM returns a corrected script, which is retried up to `max_retries` (default 3).
   - This loop enables **autonomous debugging** without user intervention.
   - Users can manually re-run scripts via right-click context menus on agent bubbles.

5. **Extended Memory and Telemetry**
   - **Chat History**: Every user message, system log, and agent answer is stored in SQLite for persistence. An **in-memory conversation cache** (`_workspace_history`, `_general_history`) is maintained live to build LLM context without re-querying the database every turn. Oldest messages are automatically trimmed when total character count exceeds safe limits (160K for workspace, 120K for general chat).
   - **Telemetry**: Tracks file access, application settings, queries, and model usage (execution time and token character counts). All these logs can be viewed and exported directly from the Menubar.

---

## Typical Workflow

1. **Start the Application** – `python laya_agentic_studio_v15.py` launches the GUI.
2. **Setup Environment** – From the Menu, you can select Themes, View Model Usage Logs, or configure Inference/Map-Reduce models in Preferences.
3. **Add Files** – Click *+ Add Files* and select any supported documents.
4. **Set Directives** – In the *Agent Directives* textbox, type or speak (using the Microphone button) a high‑level instruction.
5. **Create or Export (Optional)** – Source files are optional; directives can ask the agent to create a letter or other content from scratch. Enable *Output / Export*, choose DOCX/PDF/TXT/Markdown/HTML, and select a path to save the finished response directly.
6. **Run** – Press *▶ Initialize Analysis*.
   - The system optionally extracts data, chunks if needed, decides a strategy, synthesizes a response with the selected provider, streams it, saves requested exports, and runs generated code.
   - Check the **System Terminal** tab to view live detailed logs of the process.
   - Mouse wheel scrolling works in both Workspace and General Chat panes via a custom `bind_all` handler that walks the widget ancestry tree to route events correctly.
7. **Interact** – Use Workspace or General Chat for follow-up questions. Voice can route requests, execute directives, or read the latest response aloud; stop speech from the title bar or voice panel. Right-click messages to copy or re-run code.
8. **Review** – The *Preview* tab shows raw extraction previews for selected files in the left panel.

---

## Extending the Agent

### Adding New File Types
1. Install any required libraries (e.g., `pytesseract`, `pdfplumber`, `pandas`).
2. Extend `extract_dual_stream_from_file` with a new `elif ext == '.xyz':` block that returns a `{spatial, raw}` dictionary.

### Custom LLM or Model Configuration
- The system defaults to local Ollama. Configure remote Ollama or an OpenAI-compatible endpoint from **Models → Provider / Endpoint**.
- `ModelManager` normalizes provider responses to the `message['content']` shape expected by the UI.

### Advanced Self‑Healing
- Increase `max_retries` in `execute_agent_code` or customize the repair prompt to include unit‑test scaffolding.
- Add a sandbox (e.g., Docker or `virtualenv`) for more secure execution.

---

## Development & Testing

- **Run Unit Tests** (if added) with `pytest`.
- **Lint** using `flake8` or `pylint` – the code follows PEP‑8 conventions.
- **Debugging** – Detailed system‑level messages and tracebacks are logged directly to the **System Terminal** tab, making it easy to trace the pipeline steps, trace chunk mapping, and code executions.

---

## License & Attribution

This project is provided under the GNU General Public License v3.0. It utilizes the following third‑party components:
- **LARA** – decision engine from ConvAI Innovations.
- **Ollama** – local LLM inference platform.
- **CustomTkinter** – modern dark/light‑theme UI.
- **pdfplumber**, **python‑docx**, **pandas**, **pytesseract** – file handling utilities.
- **SpeechRecognition** – voice to text capabilities.
- **OpenAI Python client** – OpenAI-compatible providers.
- **keyring** – OS credential-vault storage for provider API keys.
- **pyttsx3** – system text-to-speech.
- **ReportLab** – PDF export.
- **Paramiko** – SSH connections and SFTP file transfer.
- **Tesseract OCR** – required for image and PDF OCR.

---

## Contact & Contributions

For issues, feature requests, or contributions, please open a GitHub issue or submit a pull request to the repository.  When extending the agent, keep the **self‑healing** philosophy in mind: any generated code should be robust enough to be automatically corrected by the LLM if it fails.

---

*End of AGENT.md*