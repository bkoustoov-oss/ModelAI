# AGENT.md

## Overview

**Agentic Studio** is a self‑healing, multimodal AI‑assistant built on top of the **LARA Decision Engine** and **Ollama** language models.  It provides a graphical user interface (GUI) that allows users to:

1. **Load a variety of data sources** – PDF, DOCX, CSV/Excel, plain‑text, and images (via OCR).
2. **Extract dual‑stream representations** – a *spatial* (semantic) and a *raw* text stream for each document.
3. **Perform context‑aware analysis** using a large language model (LLM) that can:
   - Synthesize insights from many files.
   - Generate Python scripts (e.g., `python‑docx` report generation) wrapped in Markdown code fences.
   - Execute those scripts in a sandboxed subprocess and automatically repair failures by feeding the traceback back to the LLM (self‑healing).
4. **Interactively chat** with the agent, ask follow‑up questions, and retrieve updated results.
5. **Persist session history** in a local SQLite database for reproducibility and future reference.

The application is designed for **agentic development** – i.e., creating autonomous AI agents that can reason, act, and correct themselves without human intervention beyond high‑level directives.

---

## Repository Structure

```
.
├── laya_agentic_studio_v13.py      # Main application (GUI + pipeline)
├── AGENT.md                        # 📄 Documentation you are reading now
├── requirements.txt                # Python dependencies (customtkinter, ollama, laya, etc.)
├── README.md                       # Quick start guide (outside the scope of this file)
└── ... (potential assets, models, caches)
```

### Key Modules in `laya_agentic_studio_v13.py`

| Module | Purpose |
|--------|---------|
| **DatabaseManager** | Simple SQLite wrapper that logs user/agent messages per session. |
| **extract_dual_stream_from_file** | Handles file‑type specific extraction (OCR for images, `pdfplumber` for PDFs, `python-docx` for DOCX, `pandas` for CSV/Excel, plain‑text). Returns `{spatial, raw}` streams. |
| **chunk_text** | Utility to split large text into safe‑size chunks (default 80 000 characters) to stay within LLM context windows. |
| **AgenticStudioApp** (CTk subclass) | Main GUI class – builds left panel (file list, controls) and right panel (chat + preview). Handles user actions, orchestrates the pipeline, and displays streaming output. |
| **execute_agent_code** | Detects ```python``` blocks in LLM responses, writes them to a temporary file, runs them via `subprocess.run`, captures stdout/stderr, and if execution fails, automatically prompts the LLM for a corrected script (self‑healing). |
| **run_initial_agent** | Core pipeline for the first analysis: extracts all files, performs map‑reduce chunking if needed, consults the LARA decision engine for a strategy, synthesizes a final response with the chosen LLM, streams the answer to the chat, logs it, and triggers any generated code execution. |
| **_process_followup** | Handles subsequent user queries, builds a conversation history (system prompt + prior messages), streams the agent’s answer, logs it, and again runs any code blocks. |

---

## Core Concepts for Agentic Development

1. **Dual‑Stream Extraction**
   - *Spatial*: Intended for semantic understanding (layout‑aware extraction for PDFs, OCR for images).
   - *Raw*: The exact text as read from the source, useful for precise value extraction or debugging.
   - Both streams are concatenated into a single `master_raw` string that serves as the context for the LLM.

2. **Map‑Reduce Chunking**
   - When the aggregated raw data exceeds a safe limit (`SAFE_LIMIT = 80 000` chars), the system splits the data into chunks, runs a *map* step on each chunk (prompting the LLM), then aggregates the findings. This protects the model’s context window while still processing arbitrarily large corpora.

3. **LARA Decision Engine**
   - A lightweight rule‑based/ML decision engine (`laya.load("convaiinnovations/laya")`). It receives the user directive and suggests a high‑level strategy (e.g., *code_execution*, *data_extraction*, *comparison*). The result guides the final synthesis prompt.

4. **Self‑Healing Execution**
   - After the LLM returns a response, any Python code blocks are extracted.
   - The code is saved to a temporary script and executed in a subprocess.
   - If the script fails, the traceback is sent back to the LLM with a repair prompt. The LLM returns a corrected script, which is retried up to `max_retries` (default 3).
   - This loop enables **autonomous debugging** without user intervention.

5. **Chat History Persistence**
   - Every user message, system log, and agent answer is stored in `chat_history` (SQLite). The schema includes: `session_id`, `timestamp`, `role`, `content`.
   - When a follow‑up is sent, the full history is reconstructed to provide the LLM with context.

---

## Typical Workflow

1. **Start the Application** – `python laya_agentic_studio_v13.py` launches the GUI.
2. **Add Files** – Click *Add Files…* and select any supported documents.
3. **Set Directives** – In the *Initial Agent Directives* textbox, type a high‑level instruction (e.g., “Compare the provided documents and extract common numbers.”).
4. **Enable Export (Optional)** – Check *Agent: Write Python Script to build .docx* and specify an output path. The system will automatically append a critical directive to force the LLM to generate a `python-docx` script.
5. **Run** – Press *▶ Initialize Analysis*.
   - The system extracts data, chunks if needed, decides a strategy, synthesizes a response, streams it, and runs any generated code.
6. **Interact** – Use the chat input at the bottom to ask follow‑up questions or request additional scripts. The system will keep the context and continue the self‑healing loop.
7. **Review** – The *Content Preview* tab shows raw extraction previews; the *Interactive Workspace* tab displays the conversation log.

---

## Extending the Agent

### Adding New File Types
1. Install any required libraries (e.g., `pytesseract`, `pdfplumber`, `pandas`).
2. Extend `extract_dual_stream_from_file` with a new `elif ext == '.xyz':` block that returns a `{spatial, raw}` dictionary.

### Custom LLM or Model
- Replace the `ollama.chat` calls with another provider (OpenAI, Anthropic, etc.). Ensure the API returns a dictionary with a `message['content']` field for compatibility.

### Advanced Self‑Healing
- Increase `max_retries` or customize the repair prompt to include unit‑test scaffolding.
- Add a sandbox (e.g., Docker or `virtualenv`) for more secure execution.

---

## Development & Testing

- **Run Unit Tests** (if added) with `pytest`.
- **Lint** using `flake8` or `pylint` – the code follows PEP‑8 conventions.
- **Debugging** – All system‑level messages are logged to the chat window with a ⚙️ prefix, making it easy to trace the pipeline steps.

---

## License & Attribution

This project is provided under the MIT License.  It utilizes the following third‑party components:
- **LARA** – decision engine from ConvAI Innovations.
- **Ollama** – local LLM inference platform.
- **CustomTkinter** – modern dark‑theme UI.
- **pdfplumber**, **python‑docx**, **pandas**, **pytesseract** – file handling utilities.
- **Tesseract OCR** – required for image and PDF OCR (path set to `C:\Program Files\Tesseract-OCR\tesseract.exe`).

---

## Contact & Contributions

For issues, feature requests, or contributions, please open a GitHub issue or submit a pull request to the repository.  When extending the agent, keep the **self‑healing** philosophy in mind: any generated code should be robust enough to be automatically corrected by the LLM if it fails.

---

*End of AGENT.md*