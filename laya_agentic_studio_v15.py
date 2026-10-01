import datetime
import json
import os
import re
import sqlite3
import subprocess
import sys
import threading
import uuid
import html
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
import tempfile
import shutil
import time
import webbrowser
from urllib.parse import urlsplit

# Force offline caching
os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["TRANSFORMERS_OFFLINE"] = "1"

import customtkinter as ctk

# ============================================================
# COMMIT 1 — THEME SYSTEM
# ============================================================
THEMES = {
    "dark_navy": {
        "bg_deep":    "#07090F",
        "bg_panel":   "#0C1220",
        "bg_card":    "#121A2E",
        "bg_input":   "#0E1525",
        "bg_glass":   "#10192C",
        "accent":     "#4F8EF7",
        "accent_hot": "#7B5CF0",
        "accent_glow":"#2A5ABF",
        "success":    "#22C55E",
        "warning":    "#F59E0B",
        "danger":     "#EF4444",
        "text_hi":    "#F1F5F9",
        "text_mid":   "#94A3B8",
        "text_lo":    "#3D5273",
        "border":     "#1A2A42",
        "border_hi":  "#2D4A72",
        "glow_idle":  "#22C55E",
        "glow_busy":  "#F59E0B",
        "glow_err":   "#EF4444",
    },
    "dark_slate": {
        "bg_deep":    "#0D1117",
        "bg_panel":   "#161B22",
        "bg_card":    "#1C2129",
        "bg_input":   "#12161E",
        "bg_glass":   "#1A2030",
        "accent":     "#58A6FF",
        "accent_hot": "#3FB950",
        "accent_glow":"#1F5799",
        "success":    "#3FB950",
        "warning":    "#D29922",
        "danger":     "#F85149",
        "text_hi":    "#E6EDF3",
        "text_mid":   "#8B949E",
        "text_lo":    "#484F58",
        "border":     "#30363D",
        "border_hi":  "#58A6FF",
        "glow_idle":  "#3FB950",
        "glow_busy":  "#D29922",
        "glow_err":   "#F85149",
    },
    "midnight_purple": {
        "bg_deep":    "#0A0818",
        "bg_panel":   "#110D25",
        "bg_card":    "#181030",
        "bg_input":   "#0E0B20",
        "bg_glass":   "#140E28",
        "accent":     "#A78BFA",
        "accent_hot": "#F472B6",
        "accent_glow":"#5B21B6",
        "success":    "#34D399",
        "warning":    "#FBBF24",
        "danger":     "#F87171",
        "text_hi":    "#FAF5FF",
        "text_mid":   "#C4B5FD",
        "text_lo":    "#553C9A",
        "border":     "#2E1D5E",
        "border_hi":  "#7C3AED",
        "glow_idle":  "#34D399",
        "glow_busy":  "#FBBF24",
        "glow_err":   "#F87171",
    },
    "light": {
        "bg_deep":    "#EEF2F7",
        "bg_panel":   "#FFFFFF",
        "bg_card":    "#F5F7FA",
        "bg_input":   "#EDF0F5",
        "bg_glass":   "#FAFBFD",
        "accent":     "#2563EB",
        "accent_hot": "#7C3AED",
        "accent_glow":"#BFDBFE",
        "success":    "#16A34A",
        "warning":    "#D97706",
        "danger":     "#DC2626",
        "text_hi":    "#0F172A",
        "text_mid":   "#475569",
        "text_lo":    "#94A3B8",
        "border":     "#CBD5E1",
        "border_hi":  "#2563EB",
        "glow_idle":  "#16A34A",
        "glow_busy":  "#D97706",
        "glow_err":   "#DC2626",
    },
}

PALETTE = THEMES["dark_navy"]   # active palette; mutated by apply_theme()
FONT_SCALE = {"h1": 20, "h2": 15, "body": 13, "small": 11, "mono": 12}

FILE_ICON_MAP = {
    ".pdf": "📋", ".docx": "📝", ".doc": "📝",
    ".xlsx": "📊", ".csv": "📊",
    ".txt": "📄",
    ".png": "🖼", ".jpg": "🖼", ".jpeg": "🖼",
}

OUTPUT_FORMATS = {
    "DOCX": ".docx",
    "PDF": ".pdf",
    "Text": ".txt",
    "Markdown": ".md",
    "HTML": ".html",
}

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ============================================================
# PRE-FLIGHT
# ============================================================
missing_deps = []
try:
    from docx import Document
except ImportError:
    missing_deps.append("python-docx")

try:
    import laya
    decision_engine = laya.load("convaiinnovations/laya")
except ImportError:
    missing_deps.append("laya")
    decision_engine = None

try:
    import ollama
except ImportError:
    missing_deps.append("ollama")
    ollama = None

try:
    from openai import OpenAI
except ImportError:
    missing_deps.append("openai")
    OpenAI = None

try:
    import keyring
except ImportError:
    missing_deps.append("keyring")
    keyring = None

try:
    from PIL import Image
    import pytesseract
    _tess = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(_tess):
        pytesseract.pytesseract.tesseract_cmd = _tess
except ImportError:
    missing_deps.append("pytesseract/PIL")

try:
    import speech_recognition as sr
except ImportError:
    missing_deps.append("SpeechRecognition")
    sr = None

try:
    import pyttsx3
except ImportError:
    missing_deps.append("pyttsx3")
    pyttsx3 = None

try:
    import paramiko
except ImportError:
    paramiko = None
    missing_deps.append("paramiko")

# ============================================================
# COMMIT 1 — EXTENDED DATABASE (7G)
# ============================================================
class DatabaseManager:
    def __init__(self, db_path="agentic_memory.db"):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as c:
            c.execute("PRAGMA journal_mode=WAL")
            c.executescript('''
                CREATE TABLE IF NOT EXISTS chat_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT, timestamp TEXT, role TEXT, content TEXT
                );
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY, value TEXT, updated_at TEXT
                );
                CREATE TABLE IF NOT EXISTS file_access_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT, timestamp TEXT,
                    file_path TEXT, file_size INTEGER, action TEXT
                );
                CREATE TABLE IF NOT EXISTS model_usage_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT, timestamp TEXT,
                    model_name TEXT, prompt_chars INTEGER,
                    response_chars INTEGER, duration_ms INTEGER
                );
                CREATE TABLE IF NOT EXISTS query_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT, timestamp TEXT,
                    query_text TEXT, strategy TEXT, result_summary TEXT
                );
                CREATE TABLE IF NOT EXISTS sessions (
                    session_id TEXT PRIMARY KEY,
                    started_at TEXT, files_loaded INTEGER,
                    queries_made INTEGER, model_used TEXT
                );
                CREATE TABLE IF NOT EXISTS ssh_profiles (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    alias TEXT, host TEXT, user TEXT,
                    port INTEGER, key_path TEXT, model_endpoint TEXT
                );
            ''')
            c.commit()

    def log_message(self, session_id, role, content):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as c:
            c.execute("PRAGMA journal_mode=WAL")
            c.execute("INSERT INTO chat_history (session_id,timestamp,role,content) VALUES(?,?,?,?)",
                      (session_id, ts, role, content))
            c.commit()

    def get_session_history(self, session_id, limit=10):
        with sqlite3.connect(self.db_path) as c:
            cur = c.execute(
                "SELECT role,content FROM chat_history WHERE session_id=? ORDER BY id DESC LIMIT ?",
                (session_id, limit))
            rows = cur.fetchall()
        return [{"role": r[0], "content": r[1]} for r in reversed(rows)]

    def search_history(self, query):
        with sqlite3.connect(self.db_path) as c:
            cur = c.execute(
                "SELECT session_id, timestamp, role, content FROM chat_history "
                "WHERE content LIKE ? ORDER BY id DESC LIMIT 50",
                (f"%{query}%",))
            rows = cur.fetchall()
        return [{"session_id": r[0], "timestamp": r[1], "role": r[2], "content": r[3]} for r in rows]

    def get_setting(self, key, default=None):
        with sqlite3.connect(self.db_path) as c:
            cur = c.execute("SELECT value FROM app_settings WHERE key=?", (key,))
            row = cur.fetchone()
        return row[0] if row else default

    def set_setting(self, key, value):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as c:
            c.execute("INSERT OR REPLACE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                      (key, str(value), ts))
            c.commit()

    def log_file_access(self, session_id, file_path, action="open"):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
        with sqlite3.connect(self.db_path) as c:
            c.execute("INSERT INTO file_access_log(session_id,timestamp,file_path,file_size,action) VALUES(?,?,?,?,?)",
                      (session_id, ts, file_path, size, action))
            c.commit()

    def log_model_usage(self, session_id, model, prompt_chars, resp_chars, ms):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as c:
            c.execute("INSERT INTO model_usage_log(session_id,timestamp,model_name,prompt_chars,response_chars,duration_ms) VALUES(?,?,?,?,?,?)",
                      (session_id, ts, model, prompt_chars, resp_chars, ms))
            c.commit()

    def log_query(self, session_id, query, strategy, summary):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as c:
            c.execute("INSERT INTO query_log(session_id,timestamp,query_text,strategy,result_summary) VALUES(?,?,?,?,?)",
                      (session_id, ts, query, strategy, summary))
            c.commit()

    def upsert_session(self, session_id, files_loaded, queries_made, model_used):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as c:
            c.execute("INSERT OR REPLACE INTO sessions(session_id,started_at,files_loaded,queries_made,model_used) VALUES(?,?,?,?,?)",
                      (session_id, ts, files_loaded, queries_made, model_used))
            c.commit()

    def export_session_txt(self, session_id):
        with sqlite3.connect(self.db_path) as c:
            cur = c.execute(
                "SELECT timestamp,role,content FROM chat_history WHERE session_id=? ORDER BY id",
                (session_id,))
            lines = []
            for ts, role, content in cur.fetchall():
                lines.append(f"[{ts}] {role.upper()}:\n{content}\n{'─'*60}\n")
        return "\n".join(lines)

# ============================================================
# COMMIT 2 — MODEL MANAGER (7B)
# ============================================================
class ModelManager:
    KEYRING_SERVICE = "LAYA Agentic Studio"

    def __init__(self, database=None):
        self.database = database
        self.provider_type = (database.get_setting("provider_type", "ollama")
                              if database else "ollama")
        self.endpoint = (database.get_setting(
            "provider_endpoint", "http://localhost:11434")
            if database else "http://localhost:11434")
        self._credential_id = f"{self.provider_type}:{self.endpoint}"
        self._api_key = self._load_api_key()
        self._models: list[str] = []
        self._lock = threading.RLock()
        self._client = None

    def _load_api_key(self):
        if keyring is None:
            return ""
        try:
            return keyring.get_password(self.KEYRING_SERVICE, self._credential_id) or ""
        except Exception:
            return ""

    @property
    def is_configured(self):
        return bool(self.endpoint.strip())

    def configure(self, provider_type, endpoint, api_key=None,
                   clear_api_key=False):
        provider_type = provider_type.strip().lower().replace(" ", "-")
        if provider_type not in ("ollama", "openai-compatible"):
            raise ValueError("Choose Ollama or an OpenAI-compatible provider.")
        endpoint = endpoint.strip().rstrip("/")
        parsed_endpoint = urlsplit(endpoint)
        if (parsed_endpoint.scheme not in ("http", "https") or
                not parsed_endpoint.hostname):
            raise ValueError("Enter a valid endpoint URL beginning with http:// or https://")
        if parsed_endpoint.username or parsed_endpoint.password:
            raise ValueError("Put credentials in the API key field, not in the URL.")

        credential_id = f"{provider_type}:{endpoint}"
        if clear_api_key and keyring is not None:
            try:
                keyring.delete_password(self.KEYRING_SERVICE, credential_id)
            except Exception:
                pass
            api_key = ""
        elif api_key:
            if keyring is None:
                raise RuntimeError("Install keyring to securely store API keys.")
            keyring.set_password(self.KEYRING_SERVICE, credential_id, api_key)

        with self._lock:
            previous_credential_id = self._credential_id
            self.provider_type = provider_type
            self.endpoint = endpoint
            self._credential_id = credential_id
            if api_key is not None:
                self._api_key = api_key
            elif credential_id != previous_credential_id:
                self._api_key = self._load_api_key()
            self._client = None
        if self.database:
            self.database.set_setting("provider_type", provider_type)
            self.database.set_setting("provider_endpoint", endpoint)

    def _get_client(self):
        with self._lock:
            if self._client is not None:
                return self._client
            if self.provider_type == "ollama":
                if ollama is None:
                    raise RuntimeError("Install the ollama package to use Ollama.")
                headers = ({"Authorization": f"Bearer {self._api_key}"}
                           if self._api_key else {})
                self._client = ollama.Client(
                    host=self.endpoint, headers=headers, timeout=20)
            else:
                if OpenAI is None:
                    raise RuntimeError("Install the openai package to use this provider.")
                self._client = OpenAI(
                    base_url=self.endpoint, api_key=self._api_key or "local",
                    timeout=20)
            return self._client

    @staticmethod
    def _field(value, name, default=None):
        if isinstance(value, dict):
            return value.get(name, default)
        return getattr(value, name, default)

    def list_models(self):
        client = self._get_client()
        if self.provider_type == "ollama":
            result = client.list()
            items = self._field(result, "models", []) or []
            names = [self._field(item, "model") or
                     self._field(item, "name") for item in items]
        else:
            result = client.models.list()
            items = self._field(result, "data", []) or []
            names = [self._field(item, "id") for item in items]
        return sorted({name for name in names if name})

    def refresh(self, callback=None):
        def _do():
            error = None
            try:
                names = self.list_models()
            except Exception as exc:
                names = []
                error = str(exc)
            with self._lock:
                self._models = names
            if callback:
                callback(names, error)
        threading.Thread(target=_do, daemon=True).start()

    @staticmethod
    def _ollama_message(response):
        message = ModelManager._field(response, "message", {}) or {}
        return {"message": {"content": ModelManager._field(message, "content", "") or ""}}

    def chat(self, model, messages, options=None, stream=False):
        options = options or {}
        client = self._get_client()
        if self.provider_type == "ollama":
            response = client.chat(model=model, messages=messages,
                                   options=options, stream=stream)
            if stream:
                return (self._ollama_message(chunk) for chunk in response)
            return self._ollama_message(response)

        request = {"model": model, "messages": messages, "stream": stream}
        if "temperature" in options:
            request["temperature"] = options["temperature"]
        response = client.chat.completions.create(**request)
        if stream:
            def _normalized_stream():
                for chunk in response:
                    choices = self._field(chunk, "choices", []) or []
                    if choices:
                        delta = self._field(choices[0], "delta", {}) or {}
                        content = self._field(delta, "content", "") or ""
                        if content:
                            yield {"message": {"content": content}}
            return _normalized_stream()
        choices = self._field(response, "choices", []) or []
        message = self._field(choices[0], "message", {}) if choices else {}
        return {"message": {"content": self._field(message, "content", "") or ""}}

    def pull(self, model_name, stream=True):
        if self.provider_type != "ollama":
            raise RuntimeError("Model downloads are available only for Ollama.")
        return self._get_client().pull(model_name, stream=stream)

    def get_models(self):
        with self._lock:
            return self._models[:]

# ============================================================
# DOCUMENT EXTRACTION
# ============================================================
def extract_dual_stream_from_file(file_path, log_callback=None):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.png', '.jpg', '.jpeg']:
        text = pytesseract.image_to_string(Image.open(file_path))
        return {"spatial": text, "raw": text}
    elif ext == '.pdf':
        try:
            import pdfplumber
        except ImportError:
            raise ImportError("pdfplumber not installed.")
        sp, rw = "", ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                s = page.extract_text(layout=True)
                r = page.extract_text(layout=False)
                if not s or len(s.strip()) < 15:
                    if log_callback:
                        log_callback(f"[{os.path.basename(file_path)}] OCR fallback…")
                    img = page.to_image(resolution=300).original
                    s = r = pytesseract.image_to_string(img)
                sp += (s or "") + "\n"
                rw += (r or "") + "\n"
        return {"spatial": sp, "raw": rw}
    elif ext == '.docx':
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        return {"spatial": text, "raw": text}
    elif ext == '.doc':
        try:
            import win32com.client
        except ImportError:
            raise ImportError("Legacy .doc preview requires pywin32 and Microsoft Word.")
        word = None
        document = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            document = word.Documents.Open(
                os.path.abspath(file_path), ReadOnly=True,
                AddToRecentFiles=False)
            text = document.Content.Text
        except Exception as exc:
            raise RuntimeError(
                "Could not open this .doc file. Microsoft Word must be installed.") from exc
        finally:
            if document is not None:
                document.Close(False)
            if word is not None:
                word.Quit()
        return {"spatial": text, "raw": text}
    elif ext in ['.xlsx', '.csv']:
        try:
            import pandas as pd
        except ImportError:
            raise ImportError("pandas not installed.")
        df = pd.read_csv(file_path) if ext == '.csv' else pd.read_excel(file_path)
        t = df.to_string()
        return {"spatial": t, "raw": t}
    elif ext == '.txt':
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            text = f.read()
        return {"spatial": text, "raw": text}
    else:
        raise ValueError(f"Unsupported: {ext}")

def chunk_text(text, token_limit=20000):
    cs = token_limit * 4
    return [text[i:i+cs] for i in range(0, len(text), cs)]

# ============================================================
# WIDGETS
# ============================================================
class GlowDot(tk.Canvas):
    STATES = {
        "idle": lambda: PALETTE["glow_idle"],
        "busy": lambda: PALETTE["glow_busy"],
        "error": lambda: PALETTE["glow_err"],
    }
    def __init__(self, parent, **kw):
        super().__init__(parent, width=12, height=12, highlightthickness=0,
                         bg=PALETTE["bg_deep"], **kw)
        self._state = "idle"
        self._alpha = 1.0
        self._dir = -1
        self._render()
        self._animate()

    def _render(self):
        self.delete("all")
        color = self.STATES.get(self._state, self.STATES["idle"])()
        self.configure(bg=PALETTE["bg_deep"])
        self.create_oval(2, 2, 10, 10, fill=color, outline=color)

    def _animate(self):
        self._alpha += self._dir * 0.05
        if self._alpha <= 0.3: self._dir = 1
        elif self._alpha >= 1.0: self._dir = -1
        self._render()
        self.after(70, self._animate)

    def set_state(self, state):
        self._state = state


class GlassFrame(tk.Canvas):
    """Canvas-backed frosted-glass-look panel."""
    def __init__(self, parent, width=0, height=0, **kw):
        super().__init__(parent, highlightthickness=0, bd=0,
                         bg=PALETTE["bg_deep"], **kw)
        self.bind("<Configure>", self._on_resize)

    def _on_resize(self, e=None):
        self.delete("glass")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 2 or h < 2:
            return
        # Base fill
        self.create_rectangle(0, 0, w, h,
                               fill=PALETTE["bg_panel"], outline="",
                               tags="glass")
        # Top shimmer strip
        self.create_rectangle(0, 0, w, 3,
                               fill=PALETTE["border_hi"], outline="",
                               tags="glass")
        # Left accent bar
        self.create_rectangle(0, 3, 2, h,
                               fill=PALETTE["accent_glow"], outline="",
                               tags="glass")
        # Inner glow border
        self.create_rectangle(2, 3, w-1, h-1,
                               fill="", outline=PALETTE["border"],
                               tags="glass")


class TypingIndicator(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=PALETTE["bg_card"], corner_radius=12, **kw)
        self._dots = []
        self._running = False
        self._step = 0
        for i in range(3):
            c = tk.Canvas(self, width=8, height=8, highlightthickness=0,
                          bg=PALETTE["bg_card"])
            c.grid(row=0, column=i, padx=3, pady=8)
            self._dots.append(c)
        self._render_dots()

    def _render_dots(self):
        for i, c in enumerate(self._dots):
            c.delete("all")
            color = PALETTE["accent"] if i == self._step % 3 else PALETTE["text_lo"]
            c.create_oval(1, 1, 7, 7, fill=color, outline="")

    def start(self):
        self._running = True
        self._tick()

    def _tick(self):
        if not self._running: return
        self._step += 1
        self._render_dots()
        self.after(350, self._tick)

    def stop(self):
        self._running = False


class MessageBubble(ctk.CTkFrame):
    def __init__(self, parent, sender, text, copy_callback=None,
                 rerun_callback=None, quote_callback=None,
                 speak_callback=None, **kw):
        is_user = sender == "user"
        super().__init__(parent, fg_color=PALETTE["bg_deep"],
                         corner_radius=0, **kw)
        self.grid_columnconfigure(0, weight=1)

        outer = ctk.CTkFrame(self, fg_color=PALETTE["bg_deep"], corner_radius=0)
        outer.grid(row=0, column=0, sticky="ew", padx=18, pady=7)
        outer.grid_columnconfigure(0, weight=1 if is_user else 0)
        outer.grid_columnconfigure(1, weight=0)
        outer.grid_columnconfigure(2, weight=0 if is_user else 1)

        avatar = ctk.CTkLabel(
            outer, text="YOU" if is_user else "AI", width=34, height=28,
            corner_radius=14,
            font=ctk.CTkFont(size=9, weight="bold"),
            fg_color=PALETTE["border"] if is_user else PALETTE["accent_glow"],
            text_color=PALETTE["text_hi"] if is_user else PALETTE["accent"])
        bubble_col = 1
        avatar_col = 2 if is_user else 0
        avatar.grid(row=0, column=avatar_col, padx=8, pady=(18, 0), sticky="n")

        bg = PALETTE["bg_glass"] if is_user else PALETTE["bg_card"]
        bubble = ctk.CTkFrame(outer, fg_color=bg, corner_radius=14,
                               border_width=1,
                               border_color=PALETTE["accent_glow"] if is_user
                               else PALETTE["border"])
        bubble.grid(row=0, column=bubble_col, sticky="ew")
        bubble.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            bubble, text="YOU" if is_user else "LAYA",
            font=ctk.CTkFont(size=FONT_SCALE["small"] - 1, weight="bold"),
            text_color=PALETTE["text_mid"] if is_user else PALETTE["accent"]
        ).grid(row=0, column=0, sticky="w", padx=13, pady=(8, 0))

        # Right-click context menu
        self._text = text
        self._copy_cb = copy_callback
        self._rerun_cb = rerun_callback
        self._quote_cb = quote_callback
        self._speak_cb = speak_callback
        self._code_blocks = re.findall(r'```(\w*)\n(.*?)\n```', text, re.DOTALL)

        self._render_content(bubble, text, copy_callback, rerun_callback,
                     is_user=is_user)
        self._bind_context_menu(bubble)

    def _bind_context_menu(self, widget):
        widget.bind("<Button-3>", self._show_context_menu, add="+")
        widget.bind("<Control-Button-1>", self._show_context_menu, add="+")
        for child in widget.winfo_children():
            self._bind_context_menu(child)

    def _show_context_menu(self, event):
        menu = tk.Menu(self, tearoff=0,
                       bg=PALETTE["bg_card"], fg=PALETTE["text_hi"],
                       activebackground=PALETTE["accent"],
                       activeforeground="#ffffff", bd=0, relief="flat")
        if self._copy_cb:
            menu.add_command(label="  ⎘  Copy Message",
                             command=lambda: self._copy_cb(self._text))
        for index, (language, code) in enumerate(self._code_blocks, start=1):
            if self._copy_cb:
                label = f"  Copy {language or 'code'} block {index}"
                menu.add_command(label=label,
                                 command=lambda value=code: self._copy_cb(value))
        python_blocks = [code for language, code in self._code_blocks
                         if language.lower() == "python"]
        if python_blocks and self._rerun_cb:
            menu.add_separator()
            menu.add_command(label="  ▶  Re-run Code",
                             command=lambda: self._rerun_cb(self._text))
        if self._quote_cb:
            menu.add_separator()
            menu.add_command(label="  Quote in Chat Input",
                             command=lambda: self._quote_cb(self._text))
        if self._speak_cb:
            menu.add_command(label="  Read Message Aloud",
                             command=lambda: self._speak_cb(self._text))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def _insert_inline_text(self, widget, text):
        token_pattern = re.compile(
            r'(\*\*.+?\*\*|__.+?__|~~.+?~~|`[^`]+`|\*[^*]+\*|_[^_]+_)')
        last = 0
        for match in token_pattern.finditer(text):
            if match.start() > last:
                widget.insert("end", text[last:match.start()])
            token = match.group(0)
            if token.startswith(("**", "__")):
                widget.insert("end", token[2:-2], "bold")
            elif token.startswith("~~"):
                widget.insert("end", token[2:-2], "strike")
            elif token.startswith("`"):
                widget.insert("end", token[1:-1], "inline_code")
            else:
                widget.insert("end", token[1:-1], "italic")
            last = match.end()
        if last < len(text):
            widget.insert("end", text[last:])

    def _render_rich_text(self, parent, text, row, is_user):
        cleaned = html.unescape(text)
        cleaned = re.sub(r"<br\s*/?>", "\n", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"</(?:p|div|h[1-6]|li|tr)>\s*", "\n", cleaned,
                         flags=re.IGNORECASE)
        cleaned = re.sub(r"<li\b[^>]*>", "\n• ", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"<[^>]+>", "", cleaned).replace("\r", "")
        cleaned = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1 (\2)", cleaned)
        lines = cleaned.strip().splitlines() or [""]
        width_chars = (max(20, min(64, max(map(len, lines), default=20)))
                       if is_user else 88)
        visual_lines = sum(max(1, (len(line) + width_chars - 1) // width_chars)
                           for line in lines)
        body = tk.Text(
            parent, wrap="word", width=width_chars, height=max(1, visual_lines),
            padx=13, pady=6, bd=0, relief="flat", highlightthickness=0,
            background=PALETTE["bg_glass"] if is_user else PALETTE["bg_card"],
            foreground=PALETTE["text_hi"],
            insertbackground=PALETTE["text_hi"],
            font=("Segoe UI", FONT_SCALE["body"]), spacing1=2, spacing3=3)
        body.grid(row=row, column=0, sticky="ew", padx=2, pady=(0, 7))
        body.tag_configure("bold", font=("Segoe UI", FONT_SCALE["body"], "bold"),
                           foreground=PALETTE["text_hi"])
        body.tag_configure("italic", font=("Segoe UI", FONT_SCALE["body"], "italic"))
        body.tag_configure("strike", overstrike=True)
        body.tag_configure("inline_code", font=("Consolas", FONT_SCALE["mono"]),
                           foreground="#A8FF78", background=PALETTE["bg_deep"])
        body.tag_configure("heading", font=("Segoe UI", FONT_SCALE["h2"], "bold"),
                           foreground=PALETTE["text_hi"], spacing1=8, spacing3=5)
        body.tag_configure("subheading",
                           font=("Segoe UI", FONT_SCALE["body"], "bold"),
                           foreground=PALETTE["accent"], spacing1=6)
        body.tag_configure("list_marker", foreground=PALETTE["accent"])
        body.tag_configure("table_label", font=("Segoe UI", FONT_SCALE["small"], "bold"),
                           foreground=PALETTE["accent"])

        table_headers = None
        in_table = False
        separator_pattern = re.compile(r"^\s*\|?[\s:|+-]+\|?\s*$")
        for index, raw_line in enumerate(lines):
            line = raw_line.strip()
            if not line:
                body.insert("end", "\n")
                in_table = False
                continue
            cells = [cell.strip() for cell in line.strip("|").split("|")]
            is_separator = "|" in line and separator_pattern.match(line)
            next_is_separator = (
                index + 1 < len(lines) and "|" in lines[index + 1]
                and separator_pattern.match(lines[index + 1]))
            if is_separator:
                in_table = True
                continue
            if "|" in line and (next_is_separator or in_table):
                if next_is_separator:
                    table_headers = cells
                    in_table = True
                    continue
                body.insert("end", "• ", "list_marker")
                for cell_index, value in enumerate(cells):
                    if cell_index:
                        body.insert("end", "   ")
                    label = (table_headers[cell_index]
                             if table_headers and cell_index < len(table_headers)
                             else f"Item {cell_index + 1}")
                    body.insert("end", f"{label}: ", "table_label")
                    self._insert_inline_text(body, value)
                body.insert("end", "\n")
                continue

            in_table = False
            heading = re.match(r"^(#{1,3})\s+(.*)$", line)
            bullet = re.match(r"^[-*+]\s+(.*)$", line)
            numbered = re.match(r"^(\d+[.)])\s+(.*)$", line)
            quote = re.match(r"^>\s?(.*)$", line)
            if heading:
                tag = "heading" if len(heading.group(1)) == 1 else "subheading"
                self._insert_inline_text(body, heading.group(2))
                body.tag_add(tag, "end-1l linestart", "end-1c")
            elif bullet:
                body.insert("end", "• ", "list_marker")
                self._insert_inline_text(body, bullet.group(1))
            elif numbered:
                body.insert("end", f"{numbered.group(1)} ", "list_marker")
                self._insert_inline_text(body, numbered.group(2))
            elif quote:
                self._insert_inline_text(body, quote.group(1))
                body.tag_add("italic", "end-1l linestart", "end-1c")
            else:
                self._insert_inline_text(body, line.replace("|", "  ·  "))
            body.insert("end", "\n")
        body.configure(state="disabled")

    def _render_content(self, parent, text, copy_cb, rerun_cb, is_user=False):
        code_pat = re.compile(r'```(\w*)\n(.*?)\n```', re.DOTALL)
        last = 0
        row = 1
        for m in code_pat.finditer(text):
            plain = text[last:m.start()].strip()
            if plain:
                self._render_rich_text(parent, plain, row, is_user)
                row += 1
            lang = m.group(1) or "code"
            code = m.group(2)
            cf = ctk.CTkFrame(parent, fg_color=PALETTE["bg_deep"],
                               corner_radius=8, border_width=1,
                               border_color=PALETTE["border"])
            cf.grid(row=row, column=0, sticky="ew", padx=10, pady=4)
            cf.grid_columnconfigure(0, weight=1)
            hdr = ctk.CTkFrame(cf, fg_color=PALETTE["border"], height=26,
                                corner_radius=0)
            hdr.grid(row=0, column=0, sticky="ew")
            hdr.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(hdr, text=f" {lang}",
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["accent"]).grid(row=0, column=0,
                         sticky="w", padx=8)
            if copy_cb:
                ctk.CTkButton(hdr, text="⎘", width=28, height=22,
                               fg_color="transparent",
                               hover_color=PALETTE["accent_glow"],
                               text_color=PALETTE["text_mid"],
                               command=lambda c=code: copy_cb(c)
                               ).grid(row=0, column=1, padx=4, pady=2)
            if rerun_cb:
                ctk.CTkButton(hdr, text="▶", width=28, height=22,
                               fg_color="transparent",
                               hover_color=PALETTE["success"],
                               text_color=PALETTE["text_mid"],
                               command=lambda c=code: rerun_cb(f"```python\n{c}\n```")
                               ).grid(row=0, column=2, padx=(0, 4), pady=2)
            cb = ctk.CTkTextbox(cf,
                                fg_color=PALETTE["bg_deep"],
                                text_color="#A8FF78",
                                font=ctk.CTkFont(family="Consolas",
                                                 size=FONT_SCALE["mono"]),
                                height=min(260, max(50, code.count('\n') * 18 + 36)))
            cb.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
            cb.insert("0.0", code)
            cb.configure(state="disabled")
            row += 1
            last = m.end()
        trailing = text[last:].strip()
        if trailing:
            self._render_rich_text(parent, trailing, row, is_user)


class FileCard(ctk.CTkFrame):
    def __init__(self, parent, fp, on_preview, on_remove, **kw):
        super().__init__(parent, fg_color=PALETTE["bg_card"], corner_radius=8,
                         border_width=1, border_color=PALETTE["border"], **kw)
        self.grid_columnconfigure(1, weight=1)
        ext = os.path.splitext(fp)[1].lower()
        icon = FILE_ICON_MAP.get(ext, "📄")
        name = os.path.basename(fp)
        display = name if len(name) <= 24 else name[:21] + "…"
        ctk.CTkLabel(self, text=icon, font=ctk.CTkFont(size=14),
                     width=28).grid(row=0, column=0, padx=(8, 2), pady=6)
        ctk.CTkButton(self, text=display, anchor="w",
                      fg_color="transparent", hover_color=PALETTE["border"],
                      text_color=PALETTE["text_hi"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      command=lambda: on_preview(fp)
                      ).grid(row=0, column=1, sticky="ew", pady=4)
        ctk.CTkButton(self, text="✕", width=26, height=22,
                      fg_color="transparent", hover_color=PALETTE["danger"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=11),
                      command=lambda: on_remove(fp, self)
                      ).grid(row=0, column=2, padx=(2, 6), pady=4)


# ============================================================
# SSH MANAGER
# ============================================================
class SSHManager:
    def __init__(self, log_callback=None):
        self.client = None
        self.sftp = None
        self.shell = None
        self.log_callback = log_callback
        
    def log(self, level, msg):
        if self.log_callback:
            self.log_callback(level, msg)
        else:
            print(f"[{level}] {msg}")

    def connect(self, host, port, username, password=None, key_file=None):
        if not paramiko:
            self.log("err", "Paramiko is not installed. Run: pip install paramiko")
            return False
            
        try:
            self.client = paramiko.SSHClient()
            self.client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            connect_kwargs = {
                "hostname": host,
                "port": int(port),
                "username": username,
                "timeout": 15
            }
            
            if key_file:
                pkey = None
                # Try different key types
                for key_class in [paramiko.RSAKey, paramiko.Ed25519Key, paramiko.ECDSAKey, paramiko.DSSKey]:
                    try:
                        pkey = key_class.from_private_key_file(key_file, password=password)
                        break
                    except Exception:
                        pass
                if pkey:
                    connect_kwargs["pkey"] = pkey
                else:
                    self.log("err", "Could not parse key file. Ensure format is PEM/PPK supported by paramiko.")
                    return False
            elif password:
                connect_kwargs["password"] = password
                
            self.log("info", f"SSH: Connecting to {username}@{host}:{port}...")
            self.client.connect(**connect_kwargs)
            
            self.sftp = self.client.open_sftp()
            self.shell = self.client.invoke_shell()
            self.shell.setblocking(0)
            
            self.log("ok", f"SSH: Connected to {host}")
            return True
        except Exception as e:
            self.log("err", f"SSH Connection failed: {e}")
            self.disconnect()
            return False

    def disconnect(self):
        if self.sftp:
            try: self.sftp.close()
            except: pass
        if self.shell:
            try: self.shell.close()
            except: pass
        if self.client:
            try: self.client.close()
            except: pass
        self.sftp = None
        self.shell = None
        self.client = None
        self.log("info", "SSH: Disconnected.")

    def execute_command(self, cmd):
        if not self.shell:
            self.log("warn", "SSH: Not connected.")
            return
        try:
            self.shell.send(cmd + "\n")
        except Exception as e:
            self.log("err", f"SSH: Command failed: {e}")

    def read_shell(self):
        if not self.shell:
            return ""
        try:
            if self.shell.recv_ready():
                return self.shell.recv(8192).decode('utf-8', errors='replace')
        except Exception as e:
            self.log("err", f"SSH: Read failed: {e}")
            self.disconnect()
        return ""

    def upload(self, local_path, remote_path):
        if not self.sftp: return False
        try:
            self.log("info", f"SSH: Uploading {local_path} -> {remote_path}")
            self.sftp.put(local_path, remote_path)
            self.log("ok", "SSH: Upload complete.")
            return True
        except Exception as e:
            self.log("err", f"SSH: Upload failed: {e}")
            return False

    def download(self, remote_path, local_path):
        if not self.sftp: return False
        try:
            self.log("info", f"SSH: Downloading {remote_path} -> {local_path}")
            self.sftp.get(remote_path, local_path)
            self.log("ok", "SSH: Download complete.")
            return True
        except Exception as e:
            self.log("err", f"SSH: Download failed: {e}")
            return False


# ============================================================
# MAIN APPLICATION
# ============================================================
class AgenticStudioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("LAYA Agentic Studio v15")
        self.geometry("1480x960")
        self.minsize(1200, 800)
        self.configure(fg_color=PALETTE["bg_deep"])

        # Glassmorphism: subtle window alpha
        try:
            self.wm_attributes('-alpha', 0.97)
        except Exception:
            pass

        # State
        self.db = DatabaseManager()
        self.model_var = ctk.StringVar(
            value=self.db.get_setting("inference_model", ""))
        self.map_model_var = ctk.StringVar(
            value=self.db.get_setting("map_model", ""))
        self.model_mgr = ModelManager(self.db)
        self.ssh_mgr = SSHManager(log_callback=self.append_to_terminal)
        self.session_id = str(uuid.uuid4())[:8]
        self.selected_files = []
        self.document_context = ""
        self.is_processing = False
        self.is_executing = False
        self._voice_busy = False
        if self.db.get_setting("voice_wake_upgrade_v1") != "1":
            self.voice_wake_enabled = True
            self.db.set_setting("voice_wake_enabled", "1")
            self.db.set_setting("voice_wake_upgrade_v1", "1")
        else:
            self.voice_wake_enabled = self.db.get_setting(
                "voice_wake_enabled", "1") == "1"
        self.voice_wake_phrase = self.db.get_setting(
            "voice_wake_phrase", "Ok Chacha")
        self.voice_language = self.db.get_setting("voice_language", "en-US")
        self.voice_manual_target = self.db.get_setting(
            "voice_manual_target", "Agent Directives")
        self.voice_wake_target = self.db.get_setting(
            "voice_wake_target", "Workspace Chat")
        self.voice_auto_run_directives = self.db.get_setting(
            "voice_auto_run_directives", "0") == "1"
        self.voice_speak_replies = self.db.get_setting(
            "voice_speak_replies", "0") == "1"
        self.voice_gender = self.db.get_setting("voice_gender", "Female")
        self.voice_rate = int(self.db.get_setting("voice_rate", "155"))
        self._speech_lock = threading.RLock()
        self._speech_generation = 0
        self._speech_stop_event = threading.Event()
        self._speech_engine = None
        self._speech_active = False
        self._voice_panel_speech_button = None
        self._wake_stop_event = threading.Event()
        self._wake_thread = None
        self._voice_panel = None
        self._voice_panel_entry = None
        self._voice_panel_status = None
        self._voice_panel_route = None
        self._last_agent_response = ""
        self._status_text = tk.StringVar(value="Idle")
        self._stream_agent_frame = None
        self._stream_text_acc = ""
        self._stream_label = None
        self._chat_row = 1
        self._msg_count = 0
        self._general_chat_row = 1
        self._general_stream_agent_frame = None
        self._general_stream_text_acc = ""
        self._general_stream_label = None
        self._workspace_history = []   # in-memory conversation cache
        self._general_history = []     # in-memory general chat cache
        self._active_theme = self.db.get_setting("theme", "dark_navy")
        self._zoom = 0  # font zoom offset

        # Apply persisted theme — mutate module-level PALETTE dict in-place
        if self._active_theme in THEMES:
            PALETTE.update(THEMES[self._active_theme])

        self._build_menubar()
        self._build_title_bar()
        self._build_ui()
        self._build_status_bar()
        self._start_clock()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Restore saved model
        saved_model = self.db.get_setting("inference_model")
        if saved_model:
            self.model_var.set(saved_model)
        saved_map = self.db.get_setting("map_model")
        if saved_map:
            self.map_model_var.set(saved_map)

        # Query available models in background
        self.model_mgr.refresh(callback=self._on_models_refreshed)

        if missing_deps:
            self.after(600, lambda: self.append_to_terminal("warn",
                f"Missing: {', '.join(missing_deps)}. Some features disabled."))

        # Register session
        self.db.upsert_session(self.session_id, 0, 0, self.model_var.get())
        if self.voice_wake_enabled:
            self.after(900, self._start_wake_listener)

        # Keyboard shortcuts
        self.bind("<Control-n>", lambda e: self.new_session())
        self.bind("<Control-o>", lambda e: self.add_files())
        self.bind("<Control-t>", lambda e: self._cycle_theme())
        self.bind("<Control-Return>", lambda e: self.start_pipeline())

    # --------------------------------------------------------
    # COMMIT 3 — MENUBAR (7C)
    # --------------------------------------------------------
    def _build_menubar(self):
        bar_bg  = PALETTE["bg_panel"]
        bar_fg  = PALETTE["text_hi"]
        act_bg  = PALETTE["accent"]
        card_bg = PALETTE["bg_card"]

        self.menubar = tk.Menu(self, bg=bar_bg, fg=bar_fg,
                               activebackground=act_bg,
                               activeforeground="#ffffff",
                               relief="flat", bd=0)
        self.configure(menu=self.menubar)

        def _menu(**kw):
            return tk.Menu(self.menubar, tearoff=0,
                           bg=card_bg, fg=bar_fg,
                           activebackground=act_bg,
                           activeforeground="#ffffff",
                           relief="flat", bd=0, **kw)

        # File
        fm = _menu()
        fm.add_command(label="  New Session          Ctrl+N", command=self.new_session)
        fm.add_command(label="  Add Files…           Ctrl+O", command=self.add_files)
        fm.add_separator()
        fm.add_command(label="  Export Chat as TXT", command=self.export_chat_txt)
        fm.add_command(label="  Export Chat as MD",  command=self.export_chat_md)
        fm.add_separator()
        fm.add_command(label="  Preferences",        command=self.open_settings)
        fm.add_separator()
        fm.add_command(label="  Exit",               command=self.destroy)
        self.menubar.add_cascade(label=" File ", menu=fm)

        # View
        vm = _menu()
        theme_sub = _menu()
        for t in THEMES:
            theme_sub.add_command(label=f"  {t.replace('_',' ').title()}",
                                  command=lambda th=t: self.apply_theme(th))
        vm.add_cascade(label="  Theme", menu=theme_sub)
        vm.add_separator()
        vm.add_command(label="  Toggle Terminal",  command=lambda: self.right_tabs.set("  ⚙ Terminal  "))
        vm.add_command(label="  Toggle Preview",   command=lambda: self.right_tabs.set("  🔍 Preview  "))
        vm.add_separator()
        vm.add_command(label="  Zoom In   Ctrl++", command=self._zoom_in)
        vm.add_command(label="  Zoom Out  Ctrl+-", command=self._zoom_out)
        self.menubar.add_cascade(label=" View ", menu=vm)

        # Models
        mm = _menu()
        mm.add_command(label="  Provider / Endpoint…",
                   command=self.open_provider_settings)
        mm.add_command(label="  Refresh Available Models",
                   command=self._refresh_models_ui)
        mm.add_separator()
        mm.add_command(label="  Ollama Setup…", command=self.open_ollama_setup)
        mm.add_separator()
        mm.add_command(label="  Set Inference Model…",     command=lambda: self.open_settings())
        mm.add_command(label="  Set Map-Reduce Model…",    command=lambda: self.open_settings())
        mm.add_separator()
        # Enable the SSH Remote menu entry and switch to the SSH tab when clicked.
        mm.add_command(label="  SSH Remote",
               command=lambda: self.right_tabs.set("  🔌 SSH  "))
        self.menubar.add_cascade(label=" Models ", menu=mm)

        # History
        hm = _menu()
        hm.add_command(label="  Search Chat History", command=self._search_history_ui)
        hm.add_separator()
        hm.add_command(label="  View Session Log",    command=self._view_session_log)
        hm.add_command(label="  View File Access Log",command=self._view_file_log)
        hm.add_command(label="  View Model Usage Log",command=self._view_model_log)
        hm.add_separator()
        hm.add_command(label="  Clear Session Chat",  command=self._clear_chat)
        self.menubar.add_cascade(label=" History ", menu=hm)

        # Help
        hlp = _menu()
        hlp.add_command(label="  About",             command=self._show_about)
        hlp.add_command(label="  Keyboard Shortcuts",command=self._show_shortcuts)
        self.menubar.add_cascade(label=" Help ", menu=hlp)

    # --------------------------------------------------------
    # TITLE BAR
    # --------------------------------------------------------
    def _build_title_bar(self):
        self.title_bar = tk.Frame(self, bg=PALETTE["bg_deep"], height=38)
        self.title_bar.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.title_bar.grid_columnconfigure(1, weight=1)

        left = tk.Frame(self.title_bar, bg=PALETTE["bg_deep"])
        left.grid(row=0, column=0, sticky="w", padx=10)
        self.glow_dot = GlowDot(left)
        self.glow_dot.pack(side="left", padx=(0, 8), pady=9)
        tk.Label(left, text="LAYA  Agentic Studio",
                 font=("Consolas", 12, "bold"),
                 bg=PALETTE["bg_deep"], fg=PALETTE["text_hi"]).pack(side="left")
        tk.Label(left, text=" v15",
                 font=("Consolas", 10),
                 bg=PALETTE["bg_deep"], fg=PALETTE["accent"]).pack(side="left")

        # Center: session + msg count
        mid = tk.Frame(self.title_bar, bg=PALETTE["bg_deep"])
        mid.grid(row=0, column=1)
        self._session_lbl = tk.Label(mid,
                                     text=f"SESSION  {self.session_id}  ·  0 messages",
                                     font=("Consolas", 8),
                                     bg=PALETTE["bg_deep"], fg=PALETTE["text_lo"])
        self._session_lbl.pack()

        self._voice_toggle_btn = tk.Button(
            self.title_bar, command=self._toggle_wake_listener,
            relief="flat", bd=0, padx=10, pady=4,
            font=("Segoe UI", 8, "bold"))
        self._voice_toggle_btn.grid(row=0, column=2, sticky="e", padx=10)
        self._update_wake_button()
        self._speech_stop_btn = tk.Button(
            self.title_bar, text="■ Stop voice", command=self._stop_speech,
            state="disabled", relief="flat", bd=0, padx=8, pady=4,
            font=("Segoe UI", 8, "bold"),
            bg=PALETTE["bg_card"], fg=PALETTE["text_mid"],
            activebackground=PALETTE["danger"], activeforeground="#ffffff")
        self._speech_stop_btn.grid(row=0, column=3, sticky="e", padx=(0, 10))

        # Drag
        self.title_bar.bind("<ButtonPress-1>", self._start_move)
        self.title_bar.bind("<B1-Motion>", self._do_move)
        self._drag_x = self._drag_y = 0

    def _start_move(self, e):
        self._drag_x, self._drag_y = e.x_root, e.y_root

    def _do_move(self, e):
        self.geometry(f"+{self.winfo_x()+e.x_root-self._drag_x}+{self.winfo_y()+e.y_root-self._drag_y}")
        self._drag_x, self._drag_y = e.x_root, e.y_root

    # --------------------------------------------------------
    # COMMIT 4 — LEFT PANEL FIXED LAYOUT (7D)
    # --------------------------------------------------------
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0)  # title bar
        self.grid_rowconfigure(1, weight=1)  # content
        self.grid_rowconfigure(2, weight=0)  # status

        # Fixed left panel — NO SCROLL
        self.left_panel = ctk.CTkFrame(self, fg_color=PALETTE["bg_panel"],
                                       corner_radius=0, width=280,
                                       border_width=0)
        self.left_panel.grid(row=1, column=0, sticky="nsew",
                              padx=(6, 0), pady=(2, 2))
        self.left_panel.grid_propagate(False)
        self.left_panel.grid_columnconfigure(0, weight=1)
        # Row weights: 0=sep, 1=datasrc hdr, 2=file list, 3=file btns,
        #              4=model hdr, 5=inf model, 6=map label, 7=map combo, 8=refresh,
        #              9=ctrl hdr, 10=temp slider,
        #              11=export hdr, 12=export chk+entry, 13=browse btn,
        #              14=dir hdr, 15=directives(flex), 16=run btn, 17=progress, 18=settings
        for i, w in [(0,0),(1,0),(2,1),(3,0),(4,0),(5,0),(6,0),(7,0),(8,0),
                     (9,0),(10,0),(11,0),(12,0),(13,0),
                     (14,0),(15,2),(16,0),(17,0),(18,0)]:
            self.left_panel.grid_rowconfigure(i, weight=w)

        # Glass top border on left panel
        sep = tk.Frame(self.left_panel, bg=PALETTE["accent"], height=2)
        sep.grid(row=0, column=0, sticky="ew")

        self._section_label_grid(1, "DATA SOURCES")

        # File list — internal scrollable, max visible 4
        self.file_list_frame = ctk.CTkScrollableFrame(
            self.left_panel,
            fg_color=PALETTE["bg_card"],
            corner_radius=6,
            scrollbar_fg_color=PALETTE["bg_panel"],
            height=110)
        self.file_list_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 4))
        self.file_list_frame.grid_columnconfigure(0, weight=1)

        # File buttons
        bframe = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        bframe.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 6))
        bframe.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(bframe, text="＋  Add Files", command=self.add_files,
                      fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                      height=30, corner_radius=6
                      ).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ctk.CTkButton(bframe, text="✕", command=self.clear_files,
                      fg_color=PALETTE["bg_card"], hover_color=PALETTE["danger"],
                      text_color=PALETTE["text_mid"], width=36, height=30, corner_radius=6
                      ).grid(row=0, column=1)

        self._section_label_grid(4, "MODELS")

        # Inference model combo
        self.model_combo = ctk.CTkComboBox(
            self.left_panel,
            variable=self.model_var,
            values=self.model_mgr.get_models(),
            fg_color=PALETTE["bg_input"],
            border_color=PALETTE["border"],
            button_color=PALETTE["accent"],
            button_hover_color=PALETTE["accent_hot"],
            text_color=PALETTE["text_hi"],
            dropdown_fg_color=PALETTE["bg_card"],
            dropdown_text_color=PALETTE["text_hi"],
            dropdown_hover_color=PALETTE["accent"],
            font=ctk.CTkFont(size=FONT_SCALE["small"]),
            command=lambda v: self.db.set_setting("inference_model", v))
        self.model_combo.grid(row=5, column=0, sticky="ew", padx=8, pady=(2, 3))

        ctk.CTkLabel(self.left_panel, text="Map-Reduce Model",
                     font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                     text_color=PALETTE["text_lo"]).grid(row=6, column=0,
                     sticky="w", padx=8, pady=(2, 0))
        self.map_combo = ctk.CTkComboBox(
            self.left_panel,
            variable=self.map_model_var,
            values=self.model_mgr.get_models(),
            fg_color=PALETTE["bg_input"],
            border_color=PALETTE["border"],
            button_color=PALETTE["accent_hot"],
            button_hover_color=PALETTE["accent"],
            text_color=PALETTE["text_mid"],
            dropdown_fg_color=PALETTE["bg_card"],
            dropdown_text_color=PALETTE["text_hi"],
            dropdown_hover_color=PALETTE["accent"],
            font=ctk.CTkFont(size=FONT_SCALE["small"]),
            command=lambda v: self.db.set_setting("map_model", v))
        self.map_combo.grid(row=7, column=0, sticky="ew", padx=8, pady=(0, 3))

        model_actions = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        model_actions.grid(row=8, column=0, sticky="ew", padx=8, pady=(0, 4))
        for column in range(3):
            model_actions.grid_columnconfigure(column, weight=1)
        ctk.CTkButton(model_actions, text="↺ Refresh",
                  command=self._refresh_models_ui,
                  fg_color="transparent", hover_color=PALETTE["border"],
                  text_color=PALETTE["text_lo"],
                  font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                  height=24, corner_radius=4
                  ).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(model_actions, text="Provider",
                  command=self.open_provider_settings,
                  fg_color=PALETTE["bg_card"], hover_color=PALETTE["accent_glow"],
                  text_color=PALETTE["text_hi"],
                  font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                  height=24, corner_radius=4
                  ).grid(row=0, column=1, sticky="ew", padx=3)
        ctk.CTkButton(model_actions, text="Ollama",
                  command=self.open_ollama_setup,
                  fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"],
                  text_color=PALETTE["text_mid"],
                  font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                  height=24, corner_radius=4
                  ).grid(row=0, column=2, sticky="ew", padx=(3, 0))

        self._section_label_grid(9, "INFERENCE CONTROL")

        # Temp slider
        temp_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        temp_frame.grid(row=10, column=0, sticky="ew", padx=8, pady=(2, 4))
        temp_frame.grid_columnconfigure(0, weight=1)
        self.temp_var = tk.DoubleVar(value=float(self.db.get_setting("temperature", "0.2")))
        self.temp_slider = ctk.CTkSlider(temp_frame, from_=0.0, to=1.0,
                                          variable=self.temp_var,
                                          command=self.update_temp_label,
                                          progress_color=PALETTE["accent"],
                                          button_color=PALETTE["accent_hot"],
                                          button_hover_color=PALETTE["text_hi"],
                                          height=14)
        self.temp_slider.grid(row=0, column=0, sticky="ew")
        self.temp_label = ctk.CTkLabel(temp_frame,
                                       text="0.20 — Strict",
                                       font=ctk.CTkFont(size=FONT_SCALE["small"]-1,
                                                        slant="italic"),
                                       text_color=PALETTE["success"])
        self.temp_label.grid(row=1, column=0, sticky="w")
        self.update_temp_label(self.temp_var.get())

        self._section_label_grid(11, "OUTPUT / EXPORT")

        # Export checkbox + path entry on same row sub-frame
        exp_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        exp_frame.grid(row=12, column=0, sticky="ew", padx=8, pady=(2, 2))
        exp_frame.grid_columnconfigure(2, weight=1)
        self.export_var = tk.BooleanVar(value=False)
        self.output_path_var = ctk.StringVar()
        self.output_format_var = ctk.StringVar(
            value=self.db.get_setting("output_format", "DOCX"))
        ctk.CTkCheckBox(exp_frame, text="",
                        variable=self.export_var,
                        command=self._toggle_export,
                        fg_color=PALETTE["accent"],
                        hover_color=PALETTE["accent_hot"],
                        width=22, height=22
                        ).grid(row=0, column=0, padx=(0, 4))
        self.output_format_combo = ctk.CTkComboBox(
            exp_frame, variable=self.output_format_var,
            values=list(OUTPUT_FORMATS), width=86, height=28,
            state="disabled", command=self._change_output_format,
            fg_color=PALETTE["bg_input"], border_color=PALETTE["border"],
            button_color=PALETTE["accent"],
            text_color=PALETTE["text_hi"],
            dropdown_fg_color=PALETTE["bg_card"],
            dropdown_text_color=PALETTE["text_hi"],
            font=ctk.CTkFont(size=FONT_SCALE["small"] - 1))
        self.output_format_combo.grid(row=0, column=1, padx=(0, 4))
        self.output_entry = ctk.CTkEntry(exp_frame,
                                         textvariable=self.output_path_var,
                                         placeholder_text="Output file path…",
                                         state="disabled",
                                         fg_color=PALETTE["bg_input"],
                                         text_color=PALETTE["text_hi"],
                                         border_color=PALETTE["border"],
                                         font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                                         height=28)
        self.output_entry.grid(row=0, column=2, sticky="ew")
        self.browse_out_btn = ctk.CTkButton(
            self.left_panel, text="📁  Browse Output Path",
            command=self.browse_output,
            state="disabled",
            fg_color="transparent",
            hover_color=PALETTE["border"],
            text_color=PALETTE["text_lo"],
            font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
            height=24, corner_radius=4)
        self.browse_out_btn.grid(row=13, column=0, sticky="e", padx=8, pady=(0, 4))

        directive_header = ctk.CTkFrame(self.left_panel, fg_color="transparent",
                                        height=26)
        directive_header.grid(row=14, column=0, sticky="ew", padx=8,
                              pady=(4, 0))
        directive_header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            directive_header, text="AGENT DIRECTIVES",
            font=ctk.CTkFont(size=FONT_SCALE["small"] - 1, weight="bold"),
            text_color=PALETTE["text_lo"]).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            directive_header, text="🎙 Dictate", width=78, height=23,
            command=lambda: self._open_voice_panel(
                route=self.voice_manual_target, capture=True, source="manual"),
            fg_color=PALETTE["accent_glow"], hover_color=PALETTE["accent"],
            text_color=PALETTE["text_hi"],
            font=ctk.CTkFont(size=FONT_SCALE["small"] - 1),
            corner_radius=5).grid(row=0, column=1, sticky="e")

        self.prompt_text = ctk.CTkTextbox(self.left_panel,
                                          fg_color=PALETTE["bg_input"],
                                          text_color=PALETTE["text_hi"],
                                          border_color=PALETTE["border_hi"],
                                          border_width=1,
                                          font=ctk.CTkFont(size=FONT_SCALE["body"]))
        self.prompt_text.grid(row=15, column=0, sticky="nsew", padx=8, pady=(2, 6))
        self.prompt_text.insert("0.0", "Compare the provided documents and extract common numbers.")

        # Gradient-effect Run button using Canvas
        self.run_canvas = tk.Canvas(self.left_panel, height=46,
                                    highlightthickness=0, bd=0,
                                    bg=PALETTE["bg_panel"])
        self.run_canvas.grid(row=16, column=0, sticky="ew", padx=8, pady=(0, 4))
        self.run_canvas.bind("<Configure>", self._draw_run_button)
        self.run_canvas.bind("<Button-1>", lambda e: self.start_pipeline())
        self.run_canvas.bind("<Enter>",
            lambda e: self.run_canvas.configure(cursor="hand2"))
        self.run_canvas.bind("<Leave>",
            lambda e: self.run_canvas.configure(cursor=""))

        self.progress_bar = ctk.CTkProgressBar(self.left_panel,
                                               mode="indeterminate",
                                               progress_color=PALETTE["accent_hot"],
                                               fg_color=PALETTE["border"], height=3)
        self.progress_bar.set(0)

        ctk.CTkButton(self.left_panel, text="⚙  Settings",
                      command=self.open_settings,
                      fg_color="transparent",
                      hover_color=PALETTE["border"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                      height=28, corner_radius=4
                      ).grid(row=18, column=0, sticky="ew", padx=8, pady=(0, 6))

        # ---- RIGHT PANEL ----
        self.right_tabs = ctk.CTkTabview(self, corner_radius=6,
                                         fg_color=PALETTE["bg_panel"],
                                         segmented_button_fg_color=PALETTE["bg_card"],
                                         segmented_button_selected_color=PALETTE["accent"],
                                         segmented_button_selected_hover_color=PALETTE["accent_hot"],
                                         segmented_button_unselected_color=PALETTE["bg_card"],
                                         segmented_button_unselected_hover_color=PALETTE["border"],
                                         text_color=PALETTE["text_hi"],
                                         text_color_disabled=PALETTE["text_lo"])
        self.right_tabs.grid(row=1, column=1, sticky="nsew",
                              padx=(4, 6), pady=(2, 2))

        self.tab_chat     = self.right_tabs.add("  💬 Workspace  ")
        self.tab_general  = self.right_tabs.add("  🤖 General Chat  ")
        self.tab_terminal = self.right_tabs.add("  ⚙ Terminal  ")
        self.tab_preview  = self.right_tabs.add("  🔍 Preview  ")
        self.tab_ssh      = self.right_tabs.add("  🔌 SSH  ")

        # ---- COMMIT 6 — INTERACTIVE WORKSPACE (7F) ----
        self.tab_chat.grid_rowconfigure(0, weight=1)
        self.tab_chat.grid_columnconfigure(0, weight=1)

        self.chat_scroll = ctk.CTkScrollableFrame(
            self.tab_chat,
            fg_color=PALETTE["bg_deep"],
            scrollbar_fg_color=PALETTE["bg_glass"],
            scrollbar_button_color=PALETTE["accent"],
            scrollbar_button_hover_color=PALETTE["accent_hot"])
        self.chat_scroll.grid(row=0, column=0, sticky="nsew", padx=2, pady=(2, 0))
        self.chat_scroll.grid_columnconfigure(0, weight=1)

        self.chat_scroll.bind("<Prior>", lambda e: self._scroll_chat(-10))
        self.chat_scroll.bind("<Next>", lambda e: self._scroll_chat(10))

        welcome = ctk.CTkLabel(self.chat_scroll,
                               text=f"✦  Session [{self.session_id}] ready\n"
                                    "Add files → set directives → Initialize Analysis",
                               font=ctk.CTkFont(size=FONT_SCALE["body"]),
                               text_color=PALETTE["text_lo"],
                               justify="center")
        welcome.grid(row=0, column=0, pady=40)

        self._typing_ind = TypingIndicator(self.chat_scroll)

        # Input area — glass-style
        input_bg = ctk.CTkFrame(self.tab_chat,
                                fg_color=PALETTE["bg_glass"],
                                corner_radius=10,
                                border_width=1,
                                border_color=PALETTE["border_hi"])
        input_bg.grid(row=1, column=0, sticky="ew", padx=6, pady=6)
        input_bg.grid_columnconfigure(0, weight=1)

        self.chat_input = ctk.CTkEntry(input_bg,
                                       placeholder_text="Ask the agent…  (Enter to send, Ctrl+Enter to run pipeline)",
                                       font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                       fg_color="transparent",
                                       border_width=0,
                                       text_color=PALETTE["text_hi"])
        self.chat_input.grid(row=0, column=0, sticky="ew", padx=10, ipady=11)
        self.chat_input.bind("<Return>", lambda e: self.send_followup())

        self.mic_btn = ctk.CTkButton(input_bg, text="🎤", width=40, height=36,
                                     fg_color="transparent", hover_color=PALETTE["border"],
                                     text_color=PALETTE["text_hi"],
                                     command=lambda: self._listen_to_mic(
                                         self.chat_input, self.mic_btn))
        self.mic_btn.grid(row=0, column=1, padx=(4, 0), pady=6)

        self.send_btn = ctk.CTkButton(input_bg, text="Send ⏎",
                                      font=ctk.CTkFont(size=FONT_SCALE["small"],
                                                       weight="bold"),
                                      width=86, height=36, corner_radius=8,
                                      fg_color=PALETTE["accent"],
                                      hover_color=PALETTE["accent_hot"],
                                      command=self.send_followup)
        self.send_btn.grid(row=0, column=2, padx=(4, 8), pady=6)

        # ---- GENERAL CHAT PANE ----
        self.tab_general.grid_rowconfigure(0, weight=1)
        self.tab_general.grid_columnconfigure(0, weight=1)

        self.general_scroll = ctk.CTkScrollableFrame(
            self.tab_general,
            fg_color=PALETTE["bg_deep"],
            scrollbar_fg_color=PALETTE["bg_glass"],
            scrollbar_button_color=PALETTE["accent"],
            scrollbar_button_hover_color=PALETTE["accent_hot"])
        self.general_scroll.grid(row=0, column=0, sticky="nsew", padx=2, pady=(2, 0))
        self.general_scroll.grid_columnconfigure(0, weight=1)

        self._general_typing_ind = TypingIndicator(self.general_scroll)

        gen_input_bg = ctk.CTkFrame(self.tab_general, fg_color=PALETTE["bg_glass"],
                                    corner_radius=10, border_width=1, border_color=PALETTE["border_hi"])
        gen_input_bg.grid(row=1, column=0, sticky="ew", padx=6, pady=6)
        gen_input_bg.grid_columnconfigure(0, weight=1)

        self.general_input = ctk.CTkEntry(gen_input_bg,
                                          placeholder_text="Chat with AI generally... (Enter to send)",
                                          font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                          fg_color="transparent", border_width=0,
                                          text_color=PALETTE["text_hi"])
        self.general_input.grid(row=0, column=0, sticky="ew", padx=10, ipady=11)
        self.general_input.bind("<Return>", lambda e: self.send_general_chat())

        self.gen_mic_btn = ctk.CTkButton(gen_input_bg, text="🎤", width=40, height=36,
                                         fg_color="transparent", hover_color=PALETTE["border"],
                                         text_color=PALETTE["text_hi"],
                                         command=lambda: self._listen_to_mic(
                                             self.general_input, self.gen_mic_btn))
        self.gen_mic_btn.grid(row=0, column=1, padx=(4, 0), pady=6)

        self.gen_send_btn = ctk.CTkButton(gen_input_bg, text="Send ⏎",
                                          font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                                          width=86, height=36, corner_radius=8,
                                          fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
                                          command=self.send_general_chat)
        self.gen_send_btn.grid(row=0, column=2, padx=(4, 8), pady=6)

        # Terminal
        self.tab_terminal.grid_rowconfigure(1, weight=1)
        self.tab_terminal.grid_columnconfigure(0, weight=1)

        t_bar = ctk.CTkFrame(self.tab_terminal,
                              fg_color=PALETTE["bg_card"], height=30)
        t_bar.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))
        ctk.CTkLabel(t_bar, text="System Terminal",
                     font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                     text_color=PALETTE["text_mid"]).pack(side="left", padx=10)
        ctk.CTkButton(t_bar, text="Clear", width=54, height=22,
                      fg_color="transparent", hover_color=PALETTE["border"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      command=self._clear_terminal).pack(side="right", padx=4)
        ctk.CTkButton(t_bar, text="Copy", width=54, height=22,
                      fg_color="transparent", hover_color=PALETTE["border"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      command=self._copy_terminal).pack(side="right", padx=2)

        self.terminal_log = ctk.CTkTextbox(
            self.tab_terminal,
            fg_color="#020408",
            text_color=PALETTE["success"],
            font=ctk.CTkFont(family="Consolas", size=FONT_SCALE["mono"]))
        self.terminal_log.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        self.terminal_log.tag_config("ok",   foreground=PALETTE["success"])
        self.terminal_log.tag_config("err",  foreground=PALETTE["danger"])
        self.terminal_log.tag_config("warn", foreground=PALETTE["warning"])
        self.terminal_log.tag_config("info", foreground=PALETTE["text_mid"])
        self.terminal_log.insert("0.0", "─── SYSTEM TERMINAL ─── \n", "info")

        # Preview
        self.tab_preview.grid_rowconfigure(0, weight=1)
        self.tab_preview.grid_columnconfigure(0, weight=1)
        self.preview_box = ctk.CTkTextbox(
            self.tab_preview,
            fg_color=PALETTE["bg_card"],
            text_color=PALETTE["text_hi"],
            font=ctk.CTkFont(family="Consolas", size=FONT_SCALE["small"]))
        self.preview_box.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.preview_box.insert("0.0",
            "Click a file card in the left panel to preview its extracted content.")

        self._build_ssh_tab()
        self._setup_scroll_bindings()
        self.bind_all("<Button-3>", self._show_text_context_menu, add="+")
        self.bind_all("<Control-Button-1>", self._show_text_context_menu, add="+")

    def _draw_run_button(self, e=None):
        c = self.run_canvas
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 4: return
        # Gradient simulation: 5 bands accent→accent_hot
        steps = 10
        for i in range(steps):
            ratio = i / steps
            r1 = int(0x4F + (0x7B - 0x4F) * ratio)
            g1 = int(0x8E + (0x5C - 0x8E) * ratio)
            b1 = int(0xF7 + (0xF0 - 0xF7) * ratio)
            color = f"#{r1:02x}{g1:02x}{b1:02x}"
            x0 = int(w * i / steps)
            x1 = int(w * (i + 1) / steps)
            c.create_rectangle(x0, 0, x1, h, fill=color, outline="")
        c.create_text(w // 2, h // 2, text="▶   Initialize Analysis",
                      fill="#ffffff",
                      font=("Consolas", 13, "bold"))

    # --------------------------------------------------------
    # STATUS BAR
    # --------------------------------------------------------
    def _build_status_bar(self):
        bar = tk.Frame(self, bg=PALETTE["border"], height=24)
        bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        bar.grid_columnconfigure(1, weight=1)
        self._file_badge = tk.Label(bar, text="0 files",
                                    font=("Consolas", 8),
                                    bg=PALETTE["border"], fg=PALETTE["text_lo"])
        self._file_badge.grid(row=0, column=0, padx=10, sticky="w")
        tk.Label(bar, textvariable=self._status_text,
                 font=("Consolas", 8),
                 bg=PALETTE["border"], fg=PALETTE["accent"]
                 ).grid(row=0, column=1)
        self._clock_lbl = tk.Label(bar, text="",
                                   font=("Consolas", 8),
                                   bg=PALETTE["border"], fg=PALETTE["text_lo"])
        self._clock_lbl.grid(row=0, column=2, padx=10, sticky="e")
        tk.Label(bar, textvariable=self.model_var,
                 font=("Consolas", 8),
                 bg=PALETTE["border"], fg=PALETTE["text_lo"]
                 ).grid(row=0, column=3, padx=10, sticky="e")

    def _start_clock(self):
        def tick():
            self._clock_lbl.configure(
                text=datetime.datetime.now().strftime("%H:%M:%S"))
            self.after(1000, tick)
        tick()

    def set_status(self, text, glow="idle"):
        self._status_text.set(text)
        self.glow_dot.set_state(glow)

    # --------------------------------------------------------
    # THEME SYSTEM
    # --------------------------------------------------------
    def apply_theme(self, name):
        if name not in THEMES:
            return
        PALETTE.update(THEMES[name])  # mutate in-place — no global needed
        self._active_theme = name
        self.db.set_setting("theme", name)
        # Lightweight CTk appearance toggle
        if name == "light":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")
        # Reconfigure root
        self.configure(fg_color=PALETTE["bg_deep"])
        # Notify user (full widget re-theme requires restart for deep changes)
        self.set_status(f"Theme: {name} (restart for full effect)", "idle")
        self.append_to_terminal("info",
            f"Theme changed to '{name}'. Restart for complete widget refresh.")

    def _cycle_theme(self):
        themes = list(THEMES.keys())
        idx = (themes.index(self._active_theme) + 1) % len(themes)
        self.apply_theme(themes[idx])

    # --------------------------------------------------------
    # SECTION HELPERS
    # --------------------------------------------------------
    def _section_label_grid(self, row, text):
        f = ctk.CTkFrame(self.left_panel, fg_color="transparent", height=22)
        f.grid(row=row, column=0, sticky="ew", padx=8, pady=(6, 1))
        f.pack_propagate(False)
        bar = ctk.CTkFrame(f, fg_color=PALETTE["accent"], width=3, corner_radius=0)
        bar.pack(side="left", fill="y", padx=(0, 6))
        ctk.CTkLabel(f, text=text,
                     font=ctk.CTkFont(size=FONT_SCALE["small"]-1, weight="bold"),
                     text_color=PALETTE["text_lo"]).pack(side="left", anchor="w")

    # --------------------------------------------------------
    # MODEL MANAGEMENT
    # --------------------------------------------------------
    def _on_models_refreshed(self, models, error=None):
        def _upd():
            self.model_combo.configure(values=models)
            self.map_combo.configure(values=models)
            if models:
                if self.model_var.get() not in models:
                    self.model_var.set(models[0])
                    self.db.set_setting("inference_model", models[0])
                if self.map_model_var.get() not in models:
                    self.map_model_var.set(models[0])
                    self.db.set_setting("map_model", models[0])
                self.append_to_terminal(
                    "info", f"{self.model_mgr.provider_type} model list refreshed: "
                    f"{len(models)} model(s) found.")
            elif error:
                self.append_to_terminal(
                    "warn", f"Could not list models from "
                    f"{self.model_mgr.endpoint}: {error}")
        self.after(0, _upd)

    def _refresh_models_ui(self):
        self.append_to_terminal(
            "info", f"Querying {self.model_mgr.provider_type} at "
            f"{self.model_mgr.endpoint}…")
        self.model_mgr.refresh(callback=self._on_models_refreshed)

    def open_provider_settings(self):
        win = ctk.CTkToplevel(self)
        win.title("Model Provider")
        win.geometry("640x700")
        win.minsize(580, 620)
        win.resizable(False, False)
        win.configure(fg_color=PALETTE["bg_panel"])
        win.grab_set()

        ctk.CTkLabel(
            win, text="Connect a model provider",
            font=ctk.CTkFont(size=FONT_SCALE["h1"], weight="bold"),
            text_color=PALETTE["text_hi"]).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkLabel(
            win,
            text="Use Ollama or any OpenAI-compatible local/remote endpoint. "
                 "Models are loaded from the endpoint you configure.",
            wraplength=560, justify="left",
            font=ctk.CTkFont(size=FONT_SCALE["body"]),
            text_color=PALETTE["text_mid"]).pack(anchor="w", padx=24, pady=(0, 14))

        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="x", padx=24)
        body.grid_columnconfigure(0, weight=1)
        is_local_ollama = (self.model_mgr.provider_type == "ollama" and
                           "localhost" in self.model_mgr.endpoint or
                           self.model_mgr.provider_type == "ollama" and
                           "127.0.0.1" in self.model_mgr.endpoint)
        provider_var = ctk.StringVar(value=(
            "OpenAI-compatible" if self.model_mgr.provider_type != "ollama"
            else "Local Ollama" if is_local_ollama else "Remote Ollama"))
        endpoint_var = ctk.StringVar(value=self.model_mgr.endpoint)
        ctk.CTkLabel(body, text="Provider",
                     text_color=PALETTE["text_mid"]).grid(
                         row=0, column=0, sticky="w", pady=(2, 3))
        provider_combo = ctk.CTkComboBox(
            body, variable=provider_var,
            values=["Local Ollama", "Remote Ollama", "OpenAI-compatible"],
            fg_color=PALETTE["bg_input"], border_color=PALETTE["border"],
            button_color=PALETTE["accent"], text_color=PALETTE["text_hi"],
            dropdown_fg_color=PALETTE["bg_card"],
            dropdown_text_color=PALETTE["text_hi"]
        )
        provider_combo.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        def provider_changed(selection):
            if selection == "Local Ollama":
                endpoint_var.set("http://localhost:11434")
            elif selection == "Remote Ollama":
                endpoint_var.set("https://your-ollama-server:11434")
            elif selection == "OpenAI-compatible":
                endpoint_var.set("http://localhost:1234/v1")

        provider_combo.configure(command=provider_changed)

        ctk.CTkLabel(body, text="Base URL / endpoint",
                     text_color=PALETTE["text_mid"]).grid(
                         row=2, column=0, sticky="w", pady=(2, 3))
        ctk.CTkEntry(
            body, textvariable=endpoint_var,
            placeholder_text="Ollama: http://localhost:11434 | OpenAI API: http://localhost:1234/v1",
            fg_color=PALETTE["bg_input"], border_color=PALETTE["border"],
            text_color=PALETTE["text_hi"]
        ).grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkLabel(body, text="API key (optional for local providers)",
                     text_color=PALETTE["text_mid"]).grid(
                         row=4, column=0, sticky="w", pady=(2, 3))
        key_entry = ctk.CTkEntry(
            body, show="•", placeholder_text="Leave blank to keep the saved key",
            fg_color=PALETTE["bg_input"], border_color=PALETTE["border"],
            text_color=PALETTE["text_hi"])
        key_entry.grid(row=5, column=0, sticky="ew", pady=(0, 5))
        key_status = ctk.CTkLabel(
            body,
            text=("A key is stored in the system credential vault for this endpoint."
                  if self.model_mgr._api_key else
                  "No saved key for the currently selected endpoint."),
            text_color=PALETTE["text_lo"],
            font=ctk.CTkFont(size=FONT_SCALE["small"] - 1))
        key_status.grid(row=6, column=0, sticky="w", pady=(0, 8))
        clear_key_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            body, text="Forget saved API key for this endpoint",
            variable=clear_key_var, fg_color=PALETTE["accent"],
            hover_color=PALETTE["accent_hot"],
            text_color=PALETTE["text_mid"]
        ).grid(row=7, column=0, sticky="w", pady=(0, 10))

        models_panel = ctk.CTkFrame(win, fg_color=PALETTE["bg_card"],
                                    corner_radius=8)
        models_panel.pack(fill="both", expand=True, padx=24, pady=(0, 8))
        models_panel.grid_columnconfigure(0, weight=1)
        models_heading = ctk.CTkLabel(
            models_panel, text="Available models",
            font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
            text_color=PALETTE["text_hi"])
        models_heading.grid(row=0, column=0, sticky="w", padx=10, pady=(8, 4))
        models_list = ctk.CTkScrollableFrame(
            models_panel, fg_color="transparent", height=175)
        models_list.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 6))
        models_list.grid_columnconfigure(0, weight=1)

        def render_models(models):
            for child in models_list.winfo_children():
                child.destroy()
            models_heading.configure(text=f"Available models ({len(models)})")
            if not models:
                ctk.CTkLabel(
                    models_list, text="Connect to an endpoint to list its models.",
                    text_color=PALETTE["text_lo"], anchor="w"
                ).grid(row=0, column=0, sticky="ew", padx=8, pady=8)
                return
            for index, model_name in enumerate(models):
                row = ctk.CTkFrame(models_list, fg_color=PALETTE["bg_input"],
                                   corner_radius=5)
                row.grid(row=index, column=0, sticky="ew", padx=4, pady=2)
                row.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    row, text=model_name, anchor="w",
                    font=ctk.CTkFont(size=FONT_SCALE["small"]),
                    text_color=PALETTE["text_hi"]
                ).grid(row=0, column=0, sticky="ew", padx=8, pady=6)
                ctk.CTkButton(
                    row, text="Use", width=58, height=26,
                    command=lambda name=model_name: choose_model(name, "inference"),
                    fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"]
                ).grid(row=0, column=1, padx=(4, 3), pady=3)
                ctk.CTkButton(
                    row, text="Map", width=58, height=26,
                    command=lambda name=model_name: choose_model(name, "map"),
                    fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"]
                ).grid(row=0, column=2, padx=(3, 5), pady=3)

        def choose_model(model_name, purpose):
            variable = self.model_var if purpose == "inference" else self.map_model_var
            setting = "inference_model" if purpose == "inference" else "map_model"
            variable.set(model_name)
            self.db.set_setting(setting, model_name)
            status_var.set(f"{model_name} selected for {purpose}.")

        render_models(self.model_mgr.get_models())

        status_var = ctk.StringVar(value="Not connected")
        ctk.CTkLabel(win, textvariable=status_var, anchor="w", wraplength=560,
                     text_color=PALETTE["text_mid"],
                     font=ctk.CTkFont(size=FONT_SCALE["small"])
                     ).pack(fill="x", padx=24, pady=(8, 12))

        actions = ctk.CTkFrame(win, fg_color="transparent")
        actions.pack(fill="x", padx=24, pady=(0, 18))
        actions.grid_columnconfigure(0, weight=1)

        def connect_and_list():
            provider_type = ("ollama" if provider_var.get() != "OpenAI-compatible"
                             else "openai-compatible")
            try:
                self.model_mgr.configure(
                    provider_type, endpoint_var.get(),
                    api_key=key_entry.get().strip() or None,
                    clear_api_key=clear_key_var.get())
            except Exception as exc:
                status_var.set(f"Configuration error: {exc}")
                return
            key_entry.delete(0, "end")
            key_status.configure(
                text=("A key is stored in the system credential vault for this endpoint."
                      if self.model_mgr._api_key else
                      "No API key stored; this is suitable for local providers."))
            status_var.set(f"Connecting to {self.model_mgr.endpoint}…")
            connect_button.configure(state="disabled")

            def listed(models, error):
                self._on_models_refreshed(models, error)
                def update():
                    if not win.winfo_exists():
                        return
                    connect_button.configure(state="normal")
                    if error:
                        status_var.set(f"Connection failed: {error}")
                    else:
                        render_models(models)
                        status_var.set(
                            f"Connected. {len(models)} model(s) found; "
                            "choose a model in the left panel.")
                self.after(0, update)

            self.model_mgr.refresh(callback=listed)

        connect_button = ctk.CTkButton(
            actions, text="Connect & List Models", command=connect_and_list,
            fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
            font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
            height=38)
        connect_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            actions, text="Close", command=win.destroy,
            fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"],
            text_color=PALETTE["text_hi"], height=38, width=90
        ).grid(row=0, column=1)

    def _set_setup_status(self, window, variable, message):
        def update():
            if window.winfo_exists():
                variable.set(message)
        self.after(0, update)

    def open_ollama_setup(self):
        win = ctk.CTkToplevel(self)
        win.title("Ollama Setup")
        win.geometry("640x720")
        win.minsize(580, 620)
        win.resizable(False, False)
        win.configure(fg_color=PALETTE["bg_panel"])
        win.grab_set()

        ctk.CTkLabel(win, text="Connect to Ollama",
                     font=ctk.CTkFont(size=FONT_SCALE["h1"], weight="bold"),
                     text_color=PALETTE["text_hi"]).pack(pady=(22, 8))
        ctk.CTkLabel(
            win,
            text="Connect to a local or remote Ollama server. For authenticated "
                 "remote servers, save the API key in Provider settings first.",
            justify="left", wraplength=570,
            font=ctk.CTkFont(size=FONT_SCALE["body"]),
            text_color=PALETTE["text_mid"]).pack(anchor="w", padx=24, pady=(0, 10))

        endpoint_var = ctk.StringVar(value=(
            self.model_mgr.endpoint if self.model_mgr.provider_type == "ollama"
            else "http://localhost:11434"))
        ctk.CTkLabel(win, text="Ollama endpoint",
                     text_color=PALETTE["text_mid"]).pack(anchor="w", padx=28)
        endpoint_entry = ctk.CTkEntry(
            win, textvariable=endpoint_var,
            placeholder_text="http://localhost:11434 or https://your-server:11434",
            fg_color=PALETTE["bg_input"], border_color=PALETTE["border"],
            text_color=PALETTE["text_hi"])
        endpoint_entry.pack(fill="x", padx=24, pady=(3, 8))

        status_var = ctk.StringVar(value="Enter an endpoint and list its models.")
        ctk.CTkLabel(win, textvariable=status_var, anchor="w", justify="left",
                     wraplength=570, text_color=PALETTE["text_hi"],
                     font=ctk.CTkFont(size=FONT_SCALE["small"])
                     ).pack(fill="x", padx=28, pady=(4, 6))

        action_row = ctk.CTkFrame(win, fg_color="transparent")
        action_row.pack(fill="x", padx=24, pady=(0, 8))
        action_row.grid_columnconfigure(0, weight=1)
        action_row.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            action_row, text="Check & Refresh Models",
            command=lambda: self._check_ollama_connection(
                win, status_var, endpoint_entry, render_ollama_models),
            fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"]
        ).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ctk.CTkButton(
            action_row, text="Ollama Cloud Sign-in",
            command=lambda: self._start_ollama_sign_in(win, status_var),
            fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"]
        ).grid(row=0, column=1, sticky="ew", padx=(5, 0))

        models_panel = ctk.CTkFrame(win, fg_color=PALETTE["bg_card"],
                                    corner_radius=8)
        models_panel.pack(fill="both", expand=True, padx=24, pady=(0, 8))
        models_panel.grid_columnconfigure(0, weight=1)
        models_heading = ctk.CTkLabel(
            models_panel, text="Available models",
            font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
            text_color=PALETTE["text_hi"])
        models_heading.grid(row=0, column=0, sticky="w", padx=10, pady=(7, 3))
        models_list = ctk.CTkScrollableFrame(
            models_panel, fg_color="transparent", height=160)
        models_list.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 5))
        models_list.grid_columnconfigure(0, weight=1)

        def render_ollama_models(models):
            for child in models_list.winfo_children():
                child.destroy()
            models_heading.configure(text=f"Available models ({len(models)})")
            if not models:
                ctk.CTkLabel(
                    models_list, text="No models returned by this endpoint.",
                    text_color=PALETTE["text_lo"], anchor="w"
                ).grid(row=0, column=0, sticky="ew", padx=8, pady=8)
                return
            for index, model_name in enumerate(models):
                row = ctk.CTkFrame(models_list, fg_color=PALETTE["bg_input"],
                                   corner_radius=5)
                row.grid(row=index, column=0, sticky="ew", padx=4, pady=2)
                row.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    row, text=model_name, anchor="w",
                    font=ctk.CTkFont(size=FONT_SCALE["small"]),
                    text_color=PALETTE["text_hi"]
                ).grid(row=0, column=0, sticky="ew", padx=8, pady=5)
                ctk.CTkButton(
                    row, text="Use", width=58, height=25,
                    command=lambda name=model_name: self._select_model(name, "inference"),
                    fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"]
                ).grid(row=0, column=1, padx=(4, 3), pady=3)
                ctk.CTkButton(
                    row, text="Map", width=58, height=25,
                    command=lambda name=model_name: self._select_model(name, "map"),
                    fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"]
                ).grid(row=0, column=2, padx=(3, 5), pady=3)

        render_ollama_models(self.model_mgr.get_models()
                             if self.model_mgr.provider_type == "ollama" else [])

        ctk.CTkLabel(win, text="Pull a local model",
                     font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
                     text_color=PALETTE["text_hi"]).pack(anchor="w", padx=28, pady=(4, 5))
        pull_row = ctk.CTkFrame(win, fg_color="transparent")
        pull_row.pack(fill="x", padx=24)
        model_entry = ctk.CTkEntry(
            pull_row, placeholder_text="Model name, for example llama3.2",
            fg_color=PALETTE["bg_input"], border_color=PALETTE["border"],
            text_color=PALETTE["text_hi"])
        model_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        pull_status = ctk.StringVar(value="Cloud models are provided by Ollama after sign-in; they do not need to be pulled.")
        pull_button = ctk.CTkButton(
            pull_row, text="Pull", width=74,
            command=lambda: self._pull_ollama_model(
                win, model_entry.get(), endpoint_entry, pull_status,
                pull_button, render_ollama_models),
            fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"])
        pull_button.pack(side="right")
        ctk.CTkLabel(win, textvariable=pull_status, anchor="w", justify="left",
                     wraplength=500, text_color=PALETTE["text_mid"],
                     font=ctk.CTkFont(size=FONT_SCALE["small"])
                     ).pack(fill="x", padx=28, pady=(8, 4))

        footer = ctk.CTkFrame(win, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=24, pady=18)
        ctk.CTkButton(footer, text="Install Ollama",
                      command=lambda: webbrowser.open("https://ollama.com/download"),
                      fg_color="transparent", hover_color=PALETTE["border"],
                      text_color=PALETTE["text_mid"], width=130
                      ).pack(side="left")
        ctk.CTkButton(footer, text="Create Ollama Account",
                  command=lambda: webbrowser.open("https://ollama.com/signup"),
                  fg_color="transparent", hover_color=PALETTE["border"],
                  text_color=PALETTE["text_mid"], width=155
                  ).pack(side="left", padx=4)
        ctk.CTkButton(footer, text="Close", command=win.destroy,
                      fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"],
                      text_color=PALETTE["text_hi"], width=100
                      ).pack(side="right")

        self._check_ollama_connection(
            win, status_var, endpoint_entry, render_ollama_models)

    def _select_model(self, model_name, purpose):
        variable = self.model_var if purpose == "inference" else self.map_model_var
        setting = "inference_model" if purpose == "inference" else "map_model"
        variable.set(model_name)
        self.db.set_setting(setting, model_name)
        self.set_status(f"{model_name} selected for {purpose}", "idle")

    def _check_ollama_connection(self, window, status_var, endpoint_entry,
                                 render_models):
        try:
            self.model_mgr.configure("ollama", endpoint_entry.get())
        except Exception as exc:
            status_var.set(f"Invalid Ollama endpoint: {exc}")
            return
        self._set_setup_status(
            window, status_var, f"Checking {self.model_mgr.endpoint}…")

        def listed(models, error):
            self._on_models_refreshed(models, error)
            def update():
                if not window.winfo_exists():
                    return
                render_models(models)
                if error:
                    status_var.set(
                        f"Could not connect to {self.model_mgr.endpoint}: {error}. "
                        "For authenticated remote servers, set the API key in Provider settings.")
                else:
                    cloud_count = sum("cloud" in name.lower() for name in models)
                    status_var.set(
                        f"Connected to {self.model_mgr.endpoint}. "
                        f"{len(models)} model(s) available; {cloud_count} cloud model(s).")
            self.after(0, update)

        self.model_mgr.refresh(callback=listed)

    def _start_ollama_sign_in(self, window, status_var):
        executable = shutil.which("ollama")
        if not executable:
            webbrowser.open("https://ollama.com/download")
            self._set_setup_status(
                window, status_var,
                "Ollama CLI was not found. Install Ollama, then reopen this setup.")
            return
        try:
            if os.name == "nt":
                subprocess.Popen(
                    [executable], creationflags=subprocess.CREATE_NEW_CONSOLE)
            else:
                subprocess.Popen([executable])
            self._set_setup_status(
                window, status_var,
                "Ollama opened. Complete its sign-in prompts, then check and refresh models.")
        except OSError as exc:
            self._set_setup_status(
                window, status_var, f"Could not start Ollama: {exc}")

    def _pull_ollama_model(self, window, model_name, endpoint_entry,
                           status_var, button, render_models):
        model_name = model_name.strip()
        if not model_name:
            self._set_setup_status(window, status_var, "Enter a model name to pull.")
            return
        try:
            self.model_mgr.configure("ollama", endpoint_entry.get())
        except Exception as exc:
            self._set_setup_status(window, status_var, f"Invalid Ollama endpoint: {exc}")
            return

        button.configure(state="disabled")
        self._set_setup_status(window, status_var, f"Starting download for {model_name}…")

        def pull():
            try:
                for progress in self.model_mgr.pull(model_name, stream=True):
                    status = getattr(progress, "status", None)
                    completed = getattr(progress, "completed", None)
                    total = getattr(progress, "total", None)
                    if isinstance(progress, dict):
                        status = progress.get("status", status)
                        completed = progress.get("completed", completed)
                        total = progress.get("total", total)
                    message = status or f"Downloading {model_name}…"
                    if total and completed is not None:
                        message += f" ({int(completed * 100 / total)}%)"
                    self._set_setup_status(window, status_var, message)
                self._set_setup_status(
                    window, status_var, f"{model_name} is ready.")
                def refreshed(models, error):
                    self._on_models_refreshed(models, error)
                    self.after(0, lambda: render_models(models))
                self.model_mgr.refresh(callback=refreshed)
            except Exception as exc:
                self._set_setup_status(
                    window, status_var, f"Model pull failed: {exc}")
            finally:
                self.after(0, lambda: button.configure(state="normal")
                           if button.winfo_exists() else None)

        threading.Thread(target=pull, daemon=True).start()

    # --------------------------------------------------------
    # WORKSPACE INTERACTIVITY (7F)
    # --------------------------------------------------------
    def _listen_to_mic(self, input_widget, button=None):
        if sr is None:
            messagebox.showerror("Voice Input", "SpeechRecognition is not installed.")
            return

        if self._voice_busy:
            return
        restart_wake = (self.voice_wake_enabled and self._wake_thread is not None
                        and self._wake_thread.is_alive())
        wake_thread = self._wake_thread
        if restart_wake:
            self._wake_stop_event.set()
        self._voice_busy = True
        if button:
            button.configure(state="disabled", text="…")
        self.set_status("Preparing microphone…", "busy")

        def _do_listen():
            recognizer = sr.Recognizer()
            status = "Idle"
            level = "idle"
            try:
                if restart_wake and wake_thread is not threading.current_thread():
                    wake_thread.join(timeout=2)
                with sr.Microphone() as source:
                    self.after(0, lambda: self.set_status(
                        "Calibrating microphone…", "busy"))
                    recognizer.adjust_for_ambient_noise(source, duration=0.4)
                    self.after(0, lambda: self.set_status(
                        "Listening… speak now", "busy"))
                    audio = recognizer.listen(
                        source, timeout=5, phrase_time_limit=20)
                text = recognizer.recognize_google(
                    audio, language=self.voice_language)
                self.after(0, lambda: self._insert_voice_text(input_widget, text))
                status = "Voice input ready"
            except sr.WaitTimeoutError:
                status, level = "No speech detected", "error"
            except sr.UnknownValueError:
                status, level = "Could not understand audio", "error"
            except sr.RequestError as exc:
                status, level = f"Speech service unavailable: {exc}", "error"
            except OSError as exc:
                status, level = f"Microphone unavailable: {exc}", "error"
            except Exception as exc:
                status, level = f"Voice input failed: {exc}", "error"
            finally:
                self._voice_busy = False
                self.after(0, lambda: self._finish_voice_input(
                    button, status, level, restart_wake))

        threading.Thread(target=_do_listen, daemon=True).start()

    def _insert_voice_text(self, input_widget, text):
        try:
            if isinstance(input_widget, ctk.CTkTextbox):
                current = input_widget.get("1.0", "end-1c").strip()
                input_widget.delete("1.0", "end")
                input_widget.insert("end", current + ("\n" if current else "") + text)
            else:
                current = input_widget.get()
                input_widget.delete(0, "end")
                input_widget.insert(0, current + (" " if current else "") + text)
            input_widget.focus_set()
        except tk.TclError:
            pass

    def _finish_voice_input(self, button, status, level, restart_wake=False):
        if button and button.winfo_exists():
            button.configure(state="normal", text="🎤")
        self.set_status(status, level)
        if restart_wake and self.voice_wake_enabled:
            self._start_wake_listener()

    def _update_wake_button(self):
        if not hasattr(self, "_voice_toggle_btn"):
            return
        active = self.voice_wake_enabled
        self._voice_toggle_btn.configure(
            text=f"◉ Wake {'ON' if active else 'OFF'}",
            bg=PALETTE["accent_glow"] if active else PALETTE["bg_card"],
            fg=PALETTE["text_hi"] if active else PALETTE["text_mid"],
            activebackground=PALETTE["accent"],
            activeforeground="#ffffff",
            cursor="hand2")

    def _toggle_wake_listener(self):
        self.voice_wake_enabled = not self.voice_wake_enabled
        self.db.set_setting("voice_wake_enabled", "1" if self.voice_wake_enabled else "0")
        self._update_wake_button()
        if self.voice_wake_enabled:
            self._start_wake_listener()
        else:
            self._wake_stop_event.set()
            self.set_status("Wake phrase disabled", "idle")

    def _start_wake_listener(self):
        if not self.voice_wake_enabled or sr is None:
            return
        if self._voice_busy:
            self.after(400, self._start_wake_listener)
            return
        if self._wake_thread and self._wake_thread.is_alive():
            if self._wake_stop_event.is_set():
                self.after(250, self._start_wake_listener)
            return
        self._wake_stop_event.clear()
        phrase = self.voice_wake_phrase.strip() or "Ok Chacha"
        self._wake_thread = threading.Thread(
            target=self._wake_listener_loop, args=(phrase,), daemon=True)
        self._wake_thread.start()
        self.set_status(f"Listening for '{phrase}'", "idle")
        self._update_wake_button()

    def _wake_listener_loop(self, phrase):
        recognizer = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.35)
                while not self._wake_stop_event.is_set():
                    try:
                        audio = recognizer.listen(
                            source, timeout=1, phrase_time_limit=6)
                    except sr.WaitTimeoutError:
                        continue
                    try:
                        heard = recognizer.recognize_google(
                            audio, language=self.voice_language)
                    except sr.UnknownValueError:
                        continue
                    except sr.RequestError as exc:
                        self.after(0, lambda error=str(exc): self.set_status(
                            f"Wake recognition unavailable: {error}", "error"))
                        if self._wake_stop_event.wait(3):
                            break
                        continue

                    match = self._match_wake_phrase(heard, phrase)
                    if not match:
                        continue
                    request = heard[match.end():].strip(" ,.!?;:")
                    self.after(0, lambda: self._on_wake_detected())
                    if request:
                        self.after(0, lambda value=request:
                                   self._handle_voice_utterance(value))
                        continue

                    self.after(0, lambda: self._set_voice_panel_status(
                        "Wake phrase heard. Listening for your request…"))
                    try:
                        request_audio = recognizer.listen(
                            source, timeout=8, phrase_time_limit=20)
                        request = recognizer.recognize_google(
                            request_audio, language=self.voice_language).strip()
                        if request:
                            self.after(0, lambda value=request:
                                       self._handle_voice_utterance(value))
                    except sr.WaitTimeoutError:
                        self.after(0, lambda: self._set_voice_panel_status(
                            "Ready. Speak or type a request."))
                    except sr.UnknownValueError:
                        self.after(0, lambda: self._set_voice_panel_status(
                            "I could not understand that. Try again."))
                    except sr.RequestError as exc:
                        self.after(0, lambda error=str(exc):
                                   self._set_voice_panel_status(
                                       f"Speech service unavailable: {error}"))
                        if self._wake_stop_event.wait(3):
                            break
        except Exception as exc:
            if not self._wake_stop_event.is_set():
                self.after(0, lambda error=str(exc): self.set_status(
                    f"Wake listener stopped: {error}", "error"))
                if not self._wake_stop_event.wait(4):
                    self.after(0, self._start_wake_listener)
        finally:
            if not self._wake_stop_event.is_set():
                self.after(0, self._update_wake_button)

    @staticmethod
    def _match_wake_phrase(heard, phrase):
        candidates = [phrase]
        compact = re.sub(r"[^a-z0-9]", "", phrase.lower())
        if compact in ("okchacha", "okaychacha"):
            candidates.extend(["okay chacha", "ok cha cha", "okay cha cha",
                               "chacha", "cha cha"])
        for candidate in candidates:
            words = re.findall(r"[a-z0-9]+", candidate.lower())
            if not words:
                continue
            pattern = r"(?<!\w)" + r"[\W_]+".join(
                re.escape(word) for word in words) + r"(?!\w)"
            match = re.search(pattern, heard, re.IGNORECASE)
            if match:
                return match
        return None

    def _on_wake_detected(self):
        self._open_voice_panel(route=self.voice_wake_target, source="wake")
        self._set_voice_panel_status(
            f"'{self.voice_wake_phrase}' heard. Speak or edit your request.")

    def _handle_voice_utterance(self, utterance):
        utterance = utterance.strip()
        if not utterance:
            return
        self._open_voice_panel(route=self.voice_wake_target, source="wake")
        lowered = utterance.lower().strip(" .,!?")

        if re.fullmatch(r"(?:please\s+)?(?:read|speak)(?:\s+it)?\s+aloud", lowered):
            self._read_last_response()
            return
        if re.fullmatch(r"(?:please\s+)?(?:stop|cancel)\s+"
                        r"(?:speaking|reading|voice)", lowered):
            self._stop_speech()
            return
        if re.fullmatch(r"(?:please\s+)?(?:cancel|clear|forget that)", lowered):
            self._set_voice_panel_transcript("")
            self._set_voice_panel_status("Draft cleared.")
            return

        execute_match = re.match(
            r"^(?:please\s+)?(?:execute|run|initialize)(?:\s+(.*))?$",
            utterance, re.IGNORECASE)
        if execute_match:
            request = (execute_match.group(1) or "").strip()
            if not request and self._voice_panel_entry:
                request = self._voice_panel_entry.get().strip()
            if request:
                self._route_voice_request(
                    request, "Agent Directives", execute=True)
            else:
                self._set_voice_panel_status(
                    "Say 'execute' with a request, or dictate a request first.")
            return

        route_match = re.match(
            r"^(?:please\s+)?(?:send|put|route)\s+"
            r"(?:(?:it|that|this|the request)\s+)?to\s+(?:the\s+)?"
            r"(agent\s+directives?|directives?|workspace(?:\s+chat)?)"
            r"(?:\s+(?:saying\s+)?(.+))?$",
            utterance, re.IGNORECASE)
        if route_match:
            route = ("Agent Directives" if "directive" in
                     route_match.group(1).lower() else "Workspace Chat")
            request = (route_match.group(2) or "").strip()
            if not request and self._voice_panel_entry:
                request = self._voice_panel_entry.get().strip()
            if request:
                self._route_voice_request(request, route)
                self._set_voice_panel_status(
                    f"Request sent to {route}.")
            else:
                if self._voice_panel_route:
                    self._voice_panel_route.set(route)
                self._set_voice_panel_status(
                    f"Ready. The next request will go to {route}.")
            return

        payload_route = re.match(
            r"^(?:please\s+)?send\s+(.+?)\s+to\s+(?:the\s+)?"
            r"(agent\s+directives?|directives?|workspace(?:\s+chat)?)$",
            utterance, re.IGNORECASE)
        if payload_route:
            route = ("Agent Directives" if "directive" in
                     payload_route.group(2).lower() else "Workspace Chat")
            self._route_voice_request(payload_route.group(1).strip(), route)
            self._set_voice_panel_status(f"Request sent to {route}.")
            return

        self._set_voice_panel_transcript(utterance)

    def _set_voice_panel_status(self, text):
        if self._voice_panel_status and self._voice_panel_status.winfo_exists():
            self._voice_panel_status.configure(text=text)

    def _set_voice_panel_transcript(self, text):
        if self._voice_panel_entry and self._voice_panel_entry.winfo_exists():
            self._voice_panel_entry.delete(0, "end")
            self._voice_panel_entry.insert(0, text)
            self._voice_panel_entry.focus_set()
            self._set_voice_panel_status("Transcript ready. Review, edit, then send.")

    def _open_voice_panel(self, route="Agent Directives", capture=False,
                          source="manual"):
        if self._voice_panel and self._voice_panel.winfo_exists():
            self._voice_panel.lift()
            if self._voice_panel_route:
                self._voice_panel_route.set(route)
            if capture:
                self._listen_to_mic(self._voice_panel_entry,
                                    self._voice_panel_mic_button)
            return

        panel = ctk.CTkToplevel(self)
        self._voice_panel = panel
        panel.title("LAYA Voice")
        panel.geometry("470x292")
        panel.resizable(False, False)
        panel.configure(fg_color=PALETTE["bg_panel"])
        panel.attributes("-topmost", True)
        panel.grid_columnconfigure(0, weight=1)
        panel.protocol("WM_DELETE_WINDOW", panel.destroy)

        header = ctk.CTkFrame(panel, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 2))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Voice request",
                     font=ctk.CTkFont(size=FONT_SCALE["h2"], weight="bold"),
                     text_color=PALETTE["text_hi"]).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(header, text="×", width=28, height=26,
                      fg_color="transparent", hover_color=PALETTE["danger"],
                      command=panel.destroy).grid(row=0, column=1, sticky="e")

        self._voice_panel_status = ctk.CTkLabel(
            panel, text="Ready. Speak or type a request.", anchor="w",
            font=ctk.CTkFont(size=FONT_SCALE["small"]),
            text_color=PALETTE["text_mid"])
        self._voice_panel_status.grid(row=1, column=0, sticky="ew",
                                      padx=18, pady=(0, 8))
        entry_frame = ctk.CTkFrame(panel, fg_color=PALETTE["bg_input"],
                                   corner_radius=8,
                                   border_width=1,
                                   border_color=PALETTE["border"])
        entry_frame.grid(row=2, column=0, sticky="ew", padx=18, pady=2)
        entry_frame.grid_columnconfigure(0, weight=1)
        self._voice_panel_entry = ctk.CTkEntry(
            entry_frame, placeholder_text="Your request…", height=38,
            fg_color="transparent", border_width=0,
            text_color=PALETTE["text_hi"])
        self._voice_panel_entry.grid(row=0, column=0, sticky="ew", padx=8, pady=3)
        self._voice_panel_entry.bind("<Return>", lambda event: self._send_voice_request())

        route_frame = ctk.CTkFrame(panel, fg_color="transparent")
        route_frame.grid(row=3, column=0, sticky="ew", padx=18, pady=(8, 4))
        route_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(route_frame, text="Send to",
                     text_color=PALETTE["text_mid"],
                     font=ctk.CTkFont(size=FONT_SCALE["small"])
                     ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self._voice_panel_route = ctk.StringVar(value=route)
        ctk.CTkSegmentedButton(
            route_frame, values=["Agent Directives", "Workspace Chat"],
            variable=self._voice_panel_route,
            selected_color=PALETTE["accent"],
            selected_hover_color=PALETTE["accent_hot"],
            unselected_color=PALETTE["bg_card"],
            unselected_hover_color=PALETTE["border"],
            text_color=PALETTE["text_hi"]
        ).grid(row=0, column=1, sticky="ew")

        actions = ctk.CTkFrame(panel, fg_color="transparent")
        actions.grid(row=4, column=0, sticky="ew", padx=18, pady=(6, 4))
        for column in range(3):
            actions.grid_columnconfigure(column, weight=1)
        self._voice_panel_mic_button = ctk.CTkButton(
            actions, text="🎤 Listen", width=96, height=34,
            fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"],
            command=lambda: self._listen_to_mic(
                self._voice_panel_entry, self._voice_panel_mic_button))
        self._voice_panel_mic_button.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            actions, text="Execute", width=92, height=34,
            fg_color=PALETTE["accent_glow"], hover_color=PALETTE["accent"],
            command=self._execute_voice_request
        ).grid(row=0, column=1, padx=3)
        ctk.CTkButton(
            actions, text="Send", width=92, height=34,
            fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
            command=self._send_voice_request
        ).grid(row=0, column=2, sticky="e")
        self._voice_panel_speech_button = ctk.CTkButton(
            panel, text="Read last response aloud", height=30,
            fg_color="transparent", hover_color=PALETTE["border"],
            text_color=PALETTE["text_mid"],
            command=self._read_last_response
        )
        self._voice_panel_speech_button.grid(
            row=5, column=0, sticky="e", padx=18, pady=(0, 8))
        panel.bind("<Destroy>", lambda event: self._clear_voice_panel(event, panel))
        panel.lift()
        self._voice_panel_entry.focus_set()
        if capture:
            panel.after(150, lambda: self._listen_to_mic(
                self._voice_panel_entry, self._voice_panel_mic_button))

    def _clear_voice_panel(self, event, panel):
        if event.widget is panel:
            self._voice_panel = None
            self._voice_panel_entry = None
            self._voice_panel_status = None
            self._voice_panel_route = None
            self._voice_panel_mic_button = None
            self._voice_panel_speech_button = None

    def _send_voice_request(self):
        if not self._voice_panel_entry or not self._voice_panel_entry.winfo_exists():
            return
        request = self._voice_panel_entry.get().strip()
        if not request:
            self._set_voice_panel_status("Say or type a request first.")
            return
        route = self._voice_panel_route.get()
        panel = self._voice_panel
        if panel and panel.winfo_exists():
            panel.destroy()
        self._route_voice_request(request, route)

    def _execute_voice_request(self):
        if not self._voice_panel_entry or not self._voice_panel_entry.winfo_exists():
            return
        request = self._voice_panel_entry.get().strip()
        if not request:
            self._set_voice_panel_status("Say or type an instruction to execute.")
            return
        panel = self._voice_panel
        if panel and panel.winfo_exists():
            panel.destroy()
        self._route_voice_request(request, "Agent Directives", execute=True)

    def _read_last_response(self):
        if self._last_agent_response:
            self._speak_text(self._last_agent_response)
        else:
            self._set_voice_panel_status("There is no response to read aloud yet.")

    def _route_voice_request(self, request, route, execute=False):
        if route == "Agent Directives":
            self.prompt_text.delete("1.0", "end")
            self.prompt_text.insert("1.0", request)
            self.prompt_text.focus_set()
            self.right_tabs.set("  💬 Workspace  ")
            self.set_status("Voice request placed in Agent Directives", "idle")
            if execute or self.voice_auto_run_directives:
                self.start_pipeline()
            return
        self.right_tabs.set("  💬 Workspace  ")
        self.chat_input.delete(0, "end")
        self.chat_input.insert(0, request)
        self.chat_input.focus_set()
        self.send_followup()

    def _on_close(self):
        self.voice_wake_enabled = False
        self._wake_stop_event.set()
        self._stop_speech()
        self.destroy()

    @staticmethod
    def _prepare_speech_text(text):
        speech = html.unescape(text)
        speech = re.sub(r"```.*?```", " Code block omitted. ", speech,
                        flags=re.DOTALL)
        speech = re.sub(r"`([^`]+)`", r"\1", speech)
        speech = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1", speech)
        speech = re.sub(r"https?://\S+|www\.\S+", "link omitted", speech)
        speech = re.sub(r"<br\s*/?>", "\n", speech, flags=re.IGNORECASE)
        speech = re.sub(r"</?(?:p|div|li|h[1-6])\b[^>]*>", "\n", speech,
                        flags=re.IGNORECASE)
        speech = re.sub(r"<[^>]+>", " ", speech)

        spoken_lines = []
        table_headers = None
        for raw_line in speech.splitlines():
            line = raw_line.strip()
            if not line:
                table_headers = None
                continue
            if "|" in line:
                cells = [cell.strip() for cell in line.strip("|").split("|")]
                if cells and all(re.fullmatch(r"[:\-\s]+", cell) for cell in cells):
                    continue
                if table_headers is None:
                    table_headers = cells
                    continue
                details = [
                    f"{table_headers[index] if index < len(table_headers) else 'Value'}: {cell}"
                    for index, cell in enumerate(cells) if cell]
                if details:
                    spoken_lines.append(". ".join(details))
                continue

            table_headers = None
            line = re.sub(r"^\s*#{1,6}\s*", "", line)
            line = re.sub(r"^\s*>\s*", "", line)
            line = re.sub(r"^\s*(?:[-*+]\s+|\d+[.)]\s+)", "", line)
            line = re.sub(r"\*\*(.+?)\*\*|__(.+?)__", r"\1\2", line)
            line = re.sub(r"\*(.+?)\*|_(.+?)_", r"\1\2", line)
            line = re.sub(r"~~(.+?)~~", r"\1", line)
            line = line.replace("—", ", ").replace("–", ", ")
            line = line.replace("…", ". ").replace("→", " leads to ")
            line = line.replace("|", ", ").replace(";", ", ")
            line = re.sub(r"[\U0001F000-\U0001FAFF]", "", line)
            line = re.sub(r"\s+", " ", line).strip()
            if line:
                spoken_lines.append(line)

        speech = ". ".join(spoken_lines)
        speech = re.sub(r"(?:\.\s*){2,}", ". ", speech)
        speech = re.sub(r"[,;:]\s*\.", ".", speech)
        speech = re.sub(r"\s+([,.;:!?])", r"\1", speech)
        speech = re.sub(r"([,;:])(?=\S)", r"\1 ", speech)
        return speech.strip()

    def _set_speech_controls(self, active, generation=None):
        if (generation is not None and
                generation != self._speech_generation):
            return
        self._speech_active = active
        if hasattr(self, "_speech_stop_btn") and self._speech_stop_btn.winfo_exists():
            self._speech_stop_btn.configure(
                state="normal" if active else "disabled")
        button = self._voice_panel_speech_button
        if button and button.winfo_exists():
            button.configure(
                text="Stop speaking" if active else "Read last response aloud",
                command=self._stop_speech if active else self._read_last_response)

    def _stop_speech(self):
        with self._speech_lock:
            if not self._speech_active:
                return
            self._speech_stop_event.set()
            self._speech_generation += 1
            generation = self._speech_generation
            self._speech_engine = None
        self._speech_active = False
        self._set_speech_controls(False, generation)
        self.set_status("Speech stopped", "idle")

    def _speak_text(self, text):
        if pyttsx3 is None:
            self.set_status("Text-to-speech is not installed", "error")
            return
        speech = self._prepare_speech_text(text)
        if not speech:
            return

        with self._speech_lock:
            self._speech_stop_event.set()
            self._speech_generation += 1
            generation = self._speech_generation
            stop_event = threading.Event()
            self._speech_stop_event = stop_event
            self._speech_active = True
            self._speech_engine = None
        self.after(0, lambda: self._set_speech_controls(True, generation))

        def speak():
            engine = None
            loop_started = False
            try:
                engine = pyttsx3.init()
                voices = engine.getProperty("voices") or []
                preferred_gender = self.voice_gender.lower()
                preferred_voice = next((
                    voice for voice in voices
                    if preferred_gender in str(getattr(voice, "gender", "")).lower()
                    or preferred_gender in str(getattr(voice, "name", "")).lower()),
                    None)
                if preferred_voice is not None:
                    engine.setProperty("voice", preferred_voice.id)
                engine.setProperty("rate", max(110, min(220, self.voice_rate)))
                engine.setProperty("volume", 0.95)
                with self._speech_lock:
                    if generation != self._speech_generation or stop_event.is_set():
                        return
                    self._speech_engine = engine
                engine.say(speech)
                engine.startLoop(False)
                loop_started = True
                self.after(0, lambda: self.set_status("Speaking…", "busy"))
                while not stop_event.is_set() and engine.isBusy():
                    engine.iterate()
                    stop_event.wait(0.015)
            except Exception as exc:
                error_message = f"Text-to-speech failed: {exc}"
                self.after(0, lambda: self.set_status(
                    error_message, "error"))
            finally:
                if engine is not None and loop_started:
                    try:
                        if stop_event.is_set():
                            engine.stop()
                        engine.endLoop()
                    except Exception:
                        pass
                with self._speech_lock:
                    if generation == self._speech_generation:
                        self._speech_engine = None
                        self._speech_active = False
                        self.after(0, lambda: self._set_speech_controls(
                            False, generation))
                        self.after(0, lambda: self.set_status("Idle", "idle")
                                   if not stop_event.is_set() else None)

        threading.Thread(target=speak, daemon=True).start()

    def _search_history_ui(self):
        win = ctk.CTkToplevel(self)
        win.title("Search Chat History")
        win.geometry("700x500")
        win.configure(fg_color=PALETTE["bg_panel"])
        win.grab_set()
        
        search_frame = ctk.CTkFrame(win, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=10)
        
        entry = ctk.CTkEntry(search_frame, placeholder_text="Search keyword...", fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        results_box = ctk.CTkTextbox(win, fg_color=PALETTE["bg_deep"], text_color=PALETTE["text_hi"], font=ctk.CTkFont(family="Consolas", size=12))
        results_box.pack(fill="both", expand=True, padx=10, pady=10)
        
        def _do_search(event=None):
            query = entry.get().strip()
            if not query: return
            res = self.db.search_history(query)
            results_box.configure(state="normal")
            results_box.delete("0.0", "end")
            if not res:
                results_box.insert("end", "No results found.")
            else:
                for r in res:
                    results_box.insert("end", f"[{r['timestamp']}] {r['role'].upper()} (Session {r['session_id']}):\n{r['content']}\n{'-'*60}\n")
            results_box.configure(state="disabled")
            
        entry.bind("<Return>", _do_search)
        btn = ctk.CTkButton(search_frame, text="Search", command=_do_search, fg_color=PALETTE["accent"])
        btn.pack(side="right")

    def _scroll_chat(self, units):
        try:
            self.chat_scroll._parent_canvas.yview_scroll(units, "units")
        except Exception:
            pass

    def _smooth_scroll_bottom(self):
        try:
            self.chat_scroll._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass

    def _setup_scroll_bindings(self):
        """Bind mouse wheel scrolling for workspace and general chat panels."""
        scroll_frames = (self.chat_scroll, self.general_scroll)
        for sf in scroll_frames:
            try:
                sf._parent_canvas.unbind("<MouseWheel>")
                sf.unbind("<MouseWheel>")
            except Exception:
                pass

        def _scroll_target(widget):
            while widget is not None:
                for scroll_frame in scroll_frames:
                    if widget in (scroll_frame, scroll_frame._parent_canvas,
                                  scroll_frame._parent_frame):
                        return scroll_frame._parent_canvas
                widget = getattr(widget, "master", None)
            return None

        def _scroll_by(event, units):
            canvas = _scroll_target(event.widget)
            if canvas is None:
                return
            canvas.yview_scroll(units, "units")
            return "break"

        def _on_mousewheel(event):
            delta = getattr(event, "delta", 0)
            units = int(-delta / 120)
            if units == 0 and delta:
                units = -1 if delta > 0 else 1
            return _scroll_by(event, units) if units else None

        self.bind_all("<MouseWheel>", _on_mousewheel, add="+")
        self.bind_all("<Button-4>", lambda e: _scroll_by(e, -1), add="+")
        self.bind_all("<Button-5>", lambda e: _scroll_by(e, 1), add="+")

    def _show_text_context_menu(self, event):
        widget = event.widget
        if widget.winfo_class() not in ("Entry", "Text"):
            return
        editable = widget.cget("state") not in ("disabled", "readonly")
        menu = tk.Menu(self, tearoff=0, bg=PALETTE["bg_card"],
                       fg=PALETTE["text_hi"],
                       activebackground=PALETTE["accent"],
                       activeforeground="#ffffff", bd=0, relief="flat")

        def copy_selection():
            try:
                if widget.winfo_class() == "Text":
                    value = widget.get("sel.first", "sel.last")
                else:
                    value = widget.get()[widget.index("sel.first"):
                                          widget.index("sel.last")]
            except tk.TclError:
                return
            self.clipboard_clear()
            self.clipboard_append(value)

        def cut_selection():
            try:
                copy_selection()
                if widget.winfo_class() == "Text":
                    widget.delete("sel.first", "sel.last")
                else:
                    widget.delete("sel.first", "sel.last")
            except tk.TclError:
                return

        def select_all():
            if widget.winfo_class() == "Text":
                widget.tag_add("sel", "1.0", "end-1c")
            else:
                widget.selection_range(0, "end")
            widget.icursor("end")
            widget.focus_set()

        def paste():
            try:
                value = self.clipboard_get()
                if widget.winfo_class() == "Text":
                    if widget.tag_ranges("sel"):
                        widget.delete("sel.first", "sel.last")
                    widget.insert("insert", value)
                else:
                    if widget.selection_present():
                        widget.delete("sel.first", "sel.last")
                    widget.insert("insert", value)
            except tk.TclError:
                return

        def copy_all():
            value = (widget.get("1.0", "end-1c")
                     if widget.winfo_class() == "Text" else widget.get())
            self.clipboard_clear()
            self.clipboard_append(value)

        def clear_widget():
            widget.delete("1.0" if widget.winfo_class() == "Text" else 0,
                          "end")

        menu.add_command(label="Cut", state="normal" if editable else "disabled",
                         command=cut_selection)
        menu.add_command(label="Copy Selection", command=copy_selection)
        menu.add_command(label="Paste", state="normal" if editable else "disabled",
                         command=paste)
        menu.add_separator()
        menu.add_command(label="Copy All", command=copy_all)
        menu.add_command(label="Select All", command=select_all)
        menu.add_command(label="Clear", state="normal" if editable else "disabled",
                         command=clear_widget)
        if widget.winfo_class() == "Text" and editable:
            menu.add_separator()
            menu.add_command(label="Undo", command=lambda: widget.edit_undo())
            menu.add_command(label="Redo", command=lambda: widget.edit_redo())
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def _update_msg_count(self):
        self._msg_count += 1
        self._session_lbl.configure(
            text=f"SESSION  {self.session_id}  ·  {self._msg_count} messages")

    # --------------------------------------------------------
    # SETTINGS MODAL
    # --------------------------------------------------------
    def open_settings(self):
        win = ctk.CTkToplevel(self)
        win.title("Preferences")
        win.geometry("600x760")
        win.minsize(520, 600)
        win.configure(fg_color=PALETTE["bg_panel"])
        win.grab_set()

        ctk.CTkLabel(win, text="Preferences",
                     font=ctk.CTkFont(size=FONT_SCALE["h1"], weight="bold"),
                     text_color=PALETTE["text_hi"]).pack(anchor="w", padx=24,
                                                         pady=(16, 8))
        body = ctk.CTkScrollableFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 4))

        for lbl, var in [("Inference Model", self.model_var),
                          ("Map-Reduce Model", self.map_model_var)]:
            ctk.CTkLabel(body, text=lbl,
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["text_mid"]).pack(anchor="w", padx=8, pady=(8, 2))
            ctk.CTkEntry(body, textvariable=var,
                         fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"],
                         border_color=PALETTE["border"],
                         font=ctk.CTkFont(size=FONT_SCALE["body"])).pack(fill="x", padx=8)

        tesseract_var = None
        if 'pytesseract' in sys.modules:
            ctk.CTkLabel(body, text="Tesseract Path",
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["text_mid"]).pack(anchor="w", padx=8, pady=(12, 2))
            tesseract_var = ctk.StringVar(
                value=pytesseract.pytesseract.tesseract_cmd)
            ctk.CTkEntry(body, textvariable=tesseract_var,
                         fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"],
                         border_color=PALETTE["border"],
                         font=ctk.CTkFont(size=FONT_SCALE["small"])).pack(fill="x", padx=8)

        ctk.CTkLabel(body, text="Theme",
                     font=ctk.CTkFont(size=FONT_SCALE["small"]),
                     text_color=PALETTE["text_mid"]).pack(anchor="w", padx=8, pady=(12, 2))
        tvar2 = ctk.StringVar(value=self._active_theme)
        ctk.CTkComboBox(body, variable=tvar2, values=list(THEMES.keys()),
                        fg_color=PALETTE["bg_input"],
                        border_color=PALETTE["border"],
                        button_color=PALETTE["accent"],
                        text_color=PALETTE["text_hi"],
                        dropdown_fg_color=PALETTE["bg_card"],
                        dropdown_text_color=PALETTE["text_hi"],
                        ).pack(fill="x", padx=8)

        ctk.CTkLabel(body, text="VOICE ASSISTANT",
                     font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                     text_color=PALETTE["accent"]).pack(anchor="w", padx=8,
                                                        pady=(18, 4))
        wake_enabled_var = tk.BooleanVar(value=self.voice_wake_enabled)
        ctk.CTkCheckBox(
            body, text="Enable always-listening wake phrase",
            variable=wake_enabled_var, fg_color=PALETTE["accent"],
            hover_color=PALETTE["accent_hot"],
            text_color=PALETTE["text_hi"]).pack(anchor="w", padx=8, pady=5)

        def add_voice_field(label, variable, values=None):
            ctk.CTkLabel(body, text=label,
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["text_mid"]).pack(anchor="w", padx=8,
                                                              pady=(8, 2))
            if values:
                ctk.CTkComboBox(
                    body, variable=variable, values=values,
                    fg_color=PALETTE["bg_input"],
                    border_color=PALETTE["border"],
                    button_color=PALETTE["accent"],
                    text_color=PALETTE["text_hi"],
                    dropdown_fg_color=PALETTE["bg_card"],
                    dropdown_text_color=PALETTE["text_hi"]
                ).pack(fill="x", padx=8)
            else:
                ctk.CTkEntry(
                    body, textvariable=variable,
                    fg_color=PALETTE["bg_input"],
                    border_color=PALETTE["border"],
                    text_color=PALETTE["text_hi"]
                ).pack(fill="x", padx=8)

        routes = ["Agent Directives", "Workspace Chat"]
        wake_phrase_var = ctk.StringVar(value=self.voice_wake_phrase)
        voice_language_var = ctk.StringVar(value=self.voice_language)
        manual_route_var = ctk.StringVar(value=self.voice_manual_target)
        wake_route_var = ctk.StringVar(value=self.voice_wake_target)
        add_voice_field("Wake phrase", wake_phrase_var)
        add_voice_field("Speech recognition language", voice_language_var,
                ["en-US", "en-GB", "en-IN", "hi-IN", "fr-FR",
                 "de-DE", "es-ES", "it-IT", "ja-JP", "zh-CN"])
        add_voice_field("Dictate button sends to", manual_route_var, routes)
        add_voice_field("Wake phrase sends to", wake_route_var, routes)

        auto_run_var = tk.BooleanVar(value=self.voice_auto_run_directives)
        ctk.CTkCheckBox(
            body, text="Run analysis after sending a voice directive",
            variable=auto_run_var, fg_color=PALETTE["accent"],
            hover_color=PALETTE["accent_hot"],
            text_color=PALETTE["text_hi"]).pack(anchor="w", padx=8, pady=(12, 4))
        speak_replies_var = tk.BooleanVar(value=self.voice_speak_replies)
        voice_gender_var = ctk.StringVar(value=self.voice_gender)
        add_voice_field("Read-aloud voice", voice_gender_var,
                ["Female", "Male"])
        voice_rate_var = tk.DoubleVar(value=self.voice_rate)
        rate_frame = ctk.CTkFrame(body, fg_color="transparent")
        rate_frame.pack(fill="x", padx=8, pady=(10, 2))
        rate_frame.grid_columnconfigure(0, weight=1)
        rate_label = ctk.CTkLabel(
            rate_frame, text=f"Speaking pace: {self.voice_rate}",
            text_color=PALETTE["text_mid"],
            font=ctk.CTkFont(size=FONT_SCALE["small"]))
        rate_label.grid(row=0, column=0, sticky="w")
        ctk.CTkSlider(
            rate_frame, from_=120, to=210, variable=voice_rate_var,
            command=lambda value: rate_label.configure(
                text=f"Speaking pace: {int(float(value))}"),
            progress_color=PALETTE["accent"],
            button_color=PALETTE["accent_hot"],
            button_hover_color=PALETTE["text_hi"], height=14
        ).grid(row=1, column=0, sticky="ew", pady=(2, 0))
        ctk.CTkCheckBox(
            body, text="Read AI replies aloud",
            variable=speak_replies_var, fg_color=PALETTE["accent"],
            hover_color=PALETTE["accent_hot"],
            text_color=PALETTE["text_hi"]).pack(anchor="w", padx=8, pady=4)
        ctk.CTkLabel(
            body,
            text="Wake-word detection uses online Google speech recognition and keeps the microphone active while enabled. Audio is sent for transcription. Disable it at any time with the Wake button.",
            wraplength=500, justify="left",
            font=ctk.CTkFont(size=FONT_SCALE["small"] - 1),
            text_color=PALETTE["text_lo"]).pack(fill="x", padx=8, pady=(8, 16))

        def save():
            self.db.set_setting("inference_model", self.model_var.get())
            self.db.set_setting("map_model", self.map_model_var.get())
            self.apply_theme(tvar2.get())
            if tesseract_var is not None:
                pytesseract.pytesseract.tesseract_cmd = tesseract_var.get()
            old_phrase = self.voice_wake_phrase
            old_enabled = self.voice_wake_enabled
            self.voice_wake_enabled = wake_enabled_var.get()
            self.voice_wake_phrase = wake_phrase_var.get().strip() or "Ok Chacha"
            self.voice_language = voice_language_var.get()
            self.voice_manual_target = manual_route_var.get()
            self.voice_wake_target = wake_route_var.get()
            self.voice_auto_run_directives = auto_run_var.get()
            self.voice_speak_replies = speak_replies_var.get()
            self.voice_gender = voice_gender_var.get()
            self.voice_rate = int(round(voice_rate_var.get()))
            for key, value in [
                ("voice_wake_enabled", self.voice_wake_enabled),
                ("voice_wake_phrase", self.voice_wake_phrase),
                ("voice_language", self.voice_language),
                ("voice_manual_target", self.voice_manual_target),
                ("voice_wake_target", self.voice_wake_target),
                ("voice_auto_run_directives", self.voice_auto_run_directives),
                ("voice_speak_replies", self.voice_speak_replies),
                ("voice_gender", self.voice_gender),
                ("voice_rate", self.voice_rate)]:
                self.db.set_setting(key, "1" if value is True else
                                    "0" if value is False else value)
            self._update_wake_button()
            if old_enabled != self.voice_wake_enabled or old_phrase != self.voice_wake_phrase:
                self._wake_stop_event.set()
                if self.voice_wake_enabled:
                    self.after(300, self._start_wake_listener)
            elif not self.voice_wake_enabled:
                self._wake_stop_event.set()
            win.destroy()

        ctk.CTkButton(win, text="Save & Close", command=save,
                      fg_color=PALETTE["accent"],
                      hover_color=PALETTE["accent_hot"],
                      font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
                      height=40, width=180, corner_radius=8).pack(pady=(4, 14))

    # --------------------------------------------------------
    # FILE OPERATIONS
    # --------------------------------------------------------
    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[
            ("Supported Documents and Data",
             "*.doc;*.docx;*.pdf;*.xlsx;*.csv;*.txt;*.png;*.jpg;*.jpeg"),
            ("All Files", "*.*")])
        for f in files:
            if f not in self.selected_files:
                self.selected_files.append(f)
                card = FileCard(self.file_list_frame, f,
                                on_preview=self.load_preview,
                                on_remove=self._remove_file)
                card.pack(fill="x", pady=2, padx=2)
                self.db.log_file_access(self.session_id, f, "add")
        self._update_file_badge()
        if self.selected_files and not self.output_path_var.get():
            self.output_path_var.set(
                os.path.join(os.path.dirname(self.selected_files[0]),
                             "Agentic_Report" +
                             OUTPUT_FORMATS.get(self.output_format_var.get(), ".docx")))

    def _remove_file(self, fp, card):
        if fp in self.selected_files:
            self.selected_files.remove(fp)
            self.db.log_file_access(self.session_id, fp, "remove")
        card.destroy()
        self._update_file_badge()

    def _update_file_badge(self):
        n = len(self.selected_files)
        self._file_badge.configure(text=f"{n} file{'s' if n!=1 else ''}")

    def clear_files(self):
        for fp in self.selected_files:
            self.db.log_file_access(self.session_id, fp, "clear")
        self.selected_files.clear()
        for w in self.file_list_frame.winfo_children():
            w.destroy()
        self._update_file_badge()

    def load_preview(self, fp):
        self.right_tabs.set("  🔍 Preview  ")
        self.preview_box.configure(state="normal")
        self.preview_box.delete("0.0", "end")
        self.preview_box.insert("end", f"Extracting {os.path.basename(fp)}…\n")
        self.preview_box.configure(state="disabled")
        self.db.log_file_access(self.session_id, fp, "preview")
        def _do():
            try:
                s = extract_dual_stream_from_file(fp)
                raw = s["raw"]
                content = (f"═══ {os.path.basename(fp)} ═══\n\n"
                           f"── RAW STREAM (first 4000 chars) ──\n{raw[:4000]}")
                if len(raw) > 4000:
                    content += "\n\n…[TRUNCATED]…"
            except Exception as e:
                content = f"Error:\n{e}"
            def _upd():
                self.preview_box.configure(state="normal")
                self.preview_box.delete("0.0", "end")
                self.preview_box.insert("end", content)
                self.preview_box.configure(state="disabled")
            self.after(0, _upd)
        threading.Thread(target=_do, daemon=True).start()

    def browse_output(self):
        format_name = self.output_format_var.get()
        extension = OUTPUT_FORMATS.get(format_name, ".docx")
        fn = filedialog.asksaveasfilename(
            defaultextension=extension,
            filetypes=[(f"{format_name} (*{extension})", f"*{extension}"),
                       ("All Files", "*.*")])
        if fn:
            root, existing_extension = os.path.splitext(fn)
            if existing_extension.lower() != extension:
                fn = root + extension
            self.output_path_var.set(fn)

    def _change_output_format(self, format_name):
        extension = OUTPUT_FORMATS.get(format_name, ".docx")
        self.db.set_setting("output_format", format_name)
        current_path = self.output_path_var.get().strip()
        if current_path:
            root, _ = os.path.splitext(current_path)
            self.output_path_var.set(root + extension)

    def _toggle_export(self):
        state = "normal" if self.export_var.get() else "disabled"
        self.output_entry.configure(state=state)
        self.output_format_combo.configure(state=state)
        self.browse_out_btn.configure(state=state)

    def _save_generated_output(self, content):
        if not self.export_var.get():
            return
        format_name = self.output_format_var.get()
        extension = OUTPUT_FORMATS.get(format_name, ".docx")
        output_path = self.output_path_var.get().strip()
        if not output_path:
            self.append_to_terminal("warn", "Output was not saved: no output path is set.")
            return
        root, current_extension = os.path.splitext(output_path)
        if current_extension.lower() != extension:
            output_path = root + extension
            self.output_path_var.set(output_path)

        try:
            if format_name in ("Text", "Markdown"):
                with open(output_path, "w", encoding="utf-8") as output_file:
                    output_file.write(content)
            elif format_name == "HTML":
                escaped = html.escape(content)
                html_doc = (
                    "<!doctype html><html><head><meta charset='utf-8'>"
                    "<title>LAYA generated document</title>"
                    "<style>body{font:16px/1.6 Segoe UI,Arial,sans-serif;"
                    "max-width:850px;margin:48px auto;padding:0 24px;color:#202532}"
                    "pre{white-space:pre-wrap;font:inherit}</style></head>"
                    f"<body><pre>{escaped}</pre></body></html>")
                with open(output_path, "w", encoding="utf-8") as output_file:
                    output_file.write(html_doc)
            elif format_name == "DOCX":
                document = Document()
                for line in content.splitlines():
                    trimmed = line.strip()
                    if not trimmed:
                        document.add_paragraph()
                    elif trimmed.startswith("### "):
                        document.add_heading(trimmed[4:], level=3)
                    elif trimmed.startswith("## "):
                        document.add_heading(trimmed[3:], level=2)
                    elif trimmed.startswith("# "):
                        document.add_heading(trimmed[2:], level=1)
                    elif re.match(r"^[-*+]\s+", trimmed):
                        document.add_paragraph(
                            re.sub(r"^[-*+]\s+", "", trimmed),
                            style="List Bullet")
                    elif re.match(r"^\d+[.)]\s+", trimmed):
                        document.add_paragraph(
                            re.sub(r"^\d+[.)]\s+", "", trimmed),
                            style="List Number")
                    else:
                        document.add_paragraph(trimmed)
                document.save(output_path)
            elif format_name == "PDF":
                from reportlab.lib.pagesizes import letter
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.units import inch
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

                styles = getSampleStyleSheet()
                styles.add(ParagraphStyle(
                    name="DocumentTitle", parent=styles["Title"],
                    alignment=0, spaceAfter=14))
                styles.add(ParagraphStyle(
                    name="DocumentBody", parent=styles["BodyText"],
                    leading=15, spaceAfter=7))
                story = []
                for line in content.splitlines():
                    trimmed = line.strip()
                    if not trimmed:
                        story.append(Spacer(1, 0.12 * inch))
                        continue
                    style = styles["DocumentBody"]
                    if trimmed.startswith("# "):
                        trimmed = trimmed[2:]
                        style = styles["DocumentTitle"]
                    elif trimmed.startswith("## "):
                        trimmed = trimmed[3:]
                        style = styles["Heading2"]
                    elif trimmed.startswith("### "):
                        trimmed = trimmed[4:]
                        style = styles["Heading3"]
                    elif re.match(r"^[-*+]\s+", trimmed):
                        trimmed = "-  " + re.sub(r"^[-*+]\s+", "", trimmed)
                    safe_text = html.escape(trimmed)
                    safe_text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", safe_text)
                    safe_text = re.sub(r"\*(.+?)\*", r"<i>\1</i>", safe_text)
                    story.append(Paragraph(safe_text, style))
                SimpleDocTemplate(
                    output_path, pagesize=letter,
                    rightMargin=0.75 * inch, leftMargin=0.75 * inch,
                    topMargin=0.75 * inch, bottomMargin=0.75 * inch
                ).build(story)
            self.append_to_terminal("ok", f"Generated document saved: {output_path}")
        except Exception as exc:
            self.append_to_terminal("err", f"Could not save generated document: {exc}")

    # --------------------------------------------------------
    # EXPORT & HISTORY MENU ACTIONS
    # --------------------------------------------------------
    def export_chat_txt(self):
        fn = filedialog.asksaveasfilename(defaultextension=".txt",
                                          filetypes=[("Text", "*.txt")])
        if not fn: return
        content = self.db.export_session_txt(self.session_id)
        with open(fn, "w", encoding="utf-8") as f:
            f.write(content)
        self.append_to_terminal("ok", f"Chat exported to {fn}")

    def export_chat_md(self):
        fn = filedialog.asksaveasfilename(defaultextension=".md",
                                          filetypes=[("Markdown", "*.md")])
        if not fn: return
        recs = self.db.get_session_history(self.session_id, limit=999)
        lines = [f"# Session {self.session_id}\n"]
        for r in recs:
            prefix = "**You**" if r["role"] == "user" else "**Agent**"
            lines.append(f"\n{prefix}:\n\n{r['content']}\n\n---")
        with open(fn, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        self.append_to_terminal("ok", f"Chat exported as Markdown to {fn}")

    def new_session(self):
        if messagebox.askyesno("New Session",
                               "Start a new session? Current context will be cleared."):
            self.session_id = str(uuid.uuid4())[:8]
            self.document_context = ""
            self.selected_files.clear()
            self._msg_count = 0
            for w in self.file_list_frame.winfo_children(): w.destroy()
            self._clear_chat()
            self._update_file_badge()
            self._session_lbl.configure(
                text=f"SESSION  {self.session_id}  ·  0 messages")
            self.db.upsert_session(self.session_id, 0, 0, self.model_var.get())
            self.append_to_terminal("info", f"New session started: {self.session_id}")

    def _clear_chat(self):
        for w in self.chat_scroll.winfo_children(): w.destroy()
        self._chat_row = 1
        self._stream_agent_frame = None
        self._stream_label = None
        self._stream_text_acc = ""
        self._typing_ind = TypingIndicator(self.chat_scroll)
        self._workspace_history.clear()

        if hasattr(self, 'general_scroll'):
            for w in self.general_scroll.winfo_children(): w.destroy()
            self._general_chat_row = 1
            self._general_stream_agent_frame = None
            self._general_stream_label = None
            self._general_stream_text_acc = ""
            self._general_typing_ind = TypingIndicator(self.general_scroll)
            self._general_history.clear()

    def _view_session_log(self):
        self._open_log_viewer("Session Chat Log",
            self.db.export_session_txt(self.session_id))

    def _view_file_log(self):
        with sqlite3.connect(self.db.db_path) as c:
            cur = c.execute(
                "SELECT timestamp,action,file_path,file_size FROM file_access_log "
                "WHERE session_id=? ORDER BY id", (self.session_id,))
            rows = cur.fetchall()
        text = "\n".join(f"[{r[0]}] {r[1]:8} {r[3]:>10}B  {r[2]}" for r in rows)
        self._open_log_viewer("File Access Log", text or "(empty)")

    def _view_model_log(self):
        with sqlite3.connect(self.db.db_path) as c:
            cur = c.execute(
                "SELECT timestamp,model_name,prompt_chars,response_chars,duration_ms "
                "FROM model_usage_log WHERE session_id=? ORDER BY id", (self.session_id,))
            rows = cur.fetchall()
        text = "\n".join(
            f"[{r[0]}] {r[1]} | prompt={r[2]}ch resp={r[3]}ch {r[4]}ms"
            for r in rows)
        self._open_log_viewer("Model Usage Log", text or "(empty)")

    def _open_log_viewer(self, title, content):
        win = ctk.CTkToplevel(self)
        win.title(title)
        win.geometry("760x500")
        win.configure(fg_color=PALETTE["bg_panel"])
        tb = ctk.CTkTextbox(win,
                            fg_color=PALETTE["bg_deep"],
                            text_color=PALETTE["text_hi"],
                            font=ctk.CTkFont(family="Consolas", size=11))
        tb.pack(fill="both", expand=True, padx=8, pady=8)
        tb.insert("0.0", content)
        tb.configure(state="disabled")

    def _show_about(self):
        messagebox.showinfo("About LAYA Agentic Studio v15",
            "LAYA Agentic Studio v15\n\n"
            "A self-healing multimodal AI workspace.\n"
            "Built on LAYA Decision Engine + Ollama.\n\n"
            "© 2026 — Local use only.")

    def _show_shortcuts(self):
        self._open_log_viewer("Keyboard Shortcuts",
            "Ctrl+N        New Session\n"
            "Ctrl+O        Add Files\n"
            "Ctrl+T        Cycle Theme\n"
            "Ctrl+Enter    Initialize Analysis\n"
            "Enter         Send follow-up\n"
            "Right-click   Context menu on agent bubbles\n"
        )

    def _zoom_in(self):
        FONT_SCALE["body"] = min(FONT_SCALE["body"] + 1, 20)
    def _zoom_out(self):
        FONT_SCALE["body"] = max(FONT_SCALE["body"] - 1, 9)

    # --------------------------------------------------------
    # TERMINAL
    # --------------------------------------------------------
    def append_to_terminal(self, level, text):
        icons = {"ok": "✅", "err": "❌", "warn": "⚠", "info": "⚙"}
        icon = icons.get(level, "⚙")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        msg = f"[{ts}] {icon}  {text}\n"
        def _do():
            self.terminal_log.configure(state="normal")
            self.terminal_log.insert("end", msg, level if level in ("ok","err","warn","info") else "info")
            self.terminal_log.see("end")
            self.terminal_log.configure(state="disabled")
        self.after(0, _do)

    def _clear_terminal(self):
        self.terminal_log.configure(state="normal")
        self.terminal_log.delete("0.0", "end")
        self.terminal_log.configure(state="disabled")

    def _copy_terminal(self):
        self.terminal_log.configure(state="normal")
        content = self.terminal_log.get("0.0", "end")
        self.terminal_log.configure(state="disabled")
        self.clipboard_clear()
        self.clipboard_append(content)

    # --------------------------------------------------------
    # CHAT HELPERS
    # --------------------------------------------------------
    def append_to_chat(self, sender, text, tag=None, chat_type="workspace"):
        if sender == "System":
            level = ("ok" if "✅" in text else "err" if "❌" in text
                     else "warn" if any(w in text for w in ("WARNING","WARN","warn")) else "info")
            self.append_to_terminal(level, text)
            return

        def _do():
            is_gen = (chat_type == "general")
            scroll_widget = self.general_scroll if is_gen else self.chat_scroll
            current_row = self._general_chat_row if is_gen else self._chat_row
            
            if sender == "User":
                if is_gen:
                    self._general_typing_ind.stop()
                    self._general_typing_ind.grid_remove()
                else:
                    self._hide_typing()
                
                b = MessageBubble(scroll_widget, "user", text,
                                  copy_callback=self._copy_to_clipboard)
                b.grid(row=current_row, column=0, sticky="ew", pady=2)
                
                if is_gen:
                    self._general_chat_row += 1
                else:
                    self._chat_row += 1
                    self._update_msg_count()
                
                self.after(50, lambda: scroll_widget._parent_canvas.yview_moveto(1.0))

            elif sender == "Agent" and not tag:
                if is_gen:
                    self._general_typing_ind.stop()
                    self._general_typing_ind.grid_remove()
                    self._general_stream_text_acc = ""
                else:
                    self._hide_typing()
                    self._stream_text_acc = ""
                
                frame = ctk.CTkFrame(scroll_widget,
                                     fg_color=PALETTE["bg_card"],
                                     corner_radius=12,
                                     border_width=1,
                                     border_color=PALETTE["border"])
                frame.grid(row=current_row, column=0, sticky="ew", pady=2)
                frame.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(frame, text="🤖",
                             font=ctk.CTkFont(size=14),
                             text_color=PALETTE["accent"]
                             ).grid(row=0, column=0, sticky="nw", padx=10, pady=(8, 2))
                
                lbl = ctk.CTkLabel(frame, text="",
                                   wraplength=740,
                                   justify="left",
                                   font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                   text_color=PALETTE["text_hi"])
                lbl.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
                
                if is_gen:
                    self._general_stream_agent_frame = frame
                    self._general_stream_label = lbl
                    self._general_chat_row += 1
                else:
                    self._stream_agent_frame = frame
                    self._stream_label = lbl
                    self._chat_row += 1

            elif sender == "Agent" and tag == "stream":
                if is_gen:
                    self._general_stream_text_acc += text
                    if self._general_stream_label:
                        self._general_stream_label.configure(text=self._general_stream_text_acc)
                else:
                    self._stream_text_acc += text
                    if self._stream_label:
                        self._stream_label.configure(text=self._stream_text_acc)
                self.after(30, lambda: scroll_widget._parent_canvas.yview_moveto(1.0))

        self.after(0, _do)

    def _finalise_agent_bubble(self, full_text, chat_type="workspace"):
        def _do():
            self._last_agent_response = full_text
            is_gen = (chat_type == "general")
            scroll_widget = self.general_scroll if is_gen else self.chat_scroll
            frame = self._general_stream_agent_frame if is_gen else self._stream_agent_frame
            row_idx = self._general_chat_row - 1 if is_gen else self._chat_row - 1

            if frame:
                try:
                    row_info = frame.grid_info()
                    row = row_info.get("row", row_idx)
                except Exception:
                    row = row_idx
                try:
                    frame.destroy()
                except Exception:
                    pass

                if is_gen:
                    self._general_stream_agent_frame = None
                    self._general_stream_label = None
                else:
                    self._stream_agent_frame = None
                    self._stream_label = None

                b = MessageBubble(scroll_widget, "agent", full_text,
                                  copy_callback=self._copy_to_clipboard,
                                  rerun_callback=self.execute_agent_code,
                                  quote_callback=lambda value: self._quote_to_input(
                                      value, chat_type),
                                  speak_callback=self._speak_text)
                b.grid(row=row, column=0, sticky="ew", pady=2)

                if not is_gen:
                    self._update_msg_count()
                if self.voice_speak_replies:
                    self._speak_text(full_text)
                self.after(60, lambda: scroll_widget._parent_canvas.yview_moveto(1.0))
        self.after(0, _do)

    def _show_typing(self):
        def _do():
            self._typing_ind.grid(row=self._chat_row, column=0,
                                  sticky="w", padx=16, pady=4)
            self._typing_ind.start()
            self.after(60, self._smooth_scroll_bottom)
        self.after(0, _do)

    def _hide_typing(self):
        self._typing_ind.stop()
        self._typing_ind.grid_remove()

    def _copy_to_clipboard(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.set_status("Copied!", "idle")
        self.after(2000, lambda: self.set_status("Idle", "idle"))

    def _quote_to_input(self, text, chat_type="workspace"):
        widget = self.general_input if chat_type == "general" else self.chat_input
        excerpt = " ".join(text.split())[:240]
        current = widget.get().strip()
        widget.delete(0, "end")
        widget.insert(0, f'{current} "{excerpt}" '.strip())
        widget.focus_set()

    # --------------------------------------------------------
    # PROCESSING STATE
    # --------------------------------------------------------
    def toggle_processing_state(self, processing: bool):
        self.is_processing = processing
        state = "disabled" if processing else "normal"
        self.send_btn.configure(state=state)
        if processing:
            self.run_canvas.configure(cursor="watch")
            self.progress_bar.grid(row=17, column=0, sticky="ew",
                                   padx=8, pady=(0, 2))
            self.progress_bar.start()
            self.set_status("Processing…", "busy")
        else:
            self.run_canvas.configure(cursor="")
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            if not self.is_executing:
                self.set_status("Idle", "idle")

    def update_temp_label(self, value):
        val = round(float(value), 2)
        if val <= 0.2:
            desc, color = "Strict", PALETTE["success"]
        elif val <= 0.5:
            desc, color = "Balanced", PALETTE["warning"]
        else:
            desc, color = "Creative", PALETTE["danger"]
        self.temp_label.configure(text=f"{val:.2f} — {desc}", text_color=color)
        self.db.set_setting("temperature", str(val))

    # --------------------------------------------------------
    # SELF-HEALING EXECUTION ENGINE
    # --------------------------------------------------------
    def _script_runtime_context(self):
        source_files = [os.path.abspath(path) for path in self.selected_files
                        if os.path.isfile(path)]
        pillow_version = getattr(sys.modules.get("PIL"), "__version__", "unknown")
        source_root = (os.path.dirname(source_files[0])
                       if source_files else os.getcwd())
        return (
            f"Generated scripts run on Windows with Python {sys.version.split()[0]} "
            f"and Pillow {pillow_version}. Working directory: {source_root}.\n"
            f"Selected source files (absolute paths): {json.dumps(source_files)}\n"
            "Use the absolute paths when opening source documents. Use Pillow "
            "ImageDraw.textbbox instead of removed textsize; ImageDraw.line does "
            "not support dash=, so draw dashed lines as segments."
        )

    def execute_agent_code(self, llm_response, max_retries=5):
        blocks = re.findall(r'```python\n(.*?)\n```', llm_response, re.DOTALL)
        if not blocks: return
        self.is_executing = True
        self.set_status("Executing generated script…", "busy")
        source_files = [os.path.abspath(path) for path in self.selected_files
                        if os.path.isfile(path)]
        threading.Thread(target=self._exec_thread,
                         args=(blocks, max_retries, source_files), daemon=True).start()

    def _exec_thread(self, blocks, max_retries, source_files=None):
        model = self.model_var.get()
        source_files = source_files or []
        working_directory = (os.path.dirname(source_files[0])
                             if source_files else os.getcwd())
        pillow_version = getattr(sys.modules.get("PIL"), "__version__", "unknown")
        try:
            for idx, code in enumerate(blocks):
                tmp = tempfile.mkdtemp()
                sp = os.path.join(tmp, f"agent_{self.session_id}_{idx}.py")
                cur = code
                ok = False; attempt = 1
                try:
                    while attempt <= max_retries and not ok:
                        self.append_to_terminal("info",
                            f"Script #{idx+1} attempt {attempt}/{max_retries}")
                        with open(sp, "w", encoding="utf-8") as f:
                            f.write(cur)
                        try:
                            t0 = time.time()
                            script_env = os.environ.copy()
                            script_env["LAYA_SOURCE_FILES"] = json.dumps(source_files)
                            res = subprocess.run(
                                [sys.executable, sp],
                                capture_output=True, text=True,
                                cwd=working_directory, env=script_env, timeout=60)
                            ms = int((time.time() - t0) * 1000)
                            self.db.log_model_usage(self.session_id, model,
                                                    len(cur), len(res.stdout), ms)
                            if res.returncode == 0:
                                self.append_to_terminal("ok",
                                    f"Script OK. {res.stdout.strip() or '(no output)'}")
                                ok = True
                            else:
                                err = res.stderr.strip()
                                self.append_to_terminal("err", f"Script failed:\n{err}")
                                if attempt < max_retries:
                                    self.append_to_terminal("warn", "Self-healing…")
                                    fix = self.model_mgr.chat(
                                        model=model,
                                        messages=[{"role": "user",
                                                   "content": (
                                                       f"Script failed:\n```\n{err}\n```\n"
                                                       f"Available source files: {json.dumps(source_files)}\n"
                                                       f"Runtime: Windows, Pillow {pillow_version}.\n"
                                                       f"Code:\n```python\n{cur}\n```\n"
                                                       "Fix all visible issues and preserve the script's purpose. "
                                                       "Use Pillow ImageDraw.textbbox (textsize was removed); "
                                                       "ImageDraw.line has no dash= parameter, so draw dashed "
                                                       "lines as segments. Use the provided absolute source paths "
                                                       "instead of assuming the script directory. Keep imports "
                                                       "within the installed requirements. Return ONLY the fixed "
                                                       "script in a ```python block.")}],
                                        options={"num_ctx": 32768, "temperature": 0.1},
                                        stream=False)
                                    fixed = re.findall(r'```python\n(.*?)\n```',
                                                       fix["message"]["content"], re.DOTALL)
                                    if fixed:
                                        cur = fixed[0]
                                        self.append_to_terminal("info", "Patched. Retrying…")
                                    else:
                                        self.append_to_terminal("err", "No fix returned.")
                                        break
                                else:
                                    self.append_to_terminal("err", "Max retries. Abort.")
                        except subprocess.TimeoutExpired:
                            self.append_to_terminal("err", "Timed out (>60s).")
                            break
                        except Exception as ex:
                            self.append_to_terminal("err", str(ex))
                            break
                        attempt += 1
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)
        finally:
            self.is_executing = False
            self.after(0, lambda: self.set_status("Idle", "idle"))

    # --------------------------------------------------------
    # CORE PIPELINE
    # --------------------------------------------------------
    def start_pipeline(self):
        if self.is_processing: return
        directives = self.prompt_text.get("0.0", "end").strip()
        temperature = round(self.temp_var.get(), 2)
        if not self.model_mgr.is_configured:
            messagebox.showerror("Model Provider",
                "Configure a model provider and endpoint before running analysis.")
            return
        if self.export_var.get() and not self.output_path_var.get().strip():
            self.browse_output()
            if not self.output_path_var.get().strip():
                return
        if self.export_var.get():
            directives += (
                f"\n\nPrepare the completed response as a polished, complete "
                f"{self.output_format_var.get()} document. Return the document content "
                "directly; the application will save it in the selected format. "
                "Do not write a file-generation script."
            )
        self.right_tabs.set("  💬 Workspace  ")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self.run_initial_agent,
                         args=(directives, temperature), daemon=True).start()

    def run_initial_agent(self, directives, temperature):
        try:
            self.append_to_terminal("info",
                f"Pipeline started. {len(self.selected_files)} file(s).")
            self.db.log_message(self.session_id, "user", directives)
            self._workspace_history.append({"role": "user", "content": directives})
            self.append_to_chat("User", directives)

            master_raw = ""
            for fp in self.selected_files:
                self.append_to_terminal("info", f"Extracting: {os.path.basename(fp)}")
                st = extract_dual_stream_from_file(
                    fp, log_callback=lambda m: self.append_to_terminal("info", m))
                master_raw += f'\n<file name="{os.path.basename(fp)}">\n{st["raw"]}\n</file>\n'
                self.db.log_file_access(self.session_id, fp, "extract")

            if not master_raw.strip():
                master_raw = (
                    "No source documents were attached. Create the requested content "
                    "from the user's directives without assuming an input document."
                )

            TOKEN_LIMIT = 20000
            if len(master_raw) > TOKEN_LIMIT * 4:
                self.append_to_terminal("warn",
                    f"~{len(master_raw)//4} tokens. Map-Reduce…")
                chunks = chunk_text(master_raw, TOKEN_LIMIT)
                aggregated = ""
                for i, chunk in enumerate(chunks):
                    self.append_to_terminal("info",
                        f"Mapping chunk {i+1}/{len(chunks)}…")
                    t0 = time.time()
                    r = self.model_mgr.chat(
                        model=self.map_model_var.get(),
                        messages=[{"role": "user",
                                   "content": f"Directive: {directives}\nExtract:\n{chunk}"}],
                        options={"num_ctx": 65536, "temperature": temperature},
                        stream=False)
                    ms = int((time.time() - t0) * 1000)
                    self.db.log_model_usage(self.session_id,
                        self.map_model_var.get(), len(chunk),
                        len(r["message"]["content"]), ms)
                    aggregated += f"\n[Chunk {i+1}]:\n{r['message']['content']}\n"
                self.document_context = aggregated
            else:
                self.document_context = master_raw

            strategy = "Default Synthesis"
            if decision_engine:
                try:
                    res = decision_engine.predict(
                        state=f"User Request:\n{directives}",
                        questions={"strategy": {
                            "type": "choice",
                            "instructions": "Determine analytical approach.",
                            "criteria": {
                                "code_execution": "Write a python script",
                                "data_extraction": "Pull specific values",
                                "comparison": "Compare files"
                            }}})
                    strategy = res.get("answers", res)
                    self.db.log_query(self.session_id, directives,
                                      str(strategy), "")
                except Exception as e:
                    self.append_to_terminal("warn", f"LAYA: {e}")

            self.append_to_terminal("info",
                f"Synthesizing with {self.model_var.get()} (T={temperature})…")
            self._show_typing()

            sys_p = (f"You are a master data agent. Strategy: {strategy}\n"
                     f"{self._script_runtime_context()}\n\n"
                     f"=== DATA ===\n{self.document_context}")
            t0 = time.time()
            stream = self.model_mgr.chat(
                model=self.model_var.get(),
                messages=[{"role": "system", "content": sys_p},
                          {"role": "user", "content": directives}],
                options={"num_ctx": 65536, "temperature": temperature},
                stream=True)

            self.append_to_chat("Agent", "")
            full = ""; buf = ""
            for chunk in stream:
                tok = chunk["message"]["content"]
                full += tok; buf += tok
                if len(buf) > 20 or "\n" in buf:
                    self.append_to_chat("Agent", buf, tag="stream")
                    buf = ""
            if buf: self.append_to_chat("Agent", buf, tag="stream")

            ms = int((time.time() - t0) * 1000)
            self.db.log_model_usage(self.session_id, self.model_var.get(),
                                    len(directives), len(full), ms)
            self._finalise_agent_bubble(full)
            self._workspace_history.append({"role": "agent", "content": full})
            self.db.log_message(self.session_id, "agent", full)
            self.db.log_query(self.session_id, directives, str(strategy),
                              full[:300])
            self.db.upsert_session(self.session_id, len(self.selected_files),
                                   self._msg_count, self.model_var.get())
            self.after(0, lambda response=full:
                       self._save_generated_output(response))
            self.execute_agent_code(full)

        except Exception as e:
            self.append_to_terminal("err", f"CRITICAL: {e}")
            self.set_status("Error", "error")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))

    def send_followup(self):
        if self.is_processing or self.is_executing: return
        text = self.chat_input.get().strip()
        if not text: return
        if not self.model_mgr.is_configured:
            messagebox.showerror("Model Provider",
                "Configure a model provider and endpoint before sending a message.")
            return
        self.chat_input.delete(0, "end")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self._process_followup,
                         args=(text,), daemon=True).start()

    def _process_followup(self, user_text):
        try:
            self.append_to_chat("User", user_text)
            self.db.log_message(self.session_id, "user", user_text)

            # Append to in-memory cache
            self._workspace_history.append({"role": "user", "content": user_text})

            # Trim oldest messages when context exceeds safe limit
            MAX_HISTORY_CHARS = 160000
            while (len(self._workspace_history) > 2 and
                   sum(len(m["content"]) for m in self._workspace_history) > MAX_HISTORY_CHARS):
                self._workspace_history.pop(0)
                self.append_to_terminal("info", "Trimmed oldest message from context window.")

            messages = [{"role": "system",
                         "content": (f"You are a helpful writing and data agent.\n\n"
                                     f"{self._script_runtime_context()}\n\n"
                                     f"Context:\n{self.document_context or 'No source document context is available.'}\n\n"
                                     "Wrap Python scripts in ```python blocks.")}]
            for r in self._workspace_history:
                role = "assistant" if r["role"] == "agent" else r["role"]
                messages.append({"role": role, "content": r["content"]})

            self.set_status("Reasoning…", "busy")
            self._show_typing()

            t0 = time.time()
            stream = self.model_mgr.chat(
                model=self.model_var.get(),
                messages=messages,
                options={"num_ctx": 65536,
                         "temperature": round(self.temp_var.get(), 2)},
                stream=True)

            self.append_to_chat("Agent", "")
            full = ""; buf = ""
            for chunk in stream:
                tok = chunk["message"]["content"]
                full += tok; buf += tok
                if len(buf) > 20 or "\n" in buf:
                    self.append_to_chat("Agent", buf, tag="stream")
                    buf = ""
            if buf: self.append_to_chat("Agent", buf, tag="stream")

            ms = int((time.time() - t0) * 1000)
            self.db.log_model_usage(self.session_id, self.model_var.get(),
                                    len(user_text), len(full), ms)
            self._finalise_agent_bubble(full)
            self._workspace_history.append({"role": "agent", "content": full})
            self.db.log_message(self.session_id, "agent", full)
            self.db.upsert_session(self.session_id, len(self.selected_files),
                                   self._msg_count, self.model_var.get())
            self.execute_agent_code(full)

        except Exception as e:
            self.append_to_terminal("err", f"ERROR: {e}")
            self.set_status("Error", "error")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))

    def send_general_chat(self):
        if self.is_processing or self.is_executing: return
        text = self.general_input.get().strip()
        if not text: return
        if not self.model_mgr.is_configured:
            messagebox.showerror("Model Provider",
                "Configure a model provider and endpoint before sending a message.")
            return

        self.general_input.delete(0, "end")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self._process_general_chat, args=(text,), daemon=True).start()

    def _process_general_chat(self, user_text):
        try:
            gen_session = f"gen_{self.session_id}"
            self.append_to_chat("User", user_text, chat_type="general")
            self.db.log_message(gen_session, "user", user_text)

            # Append to in-memory cache
            self._general_history.append({"role": "user", "content": user_text})

            # Trim if too long
            MAX_HISTORY_CHARS = 120000
            while (len(self._general_history) > 2 and
                   sum(len(m["content"]) for m in self._general_history) > MAX_HISTORY_CHARS):
                self._general_history.pop(0)

            messages = [{"role": "system", "content": "You are a general-purpose AI assistant. Provide helpful and concise answers."}]
            for r in self._general_history:
                role = "assistant" if r["role"] == "agent" else r["role"]
                messages.append({"role": role, "content": r["content"]})

            self.set_status("Reasoning (General)…", "busy")
            self.after(0, lambda: self._general_typing_ind.grid(row=self._general_chat_row, column=0, sticky="w", padx=16, pady=4))
            self.after(0, lambda: self._general_typing_ind.start())
            self.after(60, lambda: self.general_scroll._parent_canvas.yview_moveto(1.0))

            t0 = time.time()
            stream = self.model_mgr.chat(
                model=self.model_var.get(),
                messages=messages,
                options={"num_ctx": 16384, "temperature": round(self.temp_var.get(), 2)},
                stream=True)

            self.append_to_chat("Agent", "", chat_type="general")
            full = ""; buf = ""
            for chunk in stream:
                tok = chunk["message"]["content"]
                full += tok; buf += tok
                if len(buf) > 20 or "\n" in buf:
                    self.append_to_chat("Agent", buf, tag="stream", chat_type="general")
                    buf = ""
            if buf: self.append_to_chat("Agent", buf, tag="stream", chat_type="general")

            ms = int((time.time() - t0) * 1000)
            self.db.log_model_usage(gen_session, self.model_var.get(),
                                    len(user_text), len(full), ms)
            self._finalise_agent_bubble(full, chat_type="general")
            self._general_history.append({"role": "agent", "content": full})
            self.db.log_message(gen_session, "agent", full)
        except Exception as e:
            self.append_to_terminal("err", f"GENERAL CHAT ERROR: {e}")
            self.set_status("Error", "error")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))

    # --------------------------------------------------------
    # SSH UI INTEGRATION
    # --------------------------------------------------------
    def _build_ssh_tab(self):
        self.tab_ssh.grid_columnconfigure(0, weight=1)
        self.tab_ssh.grid_rowconfigure(1, weight=1)

        # Connection Header
        conn_frame = ctk.CTkFrame(self.tab_ssh, fg_color=PALETTE["bg_panel"])
        conn_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        
        ctk.CTkLabel(conn_frame, text="Host:").grid(row=0, column=0, padx=4, pady=4, sticky="e")
        self.ssh_host = ctk.CTkEntry(conn_frame, width=120, fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.ssh_host.grid(row=0, column=1, padx=4, pady=4)
        
        ctk.CTkLabel(conn_frame, text="Port:").grid(row=0, column=2, padx=4, pady=4, sticky="e")
        self.ssh_port = ctk.CTkEntry(conn_frame, width=50, fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.ssh_port.insert(0, "22")
        self.ssh_port.grid(row=0, column=3, padx=4, pady=4)
        
        ctk.CTkLabel(conn_frame, text="User:").grid(row=0, column=4, padx=4, pady=4, sticky="e")
        self.ssh_user = ctk.CTkEntry(conn_frame, width=100, fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.ssh_user.grid(row=0, column=5, padx=4, pady=4)
        
        ctk.CTkLabel(conn_frame, text="Pass/KeyPass:").grid(row=0, column=6, padx=4, pady=4, sticky="e")
        self.ssh_pass = ctk.CTkEntry(conn_frame, width=100, show="*", fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.ssh_pass.grid(row=0, column=7, padx=4, pady=4)
        
        self.ssh_key_path = tk.StringVar()
        ctk.CTkButton(conn_frame, text="🔑 Key", width=60, fg_color=PALETTE["accent"], command=self._ssh_browse_key).grid(row=0, column=8, padx=4, pady=4)
        self.ssh_key_lbl = ctk.CTkLabel(conn_frame, text="No key", width=100)
        self.ssh_key_lbl.grid(row=0, column=9, padx=4, pady=4)
        
        self.ssh_connect_btn = ctk.CTkButton(conn_frame, text="Connect", width=80, fg_color=PALETTE["accent"], command=self._ssh_toggle_connect)
        self.ssh_connect_btn.grid(row=0, column=10, padx=(10, 4), pady=4)

        # Tabs for Terminal vs File Transfer
        self.ssh_tabs = ctk.CTkTabview(self.tab_ssh, corner_radius=6, segmented_button_selected_color=PALETTE["accent"])
        self.ssh_tabs.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        term_tab = self.ssh_tabs.add("Terminal")
        sftp_tab = self.ssh_tabs.add("SFTP (File Transfer)")
        
        # --- SSH Terminal ---
        term_tab.grid_columnconfigure(0, weight=1)
        term_tab.grid_rowconfigure(0, weight=1)
        
        self.ssh_term_box = ctk.CTkTextbox(term_tab, font=ctk.CTkFont(family="Consolas", size=FONT_SCALE["body"]), state="disabled", fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.ssh_term_box.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        
        self.ssh_cmd_entry = ctk.CTkEntry(term_tab, placeholder_text="Enter command...", fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.ssh_cmd_entry.grid(row=1, column=0, sticky="ew", padx=4, pady=4)
        self.ssh_cmd_entry.bind("<Return>", self._ssh_send_cmd)
        
        # --- SSH SFTP ---
        sftp_tab.grid_columnconfigure(0, weight=1)
        sftp_tab.grid_columnconfigure(1, weight=1)
        sftp_tab.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(sftp_tab, text="Local File (to upload):").grid(row=0, column=0, sticky="w", padx=4)
        loc_frame = ctk.CTkFrame(sftp_tab, fg_color="transparent")
        loc_frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        loc_frame.grid_columnconfigure(0, weight=1)
        self.sftp_loc_entry = ctk.CTkEntry(loc_frame, fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.sftp_loc_entry.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(loc_frame, text="Browse", width=60, fg_color=PALETTE["accent"], command=self._ssh_browse_local_upload).grid(row=0, column=1, padx=4)
        
        ctk.CTkLabel(sftp_tab, text="Remote File (to download):").grid(row=0, column=1, sticky="w", padx=4)
        rem_frame = ctk.CTkFrame(sftp_tab, fg_color="transparent")
        rem_frame.grid(row=1, column=1, sticky="nsew", padx=4, pady=4)
        rem_frame.grid_columnconfigure(0, weight=1)
        self.sftp_rem_entry = ctk.CTkEntry(rem_frame, fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"])
        self.sftp_rem_entry.grid(row=0, column=0, sticky="ew")
        
        act_frame = ctk.CTkFrame(sftp_tab, fg_color="transparent")
        act_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ctk.CTkButton(act_frame, text="Upload (Local -> Remote)", fg_color=PALETTE["accent"], command=self._ssh_upload).grid(row=0, column=0, padx=10)
        ctk.CTkButton(act_frame, text="Download (Remote -> Local)", fg_color=PALETTE["accent"], command=self._ssh_download).grid(row=0, column=1, padx=10)

        self._ssh_poll_job = None

    def _ssh_browse_key(self):
        fn = filedialog.askopenfilename(title="Select PEM/PPK Key", filetypes=[("Key Files", "*.pem *.ppk"), ("All Files", "*.*")])
        if fn:
            self.ssh_key_path.set(fn)
            self.ssh_key_lbl.configure(text=os.path.basename(fn))

    def _ssh_toggle_connect(self):
        if self.ssh_mgr.client:
            self.ssh_mgr.disconnect()
            self.ssh_connect_btn.configure(text="Connect", fg_color=PALETTE["accent"])
            if self._ssh_poll_job:
                self.after_cancel(self._ssh_poll_job)
                self._ssh_poll_job = None
            self._ssh_append_term("\n[Disconnected]\n")
        else:
            h = self.ssh_host.get().strip()
            p = self.ssh_port.get().strip()
            u = self.ssh_user.get().strip()
            pw = self.ssh_pass.get()
            kf = self.ssh_key_path.get()
            
            if not h or not u:
                messagebox.showerror("SSH Error", "Host and User are required.")
                return
                
            def _connect_thread():
                self.ssh_connect_btn.configure(state="disabled", text="Connecting...")
                success = self.ssh_mgr.connect(h, p, u, pw if pw else None, kf if kf else None)
                if success:
                    self.after(0, lambda: self.ssh_connect_btn.configure(state="normal", text="Disconnect", fg_color="red"))
                    self.after(0, self._ssh_poll_terminal)
                else:
                    self.after(0, lambda: self.ssh_connect_btn.configure(state="normal", text="Connect", fg_color=PALETTE["accent"]))
                    
            threading.Thread(target=_connect_thread, daemon=True).start()

    def _ssh_append_term(self, text):
        self.ssh_term_box.configure(state="normal")
        self.ssh_term_box.insert("end", text)
        self.ssh_term_box.see("end")
        self.ssh_term_box.configure(state="disabled")

    def _ssh_poll_terminal(self):
        if self.ssh_mgr.client:
            out = self.ssh_mgr.read_shell()
            if out:
                self._ssh_append_term(out)
            self._ssh_poll_job = self.after(100, self._ssh_poll_terminal)
        else:
            self.ssh_connect_btn.configure(text="Connect", fg_color=PALETTE["accent"])

    def _ssh_send_cmd(self, event=None):
        cmd = self.ssh_cmd_entry.get()
        if cmd and self.ssh_mgr.client:
            self.ssh_mgr.execute_command(cmd)
            self.ssh_cmd_entry.delete(0, "end")
            
    def _ssh_browse_local_upload(self):
        fn = filedialog.askopenfilename(title="Select File to Upload")
        if fn:
            self.sftp_loc_entry.delete(0, "end")
            self.sftp_loc_entry.insert(0, fn)
            rem = self.sftp_rem_entry.get()
            if not rem:
                self.sftp_rem_entry.insert(0, "./" + os.path.basename(fn))

    def _ssh_upload(self):
        loc = self.sftp_loc_entry.get().strip()
        rem = self.sftp_rem_entry.get().strip()
        if loc and rem:
            threading.Thread(target=self.ssh_mgr.upload, args=(loc, rem), daemon=True).start()

    def _ssh_download(self):
        rem = self.sftp_rem_entry.get().strip()
        if not rem: return
        fn = filedialog.asksaveasfilename(title="Save Download As", initialfile=os.path.basename(rem))
        if fn:
            threading.Thread(target=self.ssh_mgr.download, args=(rem, fn), daemon=True).start()


if __name__ == "__main__":
    app = AgenticStudioApp()
    app.mainloop()
