import datetime
import json
import os
import re
import sqlite3
import subprocess
import sys
import threading
import uuid
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tempfile
import shutil
import time

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
    from PIL import Image
    import pytesseract
    _tess = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(_tess):
        pytesseract.pytesseract.tesseract_cmd = _tess
except ImportError:
    missing_deps.append("pytesseract/PIL")

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
    DEFAULT_MODELS = ["nemotron-3-ultra:cloud", "gpt-oss:120b-cloud",
                      "llama3:latest", "mistral:latest", "phi3:latest"]

    def __init__(self):
        self._models: list[str] = self.DEFAULT_MODELS[:]
        self._lock = threading.Lock()

    def refresh(self, callback=None):
        def _do():
            try:
                if ollama:
                    result = ollama.list()
                    names = [m["name"] for m in result.get("models", [])]
                    if names:
                        with self._lock:
                            self._models = names + [m for m in self.DEFAULT_MODELS if m not in names]
            except Exception:
                pass
            if callback:
                callback(self._models[:])
        threading.Thread(target=_do, daemon=True).start()

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
                 rerun_callback=None, **kw):
        is_user = sender == "user"
        super().__init__(parent, fg_color=PALETTE["bg_deep"],
                         corner_radius=0, **kw)
        self.grid_columnconfigure(0, weight=1)

        outer = ctk.CTkFrame(self, fg_color=PALETTE["bg_deep"], corner_radius=0)
        outer.grid(row=0, column=0, sticky="ew", padx=8, pady=4)
        outer.grid_columnconfigure(0 if not is_user else 1, weight=1)

        avatar = ctk.CTkLabel(outer, text="👤" if is_user else "🤖",
                               font=ctk.CTkFont(size=14),
                               text_color=PALETTE["text_mid"] if is_user else PALETTE["accent"])
        avatar.grid(row=0, column=1 if is_user else 0,
                    padx=(6, 0) if is_user else (0, 6), pady=4, sticky="n")

        bubble_col = 0 if is_user else 1
        bg = PALETTE["accent"] if is_user else PALETTE["bg_card"]
        bubble = ctk.CTkFrame(outer, fg_color=bg, corner_radius=14,
                               border_width=1,
                               border_color=PALETTE["border_hi"] if not is_user else bg)
        bubble.grid(row=0, column=bubble_col, sticky="ew")
        bubble.grid_columnconfigure(0, weight=1)

        # Right-click context menu
        self._text = text
        self._copy_cb = copy_callback
        self._rerun_cb = rerun_callback
        bubble.bind("<Button-3>", self._show_context_menu)

        self._render_content(bubble, text, copy_callback, rerun_callback)

    def _show_context_menu(self, event):
        menu = tk.Menu(self, tearoff=0,
                       bg=PALETTE["bg_card"], fg=PALETTE["text_hi"],
                       activebackground=PALETTE["accent"],
                       activeforeground="#ffffff", bd=0, relief="flat")
        if self._copy_cb:
            menu.add_command(label="  ⎘  Copy Message",
                             command=lambda: self._copy_cb(self._text))
        code_blocks = re.findall(r'```python\n(.*?)\n```', self._text, re.DOTALL)
        if code_blocks and self._copy_cb:
            menu.add_command(label="  ⎘  Copy Code Block",
                             command=lambda: self._copy_cb(code_blocks[0]))
        if code_blocks and self._rerun_cb:
            menu.add_separator()
            menu.add_command(label="  ▶  Re-run Code",
                             command=lambda: self._rerun_cb(self._text))
        menu.tk_popup(event.x_root, event.y_root)

    def _render_content(self, parent, text, copy_cb, rerun_cb):
        code_pat = re.compile(r'```(\w*)\n(.*?)\n```', re.DOTALL)
        last = 0; row = 0
        for m in code_pat.finditer(text):
            plain = text[last:m.start()].strip()
            if plain:
                lbl = ctk.CTkLabel(parent, text=plain, wraplength=700,
                                   justify="left",
                                   font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                   text_color=PALETTE["text_hi"])
                lbl.grid(row=row, column=0, sticky="w", padx=14, pady=(10, 4))
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
            lbl = ctk.CTkLabel(parent, text=trailing, wraplength=700,
                               justify="left",
                               font=ctk.CTkFont(size=FONT_SCALE["body"]),
                               text_color=PALETTE["text_hi"])
            lbl.grid(row=row, column=0, sticky="w", padx=14, pady=(8, 12))


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
        self.model_var = ctk.StringVar(value="nemotron-3-ultra:cloud")
        self.map_model_var = ctk.StringVar(value="gpt-oss:120b-cloud")
        self.db = DatabaseManager()
        self.model_mgr = ModelManager()
        self.session_id = str(uuid.uuid4())[:8]
        self.selected_files = []
        self.document_context = ""
        self.is_processing = False
        self.is_executing = False
        self._status_text = tk.StringVar(value="Idle")
        self._stream_agent_frame = None
        self._stream_text_acc = ""
        self._stream_label = None
        self._chat_row = 1
        self._msg_count = 0
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
        mm.add_command(label="  Refresh Available Models", command=self._refresh_models_ui)
        mm.add_separator()
        mm.add_command(label="  Set Inference Model…",     command=lambda: self.open_settings())
        mm.add_command(label="  Set Map-Reduce Model…",    command=lambda: self.open_settings())
        mm.add_separator()
        mm.add_command(label="  SSH Remote (coming soon)", state="disabled")
        self.menubar.add_cascade(label=" Models ", menu=mm)

        # History
        hm = _menu()
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
        # Row weights: 0=sep, 1=file list, 2=file btns,
        #              3=model hdr, 4=inf model, 5=map label+combo, 6=map combo, 7=refresh,
        #              8=ctrl hdr, 9=temp slider,
        #              10=export hdr, 11=export chk+entry, 12=browse btn,
        #              13=dir hdr, 14=directives(flex), 15=run btn, 16=progress, 17=settings
        for i, w in [(0,0),(1,1),(2,0),(3,0),(4,0),(5,0),(6,0),(7,0),
                     (8,0),(9,0),(10,0),(11,0),(12,0),
                     (13,0),(14,2),(15,0),(16,0),(17,0)]:
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
        self.file_list_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))
        self.file_list_frame.grid_columnconfigure(0, weight=1)

        # File buttons
        bframe = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        bframe.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 6))
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

        self._section_label_grid(3, "MODELS")

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
        self.model_combo.grid(row=4, column=0, sticky="ew", padx=8, pady=(2, 3))

        ctk.CTkLabel(self.left_panel, text="Map-Reduce Model",
                     font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                     text_color=PALETTE["text_lo"]).grid(row=5, column=0,
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
        self.map_combo.grid(row=6, column=0, sticky="ew", padx=8, pady=(0, 3))

        ctk.CTkButton(self.left_panel, text="↺  Refresh Models",
                      command=self._refresh_models_ui,
                      fg_color="transparent",
                      hover_color=PALETTE["border"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                      height=24, corner_radius=4
                      ).grid(row=7, column=0, sticky="e", padx=8, pady=(0, 4))

        self._section_label_grid(8, "INFERENCE CONTROL")

        # Temp slider
        temp_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        temp_frame.grid(row=9, column=0, sticky="ew", padx=8, pady=(2, 4))
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

        self._section_label_grid(10, "OUTPUT / EXPORT")

        # Export checkbox + path entry on same row sub-frame
        exp_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        exp_frame.grid(row=11, column=0, sticky="ew", padx=8, pady=(2, 2))
        exp_frame.grid_columnconfigure(1, weight=1)
        self.export_var = tk.BooleanVar(value=False)
        self.output_path_var = ctk.StringVar()
        ctk.CTkCheckBox(exp_frame, text="",
                        variable=self.export_var,
                        command=self._toggle_export,
                        fg_color=PALETTE["accent"],
                        hover_color=PALETTE["accent_hot"],
                        width=22, height=22
                        ).grid(row=0, column=0, padx=(0, 4))
        self.output_entry = ctk.CTkEntry(exp_frame,
                                         textvariable=self.output_path_var,
                                         placeholder_text="Output .docx path…",
                                         state="disabled",
                                         fg_color=PALETTE["bg_input"],
                                         text_color=PALETTE["text_hi"],
                                         border_color=PALETTE["border"],
                                         font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
                                         height=28)
        self.output_entry.grid(row=0, column=1, sticky="ew")
        self.browse_out_btn = ctk.CTkButton(
            self.left_panel, text="📁  Browse Output Path",
            command=self.browse_output,
            state="disabled",
            fg_color="transparent",
            hover_color=PALETTE["border"],
            text_color=PALETTE["text_lo"],
            font=ctk.CTkFont(size=FONT_SCALE["small"]-1),
            height=24, corner_radius=4)
        self.browse_out_btn.grid(row=12, column=0, sticky="e", padx=8, pady=(0, 4))

        self._section_label_grid(13, "AGENT DIRECTIVES")

        self.prompt_text = ctk.CTkTextbox(self.left_panel,
                                          fg_color=PALETTE["bg_input"],
                                          text_color=PALETTE["text_hi"],
                                          border_color=PALETTE["border_hi"],
                                          border_width=1,
                                          font=ctk.CTkFont(size=FONT_SCALE["body"]))
        self.prompt_text.grid(row=14, column=0, sticky="nsew", padx=8, pady=(2, 6))
        self.prompt_text.insert("0.0", "Compare the provided documents and extract common numbers.")

        # Gradient-effect Run button using Canvas
        self.run_canvas = tk.Canvas(self.left_panel, height=46,
                                    highlightthickness=0, bd=0,
                                    bg=PALETTE["bg_panel"])
        self.run_canvas.grid(row=15, column=0, sticky="ew", padx=8, pady=(0, 4))
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
                      ).grid(row=17, column=0, sticky="ew", padx=8, pady=(0, 6))

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
        self.tab_terminal = self.right_tabs.add("  ⚙ Terminal  ")
        self.tab_preview  = self.right_tabs.add("  🔍 Preview  ")

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

        # Bind mouse wheel directly for Windows
        self.chat_scroll.bind_all("<MouseWheel>", self._on_mousewheel)
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

        self.send_btn = ctk.CTkButton(input_bg, text="Send ⏎",
                                      font=ctk.CTkFont(size=FONT_SCALE["small"],
                                                       weight="bold"),
                                      width=86, height=36, corner_radius=8,
                                      fg_color=PALETTE["accent"],
                                      hover_color=PALETTE["accent_hot"],
                                      command=self.send_followup)
        self.send_btn.grid(row=0, column=1, padx=(4, 8), pady=6)

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
    def _on_models_refreshed(self, models):
        def _upd():
            self.model_combo.configure(values=models)
            self.map_combo.configure(values=models)
            self.append_to_terminal("info",
                f"Model list refreshed: {len(models)} model(s) found.")
        self.after(0, _upd)

    def _refresh_models_ui(self):
        self.append_to_terminal("info", "Querying available models…")
        self.model_mgr.refresh(callback=self._on_models_refreshed)

    # --------------------------------------------------------
    # WORKSPACE INTERACTIVITY (7F)
    # --------------------------------------------------------
    def _on_mousewheel(self, event):
        try:
            self.chat_scroll._parent_canvas.yview_scroll(
                int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

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
        win.geometry("520x460")
        win.configure(fg_color=PALETTE["bg_panel"])
        win.grab_set()

        ctk.CTkLabel(win, text="Preferences",
                     font=ctk.CTkFont(size=FONT_SCALE["h1"], weight="bold"),
                     text_color=PALETTE["text_hi"]).pack(pady=(20, 10))

        for lbl, var in [("Inference Model", self.model_var),
                          ("Map-Reduce Model", self.map_model_var)]:
            ctk.CTkLabel(win, text=lbl,
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["text_mid"]).pack(anchor="w", padx=30, pady=(8, 2))
            ctk.CTkEntry(win, textvariable=var, width=440,
                         fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"],
                         border_color=PALETTE["border"],
                         font=ctk.CTkFont(size=FONT_SCALE["body"])).pack(padx=30)

        if 'pytesseract' in sys.modules:
            ctk.CTkLabel(win, text="Tesseract Path",
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["text_mid"]).pack(anchor="w", padx=30, pady=(8, 2))
            tvar = ctk.StringVar(value=pytesseract.pytesseract.tesseract_cmd)
            ctk.CTkEntry(win, textvariable=tvar, width=440,
                         fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"],
                         border_color=PALETTE["border"],
                         font=ctk.CTkFont(size=FONT_SCALE["small"])).pack(padx=30)

        ctk.CTkLabel(win, text="Theme",
                     font=ctk.CTkFont(size=FONT_SCALE["small"]),
                     text_color=PALETTE["text_mid"]).pack(anchor="w", padx=30, pady=(8, 2))
        tvar2 = ctk.StringVar(value=self._active_theme)
        ctk.CTkComboBox(win, variable=tvar2, values=list(THEMES.keys()),
                        fg_color=PALETTE["bg_input"],
                        border_color=PALETTE["border"],
                        button_color=PALETTE["accent"],
                        text_color=PALETTE["text_hi"],
                        dropdown_fg_color=PALETTE["bg_card"],
                        dropdown_text_color=PALETTE["text_hi"],
                        width=440).pack(padx=30)

        def save():
            self.db.set_setting("inference_model", self.model_var.get())
            self.db.set_setting("map_model", self.map_model_var.get())
            self.apply_theme(tvar2.get())
            if 'pytesseract' in sys.modules:
                pytesseract.pytesseract.tesseract_cmd = tvar.get()
            win.destroy()

        ctk.CTkButton(win, text="Save & Close", command=save,
                      fg_color=PALETTE["accent"],
                      hover_color=PALETTE["accent_hot"],
                      font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
                      height=40, width=180, corner_radius=8).pack(pady=24)

    # --------------------------------------------------------
    # FILE OPERATIONS
    # --------------------------------------------------------
    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[
            ("All Supported", "*.docx;*.pdf;*.xlsx;*.csv;*.txt;*.png;*.jpg;*.jpeg")])
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
                             "Agentic_Report.docx"))

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
        fn = filedialog.asksaveasfilename(defaultextension=".docx",
                                          filetypes=[("Word Documents", "*.docx")])
        if fn:
            self.output_path_var.set(fn)

    def _toggle_export(self):
        state = "normal" if self.export_var.get() else "disabled"
        self.output_entry.configure(state=state)
        self.browse_out_btn.configure(state=state)

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
        self._typing_ind = TypingIndicator(self.chat_scroll)

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
    def append_to_chat(self, sender, text, tag=None):
        if sender == "System":
            level = ("ok" if "✅" in text else "err" if "❌" in text
                     else "warn" if any(w in text for w in ("WARNING","WARN","warn")) else "info")
            self.append_to_terminal(level, text)
            return

        def _do():
            if sender == "User":
                self._hide_typing()
                b = MessageBubble(self.chat_scroll, "user", text,
                                  copy_callback=self._copy_to_clipboard)
                b.grid(row=self._chat_row, column=0, sticky="ew", pady=2)
                self._chat_row += 1
                self._update_msg_count()
                self.after(50, self._smooth_scroll_bottom)

            elif sender == "Agent" and not tag:
                self._hide_typing()
                self._stream_text_acc = ""
                frame = ctk.CTkFrame(self.chat_scroll,
                                     fg_color=PALETTE["bg_card"],
                                     corner_radius=12,
                                     border_width=1,
                                     border_color=PALETTE["border"])
                frame.grid(row=self._chat_row, column=0, sticky="ew", pady=2)
                frame.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(frame, text="🤖",
                             font=ctk.CTkFont(size=14),
                             text_color=PALETTE["accent"]
                             ).grid(row=0, column=0, sticky="nw", padx=10, pady=(8, 2))
                self._stream_label = ctk.CTkLabel(frame, text="",
                                                   wraplength=740,
                                                   justify="left",
                                                   font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                                   text_color=PALETTE["text_hi"])
                self._stream_label.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
                self._stream_agent_frame = frame
                self._chat_row += 1

            elif sender == "Agent" and tag == "stream":
                self._stream_text_acc += text
                if self._stream_label:
                    self._stream_label.configure(text=self._stream_text_acc)
                self.after(30, self._smooth_scroll_bottom)

        self.after(0, _do)

    def _finalise_agent_bubble(self, full_text):
        def _do():
            if self._stream_agent_frame:
                row_info = self._stream_agent_frame.grid_info()
                row = row_info.get("row", self._chat_row - 1)
                self._stream_agent_frame.destroy()
                self._stream_agent_frame = None
                self._stream_label = None
                b = MessageBubble(self.chat_scroll, "agent", full_text,
                                  copy_callback=self._copy_to_clipboard,
                                  rerun_callback=self.execute_agent_code)
                b.grid(row=row, column=0, sticky="ew", pady=2)
                self._update_msg_count()
                self.after(60, self._smooth_scroll_bottom)
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

    # --------------------------------------------------------
    # PROCESSING STATE
    # --------------------------------------------------------
    def toggle_processing_state(self, processing: bool):
        self.is_processing = processing
        state = "disabled" if processing else "normal"
        self.send_btn.configure(state=state)
        if processing:
            self.run_canvas.configure(cursor="watch")
            self.progress_bar.grid(row=16, column=0, sticky="ew",
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
    def execute_agent_code(self, llm_response, max_retries=3):
        blocks = re.findall(r'```python\n(.*?)\n```', llm_response, re.DOTALL)
        if not blocks: return
        self.is_executing = True
        self.set_status("Executing generated script…", "busy")
        threading.Thread(target=self._exec_thread,
                         args=(blocks, max_retries), daemon=True).start()

    def _exec_thread(self, blocks, max_retries):
        model = self.model_var.get()
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
                            res = subprocess.run(
                                [sys.executable, sp],
                                capture_output=True, text=True,
                                cwd=tmp, timeout=60)
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
                                if attempt < max_retries and ollama:
                                    self.append_to_terminal("warn", "Self-healing…")
                                    fix = ollama.chat(
                                        model=model,
                                        messages=[{"role": "user",
                                                   "content": (
                                                       f"Script failed:\n```\n{err}\n```\n"
                                                       f"Code:\n```python\n{cur}\n```\n"
                                                       "Return ONLY the fixed script in a ```python block.")}],
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
        if not self.selected_files:
            messagebox.showerror("No Files", "Add at least one document first.")
            return
        # Inject docx export directive if enabled
        out_path = self.output_path_var.get().strip().replace("\\", "/")
        if self.export_var.get() and out_path:
            directives += (
                f"\n\nCRITICAL DIRECTIVE: Write a complete `python-docx` script that saves "
                f"the full analysis report to '{out_path}'. Wrap it in a ```python block."
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
            self.append_to_chat("User", directives)

            master_raw = ""
            for fp in self.selected_files:
                self.append_to_terminal("info", f"Extracting: {os.path.basename(fp)}")
                st = extract_dual_stream_from_file(
                    fp, log_callback=lambda m: self.append_to_terminal("info", m))
                master_raw += f'\n<file name="{os.path.basename(fp)}">\n{st["raw"]}\n</file>\n'
                self.db.log_file_access(self.session_id, fp, "extract")

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
                    r = ollama.chat(
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

            sys_p = (f"You are a master data agent. Strategy: {strategy}\n\n"
                     f"=== DATA ===\n{self.document_context}")
            t0 = time.time()
            stream = ollama.chat(
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
            self.db.log_message(self.session_id, "agent", full)
            self.db.log_query(self.session_id, directives, str(strategy),
                              full[:300])
            self.db.upsert_session(self.session_id, len(self.selected_files),
                                   self._msg_count, self.model_var.get())
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
        if not self.document_context:
            messagebox.showwarning("Not Ready", "Run Initialize Analysis first.")
            return
        self.chat_input.delete(0, "end")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self._process_followup,
                         args=(text,), daemon=True).start()

    def _process_followup(self, user_text):
        try:
            self.append_to_chat("User", user_text)
            self.db.log_message(self.session_id, "user", user_text)

            history = self.db.get_session_history(self.session_id, limit=10)
            total = sum(len(r["content"]) for r in history)
            if total > 160000:
                self.append_to_terminal("warn",
                    f"History trimmed (last 10 msgs, ~{total//4} tokens).")

            messages = [{"role": "system",
                         "content": (f"You are a helpful data agent.\n\n"
                                     f"Context:\n{self.document_context}\n\n"
                                     "Wrap Python scripts in ```python blocks.")}]
            for r in history:
                role = "assistant" if r["role"] == "agent" else r["role"]
                messages.append({"role": role, "content": r["content"]})

            self.set_status("Reasoning…", "busy")
            self._show_typing()

            t0 = time.time()
            stream = ollama.chat(
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
            self.db.log_message(self.session_id, "agent", full)
            self.db.upsert_session(self.session_id, len(self.selected_files),
                                   self._msg_count, self.model_var.get())
            self.execute_agent_code(full)

        except Exception as e:
            self.append_to_terminal("err", f"ERROR: {e}")
            self.set_status("Error", "error")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))


if __name__ == "__main__":
    app = AgenticStudioApp()
    app.mainloop()
