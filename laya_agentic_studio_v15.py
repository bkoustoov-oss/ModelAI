import datetime
import os
import re
import sqlite3
import subprocess
import sys
import threading
import uuid
import tkinter as tk
from tkinter import filedialog, messagebox
import tempfile
import shutil
import time

# Force offline caching
os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["TRANSFORMERS_OFFLINE"] = "1"

import customtkinter as ctk

# ==========================================
# DESIGN SYSTEM — PALETTE & TYPOGRAPHY
# ==========================================
PALETTE = {
    "bg_deep":    "#080C14",
    "bg_panel":   "#0F1623",
    "bg_card":    "#161D2E",
    "bg_input":   "#111827",
    "accent":     "#4F8EF7",
    "accent_hot": "#7B5CF0",
    "success":    "#22C55E",
    "warning":    "#F59E0B",
    "danger":     "#EF4444",
    "text_hi":    "#F1F5F9",
    "text_mid":   "#94A3B8",
    "text_lo":    "#475569",
    "border":     "#1E2D45",
    "glow_idle":  "#22C55E",
    "glow_busy":  "#F59E0B",
    "glow_err":   "#EF4444",
}

FONT_SCALE = {"h1": 20, "h2": 15, "body": 13, "small": 11, "mono": 12}

FILE_ICON_MAP = {
    ".pdf":  "📋", ".docx": "📝", ".doc": "📝",
    ".xlsx": "📊", ".csv":  "📊",
    ".txt":  "📄",
    ".png":  "🖼", ".jpg":  "🖼", ".jpeg": "🖼",
}

# Configure modern dark theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ==========================================
# PRE-FLIGHT DEPENDENCY CHECK
# ==========================================
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
    tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
except ImportError:
    missing_deps.append("pytesseract/PIL")

# ==========================================
# MODULE 1: DATABASE MANAGER
# ==========================================
class DatabaseManager:
    def __init__(self, db_path="agentic_memory.db"):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("PRAGMA journal_mode=WAL")  # FIX 6C: WAL mode for thread safety
            conn.execute('''CREATE TABLE IF NOT EXISTS chat_history (
                                id INTEGER PRIMARY KEY AUTOINCREMENT,
                                session_id TEXT,
                                timestamp TEXT,
                                role TEXT,
                                content TEXT
                            )''')
            conn.commit()

    def log_message(self, session_id, role, content):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("INSERT INTO chat_history (session_id, timestamp, role, content) VALUES (?, ?, ?, ?)",
                         (session_id, timestamp, role, content))
            conn.commit()

    def get_session_history(self, session_id, limit=10):  # FIX 6E: sliding window
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute(
                "SELECT role, content FROM chat_history WHERE session_id = ? ORDER BY id DESC LIMIT ?",
                (session_id, limit)
            )
            rows = cursor.fetchall()
        return [{"role": r[0], "content": r[1]} for r in reversed(rows)]

# ==========================================
# MODULE 2: DOCUMENT EXTRACTION ENGINE
# ==========================================
def extract_dual_stream_from_file(file_path, log_callback=None):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.png', '.jpg', '.jpeg']:
        text = pytesseract.image_to_string(Image.open(file_path))
        return {"spatial": text, "raw": text}
    elif ext == '.pdf':
        try:
            import pdfplumber
        except ImportError:
            raise ImportError("pdfplumber is not installed. Run: pip install pdfplumber")
        spatial_text, raw_text = "", ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                s_text = page.extract_text(layout=True)
                r_text = page.extract_text(layout=False)
                if not s_text or len(s_text.strip()) < 15:
                    if log_callback:
                        log_callback(f"[{os.path.basename(file_path)}] Engaging OCR on sparse page...")
                    pil_img = page.to_image(resolution=300).original
                    s_text = r_text = pytesseract.image_to_string(pil_img)
                spatial_text += (s_text or "") + "\n"
                raw_text += (r_text or "") + "\n"
        return {"spatial": spatial_text, "raw": raw_text}
    elif ext == '.docx':
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        return {"spatial": text, "raw": text}
    elif ext in ['.xlsx', '.csv']:
        try:
            import pandas as pd
        except ImportError:
            raise ImportError("pandas is not installed. Run: pip install pandas openpyxl")
        df = pd.read_csv(file_path) if ext == '.csv' else pd.read_excel(file_path)
        return {"spatial": df.to_string(), "raw": df.to_string()}
    elif ext == '.txt':
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            text = f.read()
        return {"spatial": text, "raw": text}
    else:
        raise ValueError(f"Unsupported file format: {ext}")


def chunk_text(text, token_limit=20000):
    """Heuristic chunking: ~4 chars per token."""
    chunk_size = token_limit * 4
    return [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]

# ==========================================
# MODULE 3: CUSTOM WIDGETS
# ==========================================

class GlowDot(tk.Canvas):
    """Animated status indicator dot."""
    STATES = {"idle": PALETTE["glow_idle"], "busy": PALETTE["glow_busy"], "error": PALETTE["glow_err"]}

    def __init__(self, parent, **kwargs):
        super().__init__(parent, width=12, height=12, highlightthickness=0,
                         bg=PALETTE["bg_deep"], **kwargs)
        self._state = "idle"
        self._alpha = 1.0
        self._dir = -1
        self._draw()
        self._animate()

    def _draw(self):
        self.delete("all")
        color = self.STATES[self._state]
        alpha_hex = format(int(self._alpha * 200 + 55), '02x')
        try:
            self.create_oval(2, 2, 10, 10, fill=color, outline=color)
        except Exception:
            pass

    def _animate(self):
        self._alpha += self._dir * 0.04
        if self._alpha <= 0.3:
            self._dir = 1
        elif self._alpha >= 1.0:
            self._dir = -1
        self._draw()
        self.after(60, self._animate)

    def set_state(self, state):
        self._state = state if state in self.STATES else "idle"


class TypingIndicator(ctk.CTkFrame):
    """Animated 3-dot typing indicator."""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color=PALETTE["bg_card"], corner_radius=12, **kwargs)
        self._dots = []
        self._running = False
        self._step = 0
        for i in range(3):
            c = tk.Canvas(self, width=8, height=8, highlightthickness=0, bg=PALETTE["bg_card"])
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
        if not self._running:
            return
        self._step += 1
        self._render_dots()
        self.after(350, self._tick)

    def stop(self):
        self._running = False


class MessageBubble(ctk.CTkFrame):
    """A single chat message bubble with optional code block rendering."""
    def __init__(self, parent, sender, text, copy_callback=None, **kwargs):
        is_user = sender == "user"
        bg = PALETTE["accent"] if is_user else PALETTE["bg_card"]
        super().__init__(parent, fg_color=PALETTE["bg_deep"], corner_radius=0, **kwargs)
        self.grid_columnconfigure(0, weight=1)

        outer = ctk.CTkFrame(self, fg_color=PALETTE["bg_deep"], corner_radius=0)
        outer.grid(row=0, column=0, sticky="ew", padx=8, pady=4)
        outer.grid_columnconfigure(0, weight=1)

        if is_user:
            avatar_lbl = ctk.CTkLabel(outer, text="👤", font=ctk.CTkFont(size=14),
                                      text_color=PALETTE["text_mid"])
            avatar_lbl.grid(row=0, column=1, padx=(6, 0), pady=4, sticky="n")
            bubble_col = 0
        else:
            avatar_lbl = ctk.CTkLabel(outer, text="🤖", font=ctk.CTkFont(size=14),
                                      text_color=PALETTE["accent"])
            avatar_lbl.grid(row=0, column=0, padx=(0, 6), pady=4, sticky="n")
            bubble_col = 1

        outer.grid_columnconfigure(bubble_col, weight=1)

        bubble = ctk.CTkFrame(outer, fg_color=bg, corner_radius=12)
        bubble.grid(row=0, column=bubble_col, sticky="ew")
        bubble.grid_columnconfigure(0, weight=1)

        self._render_content(bubble, text, copy_callback)

    def _render_content(self, parent, text, copy_callback):
        """Split text into plain and code segments, render each."""
        code_pattern = re.compile(r'```(\w*)\n(.*?)\n```', re.DOTALL)
        last = 0
        row = 0
        for m in code_pattern.finditer(text):
            plain = text[last:m.start()].strip()
            if plain:
                lbl = ctk.CTkLabel(parent, text=plain, wraplength=700, justify="left",
                                   font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                   text_color=PALETTE["text_hi"])
                lbl.grid(row=row, column=0, sticky="w", padx=12, pady=(8, 4))
                row += 1

            lang = m.group(1) or "code"
            code = m.group(2)
            code_frame = ctk.CTkFrame(parent, fg_color=PALETTE["bg_deep"], corner_radius=8)
            code_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=4)
            code_frame.grid_columnconfigure(0, weight=1)

            header = ctk.CTkFrame(code_frame, fg_color=PALETTE["border"], corner_radius=0, height=28)
            header.grid(row=0, column=0, sticky="ew")
            header.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(header, text=f" {lang}", font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["accent"]).grid(row=0, column=0, sticky="w", padx=8)
            if copy_callback:
                copy_btn = ctk.CTkButton(header, text="⎘ Copy", width=60, height=22,
                                         fg_color=PALETTE["bg_card"], hover_color=PALETTE["accent"],
                                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                                         command=lambda c=code: copy_callback(c))
                copy_btn.grid(row=0, column=1, padx=4, pady=2)

            code_box = ctk.CTkTextbox(code_frame, fg_color=PALETTE["bg_deep"],
                                      text_color="#A8FF78",
                                      font=ctk.CTkFont(family="Consolas", size=FONT_SCALE["mono"]),
                                      height=min(300, max(60, code.count('\n') * 18 + 40)))
            code_box.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
            code_box.insert("0.0", code)
            code_box.configure(state="disabled")
            row += 1
            last = m.end()

        trailing = text[last:].strip()
        if trailing:
            lbl = ctk.CTkLabel(parent, text=trailing, wraplength=700, justify="left",
                               font=ctk.CTkFont(size=FONT_SCALE["body"]),
                               text_color=PALETTE["text_hi"])
            lbl.grid(row=row, column=0, sticky="w", padx=12, pady=(8, 10))


class FileCard(ctk.CTkFrame):
    """Styled file card with icon, name, and remove button."""
    def __init__(self, parent, file_path, on_preview, on_remove, **kwargs):
        super().__init__(parent, fg_color=PALETTE["bg_card"], corner_radius=8,
                         border_width=1, border_color=PALETTE["border"], **kwargs)
        self.grid_columnconfigure(1, weight=1)

        ext = os.path.splitext(file_path)[1].lower()
        icon = FILE_ICON_MAP.get(ext, "📄")
        name = os.path.basename(file_path)
        display = name if len(name) <= 22 else name[:19] + "..."

        ctk.CTkLabel(self, text=icon, font=ctk.CTkFont(size=16),
                     width=30).grid(row=0, column=0, padx=(8, 4), pady=8)

        name_btn = ctk.CTkButton(self, text=display, anchor="w",
                                 fg_color="transparent", hover_color=PALETTE["border"],
                                 text_color=PALETTE["text_hi"],
                                 font=ctk.CTkFont(size=FONT_SCALE["small"]),
                                 command=lambda: on_preview(file_path))
        name_btn.grid(row=0, column=1, sticky="ew", pady=6)

        rm_btn = ctk.CTkButton(self, text="✕", width=28, height=24,
                               fg_color="transparent", hover_color=PALETTE["danger"],
                               text_color=PALETTE["text_lo"],
                               font=ctk.CTkFont(size=12),
                               command=lambda: on_remove(file_path, self))
        rm_btn.grid(row=0, column=2, padx=(4, 6), pady=6)


# ==========================================
# MODULE 4: MAIN APPLICATION
# ==========================================
class AgenticStudioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("LAYA Agentic Studio v15")
        self.geometry("1450x960")
        self.minsize(1150, 800)
        self.configure(fg_color=PALETTE["bg_deep"])

        # FIX 6A: Initialize model vars at __init__ time — no more hasattr guards
        self.model_var = ctk.StringVar(value="nemotron-3-ultra:cloud")
        self.map_model_var = ctk.StringVar(value="gpt-oss:120b-cloud")

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=0)  # title bar
        self.grid_rowconfigure(1, weight=1)  # content
        self.grid_rowconfigure(2, weight=0)  # status bar

        self.db = DatabaseManager()
        self.session_id = str(uuid.uuid4())[:8]
        self.selected_files = []
        self.document_context = ""
        self.is_processing = False
        self.is_executing = False  # FIX 6F: separate exec flag
        self._status_text = tk.StringVar(value="Idle")
        self._stream_agent_frame = None   # current streaming bubble frame ref
        self._stream_text_acc = ""        # accumulated stream text

        self._build_title_bar()
        self._build_ui()
        self._build_status_bar()
        self._start_clock()

        if missing_deps:
            self.append_to_terminal("warn",
                f"Missing dependencies: {', '.join(missing_deps)}. Some features may be unavailable.")
            self.after(600, lambda: messagebox.showwarning(
                "Missing Dependencies",
                f"Missing packages: {', '.join(missing_deps)}\n\nSee System Terminal for details."))

    # ------------------------------------------
    # TITLE BAR
    # ------------------------------------------
    def _build_title_bar(self):
        self.title_bar = tk.Frame(self, bg=PALETTE["bg_deep"], height=42)
        self.title_bar.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.title_bar.grid_columnconfigure(1, weight=1)
        self.title_bar.pack_propagate(False)

        # Glow dot + title
        left = tk.Frame(self.title_bar, bg=PALETTE["bg_deep"])
        left.grid(row=0, column=0, sticky="w", padx=12)
        self.glow_dot = GlowDot(left)
        self.glow_dot.pack(side="left", padx=(0, 8), pady=10)
        tk.Label(left, text="LAYA  Agentic Studio",
                 font=("Consolas", 13, "bold"),
                 bg=PALETTE["bg_deep"], fg=PALETTE["text_hi"]).pack(side="left")
        tk.Label(left, text=" v15",
                 font=("Consolas", 11),
                 bg=PALETTE["bg_deep"], fg=PALETTE["accent"]).pack(side="left")

        # Session badge
        self._session_badge = tk.Label(self.title_bar,
                                       text=f"SESSION  {self.session_id}",
                                       font=("Consolas", 9),
                                       bg=PALETTE["bg_deep"], fg=PALETTE["text_lo"])
        self._session_badge.grid(row=0, column=1, sticky="")

        # Win controls
        ctrl_frame = tk.Frame(self.title_bar, bg=PALETTE["bg_deep"])
        ctrl_frame.grid(row=0, column=2, sticky="e", padx=8)
        for symbol, action, color in [
            ("─", self.iconify, PALETTE["text_lo"]),
            ("□", self._toggle_maximise, PALETTE["text_mid"]),
            ("✕", self.destroy, PALETTE["danger"]),
        ]:
            btn = tk.Label(ctrl_frame, text=symbol, font=("Consolas", 13),
                           bg=PALETTE["bg_deep"], fg=color, padx=10, pady=8, cursor="hand2")
            btn.pack(side="left")
            btn.bind("<Button-1>", lambda e, a=action: a())
            btn.bind("<Enter>", lambda e, b=btn: b.configure(bg=PALETTE["bg_card"]))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(bg=PALETTE["bg_deep"]))

        # Drag support
        self.title_bar.bind("<ButtonPress-1>", self._start_move)
        self.title_bar.bind("<B1-Motion>", self._do_move)
        self._drag_x = self._drag_y = 0

    def _start_move(self, e):
        self._drag_x, self._drag_y = e.x_root, e.y_root

    def _do_move(self, e):
        dx = e.x_root - self._drag_x
        dy = e.y_root - self._drag_y
        x = self.winfo_x() + dx
        y = self.winfo_y() + dy
        self.geometry(f"+{x}+{y}")
        self._drag_x, self._drag_y = e.x_root, e.y_root

    def _toggle_maximise(self):
        self.state("normal" if self.state() == "zoomed" else "zoomed")

    # ------------------------------------------
    # STATUS BAR
    # ------------------------------------------
    def _build_status_bar(self):
        bar = tk.Frame(self, bg=PALETTE["border"], height=26)
        bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        bar.grid_columnconfigure(1, weight=1)

        self._file_badge = tk.Label(bar, text="0 files", font=("Consolas", 9),
                                    bg=PALETTE["border"], fg=PALETTE["text_lo"])
        self._file_badge.grid(row=0, column=0, padx=12, sticky="w")

        op_lbl = tk.Label(bar, textvariable=self._status_text, font=("Consolas", 9),
                          bg=PALETTE["border"], fg=PALETTE["accent"])
        op_lbl.grid(row=0, column=1)

        self._clock_lbl = tk.Label(bar, text="", font=("Consolas", 9),
                                   bg=PALETTE["border"], fg=PALETTE["text_lo"])
        self._clock_lbl.grid(row=0, column=2, padx=12, sticky="e")

        self._model_lbl = tk.Label(bar, textvariable=self.model_var, font=("Consolas", 9),
                                   bg=PALETTE["border"], fg=PALETTE["text_lo"])
        self._model_lbl.grid(row=0, column=3, padx=12, sticky="e")

    def _start_clock(self):
        def tick():
            self._clock_lbl.configure(text=datetime.datetime.now().strftime("%H:%M:%S"))
            self.after(1000, tick)
        tick()

    def set_status(self, text, glow_state="idle"):
        self._status_text.set(text)
        self.glow_dot.set_state(glow_state)

    # ------------------------------------------
    # MAIN UI
    # ------------------------------------------
    def _build_ui(self):
        # ---- LEFT PANEL ----
        self.left_panel = ctk.CTkScrollableFrame(self, corner_radius=0,
                                                 fg_color=PALETTE["bg_panel"],
                                                 scrollbar_fg_color=PALETTE["bg_panel"])
        self.left_panel.grid(row=1, column=0, padx=(8, 4), pady=(4, 4), sticky="nsew")

        self._section_label("DATA SOURCES")

        self.file_list_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        self.file_list_frame.pack(fill="x", padx=10, pady=(4, 8))

        btn_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 12))
        ctk.CTkButton(btn_frame, text="＋  Add Files", command=self.add_files,
                      fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                      height=32, corner_radius=8).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkButton(btn_frame, text="✕  Clear All", command=self.clear_files,
                      fg_color=PALETTE["bg_card"], hover_color=PALETTE["danger"],
                      text_color=PALETTE["text_mid"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      height=32, corner_radius=8).pack(side="right")

        self._section_label("EXPORT OPTIONS")
        self.export_var = tk.BooleanVar(value=False)
        self.export_chk = ctk.CTkCheckBox(self.left_panel, text="Generate .docx Report Script",
                                          variable=self.export_var, command=self.toggle_export,
                                          text_color=PALETTE["text_mid"],
                                          fg_color=PALETTE["accent"],
                                          font=ctk.CTkFont(size=FONT_SCALE["small"]))
        self.export_chk.pack(anchor="w", padx=10, pady=(4, 4))
        self.output_path_var = ctk.StringVar()
        self.output_entry = ctk.CTkEntry(self.left_panel, textvariable=self.output_path_var,
                                         state="disabled", fg_color=PALETTE["bg_input"],
                                         text_color=PALETTE["text_hi"], border_color=PALETTE["border"],
                                         font=ctk.CTkFont(size=FONT_SCALE["small"]))
        self.output_entry.pack(fill="x", padx=10, pady=(2, 4))
        self.browse_out_btn = ctk.CTkButton(self.left_panel, text="📁  Browse Output Path",
                                            command=self.browse_output, state="disabled",
                                            fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"],
                                            text_color=PALETTE["text_mid"],
                                            font=ctk.CTkFont(size=FONT_SCALE["small"]),
                                            height=30, corner_radius=6)
        self.browse_out_btn.pack(anchor="e", padx=10, pady=(0, 12))

        self._section_label("INFERENCE CONTROL")
        ctk.CTkLabel(self.left_panel, text="Temperature",
                     font=ctk.CTkFont(size=FONT_SCALE["small"]),
                     text_color=PALETTE["text_mid"]).pack(anchor="w", padx=10, pady=(4, 0))
        self.temp_var = tk.DoubleVar(value=0.2)
        self.temp_slider = ctk.CTkSlider(self.left_panel, from_=0.0, to=1.0,
                                          variable=self.temp_var, command=self.update_temp_label,
                                          progress_color=PALETTE["accent"],
                                          button_color=PALETTE["accent_hot"],
                                          button_hover_color=PALETTE["text_hi"])
        self.temp_slider.pack(fill="x", padx=10, pady=(4, 2))
        self.temp_label = ctk.CTkLabel(self.left_panel, text="0.20 — Strict / Analytical",
                                       font=ctk.CTkFont(size=FONT_SCALE["small"], slant="italic"),
                                       text_color=PALETTE["success"])
        self.temp_label.pack(anchor="w", padx=10, pady=(0, 12))

        self._section_label("AGENT DIRECTIVES")
        self.prompt_text = ctk.CTkTextbox(self.left_panel, height=130,
                                          fg_color=PALETTE["bg_input"],
                                          text_color=PALETTE["text_hi"],
                                          border_color=PALETTE["border"],
                                          border_width=1,
                                          font=ctk.CTkFont(size=FONT_SCALE["body"]))
        self.prompt_text.pack(fill="x", padx=10, pady=(4, 12))
        self.prompt_text.insert("0.0", "Compare the provided documents and extract common numbers.")

        self.run_btn = ctk.CTkButton(self.left_panel, text="▶  Initialize Analysis",
                                     font=ctk.CTkFont(size=FONT_SCALE["h2"], weight="bold"),
                                     height=48, corner_radius=10,
                                     fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
                                     command=self.start_pipeline)
        self.run_btn.pack(fill="x", padx=10, pady=(4, 8))

        self.progress_bar = ctk.CTkProgressBar(self.left_panel, mode="indeterminate",
                                               progress_color=PALETTE["accent_hot"],
                                               fg_color=PALETTE["border"],
                                               height=4)
        self.progress_bar.set(0)

        ctk.CTkButton(self.left_panel, text="⚙  Settings",
                      command=self.open_settings,
                      fg_color=PALETTE["bg_card"], hover_color=PALETTE["border"],
                      text_color=PALETTE["text_mid"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      height=32, corner_radius=6).pack(fill="x", padx=10, pady=(4, 12))

        # ---- RIGHT PANEL ----
        self.right_tabs = ctk.CTkTabview(self, corner_radius=8,
                                         fg_color=PALETTE["bg_panel"],
                                         segmented_button_fg_color=PALETTE["bg_card"],
                                         segmented_button_selected_color=PALETTE["accent"],
                                         segmented_button_selected_hover_color=PALETTE["accent_hot"],
                                         segmented_button_unselected_color=PALETTE["bg_card"],
                                         segmented_button_unselected_hover_color=PALETTE["border"],
                                         text_color=PALETTE["text_hi"],
                                         text_color_disabled=PALETTE["text_lo"])
        self.right_tabs.grid(row=1, column=1, padx=(4, 8), pady=(4, 4), sticky="nsew")

        self.tab_chat = self.right_tabs.add("  💬 Workspace  ")
        self.tab_terminal = self.right_tabs.add("  ⚙ Terminal  ")
        self.tab_preview = self.right_tabs.add("  🔍 Preview  ")

        # Chat tab
        self.tab_chat.grid_rowconfigure(0, weight=1)
        self.tab_chat.grid_columnconfigure(0, weight=1)

        self.chat_scroll = ctk.CTkScrollableFrame(self.tab_chat,
                                                   fg_color=PALETTE["bg_deep"],
                                                   scrollbar_fg_color=PALETTE["bg_panel"])
        self.chat_scroll.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self.chat_scroll.grid_columnconfigure(0, weight=1)

        # Welcome message
        welcome = ctk.CTkLabel(self.chat_scroll,
                               text=f"Session [{self.session_id}] initialized.\nAdd files and click  ▶ Initialize Analysis  to begin.",
                               font=ctk.CTkFont(size=FONT_SCALE["body"]),
                               text_color=PALETTE["text_lo"],
                               justify="center")
        welcome.grid(row=0, column=0, pady=30)
        self._chat_row = 1

        # Typing indicator (hidden by default)
        self._typing_ind = TypingIndicator(self.chat_scroll)

        # Input row
        input_frame = ctk.CTkFrame(self.tab_chat, fg_color=PALETTE["bg_card"],
                                   corner_radius=10, border_width=1,
                                   border_color=PALETTE["border"])
        input_frame.grid(row=1, column=0, sticky="ew", padx=6, pady=6)
        input_frame.grid_columnconfigure(0, weight=1)
        self.chat_input = ctk.CTkEntry(input_frame,
                                       placeholder_text="Ask the agent a follow-up question or request a script…",
                                       font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                       fg_color="transparent",
                                       border_width=0,
                                       text_color=PALETTE["text_hi"])
        self.chat_input.grid(row=0, column=0, sticky="ew", padx=8, ipady=10)
        self.chat_input.bind("<Return>", lambda e: self.send_followup())
        self.send_btn = ctk.CTkButton(input_frame, text="Send ⏎",
                                      font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                                      width=90, height=38, corner_radius=8,
                                      fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
                                      command=self.send_followup)
        self.send_btn.grid(row=0, column=1, padx=8, pady=5)

        # Terminal tab
        self.tab_terminal.grid_rowconfigure(1, weight=1)
        self.tab_terminal.grid_columnconfigure(0, weight=1)

        term_toolbar = ctk.CTkFrame(self.tab_terminal, fg_color=PALETTE["bg_card"], height=32)
        term_toolbar.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))
        ctk.CTkLabel(term_toolbar, text="System Terminal",
                     font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                     text_color=PALETTE["text_mid"]).pack(side="left", padx=10)
        ctk.CTkButton(term_toolbar, text="⊘ Clear", width=60, height=24,
                      fg_color="transparent", hover_color=PALETTE["border"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      command=lambda: (self.terminal_log.configure(state="normal"),
                                       self.terminal_log.delete("0.0", "end"),
                                       self.terminal_log.configure(state="disabled"))).pack(side="right", padx=4)
        ctk.CTkButton(term_toolbar, text="⎘ Copy", width=60, height=24,
                      fg_color="transparent", hover_color=PALETTE["border"],
                      text_color=PALETTE["text_lo"],
                      font=ctk.CTkFont(size=FONT_SCALE["small"]),
                      command=self._copy_terminal).pack(side="right", padx=4)

        self.terminal_log = ctk.CTkTextbox(self.tab_terminal,
                                            fg_color="#000000",
                                            text_color=PALETTE["success"],
                                            font=ctk.CTkFont(family="Consolas", size=FONT_SCALE["mono"]))
        self.terminal_log.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        self.terminal_log.tag_config("ok",    foreground=PALETTE["success"])
        self.terminal_log.tag_config("err",   foreground=PALETTE["danger"])
        self.terminal_log.tag_config("warn",  foreground=PALETTE["warning"])
        self.terminal_log.tag_config("info",  foreground=PALETTE["text_mid"])
        self.terminal_log.insert("0.0", "─── SYSTEM TERMINAL ───\n", "info")

        # Preview tab
        self.tab_preview.grid_rowconfigure(0, weight=1)
        self.tab_preview.grid_columnconfigure(0, weight=1)
        self.preview_box = ctk.CTkTextbox(self.tab_preview,
                                          fg_color=PALETTE["bg_card"],
                                          text_color=PALETTE["text_hi"],
                                          font=ctk.CTkFont(family="Consolas", size=FONT_SCALE["small"]))
        self.preview_box.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.preview_box.insert("0.0", "Click a file card to preview its extracted content.")

    def _section_label(self, text):
        frame = ctk.CTkFrame(self.left_panel, fg_color="transparent", height=24)
        frame.pack(fill="x", padx=10, pady=(10, 2))
        frame.pack_propagate(False)
        bar = ctk.CTkFrame(frame, fg_color=PALETTE["accent"], width=3, corner_radius=2)
        bar.pack(side="left", fill="y", padx=(0, 8))
        ctk.CTkLabel(frame, text=text,
                     font=ctk.CTkFont(size=FONT_SCALE["small"], weight="bold"),
                     text_color=PALETTE["text_lo"]).pack(side="left", anchor="w")

    # ------------------------------------------
    # UI HELPERS
    # ------------------------------------------
    def open_settings(self):
        win = ctk.CTkToplevel(self)
        win.title("Settings")
        win.geometry("480x380")
        win.configure(fg_color=PALETTE["bg_panel"])
        win.grab_set()

        ctk.CTkLabel(win, text="Application Settings",
                     font=ctk.CTkFont(size=FONT_SCALE["h1"], weight="bold"),
                     text_color=PALETTE["text_hi"]).pack(pady=(20, 16))

        for label, var in [("Inference Model (nemotron / etc.)", self.model_var),
                            ("Map-Reduce Model (chunk processing)", self.map_model_var)]:
            ctk.CTkLabel(win, text=label,
                         font=ctk.CTkFont(size=FONT_SCALE["small"]),
                         text_color=PALETTE["text_mid"]).pack(anchor="w", padx=30, pady=(8, 2))
            ctk.CTkEntry(win, textvariable=var, width=400,
                         fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"],
                         border_color=PALETTE["border"],
                         font=ctk.CTkFont(size=FONT_SCALE["body"])).pack(padx=30)

        ctk.CTkLabel(win, text="Tesseract OCR Path",
                     font=ctk.CTkFont(size=FONT_SCALE["small"]),
                     text_color=PALETTE["text_mid"]).pack(anchor="w", padx=30, pady=(8, 2))
        self.tess_path_var = ctk.StringVar(value=pytesseract.pytesseract.tesseract_cmd
                                           if 'pytesseract' in sys.modules else "Not installed")
        ctk.CTkEntry(win, textvariable=self.tess_path_var, width=400,
                     fg_color=PALETTE["bg_input"], text_color=PALETTE["text_hi"],
                     border_color=PALETTE["border"],
                     font=ctk.CTkFont(size=FONT_SCALE["small"])).pack(padx=30)

        def save_close():
            if 'pytesseract' in sys.modules and self.tess_path_var.get():
                pytesseract.pytesseract.tesseract_cmd = self.tess_path_var.get()
            win.destroy()

        ctk.CTkButton(win, text="Save & Close", command=save_close,
                      fg_color=PALETTE["accent"], hover_color=PALETTE["accent_hot"],
                      font=ctk.CTkFont(size=FONT_SCALE["body"], weight="bold"),
                      height=40, width=180, corner_radius=8).pack(pady=24)

    def toggle_export(self):
        state = "normal" if self.export_var.get() else "disabled"
        self.output_entry.configure(state=state)
        self.browse_out_btn.configure(state=state)

    def update_temp_label(self, value):
        val = round(float(value), 2)
        if val <= 0.2:
            desc, color = "Strict / Analytical", PALETTE["success"]
        elif val <= 0.5:
            desc, color = "Balanced", PALETTE["warning"]
        else:
            desc, color = "Creative / Exploratory", PALETTE["danger"]
        self.temp_label.configure(text=f"{val:.2f} — {desc}", text_color=color)

    def toggle_processing_state(self, processing: bool):
        self.is_processing = processing
        state = "disabled" if processing else "normal"
        self.run_btn.configure(state=state)
        self.send_btn.configure(state=state)
        if processing:
            self.progress_bar.pack(fill="x", padx=10, pady=(2, 6))
            self.progress_bar.start()
            self.set_status("Processing...", "busy")
        else:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            if not self.is_executing:
                self.set_status("Idle", "idle")

    def add_files(self):
        filenames = filedialog.askopenfilenames(
            filetypes=[("All Supported Files", "*.docx;*.pdf;*.xlsx;*.csv;*.txt;*.png;*.jpg;*.jpeg")])
        for f in filenames:
            if f not in self.selected_files:
                self.selected_files.append(f)
                card = FileCard(self.file_list_frame, f,
                                on_preview=self.load_preview,
                                on_remove=self._remove_file)
                card.pack(fill="x", pady=3, padx=2)
                self._slide_in(card)
        self._update_file_badge()
        if self.selected_files and not self.output_path_var.get():
            self.output_path_var.set(
                os.path.join(os.path.dirname(self.selected_files[0]), "Agentic_Report.docx"))

    def _slide_in(self, widget):
        """Animate card sliding in from left."""
        target_x = widget.winfo_x()
        widget.place(x=-300, y=widget.winfo_y())
        steps = 12
        def step(i=0):
            if i > steps:
                widget.place_forget()
                return
            x = int(-300 + (300 * i / steps))
            widget.lift()
            self.after(12 * i, step, i + 1)
        step()

    def _remove_file(self, file_path, card_widget):
        if file_path in self.selected_files:
            self.selected_files.remove(file_path)
        card_widget.destroy()
        self._update_file_badge()

    def _update_file_badge(self):
        n = len(self.selected_files)
        self._file_badge.configure(text=f"{n} file{'s' if n != 1 else ''}")

    def load_preview(self, file_path):
        self.right_tabs.set("  🔍 Preview  ")
        self.preview_box.configure(state="normal")
        self.preview_box.delete("0.0", "end")
        self.preview_box.insert("end", f"Extracting {os.path.basename(file_path)}...\n")
        self.preview_box.configure(state="disabled")

        # FIX 6D: Run extraction in background thread
        def _do():
            try:
                streams = extract_dual_stream_from_file(file_path)
                raw = streams['raw']
                content = (f"═══ {os.path.basename(file_path)} ═══\n\n"
                           f"── RAW STREAM (first 4000 chars) ──\n"
                           f"{raw[:4000]}")
                if len(raw) > 4000:
                    content += "\n\n…[TRUNCATED FOR PREVIEW]…"
            except Exception as e:
                content = f"Extraction error:\n{e}"
            def _update():
                self.preview_box.configure(state="normal")
                self.preview_box.delete("0.0", "end")
                self.preview_box.insert("end", content)
                self.preview_box.configure(state="disabled")
            self.after(0, _update)
        threading.Thread(target=_do, daemon=True).start()

    def clear_files(self):
        self.selected_files.clear()
        for w in self.file_list_frame.winfo_children():
            w.destroy()
        self._update_file_badge()

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.output_path_var.set(filename)

    def _copy_terminal(self):
        self.terminal_log.configure(state="normal")
        content = self.terminal_log.get("0.0", "end")
        self.terminal_log.configure(state="disabled")
        self.clipboard_clear()
        self.clipboard_append(content)

    # ------------------------------------------
    # LOGGING
    # ------------------------------------------
    def append_to_terminal(self, level, text):
        """level: ok | err | warn | info"""
        tag_map = {"ok": "ok", "err": "err", "warn": "warn", "info": "info"}
        icon_map = {"ok": "✅", "err": "❌", "warn": "⚠", "info": "⚙"}
        tag = tag_map.get(level, "info")
        icon = icon_map.get(level, "⚙")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        msg = f"[{ts}] {icon}  {text}\n"
        def _do():
            self.terminal_log.configure(state="normal")
            self.terminal_log.insert("end", msg, tag)
            self.terminal_log.see("end")
            self.terminal_log.configure(state="disabled")
        self.after(0, _do)

    def append_to_chat(self, sender, text, tag=None):
        if sender == "System":
            level = "ok" if "✅" in text else "err" if "❌" in text else "warn" if "WARNING" in text else "info"
            self.append_to_terminal(level, text)
            return

        def _do():
            if sender == "User":
                self._hide_typing()
                bubble = MessageBubble(self.chat_scroll, "user", text,
                                       copy_callback=self._copy_to_clipboard)
                bubble.grid(row=self._chat_row, column=0, sticky="ew", pady=2)
                self._chat_row += 1

            elif sender == "Agent" and not tag:
                self._hide_typing()
                # Create a streaming bubble — plain label for now, finalised on completion
                self._stream_text_acc = ""
                frame = ctk.CTkFrame(self.chat_scroll,
                                     fg_color=PALETTE["bg_card"], corner_radius=12)
                frame.grid(row=self._chat_row, column=0, sticky="ew", pady=2)
                frame.grid_columnconfigure(0, weight=1)

                avatar = ctk.CTkLabel(frame, text="🤖",
                                      font=ctk.CTkFont(size=14),
                                      text_color=PALETTE["accent"])
                avatar.grid(row=0, column=0, sticky="nw", padx=10, pady=(10, 4))

                self._stream_label = ctk.CTkLabel(frame, text="",
                                                   wraplength=720, justify="left",
                                                   font=ctk.CTkFont(size=FONT_SCALE["body"]),
                                                   text_color=PALETTE["text_hi"])
                self._stream_label.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
                self._stream_agent_frame = frame
                self._chat_row += 1

            elif sender == "Agent" and tag == "stream":
                self._stream_text_acc += text
                if self._stream_label:
                    self._stream_label.configure(text=self._stream_text_acc)

        self.after(0, _do)

    def _finalise_agent_bubble(self, full_text):
        """Replace streaming label frame with a proper MessageBubble."""
        def _do():
            if self._stream_agent_frame:
                row_info = self._stream_agent_frame.grid_info()
                row = row_info.get("row", self._chat_row - 1)
                self._stream_agent_frame.destroy()
                self._stream_agent_frame = None
                self._stream_label = None
                bubble = MessageBubble(self.chat_scroll, "agent", full_text,
                                       copy_callback=self._copy_to_clipboard)
                bubble.grid(row=row, column=0, sticky="ew", pady=2)
        self.after(0, _do)

    def _show_typing(self):
        def _do():
            self._typing_ind.grid(row=self._chat_row, column=0, sticky="w",
                                  padx=16, pady=4)
            self._typing_ind.start()
        self.after(0, _do)

    def _hide_typing(self):
        self._typing_ind.stop()
        self._typing_ind.grid_remove()

    def _copy_to_clipboard(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.set_status("Copied to clipboard!", "idle")
        self.after(2000, lambda: self.set_status("Idle", "idle"))

    # ------------------------------------------
    # SELF-HEALING CODE EXECUTION ENGINE
    # ------------------------------------------
    def execute_agent_code(self, llm_response, max_retries=3):
        """FIX 6F: Run in own thread with is_executing flag."""
        code_blocks = re.findall(r'```python\n(.*?)\n```', llm_response, re.DOTALL)
        if not code_blocks:
            return
        self.is_executing = True
        self.set_status("Executing generated script…", "busy")
        threading.Thread(target=self._exec_thread, args=(code_blocks, max_retries), daemon=True).start()

    def _exec_thread(self, code_blocks, max_retries):
        model_name = self.model_var.get()
        try:
            for idx, code in enumerate(code_blocks):
                # FIX 6B: use mkdtemp so we control lifecycle precisely
                temp_dir = tempfile.mkdtemp()
                script_path = os.path.join(temp_dir, f"agent_script_{self.session_id}_{idx}.py")
                current_code = code
                success = False
                attempt = 1
                try:
                    while attempt <= max_retries and not success:
                        self.append_to_terminal("info",
                            f"Script #{idx+1} — Attempt {attempt}/{max_retries}")
                        with open(script_path, "w", encoding="utf-8") as f:
                            f.write(current_code)
                        try:
                            result = subprocess.run(
                                [sys.executable, script_path],
                                capture_output=True, text=True,
                                cwd=temp_dir, timeout=60)
                            if result.returncode == 0:
                                out = result.stdout.strip() or "Completed with no output."
                                self.append_to_terminal("ok", f"Script successful.\n{out}")
                                success = True
                            else:
                                err = result.stderr.strip()
                                self.append_to_terminal("err", f"Script failed.\n{err}")
                                if attempt < max_retries and ollama:
                                    self.append_to_terminal("warn",
                                        "Self-healing: routing traceback to LLM…")
                                    fix_prompt = (
                                        f"You wrote a Python script that failed:\n\n"
                                        f"```\n{err}\n```\n\n"
                                        f"Broken code:\n```python\n{current_code}\n```\n\n"
                                        f"Output ONLY the fixed script in a ```python block."
                                    )
                                    fix = ollama.chat(
                                        model=model_name,
                                        messages=[{"role": "user", "content": fix_prompt}],
                                        options={"num_ctx": 32768, "temperature": 0.1},
                                        stream=False)
                                    fixed = re.findall(r'```python\n(.*?)\n```',
                                                       fix["message"]["content"], re.DOTALL)
                                    if fixed:
                                        current_code = fixed[0]
                                        self.append_to_terminal("info",
                                            "Patched code received. Retrying…")
                                    else:
                                        self.append_to_terminal("err",
                                            "Self-healing failed: no valid Python block returned.")
                                        break
                                else:
                                    self.append_to_terminal("err", "Max retries reached. Aborting.")
                        except subprocess.TimeoutExpired:
                            self.append_to_terminal("err", "Script timed out (>60s). Killed.")
                            break
                        except Exception as e:
                            self.append_to_terminal("err", f"System error: {str(e)}")
                            break
                        attempt += 1
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
        finally:
            self.is_executing = False
            self.after(0, lambda: self.set_status("Idle", "idle"))

    # ------------------------------------------
    # CORE PIPELINE
    # ------------------------------------------
    def start_pipeline(self):
        if self.is_processing:
            return
        directives = self.prompt_text.get("0.0", "end").strip()
        temperature = round(self.temp_var.get(), 2)
        if not self.selected_files:
            messagebox.showerror("No Files", "Please add at least one document before running.")
            return
        if self.export_var.get() and self.output_path_var.get():
            out = self.output_path_var.get().strip().replace("\\", "/")
            directives += (
                f"\n\nCRITICAL DIRECTIVE: Write a complete `python-docx` script that saves "
                f"the report to '{out}'. Wrap it in a ```python block."
            )
        self.right_tabs.set("  💬 Workspace  ")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self.run_initial_agent, args=(directives, temperature), daemon=True).start()

    def run_initial_agent(self, directives, temperature):
        try:
            self.append_to_terminal("info", f"Extracting {len(self.selected_files)} file(s)…")
            self.set_status("Extracting files…", "busy")
            self.db.log_message(self.session_id, "user", directives)
            self.append_to_chat("User", directives)

            master_raw = ""
            for fp in self.selected_files:
                self.append_to_terminal("info", f"Processing: {os.path.basename(fp)}")
                streams = extract_dual_stream_from_file(fp, log_callback=lambda m: self.append_to_terminal("info", m))
                master_raw += f'\n<file name="{os.path.basename(fp)}">\n{streams["raw"]}\n</file>\n'

            TOKEN_LIMIT = 20000
            if len(master_raw) > TOKEN_LIMIT * 4:
                self.append_to_terminal("warn",
                    f"~{len(master_raw)//4} tokens detected. Initiating Map-Reduce chunking…")
                self.set_status("Map-Reduce chunking…", "busy")
                chunks = chunk_text(master_raw, TOKEN_LIMIT)
                aggregated = ""
                for i, chunk in enumerate(chunks):
                    self.append_to_terminal("info", f"Mapping chunk {i+1}/{len(chunks)}…")
                    r = ollama.chat(
                        model=self.map_model_var.get(),
                        messages=[{"role": "user",
                                   "content": f"Directive: {directives}\nExtract targets:\n{chunk}"}],
                        options={"num_ctx": 65536, "temperature": temperature},
                        stream=False)
                    aggregated += f"\n[Chunk {i+1}]:\n{r['message']['content']}\n"
                self.document_context = aggregated
            else:
                self.document_context = master_raw

            laya_strategy = "Default Synthesis"
            if decision_engine:
                self.append_to_terminal("info", "Consulting LAYA Decision Engine…")
                self.set_status("LAYA strategy…", "busy")
                try:
                    res = decision_engine.predict(
                        state=f"User Request:\n{directives}",
                        questions={"strategy": {
                            "type": "choice",
                            "instructions": "Determine the required analytical approach.",
                            "criteria": {
                                "code_execution": "Requires writing a python script",
                                "data_extraction": "Pulling specific values",
                                "comparison": "Comparing files"
                            }
                        }})
                    laya_strategy = res.get("answers", res)
                except Exception as e:
                    self.append_to_terminal("warn", f"LAYA failed, using default. {e}")

            self.append_to_terminal("info",
                f"Synthesizing with {self.model_var.get()} (T={temperature})…")
            self.set_status("LLM synthesizing…", "busy")
            self._show_typing()

            sys_prompt = (f"You are a master data agent. Strategy: {laya_strategy}\n\n"
                          f"=== DATA ===\n{self.document_context}")
            stream = ollama.chat(
                model=self.model_var.get(),
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": directives}],
                options={"num_ctx": 65536, "temperature": temperature},
                stream=True)

            self.append_to_chat("Agent", "")
            full = ""
            buf = ""
            for chunk in stream:
                tok = chunk["message"]["content"]
                full += tok
                buf += tok
                if len(buf) > 20 or "\n" in buf:
                    self.append_to_chat("Agent", buf, tag="stream")
                    buf = ""
            if buf:
                self.append_to_chat("Agent", buf, tag="stream")

            self._finalise_agent_bubble(full)
            self.db.log_message(self.session_id, "agent", full)
            self.execute_agent_code(full)

        except Exception as e:
            self.append_to_terminal("err", f"CRITICAL ERROR: {e}")
            self.set_status("Error", "error")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))

    def send_followup(self):
        if self.is_processing or self.is_executing:
            return
        user_text = self.chat_input.get().strip()
        if not user_text:
            return
        if not self.document_context:
            messagebox.showwarning("Not Ready", "Please run Initialize Analysis first.")
            return
        self.chat_input.delete(0, "end")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self._process_followup, args=(user_text,), daemon=True).start()

    def _process_followup(self, user_text):
        try:
            self.append_to_chat("User", user_text)
            self.db.log_message(self.session_id, "user", user_text)

            # FIX 6E: Sliding window — last 10 messages
            history = self.db.get_session_history(self.session_id, limit=10)
            total_chars = sum(len(r["content"]) for r in history)
            if total_chars > 160000:
                self.append_to_terminal("warn",
                    f"History trimmed to last 10 msgs ({total_chars//4} tokens).")

            messages = [{"role": "system",
                         "content": (f"You are a helpful data agent.\n\n"
                                     f"Active Context:\n{self.document_context}\n\n"
                                     f"Wrap any Python scripts in ```python blocks.")}]
            for r in history:
                role = "assistant" if r["role"] == "agent" else r["role"]
                messages.append({"role": role, "content": r["content"]})

            self.set_status("Reasoning…", "busy")
            self.append_to_terminal("info", "Processing follow-up…")
            self._show_typing()

            stream = ollama.chat(
                model=self.model_var.get(),
                messages=messages,
                options={"num_ctx": 65536, "temperature": round(self.temp_var.get(), 2)},
                stream=True)

            self.append_to_chat("Agent", "")
            full = ""
            buf = ""
            for chunk in stream:
                tok = chunk["message"]["content"]
                full += tok
                buf += tok
                if len(buf) > 20 or "\n" in buf:
                    self.append_to_chat("Agent", buf, tag="stream")
                    buf = ""
            if buf:
                self.append_to_chat("Agent", buf, tag="stream")

            self._finalise_agent_bubble(full)
            self.db.log_message(self.session_id, "agent", full)
            self.execute_agent_code(full)

        except Exception as e:
            self.append_to_terminal("err", f"ERROR: {e}")
            self.set_status("Error", "error")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))


if __name__ == "__main__":
    app = AgenticStudioApp()
    app.mainloop()
