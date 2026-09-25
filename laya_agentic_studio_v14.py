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
import time

# Force offline caching
os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["TRANSFORMERS_OFFLINE"] = "1"

import customtkinter as ctk

# Pre-flight dependency check
missing_deps = []
try:
    from docx import Document
except ImportError:
    missing_deps.append("python-docx")

try:
    import laya
    print("Loading LAYA System 1 Decision Engine (Offline Cache)...")
    decision_engine = laya.load("convaiinnovations/laya")
except ImportError:
    missing_deps.append("laya")
    decision_engine = None

try:
    import ollama
except ImportError:
    missing_deps.append("ollama")

try:
    from PIL import Image
    import pytesseract
    tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
except ImportError:
    missing_deps.append("pytesseract/PIL")

# Configure modern dark theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ==========================================
# MODULE 1: LOCAL SQLITE DATABASE MANAGER
# ==========================================
class DatabaseManager:
    def __init__(self, db_path="agentic_memory.db"):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
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
        # Use a new connection for thread safety
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT INTO chat_history (session_id, timestamp, role, content) VALUES (?, ?, ?, ?)",
                         (session_id, timestamp, role, content))
            conn.commit()

    def get_session_history(self, session_id):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("SELECT role, content FROM chat_history WHERE session_id = ? ORDER BY id ASC", (session_id,))
            return [{"role": row[0], "content": row[1]} for row in cursor.fetchall()]

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
            raise ImportError("pdfplumber is not installed.")
        
        spatial_text, raw_text = "", ""
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                s_text = page.extract_text(layout=True)
                r_text = page.extract_text(layout=False)
                if not s_text or len(s_text.strip()) < 15:
                    if log_callback: log_callback(f"[{os.path.basename(file_path)}] Engaging OCR...")
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
            raise ImportError("pandas is not installed.")
        
        df = pd.read_csv(file_path) if ext == '.csv' else pd.read_excel(file_path)
        return {"spatial": df.to_string(), "raw": df.to_string()}
    elif ext == '.txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
            return {"spatial": text, "raw": text}
    else:
        raise ValueError(f"Unsupported file format: {ext}")

def chunk_text(text, token_limit=20000):
    """
    A heuristic chunking approach: assuming ~4 characters per token.
    Provides better alignment with LLM context windows.
    """
    chunk_size = token_limit * 4
    return [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]

# ==========================================
# MODULE 3: MAIN APPLICATION GUI
# ==========================================
class AgenticStudioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Universal LAYA Agentic Studio - v14 (Self-Healing Code Engine)")
        self.geometry("1400x950")
        self.minsize(1100, 800)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=1)
        
        self.db = DatabaseManager()
        self.session_id = str(uuid.uuid4())[:8]
        self.selected_files = [] 
        self.document_context = "" 
        self.is_processing = False
        
        # Stream buffering state
        self.stream_buffer = ""
        self.last_update_time = time.time()

        self._build_ui()
        
        if missing_deps:
            self.append_to_terminal(f"WARNING: The following dependencies are missing: {', '.join(missing_deps)}")
            self.append_to_terminal("Please pip install them for full functionality.")
            self.after(500, lambda: messagebox.showwarning("Missing Dependencies", f"Missing packages: {', '.join(missing_deps)}\n\nCheck System Terminal for details."))

    def _build_ui(self):
        # Left Panel (Controls)
        self.left_panel = ctk.CTkScrollableFrame(self, corner_radius=10)
        self.left_panel.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        title_lbl = ctk.CTkLabel(self.left_panel, text="Data Sources", font=ctk.CTkFont(size=18, weight="bold"))
        title_lbl.pack(pady=(10, 15))

        self.file_list_frame = ctk.CTkFrame(self.left_panel, fg_color="#1e1e1e")
        self.file_list_frame.pack(fill="x", padx=10, pady=(5, 5))
        
        btn_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 15))
        ctk.CTkButton(btn_frame, text="Add Files...", command=self.add_files, width=120).pack(side="left")
        ctk.CTkButton(btn_frame, text="Clear", command=self.clear_files, width=80, fg_color="#7f1d1d", hover_color="#991b1b").pack(side="right")

        self.export_var = tk.BooleanVar(value=False)
        self.export_chk = ctk.CTkCheckBox(self.left_panel, text="Agent: Write Python Script to build .docx", variable=self.export_var, command=self.toggle_export)
        self.export_chk.pack(anchor="w", padx=10, pady=(15, 5))
        
        self.output_path_var = ctk.StringVar()
        self.output_entry = ctk.CTkEntry(self.left_panel, textvariable=self.output_path_var, state="disabled")
        self.output_entry.pack(fill="x", padx=10, pady=(5, 5))
        self.browse_out_btn = ctk.CTkButton(self.left_panel, text="Browse Output", command=self.browse_output, state="disabled")
        self.browse_out_btn.pack(anchor="e", padx=10, pady=(0, 15))
        
        ctk.CTkLabel(self.left_panel, text="AI Temperature", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=10)
        self.temp_var = tk.DoubleVar(value=0.2)
        self.temp_slider = ctk.CTkSlider(self.left_panel, from_=0.0, to=1.0, variable=self.temp_var, command=self.update_temp_label)
        self.temp_slider.pack(fill="x", padx=10, pady=(5, 2))
        self.temp_label = ctk.CTkLabel(self.left_panel, text="Current: 0.2 (Strict / Analytical)", font=ctk.CTkFont(size=11, slant="italic"), text_color="#00FF66")
        self.temp_label.pack(anchor="w", padx=10, pady=(0, 15))

        ctk.CTkLabel(self.left_panel, text="Initial Agent Directives", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=10)
        self.prompt_text = ctk.CTkTextbox(self.left_panel, height=140)
        self.prompt_text.pack(fill="x", padx=10, pady=(5, 15))
        self.prompt_text.insert("0.0", "Compare the provided documents and extract common numbers.")

        self.run_btn = ctk.CTkButton(self.left_panel, text="▶ Initialize Analysis", font=ctk.CTkFont(size=14, weight="bold"), height=45, fg_color="#1d4ed8", hover_color="#1e3a8a", command=self.start_pipeline)
        self.run_btn.pack(fill="x", padx=10, pady=20)
        
        # New Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.left_panel, mode="indeterminate")
        self.progress_bar.set(0)
        # We will pack this dynamically when processing
        
        # New Settings Button
        self.settings_btn = ctk.CTkButton(self.left_panel, text="⚙ Application Settings", command=self.open_settings, fg_color="#374151", hover_color="#4b5563")
        self.settings_btn.pack(fill="x", padx=10, pady=10)

        # Right Panel (Tabs)
        self.right_tabs = ctk.CTkTabview(self, corner_radius=10)
        self.right_tabs.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        
        self.tab_chat = self.right_tabs.add("Interactive Workspace")
        self.tab_terminal = self.right_tabs.add("System Terminal")  # NEW TAB
        self.tab_preview = self.right_tabs.add("Content Preview")

        # Chat Workspace Setup
        self.tab_chat.grid_rowconfigure(0, weight=1)
        self.tab_chat.grid_columnconfigure(0, weight=1)
        self.chat_log = ctk.CTkTextbox(self.tab_chat, fg_color="#0a0a0a", text_color="#d1d5db", font=ctk.CTkFont(family="Consolas", size=13))
        self.chat_log.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.chat_log.insert("0.0", f"--- Session Initialized [ID: {self.session_id}] ---\nSelect files and click Initialize Analysis to begin.\n\n")

        input_frame = ctk.CTkFrame(self.tab_chat, fg_color="transparent")
        input_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        input_frame.grid_columnconfigure(0, weight=1)
        self.chat_input = ctk.CTkEntry(input_frame, placeholder_text="Ask follow-up questions or instruct the agent to write scripts...", font=ctk.CTkFont(size=13))
        self.chat_input.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=8)
        self.chat_input.bind("<Return>", lambda event: self.send_followup())
        self.send_btn = ctk.CTkButton(input_frame, text="Send", font=ctk.CTkFont(weight="bold"), width=80, command=self.send_followup)
        self.send_btn.grid(row=0, column=1, sticky="e", ipady=8)
        
        # System Terminal Setup
        self.tab_terminal.grid_rowconfigure(0, weight=1)
        self.tab_terminal.grid_columnconfigure(0, weight=1)
        self.terminal_log = ctk.CTkTextbox(self.tab_terminal, fg_color="#000000", text_color="#00FF00", font=ctk.CTkFont(family="Consolas", size=12))
        self.terminal_log.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.terminal_log.insert("0.0", "--- SYSTEM TERMINAL LOGS ---\nAll execution traces and system states will appear here.\n\n")

        # Preview Setup
        self.tab_preview.grid_rowconfigure(0, weight=1)
        self.tab_preview.grid_columnconfigure(0, weight=1)
        self.preview_box = ctk.CTkTextbox(self.tab_preview, fg_color="#1e1e1e", text_color="#e5e7eb", font=ctk.CTkFont(family="Consolas", size=11))
        self.preview_box.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.preview_box.insert("0.0", "Click a file from the left panel to preview its raw extraction.")

    # --- UI Helpers ---
    def open_settings(self):
        settings_win = ctk.CTkToplevel(self)
        settings_win.title("Settings")
        settings_win.geometry("400x300")
        settings_win.grab_set() # modal
        
        ctk.CTkLabel(settings_win, text="Ollama Inference Model", font=ctk.CTkFont(weight="bold")).pack(pady=(20, 5))
        if not hasattr(self, 'model_var'):
            self.model_var = ctk.StringVar(value="nemotron-3-ultra:cloud")
        ctk.CTkEntry(settings_win, textvariable=self.model_var, width=300).pack(pady=5)
        
        ctk.CTkLabel(settings_win, text="Ollama Map-Reduce Model", font=ctk.CTkFont(weight="bold")).pack(pady=(15, 5))
        if not hasattr(self, 'map_model_var'):
            self.map_model_var = ctk.StringVar(value="gpt-oss:120b-cloud")
        ctk.CTkEntry(settings_win, textvariable=self.map_model_var, width=300).pack(pady=5)
        
        ctk.CTkButton(settings_win, text="Save & Close", command=settings_win.destroy).pack(pady=30)

    def toggle_export(self):
        state = "normal" if self.export_var.get() else "disabled"
        self.output_entry.configure(state=state)
        self.browse_out_btn.configure(state=state)

    def update_temp_label(self, value):
        val = round(value, 2)
        desc = "Strict" if val <= 0.2 else "Balanced" if val <= 0.5 else "Creative"
        self.temp_label.configure(text=f"Current: {val} ({desc})")

    def toggle_processing_state(self, processing: bool):
        self.is_processing = processing
        if processing:
            self.run_btn.configure(state="disabled")
            self.send_btn.configure(state="disabled")
            self.progress_bar.pack(fill="x", padx=10, pady=(10, 0))
            self.progress_bar.start()
        else:
            self.run_btn.configure(state="normal")
            self.send_btn.configure(state="normal")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def add_files(self):
        filenames = filedialog.askopenfilenames(filetypes=[("All Supported Files", "*.docx;*.pdf;*.xlsx;*.csv;*.txt;*.png;*.jpg;*.jpeg")])
        for f in filenames:
            if f not in self.selected_files:
                self.selected_files.append(f)
                btn = ctk.CTkButton(
                    self.file_list_frame, text=f"📄 {os.path.basename(f)}", anchor="w", 
                    fg_color="transparent", text_color="#e5e7eb", hover_color="#374151",
                    command=lambda path=f: self.load_preview(path)
                )
                btn.pack(fill="x", pady=2, padx=5)
        if self.selected_files and not self.output_path_var.get():
            self.output_path_var.set(os.path.join(os.path.dirname(self.selected_files[0]), "Agentic_Report.docx"))

    def load_preview(self, file_path):
        self.right_tabs.set("Content Preview")
        self.preview_box.delete("0.0", "end")
        self.preview_box.insert("end", f"Loading extraction preview for {os.path.basename(file_path)}...\n")
        self.update()
        try:
            doc_streams = extract_dual_stream_from_file(file_path)
            preview_content = f"=== FILE: {os.path.basename(file_path)} ===\n\n--- SEMANTIC RAW STREAM (First 3000 chars) ---\n"
            preview_content += doc_streams['raw'][:3000]
            if len(doc_streams['raw']) > 3000: preview_content += "\n\n...[TRUNCATED FOR PREVIEW]...\n"
            self.preview_box.delete("0.0", "end")
            self.preview_box.insert("end", preview_content)
        except Exception as e:
            self.preview_box.delete("0.0", "end")
            self.preview_box.insert("end", f"Error extracting file:\n{e}")

    def clear_files(self):
        self.selected_files.clear()
        for widget in self.file_list_frame.winfo_children(): widget.destroy()

    def browse_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if filename: self.output_path_var.set(filename)

    # --- Chat & Logging Helpers ---
    def append_to_terminal(self, text):
        def _append():
            timestamp = datetime.datetime.now().strftime('%H:%M:%S')
            self.terminal_log.insert("end", f"[{timestamp}] ⚙️ {text}\n")
            self.terminal_log.see("end")
        self.after(0, _append)

    def append_to_chat(self, sender, text, tag=None):
        if sender == "System":
            # Reroute System logs to terminal directly
            self.append_to_terminal(text)
            return

        def _append():
            timestamp = datetime.datetime.now().strftime('%H:%M:%S')
            prefix = f"[{timestamp}] "
            
            if sender == "User": 
                msg = f"\n{prefix}👤 You:\n{text}\n"
                self.chat_log.insert("end", msg)
                self.chat_log.see("end")
            elif sender == "Agent" and not tag:
                # Initialization of agent message
                msg = f"\n{prefix}🤖 Agent:\n{text}\n"
                self.chat_log.insert("end", msg)
                self.chat_log.see("end")
            elif sender == "Agent" and tag == "stream":
                # Render buffered chunks for performance
                self.chat_log.insert("end", text)
                self.chat_log.see("end")

        self.after(0, _append)

    # --- Self-Healing Code Execution Engine (V14 Isolated) ---
    def execute_agent_code(self, llm_response, max_retries=3):
        """Finds Python blocks, executes them in a TempDir, and recursively self-corrects via LLM."""
        code_blocks = re.findall(r'```python\n(.*?)\n```', llm_response, re.DOTALL)
        if not code_blocks:
            return

        if not hasattr(self, 'model_var'):
            self.model_var = ctk.StringVar(value="nemotron-3-ultra:cloud")
        model_name = self.model_var.get()

        with tempfile.TemporaryDirectory() as temp_dir:
            self.append_to_terminal(f"Created isolated execution environment: {temp_dir}")
            
            for idx, code in enumerate(code_blocks):
                script_path = os.path.join(temp_dir, f"agent_script_{self.session_id}_{idx}.py")
                current_code = code
                success = False
                attempt = 1
                
                while attempt <= max_retries and not success:
                    self.append_to_terminal(f"Executing Python script (Attempt {attempt}/{max_retries})...")
                    try:
                        with open(script_path, "w", encoding="utf-8") as f:
                            f.write(current_code)
                        
                        # Added timeout for safety
                        result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, cwd=temp_dir, timeout=60)
                        
                        if result.returncode == 0:
                            output_msg = result.stdout.strip() if result.stdout else "Execution completed successfully with no output."
                            self.append_to_terminal(f"✅ Script successful.\nOutput:\n{output_msg}")
                            success = True
                        else:
                            error_msg = result.stderr.strip()
                            self.append_to_terminal(f"❌ Script failed.\nError:\n{error_msg}")
                            
                            if attempt < max_retries:
                                self.append_to_terminal("Initiating autonomous self-healing... Routing traceback to LLM.")
                                
                                fix_prompt = (
                                    f"You wrote a Python script that failed to execute. Here is the traceback error:\n\n"
                                    f"```\n{error_msg}\n```\n\n"
                                    f"Here is the broken code you wrote:\n\n"
                                    f"```python\n{current_code}\n```\n\n"
                                    f"Fix the syntax or logic error and output the ENTIRE corrected script inside a single ```python block. "
                                    f"Do not include conversational filler, only the fixed code."
                                )
                                
                                fix_stream = ollama.chat(
                                    model=model_name,
                                    messages=[{'role': 'user', 'content': fix_prompt}],
                                    options={'num_ctx': 32768, 'temperature': 0.1},
                                    stream=False
                                )
                                
                                new_response = fix_stream['message']['content']
                                new_code_blocks = re.findall(r'```python\n(.*?)\n```', new_response, re.DOTALL)
                                
                                if new_code_blocks:
                                    current_code = new_code_blocks[0]
                                    self.append_to_terminal("Received patched code from agent. Retrying execution...")
                                else:
                                    self.append_to_terminal("Self-healing failed: Agent did not return a valid Python block.")
                                    break
                            else:
                                self.append_to_terminal("Max retries reached. Script execution aborted.")
                    except subprocess.TimeoutExpired:
                        self.append_to_terminal(f"❌ Script execution timed out (exceeded 60s).")
                        break
                    except Exception as e:
                        self.append_to_terminal(f"❌ System Error running script: {str(e)}")
                        break
                    
                    attempt += 1

    # --- Core Pipeline Execution ---
    def start_pipeline(self):
        if self.is_processing: return
        
        directives = self.prompt_text.get("0.0", "end").strip()
        temperature = round(self.temp_var.get(), 2)

        if not self.selected_files:
            messagebox.showerror("Error", "Please add at least one input document.")
            return
        
        if self.export_var.get() and self.output_path_var.get():
            out_path = self.output_path_var.get().strip().replace("\\", "/")
            directives += (
                f"\n\nCRITICAL DIRECTIVE: You MUST write a complete, executable Python script using the `python-docx` "
                f"library to generate the highly professional final report. The script must save the document exactly to '{out_path}'. "
                f"Wrap the code inside a standard ```python code block."
            )

        self.right_tabs.set("Interactive Workspace")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self.run_initial_agent, args=(directives, temperature), daemon=True).start()

    def run_initial_agent(self, directives, temperature):
        try:
            self.append_to_terminal(f"Extracting and tagging {len(self.selected_files)} files...")
            self.db.log_message(self.session_id, "user", directives)
            self.append_to_chat("User", directives)

            master_raw = ""
            for file_path in self.selected_files:
                doc_streams = extract_dual_stream_from_file(file_path, log_callback=lambda m: self.append_to_terminal(m))
                master_raw += f"\n<file name=\"{os.path.basename(file_path)}\">\n{doc_streams['raw']}\n</file>\n"
            
            TOKEN_LIMIT = 20000 
            if len(master_raw) > (TOKEN_LIMIT * 4):
                self.append_to_terminal(f"Data massive (~{len(master_raw)//4} tokens). Initiating Map-Reduce block chunking...")
                chunks = chunk_text(master_raw, TOKEN_LIMIT)
                aggregated_insights = ""
                
                if not hasattr(self, 'map_model_var'):
                    self.map_model_var = ctk.StringVar(value="gpt-oss:120b-cloud")
                map_model = self.map_model_var.get()
                
                for idx, chunk in enumerate(chunks):
                    self.append_to_terminal(f"Mapping Chunk {idx+1}/{len(chunks)} via {map_model}...")
                    chunk_prompt = f"Directive: {directives}\nExtract only the required targets from this specific data chunk:\n{chunk}"
                    
                    chunk_stream = ollama.chat(
                        model=map_model,
                        messages=[{'role': 'user', 'content': chunk_prompt}],
                        options={'num_ctx': 65536, 'temperature': temperature},
                        stream=False
                    )
                    aggregated_insights += f"\n[Chunk {idx+1} Findings]:\n{chunk_stream['message']['content']}\n"
                
                self.document_context = aggregated_insights
            else:
                self.document_context = master_raw

            laya_strategy = "Default Synthesis"
            if decision_engine:
                self.append_to_terminal("Consulting LAYA Decision Engine for synthesis strategy...")
                try:
                    laya_result = decision_engine.predict(
                        state=f"User Request:\n{directives}",
                        questions={
                            "strategy": {
                                "type": "choice", 
                                "instructions": "Determine the required analytical approach.", 
                                "criteria": {"code_execution": "Requires writing a python script", "data_extraction": "Pulling specific values", "comparison": "Comparing files"}
                            }
                        }
                    )
                    laya_strategy = laya_result.get("answers", laya_result)
                except Exception as e:
                    self.append_to_terminal(f"LAYA engine failed, proceeding with default. Error: {e}")

            if not hasattr(self, 'model_var'):
                self.model_var = ctk.StringVar(value="nemotron-3-ultra:cloud")
            model_name = self.model_var.get()
            self.append_to_terminal(f"{model_name} synthesizing final response (Temp: {temperature})...")
            
            system_prompt = f"You are a master data agent. Synthesize a response based on the provided insights. Follow the strategy: {laya_strategy}\n\n=== SOURCE/AGGREGATED DATA ===\n{self.document_context}"
            
            response_stream = ollama.chat(
                model=model_name,
                messages=[{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': directives}],
                options={'num_ctx': 65536, 'temperature': temperature},
                stream=True
            )
            
            self.append_to_chat("Agent", "") 
            full_response = ""
            
            # Stream Buffering
            buffer = ""
            for chunk in response_stream:
                token = chunk['message']['content']
                full_response += token
                buffer += token
                
                if len(buffer) > 15 or '\n' in buffer:
                    self.append_to_chat("Agent", buffer, tag="stream")
                    buffer = ""
            
            if buffer:
                self.append_to_chat("Agent", buffer, tag="stream")
                
            self.db.log_message(self.session_id, "agent", full_response)
            self.execute_agent_code(full_response)

        except Exception as e:
            self.append_to_terminal(f"CRITICAL ERROR: {str(e)}")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))

    def send_followup(self):
        if self.is_processing: return
        user_text = self.chat_input.get().strip()
        if not user_text: return
        if not self.document_context:
            messagebox.showwarning("Warning", "Initialize an analysis first.")
            return

        self.chat_input.delete(0, "end")
        self.after(0, lambda: self.toggle_processing_state(True))
        threading.Thread(target=self._process_followup, args=(user_text,), daemon=True).start()

    def _process_followup(self, user_text):
        try:
            self.append_to_chat("User", user_text)
            self.db.log_message(self.session_id, "user", user_text)
            
            history_records = self.db.get_session_history(self.session_id)
            messages = [{'role': 'system', 'content': f"You are a helpful data agent. Active context findings:\n{self.document_context}. If asked to write a Python script, put it in a ```python block."}]
            for record in history_records:
                role = "assistant" if record["role"] == "agent" else record["role"]
                messages.append({'role': role, 'content': record["content"]})

            if not hasattr(self, 'model_var'):
                self.model_var = ctk.StringVar(value="nemotron-3-ultra:cloud")
            model_name = self.model_var.get()
            self.append_to_terminal("Reasoning follow-up query...")
            
            response_stream = ollama.chat(
                model=model_name,
                messages=messages,
                options={'num_ctx': 65536, 'temperature': round(self.temp_var.get(), 2)},
                stream=True
            )
            
            self.append_to_chat("Agent", "")
            full_response = ""
            buffer = ""
            for chunk in response_stream:
                token = chunk['message']['content']
                full_response += token
                buffer += token
                
                if len(buffer) > 15 or '\n' in buffer:
                    self.append_to_chat("Agent", buffer, tag="stream")
                    buffer = ""
            
            if buffer:
                self.append_to_chat("Agent", buffer, tag="stream")
                
            self.db.log_message(self.session_id, "agent", full_response)
            self.execute_agent_code(full_response)

        except Exception as e:
            self.append_to_terminal(f"ERROR: {str(e)}")
        finally:
            self.after(0, lambda: self.toggle_processing_state(False))

if __name__ == "__main__":
    app = AgenticStudioApp()
    app.mainloop()
