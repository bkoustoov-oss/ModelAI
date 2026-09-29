import datetime
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

# Force offline caching to load LAYA instantly without downloading or network verification
os.environ["HF_HUB_OFFLINE"] = "0"
os.environ["TRANSFORMERS_OFFLINE"] = "0"

import customtkinter as ctk
from docx import Document
from docx.shared import Inches, Pt, RGBColor
import laya
import ollama

# Configure modern dark theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

print("Loading LAYA System 1 Decision Engine (Offline Cache)...")
decision_engine = laya.load("convaiinnovations/laya")

def add_safe_heading(doc, text, level=1):
    """Adds a stylized heading safely without depending on Word's style dictionary."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    
    if level == 1:
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        run.font.size = Pt(15)
        run.font.color.rgb = RGBColor(27, 54, 93)   # Deep Navy
    elif level == 2:
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(37, 99, 235)  # Royal Blue
    else:  # Level 3+
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        run.font.size = Pt(11.5)
        run.font.color.rgb = RGBColor(55, 65, 81)   # Slate Charcoal
    return p

class AgenticStudioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("LAYA Agentic Studio - v2")
        self.geometry("1100x750")
        self.minsize(900, 600)
        
        # Grid layout: Column 0 (Controls), Column 1 (Streaming Workspace)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # --- LEFT PANEL: CONTROLS ---
        self.left_panel = ctk.CTkFrame(self, corner_radius=10)
        self.left_panel.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        ctk.CTkLabel(self.left_panel, text="Document Orchestration", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(15, 20))

        # 1. Input File Selection
        ctk.CTkLabel(self.left_panel, text="Source Document (.docx)", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=15)
        self.input_path_var = ctk.StringVar()
        self.input_entry = ctk.CTkEntry(self.left_panel, textvariable=self.input_path_var, placeholder_text="Select input file...")
        self.input_entry.pack(fill="x", padx=15, pady=(5, 5))
        ctk.CTkButton(self.left_panel, text="Browse Input", command=self.browse_input).pack(anchor="e", padx=15, pady=(0, 15))

        # 2. Output File Destination
        ctk.CTkLabel(self.left_panel, text="Save Output As", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=15)
        self.output_path_var = ctk.StringVar()
        self.output_entry = ctk.CTkEntry(self.left_panel, textvariable=self.output_path_var, placeholder_text="Choose save destination...")
        self.output_entry.pack(fill="x", padx=15, pady=(5, 5))
        ctk.CTkButton(self.left_panel, text="Browse Output", command=self.browse_output).pack(anchor="e", padx=15, pady=(0, 20))

        # 3. Custom Directives Textbox
        ctk.CTkLabel(self.left_panel, text="Agent Directives", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=15)
        self.prompt_text = ctk.CTkTextbox(self.left_panel, height=140)
        self.prompt_text.pack(fill="x", padx=15, pady=(5, 20))
        self.prompt_text.insert("0.0", "Analyze the document, identify structural gaps, and synthesize a polished visual executive synopsis focusing on core actionable optimizations. Do not use markdown tables. Format all data using structured bullet points and bold text.")

        # 4. Action Button
        self.run_btn = ctk.CTkButton(
            self.left_panel, 
            text="▶ Run Agentic Pipeline", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            height=45, 
            fg_color="#1d4ed8", 
            hover_color="#1e3a8a", 
            command=self.start_pipeline
        )
        self.run_btn.pack(fill="x", padx=15, pady=10)

        # --- RIGHT PANEL: AGENT WORKSPACE ---
        self.right_panel = ctk.CTkFrame(self, corner_radius=10, fg_color="#111111")
        self.right_panel.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")

        ctk.CTkLabel(
            self.right_panel, 
            text="Agent Workspace (Live Streaming)", 
            text_color="#00FF66", 
            font=ctk.CTkFont(family="Consolas", size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=10)

        self.console = ctk.CTkTextbox(
            self.right_panel, 
            fg_color="#0a0a0a", 
            text_color="#d1d5db", 
            font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.console.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.console.insert("0.0", "> Agent initialized in offline mode. Ready for execution...\n")

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.input_path_var.set(filename)
            base, ext = os.path.splitext(filename)
            self.output_path_var.set(f"{base}_Orchestrated{ext}")

    def browse_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.output_path_var.set(filename)

    def log(self, message, style="system"):
        """Thread-safe UI logging with distinct visual styles."""
        def _append():
            timestamp = datetime.datetime.now().strftime('%H:%M:%S')
            prefix = f"[{timestamp}] "
            if style == "header":
                msg = f"\n{'-'*50}\n{prefix}⚡ {message}\n{'-'*50}\n"
            elif style == "stream":
                msg = message
            elif style == "success":
                msg = f"\n{prefix}✅ {message}\n"
            else:
                msg = f"\n{prefix} {message}\n"
            
            self.console.insert("end", msg)
            self.console.see("end")
        self.after(0, _append)

    def start_pipeline(self):
        in_path = self.input_path_var.get().strip()
        out_path = self.output_path_var.get().strip()
        directives = self.prompt_text.get("0.0", "end").strip()

        if not in_path or not os.path.exists(in_path):
            messagebox.showerror("Error", "Please select a valid input document.")
            return
        if not out_path:
            messagebox.showerror("Error", "Please specify an output save location.")
            return

        self.run_btn.configure(state="disabled", fg_color="#4b5563")
        threading.Thread(target=self.run_agent, args=(in_path, out_path, directives), daemon=True).start()

    def run_agent(self, in_path, out_path, directives):
        try:
            # 1. READ DOCUMENT
            self.log(f"Ingesting Document: {os.path.basename(in_path)}", "header")
            doc = Document(in_path)
            doc_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            
            # 2. GENERATE IDEAS (LIVE TOKEN STREAMING)
            self.log("Agent [GPT OSS 120B] Initiating reasoning and ideation...", "header")
            
            response_stream = ollama.chat(
                model='gpt-oss:120b-cloud',
                messages=[{'role': 'user', 'content': f"Directive: {directives}\n\nContext:\n{doc_text}"}],
                options={'num_ctx': 65536, 'temperature': 0.7, 'thinking': True},
                stream=True 
            )
            
            raw_ideas = ""
            for chunk in response_stream:
                token = chunk['message']['content']
                raw_ideas += token
                self.log(token, style="stream")

            # 3. LAYA DECISION ENGINE GATE
            self.log("Agent [LAYA] Evaluating generated insight matrices...", "header")
            result = decision_engine.predict(
                state=f"Generated Suggestions:\n{raw_ideas}",
                questions={
                    "is_critical": {"type": "noul", "instructions": "Is this critical for execution?"},
                    "category": {"type": "choice", "instructions": "Classify.", "criteria": {"route_safety": "Safety", "logistics": "Logistics", "general": "General"}}
                }
            )
            laya_decisions = result.get("answers", result)
            self.log(f"LAYA Router Decisions:\n{laya_decisions}", style="system")

            # 4. FINAL POLISHED SYNTHESIS (LIVE TOKEN STREAMING)
            self.log("Agent [Nemotron Ultra] Synthesizing final structured document payload...", "header")
            
            final_stream = ollama.chat(
                model='nemotron-3-ultra:cloud',
                messages=[{'role': 'user', 'content': f"Format a polished, final list based on directive ({directives}):\n\nIdeas: {raw_ideas}\nDecisions: {laya_decisions}"}],
                options={'num_ctx': 32768, 'temperature': 0.3},
                stream=True
            )
            
            synopsis_text = ""
            for chunk in final_stream:
                token = chunk['message']['content']
                synopsis_text += token
                self.log(token, style="stream")

            # 5. DOCUMENT WRITING & CRASH-PROOF MARKDOWN PARSING
            self.log("Parsing Markdown into native Word formatting...", "header")
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Safe Top visual header
            add_safe_heading(doc, f"✨ AI Visual Synopsis & Insights ({timestamp})", level=1)

            clean_text = synopsis_text.replace('<br>', '\n').replace('<br/>', '\n')
            
            for line in clean_text.split('\n'):
                line = line.strip()
                if not line:
                    continue
                
                # Headings handled directly via custom styles (prevents KeyError crashes)
                if line.startswith('### '):
                    add_safe_heading(doc, line.replace('### ', '').replace('**', ''), level=3)
                elif line.startswith('## '):
                    add_safe_heading(doc, line.replace('## ', '').replace('**', ''), level=2)
                elif line.startswith('# '):
                    add_safe_heading(doc, line.replace('# ', '').replace('**', ''), level=1)
                
                # Paragraphs, bullet points, and inline bolding
                else:
                    p = doc.add_paragraph()
                    if line.startswith('- ') or line.startswith('* '):
                        line = line[2:].strip()
                        bullet_run = p.add_run("• ")
                        bullet_run.bold = True
                    
                    # Apply inline bold formatting via regex splitting
                    parts = re.split(r'(\*\*.*?\*\*)', line)
                    for part in parts:
                        if part.startswith('**') and part.endswith('**'):
                            bold_run = p.add_run(part[2:-2])
                            bold_run.bold = True
                        else:
                            p.add_run(part)

            doc.save(out_path)
            self.log(f"File successfully orchestrated and saved to:\n{out_path}", "success")
            
        except Exception as e:
            self.log(f"PIPELINE ERROR: {str(e)}", "header")
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal", fg_color="#1d4ed8"))

if __name__ == "__main__":
    app = AgenticStudioApp()
    app.mainloop()