import os
import threading
import customtkinter as ctk
from tkinter import filedialog

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("[WARN] tkinterdnd2 not found, drag & drop disabled")

# Safety checks for critical document library dependencies to prevent crashes
try:
    import mammoth
    HAS_MAMMOTH = True
except ImportError:
    HAS_MAMMOTH = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import pandas
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

MODE_DEPENDENCIES = {
    "MD -> Excel":  [("pandas", HAS_PANDAS), ("openpyxl", HAS_OPENPYXL)],
    "MD -> Word":   [("python-docx", HAS_DOCX)],
    "Excel -> MD":  [("pandas", HAS_PANDAS), ("openpyxl", HAS_OPENPYXL)],
    "Word -> MD":   [("mammoth", HAS_MAMMOTH)],
}

from src.core.extractors import extract_excel_to_md, extract_word_to_md
from src.core.converters import md_to_excel_from_text, md_to_word_from_text, save_markdown_from_text

# ── Conversion modes config ───────────────────────────────────────────────────

MODES = {
    "MD -> Excel":  {"in_ext": ".md",   "out_ext": ".xlsx", "in_label": "File .md",   "out_label": "Save .xlsx"},
    "MD -> Word":   {"in_ext": ".md",   "out_ext": ".docx", "in_label": "File .md",   "out_label": "Save .docx"},
    "Excel -> MD":  {"in_ext": ".xlsx", "out_ext": ".md",   "in_label": "File .xlsx", "out_label": "Save .md"},
    "Word -> MD":   {"in_ext": ".docx", "out_ext": ".md",   "in_label": "File .docx", "out_label": "Save .md"},
}

IN_FILETYPES = {
    ".md":   [("Markdown", "*.md"), ("All Files", "*.*")],
    ".xlsx": [("Excel", "*.xlsx *.xls"), ("All Files", "*.*")],
    ".docx": [("Word", "*.docx"), ("All Files", "*.*")],
}

OUT_FILETYPES = {
    ".xlsx": [("Excel", "*.xlsx")],
    ".docx": [("Word",  "*.docx")],
    ".md":   [("Markdown", "*.md")],
}

BaseClass = TkinterDnD.Tk if HAS_DND else ctk.CTk

class App(BaseClass): # type: ignore
    def __init__(self):
        super().__init__()
        self.title("Document Converter Workspace")
        self.geometry("1200x780")
        self.resizable(True, True)
        
        # Configure app-wide variables
        self.in_path = ctk.StringVar()
        self.out_path = ctk.StringVar()
        self.mode_var = ctk.StringVar(value="MD -> Excel")
        
        self._build_ui()
        self._on_mode_change()

    def _build_ui(self):
        # Grid Configuration for main workspace
        self.rowconfigure(0, weight=0) # Header
        self.rowconfigure(1, weight=1) # Main workspace split
        self.columnconfigure(0, weight=1)

        # 1. Sleek Header Panel
        header_frame = ctk.CTkFrame(self, fg_color="#1e1e24", height=70, corner_radius=0)
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header_frame.grid_propagate(False)
        
        title_lbl = ctk.CTkLabel(
            header_frame, text="Document Converter Workspace",
            font=ctk.CTkFont(family="Arial", size=20, weight="bold"), text_color="#ffffff"
        )
        title_lbl.pack(side="left", padx=25, pady=10)
        
        subtitle_lbl = ctk.CTkLabel(
            header_frame, text="Multipurpose document editing and conversion workspace",
            font=ctk.CTkFont(family="Arial", size=13, slant="italic"), text_color="#8a8a9e"
        )
        subtitle_lbl.pack(side="left", padx=10, pady=16)

        # 2. Main Workspace Split (Left Pane & Right Pane)
        workspace = ctk.CTkFrame(self, fg_color="transparent")
        workspace.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)
        workspace.columnconfigure(0, weight=6) # Left Pane gets more weight
        workspace.columnconfigure(1, weight=5) # Right Pane
        workspace.rowconfigure(0, weight=1)

        # ── LEFT PANE: Input Editor & Overview ────────────────────────────────
        left_pane = ctk.CTkFrame(workspace, fg_color="#18181c", corner_radius=12, border_width=1, border_color="#2c2c35")
        left_pane.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        left_pane.rowconfigure(0, weight=0) # Label
        left_pane.rowconfigure(1, weight=0) # Drag Drop Info
        left_pane.rowconfigure(2, weight=1) # Textbox Editor
        left_pane.rowconfigure(3, weight=0) # Footer stats
        left_pane.columnconfigure(0, weight=1)

        editor_title = ctk.CTkLabel(
            left_pane, text="INPUT EDITOR & OVERVIEW (MARKDOWN / TEXT)",
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"), text_color="#3a86ff"
        )
        editor_title.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 5))

        # File Load Area / Drag Zone
        self.load_bar = ctk.CTkFrame(left_pane, fg_color="#1c1c24", corner_radius=8, height=45)
        self.load_bar.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        self.load_bar.pack_propagate(False)
        
        self.drop_lbl = ctk.CTkLabel(
            self.load_bar, text="Drag & drop file here or click 'Browse' to load content...",
            font=ctk.CTkFont(family="Arial", size=12), text_color="#7eb8f5"
        )
        self.drop_lbl.pack(side="left", padx=15, fill="both", expand=True)

        btn_browse_in = ctk.CTkButton(
            self.load_bar, text="Browse", width=75, height=28,
            font=ctk.CTkFont(family="Arial", size=12, weight="bold"), fg_color="#2980b9", hover_color="#3498db",
            command=self._browse_input
        )
        btn_browse_in.pack(side="right", padx=8, pady=8)

        # Register Drag & Drop
        if HAS_DND:
            for w in (self.load_bar, self.drop_lbl):
                w.drop_target_register(DND_FILES) # type: ignore
                w.dnd_bind("<<Drop>>", self._on_drop) # type: ignore

        # The Live Monospace Text Editor (with Undo/Redo tracking)
        self.editor = ctk.CTkTextbox(
            left_pane, fg_color="#111115", text_color="#f8f8f2",
            font=ctk.CTkFont(family="Consolas", size=13),
            border_width=1, border_color="#2c2c35", corner_radius=8,
            undo=True
        )
        self.editor.grid(row=2, column=0, sticky="nsew", padx=15, pady=8)
        self.editor.bind("<KeyRelease>", self._update_counts)

        # Bind custom Undo/Redo events to catch exceptions and prevent duplicate actions
        self.editor.bind("<Control-z>", self._undo)
        self.editor.bind("<Control-Z>", self._undo)
        self.editor.bind("<Control-y>", self._redo)
        self.editor.bind("<Control-Y>", self._redo)
        self.editor.bind("<Control-Shift-z>", self._redo)
        self.editor.bind("<Control-Shift-Z>", self._redo)

        # Editor Footer (Stats & Actions)
        editor_footer = ctk.CTkFrame(left_pane, fg_color="transparent")
        editor_footer.grid(row=3, column=0, sticky="ew", padx=15, pady=(5, 12))
        
        self.char_lbl = ctk.CTkLabel(
            editor_footer, text="Characters: 0", font=ctk.CTkFont(family="Arial", size=12), text_color="#8a8a9e"
        )
        self.char_lbl.pack(side="left", padx=5)

        self.word_lbl = ctk.CTkLabel(
            editor_footer, text=" |  Words: 0", font=ctk.CTkFont(family="Arial", size=12), text_color="#8a8a9e"
        )
        self.word_lbl.pack(side="left", padx=5)

        btn_clear = ctk.CTkButton(
            editor_footer, text="Clear All", width=70, height=24, fg_color="#c0392b", hover_color="#e74c3c",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            command=self._clear_editor
        )
        btn_clear.pack(side="right", padx=5)

        btn_redo = ctk.CTkButton(
            editor_footer, text="Redo", width=55, height=24, fg_color="#34495e", hover_color="#415b76",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            command=self._redo
        )
        btn_redo.pack(side="right", padx=5)

        btn_undo = ctk.CTkButton(
            editor_footer, text="Undo", width=55, height=24, fg_color="#34495e", hover_color="#415b76",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            command=self._undo
        )
        btn_undo.pack(side="right", padx=5)


        # ── RIGHT PANE: Output Configuration, Preview & Actions ───────────────
        right_pane = ctk.CTkFrame(workspace, fg_color="#18181c", corner_radius=12, border_width=1, border_color="#2c2c35")
        right_pane.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        right_pane.rowconfigure(0, weight=0) # Title
        right_pane.rowconfigure(1, weight=0) # Config panel
        right_pane.rowconfigure(2, weight=1) # Output preview / Logs textbox
        right_pane.rowconfigure(3, weight=0) # Convert Button & Status
        right_pane.columnconfigure(0, weight=1)

        output_title = ctk.CTkLabel(
            right_pane, text="OUTPUT CONFIGURATION & EXPORT",
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"), text_color="#2ecc71"
        )
        output_title.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 5))

        # Config Panel
        config_frame = ctk.CTkFrame(right_pane, fg_color="#1c1c24", corner_radius=8)
        config_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        config_frame.columnconfigure(0, weight=1)

        # Mode row
        row_mode = ctk.CTkFrame(config_frame, fg_color="transparent")
        row_mode.pack(fill="x", padx=12, pady=(10, 5))
        ctk.CTkLabel(row_mode, text="Mode:", width=70, anchor="w", font=ctk.CTkFont(family="Arial", size=12, weight="bold")).pack(side="left")
        self.mode_menu = ctk.CTkOptionMenu(
            row_mode, values=list(MODES.keys()), variable=self.mode_var,
            width=200, font=ctk.CTkFont(family="Arial", size=12),
            command=self._on_mode_change
        )
        self.mode_menu.pack(side="left", padx=5)

        # Path row
        row_out = ctk.CTkFrame(config_frame, fg_color="transparent")
        row_out.pack(fill="x", padx=12, pady=(5, 10))
        self.lbl_out = ctk.CTkLabel(row_out, text="Output:", width=70, anchor="w", font=ctk.CTkFont(family="Arial", size=12, weight="bold"))
        self.lbl_out.pack(side="left")
        
        self.entry_out = ctk.CTkEntry(
            row_out, textvariable=self.out_path, placeholder_text="Select save location...",
            font=ctk.CTkFont(family="Arial", size=12)
        )
        self.entry_out.pack(side="left", fill="x", expand=True, padx=5)
        
        btn_browse_out = ctk.CTkButton(
            row_out, text="Browse", width=65, height=28,
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"), fg_color="#2980b9", hover_color="#3498db",
            command=self._browse_output
        )
        btn_browse_out.pack(side="right", padx=2)

        # Output Preview & Logs Workspace
        self.preview_box = ctk.CTkTextbox(
            right_pane, fg_color="#111115", text_color="#a0a0b2",
            font=ctk.CTkFont(family="Consolas", size=12),
            border_width=1, border_color="#2c2c35", corner_radius=8
        )
        self.preview_box.grid(row=2, column=0, sticky="nsew", padx=15, pady=8)
        self._write_preview("SYSTEM READY\n\n- Drag & drop your Markdown, Excel, or Word file into the left pane.\n- The smart extractor will automatically parse your file into Markdown for you to preview, edit, or delete any characters before exporting to a new format.")

        # Large Action Button & Status Info
        action_frame = ctk.CTkFrame(right_pane, fg_color="transparent")
        action_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=(5, 12))
        action_frame.columnconfigure(0, weight=3)
        action_frame.columnconfigure(1, weight=2)

        self.btn_convert = ctk.CTkButton(
            action_frame, text="CONVERT & SAVE", height=48,
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            fg_color="#2ecc71", hover_color="#27ae60", text_color="#ffffff",
            command=self._run_conversion
        )
        self.btn_convert.grid(row=0, column=0, sticky="ew", padx=(0, 5), pady=(0, 5))

        self.btn_open_file = ctk.CTkButton(
            action_frame, text="OPEN CREATED FILE", height=48,
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            fg_color="#1c1c24", hover_color="#2c2c35", text_color="#8a8a9e",
            state="disabled",
            command=self._open_generated_file
        )
        self.btn_open_file.grid(row=0, column=1, sticky="ew", padx=(5, 0), pady=(0, 5))

        self.status_lbl = ctk.CTkLabel(
            action_frame, text="Ready to process", font=ctk.CTkFont(family="Arial", size=12, slant="italic"), text_color="gray"
        )
        self.status_lbl.grid(row=1, column=0, columnspan=2, sticky="ew")

    # ── Internal Actions & Event Handlers ─────────────────────────────────────

    def _cfg(self):
        return MODES[self.mode_var.get()]

    def _on_mode_change(self, _=None):
        cfg = self._cfg()
        self.lbl_out.configure(text=cfg["out_label"] + ":")
        
        # Proactively update the output file path if input exists
        inp = self.in_path.get().strip()
        if inp:
            self._auto_output(inp)
        else:
            self.out_path.set("")
            
        # Validate dependencies for the selected mode
        mode = self.mode_var.get()
        missing = [name for name, available in MODE_DEPENDENCIES[mode] if not available]
        
        if missing:
            self.btn_convert.configure(state="disabled", fg_color="#c0392b", text="UNAVAILABLE")
            missing_str = ", ".join(missing)
            self._set_status(f"Missing dependency: {missing_str}", "red")
            
            tooltip_msg = f"Dependency Error:\n" \
                          f"-----------------------------------\n" \
                          f"The selected mode '{mode}' is unavailable because the following required libraries are missing:\n" \
                          f"- {missing_str}\n\n" \
                          f"To enable this mode, please run the following command in your terminal:\n" \
                          f"    pip install {' '.join(missing)}\n\n" \
                          f"Then restart this application to apply."
            self._write_preview(tooltip_msg)
        else:
            self.btn_convert.configure(state="normal", fg_color="#2ecc71", text="CONVERT & SAVE")
            self._set_status("Mode changed to: " + mode, "gray")
            # If the editor has no text, reset to standard greeting
            if not self.editor.get("1.0", "end-1c").strip():
                self._write_preview("SYSTEM READY\n\n- Drag & drop your Markdown, Excel, or Word file into the left pane.\n- The smart extractor will automatically parse your file into Markdown for you to preview, edit, or delete any characters before exporting to a new format.")

    def _clear_editor(self):
        self.editor.delete("1.0", "end")
        self.editor.edit_reset()  # Reset undo stack after manual clearing
        self.in_path.set("")
        self._update_counts()
        self._write_preview("Editor is empty. You can write your own Markdown text here!")
        self._set_status("Editor cleared", "gray")
        self.drop_lbl.configure(text="Drag & drop file here or click 'Browse' to load content...", text_color="#7eb8f5")

    def _update_counts(self, _=None):
        content = self.editor.get("1.0", "end-1c")
        chars = len(content)
        words = len(content.split())
        self.char_lbl.configure(text=f"Characters: {chars}")
        self.word_lbl.configure(text=f" |  Words: {words}")

    def _auto_output(self, in_path: str):
        base = os.path.splitext(in_path)[0]
        self.out_path.set(base + self._cfg()["out_ext"])

    def _write_preview(self, msg: str):
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", msg)
        self.preview_box.configure(state="disabled")

    def _set_status(self, msg: str, color: str = "white"):
        hex_colors = {"green": "#2ecc71", "red": "#e74c3c", "gray": "#8a8a9e", "orange": "#e67e22"}
        target_color = hex_colors.get(color, color)
        self.status_lbl.configure(text=msg, text_color=target_color)

    # ── Load file on drop or browse ───────────────────────────────────────────

    def _on_drop(self, event):
        raw_data = event.data
        if not raw_data:
            return
        
        # Parse TkinterDnD paths (handles spaces wrapped in {}, unicode, and multiple files)
        import re
        pattern = re.compile(r'\{([^}]+)\}|(\S+)')
        paths = [m.group(1) or m.group(2) for m in pattern.finditer(raw_data)]
        paths = [p for p in paths if p]

        if not paths:
            self._set_status("No valid files dropped", "orange")
            return

        # Since it is currently a single-document editor, load the first file
        target_path = paths[0]
        self._load_file_to_editor(target_path)

        # Notify if multiple files were dropped
        if len(paths) > 1:
            self._set_status(f"Loaded first file. Ignored {len(paths)-1} other files.", "orange")

    def _browse_input(self):
        # We allow loading any supported document type to auto-detect
        path = filedialog.askopenfilename(filetypes=[
            ("Supported Documents", "*.md *.xlsx *.xls *.docx"),
            ("Markdown (*.md)", "*.md"),
            ("Excel (*.xlsx, *.xls)", "*.xlsx *.xls"),
            ("Word (*.docx)", "*.docx"),
            ("All Files", "*.*")
        ])
        if path:
            self._load_file_to_editor(path)

    def _load_file_to_editor(self, path: str):
        if not os.path.exists(path):
            self._set_status("Input file does not exist!", "red")
            return

        # Check file size warning (>50MB) to prevent UI freezing and high memory usage
        file_size = os.path.getsize(path)
        if file_size > 50 * 1024 * 1024:
            from tkinter import messagebox
            file_size_mb = file_size / (1024 * 1024)
            confirm = messagebox.askyesno(
                title="Large File Warning",
                message=f"The selected file '{os.path.basename(path)}' is very large ({file_size_mb:.1f} MB).\n\n"
                        f"Loading large files may freeze the user interface or consume significant system memory (RAM).\n\n"
                        f"Are you sure you want to proceed?"
            )
            if not confirm:
                self._set_status("Loading cancelled", "orange")
                self._write_preview(f"Load Cancelled:\n\nFile '{os.path.basename(path)}' ({file_size_mb:.1f} MB) was not loaded because it exceeds the 50 MB safety threshold.")
                return

        ext = os.path.splitext(path)[1].lower()

        # Enforce dependency validation before attempting ingestion
        if ext == ".docx" and not HAS_MAMMOTH:
            self._set_status("Cannot load: mammoth library missing!", "red")
            tooltip_msg = f"Load Error:\n" \
                          f"-----------------------------------\n" \
                          f"The file '{os.path.basename(path)}' requires the 'mammoth' library for extraction, which is not installed.\n\n" \
                          f"To resolve this, please install it via:\n" \
                          f"    pip install mammoth"
            self._write_preview(tooltip_msg)
            return

        if ext in (".xlsx", ".xls") and (not HAS_OPENPYXL or not HAS_PANDAS):
            missing = []
            if not HAS_PANDAS: missing.append("pandas")
            if not HAS_OPENPYXL: missing.append("openpyxl")
            missing_str = " and ".join(missing)
            self._set_status(f"Cannot load: {missing_str} missing!", "red")
            tooltip_msg = f"Load Error:\n" \
                          f"-----------------------------------\n" \
                          f"The file '{os.path.basename(path)}' requires the following library/libraries to be read:\n" \
                          f"- {missing_str}\n\n" \
                          f"To resolve this, please run:\n" \
                          f"    pip install {' '.join(missing)}"
            self._write_preview(tooltip_msg)
            return

        self._set_status("Loading and extracting file content...", "orange")

        # Update paths
        self.in_path.set(path)
        self.drop_lbl.configure(text=f"Loaded successfully: {os.path.basename(path)}", text_color="#2ecc71")

        def task():
            try:
                content = ""
                detected_mode = None

                if ext == ".md":
                    detected_mode = "MD -> Excel"  # Default
                    with open(path, encoding="utf-8") as f:
                        content = f.read()
                elif ext in (".xlsx", ".xls"):
                    detected_mode = "Excel -> MD"
                    content = extract_excel_to_md(path)
                elif ext == ".docx":
                    detected_mode = "Word -> MD"
                    content = extract_word_to_md(path)
                else:
                    raise ValueError(f"File extension {ext} is not supported!")

                # Update UI on main thread
                def update_ui():
                    if detected_mode:
                        self.mode_var.set(detected_mode)
                        self._on_mode_change()
                    
                    self.editor.delete("1.0", "end")
                    self.editor.insert("1.0", content)
                    self.editor.edit_reset()  # Reset undo stack after loading new document content
                    self._update_counts()
                    
                    # Generate Overview log
                    preview_msg = f"LOADED DOCUMENT OVERVIEW:\n" \
                                  f"-----------------------------------\n" \
                                  f"- Filename: {os.path.basename(path)}\n" \
                                  f"- Original format: {ext.upper()}\n" \
                                  f"- Size: {os.path.getsize(path)} bytes\n" \
                                  f"- Auto-switched mode: {detected_mode}\n\n" \
                                  f"You can now read, edit, or delete any character in the left editor. Click 'Browse' in Output to change the save location."
                    
                    self._write_preview(preview_msg)
                    self._set_status("Loaded and extracted successfully!", "green")

                self.after(0, update_ui)
            except Exception as e:
                self.after(0, lambda: self._set_status(f"Load error: {e}", "red"))
                self.after(0, lambda: self._write_preview(f"Load error details:\n{e}"))

        threading.Thread(target=task, daemon=True).start()

    def _browse_output(self):
        ext = self._cfg()["out_ext"]
        path = filedialog.asksaveasfilename(
            defaultextension=ext, filetypes=OUT_FILETYPES[ext],
            initialfile=os.path.splitext(os.path.basename(self.in_path.get() or "document"))[0] + ext
        )
        if path:
            self.out_path.set(path)

    # ── Execute Document Generation & Export ─────────────────────────────────

    def _run_conversion(self):
        content = self.editor.get("1.0", "end-1c").strip()
        out = self.out_path.get().strip()
        mode = self.mode_var.get()

        if not content:
            self._set_status("Editor is empty, nothing to convert!", "orange")
            return
        if not out:
            self._set_status("Please choose a save path!", "orange")
            return

        # Validate Markdown tables to prevent malformed data parsing
        from src.core.validator import validate_md_tables
        warnings = validate_md_tables(content)
        if warnings:
            from tkinter import messagebox
            warn_msg = "The following issues were detected in your Markdown tables:\n\n" + \
                       "\n".join(f"- {w}" for w in warnings[:10])
            if len(warnings) > 10:
                warn_msg += f"\n- ... and {len(warnings) - 10} more warnings."
            warn_msg += "\n\nDo you want to proceed with the conversion anyway?"
            
            confirm = messagebox.askyesno(
                title="Table Validation Warning",
                message=warn_msg
            )
            if not confirm:
                self._set_status("Conversion cancelled", "orange")
                return

        if os.path.exists(out):
            from tkinter import messagebox
            confirm = messagebox.askyesno(
                title="Overwrite Confirmation",
                message=f"The file '{os.path.basename(out)}' already exists.\n\nDo you want to overwrite it?"
            )
            if not confirm:
                self._set_status("Conversion cancelled", "orange")
                return

        self.btn_convert.configure(state="disabled", text="Converting...")
        self.btn_open_file.configure(state="disabled", fg_color="#1c1c24", text_color="#8a8a9e")
        self._set_status("Processing conversion and writing file...", "orange")

        def task():
            try:
                # Direct dispatcher
                if mode == "MD -> Excel":
                    msg = md_to_excel_from_text(content, out)
                elif mode == "MD -> Word":
                    msg = md_to_word_from_text(content, out)
                elif mode in ("Excel -> MD", "Word -> MD"):
                    msg = save_markdown_from_text(content, out)
                else:
                    raise ValueError(f"Invalid mode {mode}!")

                color = "green" if msg.startswith("Exported") or msg.startswith("Word") or msg.startswith("Markdown") else "red"
                
                # Output Overview details
                log_details = f"DOCUMENT EXPORT LOG:\n" \
                              f"-----------------------------------\n" \
                              f"- Result: SUCCESS\n" \
                              f"- Conversion mode: {mode}\n" \
                              f"- Generated file: {os.path.basename(out)}\n" \
                              f"- Save folder: {os.path.dirname(out)}\n" \
                              f"- New size: {os.path.getsize(out) if os.path.exists(out) else 0} bytes\n\n" \
                              f"Click the 'OPEN CREATED FILE' button below to open and view your document directly!"

                self.after(0, lambda: self._write_preview(log_details))
                self.after(0, lambda: self._set_status(msg, color))
                
                if color == "green":
                    self.after(0, lambda: self.btn_open_file.configure(
                        state="normal", fg_color="#3a86ff", hover_color="#2563eb", text_color="#ffffff"
                    ))
            except Exception as e:
                self.after(0, lambda: self._set_status(f"Conversion error: {e}", "red"))
                self.after(0, lambda: self._write_preview(f"Conversion error details:\n{e}"))
            
            self.after(0, lambda: self.btn_convert.configure(state="normal", text="CONVERT & SAVE"))

        threading.Thread(target=task, daemon=True).start()

    def _open_generated_file(self):
        out = self.out_path.get().strip()
        if out and os.path.exists(out):
            try:
                os.startfile(out)
                self._set_status("Opened file in default application", "green")
            except Exception as e:
                self._set_status(f"Failed to open file: {e}", "red")
        else:
            self._set_status("File does not exist or has not been created yet!", "orange")

    def _undo(self, event=None):
        try:
            self.editor.edit_undo()
        except Exception:
            pass
        return "break"

    def _redo(self, event=None):
        try:
            self.editor.edit_redo()
        except Exception:
            pass
        return "break"
