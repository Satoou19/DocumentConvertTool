import os
import subprocess
import sys
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
    "Word -> MD":   [("python-docx", HAS_DOCX)],
}

from src.services.file_loader import load_document
from src.services.conversion_service import (
    convert_content,
    get_md_table_warnings,
    has_md_tables,
    is_output_locked,
)

# ── Configuration constants ───────────────────────────────────────────────────

# Giới hạn kích cỡ hiển thị trong textbox editor (500KB = ~500,000 ký tự)
EDITOR_DISPLAY_LIMIT = 500_000

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
        
        # Thiết lập kích thước cửa sổ phù hợp với nhiều loại màn hình (đặc biệt là laptop)
        window_width = 1150
        window_height = 680
        
        try:
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            # Đảm bảo y không bị âm và nhấc lên một chút để tránh taskbar
            y = max(y - 20, 15)
            self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        except Exception:
            self.geometry(f"{window_width}x{window_height}")
            
        self.resizable(True, True)
        
        # Configure app-wide variables
        self.in_path = ctk.StringVar()
        self.out_path = ctk.StringVar()
        self.mode_var = ctk.StringVar(value="MD -> Excel")
        self.full_content = ""  # Lưu toàn bộ content khi tải file lớn
        self.is_preview_blocked = False
        self.is_dirty = False
        self.is_processing = False
        
        self._build_ui()
        self._on_mode_change()
        
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

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
        workspace.columnconfigure(0, weight=6, uniform="workspace_split") # Left Pane gets more weight
        workspace.columnconfigure(1, weight=5, uniform="workspace_split") # Right Pane
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

        # File Load Area / Drag Zone (with subtle border for drag enter/leave highlights)
        self.load_bar = ctk.CTkFrame(left_pane, fg_color="#1c1c24", corner_radius=8, height=45, border_width=1, border_color="#2c2c35")
        self.load_bar.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        self.load_bar.pack_propagate(False)
        
        self.drop_lbl = ctk.CTkLabel(
            self.load_bar, text="Drag & drop file here or click 'Browse' to load content...",
            font=ctk.CTkFont(family="Arial", size=12), text_color="#7eb8f5"
        )
        self.drop_lbl.pack(side="left", padx=15, fill="both", expand=True)

        self.btn_browse_in = ctk.CTkButton(
            self.load_bar, text="Browse", width=75, height=28,
            font=ctk.CTkFont(family="Arial", size=12, weight="bold"), fg_color="#2980b9", hover_color="#3498db",
            command=self._browse_input
        )
        self.btn_browse_in.pack(side="right", padx=8, pady=8)

        # Register Drag & Drop with active drag hover animations
        if HAS_DND:
            for w in (self.load_bar, self.drop_lbl):
                w.drop_target_register(DND_FILES) # type: ignore
                w.dnd_bind("<<Drop>>", self._on_drop) # type: ignore
                w.dnd_bind("<<DragEnter>>", self._on_drag_enter) # type: ignore
                w.dnd_bind("<<DragLeave>>", self._on_drag_leave) # type: ignore

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

        self.btn_clear = ctk.CTkButton(
            editor_footer, text="Clear All", width=70, height=24, fg_color="#c0392b", hover_color="#e74c3c",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            command=self._clear_editor
        )
        self.btn_clear.pack(side="right", padx=5)

        self.btn_redo = ctk.CTkButton(
            editor_footer, text="Redo", width=55, height=24, fg_color="#34495e", hover_color="#415b76",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            command=self._redo
        )
        self.btn_redo.pack(side="right", padx=5)

        self.btn_undo = ctk.CTkButton(
            editor_footer, text="Undo", width=55, height=24, fg_color="#34495e", hover_color="#415b76",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            command=self._undo
        )
        self.btn_undo.pack(side="right", padx=5)


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
        
        self.btn_browse_out = ctk.CTkButton(
            row_out, text="Browse", width=65, height=28,
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"), fg_color="#2980b9", hover_color="#3498db",
            command=self._browse_output
        )
        self.btn_browse_out.pack(side="right", padx=2)

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

        self.progress_bar = ctk.CTkProgressBar(action_frame, mode="indeterminate", width=10, height=8, progress_color="#2ecc71")
        # Kept hidden initially until conversion begins

        self.status_lbl = ctk.CTkLabel(
            action_frame, text="Ready to process", font=ctk.CTkFont(family="Arial", size=12, slant="italic"), text_color="gray"
        )
        self.status_lbl.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(5, 0))

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
        if self.is_processing:
            return
        self.editor.configure(state="normal", text_color="#f8f8f2") # Mở khóa editor và đặt lại màu chữ
        self.editor.delete("1.0", "end")
        self.editor.edit_reset()  # Reset undo stack after manual clearing
        self.in_path.set("")
        self.full_content = ""  # Xóa toàn bộ content
        self.is_preview_blocked = False
        self.is_dirty = False
        self._update_counts()
        self._write_preview("Editor is empty. You can write your own Markdown text here!")
        self._set_status("Editor cleared", "gray")
        self.drop_lbl.configure(text="Drag & drop file here or click 'Browse' to load content...", text_color="#7eb8f5")

    def _update_counts(self, event=None):
        # Dùng full_content nếu file lớn hoặc bị chặn preview, nếu không thì lấy từ editor
        is_large_or_blocked = (self.full_content and len(self.full_content) > EDITOR_DISPLAY_LIMIT) or self.is_preview_blocked
        content = self.full_content if is_large_or_blocked else self.editor.get("1.0", "end-1c")
        
        if event is not None:
            self.is_dirty = True
            
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

    def _toggle_ui_state(self, enabled: bool):
        self.is_processing = not enabled
        state = "normal" if enabled else "disabled"
        
        if enabled:
            mode = self.mode_var.get()
            missing = [name for name, available in MODE_DEPENDENCIES[mode] if not available]
            if missing:
                self.btn_convert.configure(state="disabled", fg_color="#c0392b", text="UNAVAILABLE")
            else:
                self.btn_convert.configure(state="normal", fg_color="#2ecc71", text="CONVERT & SAVE")
        else:
            self.btn_convert.configure(state="disabled")

        self.btn_browse_in.configure(state=state)
        self.btn_clear.configure(state=state)
        self.btn_undo.configure(state=state)
        self.btn_redo.configure(state=state)
        self.mode_menu.configure(state=state)
        self.entry_out.configure(state=state)
        self.btn_browse_out.configure(state=state)
        
        if not enabled:
            self.editor.configure(state="disabled")
        else:
            is_large_file = self.full_content and len(self.full_content) > EDITOR_DISPLAY_LIMIT
            if self.is_preview_blocked or is_large_file:
                self.editor.configure(state="disabled", text_color="#8a8a9e")
            else:
                self.editor.configure(state="normal", text_color="#f8f8f2")

    # ── Load file on drop or browse ───────────────────────────────────────────

    def _on_drop(self, event):
        if self.is_processing:
            return
        self._on_drag_leave(event)  # Reset highlight borders immediately
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
        if self.is_processing:
            return
        # We allow loading any supported document type to auto-detect
        path = filedialog.askopenfilename(parent=self, filetypes=[
            ("Supported Documents", "*.md *.xlsx *.xls *.docx"),
            ("Markdown (*.md)", "*.md"),
            ("Excel (*.xlsx, *.xls)", "*.xlsx *.xls"),
            ("Word (*.docx)", "*.docx"),
            ("All Files", "*.*")
        ])
        if path:
            self._load_file_to_editor(path)

    def _load_file_to_editor(self, path: str):
        result = load_document(path)
        if not result.success:
            self.in_path.set("")
            self.full_content = ""
            self.editor.configure(state="normal")
            self.editor.delete("1.0", "end")  # Clear editor content on load failure

            if result.missing_dependencies:
                missing_str = " and ".join(result.missing_dependencies)
                self.drop_lbl.configure(text=f"Failed: {missing_str} missing", text_color="#e74c3c")
                self._set_status(f"Cannot load: {missing_str} missing!", "red")
                tooltip_msg = f"Load Error:\n" \
                              f"-----------------------------------\n" \
                              f"The file '{os.path.basename(path)}' requires the following library/libraries to be installed:\n" \
                              f"- {missing_str}\n\n" \
                              f"To resolve this, please run:\n" \
                              f"    pip install {' '.join(result.missing_dependencies)}"
                self._write_preview(tooltip_msg)
            else:
                short_err = result.error_short or "Load error"
                detailed_err = result.error_detail or "An unknown error occurred during load."
                self.drop_lbl.configure(text=f"Failed: {short_err}", text_color="#e74c3c")
                self._set_status(f"Load error: {short_err}", "red")
                self._write_preview(f"LOAD ERROR:\n\n{detailed_err}")
                from tkinter import messagebox
                messagebox.showerror(parent=self, title="File Ingestion Error", message=detailed_err)
            return

        self.is_preview_blocked = False
        file_size = os.path.getsize(path)
        file_size_mb = file_size / (1024 * 1024)

        if file_size > 5 * 1024 * 1024:
            from tkinter import messagebox
            confirm = messagebox.askyesno(
                parent=self,
                title="Large File Preview Option",
                message=f"The selected file '{os.path.basename(path)}' is larger than 5MB ({file_size_mb:.2f} MB).\n\n"
                        f"Loading a preview in the editor might cause the application to freeze.\n\n"
                        f"Do you want to enable the preview?\n\n"
                        f"- Choose 'Yes' to enable the preview (limited to first 500KB).\n"
                        f"- Choose 'No' to load the file without editor preview (recommended for performance)."
            )
            if not confirm:
                self.is_preview_blocked = True
        elif file_size > 1 * 1024 * 1024:
            from tkinter import messagebox
            confirm = messagebox.askyesno(
                parent=self,
                title="File Size Preview Option",
                message=f"The selected file '{os.path.basename(path)}' is larger than 1MB ({file_size_mb:.2f} MB).\n\n"
                        f"Loading a preview might take some time.\n\n"
                        f"Do you want to enable the preview?\n\n"
                        f"- Choose 'Yes' to enable the preview.\n"
                        f"- Choose 'No' to load the file without editor preview (keeps performance fast)."
            )
            if not confirm:
                self.is_preview_blocked = True

        self._toggle_ui_state(False)
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 5))
        self.progress_bar.start()

        self._set_status("Loading and extracting file content...", "orange")

        # Update paths (Set state to 'Loading' in orange, NOT loaded yet)
        self.in_path.set(path)
        self.drop_lbl.configure(text=f"Loading: {os.path.basename(path)}...", text_color="#e67e22")

        def task():
            try:
                content = ""
                detected_mode = None

                if not result.success:
                    raise RuntimeError(result.error_detail or result.error_short or "Failed to load document.")
                detected_mode = result.mode
                content = result.content
                ext = os.path.splitext(path)[1].lower()
                print(f"[DEBUG] Extracted content size: {len(content):,} chars ({len(content)/1024/1024:.2f} MB)")

                # Update UI on main thread
                def update_ui():
                    file_size = os.path.getsize(path)
                    file_size_mb = file_size / (1024 * 1024)

                    if detected_mode:
                        self.mode_var.set(detected_mode)
                        self._on_mode_change()
                    
                    # Giới hạn hiển thị nếu content quá lớn hoặc bị chặn preview
                    if self.is_preview_blocked:
                        display_content = f"[PREVIEW SKIPPED] This file is large ({file_size_mb:.2f} MB). Previewing has been skipped to maintain fast performance. You can still convert and save the file by clicking 'CONVERT & SAVE'."
                    else:
                        is_truncated = len(content) > EDITOR_DISPLAY_LIMIT
                        display_content = content[:EDITOR_DISPLAY_LIMIT] if is_truncated else content
                    
                    # Lưu toàn bộ content để dùng khi convert
                    self.full_content = content
                    
                    self.editor.configure(state="normal") # Đảm bảo mở khóa để cho phép xóa/chèn dữ liệu mới
                    self.editor.delete("1.0", "end")
                    self.editor.insert("1.0", display_content)
                    self.editor.edit_reset()  # Reset undo stack after loading new document content
                    
                    self.is_dirty = False
                    self._update_counts()
                    
                    preview_msg = f"LOADED DOCUMENT OVERVIEW:\n" \
                                  f"-----------------------------------\n" \
                                  f"- Filename: {os.path.basename(path)}\n" \
                                  f"- Original format: {ext.upper()}\n" \
                                  f"- Size: {file_size_mb:.2f} MB ({file_size:,} bytes)\n" \
                                  f"- Auto-switched mode: {detected_mode}\n"
                    
                    if self.is_preview_blocked:
                        truncated_size = len(content) / (1024 * 1024)
                        preview_msg += f"\n🚫  PREVIEW SKIPPED:\n" \
                                      f"- Full content: {truncated_size:.2f} MB\n" \
                                      f"- Note: Previewing is skipped to maintain fast performance.\n" \
                                      f"  The full content will be used when you click 'CONVERT & SAVE'.\n"
                    elif is_truncated:
                        truncated_size = len(content) / (1024 * 1024)
                        display_size = len(display_content) / (1024 * 1024)
                        preview_msg += f"\n⚠️  PREVIEW TRUNCATED:\n" \
                                      f"- Full content: {truncated_size:.2f} MB\n" \
                                      f"- Displayed: {display_size:.2f} MB (first 500KB)\n" \
                                      f"- Note: Only the first 500KB is shown in the editor for performance.\n" \
                                      f"  The full content will be used when you click 'CONVERT & SAVE'.\n"
                    
                    if not self.is_preview_blocked:
                        preview_msg += "\nYou can now read, edit, or delete any character in the left editor. Click 'Browse' in Output to change the save location."
                    else:
                        preview_msg += "\nPreview is skipped. Click 'Browse' in Output to change the save location and click 'CONVERT & SAVE' to convert the file directly."
                    
                    self._write_preview(preview_msg)
                    if self.is_preview_blocked:
                        self._set_status("Loaded (preview skipped for performance)", "orange")
                        self.drop_lbl.configure(text=f"Loaded (No preview): {os.path.basename(path)}", text_color="#2ecc71")
                    else:
                        self._set_status("Loaded and extracted successfully!" if not is_truncated else "Loaded (preview truncated for performance)", "green")
                        self.drop_lbl.configure(text=f"Loaded successfully: {os.path.basename(path)}", text_color="#2ecc71")
                    
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self._toggle_ui_state(True)
 
                self.after(0, update_ui)
            except Exception as e:
                err_msg = str(e)
                def update_error():
                    self.in_path.set("")
                    self.full_content = ""
                    self.editor.configure(state="normal")
                    self.editor.delete("1.0", "end")  # Clear editor content on load failure
                    self.drop_lbl.configure(text=f"Failed: {err_msg}", text_color="#e74c3c")
                    self._set_status(f"Load error: {err_msg}", "red")
                    self._write_preview(f"Load error details:\n{err_msg}")
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self._toggle_ui_state(True)
                self.after(0, update_error)

        threading.Thread(target=task, daemon=True).start()

    def _browse_output(self):
        if self.is_processing:
            return
        ext = self._cfg()["out_ext"]
        path = filedialog.asksaveasfilename(parent=self,
            defaultextension=ext, filetypes=OUT_FILETYPES[ext],
            initialfile=os.path.splitext(os.path.basename(self.in_path.get() or "document"))[0] + ext
        )
        if path:
            self.out_path.set(path)

    # ── Execute Document Generation & Export ─────────────────────────────────

    def _run_conversion(self):
        if self.is_processing:
            return
        # Ưu tiên dùng full_content nếu file lớn (vượt quá giới hạn hiển thị của editor) hoặc bị chặn preview được tải
        is_large_or_blocked = (self.full_content and len(self.full_content) > EDITOR_DISPLAY_LIMIT) or self.is_preview_blocked
        content = self.full_content if is_large_or_blocked else self.editor.get("1.0", "end-1c").strip()
        content = content.strip()
        out = self.out_path.get().strip()
        mode = self.mode_var.get()

        if not content:
            self._set_status("Editor is empty, nothing to convert!", "orange")
            return
        if not out:
            self._set_status("Please choose a save path!", "orange")
            return

        # Cảnh báo nếu converting file lớn (preview chỉ 500KB hoặc bị blocked nhưng convert toàn bộ)
        if self.full_content and (len(self.full_content) > EDITOR_DISPLAY_LIMIT or self.is_preview_blocked):
            from tkinter import messagebox
            full_size_mb = len(self.full_content) / (1024 * 1024)
            if self.is_preview_blocked:
                msg_text = f"You are converting a large file ({full_size_mb:.2f} MB).\n\n" \
                           f"Note: Previewing was blocked for this file, but the entire file will be converted.\n\n" \
                           f"Do you want to proceed?"
            else:
                msg_text = f"You are converting a large file ({full_size_mb:.2f} MB).\n\n" \
                           f"Note: Only the first 500KB was displayed for preview, but the entire file will be converted.\n\n" \
                           f"Any edits you made in the preview area will NOT be applied to the full conversion.\n\n" \
                           f"Do you want to proceed?"
            confirm = messagebox.askyesno(
                title="Large File Conversion",
                message=msg_text
            )
            if not confirm:
                self._set_status("Conversion cancelled", "orange")
                return

        # Validate Markdown tables to prevent malformed data parsing
        if mode == "MD -> Excel":
            if not has_md_tables(content):
                from tkinter import messagebox
                messagebox.showwarning(
                    parent=self,
                    title="No Tables Found",
                    message="No tables were found in the Markdown content.\n\n"
                            "To convert to Excel, your Markdown file must contain at least one table in standard Markdown format, for example:\n\n"
                            "| Column 1 | Column 2 |\n"
                            "| --- | --- |\n"
                            "| Value 1 | Value 2 |\n\n"
                            "Please ensure there is a separator row (like '|---|') below the header row."
                )
                self._set_status("No tables found in content", "red")
                return

        warnings = get_md_table_warnings(content)
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

            # Check if the output file is locked
            try:
                with open(out, "r+b") as f:
                    pass
            except PermissionError:
                self._set_status("Output file locked!", "red")
                messagebox.showerror(
                    parent=self,
                    title="File Lock Error",
                    message=f"The destination file '{os.path.basename(out)}' is currently open or locked by another application (e.g., Microsoft Word or Excel).\n\nPlease close the application holding the file and try again."
                )
                return
            except Exception:
                pass

        self._toggle_ui_state(False)
        self.btn_convert.configure(text="Converting...")
        self.btn_open_file.configure(state="disabled", fg_color="#1c1c24", text_color="#8a8a9e")
        self._set_status("Processing conversion and writing file...", "orange")

        # Show and start progress bar
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 5))
        self.progress_bar.start()

        def task():
            try:
                # Direct dispatcher
                msg = convert_content(mode, content, out)

                color = "green" if msg.startswith("Exported") or msg.startswith("Word") or msg.startswith("Markdown") else "red"
                
                # Output Overview details
                if color == "red":
                    log_details = f"DOCUMENT EXPORT LOG:\n" \
                                  f"-----------------------------------\n" \
                                  f"- Result: FAILED\n" \
                                  f"- Conversion mode: {mode}\n\n" \
                                  f"Error:\n{msg}"
                else:
                    log_details = f"DOCUMENT EXPORT LOG:\n" \
                                  f"-----------------------------------\n" \
                                  f"- Result: SUCCESS\n" \
                                  f"- Conversion mode: {mode}\n" \
                                  f"- Generated file: {os.path.basename(out)}\n" \
                                  f"- Save folder: {os.path.dirname(out)}\n" \
                                  f"- New size: {os.path.getsize(out) if os.path.exists(out) else 0} bytes\n\n" \
                                  f"Click the 'OPEN CREATED FILE' button below to open and view your document directly!"

                def update_success():
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self._write_preview(log_details)
                    self._set_status(msg.split("\n")[0] if "\n" in msg else msg, color)
                    self._toggle_ui_state(True)
                    if color == "green":
                        self.btn_open_file.configure(
                            state="normal", fg_color="#3a86ff", hover_color="#2563eb", text_color="#ffffff"
                        )
                        self.is_dirty = False

                self.after(0, update_success)
            except Exception as e:
                err_msg = str(e)
                def update_error():
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self._set_status(f"Conversion error: {err_msg}", "red")
                    self._write_preview(f"Conversion error details:\n{err_msg}")
                    self._toggle_ui_state(True)

                self.after(0, update_error)

        threading.Thread(target=task, daemon=True).start()

    def _open_generated_file(self):
        out = self.out_path.get().strip()
        if out and os.path.exists(out):
            try:
                if sys.platform == "win32":
                    os.startfile(out)
                elif sys.platform == "darwin":
                    subprocess.run(["open", out], check=True)
                else:
                    subprocess.run(["xdg-open", out], check=True)
                self._set_status("Opened file in default application", "green")
            except Exception as e:
                self._set_status(f"Failed to open file: {e}", "red")
        else:
            self._set_status("File does not exist or has not been created yet!", "orange")

    def _undo(self, event=None):
        if self.is_processing:
            return "break"
        try:
            self.editor.edit_undo()
        except Exception:
            pass
        return "break"

    def _redo(self, event=None):
        if self.is_processing:
            return "break"
        try:
            self.editor.edit_redo()
        except Exception:
            pass
        return "break"

    def _on_drag_enter(self, event):
        if self.is_processing:
            return
        self.load_bar.configure(border_color="#3a86ff", border_width=2)
        self.drop_lbl.configure(text="Drop the file now!", text_color="#3a86ff")

    def _on_drag_leave(self, event=None):
        if self.is_processing:
            return
        self.load_bar.configure(border_color="#2c2c35", border_width=1)
        inp = self.in_path.get().strip()
        if inp:
            self.drop_lbl.configure(text=f"Loaded successfully: {os.path.basename(inp)}", text_color="#2ecc71")
        else:
            self.drop_lbl.configure(text="Drag & drop file here or click 'Browse' to load content...", text_color="#7eb8f5")

    def _on_closing(self):
        if self.is_processing:
            from tkinter import messagebox
            confirm = messagebox.askyesno(
                title="Operation in Progress",
                message="An operation is currently in progress. Exiting may interrupt the process.\n\nAre you sure you want to exit?"
            )
            if not confirm:
                return
        elif self.is_dirty:
            from tkinter import messagebox
            confirm = messagebox.askyesno(
                title="Unsaved Changes",
                message="You have unsaved changes in the editor.\n\nAre you sure you want to exit?"
            )
            if not confirm:
                return
        self.destroy()
