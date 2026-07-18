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

# Force document modules to load and register
from src.core.registry import ModuleRegistry
import src.modules  # noqa: F401

from src.services.file_loader import load_document
from src.services.conversion_service import (
    convert_content,
    get_md_table_warnings,
    has_md_tables,
    is_output_locked,
)
from src.__version__ import __version__

# Load custom theme before any window is constructed
_theme_path = os.path.join(os.path.dirname(__file__), "theme.json")
if os.path.exists(_theme_path):
    try:
        ctk.set_default_color_theme(_theme_path)
    except Exception as e:
        print(f"[WARN] Failed to load custom theme: {e}")

# Premium Accent Palettes configuration. Format: (light_mode_color, dark_mode_color)
PALETTES = {
    "Violet Cyberpunk": {
        "text_accent_primary": ("#5d3fd3", "#725ac1"),
        "text_accent_secondary": ("#0096b4", "#00b4d8"),
        "btn_convert_fg": ("#5d3fd3", "#725ac1"),
        "btn_convert_hover": ("#4c329a", "#5d3fd3"),
        "btn_open_fg": ("#0096b4", "#00b4d8"),
        "btn_open_hover": ("#007a93", "#0096b4"),
        "bg_header": ("#ebe4ff", "#161224"),
        "bg_pane": ("#eaeaea", "#111216"),
        "bg_component": ("#ffffff", "#181a22"),
        "bg_pure_dark": ("#ffffff", "#07080a"),
        "border_color": ("#dcdcdc", "#222530")
    },
    "Emerald Obsidian": {
        "text_accent_primary": ("#059669", "#10b981"),
        "text_accent_secondary": ("#0f766e", "#14b8a6"),
        "btn_convert_fg": ("#059669", "#10b981"),
        "btn_convert_hover": ("#047857", "#059669"),
        "btn_open_fg": ("#0f766e", "#14b8a6"),
        "btn_open_hover": ("#0d5c55", "#0f766e"),
        "bg_header": ("#e6f4ea", "#121815"),
        "bg_pane": ("#eaeaea", "#111216"),
        "bg_component": ("#ffffff", "#161c19"),
        "bg_pure_dark": ("#ffffff", "#080c0a"),
        "border_color": ("#dcdcdc", "#202a25")
    },
    "Deep Ocean": {
        "text_accent_primary": ("#2563eb", "#3b82f6"),
        "text_accent_secondary": ("#0891b2", "#06b6d4"),
        "btn_convert_fg": ("#2563eb", "#3b82f6"),
        "btn_convert_hover": ("#1d4ed8", "#2563eb"),
        "btn_open_fg": ("#0891b2", "#06b6d4"),
        "btn_open_hover": ("#0e7490", "#0891b2"),
        "bg_header": ("#e8f0fe", "#0d131f"),
        "bg_pane": ("#eaeaea", "#111216"),
        "bg_component": ("#ffffff", "#121926"),
        "bg_pure_dark": ("#ffffff", "#070b12"),
        "border_color": ("#dcdcdc", "#1a2436")
    },
    "Sunset Gold": {
        "text_accent_primary": ("#d97706", "#f59e0b"),
        "text_accent_secondary": ("#ea580c", "#f97316"),
        "btn_convert_fg": ("#d97706", "#f59e0b"),
        "btn_convert_hover": ("#b45309", "#d97706"),
        "btn_open_fg": ("#ea580c", "#f97316"),
        "btn_open_hover": ("#c2410c", "#ea580c"),
        "bg_header": ("#fef3c7", "#171410"),
        "bg_pane": ("#eaeaea", "#111216"),
        "bg_component": ("#ffffff", "#1a1612"),
        "bg_pure_dark": ("#ffffff", "#0a0b0d"),
        "border_color": ("#dcdcdc", "#26201a")
    }
}

# UI Style configurations and color variables (Adapts to Light/Dark Mode)
STYLE = {
    # Typography
    "font_family_title": "Segoe UI",
    "font_family_body": "Segoe UI",
    "font_family_mono": "Consolas",
    
    # Text colors
    "text_primary": ("#1d1d1f", "#ffffff"),
    "text_muted": ("#555555", "#8f93a7"),
    "text_editor_fg": ("#1d1d1f", "#f8f8f2"),
    
    # Status colors (light_mode_color, dark_mode_color)
    "status_green": ("#0d9488", "#2ec4b6"),
    "status_red": ("#dc2626", "#e71d36"),
    "status_orange": ("#ea580c", "#ff9f1c"),
    "status_gray": ("#6b7280", "#6f738a"),
    
    # Button override colors
    "btn_clear_fg": ("#dc2626", "#c0392b"),
    "btn_clear_hover": ("#b91c1c", "#e74c3c"),
    
    "btn_utility_fg": ("#f3f4f6", "#1d202b"),
    "btn_utility_hover": ("#e5e7eb", "#2b2f42"),
    "btn_utility_border": ("#d1d5db", "#343952"),
    
    # Text search highlight tags (manually resolved in code)
    "tag_search_bg": ("#d8b4fe", "#3d2e6b"),
    "tag_search_fg": ("#000000", "#ffffff"),
    "tag_active_bg": ("#fef08a", "#b58400"),
    "tag_active_fg": ("#000000", "#ffffff")
}

# ── Configuration constants ───────────────────────────────────────────────────

# Resolve standard OS AppData directory (e.g. Roaming AppData on Windows)
appdata_dir = os.getenv('APPDATA')
if not appdata_dir:
    # Cross-platform fallback for macOS / Linux
    appdata_dir = os.path.join(os.path.expanduser("~"), ".config")

DRAFT_PATH = os.path.join(appdata_dir, "DocConvert", "draft_autosave.md")

# Giới hạn kích cỡ hiển thị trong textbox editor (500KB = ~500,000 ký tự)
EDITOR_DISPLAY_LIMIT = 500_000

# ── Conversion modes config ───────────────────────────────────────────────────

MODES = {
    "MD -> Excel":  {"in_ext": ".md",   "out_ext": ".xlsx", "in_label": "File .md",   "out_label": "Save .xlsx"},
    "MD -> Word":   {"in_ext": ".md",   "out_ext": ".docx", "in_label": "File .md",   "out_label": "Save .docx"},
    "MD -> CSV":    {"in_ext": ".md",   "out_ext": ".csv",  "in_label": "File .md",   "out_label": "Save .csv"},
    "MD -> PDF":    {"in_ext": ".md",   "out_ext": ".pdf",  "in_label": "File .md",   "out_label": "Save .pdf"},
    "MD -> HTML":   {"in_ext": ".md",   "out_ext": ".html", "in_label": "File .md",   "out_label": "Save .html"},
    "Excel -> MD":  {"in_ext": ".xlsx", "out_ext": ".md",   "in_label": "File .xlsx", "out_label": "Save .md"},
    "Word -> MD":   {"in_ext": ".docx", "out_ext": ".md",   "in_label": "File .docx", "out_label": "Save .md"},
    "CSV -> MD":    {"in_ext": ".csv",  "out_ext": ".md",   "in_label": "File .csv",  "out_label": "Save .md"},
    "PDF -> MD":    {"in_ext": ".pdf",  "out_ext": ".md",   "in_label": "File .pdf",  "out_label": "Save .md"},
    "HTML -> MD":   {"in_ext": ".html", "out_ext": ".md",   "in_label": "File .html", "out_label": "Save .md"},
}

IN_FILETYPES = {
    ".md":   [("Markdown", "*.md"), ("All Files", "*.*")],
    ".xlsx": [("Excel", "*.xlsx *.xls"), ("All Files", "*.*")],
    ".docx": [("Word", "*.docx"), ("All Files", "*.*")],
    ".csv":  [("CSV", "*.csv"), ("All Files", "*.*")],
    ".pdf":  [("PDF", "*.pdf"), ("All Files", "*.*")],
    ".html": [("HTML Document", "*.html *.htm"), ("All Files", "*.*")],
}

OUT_FILETYPES = {
    ".xlsx": [("Excel", "*.xlsx")],
    ".docx": [("Word",  "*.docx")],
    ".md":   [("Markdown", "*.md")],
    ".csv":  [("CSV", "*.csv")],
    ".pdf":  [("PDF", "*.pdf")],
    ".html": [("HTML Document", "*.html *.htm")],
}

BaseClass = TkinterDnD.Tk if HAS_DND else ctk.CTk


class OutputFileState:
    def __init__(self, in_path: str, out_path: str, duration_sec: float):
        self.in_path = in_path
        self.out_path = out_path
        self.in_filename = os.path.basename(in_path)
        self.out_filename = os.path.basename(out_path)
        self.out_dir = os.path.dirname(out_path)
        self.size_bytes = os.path.getsize(out_path) if os.path.exists(out_path) else 0
        self.duration_sec = duration_sec


class App(BaseClass): # type: ignore
    def __init__(self):
        super().__init__()
        self.title(f"Document Converter Workspace v{__version__}")
        
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
        self.output_state = None
        self._overlay_visible = False
        self._badge_timer_id = None
        self._toast_timer_id = None
        self._preview_timer = None
        
        self._block_update_dimensions = False
        
        # Theme variables
        self.appearance_mode_var = ctk.StringVar(value="Dark")
        self.current_palette_var = ctk.StringVar(value="Violet Cyberpunk")
        
        # Search & Replace panel variables
        self.matches = []
        self.current_match_idx = -1
        self.search_panel_visible = False
        
        # Register callback for appearance mode changes (to handle Windows titlebar dynamic sync)
        try:
            from customtkinter.windows.widgets.appearance_mode import AppearanceModeTracker
            AppearanceModeTracker.add(self._update_titlebar_theme)
        except Exception as e:
            print(f"[DEBUG] Failed to register appearance mode callback: {e}")
            
        self._build_ui()
        self._on_mode_change()
        
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def block_update_dimensions_event(self):
        self._block_update_dimensions = True

    def unblock_update_dimensions_event(self):
        self._block_update_dimensions = False

    def _build_ui(self):
        # Grid Configuration for main workspace
        self.rowconfigure(0, weight=0) # Header
        self.rowconfigure(1, weight=0) # Separator Line
        self.rowconfigure(2, weight=1) # Main workspace split
        self.columnconfigure(0, weight=1)

        # 1. Sleek Header Panel
        self.header_frame = ctk.CTkFrame(self, fg_color=STYLE["text_primary"], height=70, corner_radius=0, border_width=0) # Will be configured on palette change
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.header_frame.grid_propagate(False)
        
        title_lbl = ctk.CTkLabel(
            self.header_frame, text=f"Document Converter Workspace v{__version__}",
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=20, weight="bold"), text_color=STYLE["text_primary"]
        )
        title_lbl.pack(side="left", padx=25, pady=10)
        
        subtitle_lbl = ctk.CTkLabel(
            self.header_frame, text="Multipurpose document editing and conversion workspace",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=13, slant="italic"), text_color=STYLE["text_muted"]
        )
        subtitle_lbl.pack(side="left", padx=10, pady=16)

        # Theme controls frame on the right side of the header
        theme_ctrl_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent", border_width=0)
        theme_ctrl_frame.pack(side="right", padx=25, pady=10)
        
        # Color Palette Dropdown
        palette_lbl = ctk.CTkLabel(
            theme_ctrl_frame, text="Theme:", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"),
            text_color=STYLE["text_muted"]
        )
        palette_lbl.pack(side="left", padx=(0, 5))
        
        self.palette_menu = ctk.CTkOptionMenu(
            theme_ctrl_frame, 
            values=list(PALETTES.keys()), 
            variable=self.current_palette_var,
            width=140, height=28,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12),
            fg_color=STYLE["btn_utility_fg"],
            button_color=STYLE["btn_utility_fg"],
            button_hover_color=STYLE["btn_utility_hover"],
            text_color=STYLE["text_primary"],
            command=self._change_palette
        )
        self.palette_menu.pack(side="left", padx=(0, 15))
        
        # Appearance Mode Dropdown
        mode_lbl = ctk.CTkLabel(
            theme_ctrl_frame, text="Mode:", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"),
            text_color=STYLE["text_muted"]
        )
        mode_lbl.pack(side="left", padx=(0, 5))
        
        self.appearance_menu = ctk.CTkOptionMenu(
            theme_ctrl_frame, 
            values=["Dark", "Light", "System"], 
            variable=self.appearance_mode_var,
            width=90, height=28,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12),
            fg_color=STYLE["btn_utility_fg"],
            button_color=STYLE["btn_utility_fg"],
            button_hover_color=STYLE["btn_utility_hover"],
            text_color=STYLE["text_primary"],
            command=self._change_appearance_mode
        )
        self.appearance_menu.pack(side="left")

        # Thin border separator line below the header
        self.separator = ctk.CTkFrame(self, fg_color=STYLE["btn_utility_border"], height=1, corner_radius=0, border_width=0)
        self.separator.grid(row=1, column=0, sticky="ew", padx=0, pady=0)

        # 2. Main Workspace Wrapper to support rounded borders
        self.workspace_outer = ctk.CTkFrame(self, fg_color=PALETTES["Violet Cyberpunk"]["bg_pane"], corner_radius=0, border_width=0)
        self.workspace_outer.grid(row=2, column=0, sticky="nsew", padx=0, pady=0)
        
        self.workspace = ctk.CTkFrame(self.workspace_outer, fg_color=PALETTES["Violet Cyberpunk"]["bg_pane"], corner_radius=0, border_width=0)
        self.workspace.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.workspace.columnconfigure(0, weight=6, uniform="workspace_split") # Left Pane gets more weight
        self.workspace.columnconfigure(1, weight=5, uniform="workspace_split") # Right Pane
        self.workspace.rowconfigure(0, weight=1)

        # ── LEFT PANE: Input Editor & Overview ────────────────────────────────
        self.left_pane = ctk.CTkFrame(
            self.workspace, 
            fg_color=STYLE["text_primary"], 
            corner_radius=12, 
            border_width=1, 
            border_color=STYLE["btn_utility_border"]
        )
        self.left_pane.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)

        # Drop overlay frame (visual-only, sits on top of editor & toolbar when file dragged over left pane)
        self.drop_overlay = ctk.CTkFrame(
            self.left_pane,
            fg_color=("#ebe4ff", "#130f24"),
            border_width=2,
            corner_radius=12
        )
        # Sits in the center of the overlay
        self.drop_overlay_inner = ctk.CTkFrame(self.drop_overlay, fg_color="transparent")
        self.drop_overlay_inner.place(relx=0.5, rely=0.5, anchor="center")
        
        self.drop_icon_lbl = ctk.CTkLabel(
            self.drop_overlay_inner, text="📂",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=36)
        )
        self.drop_icon_lbl.pack(pady=10)
        
        self.drop_text_lbl = ctk.CTkLabel(
            self.drop_overlay_inner, text="Drop your document here",
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=16, weight="bold"),
            text_color=STYLE["text_primary"]
        )
        self.drop_text_lbl.pack(pady=5)
        
        self.drop_formats_lbl = ctk.CTkLabel(
            self.drop_overlay_inner, text="Supports: DOCX | PDF | XLS | CSV | MD | HTML",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12),
            text_color=STYLE["text_muted"]
        )
        self.drop_formats_lbl.pack(pady=5)
        self.left_pane.rowconfigure(0, weight=0) # Label
        self.left_pane.rowconfigure(1, weight=0) # Drag Drop Info
        self.left_pane.rowconfigure(2, weight=0) # Toolbar
        self.left_pane.rowconfigure(3, weight=0) # Search & Replace Panel (collapsible)
        self.left_pane.rowconfigure(4, weight=1) # Textbox Editor
        self.left_pane.rowconfigure(5, weight=0) # Footer stats
        self.left_pane.columnconfigure(0, weight=1)

        self.editor_title = ctk.CTkLabel(
            self.left_pane, text="INPUT EDITOR & OVERVIEW (MARKDOWN / TEXT)",
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=12, weight="bold"), 
            text_color=STYLE["text_primary"]
        )
        self.editor_title.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 5))

        # File Load Area / Drag Zone (with subtle border for drag enter/leave highlights)
        self.load_bar = ctk.CTkFrame(
            self.left_pane, 
            fg_color=STYLE["btn_utility_fg"], 
            corner_radius=8, 
            height=45, 
            border_width=1, 
            border_color=STYLE["btn_utility_border"]
        )
        self.load_bar.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        self.load_bar.pack_propagate(False)
        
        self.btn_browse_in = ctk.CTkButton(
            self.load_bar, text="Browse", width=75, height=28,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"), 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            command=self._browse_input
        )
        self.btn_browse_in.place(relx=1.0, rely=0.5, anchor="e", x=-8)

        self.drop_lbl = ctk.CTkLabel(
            self.load_bar, text="Drag & drop file here or click 'Browse' to load content...",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12), 
            text_color=STYLE["text_muted"]
        )
        self.drop_lbl.place(relx=0.0, rely=0.5, anchor="w", x=15)



        # The Formatting Toolbar (Word / Excel style toolbar)
        self.toolbar_frame = ctk.CTkFrame(self.left_pane, fg_color="transparent", border_width=0)
        self.toolbar_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(5, 5))
        
        btn_configs = [
            ("B", lambda: self._apply_format("**", "**"), "Bold"),
            ("I", lambda: self._apply_format("*", "*"), "Italic"),
            ("S", lambda: self._apply_format("~~", "~~"), "Strikethrough"),
            ("U", lambda: self._apply_format("<u>", "</u>"), "Underline"),
            ("Code", lambda: self._apply_format("`", "`"), "Inline Code"),
            ("Link 🔗", self._apply_link, "Link"),
            ("H1", lambda: self._apply_heading("#"), "H1"),
            ("H2", lambda: self._apply_heading("##"), "H2"),
            ("H3", lambda: self._apply_heading("###"), "H3"),
            ("List ☰", lambda: self._apply_format("- ", ""), "Bullet List"),
            ("List 🔢", lambda: self._apply_format("1. ", ""), "Numbered List"),
            ("Table ⊞", self._insert_table, "Table"),
        ]
        
        for idx, (text, command, tooltip) in enumerate(btn_configs):
            btn = ctk.CTkButton(
                self.toolbar_frame, text=text, width=45 if len(text) > 2 else 32, height=24,
                font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
                fg_color=STYLE["btn_utility_fg"],
                hover_color=STYLE["btn_utility_hover"],
                border_color=STYLE["btn_utility_border"],
                border_width=1,
                text_color=STYLE["text_primary"],
                command=command
            )
            btn.pack(side="left", padx=2)

        # The Live Monospace Text Editor (with Undo/Redo tracking)
        self.editor = ctk.CTkTextbox(
            self.left_pane, 
            fg_color=STYLE["btn_utility_fg"], 
            text_color=STYLE["text_editor_fg"],
            font=ctk.CTkFont(family=STYLE["font_family_mono"], size=13),
            border_width=1, 
            border_color=STYLE["btn_utility_border"], 
            corner_radius=8,
            undo=True
        )
        self.editor.grid(row=4, column=0, sticky="nsew", padx=15, pady=8)
        self.editor._textbox.bind("<KeyRelease>", self._on_editor_changed)
        self.editor._textbox.bind("<KeyPress>", self._on_editor_key_press)

        # Register highlight tags for search
        self._update_highlight_colors()

        # Bind custom Undo/Redo events to catch exceptions and prevent duplicate actions
        self.editor._textbox.bind("<Control-z>", self._undo)
        self.editor._textbox.bind("<Control-Z>", self._undo)
        self.editor._textbox.bind("<Control-y>", self._redo)
        self.editor._textbox.bind("<Control-Y>", self._redo)
        self.editor._textbox.bind("<Control-Shift-z>", self._redo)
        self.editor._textbox.bind("<Control-Shift-Z>", self._redo)

        # Editor Footer (Stats & Actions)
        editor_footer = ctk.CTkFrame(self.left_pane, fg_color="transparent", border_width=0)
        editor_footer.grid(row=5, column=0, sticky="ew", padx=15, pady=(5, 12))
        
        self.char_lbl = ctk.CTkLabel(
            editor_footer, text="Characters: 0", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12), 
            text_color=STYLE["text_muted"]
        )
        self.char_lbl.pack(side="left", padx=5)

        self.word_lbl = ctk.CTkLabel(
            editor_footer, text=" |  Words: 0", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12), 
            text_color=STYLE["text_muted"]
        )
        self.word_lbl.pack(side="left", padx=5)

        self.btn_md_guide = ctk.CTkButton(
            editor_footer, text="MD Guide ❔", width=85, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._show_md_guide
        )
        self.btn_md_guide.pack(side="left", padx=15)

        self.btn_clear = ctk.CTkButton(
            editor_footer, text="Clear All", width=70, height=24, 
            fg_color=STYLE["btn_clear_fg"], 
            hover_color=STYLE["btn_clear_hover"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._clear_editor
        )
        self.btn_clear.pack(side="right", padx=5)

        self.btn_find_replace = ctk.CTkButton(
            editor_footer, text="Find & Replace", width=95, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._toggle_search_panel
        )
        self.btn_find_replace.pack(side="right", padx=5)

        self.btn_redo = ctk.CTkButton(
            editor_footer, text="Redo", width=55, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._redo
        )
        self.btn_redo.pack(side="right", padx=5)

        self.btn_undo = ctk.CTkButton(
            editor_footer, text="Undo", width=55, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._undo
        )
        self.btn_undo.pack(side="right", padx=5)


        # ── RIGHT PANE: Output Configuration, Preview & Actions ───────────────
        self.right_pane = ctk.CTkFrame(
            self.workspace, 
            fg_color=STYLE["text_primary"], 
            corner_radius=12, 
            border_width=1, 
            border_color=STYLE["btn_utility_border"]
        )
        self.right_pane.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        self.right_pane.rowconfigure(0, weight=0) # Title
        self.right_pane.rowconfigure(1, weight=0) # Config panel
        self.right_pane.rowconfigure(2, weight=0) # Tab Selector
        self.right_pane.rowconfigure(3, weight=1) # Output preview / Logs or Document Preview Frame
        self.right_pane.rowconfigure(4, weight=0) # Convert Button & Status
        self.right_pane.columnconfigure(0, weight=1)

        self.output_title = ctk.CTkLabel(
            self.right_pane, text="OUTPUT CONFIGURATION & EXPORT",
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=12, weight="bold"), 
            text_color=STYLE["text_primary"]
        )
        self.output_title.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 5))

        # Config Panel
        self.config_frame = ctk.CTkFrame(
            self.right_pane, 
            fg_color=STYLE["btn_utility_fg"], 
            corner_radius=8,
            border_width=1,
            border_color=STYLE["btn_utility_border"]
        )
        self.config_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        self.config_frame.columnconfigure(0, weight=1)

        # Mode row
        row_mode = ctk.CTkFrame(self.config_frame, fg_color="transparent", border_width=0)
        row_mode.pack(fill="x", padx=12, pady=(10, 5))
        ctk.CTkLabel(
            row_mode, text="Mode:", width=90, anchor="w", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"),
            text_color=STYLE["text_muted"]
        ).pack(side="left")
        
        self.mode_menu = ctk.CTkOptionMenu(
            row_mode, values=list(MODES.keys()), variable=self.mode_var,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12),
            fg_color=STYLE["text_primary"],
            button_color=STYLE["text_primary"],
            button_hover_color=STYLE["text_primary"],
            command=self._on_mode_change
        )
        self.mode_menu.pack(side="left", fill="x", expand=True, padx=5)

        self.badge_lbl = ctk.CTkLabel(
            row_mode, text="",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            text_color=STYLE["status_green"],
            fg_color="transparent"
        )

        # Path row
        row_out = ctk.CTkFrame(self.config_frame, fg_color="transparent", border_width=0)
        row_out.pack(fill="x", padx=12, pady=(5, 10))
        self.lbl_out = ctk.CTkLabel(
            row_out, text="Output:", width=90, anchor="w", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"),
            text_color=STYLE["text_muted"]
        )
        self.lbl_out.pack(side="left")
        
        self.btn_browse_out = ctk.CTkButton(
            row_out, text="Browse", width=65, height=28,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"), 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            command=self._browse_output
        )
        self.btn_browse_out.pack(side="right", padx=2)

        self.btn_copy_path = ctk.CTkButton(
            row_out, text="Copy Path", width=75, height=28,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"), 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["status_gray"],
            state="disabled",
            command=self._copy_output_path
        )
        self.btn_copy_path.pack(side="right", padx=2)

        self.entry_out = ctk.CTkEntry(
            row_out, textvariable=self.out_path, placeholder_text="Select save location...",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12)
        )
        self.entry_out.pack(side="left", fill="x", expand=True, padx=5)

        # Segmented Tab Controller
        self.tab_frame = ctk.CTkFrame(self.right_pane, fg_color="transparent", height=35)
        self.tab_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(5, 0))
        self.tab_frame.grid_propagate(False)

        self.tab_var = ctk.StringVar(value="Document Preview")
        self.tab_selector = ctk.CTkSegmentedButton(
            self.tab_frame,
            values=["Document Preview", "System Log"],
            variable=self.tab_var,
            command=self._on_tab_change,
            height=28
        )
        self.tab_selector.pack(side="left", fill="x", expand=True)

        # Document Preview Frame (Markdown visual rendering)
        from src.ui.document_preview import DocumentPreviewFrame
        self.preview_frame = DocumentPreviewFrame(
            self.right_pane, 
            fg_color=STYLE["btn_utility_fg"],
            scrollbar_fg_color=STYLE["btn_utility_fg"],
            border_width=1, 
            border_color=STYLE["btn_utility_border"], 
            corner_radius=8
        )
        self.preview_frame.grid(row=3, column=0, sticky="nsew", padx=15, pady=8)

        # Output Logs & Text Workspace
        self.preview_box = ctk.CTkTextbox(
            self.right_pane, 
            fg_color=STYLE["btn_utility_fg"], 
            text_color=STYLE["text_editor_fg"],
            font=ctk.CTkFont(family=STYLE["font_family_mono"], size=12),
            border_width=1, 
            border_color=STYLE["btn_utility_border"], 
            corner_radius=8
        )
        # Note: self.preview_box is NOT gridded by default, managed by _on_tab_change
        self._write_preview(
            "SYSTEM READY\n\n"
            "- Drag & drop your Markdown, Excel, Word or PDF file into the left pane.\n"
            "- The smart extractor will automatically parse your file into Markdown for you to preview, edit, or delete any characters before exporting to a new format.\n\n"
            "💡 Supported Markdown formats: Headings (#), Bold (**), Italic (*), Strikethrough (~~), Underline (<u>), Inline Code (`), Links ([text](url)), Nested Lists, and Tables.\n\n"
            "Click 'MD Guide ❔' under the input editor for exact syntax examples!"
        )

        # Large Action Button & Status Info
        action_frame = ctk.CTkFrame(self.right_pane, fg_color="transparent", border_width=0)
        action_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=(5, 12))
        action_frame.columnconfigure(0, weight=3)
        action_frame.columnconfigure(1, weight=2)
        action_frame.columnconfigure(2, weight=2)

        self.btn_convert = ctk.CTkButton(
            action_frame, text="CONVERT & SAVE", height=48,
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=13, weight="bold"),
            fg_color=STYLE["text_primary"], 
            hover_color=STYLE["text_primary"], 
            text_color=STYLE["text_primary"],
            command=self._run_conversion
        )
        self.btn_convert.grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=(0, 5))

        self.btn_open_file = ctk.CTkButton(
            action_frame, text="OPEN FILE", height=48,
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=13, weight="bold"),
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"], 
            text_color=STYLE["status_gray"],
            state="disabled",
            command=self._open_generated_file
        )
        self.btn_open_file.grid(row=0, column=1, sticky="ew", padx=(4, 4), pady=(0, 5))

        self.btn_open_folder = ctk.CTkButton(
            action_frame, text="SHOW IN FOLDER", height=48,
            font=ctk.CTkFont(family=STYLE["font_family_title"], size=13, weight="bold"),
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"], 
            text_color=STYLE["status_gray"],
            state="disabled",
            command=self._open_containing_folder
        )
        self.btn_open_folder.grid(row=0, column=2, sticky="ew", padx=(4, 0), pady=(0, 5))

        self.progress_bar = ctk.CTkProgressBar(action_frame, mode="indeterminate", width=10, height=8, progress_color=STYLE["status_green"])
        # Kept hidden initially until conversion begins

        self.status_lbl = ctk.CTkLabel(
            action_frame, text="Ready to process", 
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, slant="italic"), 
            text_color=STYLE["status_gray"]
        )
        self.status_lbl.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(5, 0))

        # Build search and replace panel
        self._build_search_panel(self.left_pane)

        # Apply the default theme colors
        ctk.set_appearance_mode(self.appearance_mode_var.get())
        self._change_palette("Violet Cyberpunk")
        self._update_titlebar_theme()
        
        # Bind keyboard shortcuts on the editor widget
        self.editor._textbox.bind("<Control-f>", self._shortcut_find)
        self.editor._textbox.bind("<Control-F>", self._shortcut_find)
        self.editor._textbox.bind("<Control-h>", self._shortcut_replace)
        self.editor._textbox.bind("<Control-H>", self._shortcut_replace)
        
        # Bind keyboard shortcuts globally on the main app window
        self.bind("<Control-f>", self._shortcut_find)
        self.bind("<Control-F>", self._shortcut_find)
        self.bind("<Control-h>", self._shortcut_replace)
        self.bind("<Control-H>", self._shortcut_replace)
        
        # Register Drag & Drop with active drag hover animations
        if HAS_DND:
            for w in (self.load_bar, self.drop_lbl, self.left_pane, self.editor, self.drop_overlay):
                w.drop_target_register(DND_FILES) # type: ignore
                w.dnd_bind("<<Drop>>", self._on_drop) # type: ignore
                w.dnd_bind("<<DragEnter>>", self._on_drag_enter) # type: ignore
                w.dnd_bind("<<DragLeave>>", self._on_drag_leave) # type: ignore

        self._init_autosave()

    # ── Internal Actions & Event Handlers ─────────────────────────────────────

    def _cfg(self):
        return MODES[self.mode_var.get()]

    def _get_missing_dependencies(self) -> list[str]:
        mode = self.mode_var.get()
        if "->" in mode:
            src_fmt, dest_fmt = [part.strip() for part in mode.split("->")]
            target_fmt = dest_fmt if dest_fmt != "MD" else src_fmt
            module = ModuleRegistry.get_module_by_name(target_fmt)
            if module:
                return module.check_dependencies()
        return []

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
        missing = self._get_missing_dependencies()
        
        if missing:
            self.btn_convert.configure(state="disabled", fg_color=STYLE["status_red"], text_color_disabled=("#ffffff", "#ffffff"), text="UNAVAILABLE")
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
            palette = PALETTES[self.current_palette_var.get()]
            self.btn_convert.configure(state="normal", fg_color=palette["btn_convert_fg"], text_color=("#ffffff", "#ffffff"), text="CONVERT & SAVE")
            self._set_status("Mode changed to: " + mode, "primary")
            
            self._update_markdown_preview()
            
            if not self.editor.get("1.0", "end-1c").strip():
                self._write_preview("SYSTEM READY\n\n- Drag & drop your Markdown, Excel, Word or PDF file into the left pane.\n- The smart extractor will automatically parse your file into Markdown for you to preview, edit, or delete any characters before exporting to a new format.\n\n💡 Supported Markdown formats: Headings (#), Bold (**), Italic (*), Strikethrough (~~), Underline (<u>), Inline Code (`), Links ([text](url)), Nested Lists, and Tables.\n\nClick 'MD Guide ❔' under the input editor for exact syntax examples!")

    def _clear_editor(self):
        if self.is_processing:
            return
        self.editor.configure(state="normal", text_color=STYLE["text_editor_fg"]) # Mở khóa editor và đặt lại màu chữ
        self.editor.delete("1.0", "end")
        self.editor.edit_reset()  # Reset undo stack after manual clearing
        self.in_path.set("")
        self.full_content = ""  # Xóa toàn bộ content
        self.is_preview_blocked = False
        self.is_dirty = False
        self._update_counts()
        if os.path.exists(DRAFT_PATH):
            try:
                os.remove(DRAFT_PATH)
            except Exception:
                pass
        self._write_preview("Editor is empty. You can write your own Markdown text here!")
        self._update_markdown_preview()
        self._set_status("Editor cleared", "primary")
        palette = PALETTES[self.current_palette_var.get()]
        self.drop_lbl.configure(text="Drag & drop file here or click 'Browse' to load content...", text_color=palette["text_accent_secondary"])

    def _on_editor_changed(self, event=None):
        self.is_dirty = True
        if self.search_panel_visible:
            self._perform_search(keep_current_index=True)
        self._update_counts()
        self._schedule_preview_update()

    def _schedule_preview_update(self):
        if self._preview_timer is not None:
            self.after_cancel(self._preview_timer)
        self._preview_timer = self.after(300, self._update_markdown_preview)

    def _update_markdown_preview(self):
        self._preview_timer = None
        if self.tab_var.get() == "Document Preview":
            is_large_or_blocked = (self.full_content and len(self.full_content) > EDITOR_DISPLAY_LIMIT) or self.is_preview_blocked
            content = self.full_content if is_large_or_blocked else self.editor.get("1.0", "end-1c")
            self.preview_frame.update_preview(content)

    def _on_tab_change(self, selected_tab: str):
        if selected_tab == "Document Preview":
            self.preview_box.grid_forget()
            self.preview_frame.grid(row=3, column=0, sticky="nsew", padx=15, pady=8)
            self._update_markdown_preview()
        else:
            self.preview_frame.grid_forget()
            self.preview_box.grid(row=3, column=0, sticky="nsew", padx=15, pady=8)

    def _update_counts(self):
        is_large_or_blocked = (self.full_content and len(self.full_content) > EDITOR_DISPLAY_LIMIT) or self.is_preview_blocked
        content = self.full_content if is_large_or_blocked else self.editor.get("1.0", "end-1c")
        
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

    def _set_status(self, msg: str, color: str = "primary"):
        hex_colors = {
            "primary": STYLE["text_primary"],
            "green": STYLE["status_green"], 
            "red": STYLE["status_red"], 
            "gray": STYLE["status_gray"], 
            "orange": STYLE["status_orange"]
        }
        target_color = hex_colors.get(color, color)
        if target_color == "white":
            target_color = STYLE["text_primary"]
        self.status_lbl.configure(text=msg, text_color=target_color)

    def _toggle_ui_state(self, enabled: bool):
        self.is_processing = not enabled
        state = "normal" if enabled else "disabled"
        
        if enabled:
            missing = self._get_missing_dependencies()
            if missing:
                self.btn_convert.configure(state="disabled", fg_color=STYLE["status_red"], text_color_disabled=("#ffffff", "#ffffff"), text="UNAVAILABLE")
            else:
                palette = PALETTES[self.current_palette_var.get()]
                self.btn_convert.configure(state="normal", fg_color=palette["btn_convert_fg"], text_color=("#ffffff", "#ffffff"), text="CONVERT & SAVE")
        else:
            self.btn_convert.configure(state="disabled")

        self.btn_browse_in.configure(state=state)
        self.btn_clear.configure(state=state)
        self.btn_undo.configure(state=state)
        self.btn_redo.configure(state=state)
        self.mode_menu.configure(state=state)
        self.entry_out.configure(state=state)
        self.btn_browse_out.configure(state=state)
        
        # Configure output action buttons
        if not enabled:
            self.btn_open_file.configure(state="disabled", fg_color=STYLE["btn_utility_fg"], text_color=STYLE["status_gray"])
            self.btn_open_folder.configure(state="disabled", fg_color=STYLE["btn_utility_fg"], text_color=STYLE["status_gray"])
            self.btn_copy_path.configure(state="disabled", fg_color=STYLE["btn_utility_fg"], text_color=STYLE["status_gray"])
        else:
            if self.output_state is not None:
                palette_active = PALETTES[self.current_palette_var.get()]
                self.btn_open_file.configure(
                    state="normal", fg_color=palette_active["btn_open_fg"], hover_color=palette_active["btn_open_hover"], text_color=("#ffffff", "#ffffff")
                )
                self.btn_open_folder.configure(
                    state="normal", fg_color=palette_active["btn_open_fg"], hover_color=palette_active["btn_open_hover"], text_color=("#ffffff", "#ffffff")
                )
                self.btn_copy_path.configure(
                    state="normal", fg_color=palette_active["btn_open_fg"], hover_color=palette_active["btn_open_hover"], text_color=("#ffffff", "#ffffff")
                )
            else:
                self.btn_open_file.configure(state="disabled", fg_color=STYLE["btn_utility_fg"], text_color=STYLE["status_gray"])
                self.btn_open_folder.configure(state="disabled", fg_color=STYLE["btn_utility_fg"], text_color=STYLE["status_gray"])
                self.btn_copy_path.configure(state="disabled", fg_color=STYLE["btn_utility_fg"], text_color=STYLE["status_gray"])

        if hasattr(self, "search_entry"):
            self.search_entry.configure(state=state)
            self.replace_entry.configure(state=state)
            self.btn_prev.configure(state=state)
            self.btn_next.configure(state=state)
            self.btn_close_search.configure(state=state)
            self.btn_replace.configure(state=state)
            self.btn_replace_all.configure(state=state)
            self.btn_find_replace.configure(state=state)
            
        if not enabled:
            self.editor.configure(state="disabled")
        else:
            is_large_file = self.full_content and len(self.full_content) > EDITOR_DISPLAY_LIMIT
            if self.is_preview_blocked or is_large_file:
                self.editor.configure(state="disabled", text_color=STYLE["status_gray"])
            else:
                self.editor.configure(state="normal", text_color=STYLE["text_editor_fg"])

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
            ("Supported Documents", "*.md *.xlsx *.xls *.docx *.csv *.pdf *.html *.htm"),
            ("Markdown (*.md)", "*.md"),
            ("Excel (*.xlsx, *.xls)", "*.xlsx *.xls"),
            ("Word (*.docx)", "*.docx"),
            ("CSV (*.csv)", "*.csv"),
            ("PDF (*.pdf)", "*.pdf"),
            ("HTML Document (*.html, *.htm)", "*.html *.htm"),
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
                self.drop_lbl.configure(text=f"Failed: {missing_str} missing", text_color=STYLE["status_red"])
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
                self.drop_lbl.configure(text=f"Failed: {short_err}", text_color=STYLE["status_red"])
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
        self.drop_lbl.configure(text=f"Loading: {os.path.basename(path)}...", text_color=STYLE["status_orange"])

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
                        self._show_auto_detect_badge(detected_mode, path)
                    
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
                    
                    self.is_dirty = True
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
                        self.drop_lbl.configure(text=f"Loaded (No preview): {os.path.basename(path)}", text_color=STYLE["status_green"])
                    else:
                        self._set_status("Loaded and extracted successfully!" if not is_truncated else "Loaded (preview truncated for performance)", "green")
                        self.drop_lbl.configure(text=f"Loaded successfully: {os.path.basename(path)}", text_color=STYLE["status_green"])
                    
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self._update_markdown_preview()
                    self._toggle_ui_state(True)
 
                self.after(0, update_ui)
            except Exception as e:
                err_msg = str(e)
                def update_error():
                    self.in_path.set("")
                    self.full_content = ""
                    self.editor.configure(state="normal")
                    self.editor.delete("1.0", "end")  # Clear editor content on load failure
                    self.drop_lbl.configure(text=f"Failed: {err_msg}", text_color=STYLE["status_red"])
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

        import time
        start_time = time.time()

        palette = PALETTES[self.current_palette_var.get()]
        self._toggle_ui_state(False)
        self.btn_convert.configure(text="Converting...")
        self.btn_open_file.configure(state="disabled", fg_color=palette["bg_component"], text_color=STYLE["status_gray"])
        self.btn_open_folder.configure(state="disabled", fg_color=palette["bg_component"], text_color=STYLE["status_gray"])
        self._set_status("Processing conversion and writing file...", "orange")

        # Show and start progress bar
        self.progress_bar.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(5, 5))
        self.progress_bar.start()

        def task():
            try:
                # Direct dispatcher
                msg = convert_content(mode, content, out)
                duration_sec = time.time() - start_time

                color = "green" if msg.startswith("Exported") or msg.startswith("Word") or msg.startswith("Markdown") else "red"
                
                # Output Overview details
                if color == "red":
                    log_details = f"DOCUMENT EXPORT LOG:\n" \
                                  f"-----------------------------------\n" \
                                  f"- Result: FAILED\n" \
                                  f"- Conversion mode: {mode}\n\n" \
                                  f"Error:\n{msg}"
                else:
                    log_details = "" # Will be replaced by success card formatting

                def update_success():
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self._set_status(msg.split("\n")[0] if "\n" in msg else msg, color)
                    
                    if color == "green":
                        # Set Output State
                        in_p = self.in_path.get().strip() or "editor_content.md"
                        self.output_state = OutputFileState(in_p, out, duration_sec)
                        self._write_success_preview(mode)
                        self.is_dirty = False
                    else:
                        self.output_state = None
                        self._write_preview(log_details)
                    
                    self._toggle_ui_state(True)

                self.after(0, update_success)
            except Exception as e:
                err_msg = str(e)
                def update_error():
                    self.progress_bar.stop()
                    self.progress_bar.grid_remove()
                    self.output_state = None
                    self._set_status(f"Conversion error: {err_msg}", "red")
                    self._write_preview(f"Conversion error details:\n{err_msg}")
                    self._toggle_ui_state(True)

                self.after(0, update_error)

        threading.Thread(target=task, daemon=True).start()

    def _open_generated_file(self):
        if not self.output_state:
            self._set_status("No output file available to open!", "orange")
            return
        out = self.output_state.out_path
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
            self._set_status("File no longer exists!", "red")
            from tkinter import messagebox
            messagebox.showwarning(parent=self, title="File Not Found", message="The converted file no longer exists at the target location.")

    def _open_containing_folder(self):
        if not self.output_state:
            self._set_status("No output file available to show in folder!", "orange")
            return
        out = self.output_state.out_path
        if out and os.path.exists(out):
            try:
                if sys.platform == "win32":
                    subprocess.run(["explorer", "/select,", os.path.normpath(out)])
                elif sys.platform == "darwin":
                    subprocess.run(["open", "-R", out])
                else:
                    subprocess.run(["xdg-open", os.path.dirname(out)])
                self._set_status("Opened containing folder", "green")
            except Exception as e:
                self._set_status(f"Failed to open folder: {e}", "red")
        else:
            self._set_status("File no longer exists!", "red")
            from tkinter import messagebox
            messagebox.showwarning(parent=self, title="File Not Found", message="The converted file no longer exists at the target location.")

    def _copy_output_path(self):
        if not self.output_state:
            self._set_status("No output file available to copy path!", "orange")
            return
        out = self.output_state.out_path
        try:
            self.clipboard_clear()
            self.clipboard_append(out)
            self._set_status("✓ Path copied to clipboard", "green")
            
            if hasattr(self, "_toast_timer_id") and self._toast_timer_id is not None:
                try:
                    self.after_cancel(self._toast_timer_id)
                except Exception:
                    pass
            self.btn_copy_path.configure(text="✓ Copied!")
            self._toast_timer_id = self.after(1500, self._reset_copy_btn_text)
        except Exception as e:
            self._set_status(f"Failed to copy path: {e}", "red")

    def _reset_copy_btn_text(self):
        self.btn_copy_path.configure(text="Copy Path")
        self._toast_timer_id = None

    def _show_auto_detect_badge(self, mode, file_path):
        ext = os.path.splitext(file_path)[1].lower().replace(".", "").upper()
        # "MD -> Excel" -> "Excel"
        out_fmt = mode.split(" -> ")[1]
        
        if hasattr(self, "_badge_timer_id") and self._badge_timer_id is not None:
            try:
                self.after_cancel(self._badge_timer_id)
            except Exception:
                pass
            self._badge_timer_id = None

        self.badge_lbl.configure(text=f"✓ {ext} → {out_fmt} detected")
        self.badge_lbl.pack(side="left", padx=10)
        self._badge_timer_id = self.after(3000, self._hide_auto_detect_badge)

    def _hide_auto_detect_badge(self):
        self.badge_lbl.pack_forget()
        self._badge_timer_id = None

    def _write_success_preview(self, mode):
        # Configure tag configuration once
        try:
            self.preview_box._textbox.tag_config("success_title", foreground="#2ec4b6", font=ctk.CTkFont(family=STYLE["font_family_body"], size=14, weight="bold"))
            self.preview_box._textbox.tag_config("success_body", font=ctk.CTkFont(family=STYLE["font_family_body"], size=12))
        except Exception:
            pass

        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        
        self.preview_box.insert("end", "✓ CONVERSION COMPLETE\n\n", "success_title")
        
        state = self.output_state
        if not state:
            return
            
        size_bytes = state.size_bytes
        if size_bytes > 1024 * 1024:
            size_str = f"{size_bytes / (1024*1024):.2f} MB"
        else:
            size_str = f"{size_bytes / 1024:.2f} KB"
            
        details = (
            f"• Input:    {state.in_filename}\n"
            f"• Output:   {state.out_filename}\n"
            f"• Location: {state.out_dir}\n"
            f"• Size:     {size_str} ({size_bytes:,} bytes)\n"
            f"• Duration: {state.duration_sec:.2f} s\n\n"
            f"What would you like to do next?\n"
            f"- Click 'OPEN FILE' to view it.\n"
            f"- Click 'SHOW IN FOLDER' to highlight it.\n"
            f"- Click 'Copy Path' to copy path to clipboard."
        )
        self.preview_box.insert("end", details, "success_body")
        self.preview_box.configure(state="disabled")

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
        if not self._overlay_visible:
            self._overlay_visible = True
            palette = PALETTES[self.current_palette_var.get()]
            self.drop_overlay.configure(border_color=palette["text_accent_primary"])
            self.drop_overlay.place(relx=0.02, rely=0.1, relwidth=0.96, relheight=0.86)

    def _on_drag_leave(self, event=None):
        if self.is_processing:
            return
        try:
            # Check if pointer has actually left left_pane bounds
            x = self.winfo_pointerx() - self.left_pane.winfo_rootx()
            y = self.winfo_pointery() - self.left_pane.winfo_rooty()
            w = self.left_pane.winfo_width()
            h = self.left_pane.winfo_height()
            if x < 0 or x >= w or y < 0 or y >= h:
                self._overlay_visible = False
                self.drop_overlay.place_forget()
        except Exception:
            self._overlay_visible = False
            self.drop_overlay.place_forget()

    def _change_appearance_mode(self, mode: str):
        ctk.set_appearance_mode(mode)
        # Schedule the full background and title bar update to allow CustomTkinter's system theme resolver to settle
        self.after(100, self._execute_appearance_mode_update)

    def _update_titlebar_theme(self, _mode_arg=None):
        # Schedule the full background and title bar update to allow CustomTkinter's system theme resolver to settle
        self.after(100, self._execute_appearance_mode_update)

    def _execute_appearance_mode_update(self):
        # 1. Resolve current active mode
        appearance_mode = ctk.get_appearance_mode().lower()
        mode_idx = 0 if appearance_mode == "light" else 1
        
        # 2. Update window background
        palette = PALETTES[self.current_palette_var.get()]
        bg_color = palette["bg_pane"][mode_idx]
        try:
            self.configure(bg=bg_color)
        except Exception:
            pass
            
        # 3. Update workspace backgrounds
        self.workspace_outer.configure(fg_color=palette["bg_pane"])
        self.workspace.configure(fg_color=palette["bg_pane"])
        self.left_pane.configure(fg_color=palette["bg_pane"], border_color=palette["border_color"])
        self.right_pane.configure(fg_color=palette["bg_pane"], border_color=palette["border_color"])
        
        # 4. Update search highlights
        self._update_highlight_colors()
        
        # 5. Update Windows title bar theme
        if sys.platform != "win32":
            return
        try:
            import ctypes
            is_dark = (appearance_mode == "dark")
            hwnd = self.winfo_id()
            hwnd = ctypes.windll.user32.GetAncestor(hwnd, 2) # GA_ROOT = 2
            
            for attr in (20, 19):
                value = ctypes.c_int(1 if is_dark else 0)
                hr = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    attr,
                    ctypes.byref(value),
                    ctypes.sizeof(value)
                )
                if hr == 0:
                    break
            self.update_idletasks()
        except Exception as e:
            print(f"[DEBUG] Failed to set titlebar color: {e}")

    def _change_palette(self, palette_name: str):
        palette = PALETTES[palette_name]
        
        # 1. Update backgrounds and borders
        appearance_mode = ctk.get_appearance_mode().lower()
        mode_idx = 0 if appearance_mode == "light" else 1
        bg_color = palette["bg_pane"][mode_idx]
        try:
            self.configure(bg=bg_color)
        except Exception:
            pass
        self.header_frame.configure(fg_color=palette["bg_header"])
        self.separator.configure(fg_color=palette["border_color"])
        self.workspace_outer.configure(fg_color=palette["bg_pane"])
        self.workspace.configure(fg_color=palette["bg_pane"])
        self.left_pane.configure(fg_color=palette["bg_pane"], border_color=palette["border_color"])
        self.right_pane.configure(fg_color=palette["bg_pane"], border_color=palette["border_color"])
        self.load_bar.configure(fg_color=palette["bg_component"], border_color=palette["border_color"])
        self.config_frame.configure(fg_color=palette["bg_component"], border_color=palette["border_color"])
        
        # 2. Update titles
        self.editor_title.configure(text_color=palette["text_accent_primary"])
        self.output_title.configure(text_color=palette["text_accent_secondary"])
        
        # 3. Update editor text box & preview box
        self.editor.configure(fg_color=palette["bg_pure_dark"], border_color=palette["border_color"])
        self.preview_box.configure(fg_color=palette["bg_pure_dark"], border_color=palette["border_color"])
        
        # 4. Update entries
        self.entry_out.configure(fg_color=palette["bg_component"], border_color=palette["border_color"])
        if hasattr(self, "search_frame"):
            self.search_frame.configure(fg_color=palette["bg_component"], border_color=palette["border_color"])
            self.search_entry.configure(fg_color=palette["bg_component"], border_color=palette["border_color"])
            self.replace_entry.configure(fg_color=palette["bg_component"], border_color=palette["border_color"])
            
        # 5. Update primary and secondary action buttons
        missing = self._get_missing_dependencies()
        if missing:
            self.btn_convert.configure(fg_color=STYLE["status_red"], hover_color=STYLE["status_red"], text_color_disabled=("#ffffff", "#ffffff"))
        else:
            self.btn_convert.configure(fg_color=palette["btn_convert_fg"], hover_color=palette["btn_convert_hover"], text_color=("#ffffff", "#ffffff"))
            
        # Open file, Show folder, Copy path buttons: if normal state, use palette colors, else use component color
        for btn in (self.btn_open_file, self.btn_open_folder, self.btn_copy_path):
            if btn.cget("state") == "normal":
                btn.configure(
                    fg_color=palette["btn_open_fg"],
                    hover_color=palette["btn_open_hover"],
                    text_color=("#ffffff", "#ffffff")
                )
            else:
                btn.configure(
                    fg_color=palette["bg_component"],
                    text_color=STYLE["status_gray"]
                )
            
        # 6. Update OptionMenus button/hover colors
        self.mode_menu.configure(
            fg_color=palette["btn_convert_fg"],
            button_color=palette["btn_convert_fg"],
            button_hover_color=palette["btn_convert_hover"],
            text_color=("#ffffff", "#ffffff")
        )
        self.palette_menu.configure(
            fg_color=palette["btn_convert_fg"],
            button_color=palette["btn_convert_fg"],
            button_hover_color=palette["btn_convert_hover"],
            text_color=("#ffffff", "#ffffff")
        )
        self.appearance_menu.configure(
            fg_color=palette["btn_convert_fg"],
            button_color=palette["btn_convert_fg"],
            button_hover_color=palette["btn_convert_hover"],
            text_color=("#ffffff", "#ffffff")
        )
        
        # 7. Update status progress color
        self.progress_bar.configure(progress_color=palette["text_accent_secondary"])
        
        # 8. Update search highlight configurations manually
        self._update_highlight_colors()

        # 9. Update preview frame theme and tab selector styling
        if hasattr(self, "preview_frame"):
            self.preview_frame.set_theme(palette, STYLE)
        if hasattr(self, "tab_selector"):
            self.tab_selector.configure(
                selected_color=palette["text_accent_primary"],
                selected_hover_color=palette["btn_convert_hover"]
            )

    def _update_highlight_colors(self):
        appearance_mode = ctk.get_appearance_mode().lower() # "light" or "dark"
        mode_idx = 0 if appearance_mode == "light" else 1
        
        search_bg = STYLE["tag_search_bg"][mode_idx]
        search_fg = STYLE["tag_search_fg"][mode_idx]
        active_bg = STYLE["tag_active_bg"][mode_idx]
        active_fg = STYLE["tag_active_fg"][mode_idx]
        
        self.editor._textbox.tag_config("search_highlight", background=search_bg, foreground=search_fg)
        self.editor._textbox.tag_config("active_highlight", background=active_bg, foreground=active_fg)

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
            # Save draft immediately on exit so it is not lost
            try:
                content = self.editor.get("1.0", "end-1c")
                os.makedirs(os.path.dirname(DRAFT_PATH), exist_ok=True)
                with open(DRAFT_PATH, "w", encoding="utf-8") as f:
                    f.write(content)
            except Exception as e:
                print(f"[DEBUG] Failed to save draft on exit: {e}")
        self.destroy()

    def _build_search_panel(self, parent_pane):
        self.search_frame = ctk.CTkFrame(parent_pane, fg_color=STYLE["btn_utility_fg"], corner_radius=8, border_width=1, border_color=STYLE["btn_utility_border"])
        
        self.search_frame.columnconfigure(0, weight=0) # Labels
        self.search_frame.columnconfigure(1, weight=1) # Entries
        self.search_frame.columnconfigure(2, weight=0) # Controls / Action buttons
        
        self.search_query_var = ctk.StringVar()
        
        # Row 0: Find
        find_lbl = ctk.CTkLabel(self.search_frame, text="Find:", font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"), text_color=STYLE["text_muted"])
        find_lbl.grid(row=0, column=0, padx=(12, 5), pady=(10, 5), sticky="w")
        
        self.search_entry = ctk.CTkEntry(
            self.search_frame, placeholder_text="Type text to search...",
            textvariable=self.search_query_var,
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12),
            height=28
        )
        self.search_entry.grid(row=0, column=1, padx=5, pady=(10, 5), sticky="ew")
        self.search_entry.bind("<Return>", lambda e: self._find_next())
        self.search_entry.bind("<Escape>", lambda e: self._toggle_search_panel(show=False))
        
        # Triggers perform search on any modification (including paste)
        self.search_query_var.trace_add("write", lambda *args: self._perform_search())
        
        # Controls Frame for Find Row
        find_ctrl = ctk.CTkFrame(self.search_frame, fg_color="transparent", border_width=0)
        find_ctrl.grid(row=0, column=2, padx=(5, 12), pady=(10, 5), sticky="e")
        
        self.match_lbl = ctk.CTkLabel(find_ctrl, text="0 of 0", font=ctk.CTkFont(family=STYLE["font_family_body"], size=12), text_color=STYLE["text_muted"], width=55)
        self.match_lbl.pack(side="left", padx=5)
        
        self.btn_prev = ctk.CTkButton(
            find_ctrl, text="▲", width=24, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._find_prev
        )
        self.btn_prev.pack(side="left", padx=2)
        
        self.btn_next = ctk.CTkButton(
            find_ctrl, text="▼", width=24, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._find_next
        )
        self.btn_next.pack(side="left", padx=2)
        
        self.btn_close_search = ctk.CTkButton(
            find_ctrl, text="×", width=24, height=24, 
            fg_color=STYLE["btn_clear_fg"], 
            hover_color=STYLE["btn_clear_hover"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=14, weight="bold"),
            command=lambda: self._toggle_search_panel(show=False)
        )
        self.btn_close_search.pack(side="left", padx=(8, 0))
        
        # Row 1: Replace
        replace_lbl = ctk.CTkLabel(self.search_frame, text="Replace:", font=ctk.CTkFont(family=STYLE["font_family_body"], size=12, weight="bold"), text_color=STYLE["text_muted"])
        replace_lbl.grid(row=1, column=0, padx=(12, 5), pady=(5, 10), sticky="w")
        
        self.replace_entry = ctk.CTkEntry(
            self.search_frame, placeholder_text="Replace with...",
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=12),
            height=28
        )
        self.replace_entry.grid(row=1, column=1, padx=5, pady=(5, 10), sticky="ew")
        self.replace_entry.bind("<Return>", lambda e: self._replace_current())
        self.replace_entry.bind("<Escape>", lambda e: self._toggle_search_panel(show=False))
        
        # Controls Frame for Replace Row
        replace_ctrl = ctk.CTkFrame(self.search_frame, fg_color="transparent", border_width=0)
        replace_ctrl.grid(row=1, column=2, padx=(5, 12), pady=(5, 10), sticky="e")
        
        self.btn_replace = ctk.CTkButton(
            replace_ctrl, text="Replace", width=70, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._replace_current
        )
        self.btn_replace.pack(side="left", padx=2)
        
        self.btn_replace_all = ctk.CTkButton(
            replace_ctrl, text="Replace All", width=80, height=24, 
            fg_color=STYLE["btn_utility_fg"], 
            hover_color=STYLE["btn_utility_hover"],
            border_color=STYLE["btn_utility_border"],
            border_width=1,
            text_color=STYLE["text_primary"],
            font=ctk.CTkFont(family=STYLE["font_family_body"], size=11, weight="bold"),
            command=self._replace_all
        )
        self.btn_replace_all.pack(side="left", padx=2)

    def _toggle_search_panel(self, show=None):
        if show is None:
            show = not self.search_panel_visible
            
        self.search_panel_visible = show
        
        if show:
            # Grid into row 3 of left_pane
            self.search_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=5)
            self.search_entry.focus()
            self.search_entry.select_range(0, "end")
            self._perform_search()
        else:
            self.search_frame.grid_remove()
            self._clear_search_tags()
            self.editor.focus()

    def _shortcut_find(self, event=None):
        self._toggle_search_panel(show=True)
        self.search_entry.focus()
        self.search_entry.select_range(0, "end")
        return "break"

    def _shortcut_replace(self, event=None):
        self._toggle_search_panel(show=True)
        self.replace_entry.focus()
        self.replace_entry.select_range(0, "end")
        return "break"

    def _clear_search_tags(self):
        self.editor._textbox.tag_remove("search_highlight", "1.0", "end")
        self.editor._textbox.tag_remove("active_highlight", "1.0", "end")

    def _perform_search(self, keep_current_index=False):
        self._clear_search_tags()
        
        query = self.search_entry.get()
        if not query:
            self.matches = []
            self.current_match_idx = -1
            self.match_lbl.configure(text="0 of 0")
            return
            
        import tkinter as tk
        start_idx = "1.0"
        count_var = tk.IntVar()
        
        old_active_pos = None
        if keep_current_index and 0 <= self.current_match_idx < len(self.matches):
            old_active_pos = self.matches[self.current_match_idx][0]
            
        self.matches = []
        
        while True:
            pos = self.editor._textbox.search(query, start_idx, stopindex="end", nocase=True, count=count_var)
            if not pos:
                break
            match_len = count_var.get()
            if match_len == 0:
                break
            end_idx = f"{pos} + {match_len}c"
            self.matches.append((pos, end_idx))
            start_idx = end_idx
            
        if not self.matches:
            self.current_match_idx = -1
            self.match_lbl.configure(text="0 of 0")
            return
            
        if keep_current_index and old_active_pos:
            found = False
            for idx, (pos, _) in enumerate(self.matches):
                if pos == old_active_pos:
                    self.current_match_idx = idx
                    found = True
                    break
            if not found:
                self.current_match_idx = min(self.current_match_idx, len(self.matches) - 1)
        else:
            if self.current_match_idx < 0 or self.current_match_idx >= len(self.matches):
                self.current_match_idx = 0
                
        for pos, end_idx in self.matches:
            self.editor._textbox.tag_add("search_highlight", pos, end_idx)
            
        self._update_active_match()

    def _update_active_match(self):
        self.editor._textbox.tag_remove("active_highlight", "1.0", "end")
        
        if 0 <= self.current_match_idx < len(self.matches):
            pos, end_idx = self.matches[self.current_match_idx]
            self.editor._textbox.tag_add("active_highlight", pos, end_idx)
            self.editor._textbox.see(pos)
            self.match_lbl.configure(text=f"{self.current_match_idx + 1} of {len(self.matches)}")
        else:
            self.match_lbl.configure(text="0 of 0")

    def _find_next(self):
        if not self.matches:
            return
        self.current_match_idx = (self.current_match_idx + 1) % len(self.matches)
        self._update_active_match()
        
    def _find_prev(self):
        if not self.matches:
            return
        self.current_match_idx = (self.current_match_idx - 1) % len(self.matches)
        self._update_active_match()

    def _replace_current(self):
        if self.is_processing:
            return
        if not self.matches or self.current_match_idx < 0 or self.current_match_idx >= len(self.matches):
            return
            
        pos, end_idx = self.matches[self.current_match_idx]
        rep_text = self.replace_entry.get()
        
        old_state = self.editor._textbox.cget("state")
        self.editor.configure(state="normal")
        try:
            self.editor.delete(pos, end_idx)
            self.editor.insert(pos, rep_text)
        finally:
            self.editor.configure(state=old_state)
        
        self.is_dirty = True
        self._update_counts()
        
        self._perform_search(keep_current_index=True)

    def _replace_all(self):
        if self.is_processing:
            return
        if not self.matches:
            return
            
        rep_text = self.replace_entry.get()
        
        old_state = self.editor._textbox.cget("state")
        self.editor.configure(state="normal")
        try:
            for pos, end_idx in reversed(self.matches):
                self.editor.delete(pos, end_idx)
                self.editor.insert(pos, rep_text)
        finally:
            self.editor.configure(state=old_state)
            
        self.is_dirty = True
        self._update_counts()
        
        self.current_match_idx = -1
        self._perform_search()

    def _show_md_guide(self):
        import tkinter as tk
        guide_win = tk.Toplevel(self)
        guide_win.title("Markdown Syntax Guide")
        guide_win.geometry("450x440")
        guide_win.resizable(False, False)
        
        # Resolve background theme color
        bg_color = self.workspace.cget("fg_color")
        if isinstance(bg_color, tuple):
            bg_color = bg_color[1] if ctk.get_appearance_mode() == "Dark" else bg_color[0]
        guide_win.configure(bg=bg_color)
        
        # Make it transient & grab focus (modal behavior)
        guide_win.transient(self)
        guide_win.grab_set()
        guide_win.focus_set()
        
        fg_color = "#ffffff" if ctk.get_appearance_mode() == "Dark" else "#1d1d1f"
        
        title_lbl = tk.Label(
            guide_win, text="Supported Markdown Formatting",
            font=("Segoe UI", 14, "bold"),
            bg=bg_color,
            fg=fg_color
        )
        title_lbl.pack(pady=(15, 10))
        
        text_bg = "#1d202b" if ctk.get_appearance_mode() == "Dark" else "#f3f4f6"
        text_fg = "#f8f8f2" if ctk.get_appearance_mode() == "Dark" else "#1d1d1f"
        
        guide_txt = tk.Text(
            guide_win, width=50, height=18,
            font=("Consolas", 10),
            bg=text_bg,
            fg=text_fg,
            bd=1,
            relief="flat",
            padx=10,
            pady=10
        )
        guide_txt.pack(padx=20, pady=5)
        
        cheat_sheet = (
            "Syntax Cheatsheet:\n"
            "=========================================\n\n"
            "1. Headings (Tiêu đề):\n"
            "   # Heading 1\n"
            "   ## Heading 2\n"
            "   ### Heading 3\n\n"
            "2. Bold (In đậm):\n"
            "   **text** or __text__\n\n"
            "3. Italic (In nghiêng):\n"
            "   *text* or _text_\n\n"
            "4. Strikethrough (Gạch ngang):\n"
            "   ~~text~~\n\n"
            "5. Underline (Gạch chân):\n"
            "   <u>text</u>\n\n"
            "6. Inline Code (Mã dòng):\n"
            "   `code`\n\n"
            "7. Hyperlink (Liên kết):\n"
            "   [Link Text](https://url.com)\n\n"
            "8. Lists (Danh sách):\n"
            "   - Bullet Item 1\n"
            "     - Nested Bullet Item (Indent 2 spaces)\n"
            "   1. Numbered Item 1\n"
            "     1. Nested Numbered (Indent 2 spaces)\n\n"
            "9. Tables (Bảng dữ liệu):\n"
            "   | Header 1 | Header 2 |\n"
            "   | --- | --- |\n"
            "   | **Bold Cell** | [Link](url) |\n"
        )
        guide_txt.insert("1.0", cheat_sheet)
        guide_txt.configure(state="disabled")
        
        btn_close = tk.Button(
            guide_win, text="Close", width=12, height=1,
            font=("Segoe UI", 10, "bold"),
            bg="#5d3fd3" if ctk.get_appearance_mode() == "Dark" else "#ebe4ff",
            fg="#ffffff" if ctk.get_appearance_mode() == "Dark" else "#5d3fd3",
            activebackground=bg_color,
            bd=1,
            relief="flat",
            command=guide_win.destroy
        )
        btn_close.pack(pady=(10, 15))

    def _apply_format(self, prefix, suffix):
        import tkinter as tk
        try:
            start = self.editor._textbox.index("sel.first")
            end = self.editor._textbox.index("sel.last")
            selected_text = self.editor._textbox.get("sel.first", "sel.last")
            new_text = f"{prefix}{selected_text}{suffix}"
            self.editor._textbox.delete("sel.first", "sel.last")
            self.editor._textbox.insert(start, new_text)
            new_end = self.editor._textbox.index(f"{start} + {len(new_text)} chars")
            self.editor._textbox.tag_add("sel", start, new_end)
            self.editor.focus_set()
            self._update_counts()
        except tk.TclError:
            # If no selection, insert placeholder at cursor and highlight the placeholder
            cursor_pos = self.editor._textbox.index("insert")
            self.editor._textbox.insert(cursor_pos, f"{prefix}text{suffix}")
            text_start = self.editor._textbox.index(f"{cursor_pos} + {len(prefix)} chars")
            text_end = self.editor._textbox.index(f"{text_start} + 4 chars")
            self.editor._textbox.tag_add("sel", text_start, text_end)
            # Move insertion cursor to the start of the 'text' placeholder so typing replaces it immediately!
            self.editor._textbox.mark_set("insert", text_start)
            self.editor.focus_set()
            self._update_counts()

    def _apply_heading(self, heading_mark):
        import re
        try:
            cursor_pos = self.editor._textbox.index("insert")
            line_num = cursor_pos.split(".")[0]
            line_start = f"{line_num}.0"
            line_end = f"{line_num}.end"
            
            line_content = self.editor._textbox.get(line_start, line_end)
            cleaned_line = re.sub(r"^#{1,6}\s*", "", line_content)
            new_line = f"{heading_mark} {cleaned_line}"
            
            self.editor._textbox.delete(line_start, line_end)
            self.editor._textbox.insert(line_start, new_line)
            
            self.editor.focus_set()
            self._update_counts()
        except Exception:
            pass

    def _apply_link(self):
        from tkinter import simpledialog
        try:
            start = self.editor._textbox.index("sel.first")
            selected_text = self.editor._textbox.get("sel.first", "sel.last")
            has_selection = True
        except Exception:
            selected_text = "Link Text"
            start = self.editor._textbox.index("insert")
            has_selection = False

        url = simpledialog.askstring("Insert Link", "Enter URL:", parent=self)
        if url:
            new_text = f"[{selected_text}]({url})"
            if has_selection:
                try:
                    self.editor._textbox.delete("sel.first", "sel.last")
                except Exception:
                    pass
            self.editor._textbox.insert(start, new_text)
            
            if not has_selection:
                # Highlight only the placeholder link text
                text_start = self.editor._textbox.index(f"{start} + 1 chars")
                text_end = self.editor._textbox.index(f"{text_start} + {len(selected_text)} chars")
                self.editor._textbox.tag_add("sel", text_start, text_end)
                self.editor._textbox.mark_set("insert", text_start)
                
            self.editor.focus_set()
            self._update_counts()

    def _insert_table(self):
        cursor_pos = self.editor._textbox.index("insert")
        table_template = (
            "\n| Column 1 | Column 2 |\n"
            "| --- | --- |\n"
            "| Row 1 Cell 1 | Row 1 Cell 2 |\n"
            "| Row 2 Cell 1 | Row 2 Cell 2 |\n"
        )
        self.editor._textbox.insert(cursor_pos, table_template)
        self.editor.focus_set()
        self._update_counts()

    def _on_editor_key_press(self, event):
        if event.char and len(event.char) == 1:
            # Check if it is a printable character, Enter, or Tab
            if ord(event.char) >= 32 or event.char in ('\r', '\n', '\t'):
                try:
                    if self.editor._textbox.tag_ranges("sel"):
                        start = self.editor._textbox.index("sel.first")
                        self.editor._textbox.delete("sel.first", "sel.last")
                        self.editor._textbox.mark_set("insert", start)
                except Exception:
                    pass

    def _init_autosave(self):
        # Trigger restoration check asynchronously after 1 second to stabilize layout
        self.after(1000, self._restore_draft_if_needed)
        # Start the periodic background autosave loop
        self._periodic_autosave()

    def _restore_draft_if_needed(self):
        if os.path.exists(DRAFT_PATH):
            try:
                with open(DRAFT_PATH, "r", encoding="utf-8") as f:
                    draft_content = f.read()
                if draft_content.strip():
                    self._show_restore_dialog(draft_content)
            except Exception as e:
                print(f"[DEBUG] Failed to check/restore draft on startup: {e}")

    def _show_restore_dialog(self, draft_content):
        import tkinter as tk
        palette = PALETTES[self.current_palette_var.get()]
        
        # Helper to resolve CustomTkinter color tuples to native strings
        def resolve_color(color_val):
            if isinstance(color_val, (tuple, list)):
                if self.appearance_mode_var.get() == "Light":
                    return color_val[0]
                else:
                    return color_val[1]
            return color_val

        # Determine theme-based colors
        bg_color = resolve_color(palette.get("bg_component", "#1e1e2e"))
        text_color = "#ffffff" if self.appearance_mode_var.get() == "Dark" else "#000000"
        accent_color = resolve_color(palette.get("btn_convert_fg", "#7b2cbf"))
        accent_hover = resolve_color(palette.get("btn_convert_hover", "#9d4edd"))
        discard_color = resolve_color(STYLE.get("status_red", "#ef476f"))
        later_bg = resolve_color(STYLE.get("btn_utility_fg", "#2d3748"))
        later_hover = resolve_color(STYLE.get("btn_utility_hover", "#3a475c"))
        later_fg = resolve_color(STYLE.get("text_primary", "#ffffff"))

        dialog = tk.Toplevel(self)
        dialog.title("Recover Unsaved Draft")
        dialog.geometry("460x180")
        dialog.resizable(False, False)
        dialog.configure(bg=bg_color)
        dialog.transient(self)
        dialog.grab_set()
        dialog.focus_set()
        
        # Center dialog relative to screen
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        dlg_x = (screen_w - 460) // 2
        dlg_y = (screen_h - 180) // 2
        dialog.geometry(f"460x180+{max(dlg_x, 0)}+{max(dlg_y, 0)}")

        # Fonts
        font_family = STYLE["font_family_body"]
        title_font = (font_family, 10, "bold")
        body_font = (font_family, 11)

        lbl_msg = tk.Label(
            dialog,
            text="We found an unsaved draft from your previous session.\nWould you like to restore it?",
            font=body_font,
            bg=bg_color,
            fg=text_color,
            justify="center"
        )
        lbl_msg.pack(pady=(25, 20), padx=20)

        btn_frame = tk.Frame(dialog, bg=bg_color)
        btn_frame.pack(pady=(5, 15))

        def on_restore():
            self.editor.configure(state="normal")
            self.editor.delete("1.0", "end")
            self.editor.insert("1.0", draft_content)
            self._update_counts()
            self._set_status("Restored autosaved draft", "green")
            dialog.destroy()
            try:
                os.remove(DRAFT_PATH)
            except Exception:
                pass

        def on_discard():
            try:
                os.remove(DRAFT_PATH)
            except Exception:
                pass
            self._set_status("Discarded unsaved draft", "orange")
            dialog.destroy()

        def on_later():
            self._set_status("Draft recovery postponed", "orange")
            dialog.destroy()

        btn_restore = tk.Button(
            btn_frame, text="Restore", width=12, height=1,
            font=title_font,
            bg=accent_color,
            fg="#ffffff",
            activebackground=accent_hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            command=on_restore
        )
        btn_restore.pack(side="left", padx=10)

        btn_discard = tk.Button(
            btn_frame, text="Discard", width=12, height=1,
            font=title_font,
            bg=discard_color,
            fg="#ffffff",
            activebackground=discard_color,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            command=on_discard
        )
        btn_discard.pack(side="left", padx=10)

        btn_later = tk.Button(
            btn_frame, text="Later", width=12, height=1,
            font=title_font,
            bg=later_bg,
            fg=later_fg,
            activebackground=later_hover,
            activeforeground=later_fg,
            relief="flat",
            bd=0,
            command=on_later
        )
        btn_later.pack(side="left", padx=10)

    def _periodic_autosave(self):
        if self.is_dirty and not self.is_processing:
            try:
                content = self.editor.get("1.0", "end-1c")
                os.makedirs(os.path.dirname(DRAFT_PATH), exist_ok=True)
                with open(DRAFT_PATH, "w", encoding="utf-8") as f:
                    f.write(content)
            except Exception as e:
                print(f"[DEBUG] Failed to autosave draft: {e}")
        
        # Schedule the next autosave check in 5000 milliseconds (5 seconds)
        self.after(5000, self._periodic_autosave)
