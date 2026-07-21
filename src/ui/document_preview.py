import os
import re
import customtkinter as ctk

class DocumentPreviewFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.palette = None
        self.style = None
        self._current_text = ""
        self.columnconfigure(0, weight=1)
        self._resize_timer = None
        self.bind("<Configure>", self._on_configure, add="+")
        self._current_width = 400

    def set_theme(self, palette, style):
        self.palette = palette
        self.style = style
        self.configure(fg_color=palette["bg_pure_dark"], scrollbar_fg_color=palette["bg_pure_dark"])
        # Re-render current text with new theme
        self.update_preview(self._current_text)

    def _on_configure(self, event):
        if abs(event.width - self._current_width) > 15:
            self._current_width = event.width
            if self._resize_timer is not None:
                try:
                    self.after_cancel(self._resize_timer)
                except Exception:
                    pass
            self._resize_timer = self.after(150, self._update_wraplengths)

    def _update_wraplengths(self):
        new_width = self._current_width - 45
        if new_width < 100:
            return
        
        # Traverse and update wraplengths dynamically
        def update_widget(w, width):
            if isinstance(w, ctk.CTkLabel):
                w.configure(wraplength=width)
            elif isinstance(w, ctk.CTkFrame):
                # Recurse inside container frames (e.g. lists, code boxes)
                for child in w.winfo_children():
                    update_widget(child, width - 20)

        for child in self.winfo_children():
            update_widget(child, new_width)

        # Force scrollregion update to fit the newly wrapped widget heights
        try:
            self._parent_canvas.configure(scrollregion=self._parent_canvas.bbox("all"))
        except Exception:
            pass

    def update_preview(self, markdown_text: str):
        self._current_text = markdown_text
        
        # Clear existing widgets
        for widget in self.winfo_children():
            widget.destroy()

        if not markdown_text.strip():
            # Show empty preview placeholder
            lbl = ctk.CTkLabel(
                self, 
                text="Document Preview is empty.\nType Markdown in the editor to see visual preview.",
                font=ctk.CTkFont(family=self.style["font_family_body"] if self.style else "Segoe UI", size=12, slant="italic"),
                text_color=self.style["text_muted"] if self.style else "#8f93a7"
            )
            lbl.pack(pady=40, fill="x")
            return

        blocks = self._parse_blocks(markdown_text)
        
        for block_type, data in blocks:
            if block_type == "heading":
                self._render_heading(data)
            elif block_type == "code":
                self._render_code(data)
            elif block_type == "table":
                self._render_table(data)
            elif block_type == "list_item":
                self._render_list_item(data)
            else:
                self._render_paragraph(data)
                
        # Force a small delayed update of wraplengths to fit the initial render width
        self.after(50, self._update_wraplengths)

    def _parse_blocks(self, text: str) -> list:
        lines = text.splitlines()
        blocks = []
        current_block = []
        in_code_block = False
        in_table = False
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # Code block starts or ends
            if stripped.startswith("```"):
                if in_code_block:
                    blocks.append(("code", "\n".join(current_block)))
                    current_block = []
                    in_code_block = False
                else:
                    if current_block:
                        blocks.append(self._determine_block_type(current_block))
                        current_block = []
                    in_code_block = True
                i += 1
                continue
                
            if in_code_block:
                current_block.append(line)
                i += 1
                continue
                
            # Table lines
            if "|" in line:
                if not in_table:
                    if current_block:
                        blocks.append(self._determine_block_type(current_block))
                        current_block = []
                    in_table = True
                current_block.append(line)
                i += 1
                continue
            elif in_table:
                blocks.append(("table", current_block))
                current_block = []
                in_table = False
                
            # Empty line -> flush current block
            if not stripped:
                if current_block:
                    blocks.append(self._determine_block_type(current_block))
                    current_block = []
                i += 1
                continue
                
            # Headings
            if stripped.startswith("#"):
                if current_block:
                    blocks.append(self._determine_block_type(current_block))
                    current_block = []
                blocks.append(("heading", line))
                i += 1
                continue
                
            # List item
            if stripped.startswith(("- ", "* ", "• ")) or re.match(r"^\d+\.\s+", stripped):
                if current_block:
                    blocks.append(self._determine_block_type(current_block))
                    current_block = []
                blocks.append(("list_item", line))
                i += 1
                continue
                
            # Paragraph text accumulator
            current_block.append(line)
            i += 1
            
        # Flush final blocks
        if current_block:
            if in_code_block:
                blocks.append(("code", "\n".join(current_block)))
            elif in_table:
                blocks.append(("table", current_block))
            else:
                blocks.append(self._determine_block_type(current_block))
                
        return blocks

    def _determine_block_type(self, lines: list) -> tuple:
        text = " ".join(l.strip() for l in lines)
        return ("paragraph", text)

    def _render_heading(self, text: str):
        stripped = text.strip()
        level = 0
        while level < len(stripped) and stripped[level] == "#":
            level += 1
            
        content = stripped[level:].strip()
        
        # Strip trailing hashes and inline Markdown formatting
        content = re.sub(r"#+$", "", content).strip()
        content = self._strip_bold_italic(content)
        
        font_sizes = {1: 18, 2: 15, 3: 13}
        size = font_sizes.get(level, 12)
        weight = "bold"
        
        lbl = ctk.CTkLabel(
            self,
            text=content,
            font=ctk.CTkFont(family=self.style["font_family_title"] if self.style else "Segoe UI", size=size, weight=weight),
            text_color=self.palette["text_accent_primary"] if self.palette else None,
            anchor="w",
            justify="left"
        )
        lbl.pack(fill="x", padx=10, pady=(12, 4))
        
        if level <= 2:
            sep = ctk.CTkFrame(self, height=1, fg_color=self.palette["border_color"] if self.palette else "#e1e4e8")
            sep.pack(fill="x", padx=10, pady=(0, 4))

    def _render_paragraph(self, text: str):
        content = self._strip_bold_italic(text)
        lbl = ctk.CTkLabel(
            self,
            text=content,
            font=ctk.CTkFont(family=self.style["font_family_body"] if self.style else "Segoe UI", size=12),
            text_color=self.style["text_primary"] if self.style else "#1d1d1f",
            anchor="w",
            justify="left"
        )
        lbl.pack(fill="x", padx=10, pady=3)

    def _render_list_item(self, text: str):
        stripped = text.strip()
        depth = len(text) - len(text.lstrip())
        indent_spaces = max(10, depth * 4 + 10)
        
        content = self._strip_bold_italic(stripped)
        
        # If it starts with a standard bullet, keep a bullet character
        if content.startswith("- ") or content.startswith("* "):
            content = "• " + content[2:]
            
        lbl = ctk.CTkLabel(
            self,
            text=content,
            font=ctk.CTkFont(family=self.style["font_family_body"] if self.style else "Segoe UI", size=12),
            text_color=self.style["text_primary"] if self.style else "#1d1d1f",
            anchor="w",
            justify="left"
        )
        lbl.pack(fill="x", padx=(indent_spaces, 10), pady=2)

    def _render_code(self, text: str):
        bg = self.palette["bg_component"] if self.palette else "#f6f8fa"
        border = self.palette["border_color"] if self.palette else "#e1e4e8"
        
        frame = ctk.CTkFrame(self, fg_color=bg, border_color=border, border_width=1, corner_radius=6)
        frame.pack(fill="x", padx=10, pady=5)
        
        lbl = ctk.CTkLabel(
            frame,
            text=text.strip(),
            font=ctk.CTkFont(family=self.style["font_family_mono"] if self.style else "Consolas", size=11),
            text_color=self.style["text_editor_fg"] if self.style else "#f8f8f2",
            anchor="w",
            justify="left"
        )
        lbl.pack(fill="both", padx=10, pady=6)

    def _render_table(self, lines: list):
        from src.core.converters import strip_markdown_styles
        
        data_lines = [l.strip() for l in lines if not re.match(r"^[\|\s\-:]+$", l.strip())]
        if len(data_lines) < 1:
            return
            
        rows = []
        for l in data_lines:
            row_cells = []
            inner = l
            if inner.startswith("|"):
                inner = inner[1:]
            if inner.endswith("|"):
                inner = inner[:-1]
            for c in inner.split("|"):
                row_cells.append(strip_markdown_styles(c.strip()))
            rows.append(row_cells)
            
        if not rows:
            return
            
        max_cols = max(len(r) for r in rows)
        # Pad shorter rows
        rows = [r + [""] * (max_cols - len(r)) for r in rows]
        
        bg = self.palette["bg_component"] if self.palette else "#ffffff"
        border = self.palette["border_color"] if self.palette else "#e1e4e8"
        
        table_frame = ctk.CTkFrame(self, fg_color="transparent")
        table_frame.pack(fill="x", padx=10, pady=6)
        
        # Configure columns
        for col_idx in range(max_cols):
            table_frame.columnconfigure(col_idx, weight=1)
            
        # Draw grid cells
        for row_idx, row_cells in enumerate(rows):
            is_header = (row_idx == 0)
            row_bg = self.palette["bg_header"] if (is_header and self.palette) else bg
            if not is_header and row_idx % 2 == 0 and self.palette:
                row_bg = self.palette["bg_pane"] # Zebra striping
                
            for col_idx, cell_text in enumerate(row_cells):
                cell_frame = ctk.CTkFrame(
                    table_frame, 
                    fg_color=row_bg, 
                    border_color=border, 
                    border_width=1, 
                    corner_radius=0
                )
                cell_frame.grid(row=row_idx, column=col_idx, sticky="nsew", padx=0, pady=0)
                
                font_weight = "bold" if is_header else "normal"
                lbl = ctk.CTkLabel(
                    cell_frame,
                    text=cell_text,
                    font=ctk.CTkFont(family=self.style["font_family_body"] if self.style else "Segoe UI", size=11, weight=font_weight),
                    text_color=self.style["text_primary"] if self.style else "#1d1d1f",
                    anchor="w",
                    justify="left"
                )
                lbl.pack(fill="both", padx=6, pady=4)

    def _strip_bold_italic(self, text: str) -> str:
        # Strip basic bold ** or __, italic * or _, and html tag <u>
        t = re.sub(r"\*\*\*(.*?)\*\*\*", r"\1", text)
        t = re.sub(r"\*\*(.*?)\*\*", r"\1", t)
        t = re.sub(r"\*(.*?)\*", r"\1", t)
        t = re.sub(r"__(.*?)__", r"\1", t)
        t = re.sub(r"_(.*?)_", r"\1", t)
        t = re.sub(r"~~(.*?)~~", r"\1", t)
        t = re.sub(r"(?i)<u>(.*?)</u>", r"\1", t)
        # Convert links [text](url) to text (url)
        t = re.sub(r"\[(.*?)\]\((.*?)\)", r"\1 (\2)", t)
        return t
