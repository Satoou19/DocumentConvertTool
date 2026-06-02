"""
Universal Document Converter
Supports: MD→Excel, MD→Word, Excel→MD, Word→MD
Dependencies: customtkinter, tkinterdnd2, pandas, openpyxl, python-docx, mammoth
Install: pip install customtkinter tkinterdnd2 pandas openpyxl python-docx 
Run: python md_to_excel.py
Build .exe: pyinstaller --onefile --windowed --name "Document Converter" --icon=favicon.ico md_to_excel.py
    Lưu ý: file favicon.ico phải để cùng thư mục với md_to_excel.py trước khi chạy lệnh.
"""

import sys
import re
import os
import threading
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog

# Prevent UnicodeEncodeError when printing emojis/UTF-8 characters to standard output in Windows console
if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')
    except AttributeError:
        pass
if sys.stderr is not None:
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='backslashreplace')
    except AttributeError:
        pass

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("[WARN] tkinterdnd2 not found, drag & drop disabled")


# ── Conversion modes config ───────────────────────────────────────────────────

MODES = {
    "MD → Excel":  {"in_ext": ".md",   "out_ext": ".xlsx", "in_label": "File .md",   "out_label": "Lưu .xlsx"},
    "MD → Word":   {"in_ext": ".md",   "out_ext": ".docx", "in_label": "File .md",   "out_label": "Lưu .docx"},
    "Excel → MD":  {"in_ext": ".xlsx", "out_ext": ".md",   "in_label": "File .xlsx", "out_label": "Lưu .md"},
    "Word → MD":   {"in_ext": ".docx", "out_ext": ".md",   "in_label": "File .docx", "out_label": "Lưu .md"},
}

IN_FILETYPES = {
    ".md":   [("Markdown", "*.md"), ("All", "*.*")],
    ".xlsx": [("Excel", "*.xlsx *.xls"), ("All", "*.*")],
    ".docx": [("Word", "*.docx"), ("All", "*.*")],
}

OUT_FILETYPES = {
    ".xlsx": [("Excel", "*.xlsx")],
    ".docx": [("Word",  "*.docx")],
    ".md":   [("Markdown", "*.md")],
}


# ── Converters ────────────────────────────────────────────────────────────────

def parse_md_tables(content: str) -> list[tuple[str, pd.DataFrame]]:
    tables, lines, i = [], content.split("\n"), 0
    while i < len(lines):
        line = lines[i].strip()
        if "|" in line and not re.match(r"^[\|\s\-:]+$", line):
            table_name = f"Sheet{len(tables)+1}"
            for j in range(i-1, max(i-5, -1), -1):
                prev = lines[j].strip()
                if prev.startswith("#"):
                    table_name = re.sub(r"^#+\s*", "", prev)[:31]
                    break
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
            data_lines = [l for l in table_lines if not re.match(r"^[\|\s\-:]+$", l)]
            if len(data_lines) < 2:
                continue
            rows = [[c.strip() for c in l.split("|") if c.strip()] for l in data_lines]
            max_cols = max(len(r) for r in rows)
            rows = [r + [""] * (max_cols - len(r)) for r in rows]
            df = pd.DataFrame(rows[1:], columns=rows[0])
            print(f"[DEBUG] Table '{table_name}': {df.shape}")
            tables.append((table_name, df))
        else:
            i += 1
    return tables


def md_to_excel(in_path: str, out_path: str) -> str:
    from openpyxl.styles import Font
    from openpyxl.styles import PatternFill
    from openpyxl.styles import Alignment
    from openpyxl.styles import Border
    from openpyxl.styles import Side
    from openpyxl.utils import get_column_letter

    with open(in_path, encoding="utf-8") as f:
        content = f.read()
    tables = parse_md_tables(content)
    print(f"[DEBUG] Tables found: {len(tables)}")
    if not tables:
        return "❌ Không tìm thấy table nào trong file markdown."

    seen = {}
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in tables:
            key = name
            if key in seen:
                seen[key] += 1
                key = f"{name}_{seen[key]}"
            else:
                seen[name] = 0
            df.to_excel(writer, sheet_name=key, index=False)

            ws = writer.sheets[key]

            header_fill = PatternFill(
                "solid",
                fgColor="4472C4"
            )

            header_font = Font(
                name="Arial",
                size=13,
                bold=True,
                color="FFFFFF"
            )

            body_font = Font(
                name="Arial",
                size=13
            )

            thin = Side(
                border_style="thin",
                color="D9D9D9"
            )
            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            left_align   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
            thin_border  = Border(left=thin, right=thin, top=thin, bottom=thin)
            alt_fill     = PatternFill("solid", fgColor="F2F2F2")

            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                for cell in row:

                    cell.border = thin_border

                    if row_idx == 1:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(
                            horizontal="center",
                            vertical="center",
                            wrap_text=True
                        )
                    else:
                        cell.font = body_font
                        cell.alignment = Alignment(
                            horizontal="left",
                            vertical="center",
                            wrap_text=True
                        )

            for col in ws.columns:
                max_len = 0

                for cell in col:
                    if cell.value:
                        max_len = max(
                            max_len,
                            len(str(cell.value))
                        )

                ws.column_dimensions[
                    get_column_letter(col[0].column)
                ].width = min(max_len + 5, 50)

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            print(f"[DEBUG] Styled sheet: {key}")

    return f"✅ Xuất {len(tables)} sheet(s) → {os.path.basename(out_path)}"


def md_to_word(in_path: str, out_path: str) -> str:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

    FONT = "Arial"
    HEADING_SIZES = {1: 20, 2: 16, 3: 13, 4: 12, 5: 11, 6: 11}
    HEADING_COLORS = {
        1: "404040",
        2: "404040",
        3: "404040",
        4: "404040",
        5: "404040",
        6: "404040"
    }

    def set_font(run, size=13, bold=False, color=None):
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = RGBColor.from_string(color)

    def add_paragraph_with_font(doc, text, size=13, bold=False, color=None, style=None):
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        # Handle **bold** inline
        parts = re.split(r"(\*\*.*?\*\*)", text)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                run = p.add_run(part[2:-2])
                set_font(run, size=size, bold=True, color=color)
            else:
                run = p.add_run(part)
                set_font(run, size=size, bold=bold, color=color)
        return p

    with open(in_path, encoding="utf-8") as f:
        lines = f.readlines()

    doc = Document()

    # Set default font for document
    style = doc.styles["Normal"]
    style.font.name = FONT # type: ignore
    style.font.size = Pt(13) # type: ignore

    i = 0
    while i < len(lines):
        line = lines[i].rstrip("\n")

        # Headings
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            heading = doc.add_heading(
                m.group(2),
                level=min(level, 9)
            )
            for run in heading.runs:
                set_font(
                    run,
                    size=HEADING_SIZES[level],
                    bold=True,
                    color=HEADING_COLORS[level]
                )
            heading.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            heading.paragraph_format.space_before = Pt(10)
            heading.paragraph_format.space_after = Pt(4)
            print(f"[DEBUG] Heading {level}: {m.group(2)[:40]}")
            i += 1
            continue

        # Table
        if "|" in line and not re.match(r"^[\|\s\-:]+$", line.strip()):
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
            data_lines = [l for l in table_lines if not re.match(r"^[\|\s\-:]+$", l)]
            if len(data_lines) >= 2:
                rows = [[c.strip() for c in l.split("|") if c.strip()] for l in data_lines]
                max_cols = max(len(r) for r in rows)
                rows = [r + [""] * (max_cols - len(r)) for r in rows]
                tbl = doc.add_table(rows=len(rows), cols=max_cols)
                tbl.style = "Table Grid"
                for r_idx, row_data in enumerate(rows):
                    for c_idx, cell_text in enumerate(row_data):
                        cell = tbl.cell(r_idx, c_idx)
                        cell.text = ""
                        run = cell.paragraphs[0].add_run(cell_text)
                        if r_idx == 0:
                            set_font(run, size=13, bold=True, color="FFFFFF")
                            # Header cell background: xám đậm
                            tc = cell._tc
                            tcPr = tc.get_or_add_tcPr()
                            shd = OxmlElement("w:shd")
                            shd.set(qn("w:fill"), "4472C4")
                            shd.set(qn("w:color"), "auto")
                            shd.set(qn("w:val"), "clear")
                            tcPr.append(shd)
                        else:
                            set_font(run, size=13)
                print(f"[DEBUG] Table: {len(rows)}×{max_cols}")
                doc.add_paragraph()  # spacing after table
            continue

        # Unordered list
        m = re.match(r"^(\s*)[-*+]\s+(.*)", line)
        if m:
            add_paragraph_with_font(doc, m.group(2), style="List Bullet")
            i += 1
            continue

        # Ordered list
        m = re.match(r"^\s*\d+\.\s+(.*)", line)
        if m:
            print("[DEBUG LIST]", line)

            add_paragraph_with_font(
                doc,
                m.group(1),
                style="List Number"
            )

            i += 1
            continue

        # Horizontal rule
        if re.match(r"^[-*_]{3,}$", line.strip()):
            p = doc.add_paragraph("─" * 50)
            p.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            i += 1
            continue

        # Normal paragraph
        text = re.sub(r"`(.*?)`", r"\1", line)
        if text.strip():
            add_paragraph_with_font(doc, text)
        i += 1

    doc.save(out_path)
    return f"✅ Đã tạo Word → {os.path.basename(out_path)}"


def excel_to_md(in_path: str, out_path: str) -> str:
    xl = pd.ExcelFile(in_path)
    parts = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"[DEBUG] Sheet '{sheet}': {df.shape}")
        parts.append(f"## {sheet}\n")
        # Header
        header = "| " + " | ".join(str(c) for c in df.columns) + " |"
        sep    = "| " + " | ".join("---" for _ in df.columns) + " |"
        parts.append(header)
        parts.append(sep)
        for _, row in df.iterrows():
            parts.append("| " + " | ".join(str(v) for v in row) + " |")
        parts.append("")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return f"✅ Xuất {len(xl.sheet_names)} sheet(s) → {os.path.basename(out_path)}"


def word_to_md(in_path: str, out_path: str) -> str:
    import os
    import re
    import mammoth

    with open(in_path, "rb") as f:
        result = mammoth.convert_to_markdown(f)

    print(f"[DEBUG] mammoth messages: {result.messages}")
    markdown = result.value

    # LÀM SẠCH FILE: Xóa bỏ các ký tự back-slash (\) dư thừa do mammoth tự thêm vào
    markdown = re.sub(r"\\([.\-_~*`\[\]()#+!{}])", r"\1", markdown)

    lines = markdown.splitlines()
    global_counter = 1
    final_lines = []

    for line in lines:
        stripped = line.strip()
        
        # Chỉ bắt các dòng bắt đầu bằng "1. " ở đầu dòng (định dạng mục lớn của Mammoth)
        # và loại trừ các dòng là list con thụt lề sâu hoặc định dạng khác
        if re.match(r"^1\.\s+", stripped):
            # Sử dụng lambda để truyền trực tiếp giá trị của global_counter vào, 
            # tránh hoàn toàn lỗi "invalid group reference" khi counter >= 10
            replaced_line = re.sub(
                r"^(.*?)1\.\s+", 
                lambda m: f"{m.group(1)}{global_counter}. ", 
                line, 
                count=1
            )
            final_lines.append(replaced_line)
            global_counter += 1
        else:
            final_lines.append(line)

    markdown = "\n".join(final_lines)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(markdown)

    return f"✅ Chuyển Word → MD → {os.path.basename(out_path)}"

CONVERTERS = {
    "MD → Excel": md_to_excel,
    "MD → Word":  md_to_word,
    "Excel → MD": excel_to_md,
    "Word → MD":  word_to_md,
}


# ── GUI ───────────────────────────────────────────────────────────────────────

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BaseClass = TkinterDnD.Tk if HAS_DND else ctk.CTk


class App(BaseClass): # type: ignore
    def __init__(self):
        super().__init__()
        self.title("Document Converter")
        self.geometry("540x520")
        self.resizable(False, False)

        self.in_path  = ctk.StringVar()
        self.out_path = ctk.StringVar()
        self.mode_var = ctk.StringVar(value="MD → Excel")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 20, "pady": 6}

        # Title
        ctk.CTkLabel(self, text="📄 Document Converter",
                     font=ctk.CTkFont(size=22, weight="bold"),text_color="black").pack(pady=(22, 2), )
        ctk.CTkLabel(self, text="MD ↔ Excel / Word",
                     text_color="gray").pack()

        # Mode selector
        mode_row = ctk.CTkFrame(self, fg_color="transparent")
        mode_row.pack(pady=(14, 2))
        ctk.CTkLabel(mode_row, text="Mode:", width=55, anchor="w").pack(side="left")
        self.mode_menu = ctk.CTkOptionMenu(
            mode_row, values=list(MODES.keys()),
            variable=self.mode_var, width=200,
            command=self._on_mode_change
        )
        self.mode_menu.pack(side="left", padx=6)

        # Drag & drop zone
        self.drop_frame = ctk.CTkFrame(self, width=480, height=100,
                                       corner_radius=12, border_width=2,
                                       border_color="#3a7ebf")
        self.drop_frame.pack(padx=20, pady=(10, 6))
        self.drop_frame.pack_propagate(False)

        self.drop_label = ctk.CTkLabel(
            self.drop_frame,
            text="📂  Kéo file vào đây\nhoặc dùng nút Browse bên dưới",
            font=ctk.CTkFont(size=13), text_color="#7eb8f5"
        )
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")

        if HAS_DND:
            for w in (self.drop_frame, self.drop_label):
                w.drop_target_register(DND_FILES) # type: ignore
                w.dnd_bind("<<Drop>>", self._on_drop) # type: ignore

        # Input row
        self.row_in = ctk.CTkFrame(self, fg_color="transparent")
        self.row_in.pack(fill="x", **pad)
        self.lbl_in = ctk.CTkLabel(self.row_in, text="Input:", width=55, anchor="w")
        self.lbl_in.pack(side="left")
        ctk.CTkEntry(self.row_in, textvariable=self.in_path,
                     width=330, placeholder_text="Chọn file...").pack(side="left", padx=(4, 6))
        ctk.CTkButton(self.row_in, text="Browse", width=70,
                      command=self._browse_input).pack(side="left")

        # Output row
        row_out = ctk.CTkFrame(self, fg_color="transparent")
        row_out.pack(fill="x", **pad)
        self.lbl_out = ctk.CTkLabel(row_out, text="Output:", width=55, anchor="w")
        self.lbl_out.pack(side="left")
        ctk.CTkEntry(row_out, textvariable=self.out_path,
                     width=330, placeholder_text="Chọn nơi lưu...").pack(side="left", padx=(4, 6))
        ctk.CTkButton(row_out, text="Browse", width=70,
                      command=self._browse_output).pack(side="left")

        # Convert button
        self.btn = ctk.CTkButton(
            self, text="⚡  Convert", height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._run
        )
        self.btn.pack(padx=20, pady=(14, 4))

        # Status
        self.status_lbl = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=13))
        self.status_lbl.pack()

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _cfg(self):
        return MODES[self.mode_var.get()]

    def _on_mode_change(self, _=None):
        cfg = self._cfg()
        self.lbl_in.configure(text=cfg["in_label"] + ":")
        self.lbl_out.configure(text=cfg["out_label"] + ":")
        # Reset paths & drop zone khi đổi mode
        self.in_path.set("")
        self.out_path.set("")
        self._reset_drop_zone()
        self._set_status("")

    def _reset_drop_zone(self):
        self.drop_frame.configure(border_color="#3a7ebf")
        self.drop_label.configure(
            text="📂  Kéo file vào đây\nhoặc dùng nút Browse bên dưới",
            text_color="#7eb8f5"
        )

    def _auto_output(self, in_path: str):
        base = os.path.splitext(in_path)[0]
        self.out_path.set(base + self._cfg()["out_ext"])

    def _update_drop_zone(self, filename: str):
        self.drop_frame.configure(border_color="#2ecc71")
        self.drop_label.configure(text=f"✅  {filename}", text_color="#2ecc71")

    def _set_status(self, msg: str, color: str = "white"):
        self.status_lbl.configure(text=msg, text_color=color)

    # ── Event handlers ────────────────────────────────────────────────────────

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        expected_ext = self._cfg()["in_ext"]
        print(f"[DEBUG] Dropped: {path}")
        if path.lower().endswith(expected_ext):
            self.in_path.set(path)
            self._auto_output(path)
            self._update_drop_zone(os.path.basename(path))
        else:
            self._set_status(f"⚠️ Mode này cần file {expected_ext}", color="orange")

    def _browse_input(self):
        ext = self._cfg()["in_ext"]
        path = filedialog.askopenfilename(filetypes=IN_FILETYPES[ext])
        if path:
            self.in_path.set(path)
            self._auto_output(path)
            self._update_drop_zone(os.path.basename(path))

    def _browse_output(self):
        ext = self._cfg()["out_ext"]
        path = filedialog.asksaveasfilename(
            defaultextension=ext, filetypes=OUT_FILETYPES[ext]
        )
        if path:
            self.out_path.set(path)

    def _run(self):
        inp = self.in_path.get().strip()
        out = self.out_path.get().strip()
        mode = self.mode_var.get()

        if not inp:
            self._set_status("⚠️ Chưa chọn file input", "orange"); return
        if not out:
            self._set_status("⚠️ Chưa chọn output", "orange"); return
        if not os.path.exists(inp):
            self._set_status("❌ File không tồn tại", "red"); return

        self.btn.configure(state="disabled", text="Converting...")
        self._set_status("⏳ Đang xử lý...", "gray")

        def task():
            try:
                msg = CONVERTERS[mode](inp, out)
                color = "green" if msg.startswith("✅") else "red"
            except Exception as e:
                msg = f"❌ Lỗi: {e}"
                color = "red"
                print(f"[DEBUG] Exception: {e}")
            self.after(0, lambda: self._set_status(msg, color))
            self.after(0, lambda: self.btn.configure(state="normal", text="⚡  Convert"))

        threading.Thread(target=task, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()