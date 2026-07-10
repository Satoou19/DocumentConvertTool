import os
import re
import docx
import zipfile
from src.core.base_module import BaseDocumentModule
from src.core.registry import ModuleRegistry

class WordModule(BaseDocumentModule):
    @property
    def name(self) -> str:
        return "Word"

    @property
    def file_extensions(self) -> list[str]:
        return [".docx"]

    @property
    def required_dependencies(self) -> list[str]:
        return ["python-docx"]

    def load_to_markdown(self, file_path: str) -> str:
        """Extracts Word .docx to clean Markdown text, preserving tables, headings, bold/italic styles, and lists."""
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl
        from docx.text.paragraph import Paragraph
        from docx.table import Table

        if not zipfile.is_zipfile(file_path):
            return "⚠️ File Validation Error: Invalid file. Please select a valid DOCX document."

        doc = docx.Document(file_path)
        parts = []

        def iter_block_items(parent):
            parent_elm = parent.element.body
            for child in parent_elm.iterchildren():
                if isinstance(child, CT_P):
                    yield Paragraph(child, parent)
                elif isinstance(child, CT_Tbl):
                    yield Table(child, parent)

        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                text = block.text.strip()
                if not text:
                    parts.append("")
                    continue
                
                style_name = (block.style.name or "").lower() if block.style else ""
                
                def run_contains_image(run):
                    xml = getattr(run._element, "xml", "")
                    return "<w:drawing" in xml or "<w:pict" in xml or "<a:blip" in xml

                para_parts = []
                for run in block.runs:
                    if run_contains_image(run):
                        para_parts.append("[image]")
                        continue
                    run_text = run.text
                    if not run_text:
                        continue
                    stripped_run = run_text.strip()
                    if not stripped_run:
                        para_parts.append(run_text)
                        continue
                    
                    r_text = run_text
                    if run.bold:
                        r_text = r_text.replace(stripped_run, f"**{stripped_run}**")
                    if run.italic:
                        r_text = r_text.replace(stripped_run, f"*{stripped_run}*")
                    para_parts.append(r_text)
                
                para_text = "".join(para_parts).strip()
                if not para_text:
                    para_text = text
                
                style_name = (block.style.name or "").lower() if block.style else ""
                is_heading = style_name.startswith("heading ") or re.match(r"^(đề mục|tiêu đề)\s*\d", style_name)

                if is_heading:
                    try:
                        level = int(style_name.split()[-1])
                        parts.append("#" * level + " " + para_text)
                    except ValueError:
                        parts.append(para_text)
                elif "bullet" in style_name:
                    parts.append("- " + para_text)
                elif "number" in style_name or style_name.startswith("list"):
                    parts.append("1. " + para_text)
                else:
                    parts.append(para_text)

            elif isinstance(block, Table):
                table_parts = []
                rows_data = []
                for row in block.rows:
                    row_cells = []
                    for cell in row.cells:
                        cell_text = cell.text.strip().replace("\n", " ").replace("|", "\\|")
                        row_cells.append(cell_text)
                    rows_data.append(row_cells)
                
                if rows_data:
                    header = "| " + " | ".join(rows_data[0]) + " |"
                    sep = "| " + " | ".join("---" for _ in rows_data[0]) + " |"
                    table_parts.append(header)
                    table_parts.append(sep)
                    for row in rows_data[1:]:
                        table_parts.append("| " + " | ".join(row) + " |")
                    parts.append("\n".join(table_parts))
                    parts.append("")

        return "\n\n".join(parts)

    def save_from_markdown(self, markdown_content: str, out_path: str) -> str:
        """Converts Markdown text to formatted Word document."""
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        FONT = "Arial"
        HEADING_SIZES = {1: 20, 2: 16, 3: 13, 4: 12, 5: 11, 6: 11}
        HEADING_COLORS = {i: "404040" for i in range(1, 7)}

        def set_font(run, size=11, bold=False, color=None):
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.bold = bold
            if color:
                run.font.color.rgb = RGBColor.from_string(color)

        def add_paragraph_with_font(doc, text, size=11, bold=False, color=None, style=None):
            p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
            parts = re.split(r"(\*\*.*?\*\*)", text)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = p.add_run(part[2:-2])
                    set_font(run, size=size, bold=True, color=color)
                else:
                    run = p.add_run(part)
                    set_font(run, size=size, bold=bold, color=color)
            return p

        lines = markdown_content.splitlines()
        doc = Document()
        style = doc.styles["Normal"]
        style.font.name = FONT # type: ignore
        style.font.size = Pt(11) # type: ignore

        i = 0
        while i < len(lines):
            line = lines[i].rstrip("\r\n")

            m = re.match(r"^(#{1,6})\s+(.*)", line)
            if m:
                level = len(m.group(1))
                heading = doc.add_heading(m.group(2), level=min(level, 9))
                for run in heading.runs:
                    set_font(run, size=HEADING_SIZES[level], bold=True, color=HEADING_COLORS[level])
                heading.paragraph_format.space_before = Pt(10)
                heading.paragraph_format.space_after = Pt(4)
                i += 1
                continue

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
                                set_font(run, size=11, bold=True, color="FFFFFF")
                                tc = cell._tc
                                tcPr = tc.get_or_add_tcPr()
                                shd = OxmlElement("w:shd")
                                shd.set(qn("w:fill"), "4472C4")
                                shd.set(qn("w:color"), "auto")
                                shd.set(qn("w:val"), "clear")
                                tcPr.append(shd)
                            else:
                                set_font(run, size=11)
                    doc.add_paragraph()
                continue

            m = re.match(r"^(\s*)[-*+]\s+(.*)", line)
            if m:
                add_paragraph_with_font(doc, m.group(2), style="List Bullet")
                i += 1
                continue

            m = re.match(r"^\s*\d+\.\s+(.*)", line)
            if m:
                add_paragraph_with_font(doc, m.group(1), style="List Number")
                i += 1
                continue

            if re.match(r"^[-*_]{3,}$", line.strip()):
                p = doc.add_paragraph("─" * 50)
                p.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)
                i += 1
                continue

            text = re.sub(r"`(.*?)`", r"\1", line)
            if text.strip():
                add_paragraph_with_font(doc, text)
            i += 1

        doc.save(out_path)
        return f"Word document created successfully -> {os.path.basename(out_path)}"

ModuleRegistry.register(WordModule())
