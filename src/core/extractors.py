import re
import zipfile

def extract_excel_to_md(in_path: str) -> str:
    """Extracts Excel sheets into clean Markdown tables."""
    import pandas as pd
    xl = pd.ExcelFile(in_path)
    parts = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet).fillna("")
        parts.append(f"## {sheet}\n")
        # Generate Markdown Table representation
        if not df.empty:
            header = "| " + " | ".join(str(c) for c in df.columns) + " |"
            sep    = "| " + " | ".join("---" for _ in df.columns) + " |"
            parts.append(header)
            parts.append(sep)
            for _, row in df.iterrows():
                parts.append("| " + " | ".join(str(v) for v in row) + " |")
        else:
            parts.append("*(Empty Table)*")
        parts.append("")
    return "\n".join(parts)


def extract_word_to_md(in_path: str) -> str:
    """Extracts Word .docx to clean Markdown text, preserving tables, headings, bold/italic styles, and lists."""
    import docx
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    # Validate DOCX file
    import zipfile
    try:
        if not zipfile.is_zipfile(in_path):
            return "⚠️ File Validation Error: Invalid file. Please select a valid DOCX document."
    except Exception as e:
        return f"⚠️ File Validation Error: {e}"

    try:
        doc = docx.Document(in_path)
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
                    # Empty paragraph
                    parts.append("")
                    continue
                
                style_name = (block.style.name or "").lower() if block.style else ""
                
                def run_contains_image(run):
                    xml = getattr(run._element, "xml", "")
                    return "<w:drawing" in xml or "<w:pict" in xml or "<a:blip" in xml

                # Check runs for bold / italic formatting
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
                    
                    # Wrap formatting
                    r_text = run_text
                    if run.bold:
                        r_text = r_text.replace(stripped_run, f"**{stripped_run}**")
                    if run.italic:
                        r_text = r_text.replace(stripped_run, f"*{stripped_run}*")
                    para_parts.append(r_text)
                
                para_text = "".join(para_parts).strip()
                if not para_text:
                    para_text = text

                # Apply block formatting based on style
                if style_name.startswith("heading "):
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
                # Convert tables to Markdown tables
                table_parts = []
                rows_data = []
                for row in block.rows:
                    row_cells = []
                    for cell in row.cells:
                        cell_text = cell.text.strip().replace("\n", " ")
                        row_cells.append(cell_text)
                    rows_data.append(row_cells)
                
                if rows_data:
                    # Header row
                    header = "| " + " | ".join(rows_data[0]) + " |"
                    # Separator line
                    sep = "| " + " | ".join("---" for _ in rows_data[0]) + " |"
                    table_parts.append(header)
                    table_parts.append(sep)
                    # Data rows
                    for row in rows_data[1:]:
                        table_parts.append("| " + " | ".join(row) + " |")
                    parts.append("\n".join(table_parts))
                    parts.append("") # Blank line after table

        return "\n\n".join(parts)
    except Exception as e:
        error_msg = f"⚠️ Word Extraction Error:\n\nDetails: {str(e)}\n\nFile: {in_path}"
        print(f"[ERROR] extract_word_to_md: {error_msg}")
        return error_msg


def extract_csv_to_md(in_path: str) -> str:
    """Extracts CSV table into clean Markdown table."""
    import pandas as pd
    try:
        # Read using utf-8-sig to preserve BOM and unicode text (e.g. Vietnamese)
        df = pd.read_csv(in_path, encoding="utf-8-sig", keep_default_na=False)
        if df.empty:
            return "*(Empty Table)*"
        
        parts = []
        # Generate Markdown Table representation
        header = "| " + " | ".join(str(c) for c in df.columns) + " |"
        sep    = "| " + " | ".join("---" for _ in df.columns) + " |"
        parts.append(header)
        parts.append(sep)
        for _, row in df.iterrows():
            parts.append("| " + " | ".join(str(v).replace("\n", " ") for v in row) + " |")
        return "\n".join(parts)
    except Exception as e:
        error_msg = f"⚠️ CSV Extraction Error:\n\nDetails: {str(e)}\n\nFile: {in_path}"
        print(f"[ERROR] extract_csv_to_md: {error_msg}")
        return error_msg

