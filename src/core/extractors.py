import os
import re
import pandas as pd

def extract_excel_to_md(in_path: str) -> str:
    """Extracts Excel sheets into clean Markdown tables."""
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
    """Extracts Word .docx to clean Markdown text using Mammoth."""
    import mammoth
    with open(in_path, "rb") as f:
        result = mammoth.convert_to_markdown(f)
    markdown = result.value
    markdown = re.sub(r"\\([.\-_~*`\[\]()#+!{}])", r"\1", markdown)
    
    # Clean list numbers
    lines = markdown.splitlines()
    global_counter = 1
    final_lines = []
    for line in lines:
        stripped = line.strip()
        if re.match(r"^1\.\s+", stripped):
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
    return "\n".join(final_lines)
