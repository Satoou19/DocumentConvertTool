import os
import pandas as pd
from src.core.base_module import BaseDocumentModule
from src.core.registry import ModuleRegistry

class CSVModule(BaseDocumentModule):
    @property
    def name(self) -> str:
        return "CSV"

    @property
    def file_extensions(self) -> list[str]:
        return [".csv"]

    @property
    def required_dependencies(self) -> list[str]:
        return ["pandas"]

    def load_to_markdown(self, file_path: str) -> str:
        """Extracts CSV table into clean Markdown table."""
        # Read using utf-8-sig to preserve BOM and unicode text (e.g. Vietnamese)
        df = pd.read_csv(file_path, encoding="utf-8-sig", keep_default_na=False)
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

    def save_from_markdown(self, markdown_content: str, out_path: str) -> str:
        """Converts Markdown content line-by-line into a CSV file, writing text rows to Column A and table cells across columns."""
        import csv
        import re
        from src.core.converters import strip_markdown_styles

        lines = markdown_content.splitlines()
        
        with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            
            for line in lines:
                stripped = line.strip()
                if not stripped:
                    writer.writerow([])
                    continue
                    
                # Skip separator lines
                if "|" in stripped and re.match(r"^[\|\s\-:]+$", stripped):
                    continue
                    
                # Check if it's a table row
                if "|" in stripped:
                    inner_line = stripped
                    if inner_line.startswith("|"):
                        inner_line = inner_line[1:]
                    if inner_line.endswith("|"):
                        inner_line = inner_line[:-1]
                        
                    cells = [strip_markdown_styles(c.strip()) for c in inner_line.split("|")]
                    writer.writerow(cells)
                else:
                    # Write plain text row in Column A (remove heading markdown markers)
                    match_heading = re.match(r"^(#{1,6})\s+(.*)", stripped)
                    if match_heading:
                        text_val = strip_markdown_styles(match_heading.group(2))
                    else:
                        text_val = strip_markdown_styles(line)
                    writer.writerow([text_val])
                    
        return f"Exported successfully to CSV -> {os.path.basename(out_path)}"

ModuleRegistry.register(CSVModule())
