import os
import pandas as pd
from src.core.base_module import BaseDocumentModule
from src.core.registry import ModuleRegistry
from src.core.converters import parse_md_tables

class ExcelModule(BaseDocumentModule):
    @property
    def name(self) -> str:
        return "Excel"

    @property
    def file_extensions(self) -> list[str]:
        return [".xlsx", ".xls"]

    @property
    def required_dependencies(self) -> list[str]:
        return ["pandas", "openpyxl"]

    def load_to_markdown(self, file_path: str) -> str:
        """Extracts Excel sheets into clean Markdown tables."""
        xl = pd.ExcelFile(file_path)
        parts = []
        for sheet in xl.sheet_names:
            df = xl.parse(sheet).fillna("")
            parts.append(f"## {sheet}\n")
            if not df.empty:
                header = "| " + " | ".join(str(c) for c in df.columns) + " |"
                sep    = "| " + " | ".join("---" for _ in df.columns) + " |"
                parts.append(header)
                parts.append(sep)
                for _, row in df.iterrows():
                    cells = [str(v).replace("\n", " ").replace("|", "\\|") for v in row]
                    parts.append("| " + " | ".join(cells) + " |")
            else:
                parts.append("*(Empty Table)*")
            parts.append("")
        return "\n".join(parts)

    def save_from_markdown(self, markdown_content: str, out_path: str) -> str:
        """Converts Markdown tables to formatted Excel spreadsheet."""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        tables = parse_md_tables(markdown_content)
        if not tables:
            return (
                "No tables found in the Markdown content.\n\n"
                "To convert to Excel, please ensure your Markdown content has tables that follow the standard Markdown format, for example:\n\n"
                "| Column 1 | Column 2 |\n"
                "| --- | --- |\n"
                "| Value 1 | Value 2 |\n\n"
                "Make sure you include the separator row (the line with dashes like '| --- | --- |') below the header row."
            )

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

                header_fill = PatternFill("solid", fgColor="4472C4")
                header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
                body_font = Font(name="Arial", size=11)
                thin = Side(border_style="thin", color="D9D9D9")
                thin_border  = Border(left=thin, right=thin, top=thin, bottom=thin)

                for row_idx, row in enumerate(ws.iter_rows(), start=1):
                    for cell in row:
                        cell.border = thin_border
                        if row_idx == 1:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        else:
                            cell.font = body_font
                            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                for col in ws.columns:
                    max_len = 0
                    for cell in col:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 5, 50)

                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

        return f"Exported {len(tables)} sheet(s) successfully -> {os.path.basename(out_path)}"

ModuleRegistry.register(ExcelModule())
