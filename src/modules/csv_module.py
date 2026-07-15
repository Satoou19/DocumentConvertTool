import os
import pandas as pd
from src.core.base_module import BaseDocumentModule
from src.core.registry import ModuleRegistry
from src.core.converters import parse_md_tables

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
        """Converts the first Markdown table from the content text into a CSV file."""
        tables = parse_md_tables(markdown_content)
        if not tables:
            return (
                "No tables found in the Markdown content.\n\n"
                "To convert to CSV, please ensure your Markdown content has tables that follow the standard Markdown format, for example:\n\n"
                "| Column 1 | Column 2 |\n"
                "| --- | --- |\n"
                "| Value 1 | Value 2 |\n\n"
                "Make sure you include the separator row (the line with dashes like '| --- | --- |') below the header row."
            )
        
        # Save the first parsed table
        name, df = tables[0]
        from src.core.converters import strip_markdown_styles
        for col in df.columns:
            df[col] = df[col].apply(lambda x: strip_markdown_styles(str(x)))
        df.columns = [strip_markdown_styles(str(c)) for c in df.columns]
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
        return f"Exported table successfully to CSV -> {os.path.basename(out_path)}"

ModuleRegistry.register(CSVModule())
