import os
from typing import List

from src.core.converters import (
    md_to_excel_from_text,
    md_to_word_from_text,
    save_markdown_from_text,
    parse_md_tables,
    md_to_csv_from_text,
)
from src.core.validator import validate_md_tables


def has_md_tables(content: str) -> bool:
    return bool(parse_md_tables(content))


def get_md_table_warnings(content: str) -> List[str]:
    return validate_md_tables(content)


def is_output_locked(out_path: str) -> bool:
    if not os.path.exists(out_path):
        return False
    try:
        with open(out_path, "r+b"):
            return False
    except PermissionError:
        return True
    except Exception:
        return False


def convert_content(mode: str, content: str, out_path: str) -> str:
    if mode == "MD -> Excel":
        return md_to_excel_from_text(content, out_path)
    if mode == "MD -> Word":
        return md_to_word_from_text(content, out_path)
    if mode == "MD -> CSV":
        return md_to_csv_from_text(content, out_path)
    if mode in ("Excel -> MD", "Word -> MD", "CSV -> MD"):
        return save_markdown_from_text(content, out_path)
    raise ValueError(f"Invalid mode {mode}!")
