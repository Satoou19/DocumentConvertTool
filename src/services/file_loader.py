import os
from dataclasses import dataclass
from typing import List, Optional

from src.core.extractors import extract_excel_to_md, extract_word_to_md
from src.core.validator import validate_file_integrity

try:
    import pandas  # noqa: F401
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import openpyxl  # noqa: F401
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import docx  # noqa: F401
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


@dataclass
class LoadResult:
    success: bool
    content: str = ""
    mode: str = ""
    error_short: Optional[str] = None
    error_detail: Optional[str] = None
    missing_dependencies: Optional[List[str]] = None


def get_missing_dependencies_for_path(path: str) -> List[str]:
    ext = os.path.splitext(path)[1].lower()
    missing = []
    if ext in (".xlsx", ".xls"):
        if not HAS_PANDAS:
            missing.append("pandas")
        if not HAS_OPENPYXL:
            missing.append("openpyxl")
    elif ext == ".docx":
        if not HAS_DOCX:
            missing.append("python-docx")
    return missing


def load_document(path: str) -> LoadResult:
    integrity_error = validate_file_integrity(path)
    if integrity_error:
        short, detail = integrity_error
        return LoadResult(False, error_short=short, error_detail=detail)

    missing = get_missing_dependencies_for_path(path)
    if missing:
        missing_msg = " and ".join(missing)
        return LoadResult(
            False,
            error_short=f"{missing_msg} missing",
            error_detail="",
            missing_dependencies=missing,
        )

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".md":
            mode = "MD -> Excel"
            with open(path, encoding="utf-8") as f:
                content = f.read()
        elif ext in (".xlsx", ".xls"):
            mode = "Excel -> MD"
            content = extract_excel_to_md(path)
        elif ext == ".docx":
            mode = "Word -> MD"
            content = extract_word_to_md(path)
        else:
            return LoadResult(False, error_short="Unsupported extension", error_detail=f"File extension {ext} is not supported.")
    except Exception as exc:
        return LoadResult(False, error_short="Load error", error_detail=str(exc))

    return LoadResult(True, content=content, mode=mode)
