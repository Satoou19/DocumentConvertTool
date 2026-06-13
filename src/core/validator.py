import re
import os
import zipfile

def validate_md_tables(content: str) -> list[str]:
    """
    Validates Markdown tables in the content.
    Returns a list of warning strings. If empty, the tables are valid.
    """
    warnings = []
    lines = content.split("\n")
    i = 0
    table_index = 0
    
    while i < len(lines):
        line = lines[i].strip()
        if "|" in line and not re.match(r"^[\|\s\-:]+$", line):
            table_index += 1
            # Determine table name
            table_name = f"Table #{table_index}"
            for j in range(i-1, max(i-5, -1), -1):
                prev = lines[j].strip()
                if prev.startswith("#"):
                    table_name = re.sub(r"^#+\s*", "", prev)[:31]
                    break
            
            # Extract all table lines (consecutive lines containing '|')
            table_start_line = i + 1
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
                
            # 1. Check if table is too short
            if len(table_lines) < 2:
                warnings.append(
                    f"'{table_name}' starting at line {table_start_line}: "
                    "The table is incomplete. A valid table must have at least a header row and a separator line (e.g., '| Header |' followed by '|---|')."
                )
                continue
                
            # 2. Check for missing or malformed separator line at row index 1
            separator_line = table_lines[1]
            if not re.match(r"^[\|\s\-:]+$", separator_line):
                warnings.append(
                    f"'{table_name}' starting at line {table_start_line}: "
                    "Missing or incorrect separator line below the header. Please ensure the second line looks like '|---|---|'."
                )
                
            # Parse rows to check column counts
            # Filter out separator lines for column checking
            data_lines = [l for l in table_lines if not re.match(r"^[\|\s\-:]+$", l)]
            if len(data_lines) < 2:
                # No data rows (only header and separator)
                continue
                
            # Extract cells for each row
            rows = [[c.strip() for c in l.split("|") if c.strip()] for l in data_lines]
            header_col_count = len(rows[0])
            
            # Check column counts row by row
            for row_idx, r in enumerate(rows):
                if len(r) != header_col_count:
                    try:
                        line_num = table_start_line + table_lines.index(data_lines[row_idx])
                    except ValueError:
                        line_num = table_start_line
                    warnings.append(
                        f"'{table_name}' at line {line_num}: "
                        f"This row has {len(r)} columns but the header has {header_col_count}. Please check if you are missing a '|' separator."
                    )
        else:
            i += 1
            
    return warnings

def validate_file_integrity(path: str) -> tuple[str, str] | None:
    """
    Validates the integrity of the file.
    Returns a tuple of (short_reason, detailed_reason) if an issue is detected, or None if valid.
    """
    if not os.path.exists(path):
        return ("File does not exist", "The selected file does not exist. Please check the file path.")

    ext = os.path.splitext(path)[1].lower()
    if ext not in (".md", ".docx", ".xlsx", ".xls"):
        return (
            "Unsupported file extension",
            f"The file extension '{ext}' is not supported. Please select a valid document file (.md, .docx, .xlsx, .xls)."
        )

    # 1. Check if the file is locked / open in another application
    try:
        with open(path, "rb") as f:
            first_bytes = f.read(8)
    except PermissionError:
        return ("File is locked or open in another app", "The file is currently locked or open in another application (e.g., Microsoft Word or Excel). Please close it and try again.")
    except Exception as e:
        return ("File access denied", f"Unable to access the file: {e}")

    # 2. Check if the file is empty (0 bytes)
    try:
        if os.path.getsize(path) == 0:
            return ("File is empty (0 bytes)", "The selected file is empty (0 bytes). Please select a file with content.")
    except Exception as e:
        return ("Unable to check file size", f"Unable to check file size: {e}")

    ext = os.path.splitext(path)[1].lower()

    # 3. Check for specific formats
    if ext == ".md":
        # Check if it is a binary file instead of text
        try:
            with open(path, "r", encoding="utf-8") as f:
                chunk = f.read(4096)
                if "\x00" in chunk:
                    return ("Fake extension: binary file", "Fake file extension detected. The file content is binary, but it has a '.md' extension.")
        except UnicodeDecodeError:
            return ("Fake extension: not UTF-8 text", "Fake file extension detected. The file content is not valid UTF-8 text, but it has a '.md' extension.")
        except Exception as e:
            return ("Error reading file", f"Error reading text file: {e}")

    elif ext in (".docx", ".xlsx"):
        # Check if it is a valid zip archive
        is_zip = False
        try:
            is_zip = zipfile.is_zipfile(path)
        except Exception:
            pass

        if not is_zip:
            # Check if it's password protected (encrypted OOXML files are OLE compound documents)
            if first_bytes == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                return ("File is password-protected or legacy format", "The file appears to be password-protected (encrypted) or in an incompatible legacy format (e.g. old .xls/.doc format renamed to .xlsx/.docx). Please decrypt it and ensure it is in the correct modern format.")
            elif first_bytes.startswith(b"PK\x03\x04"):
                return ("Corrupted ZIP structure", "The file appears to be a corrupted or incomplete Word/Excel archive (corrupted ZIP structure).")
            else:
                return ("Fake extension: invalid format", f"Fake file extension detected. The file content does not match its '{ext}' extension.")

        # If it is a zip archive, check for corruption and check if it has OOXML structure
        try:
            with zipfile.ZipFile(path) as zf:
                # Test zip structure integrity
                bad_file = zf.testzip()
                if bad_file is not None:
                    return ("File is corrupted", f"The file appears to be corrupted. Incomplete zip structure at: {bad_file}")
                
                # Check for fake extension by verifying OOXML folder/file structure
                namelist = zf.namelist()
                if ext == ".docx" and "word/document.xml" not in namelist:
                    return ("Fake extension: missing Word structure", "Fake file extension detected. The file is a ZIP archive but does not contain a valid Word document structure.")
                if ext == ".xlsx" and not any(name.startswith("xl/") for name in namelist):
                    return ("Fake extension: missing Excel structure", "Fake file extension detected. The file is a ZIP archive but does not contain a valid Excel workbook structure.")
        except zipfile.BadZipFile:
            return ("Invalid zip archive", "The file is corrupted or is not a valid zip archive.")
        except Exception as e:
            return ("Corrupted or invalid file", f"The file appears to be corrupted or invalid: {e}")

    elif ext == ".xls":
        # Legacy Excel file
        if first_bytes != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            return ("Fake extension: not legacy Excel", "Fake file extension detected. The file content does not match the legacy Excel (.xls) format.")

    return None
