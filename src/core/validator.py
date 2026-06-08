import re

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
