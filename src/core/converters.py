import os
import re

def parse_md_tables(content: str) -> list:
    import pandas as pd
    tables, lines, i = [], content.split("\n"), 0
    while i < len(lines):
        line = lines[i].strip()
        if "|" in line and not re.match(r"^[\|\s\-:]+$", line):
            table_name = f"Sheet{len(tables)+1}"
            for j in range(i-1, max(i-5, -1), -1):
                prev = lines[j].strip()
                if prev.startswith("#"):
                    table_name = re.sub(r"^#+\s*", "", prev)
                    table_name = re.sub(r'[\\/?*\[\]:]', "_", table_name)[:31]
                    break
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
            data_lines = [l for l in table_lines if not re.match(r"^[\|\s\-:]+$", l)]
            if len(data_lines) < 2:
                continue
            rows = [[c.strip() for c in l.split("|") if c.strip()] for l in data_lines]
            max_cols = max(len(r) for r in rows)
            rows = [r + [""] * (max_cols - len(r)) for r in rows]
            df = pd.DataFrame(rows[1:], columns=rows[0])
            tables.append((table_name, df))
        else:
            i += 1
    return tables


def save_markdown_from_text(content: str, out_path: str) -> str:
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)
    return f"Markdown file saved successfully -> {os.path.basename(out_path)}"
