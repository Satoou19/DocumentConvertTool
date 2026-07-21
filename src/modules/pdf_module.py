import os
import sys
from src.core.base_module import BaseDocumentModule
from src.core.registry import ModuleRegistry

class PDFModule(BaseDocumentModule):
    @property
    def name(self) -> str:
        return "PDF"

    @property
    def file_extensions(self) -> list[str]:
        return [".pdf"]

    @property
    def required_dependencies(self) -> list[str]:
        return ["markitdown", "pdfplumber", "markdown-pdf"]

    def load_to_markdown(self, file_path: str) -> str:
        """
        Extracts PDF content to clean Markdown text with table structure recognition using pdfplumber.
        
        Supports table stitching across page breaks when tables are contiguous 
        (separated only by page breaks or empty space/headers) and merges continuation cells.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        try:
            # pyrefly: ignore [missing-import]
            import pdfplumber
            
            def merge_non_overlapping_columns(rows):
                if not rows or len(rows) < 2:
                    return rows
                num_cols = len(rows[0])
                cols = list(range(num_cols))
                changed = True
                while changed:
                    changed = False
                    for i in range(len(cols) - 1):
                        c1 = cols[i]
                        c2 = cols[i+1]
                        overlap = False
                        for row in rows:
                            if row[c1] and row[c2]:
                                overlap = True
                                break
                        if not overlap:
                            for row in rows:
                                if row[c2]:
                                    row[c1] = row[c2]
                            cols.pop(i+1)
                            changed = True
                            break
                rebuilt = []
                for row in rows:
                    rebuilt.append([row[c] for c in cols])
                return rebuilt

            def clean_table(table_data):
                if not table_data:
                    return []
                valid_rows = [r for r in table_data if r is not None]
                if not valid_rows:
                    return []
                max_cols = max(len(r) for r in valid_rows)
                cleaned = []
                for row in valid_rows:
                    cleaned_row = []
                    for cell in row:
                        val = str(cell).strip() if cell is not None else ""
                        cleaned_row.append(val)
                    if len(cleaned_row) < max_cols:
                        cleaned_row.extend([""] * (max_cols - len(cleaned_row)))
                    cleaned.append(cleaned_row[:max_cols])
                merged = merge_non_overlapping_columns(cleaned)
                if not merged:
                    return []
                num_cols = len(merged[0])
                active_cols = []
                for col_idx in range(num_cols):
                    has_val = False
                    for row in merged:
                        if row[col_idx]:
                            has_val = True
                            break
                    if has_val:
                        active_cols.append(col_idx)
                if not active_cols:
                    return []
                rebuilt = []
                for row in merged:
                    rebuilt.append([row[idx] for idx in active_cols])
                return rebuilt

            def map_row_to_parent_columns(t2_row, t2_cols, t1_cols):
                parent_row = [""] * len(t1_cols)
                for j, cell in enumerate(t2_row):
                    val = str(cell).strip() if cell is not None else ""
                    tx0 = t2_cols[j].bbox[0]
                    tx1 = t2_cols[j].bbox[2]
                    
                    best_idx = -1
                    max_overlap = -1
                    for idx, col in enumerate(t1_cols):
                        cx0 = col.bbox[0]
                        cx1 = col.bbox[2]
                        overlap = min(tx1, cx1) - max(tx0, cx0)
                        if overlap > max_overlap:
                            max_overlap = overlap
                            best_idx = idx
                    if best_idx != -1 and max_overlap > 0:
                        parent_row[best_idx] = val
                return parent_row

            def should_merge_first_row(last_row, first_row):
                if len(last_row) == len(first_row) and len(first_row) > 1:
                    if not first_row[0]:
                        # Case A: Last row of parent has empty cell where first row of child has content
                        for c in range(1, len(first_row)):
                            if not last_row[c] and first_row[c]:
                                return True
                        # Case B: Both have content in the last column but parent cell does not end with sentence punctuation
                        if last_row[-1] and first_row[-1]:
                            last_char = last_row[-1].strip()[-1] if last_row[-1].strip() else ""
                            if last_char not in (".", "!", "?", ":"):
                                return True
                return False

            def format_markdown_table(table_data):
                cleaned = clean_table(table_data)
                if not cleaned:
                    return ""
                headers = cleaned[0]
                rows = cleaned[1:]
                parts = []
                parts.append("| " + " | ".join(h.replace("\n", "<br>").replace("|", "\\|") for h in headers) + " |")
                parts.append("| " + " | ".join("---" for _ in headers) + " |")
                for r in rows:
                    parts.append("| " + " | ".join(cell.replace("\n", "<br>").replace("|", "\\|") for cell in r) + " |")
                return "\n".join(parts)

            def _extract_rich_text(cropped_page):
                chars = cropped_page.chars
                if not chars:
                    return ""
                
                # Sort characters by top coordinate first
                chars_sorted = sorted(chars, key=lambda c: (c["top"], c["x0"]))
                
                # Group characters into lines using a tolerance of 3 pixels
                lines = []
                current_line = []
                current_top = None
                
                for char in chars_sorted:
                    top = char["top"]
                    if current_top is None:
                        current_top = top
                        current_line.append(char)
                    elif abs(top - current_top) <= 3:
                        current_line.append(char)
                    else:
                        lines.append(current_line)
                        current_line = [char]
                        current_top = top
                if current_line:
                    lines.append(current_line)
                    
                line_texts = []
                for line in lines:
                    # Sort characters on the line from left to right
                    line.sort(key=lambda c: c["x0"])
                    
                    # Merge characters into runs that share the same font properties
                    runs = []
                    current_run_text = ""
                    current_font = None
                    total_size = 0
                    char_count = 0
                    
                    for char in line:
                        font = char.get("fontname", "").lower()
                        bold = "bold" in font
                        italic = "italic" in font or "oblique" in font
                        
                        size = char.get("size", 10)
                        total_size += size
                        char_count += 1
                        
                        style_key = (bold, italic)
                        if current_font is None:
                            current_font = style_key
                            current_run_text = char["text"]
                        elif style_key == current_font:
                            # Check distance to previous character to insert spaces
                            last_char = line[line.index(char) - 1]
                            gap = char["x0"] - last_char["x1"]
                            char_width = char["x1"] - char["x0"]
                            if gap > char_width * 0.25 and not current_run_text.endswith(" ") and char["text"] != " ":
                                current_run_text += " "
                            current_run_text += char["text"]
                        else:
                            runs.append((current_run_text, current_font))
                            current_font = style_key
                            current_run_text = char["text"]
                    if current_run_text:
                        runs.append((current_run_text, current_font))
                        
                    # Calculate average line size for heading detection
                    avg_size = total_size / char_count if char_count > 0 else 10
                    
                    # Format runs into markdown syntax
                    formatted_runs = []
                    for text, (bold, italic) in runs:
                        stripped = text.strip()
                        if not stripped:
                            formatted_runs.append(text)
                            continue
                            
                        leading_spaces = text[:len(text) - len(text.lstrip())]
                        trailing_spaces = text[len(text.rstrip()):]
                        
                        val = stripped
                        if bold and italic:
                            val = f"***{val}***"
                        elif bold:
                            val = f"**{val}**"
                        elif italic:
                            val = f"*{val}*"
                            
                        formatted_runs.append(f"{leading_spaces}{val}{trailing_spaces}")
                        
                    line_text = "".join(formatted_runs).strip()
                    
                    if line_text:
                        # Prepend heading tags if the line has larger text size
                        if avg_size >= 20:
                            line_text = f"# {line_text}"
                        elif 16 <= avg_size < 20:
                            line_text = f"## {line_text}"
                        elif 12.5 <= avg_size < 16:
                            line_text = f"### {line_text}"
                            
                    line_texts.append(line_text)
                    
                return "\n".join(line_texts)

            doc_elements = []
            settings = {"snap_tolerance": 10, "join_tolerance": 10}
            
            with pdfplumber.open(file_path) as pdf:
                for page_idx, page in enumerate(pdf.pages):
                    tables = sorted(page.find_tables(table_settings=settings), key=lambda t: t.bbox[1])
                    current_y = 0
                    
                    for t in tables:
                        bx0, btop, bx1, bbottom = t.bbox
                        if btop > current_y + 2:
                            cropped = page.crop((0, current_y, page.width, btop))
                            text_slice = _extract_rich_text(cropped)
                            if text_slice:
                                doc_elements.append({"type": "text", "content": text_slice})
                        
                        doc_elements.append({
                            "type": "table",
                            "content": t.extract(),
                            "bbox": t.bbox,
                            "columns": t.columns
                        })
                        current_y = bbottom
                        
                    if current_y < page.height:
                        cropped = page.crop((0, current_y, page.width, page.height))
                        text_slice = _extract_rich_text(cropped)
                        if text_slice:
                            doc_elements.append({"type": "text", "content": text_slice})
                            
                    # Add page break marker (except for the last page)
                    if page_idx < len(pdf.pages) - 1:
                        doc_elements.append({"type": "page_break", "content": "\n\n---\n\n"})

            # STITCHING PROCESS
            changed = True
            while changed:
                changed = False
                for idx in range(len(doc_elements) - 1):
                    el1 = doc_elements[idx]
                    if el1["type"] != "table":
                        continue
                        
                    next_table_idx = -1
                    has_page_break = False
                    accumulated_gap_text_len = 0
                    for j in range(idx + 1, len(doc_elements)):
                        el_next = doc_elements[j]
                        if el_next["type"] == "table":
                            next_table_idx = j
                            break
                        elif el_next["type"] == "page_break":
                            has_page_break = True
                            continue
                        elif el_next["type"] == "text":
                            text_content = el_next["content"].strip()
                            if text_content:
                                accumulated_gap_text_len += len(text_content)
                        else:
                            # Set an arbitrary high value to deliberately block stitching 
                            # if an unexpected or non-text element is found in the gap.
                            accumulated_gap_text_len += 9999
                            
                    if next_table_idx != -1 and has_page_break and accumulated_gap_text_len < 150:
                        el3 = doc_elements[next_table_idx]
                        t1_cols = el1["columns"]
                        t2_cols = el3["columns"]
                        t2_rows = el3["content"]
                        
                        mapped_rows = []
                        for row in t2_rows:
                            mapped = map_row_to_parent_columns(row, t2_cols, t1_cols)
                            mapped_rows.append(mapped)
                            
                        # Perform cell-continuation merge if needed
                        if el1["content"] and mapped_rows:
                            parent_last_row = el1["content"][-1]
                            child_first_row = mapped_rows[0]
                            if should_merge_first_row(parent_last_row, child_first_row):
                                for col_idx in range(len(parent_last_row)):
                                    val_T2 = child_first_row[col_idx]
                                    if val_T2:
                                        if parent_last_row[col_idx]:
                                            parent_last_row[col_idx] = parent_last_row[col_idx] + " " + val_T2
                                        else:
                                            parent_last_row[col_idx] = val_T2
                                mapped_rows.pop(0)
                        
                        el1["content"].extend(mapped_rows)
                        
                        # Remove elements between idx and next_table_idx
                        del doc_elements[idx + 1 : next_table_idx + 1]
                        changed = True
                        break

            # BUILD RENDERED OUTPUT
            output_parts = []
            for el in doc_elements:
                if el["type"] == "text":
                    output_parts.append(el["content"].strip())
                elif el["type"] == "table":
                    md_table = format_markdown_table(el["content"])
                    if md_table:
                        output_parts.append(md_table)
                elif el["type"] == "page_break":
                    output_parts.append(el["content"])
                    
            if not output_parts:
                return "*(Empty PDF)*"
            return "\n\n".join(output_parts)
            
        except Exception as e:
            # Print debug log before falling back
            print(f"[DEBUG] pdfplumber table extraction failed: {e}. Falling back to markitdown.", file=sys.stderr)
            try:
                # pyrefly: ignore [missing-import]
                from markitdown import MarkItDown
                md = MarkItDown()
                result = md.convert(file_path)
                if not result or not result.text_content:
                    return "*(Empty PDF)*"
                return result.text_content
            except Exception as inner_e:
                raise RuntimeError(f"PDF Ingestion Error: Failed to extract text layer from PDF file. Detail: {str(inner_e)}")

    def save_from_markdown(self, markdown_content: str, out_path: str) -> str:
        """Converts Markdown text to formatted PDF document using markdown-pdf."""
        import re
        try:
            # Pre-process Markdown: replace ~~text~~ with <del>text</del> for strikethrough support
            html_content = re.sub(r"~~(.*?)~~", r"<del>\1</del>", markdown_content)

            from markdown_pdf import MarkdownPdf, Section
            
            # CSS Stylesheet for professional layout matching the app's aesthetics
            css = """
            body {
                font-family: 'Segoe UI', Arial, sans-serif;
                line-height: 1.5;
                color: #333333;
            }
            h1, h2, h3 {
                color: #1a1a1a;
                margin-top: 20px;
                margin-bottom: 10px;
            }
            h1 { border-bottom: 1px solid #eaecef; padding-bottom: 5px; }
            table {
                border-collapse: collapse;
                width: 100%;
                margin-top: 15px;
                margin-bottom: 15px;
            }
            th, td {
                border: 1px solid #cccccc;
                padding: 8px 12px;
                text-align: left;
            }
            th {
                background-color: #f6f8fa;
                font-weight: bold;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            code {
                background-color: #f0f0f0;
                padding: 2px 6px;
                font-family: 'Consolas', monospace;
                font-size: 0.9em;
                border-radius: 3px;
            }
            del {
                text-decoration: line-through;
                color: #6a737d;
            }
            a {
                color: #0366d6;
                text-decoration: none;
            }
            a:hover {
                text-decoration: underline;
            }
            """
            
            pdf = MarkdownPdf(toc_level=2)
            pdf.add_section(Section(html_content), user_css=css)
            
            # Ensure output directory exists
            out_dir = os.path.dirname(out_path)
            if out_dir and not os.path.exists(out_dir):
                os.makedirs(out_dir, exist_ok=True)
                
            pdf.save(out_path)
            return f"Exported successfully to {os.path.basename(out_path)}"
        except Exception as e:
            raise RuntimeError(f"PDF Export Error: Failed to generate PDF document. Detail: {str(e)}")

ModuleRegistry.register(PDFModule())
