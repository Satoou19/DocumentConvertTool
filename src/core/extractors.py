import re
import zipfile

def extract_excel_to_md(in_path: str) -> str:
    """Extracts Excel sheets into clean Markdown tables."""
    import pandas as pd
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
    """Extracts Word .docx to clean Markdown text using Mammoth.
    Images are stripped from output to prevent massive base64 strings from bloating the content."""
    import mammoth
    
    # Validate DOCX file trước (DOCX là ZIP archive)
    try:
        if not zipfile.is_zipfile(in_path):
            return "⚠️ File Validation Error: Invalid file. Please select a valid DOCX document."
    except Exception as e:
        return f"⚠️ File Validation Error: {e}"
    
    try:
        with open(in_path, "rb") as f:
            # Extract markdown từ DOCX
            result = mammoth.convert_to_markdown(f)
        
        # Kiểm tra kết quả từ Mammoth
        if not result or not result.value:
            return "⚠️ Extract Error: Mammoth returned empty result. File may be corrupted or has no text content."
        
        markdown = result.value
        
        # Kiểm tra encoding - ensure it's valid UTF-8 string
        if isinstance(markdown, bytes):
            try:
                markdown = markdown.decode('utf-8', errors='replace')
            except Exception as e:
                return f"⚠️ Encoding Error: {e}"
        
        # Nếu vẫn là binary/garbled, báo lỗi
        if not isinstance(markdown, str):
            return f"⚠️ Extract Error: Unexpected data type from Mammoth: {type(markdown)}"
        
        # Replace base64 images với placeholder thân thiện
        # ![alt text](data:image/png;base64,...) → [📷 Image 1: alt text] hoặc [📷 Image 1]
        image_counter = [1]  # Dùng list để modify trong nested function
        
        def replace_image(match):
            alt_text = match.group(1).strip()
            if alt_text:
                placeholder = f"[📷 Image {image_counter[0]}: {alt_text}]"
            else:
                placeholder = f"[📷 Image {image_counter[0]}]"
            image_counter[0] += 1
            return placeholder
        
        markdown = re.sub(r'!\[([^\]]*)\]\(data:image/[^)]*\)', replace_image, markdown)
        
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
    
    except Exception as e:
        # Log chi tiết lỗi
        error_msg = f"⚠️ Word Extraction Error:\n\nDetails: {str(e)}\n\nFile: {in_path}"
        print(f"[ERROR] extract_word_to_md: {error_msg}")
        return error_msg
