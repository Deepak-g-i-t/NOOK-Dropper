"""
Converter Engine for Nook Drop Changer Plugin.
Handles conversion across Text, Images, PDF, Word (.docx), Excel (.xlsx), CSV, and HTML.
Supports Windows clipboard integration and drag-and-dropped files.
"""

import os
import sys
import re
import csv
import json
import time
import ctypes
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from io import BytesIO

# Third-party conversion libraries
from PIL import Image, ImageDraw, ImageFont, ImageOps
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import reportlab
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

import pypdf


def get_default_output_dir():
    """Returns the default directory where converted files are saved."""
    # 1. If running under Nook, use NOOK_DATA_DIR/converted
    nook_data = os.environ.get("NOOK_DATA_DIR")
    if nook_data:
        dest = Path(nook_data) / "converted"
        try:
            dest.mkdir(parents=True, exist_ok=True)
            return str(dest)
        except Exception:
            pass

    # 2. Try User Downloads folder
    try:
        dest = Path.home() / "Downloads" / "NookConverted"
        dest.mkdir(parents=True, exist_ok=True)
        return str(dest)
    except Exception:
        pass

    # 3. Fallback to temp directory
    import tempfile
    dest = Path(tempfile.gettempdir()) / "NookConverted"
    dest.mkdir(parents=True, exist_ok=True)
    return str(dest)


def safe_filename(name, max_len=40):
    """Sanitizes a string to make a clean filename."""
    cleaned = re.sub(r'[\\/*?:"<>| \r\n\t]+', '_', name.strip())
    cleaned = cleaned.strip('._')
    return cleaned[:max_len] if cleaned else "converted"


def get_unique_output_path(base_name, extension, output_dir=None, include_timestamp=False):
    """Generates an unused output filepath to avoid overwriting existing files."""
    if output_dir is None:
        output_dir = get_default_output_dir()
    os.makedirs(output_dir, exist_ok=True)

    extension = extension.lstrip('.')
    clean_base = safe_filename(base_name)
    if not clean_base:
        clean_base = "converted"

    if include_timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        candidate_name = f"{clean_base}_{timestamp}.{extension}"
        return os.path.join(output_dir, candidate_name)

    candidate_name = f"{clean_base}.{extension}"
    candidate_path = os.path.join(output_dir, candidate_name)
    if not os.path.exists(candidate_path):
        return candidate_path

    counter = 1
    while os.path.exists(candidate_path):
        candidate_path = os.path.join(output_dir, f"{clean_base} ({counter}).{extension}")
        counter += 1
    return candidate_path


# ---------------------------------------------------------------------------
# Windows Clipboard Inspector
# ---------------------------------------------------------------------------

def read_windows_clipboard():
    """
    Inspects Windows clipboard for:
    1. Image (e.g. copied from browser or snipping tool)
    2. Dropped/copied files from Windows Explorer
    3. Copied text / web page selections
    Returns dict: {"kind": "image"|"files"|"text", "data": ..., "preview": str} or None.
    """
    # 1. Try PIL ImageGrab (checks image or file drop list)
    try:
        from PIL import ImageGrab
        clip = ImageGrab.grabclipboard()
        if isinstance(clip, Image.Image):
            w, h = clip.size
            return {
                "kind": "image",
                "data": clip,
                "preview": f"Clipboard Image ({w}x{h} {clip.mode})",
                "source": "Clipboard",
                "name": f"clipboard_image_{w}x{h}"
            }
        elif isinstance(clip, list) and clip and all(isinstance(p, str) for p in clip):
            existing_files = [p for p in clip if os.path.exists(p)]
            if existing_files:
                first_name = os.path.basename(existing_files[0])
                count = len(existing_files)
                preview = f"{first_name}" if count == 1 else f"{first_name} (+{count-1} more)"
                return {
                    "kind": "files",
                    "data": existing_files,
                    "preview": preview,
                    "source": "Clipboard Files",
                    "name": Path(existing_files[0]).stem
                }
    except Exception:
        pass

    # 2. Try Windows Unicode Text via user32
    try:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        if user32.OpenClipboard(None):
            try:
                CF_UNICODETEXT = 13
                handle = user32.GetClipboardData(CF_UNICODETEXT)
                if handle:
                    kernel32.GlobalLock.restype = ctypes.c_wchar_p
                    ptr = kernel32.GlobalLock(handle)
                    if ptr:
                        text = str(ptr)
                        kernel32.GlobalUnlock(handle)
                        text_stripped = text.strip()
                        if text_stripped:
                            first_line = text_stripped.splitlines()[0][:50]
                            char_count = len(text_stripped)
                            return {
                                "kind": "text",
                                "data": text_stripped,
                                "preview": f'"{first_line}..." ({char_count} chars)',
                                "source": "Clipboard Text",
                                "name": safe_filename(first_line, 24) or "clipboard_text"
                            }
            finally:
                user32.CloseClipboard()
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Input Analysis & Detection
# ---------------------------------------------------------------------------

IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif', '.tiff', '.tif', '.ico'}
SPREADSHEET_EXTENSIONS = {'.xlsx', '.xls', '.csv', '.tsv'}
DOC_EXTENSIONS = {'.docx', '.doc'}
PDF_EXTENSIONS = {'.pdf'}
TEXT_EXTENSIONS = {'.txt', '.md', '.json', '.xml', '.html', '.htm', '.py', '.js', '.css', '.c', '.cpp', '.rs', '.java', '.log', '.sql', '.yaml', '.yml'}


def inspect_file(filepath):
    """Inspects a file path and returns type metadata."""
    p = Path(filepath)
    if not p.exists():
        return None
    ext = p.suffix.lower()
    size_bytes = p.stat().st_size
    size_str = format_file_size(size_bytes)

    if ext in IMAGE_EXTENSIONS:
        try:
            with Image.open(filepath) as img:
                w, h = img.size
                format_name = img.format or ext[1:].upper()
            return {
                "kind": "image",
                "path": str(p),
                "name": p.stem,
                "ext": ext,
                "preview": f"{p.name} ({w}x{h} {format_name}, {size_str})",
                "source": "File",
                "size_str": size_str,
                "image_size": (w, h)
            }
        except Exception:
            pass

    if ext in SPREADSHEET_EXTENSIONS:
        return {
            "kind": "spreadsheet",
            "path": str(p),
            "name": p.stem,
            "ext": ext,
            "preview": f"{p.name} (Spreadsheet, {size_str})",
            "source": "File",
            "size_str": size_str
        }

    if ext in PDF_EXTENSIONS:
        try:
            reader = pypdf.PdfReader(filepath)
            num_pages = len(reader.pages)
            return {
                "kind": "pdf",
                "path": str(p),
                "name": p.stem,
                "ext": ext,
                "preview": f"{p.name} ({num_pages} {'page' if num_pages == 1 else 'pages'}, {size_str})",
                "source": "File",
                "size_str": size_str,
                "pages": num_pages
            }
        except Exception:
            return {
                "kind": "pdf",
                "path": str(p),
                "name": p.stem,
                "ext": ext,
                "preview": f"{p.name} (PDF, {size_str})",
                "source": "File",
                "size_str": size_str
            }

    if ext in DOC_EXTENSIONS:
        return {
            "kind": "word",
            "path": str(p),
            "name": p.stem,
            "ext": ext,
            "preview": f"{p.name} (Word Document, {size_str})",
            "source": "File",
            "size_str": size_str
        }

    # Text / Code fallback
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            sample = f.read(200).strip()
        first_line = sample.splitlines()[0][:40] if sample else ""
        return {
            "kind": "text_file",
            "path": str(p),
            "name": p.stem,
            "ext": ext,
            "preview": f"{p.name} ({size_str})",
            "source": "File",
            "size_str": size_str,
            "sample": first_line
        }
    except Exception:
        return {
            "kind": "unknown",
            "path": str(p),
            "name": p.stem,
            "ext": ext,
            "preview": f"{p.name} ({size_str})",
            "source": "File",
            "size_str": size_str
        }


def format_file_size(size_bytes):
    """Formats file size into human-readable string."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}" if unit != 'B' else f"{size_bytes} B"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"


def get_available_target_formats(input_info):
    """Returns the list of target format chips/options for a given input."""
    kind = input_info.get("kind")

    if kind in ("text", "text_file"):
        return [
            {"id": "docx", "label": "Word (.docx)", "icon": "file-doc"},
            {"id": "pdf", "label": "PDF (.pdf)", "icon": "file-pdf"},
            {"id": "xlsx", "label": "Excel (.xlsx)", "icon": "file-xls"},
            {"id": "txt", "label": "TXT (.txt)", "icon": "file-text"},
            {"id": "md", "label": "Markdown (.md)", "icon": "file-text"},
            {"id": "png", "label": "Image Card (.png)", "icon": "file-image"},
            {"id": "html", "label": "HTML (.html)", "icon": "code"},
        ]

    elif kind == "image":
        return [
            {"id": "jpeg", "label": "JPEG (.jpg)", "icon": "file-image"},
            {"id": "png", "label": "PNG (.png)", "icon": "file-image"},
            {"id": "webp", "label": "WEBP (.webp)", "icon": "file-image"},
            {"id": "pdf", "label": "PDF (.pdf)", "icon": "file-pdf"},
            {"id": "docx", "label": "Word (.docx)", "icon": "file-doc"},
            {"id": "ico", "label": "Icon (.ico)", "icon": "app-window"},
        ]

    elif kind == "spreadsheet":
        return [
            {"id": "xlsx", "label": "Excel (.xlsx)", "icon": "file-xls"},
            {"id": "pdf", "label": "PDF Report (.pdf)", "icon": "file-pdf"},
            {"id": "docx", "label": "Word Table (.docx)", "icon": "file-doc"},
            {"id": "csv", "label": "CSV (.csv)", "icon": "table"},
            {"id": "txt", "label": "TXT (.txt)", "icon": "file-text"},
            {"id": "html", "label": "HTML Table (.html)", "icon": "code"},
        ]

    elif kind == "pdf":
        return [
            {"id": "txt", "label": "Text (.txt)", "icon": "file-text"},
            {"id": "docx", "label": "Word (.docx)", "icon": "file-doc"},
        ]

    elif kind == "word":
        return [
            {"id": "pdf", "label": "PDF (.pdf)", "icon": "file-pdf"},
            {"id": "txt", "label": "Text (.txt)", "icon": "file-text"},
            {"id": "html", "label": "HTML (.html)", "icon": "code"},
        ]

    elif kind == "url":
        return [
            {"id": "url", "label": "Internet Shortcut (.url)", "icon": "globe"},
            {"id": "txt", "label": "Text Link (.txt)", "icon": "file-text"},
            {"id": "pdf", "label": "Link Card (.pdf)", "icon": "file-pdf"},
            {"id": "html", "label": "HTML Bookmark (.html)", "icon": "code"},
        ]

    elif kind == "html":
        return [
            {"id": "pdf", "label": "PDF (.pdf)", "icon": "file-pdf"},
            {"id": "docx", "label": "Word (.docx)", "icon": "file-doc"},
            {"id": "txt", "label": "Clean Text (.txt)", "icon": "file-text"},
            {"id": "html", "label": "HTML (.html)", "icon": "code"},
        ]

    return [
        {"id": "txt", "label": "TXT (.txt)", "icon": "file-text"}
    ]


# ---------------------------------------------------------------------------
# Tabular Text Parser (for Smart Excel/CSV/Table Conversion)
# ---------------------------------------------------------------------------

def parse_text_as_table(text):
    """
    Attempts to intelligently parse arbitrary text into a 2D table (list of rows).
    Supports:
    1. JSON array of objects or list of lists
    2. CSV or TSV (e.g. copied from Google Sheets / Excel / web table)
    3. Markdown table (| col1 | col2 |)
    4. Semicolon or comma delimited lines
    5. Plain lines with line numbering fallback
    """
    cleaned = text.strip()
    if not cleaned:
        return [["Content"], ["(empty)"]]

    # 1. Try JSON
    if (cleaned.startswith("[") and cleaned.endswith("]")) or (cleaned.startswith("{") and cleaned.endswith("}")):
        try:
            data = json.loads(cleaned)
            if isinstance(data, list) and data and isinstance(data[0], dict):
                headers = list(data[0].keys())
                rows = [headers]
                for item in data:
                    rows.append([str(item.get(h, "")) for h in headers])
                return rows
            elif isinstance(data, list) and data and isinstance(data[0], list):
                return [[str(c) for c in r] for r in data]
            elif isinstance(data, dict):
                return [["Key", "Value"]] + [[str(k), str(v)] for k, v in data.items()]
        except Exception:
            pass

    # 2. Check for Markdown table (| Col 1 | Col 2 |)
    lines = [ln.strip() for ln in cleaned.splitlines() if ln.strip()]
    if len(lines) >= 2 and all(ln.startswith("|") and ln.endswith("|") for ln in lines[:2]):
        rows = []
        for line in lines:
            if re.match(r"^\|(\s*[-:]+\s*\|)+$", line):
                continue  # skip separator line
            cols = [col.strip() for col in line.strip("|").split("|")]
            rows.append(cols)
        if rows:
            return rows

    # 3. Try CSV / TSV sniff
    try:
        sample = "\n".join(lines[:20])
        dialect = csv.Sniffer().sniff(sample, delimiters=['\t', ',', ';', '|'])
        reader = csv.reader(lines, dialect)
        parsed = list(reader)
        if parsed and len(parsed[0]) > 1:
            return parsed
    except Exception:
        pass

    # 4. Check for Tab-separated data (very common from clipboard copy from web tables / Excel)
    if any("\t" in ln for ln in lines):
        rows = [ln.split("\t") for ln in lines]
        return rows

    # 5. Check for Comma-separated
    if any("," in ln for ln in lines):
        try:
            reader = csv.reader(lines, delimiter=",")
            parsed = list(reader)
            if parsed and len(parsed[0]) > 1:
                return parsed
        except Exception:
            pass

    # Fallback: Plain text rows with line numbering
    rows = [["Line", "Content"]]
    for idx, ln in enumerate(lines, 1):
        rows.append([str(idx), ln])
    return rows


# ---------------------------------------------------------------------------
# Converters: Text to Targets
# ---------------------------------------------------------------------------

def text_to_docx(text, base_name="document", output_dir=None):
    """Converts text into a styled Microsoft Word (.docx) document."""
    out_path = get_unique_output_path(base_name, "docx", output_dir)
    doc = Document()

    # Document Title
    title = doc.add_heading(level=0)
    title_run = title.add_run(base_name.replace("_", " ").title())
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # Markdown headings
        if stripped.startswith("### "):
            h = doc.add_heading(stripped[4:], level=3)
            h.paragraph_format.space_before = Pt(8)
            h.paragraph_format.space_after = Pt(4)
        elif stripped.startswith("## "):
            h = doc.add_heading(stripped[3:], level=2)
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(4)
        elif stripped.startswith("# "):
            h = doc.add_heading(stripped[2:], level=1)
            h.paragraph_format.space_before = Pt(16)
            h.paragraph_format.space_after = Pt(6)
        # Bullet list
        elif stripped.startswith("- ") or stripped.startswith("* "):
            p = doc.add_paragraph(stripped[2:], style='List Bullet')
            p.paragraph_format.space_after = Pt(2)
        # Numbered list
        elif re.match(r"^\d+\.\s", stripped):
            item_text = re.sub(r"^\d+\.\s", "", stripped)
            p = doc.add_paragraph(item_text, style='List Number')
            p.paragraph_format.space_after = Pt(2)
        # Regular paragraph
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(line)
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(51, 65, 85)

        i += 1

    # Footer note
    footer = doc.sections[0].footer
    f_p = footer.paragraphs[0]
    f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_run = f_p.add_run("Generated by Nook Drop Changer • " + datetime.now().strftime("%b %d, %Y"))
    f_run.font.size = Pt(8.5)
    f_run.font.color.rgb = RGBColor(148, 163, 184)

    doc.save(out_path)
    return out_path


def text_to_pdf(text, base_name="document", output_dir=None):
    """Converts text into a clean, modern PDF document."""
    out_path = get_unique_output_path(base_name, "pdf", output_dir)

    pdf = SimpleDocTemplate(
        out_path,
        pagesize=letter,
        rightMargin=48,
        leftMargin=48,
        topMargin=48,
        bottomMargin=48
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'NookTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=14
    )

    h2_style = ParagraphStyle(
        'NookH2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=colors.HexColor('#2563EB'),
        spaceBefore=12,
        spaceAfter=6
    )

    body_style = ParagraphStyle(
        'NookBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10.5,
        leading=15,
        textColor=colors.HexColor('#334155'),
        spaceAfter=6
    )

    bullet_style = ParagraphStyle(
        'NookBullet',
        parent=body_style,
        leftIndent=16,
        firstLineIndent=-10,
        spaceAfter=3
    )

    code_style = ParagraphStyle(
        'NookCode',
        parent=styles['Code'],
        fontName='Courier',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#0F172A'),
        backColor=colors.HexColor('#F1F5F9'),
        borderPadding=6,
        spaceBefore=4,
        spaceAfter=6
    )

    elements = []
    # Title
    elements.append(Paragraph(base_name.replace("_", " ").title(), title_style))
    elements.append(Spacer(1, 4))

    def escape_xml(s):
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            elements.append(Spacer(1, 4))
            continue

        if stripped.startswith("# "):
            elements.append(Paragraph(escape_xml(stripped[2:]), title_style))
        elif stripped.startswith("## ") or stripped.startswith("### "):
            header_text = stripped.lstrip("#").strip()
            elements.append(Paragraph(escape_xml(header_text), h2_style))
        elif stripped.startswith("- ") or stripped.startswith("* "):
            elements.append(Paragraph(f"&bull; {escape_xml(stripped[2:])}", bullet_style))
        elif re.match(r"^\d+\.\s", stripped):
            elements.append(Paragraph(escape_xml(stripped), bullet_style))
        elif line.startswith("    ") or line.startswith("\t"):
            elements.append(Paragraph(escape_xml(line), code_style))
        else:
            elements.append(Paragraph(escape_xml(line), body_style))

    pdf.build(elements)
    return out_path


def text_to_excel(text, base_name="data", output_dir=None):
    """Converts text (structured CSV/TSV/JSON/Markdown or plain lines) into an Excel workbook."""
    out_path = get_unique_output_path(base_name, "xlsx", output_dir)
    table_data = parse_text_as_table(text)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = safe_filename(base_name, 25)
    ws.views.sheetView[0].showGridLines = True

    # Styling definitions
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10, color="1E293B")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    for row_idx, row in enumerate(table_data, 1):
        for col_idx, cell_value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val_str = str(cell_value)
            
            # Numeric conversion if feasible
            if row_idx > 1 and val_str.replace('.', '', 1).isdigit():
                try:
                    cell.value = float(val_str) if '.' in val_str else int(val_str)
                except ValueError:
                    cell.value = val_str
            else:
                cell.value = val_str

            cell.border = thin_border
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = data_font
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if row_idx % 2 == 0:
                    cell.fill = alt_fill

    # Set row heights
    ws.row_dimensions[1].height = 26
    for r in range(2, len(table_data) + 1):
        ws.row_dimensions[r].height = 20

    # Auto-fit column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_len = len(str(cell.value or ''))
            if val_len > max_len:
                max_len = val_len
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 50)

    wb.save(out_path)
    return out_path


def text_to_image_card(text, base_name="snippet", output_dir=None):
    """Renders text into a beautiful social/developer card image."""
    out_path = get_unique_output_path(base_name, "png", output_dir)

    lines = text.strip().splitlines()[:30]  # cap lines for image card
    if not lines:
        lines = ["(Empty snippet)"]

    # Calculate dimensions
    max_line_len = max(len(ln) for ln in lines)
    img_width = max(min(max_line_len * 11 + 80, 1200), 600)
    img_height = max(len(lines) * 24 + 110, 240)

    img = Image.new("RGB", (img_width, img_height), color=(15, 23, 42))  # slate-900
    draw = ImageDraw.Draw(img)

    # Window controls (macOS/Notch style dots)
    dot_y = 24
    draw.ellipse((26, dot_y, 38, dot_y + 12), fill=(239, 68, 68))   # red
    draw.ellipse((46, dot_y, 58, dot_y + 12), fill=(245, 158, 11))  # yellow
    draw.ellipse((66, dot_y, 78, dot_y + 12), fill=(16, 185, 129))  # green

    # Card title
    title_text = base_name.replace("_", " ").title()
    draw.text((img_width // 2, dot_y + 6), title_text, fill=(148, 163, 184), anchor="mm")

    # Separator
    draw.line((20, 52, img_width - 20, 52), fill=(30, 41, 59), width=1)

    # Draw code/text lines
    y = 70
    for line in lines:
        draw.text((32, y), line, fill=(241, 245, 249))
        y += 24

    # Footer
    draw.text((img_width - 32, img_height - 20), "Nook Drop Changer", fill=(71, 85, 105), anchor="rm")

    img.save(out_path, "PNG")
    return out_path


def text_to_txt(text, base_name="text", output_dir=None):
    """Saves text as a clean UTF-8 text file."""
    out_path = get_unique_output_path(base_name, "txt", output_dir)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    return out_path


def text_to_html(text, base_name="document", output_dir=None):
    """Converts text into an elegant responsive HTML document."""
    out_path = get_unique_output_path(base_name, "html", output_dir)
    import html as html_lib

    title = base_name.replace("_", " ").title()
    escaped = html_lib.escape(text)

    # Wrap paragraphs
    paragraphs = []
    for line in text.splitlines():
        st = line.strip()
        if not st:
            continue
        if st.startswith("# "):
            paragraphs.append(f"<h1>{html_lib.escape(st[2:])}</h1>")
        elif st.startswith("## "):
            paragraphs.append(f"<h2>{html_lib.escape(st[3:])}</h2>")
        elif st.startswith("- "):
            paragraphs.append(f"<li>{html_lib.escape(st[2:])}</li>")
        else:
            paragraphs.append(f"<p>{html_lib.escape(line)}</p>")

    content_html = "\n".join(paragraphs)

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background-color: #0F172A;
    color: #E2E8F0;
    margin: 0;
    padding: 40px 20px;
    display: flex;
    justify-content: center;
  }}
  .container {{
    max-width: 760px;
    width: 100%;
    background: #1E293B;
    border-radius: 12px;
    padding: 36px 44px;
    box-shadow: 0 10px 25px -5px rgba(0,0,0,0.5);
    border: 1px solid #334155;
  }}
  h1 {{ color: #60A5FA; font-size: 26px; margin-top: 0; border-bottom: 1px solid #334155; padding-bottom: 12px; }}
  h2 {{ color: #93C5FD; font-size: 20px; margin-top: 24px; }}
  p {{ line-height: 1.65; color: #CBD5E1; margin: 10px 0; }}
  li {{ margin: 6px 0; color: #CBD5E1; }}
  .footer {{ margin-top: 36px; padding-top: 14px; border-top: 1px solid #334155; font-size: 12px; color: #64748B; text-align: right; }}
</style>
</head>
<body>
  <div class="container">
    <h1>{title}</h1>
    {content_html}
    <div class="footer">Converted with Nook Drop Changer • {datetime.now().strftime('%b %d, %Y')}</div>
  </div>
</body>
</html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return out_path


# ---------------------------------------------------------------------------
# Converters: Image to Targets
# ---------------------------------------------------------------------------

def image_to_format(img_or_path, target_ext, base_name="image", output_dir=None):
    """
    Converts an image (PIL Image or file path) to JPEG, PNG, WEBP, ICO, or PDF.
    """
    target_ext = target_ext.lower().lstrip('.')
    out_path = get_unique_output_path(base_name, target_ext, output_dir)

    if isinstance(img_or_path, (str, Path)):
        img = Image.open(img_or_path)
    else:
        img = img_or_path

    # Convert modes based on target
    if target_ext in ('jpg', 'jpeg'):
        # Convert RGBA or Palette to RGB with clean white background
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            bg.paste(img, mask=img.split()[3])
            bg.save(out_path, "JPEG", quality=92, optimize=True)
        else:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(out_path, "JPEG", quality=92, optimize=True)

    elif target_ext == 'png':
        img.save(out_path, "PNG", optimize=True)

    elif target_ext == 'webp':
        img.save(out_path, "WEBP", quality=90)

    elif target_ext == 'ico':
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
        img.save(out_path, format='ICO', sizes=sizes)

    elif target_ext == 'pdf':
        # Single or multi-page image to PDF
        if img.mode != 'RGB':
            rgb_img = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == 'RGBA':
                rgb_img.paste(img, mask=img.split()[3])
            else:
                rgb_img.paste(img.convert('RGB'))
            rgb_img.save(out_path, "PDF", resolution=100.0)
        else:
            img.save(out_path, "PDF", resolution=100.0)

    elif target_ext == 'docx':
        doc = Document()
        doc.add_heading(base_name.replace("_", " ").title(), level=1)

        # Save temporary png to insert
        target_dir = output_dir or get_default_output_dir()
        temp_img_path = os.path.join(target_dir, f"temp_{int(time.time()*1000)}.png")
        img.save(temp_img_path, "PNG")

        # Proportional resize to max 6.0 inches width
        w, h = img.size
        aspect = h / max(w, 1)
        target_w = min(6.0, w / 96.0)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(temp_img_path, width=Inches(target_w))

        # Caption
        cap = doc.add_paragraph(f"{base_name} • {w}x{h} px • {datetime.now().strftime('%b %d, %Y')}")
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.style = 'Caption'

        doc.save(out_path)
        try:
            os.remove(temp_img_path)
        except OSError:
            pass

    return out_path


# ---------------------------------------------------------------------------
# Converters: Spreadsheet to Targets (XLSX, CSV, PDF, Word, HTML, TXT)
# ---------------------------------------------------------------------------

def read_spreadsheet_to_table(filepath):
    """Reads .xlsx, .csv, or .tsv file into a 2D list of strings."""
    ext = Path(filepath).suffix.lower()
    if ext == '.xlsx':
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        table = []
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                table.append([str(c) if c is not None else "" for c in row])
        return table
    else:  # CSV or TSV
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            sample = f.read(2048)
            f.seek(0)
            delimiter = '\t' if ext == '.tsv' else ','
            try:
                dialect = csv.Sniffer().sniff(sample)
                reader = csv.reader(f, dialect)
            except Exception:
                reader = csv.reader(f, delimiter=delimiter)
            return list(reader)


def spreadsheet_to_format(filepath, target_format, output_dir=None, custom_name=None):
    """Converts a spreadsheet (.xlsx/.csv) to PDF, Word, CSV, Excel, HTML, or TXT."""
    table_data = read_spreadsheet_to_table(filepath)
    base_name = custom_name or Path(filepath).stem

    if target_format == 'pdf':
        out_path = get_unique_output_path(base_name, "pdf", output_dir)
        pdf = SimpleDocTemplate(out_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
        elements = [Paragraph(f"<b>{base_name}</b>", getSampleStyleSheet()['Title']), Spacer(1, 12)]

        # Truncate row columns if too wide for page
        max_cols = min(len(table_data[0]) if table_data else 1, 8)
        trimmed_data = [[cell[:40] for cell in row[:max_cols]] for row in table_data[:100]]

        t = Table(trimmed_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F8FAFC')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ]))
        elements.append(t)
        pdf.build(elements)
        return out_path

    elif target_format == 'docx':
        out_path = get_unique_output_path(base_name, "docx", output_dir)
        doc = Document()
        doc.add_heading(base_name.replace("_", " ").title(), level=1)

        if table_data:
            num_rows = min(len(table_data), 150)
            num_cols = len(table_data[0])
            word_table = doc.add_table(rows=num_rows, cols=num_cols)
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            word_table.style = 'Light Shading Accent 1'

            for r_idx in range(num_rows):
                for c_idx in range(num_cols):
                    val = table_data[r_idx][c_idx] if c_idx < len(table_data[r_idx]) else ""
                    word_table.rows[r_idx].cells[c_idx].text = str(val)

        doc.save(out_path)
        return out_path

    elif target_format == 'csv':
        out_path = get_unique_output_path(base_name, "csv", output_dir)
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(table_data)
        return out_path

    elif target_format == 'xlsx':
        # Re-encode as formatted Excel
        text_repr = "\n".join(["\t".join(r) for r in table_data])
        return text_to_excel(text_repr, base_name, output_dir)

    elif target_format == 'txt':
        out_path = get_unique_output_path(base_name, "txt", output_dir)
        with open(out_path, "w", encoding="utf-8") as f:
            for r in table_data:
                f.write("\t".join(r) + "\n")
        return out_path

    elif target_format == 'html':
        rows_html = []
        for r_idx, r in enumerate(table_data):
            tag = "th" if r_idx == 0 else "td"
            cells = "".join(f"<{tag}>{c}</{tag}>" for c in r)
            rows_html.append(f"<tr>{cells}</tr>")
        table_html = f"<table border='1' cellpadding='8' style='border-collapse:collapse;width:100%;'>\n" + "\n".join(rows_html) + "\n</table>"
        return text_to_html(table_html, base_name, output_dir)

    return None


# ---------------------------------------------------------------------------
# Converters: PDF to Targets (TXT, DOCX)
# ---------------------------------------------------------------------------

def pdf_to_format(filepath, target_format, output_dir=None, custom_name=None):
    """Extracts text or content from PDF and converts to TXT or DOCX."""
    base_name = custom_name or Path(filepath).stem
    reader = pypdf.PdfReader(filepath)
    extracted_text = []

    for idx, page in enumerate(reader.pages):
        t = page.extract_text()
        if t:
            extracted_text.append(f"--- Page {idx + 1} ---\n" + t)

    full_text = "\n\n".join(extracted_text)
    if not full_text.strip():
        full_text = f"(No selectable text found in {Path(filepath).name}. The PDF might be scanned/image-based.)"

    if target_format == 'txt':
        return text_to_txt(full_text, base_name, output_dir)
    elif target_format == 'docx':
        return text_to_docx(full_text, base_name, output_dir)
    return None


# ---------------------------------------------------------------------------
# Converters: Word (.docx) to Targets (PDF, TXT, HTML)
# ---------------------------------------------------------------------------

def word_to_format(filepath, target_format, output_dir=None, custom_name=None):
    """Converts a Word .docx document into TXT, PDF, or HTML."""
    base_name = custom_name or Path(filepath).stem
    doc = Document(filepath)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    full_text = "\n\n".join(paragraphs)

    if target_format == 'txt':
        return text_to_txt(full_text, base_name, output_dir)
    elif target_format == 'pdf':
        return text_to_pdf(full_text, base_name, output_dir)
    elif target_format == 'html':
        return text_to_html(full_text, base_name, output_dir)
    return None


# ---------------------------------------------------------------------------
# Unified Conversion Dispatcher
# ---------------------------------------------------------------------------

def open_folder_picker_async(initial_dir, callback):
    """Opens a native Windows folder browser dialog in a background thread."""
    import threading
    def _run():
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes("-topmost", 1)
            folder = filedialog.askdirectory(initialdir=initial_dir or str(Path.home()), title="Select Output Folder")
            root.destroy()
            if folder and os.path.isdir(folder):
                callback(os.path.abspath(folder))
        except Exception:
            pass
    threading.Thread(target=_run, daemon=True).start()


def execute_conversion(input_obj, target_format, output_dir=None, custom_name=None):

    """
    Main dispatch function:
    Takes an input object (from read_windows_clipboard() or inspect_file())
    and executes the conversion to target_format.
    Supports custom_name and custom output_dir.
    Returns: {"success": bool, "output_path": str, "error": str, "name": str, "size_str": str}
    """
    try:
        kind = input_obj.get("kind")
        name = custom_name or input_obj.get("custom_name") or input_obj.get("name", "converted")
        output_path = None

        if kind == "text":
            text = input_obj.get("data", "")
            if target_format == "docx":
                output_path = text_to_docx(text, name, output_dir)
            elif target_format == "pdf":
                output_path = text_to_pdf(text, name, output_dir)
            elif target_format == "xlsx":
                output_path = text_to_excel(text, name, output_dir)
            elif target_format in ("png", "jpeg", "jpg"):
                output_path = text_to_image_card(text, name, output_dir)
            elif target_format == "txt":
                output_path = text_to_txt(text, name, output_dir)
            elif target_format == "md":
                output_path = get_unique_output_path(name, "md", output_dir)
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(text)
            elif target_format == "html":
                output_path = text_to_html(text, name, output_dir)

        elif kind == "text_file":
            filepath = input_obj.get("path")
            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
            if target_format == "docx":
                output_path = text_to_docx(text, name, output_dir)
            elif target_format == "pdf":
                output_path = text_to_pdf(text, name, output_dir)
            elif target_format == "xlsx":
                output_path = text_to_excel(text, name, output_dir)
            elif target_format in ("png", "jpeg", "jpg"):
                output_path = text_to_image_card(text, name, output_dir)
            elif target_format == "txt":
                output_path = text_to_txt(text, name, output_dir)
            elif target_format == "md":
                output_path = get_unique_output_path(name, "md", output_dir)
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(text)
            elif target_format == "html":
                output_path = text_to_html(text, name, output_dir)

        elif kind == "image":
            img = input_obj.get("data") or input_obj.get("path")
            output_path = image_to_format(img, target_format, name, output_dir)

        elif kind == "spreadsheet":
            filepath = input_obj.get("path")
            output_path = spreadsheet_to_format(filepath, target_format, output_dir, custom_name=name)

        elif kind == "pdf":
            filepath = input_obj.get("path")
            output_path = pdf_to_format(filepath, target_format, output_dir, custom_name=name)

        elif kind == "word":
            filepath = input_obj.get("path")
            output_path = word_to_format(filepath, target_format, output_dir, custom_name=name)

        elif kind == "url":
            url_val = input_obj.get("data") or input_obj.get("url", "")
            if target_format == "url":
                output_path = get_unique_output_path(name, "url", output_dir)
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(f"[InternetShortcut]\nURL={url_val}\n")
            elif target_format == "txt":
                output_path = text_to_txt(f"Web Link: {name}\nURL: {url_val}\nSaved: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", name, output_dir)
            elif target_format == "pdf":
                output_path = text_to_pdf(f"Web Link Bookmark\n\nTitle: {name}\nURL: {url_val}\nDate: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", name, output_dir)
            elif target_format == "html":
                output_path = get_unique_output_path(name, "html", output_dir)
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(f'<!DOCTYPE html><html><head><meta http-equiv="refresh" content="0; url={url_val}"><title>{name}</title></head><body>Redirecting to <a href="{url_val}">{url_val}</a>...</body></html>')

        elif kind == "html":
            raw_html = input_obj.get("data", "")
            clean_text = re.sub(r'<[^>]+>', ' ', raw_html)
            clean_text = re.sub(r'\s+', ' ', clean_text).strip()
            if target_format == "pdf":
                output_path = text_to_pdf(clean_text, name, output_dir)
            elif target_format == "docx":
                output_path = text_to_docx(clean_text, name, output_dir)
            elif target_format == "txt":
                output_path = text_to_txt(clean_text, name, output_dir)
            elif target_format == "html":
                output_path = text_to_html(clean_text, name, output_dir)

        if output_path and os.path.exists(output_path):
            size_b = os.path.getsize(output_path)
            return {
                "success": True,
                "output_path": output_path,
                "name": os.path.basename(output_path),
                "size_str": format_file_size(size_b),
                "error": None
            }
        else:
            return {
                "success": False,
                "output_path": None,
                "name": None,
                "size_str": None,
                "error": f"Unsupported conversion: {kind} -> {target_format}"
            }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {
            "success": False,
            "output_path": None,
            "name": None,
            "size_str": None,
            "error": str(e)
        }
