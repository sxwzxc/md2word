"""
Python Cloud Function - Markdown to DOCX Converter
api/convert/md-to-docx.py → POST /api/convert/md-to-docx
Accepts a markdown text body and returns the .docx file as a base64-encoded
JSON response (so the binary survives any transport re-encoding on the platform).
"""
import re
import json
import base64
from io import BytesIO
from http.server import BaseHTTPRequestHandler

from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# Inline formatting helpers
# ---------------------------------------------------------------------------

_INLINE_RE = re.compile(
    r'(\*\*[^*]+\*\*'          # **bold**
    r'|\*\([^*]+\)\*'          # *(italic)*  (rare, keep simple)
    r'|\*[^*\s][^*]*\*'        # *italic*
    r'|`[^`]+`'                # `inline code`
    r'|\[[^\]]+\]\([^)]+\)'    # [text](url)
    r'|~~[^~]+~~)'             # ~~strikethrough~~
)


def _add_shading(paragraph, fill="F2F2F2"):
    """Apply background shading to a paragraph (used for code blocks)."""
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    p_pr.append(shd)


# East Asian font name — ensures Chinese / Japanese / Korean characters
# render with a real font instead of showing as missing-glyph boxes (叉叉).
_EA_FONT = '微软雅黑'


def _set_run_fonts(run, ascii_font='Calibri', ea_font=_EA_FONT):
    """Set both ASCII and East-Asian fonts on a run so CJK text renders."""
    run.font.name = ascii_font
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn('w:rFonts'))
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    r_fonts.set(qn('w:eastAsia'), ea_font)


def _add_formatted_runs(paragraph, text):
    """Parse inline markdown and add styled runs to *paragraph*."""
    if not text:
        return
    parts = _INLINE_RE.split(text)
    for part in parts:
        if not part:
            continue
        # Bold
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
            _set_run_fonts(run)
        # Strikethrough
        elif part.startswith('~~') and part.endswith('~~') and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.font.strike = True
            _set_run_fonts(run)
        # Inline code
        elif part.startswith('`') and part.endswith('`') and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            _set_run_fonts(run, ascii_font='Consolas')
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)
        # Link [text](url)
        elif part.startswith('[') and ']' in part and '(' in part:
            m = re.match(r'\[([^\]]+)\]\(([^)]+)\)', part)
            if m:
                run = paragraph.add_run(m.group(1))
                _set_run_fonts(run)
                run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                run.underline = True
            else:
                run = paragraph.add_run(part)
                _set_run_fonts(run)
        # Italic (single *)
        elif part.startswith('*') and part.endswith('*') and not part.startswith('**') and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
            _set_run_fonts(run)
        else:
            run = paragraph.add_run(part)
            _set_run_fonts(run)


# ---------------------------------------------------------------------------
# Block-level markdown → docx conversion
# ---------------------------------------------------------------------------

def _convert_table(doc, header_row, data_rows):
    """Add a markdown table to the document."""
    cols = len(header_row)
    table = doc.add_table(rows=1, cols=cols)
    table.style = 'Table Grid'
    # Header
    for idx, cell_text in enumerate(header_row):
        cell = table.rows[0].cells[idx]
        cell.text = ''
        p = cell.paragraphs[0]
        _add_formatted_runs(p, cell_text.strip())
        for run in p.runs:
            run.bold = True
            _set_run_fonts(run)
    # Data
    for row_cells in data_rows:
        cells = table.add_row().cells
        for idx in range(cols):
            cells[idx].text = ''
            _add_formatted_runs(cells[idx].paragraphs[0], row_cells[idx].strip() if idx < len(row_cells) else '')


def _setup_styles(doc):
    """Configure document styles with proper East-Asian fonts and heading sizes."""
    from docx.shared import Pt, RGBColor

    # --- Normal (body) style ---
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)
    # Set East-Asian font on the style level
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn('w:rFonts'))
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    r_fonts.set(qn('w:eastAsia'), _EA_FONT)
    r_fonts.set(qn('w:ascii'), 'Calibri')
    r_fonts.set(qn('w:hAnsi'), 'Calibri')

    # --- Heading styles: larger sizes + East-Asian font ---
    heading_config = {
        'Heading 1': (Pt(28), RGBColor(0x1F, 0x38, 0x64)),  # 28pt, dark blue
        'Heading 2': (Pt(22), RGBColor(0x2E, 0x4B, 0x8B)),  # 22pt
        'Heading 3': (Pt(16), RGBColor(0x37, 0x76, 0xAB)),  # 16pt
        'Heading 4': (Pt(14), RGBColor(0x44, 0x44, 0x44)),  # 14pt
        'Heading 5': (Pt(12), RGBColor(0x66, 0x66, 0x66)),  # 12pt
        'Heading 6': (Pt(11), RGBColor(0x66, 0x66, 0x66)),  # 11pt
    }
    for style_name, (size, color) in heading_config.items():
        try:
            hs = doc.styles[style_name]
        except KeyError:
            continue
        hs.font.size = size
        hs.font.color.rgb = color
        hs.font.bold = True
        # East-Asian font for headings
        h_rpr = hs.element.get_or_add_rPr()
        h_rfonts = h_rpr.find(qn('w:rFonts'))
        if h_rfonts is None:
            h_rfonts = OxmlElement('w:rFonts')
            h_rpr.append(h_rfonts)
        h_rfonts.set(qn('w:eastAsia'), _EA_FONT)
        h_rfonts.set(qn('w:ascii'), 'Calibri')
        h_rfonts.set(qn('w:hAnsi'), 'Calibri')


def convert_markdown_to_docx(markdown_text):
    """Convert markdown text to a .docx byte string."""
    doc = Document()

    # Configure styles with East-Asian fonts and readable heading sizes
    _setup_styles(doc)

    lines = markdown_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]

        # ---- Fenced code block ----
        stripped = line.strip()
        if stripped.startswith('```'):
            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            p = doc.add_paragraph()
            _add_shading(p)
            run = p.add_run('\n'.join(code_lines))
            _set_run_fonts(run, ascii_font='Consolas')
            run.font.size = Pt(9.5)
            continue

        # ---- Heading ----
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            _add_formatted_runs(doc.add_heading(level=level), m.group(2).strip())
            i += 1
            continue

        # ---- Horizontal rule ----
        if re.match(r'^\s*([-*_])\1{2,}\s*$', line):
            p = doc.add_paragraph()
            _add_shading(p, 'CCCCCC')
            i += 1
            continue

        # ---- Table ----
        if '|' in line and i + 1 < n and re.match(r'^\s*\|?[\s:|-]+\|?\s*$', lines[i + 1]):
            header_cells = [c for c in line.strip().strip('|').split('|')]
            i += 2  # skip separator
            data_rows = []
            while i < n and '|' in lines[i] and lines[i].strip():
                data_rows.append([c for c in lines[i].strip().strip('|').split('|')])
                i += 1
            _convert_table(doc, header_cells, data_rows)
            continue

        # ---- Blockquote ----
        if stripped.startswith('>'):
            quote_lines = []
            while i < n and lines[i].strip().startswith('>'):
                quote_lines.append(re.sub(r'^\s*>\s?', '', lines[i]))
                i += 1
            p = doc.add_paragraph('\n'.join(quote_lines))
            try:
                p.style = doc.styles['Intense Quote']
            except KeyError:
                p.style = doc.styles['Quote']
            continue

        # ---- Unordered list ----
        if re.match(r'^\s*[-*+]\s+', line):
            while i < n and re.match(r'^\s*[-*+]\s+', lines[i]):
                text = re.sub(r'^\s*[-*+]\s+', '', lines[i])
                p = doc.add_paragraph(style='List Bullet')
                _add_formatted_runs(p, text)
                i += 1
            continue

        # ---- Ordered list ----
        if re.match(r'^\s*\d+\.\s+', line):
            while i < n and re.match(r'^\s*\d+\.\s+', lines[i]):
                text = re.sub(r'^\s*\d+\.\s+', '', lines[i])
                p = doc.add_paragraph(style='List Number')
                _add_formatted_runs(p, text)
                i += 1
            continue

        # ---- Blank line ----
        if not stripped:
            i += 1
            continue

        # ---- Regular paragraph (gather consecutive non-empty, non-special lines) ----
        para_lines = [line]
        i += 1
        while i < n:
            nxt = lines[i]
            if (not nxt.strip()
                    or nxt.strip().startswith('```')
                    or re.match(r'^#{1,6}\s', nxt)
                    or re.match(r'^\s*[-*+]\s+', nxt)
                    or re.match(r'^\s*\d+\.\s+', nxt)
                    or nxt.strip().startswith('>')
                    or re.match(r'^\s*([-*_])\1{2,}\s*$', nxt)):
                break
            para_lines.append(nxt)
            i += 1
        p = doc.add_paragraph()
        _add_formatted_runs(p, ' '.join(l.strip() for l in para_lines))

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------

class handler(BaseHTTPRequestHandler):
    """POST /api/convert/md-to-docx — convert markdown body to base64-encoded .docx JSON."""

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            raw = self.rfile.read(length) if length else b''
            try:
                markdown_text = raw.decode('utf-8')
            except UnicodeDecodeError:
                markdown_text = raw.decode('utf-8', errors='replace')

            if not markdown_text.strip():
                self._json_error(400, "Request body is empty. Send markdown text as the raw POST body.")
                return

            docx_bytes = convert_markdown_to_docx(markdown_text)
            docx_b64 = base64.b64encode(docx_bytes).decode('ascii')

            payload = json.dumps({
                "ok": True,
                "filename": "converted.docx",
                "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "size": len(docx_bytes),
                "data": docx_b64,
            })

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('X-Powered-By', 'Python Cloud Function')
            self.end_headers()
            self.wfile.write(payload.encode('utf-8'))

        except Exception as exc:  # noqa: BLE001
            self._json_error(500, f"Conversion failed: {exc}")

    def do_GET(self):
        """Convenience info endpoint."""
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({
            "route": "/api/convert/md-to-docx",
            "method": "POST",
            "description": "Send markdown text as the raw POST body; receive a JSON with base64-encoded .docx in 'data'.",
            "contentType": "text/plain (utf-8)",
            "responseShape": {"ok": "bool", "filename": "str", "mime": "str", "size": "int", "data": "base64 str"},
        }).encode('utf-8'))

    def _json_error(self, status, message):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({"ok": False, "error": message}).encode('utf-8'))
