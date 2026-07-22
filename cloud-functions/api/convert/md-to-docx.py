"""
Python Cloud Function - Markdown to DOCX Converter
api/convert/md-to-docx.py → POST /api/convert/md-to-docx
Accepts a markdown text body and returns the .docx file as a base64-encoded
JSON response (so the binary survives any transport re-encoding on the platform).

Query param ?mode=1 (default, basic) or ?mode=2 (elegant typography).
"""
import re
import json
import base64
from io import BytesIO
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_LINE_SPACING
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


# East Asian font name — ensures Chinese / Japanese / Korean characters
# render with a real font instead of showing as missing-glyph boxes (叉叉).
_EA_FONT = '微软雅黑'

# Soft body text color for elegant mode (instead of harsh pure black).
_BODY_COLOR = RGBColor(0x33, 0x33, 0x33)


def _add_shading(paragraph, fill="F2F2F2"):
    """Apply background shading to a paragraph (used for code blocks)."""
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    p_pr.append(shd)


def _set_run_fonts(run, ascii_font='Calibri', ea_font=_EA_FONT):
    """Set both ASCII and East-Asian fonts on a run so CJK text renders."""
    run.font.name = ascii_font
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn('w:rFonts'))
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    _strip_theme_fonts(r_fonts)
    r_fonts.set(qn('w:eastAsia'), ea_font)
    r_fonts.set(qn('w:ascii'), ascii_font)
    r_fonts.set(qn('w:hAnsi'), ascii_font)


def _add_paragraph_border(paragraph, side='left', color='3776AB', sz='24'):
    """Add a single-side colored border to a paragraph (for blockquotes)."""
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn('w:pBdr'))
    if p_bdr is None:
        p_bdr = OxmlElement('w:pBdr')
        p_pr.append(p_bdr)
    bdr = OxmlElement(f'w:{side}')
    bdr.set(qn('w:val'), 'single')
    bdr.set(qn('w:sz'), sz)       # eighths of a point; 24 = 3pt
    bdr.set(qn('w:space'), '8')   # gap from text (twips)
    bdr.set(qn('w:color'), color)
    p_bdr.append(bdr)


def _set_paragraph_spacing(paragraph, before=None, after=None, line=None):
    """Set paragraph spacing. before/after in Pt, line as a multiple (e.g. 1.5)."""
    pf = paragraph.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE


def _add_hyperlink(paragraph, text, url, color='0563C1'):
    """Add a real, clickable hyperlink run to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True,
    )
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    r_pr = OxmlElement('w:rPr')

    r_fonts = OxmlElement('w:rFonts')
    r_fonts.set(qn('w:eastAsia'), _EA_FONT)
    r_fonts.set(qn('w:ascii'), 'Calibri')
    r_fonts.set(qn('w:hAnsi'), 'Calibri')
    r_pr.append(r_fonts)

    c = OxmlElement('w:color')
    c.set(qn('w:val'), color)
    r_pr.append(c)

    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    r_pr.append(u)

    new_run.append(r_pr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def _add_formatted_runs(paragraph, text, doc=None, mode=1):
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
            if mode == 2:
                run.font.color.rgb = _BODY_COLOR
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
                link_text, url = m.group(1), m.group(2)
                if mode == 2 and doc is not None:
                    _add_hyperlink(paragraph, link_text, url)
                else:
                    run = paragraph.add_run(link_text)
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
            if mode == 2:
                run.font.color.rgb = _BODY_COLOR


# ---------------------------------------------------------------------------
# Block-level markdown → docx conversion
# ---------------------------------------------------------------------------

def _shade_cell(cell, fill):
    """Apply background fill to a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    tc_pr.append(shd)


def _convert_table(doc, header_row, data_rows, mode=1):
    """Add a markdown table to the document."""
    cols = len(header_row)
    table = doc.add_table(rows=1, cols=cols)
    table.style = 'Table Grid'
    # Header
    for idx, cell_text in enumerate(header_row):
        cell = table.rows[0].cells[idx]
        cell.text = ''
        if mode == 2:
            _shade_cell(cell, '2B579A')  # Word blue header
        p = cell.paragraphs[0]
        if mode == 2:
            p.alignment = 1  # center
        _add_formatted_runs(p, cell_text.strip(), doc=doc, mode=mode)
        for run in p.runs:
            run.bold = True
            _set_run_fonts(run)
            if mode == 2:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Data
    for row_idx, row_cells in enumerate(data_rows):
        cells = table.add_row().cells
        for idx in range(cols):
            cell = cells[idx]
            cell.text = ''
            if mode == 2:
                # Zebra striping
                _shade_cell(cell, 'F2F6FC' if row_idx % 2 == 0 else 'FFFFFF')
            _add_formatted_runs(
                cell.paragraphs[0],
                row_cells[idx].strip() if idx < len(row_cells) else '',
                doc=doc, mode=mode,
            )


def _strip_theme_fonts(r_fonts):
    """Remove *Theme font attributes so explicit font names take precedence.

    The default python-docx template defines heading styles with
    asciiTheme/eastAsiaTheme/hAnsiTheme that point at theme fonts whose
    East-Asian slot is empty.  Word/WPS then falls back to a font that
    cannot render CJK, showing missing-glyph dots/boxes in front of
    headings.  Stripping these attributes lets our explicit
    ascii=Calibri / eastAsia=微软雅黑 win.
    """
    for attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        if r_fonts.get(qn(attr)) is not None:
            del r_fonts.attrib[qn(attr)]


def _strip_all_numbering(doc):
    """Remove ``<w:numPr>`` from non-list paragraph styles (Heading1-6,
    Subtitle, Title, etc.) and document defaults.

    The python-docx default template ships styles like ``Subtitle`` that
    carry a ``<w:numPr>`` element.  Some renderers (notably WPS Office)
    may render a stray bullet/dot if any heading-related style contains
    a numPr.  We strip numPr from every paragraph style EXCEPT list
    styles (ListBullet, ListNumber, …) so that real lists keep their
    bullets/numbers while headings stay clean.
    """
    styles_element = doc.styles.element
    for style in styles_element.findall(qn('w:style')):
        style_id = style.get(qn('w:styleId')) or ''
        # Keep numPr on list styles — they NEED it for bullets/numbers.
        if style_id.startswith('List') or 'List' in style_id:
            continue
        p_pr = style.find(qn('w:pPr'))
        if p_pr is not None:
            for num_pr in p_pr.findall(qn('w:numPr')):
                p_pr.remove(num_pr)
    # Also clean docDefaults/pPrDefault
    doc_defaults = styles_element.find(qn('w:docDefaults'))
    if doc_defaults is not None:
        p_pr_default = doc_defaults.find(qn('w:pPrDefault'))
        if p_pr_default is not None:
            p_pr = p_pr_default.find(qn('w:pPr'))
            if p_pr is not None:
                for num_pr in p_pr.findall(qn('w:numPr')):
                    p_pr.remove(num_pr)


def _setup_styles(doc, mode=1):
    """Configure document styles with proper East-Asian fonts and heading sizes."""
    # --- Normal (body) style ---
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)
    if mode == 2:
        normal.font.color.rgb = _BODY_COLOR
        # 1.5 line spacing + space after for readability
        pf = normal.paragraph_format
        pf.line_spacing = 1.5
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.space_after = Pt(6)
    # Set East-Asian font on the style level
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn('w:rFonts'))
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    _strip_theme_fonts(r_fonts)
    r_fonts.set(qn('w:eastAsia'), _EA_FONT)
    r_fonts.set(qn('w:ascii'), 'Calibri')
    r_fonts.set(qn('w:hAnsi'), 'Calibri')

    # --- Heading styles: larger sizes + East-Asian font ---
    if mode == 2:
        heading_config = {
            'Heading 1': (Pt(26), RGBColor(0x1A, 0x2A, 0x4F), 18, 8),
            'Heading 2': (Pt(20), RGBColor(0x2B, 0x57, 0x9A), 14, 6),
            'Heading 3': (Pt(15), RGBColor(0x37, 0x76, 0xAB), 10, 4),
            'Heading 4': (Pt(13), RGBColor(0x44, 0x55, 0x66), 8, 4),
            'Heading 5': (Pt(12), RGBColor(0x66, 0x66, 0x66), 6, 3),
            'Heading 6': (Pt(11), RGBColor(0x66, 0x66, 0x66), 6, 3),
        }
    else:
        heading_config = {
            'Heading 1': (Pt(28), RGBColor(0x1F, 0x38, 0x64), None, None),
            'Heading 2': (Pt(22), RGBColor(0x2E, 0x4B, 0x8B), None, None),
            'Heading 3': (Pt(16), RGBColor(0x37, 0x76, 0xAB), None, None),
            'Heading 4': (Pt(14), RGBColor(0x44, 0x44, 0x44), None, None),
            'Heading 5': (Pt(12), RGBColor(0x66, 0x66, 0x66), None, None),
            'Heading 6': (Pt(11), RGBColor(0x66, 0x66, 0x66), None, None),
        }
    for style_name, (size, color, before, after) in heading_config.items():
        try:
            hs = doc.styles[style_name]
        except KeyError:
            continue
        hs.font.size = size
        hs.font.color.rgb = color
        hs.font.bold = True
        if mode == 2 and before is not None:
            hs.paragraph_format.space_before = Pt(before)
            hs.paragraph_format.space_after = Pt(after)
        # East-Asian font for headings — strip theme refs first
        h_rpr = hs.element.get_or_add_rPr()
        h_rfonts = h_rpr.find(qn('w:rFonts'))
        if h_rfonts is None:
            h_rfonts = OxmlElement('w:rFonts')
            h_rpr.append(h_rfonts)
        _strip_theme_fonts(h_rfonts)
        h_rfonts.set(qn('w:eastAsia'), _EA_FONT)
        h_rfonts.set(qn('w:ascii'), 'Calibri')
        h_rfonts.set(qn('w:hAnsi'), 'Calibri')


def _setup_page(doc, mode=1):
    """Set page margins. Mode 2 uses comfortable 2.54 cm margins."""
    if mode != 2:
        return
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)


def _add_heading(doc, level, text, mode=1):
    """Add a heading using a clean paragraph that bypasses python-docx's
    default ``Heading`` style.

    The default ``Heading1``–``Heading6`` styles in the python-docx
    template can carry an inherited list-numbering reference (or a linked
    character style) that renders as a stray bullet/dot in front of the
    heading text.  Building the paragraph ourselves — with NO pStyle and
    NO numPr — eliminates the dot in all renderers (Word, WPS, etc.).

    Typography follows the common Chinese paper/report convention:
    黑体 (SimHei) bold for East-Asian text, with sizes close to
    三号/四号/小四 (16/14/12 pt) at levels 1–3.
    """
    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()

    # Outline level — keeps the document outline / TOC working even though
    # we don't use the built-in Heading style.
    outline = OxmlElement('w:outlineLvl')
    outline.set(qn('w:val'), str(level - 1))
    p_pr.append(outline)

    # 黑体 bold + paper-style sizes/colors.  Mode 2 (elegant) uses slightly
    # larger sizes & richer colors; mode 1 (basic) uses plainer colors.
    if mode == 2:
        heading_cfg = {
            1: (Pt(22), RGBColor(0x1A, 0x2A, 0x4F), 18, 8),  # 二号  深蓝
            2: (Pt(18), RGBColor(0x2B, 0x57, 0x9A), 14, 6),  # 三号  中蓝
            3: (Pt(15), RGBColor(0x37, 0x76, 0xAB), 10, 4),  # 小三  浅蓝
            4: (Pt(13), RGBColor(0x44, 0x55, 0x66), 8, 4),   # 四号  深灰蓝
            5: (Pt(12), RGBColor(0x55, 0x55, 0x55), 6, 3),   # 小四  中灰
            6: (Pt(11), RGBColor(0x66, 0x66, 0x66), 6, 3),   # 五号  浅灰
        }
    else:
        heading_cfg = {
            1: (Pt(22), RGBColor(0x1F, 0x38, 0x64), 12, 6),
            2: (Pt(18), RGBColor(0x2E, 0x4B, 0x8B), 10, 5),
            3: (Pt(15), RGBColor(0x37, 0x76, 0xAB), 8, 4),
            4: (Pt(13), RGBColor(0x44, 0x44, 0x44), 6, 3),
            5: (Pt(12), RGBColor(0x66, 0x66, 0x66), 6, 3),
            6: (Pt(11), RGBColor(0x66, 0x66, 0x66), 6, 3),
        }
    size, color, before, after = heading_cfg.get(level, heading_cfg[6])

    # Paragraph spacing — keep with next so a heading never sits alone
    # at the bottom of a page.
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = True
    pf.keep_together = True

    # Inline-format the heading text (preserves **bold**, `code`, links),
    # then override every run with the heading's bold / size / color and
    # the 黑体 East-Asian font.
    _add_formatted_runs(p, text, doc=doc, mode=mode)
    for run in p.runs:
        run.bold = True
        run.font.size = size
        run.font.color.rgb = color
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.find(qn('w:rFonts'))
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            r_pr.append(r_fonts)
        _strip_theme_fonts(r_fonts)
        # 黑体 (SimHei) — the standard Chinese paper/report heading typeface.
        r_fonts.set(qn('w:eastAsia'), '黑体')

    return p


def convert_markdown_to_docx(markdown_text, mode=1):
    """Convert markdown text to a .docx byte string."""
    doc = Document()

    # Configure styles with East-Asian fonts and readable heading sizes
    _setup_styles(doc, mode=mode)
    _setup_page(doc, mode=mode)
    # Strip ALL numbering definitions from every style — this guarantees
    # no stray bullet/dot can appear before headings or body text in any
    # renderer (Word, WPS Office, LibreOffice, Google Docs, …).
    _strip_all_numbering(doc)

    lines = markdown_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    i = 0
    n = len(lines)
    first_heading = None

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
            if mode == 2:
                _add_shading(p, 'F5F5F5')
                _add_paragraph_border(p, 'left', '3776AB', '18')
                _set_paragraph_spacing(p, before=6, after=6, line=1.0)
                p.paragraph_format.left_indent = Cm(0.5)
            else:
                _add_shading(p)
            run = p.add_run('\n'.join(code_lines))
            _set_run_fonts(run, ascii_font='Consolas')
            run.font.size = Pt(9.5)
            continue

        # ---- Heading ----
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            heading_text = m.group(2).strip()
            if first_heading is None and level == 1:
                first_heading = heading_text
            _add_heading(doc, level, heading_text, mode=mode)
            i += 1
            continue

        # ---- Horizontal rule ----
        if re.match(r'^\s*([-*_])\1{2,}\s*$', line):
            p = doc.add_paragraph()
            if mode == 2:
                _add_paragraph_border(p, 'bottom', 'CCCCCC', '6')
                _set_paragraph_spacing(p, before=8, after=8)
            else:
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
            _convert_table(doc, header_cells, data_rows, mode=mode)
            continue

        # ---- Blockquote ----
        if stripped.startswith('>'):
            quote_lines = []
            while i < n and lines[i].strip().startswith('>'):
                quote_lines.append(re.sub(r'^\s*>\s?', '', lines[i]))
                i += 1
            quote_text = '\n'.join(quote_lines)
            p = doc.add_paragraph()
            if mode == 2:
                _add_paragraph_border(p, 'left', '3776AB', '24')
                _set_paragraph_spacing(p, before=6, after=6, line=1.4)
                p.paragraph_format.left_indent = Cm(0.6)
                run = p.add_run(quote_text)
                _set_run_fonts(run)
                run.italic = True
                run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            else:
                run = p.add_run(quote_text)
                _set_run_fonts(run)
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
                _add_formatted_runs(p, text, doc=doc, mode=mode)
                if mode == 2:
                    _set_paragraph_spacing(p, before=2, after=2, line=1.4)
                i += 1
            continue

        # ---- Ordered list ----
        if re.match(r'^\s*\d+\.\s+', line):
            while i < n and re.match(r'^\s*\d+\.\s+', lines[i]):
                text = re.sub(r'^\s*\d+\.\s+', '', lines[i])
                p = doc.add_paragraph(style='List Number')
                _add_formatted_runs(p, text, doc=doc, mode=mode)
                if mode == 2:
                    _set_paragraph_spacing(p, before=2, after=2, line=1.4)
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
        _add_formatted_runs(p, ' '.join(l.strip() for l in para_lines), doc=doc, mode=mode)

    # Set document title from first H1 (metadata)
    if first_heading:
        try:
            doc.core_properties.title = first_heading
        except Exception:
            pass

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
            # Parse mode from query string (?mode=1 or ?mode=2)
            parsed = urlparse(self.path)
            qs = parse_qs(parsed.query)
            mode_str = qs.get('mode', ['1'])[0]
            mode = 2 if mode_str == '2' else 1

            length = int(self.headers.get('Content-Length', 0))
            raw = self.rfile.read(length) if length else b''
            try:
                markdown_text = raw.decode('utf-8')
            except UnicodeDecodeError:
                markdown_text = raw.decode('utf-8', errors='replace')

            if not markdown_text.strip():
                self._json_error(400, "Request body is empty. Send markdown text as the raw POST body.")
                return

            docx_bytes = convert_markdown_to_docx(markdown_text, mode=mode)
            docx_b64 = base64.b64encode(docx_bytes).decode('ascii')

            payload = json.dumps({
                "ok": True,
                "mode": mode,
                "ver": "v3-no-dot",
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
            "modes": {
                "1": "basic — clean conversion with CJK font support",
                "2": "elegant — 1.5 line spacing, styled headings, code borders, zebra tables, hyperlinks, blockquote bars",
            },
            "usage": "POST /api/convert/md-to-docx?mode=2  (default mode=1)",
            "responseShape": {"ok": "bool", "mode": "int", "filename": "str", "mime": "str", "size": "int", "data": "base64 str"},
        }).encode('utf-8'))

    def _json_error(self, status, message):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({"ok": False, "error": message}).encode('utf-8'))
