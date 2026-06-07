"""
Approximate page-count estimator for the FYP report.
Models A4 layout: text area 147mm wide x 247mm tall, TNR 12pt @ 1.5 spacing.
Walks the document body in order, accounting for paragraphs, headings,
tables, images and explicit page breaks.
"""
import os
from docx import Document
from docx.oxml.ns import qn
from PIL import Image

DOCX = '/projects/sandbox/fyp_report/output/FYP_Report_BioComposites.docx'
MEDIA_DIR = None  # use embedded sizes via zip

PAGE_H_MM = 247.0          # usable height (297 - 25 - 25)
TEXT_W_MM = 147.0          # usable width (210 - 38 - 25)
LINE_MM_12 = 12 * 1.5 * 0.3528   # 12pt * 1.5 spacing -> mm  (~6.35mm)
CHARS_PER_LINE_12 = 78     # approx chars per line at 12pt justified in 147mm
PT2MM = 0.3528

doc = Document(DOCX)

# Map relationship rIds -> embedded image dimensions (px) for inline images
import zipfile
z = zipfile.ZipFile(DOCX)
# Build rId -> media path
rels = {}
import re
rel_xml = z.read('word/_rels/document.xml.rels').decode('utf-8', 'ignore')
for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rel_xml):
    rels[m.group(1)] = m.group(2)

def img_dims(target):
    path = 'word/' + target.replace('../', '')
    try:
        with z.open(path) as f:
            im = Image.open(f); return im.size
    except Exception:
        return (1000, 750)

def emu_to_mm(emu):
    return emu / 914400.0 * 25.4

body = doc.element.body
total_mm = 0.0
pages = 1

def add(h):
    global total_mm, pages
    if total_mm + h > PAGE_H_MM:
        pages += 1
        total_mm = h
    else:
        total_mm += h

def newpage():
    global total_mm, pages
    pages += 1
    total_mm = 0.0

# Iterate top-level elements in order
from docx.text.paragraph import Paragraph
from docx.table import Table

for child in body.iterchildren():
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        para = Paragraph(child, doc)
        # explicit page break?
        xml = child.xml
        is_break = 'w:type="page"' in xml or 'lastRenderedPageBreak' in xml
        # space before/after
        pf = para.paragraph_format
        sb = (pf.space_before.pt if pf.space_before else 0) * PT2MM
        sa = (pf.space_after.pt if pf.space_after else 6) * PT2MM
        # inline image?
        if 'graphicData' in xml or 'pic:pic' in xml or '<wp:inline' in xml or 'a:blip' in xml:
            # find blip embed + extent
            ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', xml)
            if ext:
                h_mm = emu_to_mm(int(ext.group(2)))
            else:
                h_mm = 60
            add(h_mm + sb + sa)
            continue
        text = para.text
        # determine font size of first run
        fs = 12
        for r in para.runs:
            if r.font.size:
                fs = r.font.size.pt; break
        line_mm = fs * 1.5 * PT2MM
        cpl = max(20, int(CHARS_PER_LINE_12 * 12 / fs))
        nlines = max(1, -(-len(text) // cpl)) if text.strip() else 1
        h = nlines * line_mm + sb + sa
        add(h)
        if is_break:
            newpage()
    elif tag == 'tbl':
        tbl = Table(child, doc)
        ncols = len(tbl.columns) if tbl.columns else 1
        # chars per column at ~9.5pt across 147mm split among columns
        cpc = max(8, int((CHARS_PER_LINE_12 * 12 / 9.5) / max(1, ncols)))
        line_mm = 9.5 * 1.15 * PT2MM   # single-spaced body line at ~9.5pt
        tbl_mm = 0.0
        for row in tbl.rows:
            maxlines = 1
            for cell in row.cells:
                txt = cell.text
                nl = max(1, -(-len(txt) // cpc))
                maxlines = max(maxlines, nl)
            tbl_mm += maxlines * line_mm + 2.2   # + cell padding
        add(tbl_mm + 3)

print(f"Estimated pages: ~{pages}")
print(f"(Heuristic A4 model; actual Word pagination may vary by a few pages.)")
