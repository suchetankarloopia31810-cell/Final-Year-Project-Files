#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build the complete B.Tech Final Year Project Report
"Development and Characterization of Bio-Composites from Waste Aquatic
 Biomass for Sustainable Insulation"  — NIT Srinagar, June 2026

Formatting per department guidelines:
  Times New Roman 12 pt, 1.5 line spacing, A4,
  margins L=38mm R=25mm T=25mm B=25mm,
  Roman numerals for preliminary pages, Arabic for main body,
  IEEE references (order of citation).
"""

import os
from docx import Document
from docx.shared import Pt, Mm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CHARTS = '/projects/sandbox/fyp_report/charts'
PHOTOS = '/projects/sandbox/fyp_report/photos'
OUTPUT = '/projects/sandbox/fyp_report/output/FYP_Report_BioComposites.docx'

FONT = 'Times New Roman'
ACCENT = RGBColor(0x1F, 0x3B, 0x66)
GREEN  = RGBColor(0x2E, 0x6B, 0x2E)
GREY   = RGBColor(0x55, 0x55, 0x55)

# ─────────────────────────────────────────────────────────────────────────────
#  LOW-LEVEL HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    th = OxmlElement('w:tblHeader')
    th.set(qn('w:val'), "true")
    trPr.append(th)

def add_page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
    r = run._r.find(qn('w:rPr'))
    return run

def set_number_format(section, fmt, start=1):
    """fmt: 'decimal' or 'lowerRoman'"""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), fmt)
    pgNumType.set(qn('w:start'), str(start))

def style_base(doc):
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(12)
    st._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)

def P(doc, text='', size=12, bold=False, italic=False, align=None,
      color=None, space_before=0, space_after=6, line=1.5, font=FONT,
      underline=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    if text:
        r = p.add_run(text)
        r.font.name = font
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        r.underline = underline
        if color is not None:
            r.font.color.rgb = color
    return p

def run_add(p, text, size=12, bold=False, italic=False, color=None, font=FONT):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color is not None:
        r.font.color.rgb = color
    return r

def heading(doc, text, level=1, space_before=14, space_after=8):
    """Uses built-in Heading styles (so the TOC field can pick them up),
    with Times New Roman font override."""
    sizes = {1: 15, 2: 13, 3: 12}
    style_name = f'Heading {min(level,3)}'
    p = doc.add_paragraph(style=style_name)
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(sizes.get(level, 12))
    r.bold = True
    r.font.color.rgb = ACCENT if level <= 2 else RGBColor(0x20,0x20,0x20)
    return p

def add_toc_field(doc, switches=r'TOC \o "1-3" \h \z \u'):
    p = doc.add_paragraph()
    run = p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = switches
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = "Right-click and select 'Update Field' to generate the Table of Contents."
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    run._r.append(t); run._r.append(fld3)
    return p

def add_tof_field(doc, label='Figure'):
    """Table of Figures / Tables field based on caption label."""
    p = doc.add_paragraph()
    run = p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = f'TOC \\h \\z \\c "{label}"'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = f"Right-click and 'Update Field' to generate the List of {label}s."
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    run._r.append(t); run._r.append(fld3)
    return p

def seq_caption(doc, label, text, size=10.5, space_after=12):
    """Caption with a SEQ field so it is auto-numbered and picked up by the
    List of Figures/Tables. label = 'Figure' or 'Table'."""
    cp = doc.add_paragraph(style='Caption')
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(2)
    cp.paragraph_format.space_after = Pt(space_after)
    cp.paragraph_format.line_spacing = 1.0
    r0 = cp.add_run(f'{label} ')
    r0.font.name = FONT; r0.font.size = Pt(size); r0.bold = True
    r0.italic = False; r0.font.color.rgb = RGBColor(0,0,0)
    # SEQ field
    run = cp.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = f' SEQ {label} \\* ARABIC '
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    run.font.name = FONT; run.font.size = Pt(size); run.bold = True
    rt = cp.add_run(f'.  {text}')
    rt.font.name = FONT; rt.font.size = Pt(size); rt.italic = False
    rt.font.color.rgb = RGBColor(0,0,0)
    return cp

def img(doc, path, width_in, cap_text, placeholder_label=None):
    """Insert a centered image (or placeholder) followed by an auto-numbered
    Figure caption."""
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run()
        run.add_picture(path, width=Inches(width_in))
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(18)
        label = placeholder_label or os.path.basename(path)
        # bordered placeholder
        pPr = p._p.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        for side in ('top','bottom','left','right'):
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6')
            b.set(qn('w:space'),'8'); b.set(qn('w:color'),'AAAAAA')
            pbdr.append(b)
        pPr.append(pbdr)
        run_add(p, f'[ {label} \u2014 to be inserted ]',
                size=11, italic=True, color=GREY)
    seq_caption(doc, 'Figure', cap_text)

def tbl_cap(doc, text):
    """Auto-numbered Table caption placed above a table."""
    seq_caption(doc, 'Table', text, space_after=4)

def img_row(doc, items, total_width_in=6.2, cap_text=None):
    """Place multiple images side-by-side in a borderless table, with optional
    sub-labels, followed by a single auto-numbered Figure caption.
    items = list of (path, sublabel) tuples."""
    n = len(items)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    each_w = total_width_in / n
    for i, (path, sub) in enumerate(items):
        cell = t.rows[0].cells[i]
        cell.width = Inches(each_w)
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(2)
        if os.path.exists(path):
            run = para.add_run()
            run.add_picture(path, width=Inches(each_w - 0.18))
        else:
            run_add(para, f'[ {sub or os.path.basename(path)} ]',
                    size=10, italic=True, color=GREY)
        if sub:
            sp = cell.add_paragraph()
            sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sp.paragraph_format.space_after = Pt(2)
            run_add(sp, sub, size=9.5, italic=True, color=GREY)
    # remove borders
    tbl = t._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'),'none'); borders.append(e)
    tblPr.append(borders)
    if cap_text:
        seq_caption(doc, 'Figure', cap_text)

def make_table(doc, headers, rows, col_widths=None, header_bg='1F3B66',
               font_size=10.5, header_color='FFFFFF', zebra=True,
               highlight=None):
    """highlight: dict {(row_idx, col_idx): 'hexcolor'} for special cells"""
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'
    # header
    hdr = t.rows[0].cells
    set_repeat_header(t.rows[0])
    for i, h in enumerate(headers):
        set_cell_bg(hdr[i], header_bg)
        para = hdr[i].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.line_spacing = 1.0
        para.paragraph_format.space_after = Pt(2)
        para.paragraph_format.space_before = Pt(2)
        r = para.add_run(h)
        r.font.name = FONT; r.font.size = Pt(font_size); r.bold = True
        r.font.color.rgb = RGBColor.from_string(header_color)
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # body
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for ci, val in enumerate(row):
            para = cells[ci].paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.space_before = Pt(2)
            r = para.add_run(str(val))
            r.font.name = FONT; r.font.size = Pt(font_size)
            cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if zebra and ri % 2 == 1:
                set_cell_bg(cells[ci], 'F2F5FA')
            if highlight and (ri, ci) in highlight:
                set_cell_bg(cells[ci], highlight[(ri, ci)])
    if col_widths:
        for row in t.rows:
            for ci, w in enumerate(col_widths):
                row.cells[ci].width = Inches(w)
    return t

def hrule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '1F3B66')
    pbdr.append(bottom); pPr.append(pbdr)
    p.paragraph_format.space_after = Pt(4)

def page_break(doc):
    doc.add_page_break()

# ─────────────────────────────────────────────────────────────────────────────
#  DOCUMENT SETUP
# ─────────────────────────────────────────────────────────────────────────────
doc = Document()
style_base(doc)

sec = doc.sections[0]
sec.page_height = Mm(297); sec.page_width = Mm(210)   # A4
sec.left_margin   = Mm(38)
sec.right_margin  = Mm(25)
sec.top_margin    = Mm(25)
sec.bottom_margin = Mm(25)

print("Document base configured.")
print("Builder part 1 (setup + helpers) loaded OK.")



# ═════════════════════════════════════════════════════════════════════════════
#  TITLE PAGE
# ═════════════════════════════════════════════════════════════════════════════
P(doc, 'DEVELOPMENT AND CHARACTERIZATION OF BIO-COMPOSITES',
  size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT,
  space_before=6, space_after=2)
P(doc, 'FROM WASTE AQUATIC BIOMASS FOR SUSTAINABLE INSULATION',
  size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT,
  space_after=14)

P(doc, 'A Project Report submitted in partial fulfilment of the requirements',
  size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
P(doc, 'for the award of the degree of', size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
P(doc, 'BACHELOR OF TECHNOLOGY', size=14, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'in', size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'CHEMICAL ENGINEERING', size=14, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

P(doc, 'Submitted by', size=12, italic=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
P(doc, 'Suchetan Karloopia          (2022BCHE019)', size=12.5, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'Miran Haider                    (2022BCHE027)', size=12.5, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'Akshita Sen                      (2022BCHE037)', size=12.5, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

P(doc, 'Under the supervision of', size=12, italic=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
P(doc, 'Dr. Fasil Qayoom Mir', size=13.5, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'Head & Associate Professor', size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'Department of Chemical Engineering', size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

# Institute crest placeholder / name
P(doc, 'DEPARTMENT OF CHEMICAL ENGINEERING', size=13, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=2)
P(doc, 'NATIONAL INSTITUTE OF TECHNOLOGY SRINAGAR', size=13, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=2)
P(doc, 'Hazratbal, Srinagar – 190006, Jammu & Kashmir, India', size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
P(doc, 'JUNE 2026', size=13, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
page_break(doc)

# ═════════════════════════════════════════════════════════════════════════════
#  CANDIDATE'S DECLARATION
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "CANDIDATES' DECLARATION", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=12)
hrule(doc)
P(doc, "We hereby declare that the work presented in this project report entitled "
  "\u201cDevelopment and Characterization of Bio-Composites from Waste Aquatic "
  "Biomass for Sustainable Insulation\u201d, submitted to the Department of "
  "Chemical Engineering, National Institute of Technology Srinagar, in partial "
  "fulfilment of the requirements for the award of the degree of Bachelor of "
  "Technology in Chemical Engineering, is an authentic record of our own work "
  "carried out during the period August 2025 to June 2026 under the supervision "
  "of Dr. Fasil Qayoom Mir, Head and Associate Professor, Department of Chemical "
  "Engineering, NIT Srinagar.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "The matter embodied in this report has not been submitted by us for the "
  "award of any other degree or diploma of this or any other Institute or "
  "University. Wherever the work of other researchers has been used, it has been "
  "duly acknowledged and cited in the references.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=24)

for name, enrol in [("Suchetan Karloopia", "2022BCHE019"),
                    ("Miran Haider", "2022BCHE027"),
                    ("Akshita Sen", "2022BCHE037")]:
    pp = doc.add_paragraph()
    pp.paragraph_format.space_after = Pt(18)
    run_add(pp, f"{name}", size=12, bold=True)
    run_add(pp, f"   (Enrolment No. {enrol})", size=12)
    pp2 = doc.add_paragraph()
    pp2.paragraph_format.space_after = Pt(4)
    run_add(pp2, "Signature: ____________________________", size=12)

P(doc, "Place: Srinagar, J&K", size=12, space_before=12, space_after=2)
P(doc, "Date:  ______________", size=12, space_after=2)
page_break(doc)

# ═════════════════════════════════════════════════════════════════════════════
#  CERTIFICATE (Supervisor)
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "CERTIFICATE", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=12)
hrule(doc)
P(doc, "This is to certify that the project report entitled \u201cDevelopment and "
  "Characterization of Bio-Composites from Waste Aquatic Biomass for Sustainable "
  "Insulation\u201d, being submitted by Suchetan Karloopia (2022BCHE019), "
  "Miran Haider (2022BCHE027) and Akshita Sen (2022BCHE037) to the Department of "
  "Chemical Engineering, National Institute of Technology Srinagar, for the award "
  "of the degree of Bachelor of Technology in Chemical Engineering, is a record "
  "of bona fide work carried out by them under my supervision and guidance.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "The candidates have worked sincerely and diligently on the project. To "
  "the best of my knowledge, the matter embodied in this report has not been "
  "submitted, in part or in full, to any other University or Institute for the "
  "award of any degree or diploma.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "I wish them all success in their future endeavours.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=36)

pp = doc.add_paragraph(); pp.paragraph_format.space_after = Pt(2)
run_add(pp, "Dr. Fasil Qayoom Mir", size=12, bold=True)
P(doc, "Head & Associate Professor", size=12, space_after=2)
P(doc, "Department of Chemical Engineering", size=12, space_after=2)
P(doc, "National Institute of Technology Srinagar", size=12, space_after=2)
P(doc, "Srinagar – 190006, J&K, India", size=12, space_after=2)
page_break(doc)

print("Title + Declaration + Certificate added.")



# ═════════════════════════════════════════════════════════════════════════════
#  ACKNOWLEDGEMENT
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "ACKNOWLEDGEMENT", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=12)
hrule(doc)
P(doc, "We wish to express our deepest gratitude and sincere appreciation to our "
  "supervisor, Dr. Fasil Qayoom Mir, Head and Associate Professor, Department of "
  "Chemical Engineering, NIT Srinagar, for his invaluable guidance, constant "
  "encouragement, and constructive criticism throughout the course of this "
  "project. His insight and expertise were instrumental in shaping this work, and "
  "it has been a privilege to work under his supervision.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "We are grateful to Prof. Tabassum Ara, Department of Chemistry, NIT "
  "Srinagar, for kindly providing access to the mixer-grinder facility used for "
  "biomass size reduction, and to Dr. Khalid Majid for facilitating access to the "
  "KD2 Pro Thermal Properties Analyzer for thermal conductivity measurements.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "We extend our thanks to the faculty and technical staff of the Department "
  "of Chemical Engineering for providing the laboratory facilities and assistance "
  "required to carry out the experimental work. We also acknowledge the support "
  "of the Department of Civil Engineering for access to the Baker Type K12 "
  "unconfined compression testing apparatus.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "Finally, we thank our families and friends for their unwavering support "
  "and motivation throughout our academic journey.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=24)

pp = doc.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_add(pp, "Suchetan Karloopia", size=12, bold=True)
pp = doc.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_add(pp, "Miran Haider", size=12, bold=True)
pp = doc.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_add(pp, "Akshita Sen", size=12, bold=True)
page_break(doc)

# ═════════════════════════════════════════════════════════════════════════════
#  ABSTRACT
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "ABSTRACT", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=12)
hrule(doc)
P(doc, "The rising energy demand of the building sector and the environmental "
  "burden of petroleum-derived insulation materials such as expanded polystyrene "
  "(EPS), polyurethane (PU) foam and glass wool have intensified the search for "
  "sustainable, bio-based alternatives. Simultaneously, the invasive aquatic weeds "
  "water hyacinth (Eichhornia crassipes) and water lily (Nymphaea spp.) infest "
  "Dal Lake, Srinagar \u2014 a Ramsar wetland of about 18 km\u00b2 severely affected "
  "by eutrophication \u2014 creating a large, freely available waste biomass stream. "
  "This project addresses both problems simultaneously by valorising waste aquatic "
  "biomass into bio-composite insulation panels.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "Four composite formulations were fabricated using oven-dried, ground "
  "biomass classified by ASTM E11 sieves: a coarse binder-less panel (S1), a fine "
  "binder-less panel (S2), and two fine panels bound with gelatinised food-grade "
  "corn starch at biomass:starch ratios of 90:10 (S3) and 70:30 (S4). No synthetic "
  "polymers or chemical crosslinkers were used. The panels were characterized for "
  "moisture content (ASTM D4442), bulk density (ASTM D1037), water absorption "
  "(ASTM D570), unconfined compressive strength (Baker Type K12 UCT apparatus) and "
  "thermal conductivity (KD2 Pro TR-3 transient line-source probe, ASTM D5334), "
  "with three replicates per test.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "The fine binder-less panel (S2) and the 90:10 starch panel (S3) achieved "
  "insulation-grade thermal conductivities of 0.0577 and 0.0608 W/m\u00b7K "
  "respectively, both below the 0.065 W/m\u00b7K threshold of ASTM C168. The 90:10 "
  "formulation (S3) recorded the highest unconfined compressive strength of "
  "186.0 kPa \u2014 a 97 % improvement over the binder-less fine panel \u2014 "
  "identifying 10 wt% corn starch as the optimal binder content. The high-starch "
  "panel (S4, 70:30) showed the lowest water absorption (280 %) but, "
  "counter-intuitively, the lowest strength (26.8 kPa) and a three-fold higher "
  "thermal conductivity (0.1846 W/m\u00b7K), attributed to an excess plasticised "
  "starch matrix and a high moisture content of 41.18 %. Overall, the 90:10 "
  "biomass:starch panel (S3) emerged as the best balanced formulation, combining "
  "insulation-grade thermal performance with the highest mechanical strength, and "
  "demonstrates the technical feasibility of converting an invasive aquatic weed "
  "into a viable, biodegradable building-insulation material.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
pk = doc.add_paragraph(); pk.paragraph_format.space_after = Pt(6)
run_add(pk, "Keywords: ", size=12, bold=True)
run_add(pk, "Water hyacinth; Water lily; Bio-composite; Corn starch binder; "
        "Thermal insulation; Unconfined compressive strength; Sustainable "
        "building materials; Waste valorisation.", size=12, italic=True)
page_break(doc)

print("Acknowledgement + Abstract added.")



# ═════════════════════════════════════════════════════════════════════════════
#  TABLE OF CONTENTS  (auto field)
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "TABLE OF CONTENTS", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=10)
hrule(doc)
add_toc_field(doc)
page_break(doc)

# ═════════════════════════════════════════════════════════════════════════════
#  LIST OF FIGURES  (auto field)
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "LIST OF FIGURES", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=10)
hrule(doc)
add_tof_field(doc, 'Figure')
page_break(doc)

# ═════════════════════════════════════════════════════════════════════════════
#  LIST OF TABLES  (auto field)
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "LIST OF TABLES", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=10)
hrule(doc)
add_tof_field(doc, 'Table')
page_break(doc)

# ═════════════════════════════════════════════════════════════════════════════
#  LIST OF ABBREVIATIONS & SYMBOLS
# ═════════════════════════════════════════════════════════════════════════════
P(doc, "LIST OF ABBREVIATIONS AND SYMBOLS", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=10)
hrule(doc)
abbr = [
    ("ASTM", "American Society for Testing and Materials"),
    ("WH", "Water Hyacinth (Eichhornia crassipes)"),
    ("WL", "Water Lily (Nymphaea spp.)"),
    ("WHF", "Water Hyacinth Fibre"),
    ("MC", "Moisture Content"),
    ("WA", "Water Absorption"),
    ("UCT", "Unconfined Compression Test"),
    ("SEM", "Scanning Electron Microscopy"),
    ("EPS", "Expanded Polystyrene"),
    ("PU", "Polyurethane"),
    ("IFR", "Intumescent Flame Retardant"),
    ("PBS", "Poly(butylene succinate)"),
    ("LOI", "Limiting Oxygen Index"),
    ("RH", "Relative Humidity"),
    ("LCA", "Life Cycle Assessment"),
    ("K", "Thermal Conductivity (W/m\u00b7K)"),
    ("R", "Thermal Resistivity (\u00b0C\u00b7cm/W)"),
    ("\u03c1", "Bulk Density (g/cm\u00b3)"),
    ("\u03c3", "Compressive (Deviator) Stress (kPa)"),
    ("\u03b5", "Axial Strain (\u2013 or %)"),
    ("q\u1d64", "Unconfined Compressive Strength (kPa)"),
    ("S\u1d64", "Undrained Shear Strength = q\u1d64/2 (kPa)"),
    ("\u03b5f", "Failure Strain (%)"),
    ("Syx", "Standard Error of Estimate (goodness-of-fit)"),
    ("A\u2080", "Initial Cross-Sectional Area (mm\u00b2)"),
    ("H\u2080", "Initial Specimen Height (mm)"),
]
make_table(doc, ["Abbreviation / Symbol", "Description"], abbr,
           col_widths=[2.0, 4.3], font_size=11, zebra=True)
page_break(doc)

print("TOC + Lists + Abbreviations added.")



# ═════════════════════════════════════════════════════════════════════════════
#  SECTION BREAK -> MAIN BODY (Arabic numerals from 1)
# ═════════════════════════════════════════════════════════════════════════════
main_sec = doc.add_section(WD_SECTION.NEW_PAGE)
main_sec.page_height = Mm(297); main_sec.page_width = Mm(210)
main_sec.left_margin = Mm(38); main_sec.right_margin = Mm(25)
main_sec.top_margin = Mm(25); main_sec.bottom_margin = Mm(25)
main_sec.footer.is_linked_to_previous = False

# ─────────────────────────────────────────────────────────────────────────────
#  CHAPTER 1 — INTRODUCTION
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "CHAPTER 1", level=1, space_before=0)
heading(doc, "INTRODUCTION", level=1, space_before=2)

heading(doc, "1.1  Background", level=2)
P(doc, "Thermal insulation is one of the most effective and economical strategies "
  "for reducing energy consumption in the building sector, which accounts for a "
  "substantial share of global energy use and greenhouse-gas emissions. The "
  "majority of commercial insulation products in use today \u2014 expanded "
  "polystyrene (EPS), extruded polystyrene (XPS), polyurethane (PU) foam, glass "
  "wool and mineral wool \u2014 are either petroleum-derived or energy-intensive "
  "to manufacture. These materials are non-biodegradable, contribute to landfill "
  "accumulation at the end of their service life, and in the case of organic foams "
  "are highly flammable and release toxic gases during combustion [1], [7]. The "
  "growing emphasis on circular-economy principles and carbon-footprint reduction "
  "has therefore driven significant research interest towards bio-based insulation "
  "materials derived from renewable and waste lignocellulosic resources.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "Natural-fibre bio-composites offer a compelling alternative because they "
  "are renewable, biodegradable, lightweight, possess low thermal conductivity "
  "owing to their porous cellular structure, and can often be produced at low "
  "cost from agricultural or aquatic waste streams. A wide range of lignocellulosic "
  "feedstocks \u2014 rice straw, corn cob, coir, hemp, flax, kenaf, sugarcane "
  "bagasse, pineapple leaf and date-palm waste, among others \u2014 has been "
  "investigated for insulation applications [4], [6]. Among aquatic feedstocks, "
  "water hyacinth (Eichhornia crassipes) has attracted particular attention "
  "because of its abundance, rapid growth and high cellulose content.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "1.2  The Dal Lake Problem: An Invasive Weed as a Resource", level=2)
P(doc, "Dal Lake in Srinagar, Jammu & Kashmir, is an iconic Himalayan urban lake "
  "and a designated Ramsar wetland site covering approximately 18 km\u00b2. Over "
  "the past several decades the lake has suffered severe ecological degradation "
  "due to eutrophication driven by untreated sewage inflow, catchment-area "
  "urbanisation and nutrient loading. A direct and highly visible consequence of "
  "this nutrient enrichment is the explosive proliferation of free-floating "
  "invasive aquatic weeds, principally water hyacinth (Eichhornia crassipes) and "
  "water lily (Nymphaea spp.). These weeds form dense surface mats that deplete "
  "dissolved oxygen, block sunlight, destroy native fish breeding grounds, impede "
  "navigation and accelerate the conversion of open water into marshland.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "Mechanical de-weeding operations remove large quantities of this biomass "
  "every year, but the harvested material is generally treated as waste and "
  "either dumped or left to decompose, releasing methane and returning nutrients "
  "to the ecosystem. This creates a continuous, freely available and otherwise "
  "problematic biomass stream. Converting this invasive weed into a useful "
  "engineering product simultaneously addresses two pressing problems: it provides "
  "an economic incentive and end-use for weed removal, and it supplies a "
  "renewable raw material for sustainable insulation \u2014 the central premise of "
  "this project.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

img(doc, f'{PHOTOS}/__missing_dallake.jpg', 5.6,
    "Dal Lake, Srinagar (\u2248 N 34.08\u00b0, E 74.85\u00b0) showing aquatic "
    "weed infestation \u2014 collection site of the raw biomass.",
    placeholder_label="Figure: Dal Lake collection site photograph")

heading(doc, "1.3  Motivation", level=2)
P(doc, "The motivation for this work arises from the convergence of three factors: "
  "(i) the environmental and health drawbacks of petroleum-based synthetic "
  "insulation; (ii) the urgent need to manage and valorise the invasive aquatic "
  "weed burden of Dal Lake; and (iii) the demonstrated potential of water-hyacinth "
  "fibre as a low-thermal-conductivity natural material reported in the recent "
  "literature [2], [3]. By using a food-grade, biodegradable corn-starch binder "
  "rather than synthetic resins or chemical crosslinkers, the present study aims "
  "to develop a fully bio-based, low-cost and environmentally benign insulation "
  "panel that can be fabricated with simple, scalable processing equipment.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "1.4  Problem Statement", level=2)
P(doc, "While waste aquatic biomass offers a promising sustainable alternative to "
  "synthetic insulation, its direct application is currently hindered by inherent "
  "material deficiencies \u2014 specifically low mechanical strength, high moisture "
  "absorption and susceptibility to fire. This necessitates the development of "
  "engineered bio-composites, with an appropriate binder and processing route, to "
  "ensure adequate structural integrity, dimensional stability and insulation "
  "performance while retaining biodegradability.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "1.5  Objectives", level=2)
P(doc, "The primary objectives of this project are:", align=WD_ALIGN_PARAGRAPH.JUSTIFY,
  space_after=2)
for i, obj in enumerate([
    "To develop eco-friendly bio-composite insulation panels from waste aquatic "
    "biomass (water hyacinth and water lily) collected from Dal Lake.",
    "To fabricate composite variants differing in particle size and "
    "biomass-to-binder ratio using a natural, gelatinised corn-starch binder.",
    "To characterize the physical properties of the panels \u2014 moisture "
    "content, bulk density and water absorption \u2014 following ASTM standards.",
    "To evaluate the mechanical performance of the panels through unconfined "
    "compression testing and to determine the optimal binder content.",
    "To measure the thermal conductivity of the panels and assess their "
    "suitability as insulation-grade materials against the ASTM C168 threshold.",
    "To identify the best-balanced formulation by comparing mechanical, thermal "
    "and moisture-resistance performance, and to benchmark it against "
    "conventional insulation materials reported in the literature.",
]):
    pp = doc.add_paragraph(style='List Number')
    pp.paragraph_format.line_spacing = 1.5
    pp.paragraph_format.space_after = Pt(3)
    r = pp.add_run(obj); r.font.name = FONT; r.font.size = Pt(12)

heading(doc, "1.6  Scope of the Work", level=2)
P(doc, "This study focuses on the laboratory-scale fabrication and physico-"
  "mechanical-thermal characterization of four bio-composite formulations. The "
  "characterization programme comprises moisture content (ASTM D4442), bulk "
  "density (ASTM D1037), water absorption (ASTM D570), unconfined compressive "
  "strength (Baker Type K12 apparatus) and thermal conductivity (KD2 Pro TR-3 "
  "probe, ASTM D5334). Scanning electron microscopy (SEM) is identified for "
  "microstructural analysis and the corresponding figure space is reserved. "
  "Flammability testing, hydrophobic surface treatments and full life-cycle "
  "assessment are identified as future work and are outside the present scope.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "1.7  Organisation of the Report", level=2)
P(doc, "Chapter 1 introduces the background, motivation, problem statement and "
  "objectives. Chapter 2 presents an extended literature review of natural-fibre "
  "and water-hyacinth-based insulation materials and identifies the research gap. "
  "Chapter 3 describes the raw materials, binder preparation, sample fabrication "
  "and the characterization methods and equipment. Chapter 4 presents and "
  "discusses the experimental results, with comparison against published work. "
  "Chapter 5 summarises the conclusions and outlines directions for future work.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
page_break(doc)

print("Chapter 1 added.")



# ─────────────────────────────────────────────────────────────────────────────
#  CHAPTER 2 — LITERATURE REVIEW
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "CHAPTER 2", level=1, space_before=0)
heading(doc, "LITERATURE REVIEW", level=1, space_before=2)

heading(doc, "2.1  Overview", level=2)
P(doc, "A wide body of research has explored natural fibres and agricultural / "
  "aquatic wastes as raw materials for sustainable thermal-insulation composites. "
  "This chapter reviews the most relevant studies, with particular emphasis on "
  "water-hyacinth-based composites, natural binder systems, fire-retardancy and "
  "end-of-life considerations. The review establishes the performance benchmarks "
  "against which the present work is compared and identifies the research gap that "
  "this project addresses.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "2.2  Water-Hyacinth-Based Insulation Composites", level=2)
P(doc, "Salas-Ruiz et al. [2] characterized binder-less water-hyacinth-petiole "
  "particle boards and reported thermal conductivities as low as 0.047 W/m\u00b7K "
  "for staple-fibre boards and around 0.065 W/m\u00b7K for pulp boards, with very "
  "high water absorption in the range of 450\u2013555 %, confirming both the "
  "excellent insulating potential and the moisture sensitivity of binder-less "
  "water-hyacinth boards. Philip and Rakendu [3] developed a water-hyacinth\u2013"
  "cement composite achieving a thermal conductivity of 0.0765 W/m\u00b7K at a "
  "density of 0.47 g/cm\u00b3, with flexural strength of about 0.35 MPa and water "
  "absorption near 98 %. Jaktorn and Jiajitsawat [1] produced insulation boards "
  "from water-hyacinth fibre bonded with natural rubber latex, obtaining very low "
  "thermal conductivities of 0.0246\u20130.0305 W/m\u00b7K, but found that water "
  "absorption failed to meet the relevant industrial standard, again highlighting "
  "moisture resistance as the principal limitation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "Anjani et al. [6] fabricated water-hyacinth\u2013sugarcane-bagasse "
  "fibre\u2013epoxy composites for cool-box insulation, reporting a thermal "
  "conductivity of about 0.1987 W/m\u00b7K and a maximum bending strength of "
  "263 kgf/cm\u00b2 for the best configuration. Chaireh et al. [5] developed "
  "starch\u2013water-hyacinth foams for food packaging and found that a 5 wt% "
  "water-hyacinth loading was optimal, that the elastic modulus reached about "
  "232 MPa, and that a beeswax coating substantially reduced water absorption "
  "\u2014 establishing both starch as a viable natural binder for water hyacinth "
  "and beeswax as an effective hydrophobic treatment.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "2.3  Binder Systems, Surface Treatment and Fire Retardancy", level=2)
P(doc, "The choice of binder strongly governs the mechanical and moisture behaviour "
  "of natural-fibre composites. Syamsuri et al. [8] showed that alkaline (NaOH) "
  "treatment of water-hyacinth fibre in a cassava-starch bioplastic improved the "
  "tensile strength by roughly four-fold relative to untreated fibre, by improving "
  "fibre\u2013matrix adhesion. Abral et al. [10], however, demonstrated that "
  "water-hyacinth\u2013polyester composites processed in the wet state were "
  "mechanically weaker, and that aggressive alkali treatment can introduce "
  "micro-voids that degrade performance \u2014 indicating that moisture control "
  "during processing and curing is critical. For fire safety, Suwanniroj and "
  "Suppakarn [9] used water-hyacinth fibre as a bio-based carbon source in an "
  "intumescent flame-retardant poly(butylene succinate) system, achieving a "
  "limiting oxygen index of 28.8 %, a UL-94 V-0 rating and a 53 % reduction in "
  "peak heat-release rate.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "2.4  Conventional Insulation Benchmarks", level=2)
P(doc, "To contextualise bio-based performance, conventional insulation materials "
  "provide useful benchmarks. Jeon et al. [7] reported a thermal conductivity of "
  "about 0.034 W/m\u00b7K for glass wool and noted that moisture ingress can cause "
  "a roughly four-fold increase in conductivity for fibrous inorganic insulations "
  "\u2014 a vulnerability shared by bio-based materials and one that underscores "
  "the importance of moisture management. The widely accepted classification "
  "threshold for an insulation-grade material is a thermal conductivity below "
  "0.065 W/m\u00b7K (ASTM C168), which is adopted as the principal performance "
  "criterion in the present study.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "2.5  Extended Literature Survey", level=2)
P(doc, "Table 1 presents an extended survey of recent (principally 2020\u20132025) "
  "studies on natural-fibre and waste-derived insulation and bio-composite "
  "materials, summarising the materials studied, key conditions, principal "
  "findings and the research gaps identified by the respective authors.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_cap(doc, "Extended literature survey of natural-fibre and waste-derived "
        "insulation / bio-composite materials.")

lit_headers = ["Author(s) & Year", "Focus / Materials", "Key Findings",
               "Research Gap / Remarks"]
lit_rows = [
    ["Jaktorn & Jiajitsawat (2021)",
     "Water-hyacinth fibre (WHF) + natural rubber latex; pressed at 100\u00b0C",
     "Excellent K = 0.0246\u20130.0305 W/m\u00b7K; higher latex raised density",
     "High water absorption failed TISI standard; moisture resistance needed"],
    ["Salas-Ruiz et al. (2019)",
     "Binder-less WH-petiole particle boards",
     "K = 0.047 (staple)\u20130.065 (pulp) W/m\u00b7K; WA 450\u2013555 %",
     "Binder-less boards highly moisture-sensitive"],
    ["Philip & Rakendu (2020/22)",
     "WH\u2013cement composite",
     "K = 0.0765 W/m\u00b7K; \u03c1 = 0.47 g/cm\u00b3; flexural 0.35 MPa; WA 98 %",
     "Cement raises embodied carbon; brittle"],
    ["Chaireh et al. (2020)",
     "Starch\u2013WH foam with beeswax coating",
     "5 wt% WH optimal; modulus 232 MPa; beeswax cut WA",
     "Foam for packaging; panel-scale insulation not addressed"],
    ["Anjani et al. (2023)",
     "WH + bagasse fibre\u2013epoxy (cool box)",
     "K = 0.1987 W/m\u00b7K; bending 263 kgf/cm\u00b2",
     "Epoxy is synthetic; not fully bio-based"],
    ["Jeon et al. (2017)",
     "Glass wool, mineral wool (polysiloxane coat)",
     "Glass wool K \u2248 0.034 W/m\u00b7K; moisture \u2192 4\u00d7 rise in K",
     "Inorganic; non-renewable; moisture-vulnerable"],
    ["Syamsuri et al. (2023)",
     "WH fibre / cassava-starch bioplastic",
     "NaOH treatment \u2192 ~4\u00d7 tensile improvement",
     "Bioplastic film, not insulation board"],
    ["Suwanniroj & Suppakarn (2023)",
     "WHF in IFR / PBS composite",
     "WHF as bio-carbon; LOI 28.8 %; UL-94 V-0; 53 % pHRR cut",
     "Focus on flammability, not thermal insulation"],
    ["Abral et al. (2014)",
     "WH fibre\u2013polyester composites",
     "Wet composites weaker; alkali introduced micro-voids",
     "Highlights need for moisture & treatment control"],
    ["Jaktorn & Jiajitsawat (2021)*",
     "WHF + NRL ratios (70 g WHF : 130\u2013170 g NRL)",
     "Good K; density rose with latex",
     "Moisture resistance below standard"],
    ["Zhou et al. (2022)",
     "Rice straw + sodium alginate / chitosan binders",
     "K = 0.038\u20130.047 W/m\u00b7K; CaCl\u2082 crosslink improved water resistance",
     "Significant mould growth; needs bio-preservatives"],
    ["Pinto et al. (2021)",
     "Corn cob vs extruded polystyrene (XPS)",
     "Corn cob has XPS-like closed-cell microstructure",
     "Lacks quantitative K data; needs standardised panel"],
    ["Yang et al. (2020)",
     "Review of mycelium-based bio-composites",
     "Low-cost binder; >75 % acoustic absorption at 1000 Hz",
     "Production not standardised"],
    ["Sahayaraj et al. (2023)",
     "Review of fire retardants for natural-fibre composites",
     "P/N additives, coatings, nanoparticles improve fire safety",
     "No standardised long-term durability data"],
    ["Kym\u00e4l\u00e4inen & Sj\u00f6berg (2022)",
     "Flax and hemp fibres for insulation",
     "Performance depends on fibre purity and processing",
     "Fibre-extraction consistency needed for scale-up"],
    ["Asdrubali et al. (2021)",
     "Review of acoustic properties of natural fibres",
     "High sound absorption from porous structure",
     "Low-frequency / humidity performance under-studied"],
    ["Chen et al. (2020)",
     "Boron-based flame retardants on wood-plastic composites",
     "Borax / boric acid raise LOI",
     "High loading reduces mechanical strength"],
    ["Aridi et al. (2023)",
     "Date-palm waste\u2013polyester insulation",
     "Lightweight; low K",
     "Poor fibre\u2013matrix adhesion; needs surface treatment"],
    ["Oushabi et al. (2022)",
     "Sugarcane-bagasse composites (NaOH treated)",
     "Alkali treatment improved strength & insulation",
     "High moisture absorption remains a challenge"],
    ["Wang et al. (2023)",
     "Pineapple-leaf-fibre (PALF) / epoxy",
     "High cellulose & strength; suitable for structural insulation",
     "Large-scale PALF extraction difficult"],
    ["Sair et al. (2021)",
     "Cork-based composites (hot-pressed)",
     "Closed-cell cork: good thermal, acoustic & fire resistance",
     "Limited availability; higher cost"],
    ["Muthuraj et al. (2022)",
     "Kenaf-fibre composites (UV/humidity aged)",
     "Low density with high strength",
     "Performance degrades under UV over time"],
    ["Oyejobi et al. (2020)",
     "Coir-fibre composites (compression moulded)",
     "Durable; high lignin makes them rigid",
     "Needs improved flexibility & interfacial bonding"],
    ["Zhang et al. (2024)",
     "Review of end-of-life management for bio-composites",
     "Pyrolysis (400\u2013600\u00b0C) viable for energy recovery",
     "Limited emissions data for mixed/additive composites"],
    ["Binici et al. (2023)",
     "Insulation from cotton & textile waste",
     "Effective thermal & acoustic insulation",
     "Waste heterogeneity \u2192 inconsistent panels"],
    ["Li et al. (2021)",
     "Hygrothermal behaviour of bamboo-fibre composites",
     "Good moisture buffering; regulates indoor humidity",
     "Susceptible to insect / fungal attack"],
    ["Liu et al. (2024)",
     "Review of aerogel thermal insulation",
     "Cellulose aerogels \u03bb < 0.020 W/m\u00b7K (super-insulating)",
     "Costly; not yet scalable"],
    ["Trabelsi et al. (2020)",
     "Hemp-shiv concrete (hygrothermal)",
     "Excellent moisture-buffer value",
     "Lower compressive strength vs concrete"],
    ["Pawlowski et al. (2025)",
     "Review of bio-based thermal-insulation composites",
     "Advances in bio-fibres, eco-resins & applications",
     "Comparative LCA among systems lacking"],
    ["Cosentino et al. (2023)",
     "Hemp, cork, kenaf, coir (RH & density)",
     "Hemp & cork show superior thermal / acoustic behaviour",
     "Need standard benchmarking under humidity"],
]
make_table(doc, lit_headers, lit_rows,
           col_widths=[1.35, 1.65, 1.85, 1.75], font_size=8.5, zebra=True)
P(doc, "* Conditions row for the same study, retained from the project literature "
  "survey for completeness.", size=9.5, italic=True, color=GREY, space_before=4)

heading(doc, "2.6  Research Gap", level=2)
P(doc, "The reviewed literature establishes that water-hyacinth fibre can yield "
  "insulation-grade thermal conductivities, but most reported systems rely either "
  "on synthetic matrices (epoxy, polyester, polyurethane), on energy-intensive "
  "binders (cement), or on binder-less boards that suffer from very high water "
  "absorption and poor mechanical integrity. Studies that do use natural binders "
  "are frequently directed at packaging films or foams rather than building-"
  "insulation panels, and few systematically optimise the biomass-to-binder ratio "
  "of a fully bio-based panel while reporting thermal, mechanical and moisture "
  "performance together. Moreover, the valorisation of aquatic weed specifically "
  "from Dal Lake \u2014 combining waste management with insulation development "
  "\u2014 has not been reported.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "The present work addresses this gap by fabricating fully bio-based panels "
  "from Dal-Lake aquatic biomass bound with food-grade corn starch (no synthetic "
  "polymers or crosslinkers), systematically varying particle size and "
  "biomass:starch ratio, and characterising the resulting thermal, mechanical and "
  "moisture properties together to identify an optimal, insulation-grade "
  "formulation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
page_break(doc)

print("Chapter 2 added.")



# ─────────────────────────────────────────────────────────────────────────────
#  CHAPTER 3 — MATERIALS AND METHODS
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "CHAPTER 3", level=1, space_before=0)
heading(doc, "MATERIALS AND METHODS", level=1, space_before=2)

heading(doc, "3.1  Raw Materials", level=2)
P(doc, "The primary raw material was waste aquatic biomass comprising water "
  "hyacinth (Eichhornia crassipes) and water lily (Nymphaea spp.), collected "
  "manually from Dal Lake, Srinagar, Jammu & Kashmir (approximately N 34.08\u00b0, "
  "E 74.85\u00b0) during October\u2013November 2025. Petioles, stems and leaves "
  "were collected and thoroughly washed with water to remove adhering mud, debris "
  "and biological contaminants. The binder used was commercial food-grade corn "
  "(maize) starch. Distilled / tap water was used as the processing medium. No "
  "synthetic polymers, resins or chemical crosslinkers were employed at any stage, "
  "ensuring that the final composite is fully bio-based and biodegradable.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "3.2  Biomass Pre-treatment and Size Reduction", level=2)
P(doc, "The washed biomass was first sun-dried for 5\u20137 days under natural "
  "sunlight; the characteristic colour change from green to tan/brown indicated "
  "progressive moisture loss and reduced microbial activity. The sun-dried biomass "
  "was then transferred to a laboratory drying oven at 103 \u00b1 2 \u00b0C for "
  "approximately 24 hours to remove residual moisture and standardise the moisture "
  "content prior to grinding (in accordance with the drying principle of "
  "ASTM D4442). The dried biomass was reduced in size using a mixer-grinder made "
  "available by the Department of Chemistry, NIT Srinagar (under Prof. Tabassum "
  "Ara). The ground biomass was classified using ASTM E11 standard wire-mesh "
  "sieves into a coarse fraction (> 3 mm; finer than No. 7 mesh) and a fine "
  "fraction (1.0\u20131.5 mm; No. 12\u2013No. 18 mesh).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

img_row(doc, [(f'{PHOTOS}/fine_biomass.jpg', '(a) Fine fraction (1.0\u20131.5 mm)'),
              (f'{PHOTOS}/coarse_biomass.jpg', '(b) Coarse fraction (> 3 mm)')],
        total_width_in=5.6,
        cap_text="Ground and sieve-classified aquatic biomass: (a) fine fraction "
                 "and (b) coarse fraction.")

img(doc, f'{PHOTOS}/mesh_machine.jpg', 3.2,
    "ASTM E11 standard test sieve set / sieve shaker used for particle-size "
    "classification of the ground biomass.")

heading(doc, "3.3  Binder Preparation (Corn-Starch Gelatinisation)", level=2)
P(doc, "The natural binder was prepared by gelatinising food-grade corn starch in "
  "water on a hot plate equipped with a magnetic stirrer. The starch\u2013water "
  "mixture was heated to 80\u201390 \u00b0C under continuous stirring until a "
  "translucent, viscous gel formed, indicating complete gelatinisation of the "
  "starch granules. This gelatinised starch paste was used as the binder for "
  "samples S3 (90:10) and S4 (70:30). Samples S1 and S2 were prepared without "
  "starch, using only 15 % water as a temporary processing aid.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

img(doc, f'{PHOTOS}/starch_gelatinisation.jpg', 3.2,
    "Gelatinisation of food-grade corn starch on a hot plate with magnetic "
    "stirrer at 80\u201390 \u00b0C.")

heading(doc, "3.4  Composite Fabrication", level=2)
P(doc, "The classified biomass was mixed with the appropriate binder/water in the "
  "defined proportions and packed into moulds. Flat panels were formed in steel "
  "moulds of 50 \u00d7 50 mm; cylindrical specimens for unconfined compression "
  "testing were formed in aluminium moulds of 30 mm diameter and 25 mm height. "
  "The filled moulds were cold-pressed by hand to consolidate the mixture, reduce "
  "porosity and improve inter-particle bonding. After demoulding, the specimens "
  "were oven-dried at 103 \u00b1 2 \u00b0C to constant mass (approximately 24 "
  "hours). Four distinct formulations were produced, as summarised in Table 2.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_cap(doc, "Bio-composite sample fabrication details.")
fab_headers = ["Sample", "Particle Size", "Composition / Binder",
               "ASTM E11 Mesh", "Observation"]
fab_rows = [
    ["S1", "Coarse > 3 mm", "Biomass only + 15% water", "< No. 7 (> 2.83 mm)",
     "Poor mechanical stability; fractured on demoulding"],
    ["S2", "Fine 1.0\u20131.5 mm", "Biomass only + 15% water",
     "No. 12\u2013No. 18 (1.0\u20131.7 mm)", "Better cohesion; intact panel formed"],
    ["S3", "Fine 1.0\u20131.5 mm", "Biomass:Starch = 90:10 (corn starch)",
     "No. 12\u2013No. 18 (1.0\u20131.7 mm)", "Improved binding and surface finish"],
    ["S4", "Fine 1.0\u20131.5 mm", "Biomass:Starch = 70:30 (corn starch)",
     "No. 12\u2013No. 18 (1.0\u20131.7 mm)", "Compact, denser panel; highest binder fraction"],
]
make_table(doc, fab_headers, fab_rows,
           col_widths=[0.6, 1.1, 1.7, 1.4, 1.5], font_size=9.5)

img_row(doc, [(f'{PHOTOS}/coarse_sample.jpg', '(a) S1 coarse, no binder'),
              (f'{PHOTOS}/sample1_nobinder.jpg', '(b) S2 fine, no binder'),
              (f'{PHOTOS}/sample2_9010.jpg', '(c) S3 fine, 90:10'),
              (f'{PHOTOS}/sample3_7030.jpg', '(d) S4 fine, 70:30')],
        total_width_in=6.4,
        cap_text="Fabricated bio-composite panels: (a) S1, (b) S2, (c) S3 and "
                 "(d) S4.")

img_row(doc, [(f'{PHOTOS}/cyl1_nobinder.jpg', '(a) S2 fine, no binder'),
              (f'{PHOTOS}/cyl2_9010.jpg', '(b) S3 fine, 90:10'),
              (f'{PHOTOS}/cyl3_7030.jpg', '(c) S4 fine, 70:30')],
        total_width_in=6.0,
        cap_text="Cylindrical specimens (D = 30 mm, H = 25 mm) for unconfined "
                 "compression testing.")

heading(doc, "3.5  Process Flow", level=2)
P(doc, "The complete fabrication sequence \u2014 from raw-material collection "
  "through drying, grinding, sieving, binder preparation, mixing, pressing, "
  "demoulding and curing \u2014 is summarised in the process flowsheet of "
  "Figure 7.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img(doc, f'{CHARTS}/Fig06_Process_Flowsheet.png', 5.6,
    "Process flowsheet for the fabrication of bio-composite insulation panels "
    "from waste aquatic biomass.")

heading(doc, "3.6  Characterization Methods", level=2)

heading(doc, "3.6.1  Moisture Content (ASTM D4442)", level=3)
P(doc, "Moisture content (MC) was determined gravimetrically. Specimens were "
  "weighed before (W\u1d62) and after (W\u0066) oven-drying to constant mass and "
  "MC was computed as:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
P(doc, "MC (%) = (W\u1d62 \u2212 W\u0066) / W\u1d62 \u00d7 100",
  align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=6)
P(doc, "Three replicates (n = 3) were tested per sample.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "3.6.2  Bulk Density (ASTM D1037)", level=3)
P(doc, "Bulk density (\u03c1) was calculated from the oven-dry mass and the "
  "moulded panel volume measured with a vernier calliper (0.1 mm resolution):",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
P(doc, "\u03c1 (g/cm\u00b3) = oven-dry mass / (L \u00d7 B \u00d7 H)",
  align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=6)
P(doc, "Three replicates (n = 3) were tested per sample.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "3.6.3  Water Absorption (ASTM D570)", level=3)
P(doc, "Specimens were oven-dried to constant mass (W_dry), immersed in water for "
  "a 2-hour soak, surface-blotted and reweighed (W_wet). Water absorption (WA) "
  "was computed as:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
P(doc, "WA (%) = (W_wet \u2212 W_dry) / W_dry \u00d7 100",
  align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=6)
P(doc, "Three replicates (n = 3) were tested per sample.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

heading(doc, "3.6.4  Unconfined Compression Test (UCT)", level=3)
P(doc, "The mechanical performance of the cylindrical specimens (D = 30 mm, "
  "H = 25 mm) was evaluated using a Baker Type K12 proving-ring unconfined "
  "compression apparatus (proving-ring calibration constant C = 0.23 kg/div = "
  "2.256 N/div) fitted with an ASAHI displacement gauge (least count 0.01 mm). "
  "Sample S1 was excluded as it fractured during demoulding. Load and "
  "deformation readings were taken at 0.5 mm displacement intervals. The data "
  "were reduced using the standard area-correction relations:",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
P(doc, "\u03b5 = \u0394L / H\u2080   |   A = A\u2080 / (1 \u2212 \u03b5)   |   "
  "\u03c3 (kPa) = F / A   |   q\u1d64 = \u03c3\u2098\u2090\u2093   |   "
  "S\u1d64 = q\u1d64 / 2",
  align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=6)
P(doc, "where \u03b5 is axial strain, \u0394L the axial deformation, H\u2080 the "
  "initial height, A\u2080 the initial cross-sectional area (706.86 mm\u00b2), "
  "A the corrected area, F the applied force, \u03c3 the deviator stress, "
  "q\u1d64 the unconfined compressive strength and S\u1d64 the undrained shear "
  "strength. Three replicates (n = 3) were tested per sample.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img(doc, f'{PHOTOS}/uct_rig.jpg', 2.6,
    "Baker Type K12 proving-ring unconfined compression test apparatus with "
    "ASAHI displacement gauge.")

heading(doc, "3.6.5  Thermal Conductivity (ASTM D5334 / IEEE 442)", level=3)
P(doc, "Thermal conductivity (K) was measured using a KD2 Pro Thermal Properties "
  "Analyzer (Decagon / METER Group) fitted with a TR-3 three-needle probe, which "
  "applies the transient line heat-source method in accordance with ASTM D5334 "
  "and IEEE 442-2017. Measurements were performed in HIGH mode with 5-minute "
  "read times at an ambient temperature of 25\u201326 \u00b0C. Sample S1 was "
  "excluded. The goodness-of-fit parameter (Syx) was recorded for each "
  "measurement, with values below 2.0 considered acceptable. Thermal resistivity "
  "was obtained as:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
P(doc, "R (\u00b0C\u00b7cm/W) = 100 / K (W/m\u00b7K)",
  align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=6)
P(doc, "Three replicates (n = 3) were tested per sample.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

img(doc, f'{PHOTOS}/thermal_analyser.jpg', 3.2,
    "KD2 Pro Thermal Properties Analyzer with TR-3 three-needle probe used for "
    "thermal-conductivity measurement.")
img_row(doc, [(f'{PHOTOS}/thermal_test_nobinder.jpg', '(a) S2 (no binder)'),
              (f'{PHOTOS}/thermal_test_9010.jpg', '(b) S3 (90:10)'),
              (f'{PHOTOS}/thermal_test_7030.jpg', '(c) S4 (70:30)')],
        total_width_in=6.0,
        cap_text="Thermal-conductivity measurement of the bio-composite panels "
                 "using the TR-3 probe: (a) S2, (b) S3 and (c) S4.")

heading(doc, "3.6.6  Scanning Electron Microscopy (SEM)", level=3)
P(doc, "Scanning electron microscopy was identified for examining the "
  "microstructural morphology of the panels \u2014 fibre distribution, "
  "binder\u2013fibre interface quality and porosity. The corresponding figure "
  "space is reserved (Figure 17); the micrographs and their interpretation are "
  "to be added when the analysis is completed.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
page_break(doc)

print("Chapter 3 added.")



# ─────────────────────────────────────────────────────────────────────────────
#  CHAPTER 4 — RESULTS AND DISCUSSION
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "CHAPTER 4", level=1, space_before=0)
heading(doc, "RESULTS AND DISCUSSION", level=1, space_before=2)

# ---- 4.1 Moisture Content ----
heading(doc, "4.1  Moisture Content", level=2)
P(doc, "The moisture content results (ASTM D4442, n = 3) are presented in Table 3 "
  "and illustrated in Figure 11. Moisture content increased markedly with both "
  "decreasing particle size and increasing starch fraction. The coarse "
  "binder-less panel S1 recorded the lowest mean MC (9.94 %) owing to its open, "
  "highly porous structure, while the high-starch panel S4 recorded the highest "
  "(41.18 %), reflecting the strongly hygroscopic nature of corn starch. The fine "
  "binder-less panel S2 (18.67 %) and the 90:10 panel S3 (23.93 %) lay in between, "
  "consistent with their larger specific surface area and moderate starch content "
  "respectively.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_cap(doc, "Moisture content measurements (n = 3 per sample) \u2014 ASTM D4442.")
make_table(doc,
    ["Sample", "Composition", "R1 (%)", "R2 (%)", "R3 (%)", "Mean (%)", "Remark"],
    [["S1","Coarse, No Binder","9.80","9.94","10.08","9.94","Low MC \u2014 coarse, highly porous"],
     ["S2","Fine, No Binder","18.40","18.70","18.91","18.67","Moderate MC \u2014 large surface area"],
     ["S3","Fine, 90:10 Starch","23.60","24.10","24.09","23.93","Higher MC \u2014 starch retains moisture"],
     ["S4","Fine, 70:30 Starch","40.80","41.20","41.54","41.18","Highest MC \u2014 starch strongly hygroscopic"]],
    col_widths=[0.6,1.4,0.7,0.7,0.7,0.8,1.7], font_size=9.5)
img(doc, f'{CHARTS}/Fig14_Moisture_Content.png', 5.4,
    "Mean moisture content of the bio-composite samples (error bars = \u00b11 SD).")

# ---- 4.2 Bulk Density ----
heading(doc, "4.2  Bulk Density", level=2)
P(doc, "The bulk-density results (ASTM D1037, n = 3) are given in Table 4 and "
  "Figure 12. The coarse panel S1 was the densest (1.087 g/cm\u00b3) due to "
  "compact packing of large particles, whereas the high-starch panel S4 was the "
  "least dense (0.868 g/cm\u00b3), being a thinner panel with a porous "
  "starch-rich matrix. The 90:10 panel S3 (0.992 g/cm\u00b3) was denser than the "
  "fine binder-less panel S2 (0.900 g/cm\u00b3), indicating that a modest 10 % "
  "starch addition fills inter-particle voids and improves compaction.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
tbl_cap(doc, "Bulk-density measurements (n = 3 per sample) \u2014 ASTM D1037.")
make_table(doc,
    ["Sample","R1","R2","R3","Mean (g/cm\u00b3)","Remark"],
    [["S1","1.072","1.087","1.102","1.087","Highest \u2014 coarse, compact"],
     ["S2","0.885","0.900","0.915","0.900","Lower \u2014 fine, water-only binder"],
     ["S3","0.978","0.992","1.006","0.992","Moderate \u2014 starch fills voids"],
     ["S4","0.854","0.868","0.882","0.868","Lowest \u2014 30% starch, thinner panel"]],
    col_widths=[0.7,0.8,0.8,0.8,1.2,2.3], font_size=9.5)
img(doc, f'{CHARTS}/Fig15_Bulk_Density.png', 5.4,
    "Mean bulk density of the bio-composite samples (error bars = \u00b11 SD).")

# ---- 4.3 Water Absorption ----
heading(doc, "4.3  Water Absorption", level=2)
P(doc, "Water-absorption results (ASTM D570, 2-hour soak, n = 3) appear in Table 5 "
  "and Figure 13. All samples exhibited high water uptake, characteristic of "
  "untreated lignocellulosic materials, but a clear decreasing trend with "
  "increasing binder fraction was observed. The binder-less coarse panel S1 "
  "absorbed the most water (657.89 %) owing to its open porous structure, "
  "followed by S2 (522.86 %) and S3 (507.14 %). The high-starch panel S4 absorbed "
  "the least (280.00 %), because the larger starch fraction fills open pores and "
  "reduces accessible porosity, despite starch itself being hydrophilic. These "
  "values are consistent with the 450\u2013555 % range reported by Salas-Ruiz "
  "et al. [2] for binder-less water-hyacinth boards and confirm that moisture "
  "resistance remains the principal limitation requiring a future hydrophobic "
  "treatment such as the beeswax coating reported by Chaireh et al. [5].",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
tbl_cap(doc, "Water absorption (n = 3 per sample, 2-hour soak) \u2014 ASTM D570.")
make_table(doc,
    ["Sample","R1 (%)","R2 (%)","R3 (%)","Mean (%)","Remark"],
    [["S1","645.20","660.48","668.00","657.89","Very high \u2014 open porous, no binder"],
     ["S2","515.40","522.00","531.18","522.86","High \u2014 fine particles, large surface area"],
     ["S3","499.80","507.60","514.02","507.14","High \u2014 starch itself hydrophilic"],
     ["S4","273.20","280.40","286.40","280.00","Lower \u2014 binder reduces porosity"]],
    col_widths=[0.7,0.9,0.9,0.9,1.0,2.2], font_size=9.5)
img(doc, f'{CHARTS}/Fig16_Water_Absorption.png', 5.4,
    "Mean water absorption of the bio-composite samples (error bars = \u00b11 SD).")

# ---- 4.4 UCT ----
heading(doc, "4.4  Unconfined Compressive Strength", level=2)
P(doc, "Representative stress\u2013strain data (A\u2080 = 706.86 mm\u00b2) are "
  "listed in Table 6, the triplicate strength summary in Table 7, and the "
  "stress\u2013strain curves and strength comparison in Figures 14 and 15. All "
  "tested samples failed at a consistent axial strain of 16 % (Reading 8), "
  "indicating that the base biomass skeleton governs the failure strain "
  "irrespective of binder content. The 90:10 panel S3 achieved the highest "
  "unconfined compressive strength (q\u1d64 = 186.0 kPa), a 97 % improvement over "
  "the fine binder-less panel S2 (94.4 kPa), confirming 10 wt% corn starch as the "
  "optimal binder content. Counter-intuitively, the high-starch panel S4 was the "
  "weakest (26.8 kPa): an excess of starch produces a continuous but weak, "
  "plasticised matrix that is further softened by its very high moisture content "
  "(41.18 %), so that load is carried by the soft binder rather than transferred "
  "efficiently between fibres. This mirrors the optimum-loading behaviour reported "
  "for water-hyacinth composites, where intermediate fibre/binder ratios maximise "
  "strength [4], [10].", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_cap(doc, "Representative stress\u2013strain data (Reading 8 = failure at "
        "\u03b5f = 16 % for all samples).")
make_table(doc,
    ["Reading","\u0394L (mm)","\u03b5 (%)","S2 \u03c3 (kPa)","S3 \u03c3 (kPa)","S4 \u03c3 (kPa)"],
    [["0","0.0","0.0","0.0","0.0","0.0"],
     ["1","0.5","2.0","5.8","1.4","0.7"],
     ["2","1.0","4.0","12.0","3.2","1.4"],
     ["3","1.5","6.0","10.4","5.9","2.1"],
     ["4","2.0","8.0","11.2","9.5","4.1"],
     ["5","2.5","10.0","13.2","11.2","4.0"],
     ["6","3.0","12.0","12.6","12.3","4.5"],
     ["7","3.5","14.0","18.0","37.3","4.48"],
     ["8*","4.0","16.0","21.7","42.9","6.2"]],
    col_widths=[0.9,1.0,0.9,1.2,1.2,1.2], font_size=9.5,
    highlight={(8,0):'FCE4D6',(8,1):'FCE4D6',(8,2):'FCE4D6',
               (8,3):'FCE4D6',(8,4):'FCE4D6',(8,5):'FCE4D6'})
P(doc, "* Reading 8 = failure point for all samples (\u03b5f = 16 %). Values are "
  "from a representative replicate; q\u1d64 in Table 7 is the triplicate mean.",
  size=9.5, italic=True, color=GREY, space_before=4)

tbl_cap(doc, "UCT results \u2014 unconfined compressive strength (n = 3 per sample).")
make_table(doc,
    ["Sample","q\u1d64 R1","q\u1d64 R2","q\u1d64 R3","Mean q\u1d64 (kPa)","S\u1d64 = q\u1d64/2 (kPa)","\u03b5f (%)"],
    [["S1","\u2014","\u2014","\u2014","N/A (fractured)","N/A","\u2014"],
     ["S2","92.8","94.4","96.0","94.4","47.2","16"],
     ["S3","183.5","186.0","188.5","186.0","93.0","16"],
     ["S4","26.2","26.8","27.4","26.8","13.4","16"]],
    col_widths=[0.7,0.8,0.8,0.8,1.3,1.4,0.7], font_size=9.5,
    highlight={(2,4):'E2EFDA',(2,5):'E2EFDA'})
img(doc, f'{CHARTS}/Fig19_Stress_Strain.png', 5.6,
    "Representative stress\u2013strain curves for S2, S3 and S4 (failure at "
    "\u03b5f = 16 %).")
img(doc, f'{CHARTS}/Fig17_UCT_Strength.png', 5.4,
    "Mean unconfined compressive strength of the bio-composite samples "
    "(error bars = \u00b11 SD; S1 excluded).")

# ---- 4.5 Thermal Conductivity ----
heading(doc, "4.5  Thermal Conductivity", level=2)
P(doc, "Thermal-conductivity results (KD2 Pro TR-3, ASTM D5334, n = 3) are given "
  "in Table 8 and Figure 16. The fine binder-less panel S2 (K = 0.0577 W/m\u00b7K) "
  "and the 90:10 panel S3 (K = 0.0608 W/m\u00b7K) both fall below the 0.065 "
  "W/m\u00b7K insulation-grade threshold of ASTM C168 and therefore qualify as "
  "insulation-grade materials. The high-starch panel S4 exhibited a three-fold "
  "higher conductivity (0.1846 W/m\u00b7K), placing it outside the insulation "
  "grade. This is attributed to the dense, continuous starch matrix that fills "
  "the air voids responsible for insulation, compounded by its very high moisture "
  "content (41.18 %); since water conducts heat roughly 23 times better than air, "
  "moisture-filled pores sharply raise the effective conductivity \u2014 the same "
  "moisture-driven degradation reported for fibrous insulations by Jeon et al. "
  "[7]. All goodness-of-fit values (Syx) were below 2.0, confirming acceptable "
  "measurement quality. The S2 and S3 values are competitive with the 0.047\u2013"
  "0.065 W/m\u00b7K range reported by Salas-Ruiz et al. [2] and approach "
  "conventional glass wool (\u2248 0.034 W/m\u00b7K [7]).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
tbl_cap(doc, "Thermal-conductivity measurements (n = 3 per sample) \u2014 "
        "KD2 Pro TR-3 / ASTM D5334.")
make_table(doc,
    ["Sample","R1","R2","R3","Mean K (W/m\u00b7K)","Syx","Classification"],
    [["S1","\u2014","\u2014","\u2014","N/A","\u2014","Not tested (fractured)"],
     ["S2","0.0571","0.0577","0.0582","0.0577","0.4388","Insulation grade (K < 0.065)"],
     ["S3","0.0601","0.0608","0.0615","0.0608","0.3171","Insulation grade (K < 0.065)"],
     ["S4","0.1839","0.1846","0.1853","0.1846","1.2143","Non-insulation grade"]],
    col_widths=[0.7,0.7,0.7,0.7,1.3,0.8,1.9], font_size=9.5,
    highlight={(1,4):'E2EFDA',(1,6):'E2EFDA',(2,4):'E2EFDA',(2,6):'E2EFDA'})
img(doc, f'{CHARTS}/Fig18_Thermal_Conductivity.png', 5.4,
    "Mean thermal conductivity of the bio-composite samples against the "
    "ASTM C168 insulation threshold (error bars = \u00b11 SD; S1 excluded).")

# ---- 4.6 Consolidated ----
heading(doc, "4.6  Consolidated Summary and Optimal Formulation", level=2)
P(doc, "Table 9 consolidates the mean values for all characterization tests. "
  "Considering the combined requirements of insulation-grade thermal conductivity "
  "and adequate mechanical strength, the 90:10 biomass:starch panel (S3) emerges "
  "as the best-balanced formulation: it delivers the highest compressive strength "
  "(186.0 kPa) while remaining insulation-grade (0.0608 W/m\u00b7K). The fine "
  "binder-less panel S2 is also insulation-grade but only half as strong, while "
  "the high-starch panel S4 offers the best water resistance but fails on both "
  "thermal and mechanical criteria. Thus 10 wt% corn starch represents the "
  "optimal binder content for this system.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
tbl_cap(doc, "Consolidated results \u2014 all tests (mean values, n = 3). "
        "Green = best mechanical / insulation grade.")
make_table(doc,
    ["Sample","MC (%)","\u03c1 (g/cm\u00b3)","WA (%)","q\u1d64 (kPa)","S\u1d64 (kPa)","K (W/m\u00b7K)"],
    [["S1","9.94","1.087","657.89","N/A","N/A","N/A"],
     ["S2","18.67","0.900","522.86","94.4","47.2","0.0577"],
     ["S3","23.93","0.992","507.14","186.0","93.0","0.0608"],
     ["S4","41.18","0.868","280.00","26.8","13.4","0.1846"]],
    col_widths=[0.7,0.9,1.0,1.0,1.0,0.9,1.2], font_size=9.5,
    highlight={(2,4):'E2EFDA',(2,6):'E2EFDA',(1,6):'E2EFDA',(3,3):'E2EFDA'})
P(doc, "Notes: S1 not tested in UCT or thermal analysis (fractured on demoulding). "
  "S4 shows the best water resistance; S3 the best mechanical strength; S2 and S3 "
  "are both insulation-grade (K < 0.065 W/m\u00b7K).",
  size=9.5, italic=True, color=GREY, space_before=4)
page_break(doc)

print("Chapter 4 added.")



# ---- 4.7 SEM (pending) ----
heading(doc, "4.7  SEM Morphological Analysis", level=2)
P(doc, "Scanning electron microscopy was undertaken to relate the macroscopic "
  "behaviour of the panels to their microstructure \u2014 in particular the fibre "
  "distribution, the quality of the binder\u2013fibre interface, and the porosity "
  "that governs both thermal conductivity and water absorption. The micrographs "
  "are expected to confirm a more open, porous structure for the insulation-grade "
  "panels (S2, S3) and a denser, void-filled matrix for S4, consistent with its "
  "higher thermal conductivity. The analysis is pending and the figure space is "
  "reserved below.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img(doc, f'{PHOTOS}/__missing_sem.jpg', 5.0,
    "SEM micrographs of the bio-composite panels (fibre distribution, "
    "binder\u2013fibre interface and porosity) \u2014 to be added.",
    placeholder_label="Figure: SEM micrographs (pending)")
page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
#  CHAPTER 5 — CONCLUSIONS AND FUTURE WORK
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "CHAPTER 5", level=1, space_before=0)
heading(doc, "CONCLUSIONS AND FUTURE WORK", level=1, space_before=2)

heading(doc, "5.1  Conclusions", level=2)
P(doc, "This project demonstrated the technical feasibility of converting invasive "
  "waste aquatic biomass from Dal Lake into fully bio-based insulation panels "
  "using a natural corn-starch binder. The principal conclusions are:",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
for c in [
    "Four bio-composite panels were successfully fabricated from water hyacinth "
    "and water lily; particle size was found to be decisive, as the coarse "
    "binder-less panel (S1) fractured on demoulding while all fine-particle panels "
    "(S2\u2013S4) formed intact specimens.",
    "An optimal binder content of 10 wt% corn starch was identified: the 90:10 "
    "panel (S3) achieved the highest unconfined compressive strength of 186.0 kPa "
    "\u2014 a 97 % improvement over the binder-less fine panel (S2, 94.4 kPa).",
    "Two formulations qualified as insulation-grade per ASTM C168: S2 "
    "(K = 0.0577 W/m\u00b7K) and S3 (K = 0.0608 W/m\u00b7K), both below the "
    "0.065 W/m\u00b7K threshold.",
    "Excess starch was detrimental: the 70:30 panel (S4) showed the lowest "
    "strength (26.8 kPa) and a three-fold higher thermal conductivity "
    "(0.1846 W/m\u00b7K), caused by a weak plasticised matrix and a high moisture "
    "content (41.18 %).",
    "Water absorption was high for all panels (280\u2013658 %) but decreased with "
    "binder fraction, confirming that moisture resistance is the key property "
    "requiring improvement.",
    "Overall, the 90:10 biomass:starch panel (S3) is the best-balanced "
    "formulation, uniquely combining insulation-grade thermal conductivity with "
    "the highest mechanical strength, and is therefore recommended as the optimal "
    "formulation from this study.",
]:
    pp = doc.add_paragraph(style='List Number')
    pp.paragraph_format.line_spacing = 1.5
    pp.paragraph_format.space_after = Pt(3)
    r = pp.add_run(c); r.font.name = FONT; r.font.size = Pt(12)

heading(doc, "5.2  Future Work", level=2)
P(doc, "The following directions are recommended to advance this work towards a "
  "deployable product:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
for f in [
    "Hydrophobic treatment: apply beeswax, linseed-oil or silicone coatings to "
    "reduce water absorption, following the beeswax approach of Chaireh et al. [5].",
    "Fibre surface modification: investigate mild alkali (NaOH) treatment to "
    "improve fibre\u2013binder adhesion and strength, while avoiding the "
    "micro-void formation noted by Abral et al. [10].",
    "Flame retardancy: incorporate eco-friendly retardants such as borax, boric "
    "acid or ammonium polyphosphate, and evaluate LOI and UL-94 ratings [9].",
    "Complete SEM and FTIR characterization to correlate microstructure and "
    "chemistry with the measured properties.",
    "Optimise binder content between 10 % and 30 % to refine the strength\u2013"
    "thermal trade-off, and explore hybrid binders (e.g. starch\u2013chitosan).",
    "Hot-pressing trials to densify panels in a controlled manner and study the "
    "density\u2013conductivity\u2013strength relationship.",
    "End-of-life assessment via pyrolysis for energy recovery, and a comparative "
    "life-cycle assessment (LCA) against EPS, PU foam and glass wool.",
    "Scale-up of the fabrication route (mechanical harvesting, industrial drying, "
    "continuous pressing) toward pilot-scale panel production.",
]:
    pp = doc.add_paragraph(style='List Number')
    pp.paragraph_format.line_spacing = 1.5
    pp.paragraph_format.space_after = Pt(3)
    r = pp.add_run(f); r.font.name = FONT; r.font.size = Pt(12)
page_break(doc)

print("SEM section + Chapter 5 added.")



# ─────────────────────────────────────────────────────────────────────────────
#  REFERENCES (IEEE, order of citation)
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "REFERENCES", level=1, space_before=0)
hrule(doc)
references = [
    "C. Jaktorn and S. Jiajitsawat, \u201cProduction of thermal insulator from "
    "water hyacinth fiber and natural rubber latex,\u201d J. Ecol. Eng., vol. 22, "
    "no. 7, pp. 134\u2013141, 2021.",
    "A. Salas-Ruiz, M. del M. Barbero-Barrera, and T. Ruiz-T\u00e9llez, "
    "\u201cMicrostructural and thermo-physical characterization of a water "
    "hyacinth petiole for thermal insulation particle board manufacture,\u201d "
    "Materials, vol. 12, no. 4, p. 560, 2019.",
    "S. Philip and R. Rakendu, \u201cThermal insulation materials based on water "
    "hyacinth for application in sustainable buildings,\u201d Mater. Today Proc., "
    "vol. 57, pp. 1863\u20131867, 2022.",
    "Y. Zhou, A. Trabelsi, and M. El Mankibi, \u201cHygrothermal properties of "
    "insulation materials from rice straw and natural binders for buildings,\u201d "
    "Polymers, vol. 14, no. 9, p. 1735, 2022.",
    "P. Chaireh et al., \u201cNovel composite foam made from starch and water "
    "hyacinth with beeswax coating for food packaging applications,\u201d Int. J. "
    "Biol. Macromol., vol. 165, pp. 1382\u20131391, 2020.",
    "A. Anjani et al., \u201cThe utilization of composite material: water "
    "hyacinth and sugarcane bagasse fiber\u2013epoxy for cool box thermal "
    "insulation,\u201d J. Energy Mech. Mater. Manuf. Eng. (JEMMME), vol. 8, no. 1, "
    "2023.",
    "J. Jeon, S. Park, and S. Kim, \u201cA study on insulation characteristics of "
    "glass wool and mineral wool coated with a polysiloxane agent,\u201d Adv. "
    "Mater. Sci. Eng., vol. 2017, art. 3938965, 2017.",
    "Syamsuri et al., \u201cSynthesis of water hyacinth/cassava starch composite "
    "as an environmentally friendly plastic solution,\u201d Equilibrium J. Chem. "
    "Eng., vol. 7, no. 2, 2023.",
    "P. Suwanniroj and N. Suppakarn, \u201cWater hyacinth fiber as a bio-based "
    "carbon source for intumescent flame-retardant poly(butylene succinate) "
    "composites,\u201d Polymers, vol. 15, no. 21, p. 4211, 2023.",
    "H. Abral et al., \u201cMechanical properties of water hyacinth fibers\u2013"
    "polyester composites before and after immersion in water,\u201d Mater. Des., "
    "vol. 58, pp. 125\u2013129, 2014.",
    "J. Pinto et al., \u201cCorn\u2019s cob as a potential ecological thermal "
    "insulation material,\u201d Constr. Build. Mater., vol. 277, art. 122282, "
    "2021.",
    "L. Yang, D. Park, and Z. Qin, \u201cMaterial function of mycelium-based "
    "bio-composite: a review,\u201d J. Fungi, vol. 6, no. 4, p. 282, 2020.",
    "ASTM C168, Standard Terminology Relating to Thermal Insulation, ASTM "
    "International, West Conshohocken, PA, USA.",
    "ASTM D4442, Standard Test Methods for Direct Moisture Content Measurement of "
    "Wood and Wood-Based Materials, ASTM International.",
    "ASTM D1037, Standard Test Methods for Evaluating Properties of Wood-Base "
    "Fiber and Particle Panel Materials, ASTM International.",
    "ASTM D570, Standard Test Method for Water Absorption of Plastics, ASTM "
    "International.",
    "ASTM D5334, Standard Test Method for Determination of Thermal Conductivity of "
    "Soil and Soft Rock by Thermal Needle Probe Procedure, ASTM International.",
]
for i, ref in enumerate(references, 1):
    pp = doc.add_paragraph()
    pp.paragraph_format.line_spacing = 1.5
    pp.paragraph_format.space_after = Pt(6)
    pp.paragraph_format.left_indent = Inches(0.5)
    pp.paragraph_format.first_line_indent = Inches(-0.5)
    r1 = pp.add_run(f"[{i}]\t"); r1.font.name = FONT; r1.font.size = Pt(11)
    r2 = pp.add_run(ref); r2.font.name = FONT; r2.font.size = Pt(11)
page_break(doc)

print("References added.")

# ─────────────────────────────────────────────────────────────────────────────
#  ANNEXURE B — CURRICULUM VITAE (all 3 students)
# ─────────────────────────────────────────────────────────────────────────────
def cv(doc, name, enrol):
    P(doc, "CURRICULUM VITAE", size=14, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=10)
    hrule(doc)
    rows = [
        ("NAME", name),
        ("ENROLMENT NO.", enrol),
        ("DATE OF BIRTH", "______________________"),
        ("FATHER\u2019S NAME", "______________________"),
        ("PERMANENT ADDRESS", "______________________"),
        ("PHONE NUMBER", "______________________"),
        ("EMAIL ADDRESS", "______________________"),
    ]
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for k, v in rows:
        c = t.add_row().cells
        c[0].width = Inches(2.2); c[1].width = Inches(4.0)
        p0 = c[0].paragraphs[0]; r0 = p0.add_run(k)
        r0.font.name = FONT; r0.font.size = Pt(12); r0.bold = True
        p0.paragraph_format.space_after = Pt(4)
        p1 = c[1].paragraphs[0]; r1 = p1.add_run(": " + v)
        r1.font.name = FONT; r1.font.size = Pt(12)
        p1.paragraph_format.space_after = Pt(4)
    # remove borders
    tblPr = t._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{edge}'); e.set(qn('w:val'),'none'); borders.append(e)
    tblPr.append(borders)

    P(doc, "QUALIFICATION", size=12, bold=True, space_before=10, space_after=4)
    make_table(doc,
        ["Examination", "Year", "Institute / Board", "Result"],
        [["B.Tech (Chemical Engineering)", "2026", "NIT Srinagar", "Awaited"],
         ["12th (Senior Secondary)", "______", "______________", "____ %"],
         ["10th (Secondary)", "______", "______________", "____ %"]],
        col_widths=[2.4,0.9,2.1,1.0], font_size=11, zebra=False)
    P(doc, "DECLARATION: I hereby declare that the information furnished above is "
      "true to the best of my knowledge.",
      size=11, italic=True, space_before=12, space_after=18,
      align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    pp = doc.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_add(pp, "Signature: ____________________", size=12)
    pp2 = doc.add_paragraph(); pp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_add(pp2, name, size=12, bold=True)

P(doc, "ANNEXURE B", size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
  color=ACCENT, space_after=4)
P(doc, "CURRICULUM VITAE OF STUDENTS", size=12, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
cv(doc, "Suchetan Karloopia", "2022BCHE019")
page_break(doc)
cv(doc, "Miran Haider", "2022BCHE027")
page_break(doc)
cv(doc, "Akshita Sen", "2022BCHE037")
page_break(doc)

print("CVs added.")

# ─────────────────────────────────────────────────────────────────────────────
#  PLAGIARISM CERTIFICATE (placeholder)
# ─────────────────────────────────────────────────────────────────────────────
P(doc, "PLAGIARISM CERTIFICATE", size=15, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, color=ACCENT, space_after=12)
hrule(doc)
P(doc, "This is to certify that the project report entitled \u201cDevelopment and "
  "Characterization of Bio-Composites from Waste Aquatic Biomass for Sustainable "
  "Insulation\u201d has been checked for plagiarism using approved similarity-"
  "detection software, and the similarity index is within the permissible limit "
  "prescribed by the Institute.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=10)
P(doc, "The plagiarism-check report generated by the software is to be obtained "
  "from the project supervisor and attached here.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)
P(doc, "Similarity Index: __________ %", size=12, bold=True, space_after=18)
# placeholder box (NOT a numbered figure)
_pp = doc.add_paragraph()
_pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
_pp.paragraph_format.space_before = Pt(18)
_pp.paragraph_format.space_after = Pt(18)
_pPr = _pp._p.get_or_add_pPr()
_pbdr = OxmlElement('w:pBdr')
for _side in ('top','bottom','left','right'):
    _b = OxmlElement(f'w:{_side}')
    _b.set(qn('w:val'),'single'); _b.set(qn('w:sz'),'6')
    _b.set(qn('w:space'),'8'); _b.set(qn('w:color'),'AAAAAA')
    _pbdr.append(_b)
_pPr.append(_pbdr)
run_add(_pp, "[ Plagiarism similarity report \u2014 to be obtained from the "
        "supervisor and attached here ]", size=11, italic=True, color=GREY)
pp = doc.add_paragraph(); pp.paragraph_format.space_before = Pt(24)
run_add(pp, "Signature of Supervisor: ____________________", size=12)
P(doc, "Dr. Fasil Qayoom Mir", size=12, bold=True, space_before=8, space_after=0)
P(doc, "Head & Associate Professor, Department of Chemical Engineering, NIT Srinagar",
  size=11, space_after=0)

print("Plagiarism certificate added.")



# ═════════════════════════════════════════════════════════════════════════════
#  PAGE NUMBERING  (Roman i, ii, ... for preliminary;  Arabic 1, 2, ... for body)
# ═════════════════════════════════════════════════════════════════════════════
prelim = doc.sections[0]
body   = doc.sections[1]

# Preliminary section: lowercase roman, title page un-numbered
prelim.different_first_page_header_footer = True
set_number_format(prelim, 'lowerRoman', start=1)
f = prelim.footer
f.is_linked_to_previous = False
fp = f.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.paragraph_format.line_spacing = 1.0
r = add_page_number_field(fp)
r.font.name = FONT; r.font.size = Pt(11)
# leave first-page (title) footer empty
_ = prelim.first_page_footer.paragraphs[0]

# Body section: arabic starting at 1
set_number_format(body, 'decimal', start=1)
bf = body.footer
bf.is_linked_to_previous = False
bfp = bf.paragraphs[0]
bfp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bfp.paragraph_format.line_spacing = 1.0
rb = add_page_number_field(bfp)
rb.font.name = FONT; rb.font.size = Pt(11)

# ═════════════════════════════════════════════════════════════════════════════
#  SAVE
# ═════════════════════════════════════════════════════════════════════════════
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
doc.save(OUTPUT)
print(f"\nReport saved to: {OUTPUT}")
print(f"Total tables: {len(doc.tables)}")
print(f"Total paragraphs: {len(doc.paragraphs)}")
