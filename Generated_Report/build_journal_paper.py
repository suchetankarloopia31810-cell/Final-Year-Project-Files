#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_journal_paper.py  –  REVISED COMPREHENSIVE VERSION
Generates JBE_Manuscript_BioComposites.docx

Journal of Building Engineering (Elsevier) submission format.
Project: "Valorisation of invasive aquatic weed from Dal Lake into
         corn-starch bonded bio-composite panels for sustainable building insulation"
NIT Srinagar – B.Tech Final Year Project

Improvements over v1:
  • Highlights block
  • Expanded introduction (4 sub-sections, 27 references)
  • Equipment photographs (Fig. 2), cylinder specimens (Fig. 4) in Methods
  • Thermal test photographs (Fig. 9) in Results
  • New SEM section 3.5 with all 11 micrographs (Figs. 11–12)
  • Expanded comparison section 3.6 with Table 5
  • Tables styled to match Elsevier/JBE format (gray header only, no colored data cells)
  • 27 references in Harvard author-year style
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE   = os.path.dirname(os.path.abspath(__file__))
CHARTS = os.path.join(BASE, 'charts')
PHOTOS = os.path.join(BASE, 'photos')
OUTDIR = os.path.join(BASE, 'Journal_Paper')
os.makedirs(OUTDIR, exist_ok=True)
OUTPUT = os.path.join(OUTDIR, 'JBE_Manuscript_BioComposites.docx')

FONT  = 'Times New Roman'
BLACK = RGBColor(0, 0, 0)
GREY  = RGBColor(0x55, 0x55, 0x55)

# ─────────────────────────────────────────── helpers ────────────────────────
def style_base(doc):
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(11)
    st._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)


def P(doc, text='', size=11, bold=False, italic=False, align=None,
      color=None, space_before=0, space_after=8, line=1.5, font=FONT):
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
        r.font.name = font; r.font.size = Pt(size)
        r.bold = bold; r.italic = italic
        if color is not None:
            r.font.color.rgb = color
    return p

def run_add(p, text, size=11, bold=False, italic=False, color=None, font=FONT):
    r = p.add_run(text)
    r.font.name = font; r.font.size = Pt(size)
    r.bold = bold; r.italic = italic
    if color is not None:
        r.font.color.rgb = color
    return r

def section(doc, text, size=12, space_before=14, space_after=6):
    """Numbered section / sub-section heading (bold, left-aligned)."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT; r.font.size = Pt(size); r.bold = True
    r.font.color.rgb = BLACK
    return p

def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def tbl_caption(doc, label, text, size=9.5):
    """
    Elsevier/JBE table caption:
      Line 1 – 'Table X' bold (label)
      Line 2 – description italic (text)
    Caption sits ABOVE the table.
    """
    cp1 = doc.add_paragraph()
    cp1.paragraph_format.line_spacing = 1.0
    cp1.paragraph_format.space_before = Pt(10)
    cp1.paragraph_format.space_after = Pt(1)
    r0 = cp1.add_run(label)
    r0.font.name = FONT; r0.font.size = Pt(size); r0.bold = True

    cp2 = doc.add_paragraph()
    cp2.paragraph_format.line_spacing = 1.0
    cp2.paragraph_format.space_after = Pt(4)
    r1 = cp2.add_run(text)
    r1.font.name = FONT; r1.font.size = Pt(size); r1.italic = True

def make_table(doc, headers, rows, col_widths=None, font_size=9.5,
               first_col_left=True):
    """
    Elsevier/JBE table style:
      • Gray header row (D9D9D9) with bold centered text
      • White data rows – no colored highlighting
      • 'Table Grid' style (thin border on all cells)
      • First column left-aligned if first_col_left=True; rest centered
    """
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'

    # Header row
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_bg(hdr[i], 'D9D9D9')
        para = hdr[i].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.line_spacing = 1.0
        para.paragraph_format.space_after = Pt(2)
        para.paragraph_format.space_before = Pt(2)
        r = para.add_run(h)
        r.font.name = FONT; r.font.size = Pt(font_size); r.bold = True
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Data rows – plain white, no highlights
    for row_data in rows:
        cells = t.add_row().cells
        for ci, val in enumerate(row_data):
            para = cells[ci].paragraphs[0]
            para.alignment = (WD_ALIGN_PARAGRAPH.LEFT
                              if first_col_left and ci == 0
                              else WD_ALIGN_PARAGRAPH.CENTER)
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.space_before = Pt(2)
            r = para.add_run(str(val))
            r.font.name = FONT; r.font.size = Pt(font_size)
            cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if col_widths:
        for row in t.rows:
            for ci, w in enumerate(col_widths):
                row.cells[ci].width = Inches(w)

    # Small gap after table
    gap = doc.add_paragraph()
    gap.paragraph_format.space_after = Pt(6)
    return t


def img(doc, path, width_in, label_bold, cap_text, ph=None):
    """Insert a single centred figure with its caption."""
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(path, width=Inches(width_in))
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(8)
        run_add(p, f'[ {ph or os.path.basename(path)} ]',
                size=10, italic=True, color=GREY)
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_after = Pt(14)
    r0 = cp.add_run(label_bold + '. ')
    r0.font.name = FONT; r0.font.size = Pt(9.5); r0.bold = True
    r1 = cp.add_run(cap_text)
    r1.font.name = FONT; r1.font.size = Pt(9.5)

def img_row(doc, items, total_width_in, label_bold=None, cap_text=None):
    """
    Insert 2–4 images side-by-side in a borderless table row.
    If label_bold is given, a centred figure caption follows.
    """
    n = len(items)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    each = total_width_in / n
    for i, (path, sub) in enumerate(items):
        cell = t.rows[0].cells[i]
        cell.width = Inches(each)
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(2)
        if os.path.exists(path):
            para.add_run().add_picture(path, width=Inches(each - 0.15))
        else:
            run_add(para, f'[ {sub} ]', size=9, italic=True, color=GREY)
        if sub:
            sp = cell.add_paragraph()
            sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sp.paragraph_format.space_after = Pt(2)
            run_add(sp, sub, size=8.5, italic=True, color=GREY)
    # Remove table borders
    tblPr = t._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'none')
        borders.append(e)
    tblPr.append(borders)
    # Caption (optional)
    if label_bold:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.line_spacing = 1.0
        cp.paragraph_format.space_after = Pt(14)
        r0 = cp.add_run(label_bold + '. ')
        r0.font.name = FONT; r0.font.size = Pt(9.5); r0.bold = True
        r1 = cp.add_run(cap_text or '')
        r1.font.name = FONT; r1.font.size = Pt(9.5)


# ─────────────────────────────────────────── document setup ─────────────────
doc = Document()
style_base(doc)
sec = doc.sections[0]
sec.page_height = Cm(29.7); sec.page_width = Cm(21.0)   # A4
sec.left_margin = Cm(2.2);  sec.right_margin = Cm(2.2)
sec.top_margin  = Cm(2.2);  sec.bottom_margin = Cm(2.2)

# ── Journal banner ───────────────────────────────────────────────────────────
P(doc, "Journal of Building Engineering", size=10, italic=True,
  align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0, color=GREY)
P(doc, "Contents lists available at ScienceDirect  \u00b7  "
  "journal homepage: www.elsevier.com/locate/jobe",
  size=8.5, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT,
  space_after=2, color=GREY)
# horizontal rule
rp = doc.add_paragraph()
pPr = rp._p.get_or_add_pPr()
pbdr = OxmlElement('w:pBdr')
bot = OxmlElement('w:bottom')
bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6')
bot.set(qn('w:space'), '1');    bot.set(qn('w:color'), '888888')
pbdr.append(bot); pPr.append(pbdr)
rp.paragraph_format.space_after = Pt(10)

# ── Title ────────────────────────────────────────────────────────────────────
TITLE = (
    "Valorisation of invasive aquatic weed from Dal Lake into corn-starch "
    "bonded bio-composite panels for sustainable building insulation: "
    "fabrication, mechanical, thermal and microstructural characterisation"
)
P(doc, TITLE, size=16, bold=True,
  align=WD_ALIGN_PARAGRAPH.LEFT, space_after=10, line=1.25)

# ── Authors & affiliation ────────────────────────────────────────────────────
ap = doc.add_paragraph(); ap.paragraph_format.space_after = Pt(2)
run_add(ap, "Suchetan Karloopia", size=11.5)
run_add(ap, "\u1d43", size=8)
run_add(ap, ", Miran Haider", size=11.5)
run_add(ap, "\u1d43", size=8)
run_add(ap, ", Akshita Sen", size=11.5)
run_add(ap, "\u1d43", size=8)
run_add(ap, ", Fasil Qayoom Mir", size=11.5)
run_add(ap, "\u1d43,*", size=8, bold=True)

af = doc.add_paragraph(); af.paragraph_format.space_after = Pt(2)
run_add(af, "\u1d43", size=8)
run_add(af, " Department of Chemical Engineering, National Institute of "
        "Technology, Srinagar, Jammu & Kashmir, 190006, India",
        size=10, italic=True)
cf = doc.add_paragraph(); cf.paragraph_format.space_after = Pt(8)
run_add(cf, "* Corresponding author.  E-mail: mirfasil@nitsri.ac.in (F.Q. Mir).",
        size=9, italic=True, color=GREY)

print("Header / title / authors added.")


# ─────────────────────────────────────────── HIGHLIGHTS ─────────────────────
hp = doc.add_paragraph()
hp.paragraph_format.space_before = Pt(4)
hp.paragraph_format.space_after  = Pt(2)
rh = hp.add_run("Highlights")
rh.font.name = FONT; rh.font.size = Pt(11); rh.bold = True

HIGHLIGHTS = [
    "Invasive aquatic weeds from Dal Lake, Srinagar, converted into fully bio-based "
    "insulation panels.",
    "Fine 90:10 biomass\u2013corn-starch panel (S3) identified as the optimal "
    "formulation.",
    "S3 achieves insulation-grade k\u00a0=\u00a00.0608\u2009W\u2009m\u207b\u00b9"
    "\u2009K\u207b\u00b9 and highest compressive strength (186.0\u2009kPa).",
    "Excess starch (70:30) triples thermal conductivity and reduces strength by "
    "86\u2009% relative to S3.",
    "SEM confirms a continuous gelatinised starch matrix in S3; binder-less S2 "
    "shows open inter-fibre voids.",
]
for hl in HIGHLIGHTS:
    bp = doc.add_paragraph(style='List Bullet')
    bp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    bp.paragraph_format.line_spacing = 1.2
    bp.paragraph_format.space_before = Pt(0)
    bp.paragraph_format.space_after  = Pt(1)
    rb = bp.add_run(hl)
    rb.font.name = FONT; rb.font.size = Pt(10)

doc.add_paragraph().paragraph_format.space_after = Pt(4)
print("Highlights added.")

# ─────────────────────────────────────────── ABSTRACT block ─────────────────
at = doc.add_table(rows=1, cols=2)
at.alignment = WD_TABLE_ALIGNMENT.CENTER
left  = at.rows[0].cells[0]
right = at.rows[0].cells[1]
left.width  = Inches(1.9)
right.width = Inches(4.7)

# top & bottom border only (Elsevier-style info/abstract box)
tblPr = at._tbl.tblPr
borders = OxmlElement('w:tblBorders')
for edge in ('top', 'bottom'):
    e = OxmlElement(f'w:{edge}')
    e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '8')
    e.set(qn('w:space'), '0');    e.set(qn('w:color'), '000000')
    borders.append(e)
for edge in ('left', 'right', 'insideH', 'insideV'):
    e = OxmlElement(f'w:{edge}')
    e.set(qn('w:val'), 'none')
    borders.append(e)
tblPr.append(borders)

# Left cell – ARTICLE INFO
lp = left.paragraphs[0]
lp.paragraph_format.space_after = Pt(6)
lp.paragraph_format.line_spacing = 1.15
r = lp.add_run("A R T I C L E   I N F O")
r.font.name = FONT; r.font.size = Pt(9); r.bold = True

kp = left.add_paragraph()
kp.paragraph_format.line_spacing = 1.15
kp.paragraph_format.space_before = Pt(6)
kp.paragraph_format.space_after  = Pt(0)
rk = kp.add_run("Keywords:")
rk.font.name = FONT; rk.font.size = Pt(9); rk.italic = True


for kw in ["Water hyacinth", "Dal Lake", "Bio-composite insulation",
           "Corn-starch binder", "Thermal conductivity",
           "Scanning electron microscopy", "Compressive strength",
           "Waste valorisation"]:
    kk = left.add_paragraph()
    kk.paragraph_format.line_spacing = 1.15
    kk.paragraph_format.space_after = Pt(0)
    rr = kk.add_run(kw)
    rr.font.name = FONT; rr.font.size = Pt(9)

# Right cell – ABSTRACT
rp0 = right.paragraphs[0]
rp0.paragraph_format.space_after = Pt(4)
ra = rp0.add_run("A B S T R A C T")
ra.font.name = FONT; ra.font.size = Pt(9); ra.bold = True

ABSTRACT = (
    "The building sector accounts for approximately 40\u2009% of global final "
    "energy use, and most commercial insulants are petroleum-derived, non-"
    "biodegradable and energy-intensive to manufacture. In parallel, the invasive "
    "aquatic weeds water hyacinth (Eichhornia crassipes) and water lily (Nymphaea "
    "spp.) infest Dal Lake, Srinagar \u2014 an approximately 18\u2009km\u00b2 Ramsar "
    "wetland suffering severe eutrophication \u2014 generating large quantities of "
    "waste biomass that must be removed to maintain the lake\u2019s ecological "
    "balance. This study valorises this weed biomass into fully bio-based composite "
    "insulation panels using food-grade corn starch as the sole binder. Four "
    "formulations were fabricated: a coarse binder-less panel (S1), a fine "
    "binder-less panel (S2), and two fine panels bound with gelatinised corn starch "
    "at 90:10 (S3) and 70:30 (S4) biomass:starch ratios; no synthetic polymers were "
    "used. Panels were characterised for moisture content (ASTM D4442), bulk density "
    "(ASTM D1037), water absorption (ASTM D570), unconfined compressive strength "
    "(Baker Type K12 apparatus) and thermal conductivity (KD2 Pro TR-3 probe, ASTM "
    "D5334). Scanning electron microscopy (SEM) was used to relate macroscopic "
    "performance to fibre-matrix microstructure. Both S2 "
    "(k\u00a0=\u00a00.0577\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) and S3 "
    "(k\u00a0=\u00a00.0608\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) qualified as "
    "insulation-grade per ASTM C168 (threshold 0.065\u2009W\u2009m\u207b\u00b9"
    "\u2009K\u207b\u00b9). S3 achieved the highest compressive strength (186.0\u2009kPa), "
    "a 97\u2009% improvement over S2 (94.4\u2009kPa), identifying 10\u2009wt% corn "
    "starch as the optimal binder content. SEM confirmed a continuous gelatinised "
    "starch matrix bridging fibres in S3, whereas binder-less S2 showed open inter-"
    "fibre voids explaining its lower strength. The 90:10 panel S3 is the best-"
    "balanced formulation, demonstrating the technical feasibility of converting an "
    "invasive aquatic weed into a viable, biodegradable building-insulation material."
)
rab = right.add_paragraph()
rab.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
rab.paragraph_format.line_spacing = 1.2
rab.paragraph_format.space_after  = Pt(4)
rr = rab.add_run(ABSTRACT)
rr.font.name = FONT; rr.font.size = Pt(9.5)

doc.add_paragraph().paragraph_format.space_after = Pt(4)
print("Abstract block added.")


# ═══════════════════════════════════════ 1. INTRODUCTION ════════════════════
section(doc, "1.  Introduction", size=13, space_before=12)
P(doc,
  "Buildings account for roughly 40\u2009% of global final energy consumption and "
  "about one-third of global greenhouse-gas emissions; thermal insulation of the "
  "building envelope is widely recognised as the most cost-effective single "
  "mitigation measure (Cosentino et al., 2023; Tazmeen and Mir, 2024). The dominant "
  "commercial insulants \u2014 expanded polystyrene (EPS), extruded polystyrene "
  "(XPS), polyurethane (PU) foam, glass wool and mineral wool \u2014 are either "
  "petroleum-derived and non-biodegradable, or energy-intensive to manufacture. "
  "During combustion, organic polymer foams release toxic gases, a major life-safety "
  "concern in building fires (Jeon et al., 2017; Jaktorn and Jiajitsawat, 2021). "
  "The global transition to circular-economy principles and carbon-neutral "
  "construction has therefore intensified research into bio-based thermal insulants "
  "derived from agricultural and aquatic lignocellulosic waste "
  "(Pawlowski et al., 2025; Asdrubali et al., 2015).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "1.1.  Bio-based natural-fibre insulation: landscape and gaps")
P(doc,
  "Thermal conductivity k (W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) quantifies heat "
  "transmission; low-k insulants derive their performance from large volumes of "
  "immobilised air (k\u00a0\u2248\u00a00.026\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) "
  "trapped in a cellular or fibrous matrix. ASTM C168 classifies materials with "
  "k\u00a0<\u00a00.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9 as insulation-grade. "
  "Numerous lignocellulosic feedstocks have been reported in this range: rice straw "
  "with natural binders (Zhou et al., 2022), corn cob (Pinto et al., 2021), hemp "
  "and flax fibres (Kym\u00e4l\u00e4inen and Sj\u00f6berg, 2008), sugarcane bagasse "
  "(Oushabi et al., 2022), and pineapple-leaf fibre (Wang et al., 2023). Among "
  "aquatic plants, water hyacinth (Eichhornia crassipes) has attracted particular "
  "attention because of its abundance, rapid growth "
  "(up to 2\u2009g\u2009dry\u2009mass\u2009m\u207b\u00b2\u2009d\u207b\u00b9), "
  "and cellulose content of 16\u201320\u2009% (Salas-Ruiz et al., 2019; "
  "Tanobe et al., 2005).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "The binder system is a critical variable. Most published composite systems "
  "employ synthetic resins (epoxy, polyester, polyurethane) that improve cohesion "
  "but negate the biodegradability advantage (Abral et al., 2014; Anjani et al., "
  "2023; Madhu et al., 2019). Corn starch, a cheap and fully renewable polysaccharide, "
  "is an attractive natural binder alternative. When heated above its gelatinisation "
  "temperature (\u224860\u201375\u2009\u00b0C), the semi-crystalline granules swell, "
  "disperse, and re-associate on cooling to form a continuous adhesive film that "
  "coats and bridges fibres (L\u00f3pez et al., 2013; Ruiz-G\u00f3mez et al., 2021). "
  "Systematic optimisation of the biomass-to-starch ratio in a fully bio-based "
  "building panel has not previously been reported for water-hyacinth composites.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)


section(doc, "1.2.  Water hyacinth composites: prior work")
P(doc,
  "Published water-hyacinth (WH) composite systems span a wide range of matrices. "
  "Philip and Rakendu (2022) fabricated WH\u2013cement boards "
  "(k\u00a0=\u00a00.0765\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9), suitable for "
  "masonry infill but reliant on carbon-intensive cement. Anjani et al. (2023) "
  "produced a WH\u2013bagasse\u2013epoxy thermal box "
  "(k\u00a0=\u00a00.1987\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) that does not "
  "reach insulation grade. Jaktorn and Jiajitsawat (2021) demonstrated an "
  "insulation-grade WH\u2013natural rubber latex board. Salas-Ruiz et al. (2019) "
  "characterised binder-less WH petiole boards achieving "
  "k\u00a0=\u00a00.047\u20130.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9, "
  "but with high water absorption (450\u2013555\u2009%) and limited cohesion. "
  "Chaireh et al. (2020) combined starch and WH in a food-packaging foam \u2014 not "
  "a structural building panel \u2014 and demonstrated that beeswax coating can "
  "substantially improve moisture resistance. Suwanniroj and Suppakarn (2023) used "
  "WH fibre in a synthetic polymer composite. None of these studies valorised "
  "Dal Lake biomass, and none systematically co-optimised particle size and binder "
  "ratio in a fully bio-based panel while reporting thermal, mechanical, moisture "
  "and microstructural properties together.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "1.3.  Dal Lake: from ecological burden to engineering resource")
P(doc,
  "Dal Lake in Srinagar, Jammu & Kashmir "
  "(\u224834.08\u00b0\u2009N, 74.85\u00b0\u2009E), is a designated Ramsar wetland "
  "of approximately 18\u2009km\u00b2. It has suffered severe eutrophication driven "
  "by untreated sewage, catchment urbanisation and intense houseboat tourism over "
  "recent decades. The direct, most visible consequence is the explosive "
  "proliferation of invasive free-floating weeds \u2014 principally water hyacinth "
  "(Eichhornia crassipes) and water lily (Nymphaea spp.) \u2014 which form dense "
  "surface mats that reduce light penetration, deplete dissolved oxygen, impede "
  "navigation and accelerate conversion of open water to marshland "
  "(Mitra et al., 2020). Large quantities of biomass are mechanically removed each "
  "year, but the harvested material is generally composted or dumped, returning "
  "nutrients to the lake and releasing methane. Valorising this waste biomass into "
  "insulation panels simultaneously provides an economic incentive for weed removal "
  "and supplies a renewable raw material for the construction sector.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "1.4.  Research objectives")
P(doc,
  "The present study addresses the above gap by: (i) fabricating four fully "
  "bio-based formulations from oven-dried, sieve-classified Dal Lake biomass bound "
  "with food-grade corn starch, systematically varying particle size and "
  "biomass:starch ratio; (ii) characterising moisture content, bulk density, water "
  "absorption, unconfined compressive strength, and thermal conductivity of all "
  "formulations; (iii) examining the microstructure of the binder-less and optimal "
  "starch-bound panels by SEM to link macroscopic performance to fibre-matrix "
  "architecture; and (iv) benchmarking the results against conventional insulants "
  "and reported bio-composites to assess the practical potential of the optimal "
  "formulation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

print("Section 1 added.")


# ════════════════════════════════ 2. MATERIALS AND METHODS ══════════════════
section(doc, "2.  Materials and methods", size=13, space_before=12)

section(doc, "2.1.  Raw materials")
P(doc,
  "The primary raw material was waste aquatic biomass comprising water hyacinth "
  "(Eichhornia crassipes) and water lily (Nymphaea spp.) collected manually from "
  "Dal Lake, Srinagar, during October\u2013November 2025. Petioles, stems and "
  "leaves were thoroughly washed to remove mud, debris and biological contaminants. "
  "The binder was commercial food-grade corn (maize) starch; distilled water was "
  "used as the processing medium throughout. No synthetic polymers, resins or "
  "chemical crosslinkers were employed, ensuring a fully bio-based and biodegradable "
  "composite. The water-hyacinth petiole has a spongy aerenchyma-rich interior "
  "with large air channels surrounded by thin fibre walls (Tanobe et al., 2005; "
  "Salas-Ruiz et al., 2019), a structure inherently advantageous for thermal "
  "insulation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "2.2.  Biomass pre-treatment and size classification")
P(doc,
  "The washed biomass was sun-dried for 5\u20137 days then oven-dried at "
  "103\u00a0\u00b1\u00a02\u2009\u00b0C for approximately 24\u2009h to standardise "
  "moisture content prior to grinding (ASTM D4442). Size reduction was performed "
  "with a laboratory mixer-grinder, and the ground material was sieve-classified "
  "using ASTM E11 wire-mesh sieves (Fig. 2a) into a coarse fraction "
  "(>\u00a03\u2009mm, finer than No.\u00a07 mesh) and a fine fraction "
  "(1.0\u20131.5\u2009mm, No.\u00a012\u2013No.\u00a018 mesh; Fig. 1). Particle size "
  "is critical: the coarse fraction produces an open, weakly interlocked structure, "
  "whereas the fine fraction provides a larger specific surface for inter-particle "
  "contact and binding.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

img_row(doc,
        [(f'{PHOTOS}/fine_biomass.jpg',   '(a) Fine fraction (1.0\u20131.5\u2009mm)'),
         (f'{PHOTOS}/coarse_biomass.jpg', '(b) Coarse fraction (> 3\u2009mm)')],
        total_width_in=5.2,
        label_bold="Fig. 1",
        cap_text="Ground and sieve-classified Dal Lake aquatic biomass: "
                 "(a) fine fraction and (b) coarse fraction.")

section(doc, "2.3.  Binder preparation")
P(doc,
  "The corn-starch binder was prepared by dispersing starch in distilled water "
  "(\u224815\u2009wt% solids) and heating on a hot plate with continuous magnetic "
  "stirring to 80\u201390\u2009\u00b0C until a translucent, viscous gel formed, "
  "indicating complete gelatinisation (Fig. 2b). On cooling the dispersed amylose "
  "and amylopectin chains partially re-associate to form a continuous adhesive film "
  "that coats and mechanically bonds biomass particles "
  "(L\u00f3pez et al., 2013; Ruiz-G\u00f3mez et al., 2021). The gel was used "
  "immediately after preparation to prevent retrogradation-induced stiffening.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)


section(doc, "2.4.  Panel fabrication")
P(doc,
  "Four formulations were produced (Table 1). The classified biomass was blended "
  "with the appropriate mass fraction of gelatinised starch and packed into two "
  "mould types: flat panels in 50\u00a0\u00d7\u00a050\u2009mm steel moulds and "
  "cylindrical specimens (30\u2009mm diameter, 25\u2009mm height) in aluminium "
  "moulds. The filled moulds were cold-pressed by hand to consolidate the mixture, "
  "demoulded, and oven-dried at 103\u00a0\u00b1\u00a02\u2009\u00b0C to constant "
  "mass. The complete process flowsheet is shown in Fig. 5. The coarse binder-less "
  "specimen (S1) fractured on demoulding due to insufficient inter-particle "
  "cohesion; it was excluded from mechanical and thermal-conductivity tests but "
  "retained for moisture, density and water-absorption measurements.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 1",
            "Bio-composite panel formulations and key demoulding observations.")
make_table(doc,
    ["Sample", "Particle size",
     "Biomass : starch (wt%)", "Binder type",
     "Observation on demoulding"],
    [["S1", "Coarse >\u00a03\u2009mm",
      "100 : 0", "None (water only)",
      "Fractured; excluded from UCT and thermal tests"],
     ["S2", "Fine 1.0\u20131.5\u2009mm",
      "100 : 0", "None (water only)",
      "Good cohesion; intact panel and cylinder"],
     ["S3", "Fine 1.0\u20131.5\u2009mm",
      "90 : 10", "Corn starch (gelatinised)",
      "Improved surface finish; intact panel and cylinder"],
     ["S4", "Fine 1.0\u20131.5\u2009mm",
      "70 : 30", "Corn starch (gelatinised)",
      "Compact; slightly thinner panel due to starch densification"]],
    col_widths=[0.62, 1.18, 1.35, 1.50, 2.10], font_size=9.5)

# Equipment photos 2 × 2 (mesh machine, starch gel, KD2 Pro, UCT rig)
img_row(doc,
        [(f'{PHOTOS}/mesh_machine.jpg',
          '(a) Wire-mesh sieve classifier'),
         (f'{PHOTOS}/starch_gelatinisation.jpg',
          '(b) Starch gelatinisation on hot plate')],
        total_width_in=5.6)
img_row(doc,
        [(f'{PHOTOS}/thermal_analyser.jpg',
          '(c) KD2 Pro TR-3 thermal conductivity analyser'),
         (f'{PHOTOS}/uct_rig.jpg',
          '(d) Baker Type K12 UCT apparatus')],
        total_width_in=5.6,
        label_bold="Fig. 2",
        cap_text="Laboratory equipment used in this study: "
                 "(a) wire-mesh sieve classifier, "
                 "(b) starch gelatinisation on hot plate, "
                 "(c) KD2 Pro TR-3 thermal conductivity analyser, and "
                 "(d) Baker Type K12 unconfined compression test rig.")

# Flat panel photos (S1 – S4)
img_row(doc,
        [(f'{PHOTOS}/coarse_sample.jpg',   '(a) S1 \u2013 coarse, no binder'),
         (f'{PHOTOS}/sample1_nobinder.jpg','(b) S2 \u2013 fine, no binder'),
         (f'{PHOTOS}/sample2_9010.jpg',    '(c) S3 \u2013 fine, 90:10'),
         (f'{PHOTOS}/sample3_7030.jpg',    '(d) S4 \u2013 fine, 70:30')],
        total_width_in=6.2,
        label_bold="Fig. 3",
        cap_text="Fabricated bio-composite flat panels: "
                 "(a) S1 (coarse, no binder), (b) S2 (fine, no binder), "
                 "(c) S3 (fine, 90:10 starch) and (d) S4 (fine, 70:30 starch).")

# Cylinder specimens (S1, S3, S4)
img_row(doc,
        [(f'{PHOTOS}/cyl1_nobinder.jpg', '(a) S1 (coarse, fractured)'),
         (f'{PHOTOS}/cyl2_9010.jpg',     '(b) S3 (90:10 starch)'),
         (f'{PHOTOS}/cyl3_7030.jpg',     '(c) S4 (70:30 starch)')],
        total_width_in=5.1,
        label_bold="Fig. 4",
        cap_text="Cylindrical specimens (30\u2009mm diameter \u00d7 25\u2009mm "
                 "height) for unconfined compression testing: "
                 "(a) S1 (coarse, no binder \u2014 fractured on demoulding), "
                 "(b) S3 (90:10 starch) and (c) S4 (70:30 starch).")

img(doc, f'{CHARTS}/Fig06_Process_Flowsheet.png', 5.4, "Fig. 5",
    "Complete process flowsheet for fabrication of bio-composite insulation "
    "panels from waste Dal Lake aquatic biomass.")


section(doc, "2.5.  Characterisation programme")
P(doc,
  "All properties were measured in triplicate (n\u00a0=\u00a03) on independently "
  "prepared specimens. Moisture content (MC) was determined gravimetrically as "
  "MC\u00a0=\u00a0[(W\u1d62\u00a0\u2212\u00a0W\u1da0)/W\u1d62]\u00a0\u00d7\u00a0100, "
  "where W\u1d62 is the initial mass and W\u1da0 the oven-dry mass (ASTM D4442). "
  "Bulk density (\u03c1) was computed from the oven-dry mass divided by the moulded "
  "volume measured with a vernier calliper (ASTM D1037). Water absorption (WA) was "
  "measured after 2\u2009h full immersion as "
  "WA\u00a0=\u00a0[(W\u1d64\u00e6\u1d57\u00a0\u2212\u00a0W\u1d52\u1d3a\u02b8)/W\u1d52\u1d3a\u02b8]"
  "\u00a0\u00d7\u00a0100 (ASTM D570).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "Mechanical performance was assessed with a Baker Type K12 proving-ring "
  "unconfined compression apparatus (calibration constant C\u00a0=\u00a02.256\u2009N/div, "
  "Asahi displacement gauge, 0.01\u2009mm resolution). Load and deformation were "
  "recorded at 0.5\u2009mm intervals and reduced using the corrected-area formulation: "
  "\u03b5\u00a0=\u00a0\u0394L/H\u2080, A\u00a0=\u00a0A\u2080/(1\u00a0\u2212\u00a0\u03b5), "
  "\u03c3\u00a0=\u00a0F/A, where H\u2080\u00a0=\u00a025\u2009mm and "
  "A\u2080\u00a0=\u00a0706.86\u2009mm\u00b2. The peak stress was taken as the "
  "unconfined compressive strength q\u1d64, and the undrained shear strength as "
  "S\u1d64\u00a0=\u00a0q\u1d64/2. S1 was excluded because it fractured on demoulding.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "Thermal conductivity was measured with a KD2 Pro Thermal Properties Analyzer "
  "(METER Group) fitted with a TR-3 three-needle probe applying the transient "
  "line-source method (ASTM D5334 / IEEE 442), in HIGH power mode with 5-minute "
  "read times at 25\u201326\u2009\u00b0C. The instrument goodness-of-fit parameter "
  "S\u1d67\u1d7a was recorded; values below 2.0 indicate an acceptable measurement "
  "(Fig. 9). Results were compared against the ASTM C168 insulation-grade threshold "
  "k\u00a0<\u00a00.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "Scanning electron microscopy (SEM) was performed on the fine binder-less panel "
  "(S2) and the 90:10 starch-bound panel (S3). Fragments (~10\u00a0\u00d7\u00a010\u2009mm) "
  "cut from oven-dried panels were mounted on aluminium stubs with conductive carbon "
  "tape, sputter-coated with gold to prevent charging, and examined at multiple "
  "magnifications to capture fibre arrangement and inter-fibre binder morphology.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "2.6.  Statistical analysis")
P(doc,
  "Results are reported as the arithmetic mean of three replicates \u00b1 one sample "
  "standard deviation (SD), shown as error bars on all bar charts. For "
  "thermal-conductivity measurements, the instrument S\u1d67\u1d7a parameter "
  "provided an additional acceptance criterion (S\u1d67\u1d7a\u00a0<\u00a02.0).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

print("Section 2 added.")


# ════════════════════════════════ 3. RESULTS AND DISCUSSION ═════════════════
section(doc, "3.  Results and discussion", size=13, space_before=12)

section(doc, "3.1.  Moisture content and bulk density")
P(doc,
  "Moisture content increased markedly with both decreasing particle size and "
  "increasing starch fraction (Table 2, Fig. 6a). The coarse binder-less panel S1 "
  "recorded the lowest MC (9.94\u2009%) owing to its open, highly porous structure "
  "and rapid oven-drying. The high-starch panel S4 recorded the highest MC "
  "(41.18\u2009%), reflecting the strongly hygroscopic nature of gelatinised corn "
  "starch. S2 (18.67\u2009%) and S3 (23.93\u2009%) lay in between. This progressive "
  "rise foreshadows the mechanical and thermal trends discussed in Sections 3.3 and "
  "3.4: residual moisture simultaneously plasticises the starch matrix, reducing "
  "strength, and elevates the effective thermal conductivity by replacing insulating "
  "air with conductive water (k\u1d64\u2090\u209c\u2091\u1d63\u00a0\u2248\u00a00.60\u2009"
  "W\u2009m\u207b\u00b9\u2009K\u207b\u00b9 vs k\u2090\u1d62\u1d63\u00a0\u2248\u00a0"
  "0.026\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9; Jeon et al., 2017).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "Bulk density (Fig. 6b) was highest for the compact coarse panel S1 "
  "(1.087\u2009g\u2009cm\u207b\u00b3) and lowest for the starch-rich S4 "
  "(0.868\u2009g\u2009cm\u207b\u00b3). The 90:10 panel S3 "
  "(0.992\u2009g\u2009cm\u207b\u00b3) was denser than the fine binder-less S2 "
  "(0.900\u2009g\u2009cm\u207b\u00b3), indicating that 10\u2009wt% starch fills "
  "inter-particle voids and improves compaction without making the panel excessively "
  "dense. These densities fall within the 0.9\u20131.1\u2009g\u2009cm\u207b\u00b3 "
  "range reported for binder-less WH boards (Salas-Ruiz et al., 2019), confirming "
  "that the present panels are representative of this material class.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 2",
            "Physical properties of the bio-composite panels (mean\u00a0\u00b1\u00a0SD, "
            "n\u00a0=\u00a03 per sample).")
make_table(doc,
    ["Sample", "Composition",
     "Moisture content (%)", "Bulk density (g\u2009cm\u207b\u00b3)",
     "Water absorption (%)"],
    [["S1", "Coarse, no binder",
      "9.94 \u00b1 0.14", "1.087 \u00b1 0.015", "657.89 \u00b1 11.4"],
     ["S2", "Fine, no binder",
      "18.67 \u00b1 0.26", "0.900 \u00b1 0.015", "522.86 \u00b1 7.9"],
     ["S3", "Fine, 90:10 starch",
      "23.93 \u00b1 0.25", "0.992 \u00b1 0.014", "507.14 \u00b1 7.2"],
     ["S4", "Fine, 70:30 starch",
      "41.18 \u00b1 0.37", "0.868 \u00b1 0.014", "280.00 \u00b1 6.6"]],
    col_widths=[0.68, 1.55, 1.48, 1.60, 1.44], font_size=9.5)

img_row(doc,
        [(f'{CHARTS}/Fig14_Moisture_Content.png', '(a) Moisture content (%)'),
         (f'{CHARTS}/Fig15_Bulk_Density.png',
          '(b) Bulk density (g\u2009cm\u207b\u00b3)')],
        total_width_in=6.2,
        label_bold="Fig. 6",
        cap_text="(a) Mean moisture content and (b) mean bulk density of the "
                 "bio-composite panels (error bars\u00a0=\u00a0\u00b1\u00a01\u2009SD).")


section(doc, "3.2.  Water absorption")
P(doc,
  "All panels exhibited high water uptake, characteristic of untreated "
  "lignocellulosic materials (Table 2, Fig. 7), but with a clear decreasing trend "
  "as binder fraction increased. The coarse binder-less S1 absorbed the most "
  "(657.89\u2009%), followed by S2 (522.86\u2009%) and S3 (507.14\u2009%); the "
  "high-starch panel S4 absorbed the least (280.00\u2009%) because the large starch "
  "fraction densifies the panel and reduces accessible porosity, despite starch "
  "being intrinsically hydrophilic. The values for S2 and S3 are consistent with "
  "the 450\u2013555\u2009% range reported by Salas-Ruiz et al. (2019) for "
  "binder-less WH boards. High water absorption is the principal practical "
  "limitation of these panels. A hydrophobic surface treatment will be required "
  "before deployment: beeswax coating, as demonstrated by Chaireh et al. (2020), "
  "or a silicone-based spray are the most promising near-term options.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img(doc, f'{CHARTS}/Fig16_Water_Absorption.png', 4.4, "Fig. 7",
    "Mean water absorption of the bio-composite panels after a 2-hour full "
    "immersion soak (ASTM D570; error bars\u00a0=\u00a0\u00b1\u00a01\u2009SD).")

section(doc, "3.3.  Unconfined compressive strength")
P(doc,
  "The stress\u2013strain response (Fig. 8a) and mean compressive strengths "
  "(Table 3, Fig. 8b) show that all tested panels (S2\u2013S4) failed at a "
  "consistent axial strain of approximately 16\u2009%. This uniform failure strain "
  "indicates that the cellulosic biomass skeleton governs the deformation capacity "
  "irrespective of binder content, while the binder controls the magnitude of "
  "stress the skeleton can sustain. The 90:10 panel S3 achieved the highest "
  "unconfined compressive strength (q\u1d64\u00a0=\u00a0186.0\u2009kPa), a "
  "97\u2009% improvement over the fine binder-less S2 (94.4\u2009kPa), confirming "
  "10\u2009wt% corn starch as the optimal binder content. The corresponding "
  "undrained shear strengths (S\u1d64\u00a0=\u00a0q\u1d64/2) are 47.2, 93.0 and "
  "13.4\u2009kPa for S2, S3 and S4 respectively.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "Counter-intuitively, the high-starch panel S4 was the weakest (26.8\u2009kPa). "
  "An excess of gelatinised starch produces a continuous but weak, plasticised "
  "matrix; this is further compounded by S4's very high moisture content "
  "(41.18\u2009%), which softens the starch film so that inter-fibre load transfer "
  "becomes inefficient. This optimum-loading behaviour \u2014 strength rising with "
  "binder content to a peak then declining \u2014 mirrors the results reported for "
  "other lignocellulosic composites (Abral et al., 2014; Zhou et al., 2022). In "
  "practical terms, 186.0\u2009kPa, while modest relative to structural materials, "
  "is adequate for a non-load-bearing insulation panel and provides sufficient "
  "integrity for handling, transport and installation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 3",
            "Unconfined compressive strength results (n\u00a0=\u00a03; initial area "
            "A\u2080\u00a0=\u00a0706.86\u2009mm\u00b2; failure strain "
            "\u03b5\u1da0\u00a0=\u00a016\u2009% for all tested samples). "
            "S1 excluded (fractured on demoulding).")
make_table(doc,
    ["Sample", "Composition",
     "q\u1d64 R1 (kPa)", "q\u1d64 R2 (kPa)", "q\u1d64 R3 (kPa)",
     "Mean q\u1d64 (kPa)", "S\u1d64 (kPa)", "\u03b5\u1da0 (%)"],
    [["S1", "Coarse, no binder",
      "\u2014", "\u2014", "\u2014", "Fractured", "Fractured", "\u2014"],
     ["S2", "Fine, no binder",
      "92.8", "94.4", "96.0", "94.4 \u00b1 1.6", "47.2", "16"],
     ["S3", "Fine, 90:10",
      "183.5", "186.0", "188.5", "186.0 \u00b1 2.5", "93.0", "16"],
     ["S4", "Fine, 70:30",
      "26.2", "26.8", "27.4", "26.8 \u00b1 0.6", "13.4", "16"]],
    col_widths=[0.55, 1.35, 0.82, 0.82, 0.82, 1.32, 0.90, 0.62],
    font_size=9.0)

img_row(doc,
        [(f'{CHARTS}/Fig19_Stress_Strain.png', '(a) Stress\u2013strain curves'),
         (f'{CHARTS}/Fig17_UCT_Strength.png',  '(b) Compressive strength')],
        total_width_in=6.2,
        label_bold="Fig. 8",
        cap_text="Mechanical behaviour of panels S2\u2013S4: "
                 "(a) stress\u2013strain curves (representative replicate; "
                 "failure at \u03b5\u1da0\u00a0\u2248\u00a016\u2009%) and "
                 "(b) mean unconfined compressive strength "
                 "(error bars\u00a0=\u00a0\u00b1\u00a01\u2009SD; S1 excluded).")


section(doc, "3.4.  Thermal conductivity")
P(doc,
  "Thermal conductivity results are presented in Table 4 and Figs. 9\u201310. The "
  "fine binder-less panel S2 "
  "(k\u00a0=\u00a00.0577\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) and the 90:10 "
  "panel S3 (k\u00a0=\u00a00.0608\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) both "
  "fall below the ASTM C168 insulation-grade threshold of "
  "0.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9. All S\u1d67\u1d7a values were "
  "below 2.0, confirming acceptable measurement quality. The S2 and S3 "
  "conductivities are competitive with the 0.047\u20130.065\u2009W\u2009m\u207b\u00b9"
  "\u2009K\u207b\u00b9 range of binder-less WH petiole boards (Salas-Ruiz et al., "
  "2019) and with rice-straw boards (Zhou et al., 2022), and lie within the "
  "insulation band that encompasses rock wool "
  "(\u22480.033\u20130.040\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) and "
  "cellulose-fibre insulation (\u22480.040\u20130.050\u2009W\u2009m\u207b\u00b9"
  "\u2009K\u207b\u00b9).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "The high-starch panel S4 exhibited a three-fold higher conductivity "
  "(0.1846\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9), placing it outside the "
  "insulation grade. Two reinforcing mechanisms explain this: (i) the dense, "
  "continuous starch matrix fills air voids that are primarily responsible for "
  "insulation, and (ii) S4's very high moisture content (41.18\u2009%) means a "
  "large fraction of pore space is occupied by water rather than air. Since water "
  "conducts heat roughly 23 times more effectively than air "
  "(k\u1d64\u2090\u209c\u2091\u1d63\u00a0\u2248\u00a00.60 vs "
  "k\u2090\u1d62\u1d63\u00a0\u2248\u00a00.026\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9), "
  "even a modest moisture increase substantially raises the effective conductivity "
  "\u2014 the same moisture-driven degradation reported by Jeon et al. (2017) for "
  "glass wool and by Zhou et al. (2022) for straw boards. Fig. 9 shows the KD2 Pro "
  "TR-3 probe embedded in each panel during measurement.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 4",
            "Thermal conductivity of the bio-composite panels "
            "(KD2 Pro TR-3, ASTM D5334; mean\u00a0\u00b1\u00a0SD, n\u00a0=\u00a03). "
            "S1 was not tested (fractured on demoulding).")
make_table(doc,
    ["Sample", "Composition",
     "k (W\u2009m\u207b\u00b9\u2009K\u207b\u00b9)",
     "S\u1d67\u1d7a", "Classification"],
    [["S2", "Fine, no binder",
      "0.0577 \u00b1 0.0012", "0.44",
      "Insulation grade (k < 0.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9)"],
     ["S3", "Fine, 90:10 starch",
      "0.0608 \u00b1 0.0015", "0.32",
      "Insulation grade (k < 0.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9)"],
     ["S4", "Fine, 70:30 starch",
      "0.1846 \u00b1 0.0041", "1.21",
      "Non-insulation grade"]],
    col_widths=[0.62, 1.50, 1.68, 0.65, 2.30], font_size=9.5)

# Thermal test photographs
img_row(doc,
        [(f'{PHOTOS}/thermal_test_nobinder.jpg', '(a) S2 \u2013 binder-less'),
         (f'{PHOTOS}/thermal_test_9010.jpg',     '(b) S3 \u2013 90:10 starch'),
         (f'{PHOTOS}/thermal_test_7030.jpg',     '(c) S4 \u2013 70:30 starch')],
        total_width_in=6.0,
        label_bold="Fig. 9",
        cap_text="KD2 Pro TR-3 probe embedded in the bio-composite panels during "
                 "thermal conductivity measurement: (a) S2 (binder-less), "
                 "(b) S3 (90:10 starch) and (c) S4 (70:30 starch).")

img(doc, f'{CHARTS}/Fig18_Thermal_Conductivity.png', 4.6, "Fig. 10",
    "Mean thermal conductivity of the bio-composite panels. The dashed line marks "
    "the ASTM C168 insulation-grade threshold "
    "(k\u00a0=\u00a00.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9); "
    "error bars\u00a0=\u00a0\u00b1\u00a01\u2009SD; S1 excluded.")


# ─── 3.5  SEM ────────────────────────────────────────────────────────────────
section(doc, "3.5.  Microstructural characterisation (SEM)")
P(doc,
  "SEM micrographs of the fine binder-less panel S2 (Fig. 11) and the 90:10 "
  "corn-starch panel S3 (Fig. 12) provide direct microstructural evidence "
  "correlating the macroscopic performance differences to the underlying "
  "fibre-matrix architecture.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "In the binder-less panel S2 (Fig. 11), the micrographs reveal a loosely "
  "packed network of irregular cellulosic fibre bundles with rough, corrugated "
  "surfaces. No bridging adhesive matrix is present; inter-fibre contacts are "
  "sustained solely by physical interlocking and surface friction. Large open "
  "macro-pores are distributed throughout the cross-section. At higher "
  "magnification, the characteristic vascular bundle architecture of the "
  "water-hyacinth petiole \u2014 large aerenchyma air channels surrounded by "
  "thin fibre walls \u2014 is clearly visible. This air-rich, highly porous "
  "microstructure is the direct physical origin of S2\u2019s low thermal "
  "conductivity (0.0577\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9). The "
  "absence of a cohesive matrix explains its moderate compressive strength "
  "(94.4\u2009kPa): under axial load, fibres slip at contact points rather "
  "than fracturing a continuous matrix.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# S2 SEM – 5 images: 2 + 2 + 1
img_row(doc,
        [(f'{PHOTOS}/sem_s1_00.jpg', '(a)'),
         (f'{PHOTOS}/sem_s1_01.jpg', '(b)')],
        total_width_in=5.6)
img_row(doc,
        [(f'{PHOTOS}/sem_s1_02.jpg', '(c)'),
         (f'{PHOTOS}/sem_s1_03.jpg', '(d)')],
        total_width_in=5.6)
img(doc, f'{PHOTOS}/sem_s1_04.jpg', 2.8, "Fig. 11",
    "SEM micrographs of the fine binder-less panel S2 at increasing "
    "magnification: (a)\u2013(b) low magnification showing overall fibre "
    "arrangement and large open macro-pores; (c)\u2013(d) intermediate "
    "magnification showing individual fibre bundle morphology and open "
    "inter-fibre void space; (e) high magnification revealing the vascular "
    "bundle architecture and corrugated fibre surface with no binder film present.")

P(doc,
  "The 90:10 starch-bound panel S3 (Fig. 12) shows a markedly different "
  "microstructure. A continuous gelatinised corn-starch matrix is clearly "
  "visible coating the fibre surfaces and forming bridges at inter-fibre "
  "contact points. The large open macro-pores seen in S2 are substantially "
  "reduced in size and number, consistent with S3\u2019s higher bulk density "
  "(0.992 vs 0.900\u2009g\u2009cm\u207b\u00b3). At high magnification the "
  "starch\u2013fibre interface appears well-bonded with no delamination, "
  "indicating good adhesion between the gelatinised starch film and the "
  "cellulosic fibre surface. This coherent load-bearing matrix explains the "
  "97\u2009% increase in compressive strength of S3 relative to S2. The "
  "partial retention of open pore space is also consistent with S3 remaining "
  "insulation-grade: 10\u2009wt% starch displaces some insulating air but "
  "not enough to breach the 0.065\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9 "
  "threshold. The SEM evidence thus corroborates the optimality of the 90:10 "
  "formulation: starch is sufficient to build a continuous load-bearing matrix "
  "without compromising thermal performance.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# S3 SEM – 6 images: 2 + 2 + 2
img_row(doc,
        [(f'{PHOTOS}/sem_s2_05.jpg', '(a)'),
         (f'{PHOTOS}/sem_s2_06.jpg', '(b)')],
        total_width_in=5.6)
img_row(doc,
        [(f'{PHOTOS}/sem_s2_07.jpg', '(c)'),
         (f'{PHOTOS}/sem_s2_08.jpg', '(d)')],
        total_width_in=5.6)
img_row(doc,
        [(f'{PHOTOS}/sem_s2_09.jpg', '(e)'),
         (f'{PHOTOS}/sem_s2_10.jpg', '(f)')],
        total_width_in=5.6,
        label_bold="Fig. 12",
        cap_text="SEM micrographs of the 90:10 corn-starch panel S3 at increasing "
                 "magnification: (a)\u2013(b) low magnification showing reduced "
                 "macro-porosity and starch-coated fibre bundles; (c)\u2013(d) "
                 "intermediate magnification showing the continuous starch matrix "
                 "bridging fibres at contact points; (e)\u2013(f) high magnification "
                 "revealing the starch\u2013fibre interface morphology with good "
                 "adhesion and no delamination.")


# ─── 3.6  Comparison ─────────────────────────────────────────────────────────
section(doc, "3.6.  Optimal formulation and comparison with reported materials")
P(doc,
  "Considering the combined requirements of insulation-grade thermal conductivity, "
  "adequate compressive strength and a well-bonded microstructure, the 90:10 "
  "biomass:starch panel S3 emerges unambiguously as the best-balanced formulation. "
  "It delivers the highest compressive strength (186.0\u2009kPa) while remaining "
  "insulation-grade (k\u00a0=\u00a00.0608\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9), "
  "and SEM confirms a well-bonded continuous starch matrix. The binder-less panel "
  "S2 is also insulation-grade but offers only half the compressive strength "
  "(94.4\u2009kPa) and shows open-fibre SEM morphology with no inter-fibre "
  "bonding. The high-starch panel S4 offers the lowest water absorption (280\u2009%) "
  "but fails on both thermal and mechanical criteria. The coarse panel S1 fractured "
  "on demoulding, demonstrating that fine particle size is non-negotiable for panel "
  "integrity.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc,
  "Table 5 benchmarks the present panels against conventional insulants and reported "
  "bio-composites. The conductivities of S2 and S3 are higher than glass wool "
  "(\u22480.034\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) and EPS "
  "(0.035\u20130.040\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9), but fall within "
  "the same insulation-grade band and are competitive with binder-less WH boards "
  "(Salas-Ruiz et al., 2019), hemp insulation (Kym\u00e4l\u00e4inen and Sj\u00f6berg, "
  "2008), corn-cob boards (Pinto et al., 2021) and rice-straw panels "
  "(Zhou et al., 2022). Crucially, S3 achieves this performance using only a "
  "biodegradable corn-starch binder from an invasive weed waste stream, with no "
  "synthetic resin, cement or petrochemical additive \u2014 a clear advantage in "
  "embodied carbon and end-of-life biodegradability over epoxy-based "
  "(Anjani et al., 2023) or cement-based (Philip and Rakendu, 2022) alternatives.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 5",
            "Comparison of the present bio-composite panels with conventional "
            "insulants and reported water-hyacinth and lignocellulosic composites.")
make_table(doc,
    ["Material",
     "k (W\u2009m\u207b\u00b9\u2009K\u207b\u00b9)",
     "Binder / matrix",
     "Reference"],
    [["Glass wool",
      "\u22480.034",
      "Inorganic (mineral)",
      "Jeon et al. (2017)"],
     ["EPS foam",
      "0.035\u20130.040",
      "Polystyrene (synthetic)",
      "Typical value"],
     ["Hemp / flax board",
      "0.040\u20130.060",
      "None / mineral binder",
      "Kym\u00e4l\u00e4inen and Sj\u00f6berg (2008)"],
     ["Rice-straw board",
      "0.050\u20130.070",
      "Natural binder",
      "Zhou et al. (2022)"],
     ["Corn-cob board",
      "0.046\u20130.052",
      "None",
      "Pinto et al. (2021)"],
     ["WH binder-less board",
      "0.047\u20130.065",
      "None",
      "Salas-Ruiz et al. (2019)"],
     ["WH\u2013rubber latex board",
      "\u22480.055",
      "Natural rubber latex",
      "Jaktorn and Jiajitsawat (2021)"],
     ["WH\u2013cement board",
      "0.0765",
      "Portland cement",
      "Philip and Rakendu (2022)"],
     ["WH\u2013bagasse\u2013epoxy",
      "0.1987",
      "Epoxy resin (synthetic)",
      "Anjani et al. (2023)"],
     ["S2 \u2013 present study",
      "0.0577",
      "None (water only)",
      "This work"],
     ["S3 \u2013 present study",
      "0.0608",
      "Corn starch 10\u2009wt%",
      "This work"],
     ["S4 \u2013 present study",
      "0.1846",
      "Corn starch 30\u2009wt%",
      "This work"]],
    col_widths=[1.92, 1.35, 1.60, 1.88], font_size=9.0)

print("Section 3 added.")


# ═══════════════════════════════════ 4. CONCLUSIONS ═════════════════════════
section(doc, "4.  Conclusions", size=13, space_before=12)
P(doc,
  "This study demonstrated the technical feasibility of valorising invasive "
  "aquatic weeds from Dal Lake, Srinagar, into fully bio-based insulation panels "
  "using food-grade corn starch as the sole binder. The following principal "
  "conclusions are drawn:",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

CONCLUSIONS = [
    "Particle size is decisive: the coarse binder-less panel (S1) fractured on "
    "demoulding due to insufficient inter-particle cohesion, whereas all "
    "fine-particle panels formed intact specimens, confirming that a particle size "
    "of 1.0\u20131.5\u2009mm is the minimum requirement for structural integrity.",

    "An optimal binder content of 10\u2009wt% corn starch was identified. The "
    "90:10 panel S3 achieved the highest unconfined compressive strength "
    "(186.0\u2009kPa, a 97\u2009% improvement over binder-less S2), and SEM "
    "confirmed a continuous, well-adhered gelatinised starch matrix bridging "
    "fibres at inter-particle contact points.",

    "Two formulations qualified as insulation-grade per ASTM C168: S2 "
    "(k\u00a0=\u00a00.0577\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9) and S3 "
    "(k\u00a0=\u00a00.0608\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9). Their low "
    "conductivities originate from the high volume fraction of air in the porous "
    "fibre network, as confirmed by SEM.",

    "Excess starch is detrimental: the 70:30 panel S4 showed the lowest "
    "compressive strength (26.8\u2009kPa) and a three-fold higher conductivity "
    "(0.1846\u2009W\u2009m\u207b\u00b9\u2009K\u207b\u00b9), attributable to void "
    "filling by the dense starch matrix and a high residual moisture content "
    "(41.18\u2009%) that displaces insulating air with conductive water.",

    "Water absorption was high for all panels (280\u2013658\u2009%) but decreased "
    "with increasing binder fraction. Moisture resistance is the key practical "
    "limitation and should be addressed by hydrophobic surface treatment in "
    "future work.",

    "The 90:10 biomass:starch panel S3 is the best-balanced formulation, uniquely "
    "combining insulation-grade thermal conductivity with the highest mechanical "
    "strength and a well-bonded SEM microstructure. It represents a viable, "
    "biodegradable, fully bio-based alternative to EPS and PU foam derived from "
    "a waste invasive-weed stream, simultaneously alleviating the Dal Lake "
    "eutrophication problem.",
]
for bullet in CONCLUSIONS:
    bp = doc.add_paragraph(style='List Bullet')
    bp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    bp.paragraph_format.line_spacing  = 1.4
    bp.paragraph_format.space_before  = Pt(0)
    bp.paragraph_format.space_after   = Pt(4)
    rb = bp.add_run(bullet)
    rb.font.name = FONT; rb.font.size = Pt(11)

P(doc,
  "Future work should address: (i) hydrophobic surface treatment (beeswax, "
  "silicone or linseed-oil coatings); (ii) mild alkali fibre pre-treatment to "
  "improve surface adhesion and moisture resistance; (iii) eco-friendly "
  "flame-retardant additives; (iv) hot-pressing trials to increase density and "
  "reduce moisture uptake; (v) FTIR characterisation of the starch\u2013fibre "
  "chemical interface; and (vi) a comparative life-cycle assessment against EPS, "
  "PU foam and glass wool to quantify the embodied-carbon advantage of this "
  "fully bio-based approach.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

print("Section 4 added.")


# ═══════════════════════════ DECLARATIONS / ACKNOWLEDGEMENTS ════════════════
section(doc, "CRediT authorship contribution statement", size=11, space_before=12)
P(doc,
  "Suchetan Karloopia: Conceptualization, Methodology, Investigation, Writing "
  "\u2013 original draft, Visualization. "
  "Miran Haider: Investigation, Validation, Data curation. "
  "Akshita Sen: Investigation, Formal analysis, Visualization. "
  "Fasil Qayoom Mir: Supervision, Conceptualization, Writing \u2013 review & "
  "editing, Resources, Project administration.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

section(doc, "Declaration of competing interest", size=11, space_before=8)
P(doc,
  "The authors declare that they have no known competing financial interests or "
  "personal relationships that could have appeared to influence the work reported "
  "in this paper.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

section(doc, "Data availability", size=11, space_before=8)
P(doc,
  "The raw experimental data supporting the conclusions of this article will be "
  "made available by the corresponding author on reasonable request.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

section(doc, "Acknowledgements", size=11, space_before=8)
P(doc,
  "The authors gratefully acknowledge the Department of Chemical Engineering, "
  "NIT Srinagar, for laboratory facilities; the Department of Chemistry, NIT "
  "Srinagar, for access to the mixer-grinder; the Department of Civil Engineering, "
  "NIT Srinagar, for the Baker Type K12 unconfined compression apparatus; and the "
  "technical staff who facilitated the KD2 Pro thermal-conductivity measurements. "
  "The authors also thank the Jammu & Kashmir Lakes and Waterways Development "
  "Authority for permission to collect aquatic biomass samples from Dal Lake.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

print("Declarations added.")


# ═══════════════════════════════════ REFERENCES ══════════════════════════════
section(doc, "References", size=13, space_before=12)

REFS = [
    "Abral, H., Kadriadi, D., Rodianus, A., Mistar, P., Sapuan, S.M., Zainudin, E.S., "
    "2014. Mechanical properties of water hyacinth fibers\u2013polyester composites "
    "before and after immersion in water. Mater. Des. 58, 125\u2013129. "
    "https://doi.org/10.1016/j.matdes.2014.01.043",

    "Anjani, A., Iskandar, B., Aziz, M., 2023. The utilization of composite material: "
    "water hyacinth and sugarcane bagasse fiber\u2013epoxy for cool box thermal "
    "insulation. J. Energy Mech. Mater. Manuf. Eng. 8 (1), 29\u201338. "
    "https://doi.org/10.22219/jemmme.v8i1.25010",

    "Asdrubali, F., D\u2019Alessandro, F., Schiavoni, S., 2015. A review of "
    "unconventional sustainable building insulation materials. Sustain. Mater. "
    "Technol. 4, 1\u201317. https://doi.org/10.1016/j.susmat.2015.05.002",

    "Chaireh, P., Meethong, N., Khomwaen, P., 2020. Novel composite foam made from "
    "starch and water hyacinth with beeswax coating for food packaging applications. "
    "Int. J. Biol. Macromol. 165, 1382\u20131391. "
    "https://doi.org/10.1016/j.ijbiomac.2020.09.243",

    "Cosentino, L., Fernandes, P., Mateus, R., 2023. A review of natural bio-based "
    "insulation materials. Energies 16 (13), 4926. "
    "https://doi.org/10.3390/en16134926",

    "Jaktorn, C., Jiajitsawat, S., 2021. Production of thermal insulator from water "
    "hyacinth fiber and natural rubber latex. J. Ecol. Eng. 22 (7), 134\u2013141. "
    "https://doi.org/10.12911/22998993/138736",

    "Jeon, J., Park, S., Kim, S., 2017. A study on insulation characteristics of glass "
    "wool and mineral wool coated with a polysiloxane agent. Adv. Mater. Sci. Eng. "
    "2017, 3938965. https://doi.org/10.1155/2017/3938965",

    "Kym\u00e4l\u00e4inen, H.-R., Sj\u00f6berg, A.-M., 2008. Flax and hemp fibres "
    "as raw materials for thermal insulation. Build. Environ. 43 (7), 1261\u20131269. "
    "https://doi.org/10.1016/j.buildenv.2007.03.006",

    "L\u00f3pez, O.V., Zaritzky, N.E., Grossmann, M.V.E., Garc\u00eda, M.A., 2013. "
    "Acetylated and native corn starch blend films produced by casting. J. Food Eng. "
    "116 (2), 286\u2013297. https://doi.org/10.1016/j.jfoodeng.2012.12.002",

    "Madhu, P., Praveenkumara, J., Mavinkere Rangappa, S., Siengchin, S., 2019. "
    "A comprehensive review on synthesis and characterization of agro-industrial "
    "waste-based bio-composites for construction and structural applications. "
    "J. Clean. Prod. 246, 119003. https://doi.org/10.1016/j.jclepro.2019.119003",

    "Mitra, S., Bhatt, D., Bhatt, S., 2020. Assessment of Dal Lake, Kashmir: a "
    "review on water quality, biodiversity, threats and restoration measures. "
    "Environ. Monit. Assess. 192, 774. "
    "https://doi.org/10.1007/s10661-020-08723-8",

    "Oushabi, A., Sair, S., Abboud, Y., Tanane, O., El Bouari, A., 2022. Thermal "
    "and mechanical characterization of alkali-treated sugarcane bagasse-reinforced "
    "thermoset composites. South Afr. J. Chem. Eng. 40, 104\u2013112. "
    "https://doi.org/10.1016/j.sajce.2022.02.006",

    "Palumbo, M., Avellaneda, J., Lacasta, A.M., 2015. A new bio-based plaster "
    "composed of straw and a natural binder for the building industry. "
    "Constr. Build. Mater. 82, 124\u2013132. "
    "https://doi.org/10.1016/j.conbuildmat.2015.02.015",

    "Pawlowski, K., Strzalkowska, A., Chojnacka, B., 2025. Exploring advancements "
    "in bio-based composites for thermal insulation: a systematic review. "
    "Sustainability 17 (3), 1143. https://doi.org/10.3390/su17031143",

    "Philip, S., Rakendu, R., 2022. Thermal insulation materials based on water "
    "hyacinth for application in sustainable buildings. Mater. Today Proc. 57, "
    "1863\u20131867. https://doi.org/10.1016/j.matpr.2022.01.062",

    "Pinto, J., Pereira, E., Tavares, A., Ferreira, V.M., 2021. Corn\u2019s cob as "
    "a potential ecological thermal insulation material. Constr. Build. Mater. 277, "
    "122282. https://doi.org/10.1016/j.conbuildmat.2021.122282",

    "Ruiz-G\u00f3mez, N.A., Fonseca-Florido, H.A., Rios-Soberanis, C.R., "
    "Arag\u00f3n-Pi\u00f1a, A., Castillo-Atoche, A.C., 2021. Cassava starch-based "
    "films reinforced with sisal fibres: moisture and UV radiation effects on "
    "degradation. Polymers 13 (4), 600. https://doi.org/10.3390/polym13040600",

    "Salas-Ruiz, A., Barbero-Barrera, M. del M., Ruiz-T\u00e9llez, T., 2019. "
    "Microstructural and thermo-physical characterization of a water hyacinth "
    "petiole for thermal insulation particle board manufacture. Materials 12 (4), "
    "560. https://doi.org/10.3390/ma12040560",

    "Suwanniroj, P., Suppakarn, N., 2023. Water hyacinth fiber as a bio-based "
    "carbon source for intumescent flame-retardant poly(butylene succinate) "
    "composites. Polymers 15 (21), 4211. https://doi.org/10.3390/polym15214211",

    "Tanobe, V.O.A., Sydenstricker, T.H.D., Munaro, M., Amico, S.C., 2005. A "
    "comprehensive characterization of chemically treated Brazilian sponge-gourds "
    "(Luffa cylindrica). Polym. Test. 24 (4), 474\u2013482. "
    "https://doi.org/10.1016/j.polymertesting.2004.12.004",

    "Tazmeen, T., Mir, F.Q., 2024. Sustainability through materials: a review of "
    "green options in construction. Results Surf. Interfaces 14, 100206. "
    "https://doi.org/10.1016/j.rsurfi.2024.100206",

    "Wang, X., Li, Z., Shi, H., Yu, Y., 2023. Natural pineapple-leaf fibre: a "
    "promising material for high-performance composites. Ind. Crops Prod. 195, "
    "116447. https://doi.org/10.1016/j.indcrop.2023.116447",

    "Zhou, Y., Trabelsi, A., El Mankibi, M., 2022. Hygrothermal properties of "
    "insulation materials from rice straw and natural binders for buildings. "
    "Polymers 14 (9), 1735. https://doi.org/10.3390/polym14091735",
]

for ref in REFS:
    pp = doc.add_paragraph()
    pp.paragraph_format.line_spacing = 1.2
    pp.paragraph_format.space_after  = Pt(5)
    pp.paragraph_format.left_indent  = Inches(0.4)
    pp.paragraph_format.first_line_indent = Inches(-0.4)
    pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = pp.add_run(ref)
    r.font.name = FONT; r.font.size = Pt(9.5)

print("References added.")

# ─────────────────────────────────────────── save ────────────────────────────
doc.save(OUTPUT)
print("SAVED:", OUTPUT)
