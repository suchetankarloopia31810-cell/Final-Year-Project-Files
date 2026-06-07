#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build a research-article manuscript formatted for the
JOURNAL OF BUILDING ENGINEERING (Elsevier), based on the B.Tech final-year
project "Development and Characterization of Bio-Composites from Waste Aquatic
Biomass for Sustainable Insulation" (NIT Srinagar).

The layout mirrors the supervisor's sample Elsevier paper:
  - title, authors, affiliation, corresponding-author footnote
  - "ARTICLE INFO / ABSTRACT" block with keywords
  - numbered sections (1. Introduction ... Conclusions)
  - author-year (Harvard) in-text citations
  - alphabetical reference list with DOIs
Single-column manuscript style (the form Elsevier requests at submission).
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

# ---------------------------------------------------------------- helpers ----
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
    """Numbered top-level / sub section heading (bold, left)."""
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

def caption(doc, label_bold, text, size=9.5, space_before=4, space_after=12):
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_before = Pt(space_before)
    cp.paragraph_format.space_after = Pt(space_after)
    r0 = cp.add_run(label_bold + '. ')
    r0.font.name = FONT; r0.font.size = Pt(size); r0.bold = True
    r1 = cp.add_run(text)
    r1.font.name = FONT; r1.font.size = Pt(size)
    return cp

def img(doc, path, width_in, label_bold, cap_text, ph=None):
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(); run.add_picture(path, width=Inches(width_in))
    else:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(8)
        run_add(p, f'[ {ph or os.path.basename(path)} ]', size=10,
                italic=True, color=GREY)
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_after = Pt(14)
    r0 = cp.add_run(label_bold + '. '); r0.font.name = FONT
    r0.font.size = Pt(9.5); r0.bold = True
    r1 = cp.add_run(cap_text); r1.font.name = FONT; r1.font.size = Pt(9.5)

def img_row(doc, items, total_width_in, label_bold, cap_text):
    n = len(items)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    each = total_width_in / n
    for i, (path, sub) in enumerate(items):
        cell = t.rows[0].cells[i]; cell.width = Inches(each)
        para = cell.paragraphs[0]; para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(2)
        if os.path.exists(path):
            run = para.add_run(); run.add_picture(path, width=Inches(each - 0.15))
        else:
            run_add(para, f'[ {sub} ]', size=9, italic=True, color=GREY)
        if sub:
            sp = cell.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sp.paragraph_format.space_after = Pt(2)
            run_add(sp, sub, size=8.5, italic=True, color=GREY)
    tblPr = t._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{edge}'); e.set(qn('w:val'), 'none'); borders.append(e)
    tblPr.append(borders)
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_after = Pt(14)
    r0 = cp.add_run(label_bold + '. '); r0.font.name = FONT
    r0.font.size = Pt(9.5); r0.bold = True
    r1 = cp.add_run(cap_text); r1.font.name = FONT; r1.font.size = Pt(9.5)

def tbl_caption(doc, label_bold, text, size=9.5):
    cp = doc.add_paragraph()
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_before = Pt(10)
    cp.paragraph_format.space_after = Pt(3)
    r0 = cp.add_run(label_bold); r0.font.name = FONT
    r0.font.size = Pt(size); r0.bold = True
    cp2 = doc.add_paragraph()
    cp2.paragraph_format.line_spacing = 1.0
    cp2.paragraph_format.space_after = Pt(4)
    r1 = cp2.add_run(text); r1.font.name = FONT; r1.font.size = Pt(size)
    r1.italic = True

def make_table(doc, headers, rows, col_widths=None, header_bg='D9D9D9',
               font_size=9.5, zebra=False, highlight=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_bg(hdr[i], header_bg)
        para = hdr[i].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.line_spacing = 1.0
        para.paragraph_format.space_after = Pt(2)
        para.paragraph_format.space_before = Pt(2)
        r = para.add_run(h); r.font.name = FONT; r.font.size = Pt(font_size)
        r.bold = True
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for ci, val in enumerate(row):
            para = cells[ci].paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.space_before = Pt(2)
            r = para.add_run(str(val)); r.font.name = FONT
            r.font.size = Pt(font_size)
            cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if zebra and ri % 2 == 1:
                set_cell_bg(cells[ci], 'F2F2F2')
            if highlight and (ri, ci) in highlight:
                set_cell_bg(cells[ci], highlight[(ri, ci)])
    if col_widths:
        for row in t.rows:
            for ci, w in enumerate(col_widths):
                row.cells[ci].width = Inches(w)
    return t

# ---------------------------------------------------------------- document ---
doc = Document()
style_base(doc)
sec = doc.sections[0]
sec.page_height = Cm(29.7); sec.page_width = Cm(21.0)        # A4
sec.left_margin = Cm(2.2); sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2.2)

# Journal banner (mimics the sample's running header)
P(doc, "Journal of Building Engineering", size=10, italic=True,
  align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0, color=GREY)
P(doc, "Contents lists available at ScienceDirect  \u00b7  journal homepage: "
  "www.elsevier.com/locate/jobe", size=8.5, italic=True,
  align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, color=GREY)
# rule
rp = doc.add_paragraph()
pPr = rp._p.get_or_add_pPr(); pbdr = OxmlElement('w:pBdr')
bottom = OxmlElement('w:bottom')
bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '6')
bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '888888')
pbdr.append(bottom); pPr.append(pbdr); rp.paragraph_format.space_after = Pt(10)

# Title
TITLE = ("Valorisation of invasive aquatic weed from Dal Lake into corn-starch "
         "bonded bio-composite panels for sustainable building insulation")
P(doc, TITLE, size=16, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
  space_after=10, line=1.25)

# Authors
ap = doc.add_paragraph(); ap.paragraph_format.space_after = Pt(2)
run_add(ap, "Suchetan Karloopia", size=11.5)
run_add(ap, "\u1d43", size=8)
run_add(ap, ", Miran Haider", size=11.5)
run_add(ap, "\u1d43", size=8)
run_add(ap, ", Akshita Sen", size=11.5)
run_add(ap, "\u1d43", size=8)
run_add(ap, ", Fasil Qayoom Mir", size=11.5)
run_add(ap, "\u1d43,", size=8)
run_add(ap, "*", size=10, bold=True)

# Affiliation
af = doc.add_paragraph(); af.paragraph_format.space_after = Pt(2)
run_add(af, "\u1d43", size=8)
run_add(af, " Department of Chemical Engineering, National Institute of "
        "Technology, Srinagar, Jammu & Kashmir, 190006, India", size=10,
        italic=True)
cf = doc.add_paragraph(); cf.paragraph_format.space_after = Pt(10)
run_add(cf, "* Corresponding author. E-mail address: mirfasil@nitsri.ac.in "
        "(F.Q. Mir).", size=9, italic=True, color=GREY)

print("Header + title + authors added.")

# ---------------------------------------------------------- ABSTRACT block ---
# Two-cell borderless table: left = ARTICLE INFO/keywords, right = ABSTRACT
at = doc.add_table(rows=1, cols=2)
at.alignment = WD_TABLE_ALIGNMENT.CENTER
left = at.rows[0].cells[0]; right = at.rows[0].cells[1]
left.width = Inches(1.9); right.width = Inches(4.7)

# top & bottom border for the block
tbl = at._tbl; tblPr = tbl.tblPr
borders = OxmlElement('w:tblBorders')
for edge in ('top', 'bottom'):
    e = OxmlElement(f'w:{edge}')
    e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '8')
    e.set(qn('w:space'), '0'); e.set(qn('w:color'), '000000')
    borders.append(e)
for edge in ('left', 'right', 'insideH', 'insideV'):
    e = OxmlElement(f'w:{edge}'); e.set(qn('w:val'), 'none'); borders.append(e)
tblPr.append(borders)

# left column: ARTICLE INFO + keywords
lp = left.paragraphs[0]; lp.paragraph_format.space_after = Pt(6)
lp.paragraph_format.line_spacing = 1.15
r = lp.add_run("A R T I C L E   I N F O"); r.font.name = FONT
r.font.size = Pt(9); r.bold = True
kp = left.add_paragraph(); kp.paragraph_format.line_spacing = 1.15
kp.paragraph_format.space_before = Pt(6); kp.paragraph_format.space_after = Pt(0)
rk = kp.add_run("Keywords:"); rk.font.name = FONT; rk.font.size = Pt(9)
rk.italic = True
for kw in ["Water hyacinth", "Bio-composite", "Corn-starch binder",
           "Thermal insulation", "Compressive strength", "Waste valorisation"]:
    kk = left.add_paragraph(); kk.paragraph_format.line_spacing = 1.15
    kk.paragraph_format.space_after = Pt(0)
    rr = kk.add_run(kw); rr.font.name = FONT; rr.font.size = Pt(9)

# right column: ABSTRACT
rp0 = right.paragraphs[0]; rp0.paragraph_format.space_after = Pt(4)
ra = rp0.add_run("A B S T R A C T"); ra.font.name = FONT
ra.font.size = Pt(9); ra.bold = True

ABSTRACT = (
    "The building sector's rising energy demand and the environmental burden of "
    "petroleum-derived insulants such as expanded polystyrene and polyurethane foam "
    "have intensified the search for sustainable, bio-based alternatives. In "
    "parallel, the invasive aquatic weeds water hyacinth (Eichhornia crassipes) and "
    "water lily (Nymphaea spp.) infest Dal Lake, Srinagar \u2014 a Ramsar wetland of "
    "about 18 km\u00b2 affected by eutrophication \u2014 generating a large, freely "
    "available waste-biomass stream. This study valorises this weed into bio-composite "
    "insulation panels. Four fully bio-based formulations were fabricated from "
    "oven-dried, sieve-classified biomass: a coarse binder-less panel (S1), a fine "
    "binder-less panel (S2), and two fine panels bound with gelatinised food-grade "
    "corn starch at biomass:starch ratios of 90:10 (S3) and 70:30 (S4); no synthetic "
    "polymers or crosslinkers were used. Panels were characterized for moisture "
    "content (ASTM D4442), bulk density (ASTM D1037), water absorption (ASTM D570), "
    "unconfined compressive strength and thermal conductivity (KD2 Pro TR-3 transient "
    "line-source probe, ASTM D5334), with three replicates per test. The fine "
    "binder-less panel (S2) and the 90:10 panel (S3) achieved insulation-grade thermal "
    "conductivities of 0.0577 and 0.0608 W m\u207b\u00b9 K\u207b\u00b9, respectively, "
    "both below the 0.065 W m\u207b\u00b9 K\u207b\u00b9 threshold of ASTM C168. The "
    "90:10 panel recorded the highest unconfined compressive strength (186.0 kPa), a "
    "97 % improvement over the binder-less fine panel, identifying 10 wt% corn starch "
    "as the optimal binder content. The 70:30 panel showed the lowest water absorption "
    "(280 %) but the lowest strength (26.8 kPa) and a three-fold higher conductivity "
    "(0.1846 W m\u207b\u00b9 K\u207b\u00b9), attributed to an excess plasticised starch "
    "matrix and a high moisture content (41.18 %). The 90:10 biomass:starch panel "
    "emerged as the best-balanced formulation, demonstrating the technical feasibility "
    "of converting an invasive aquatic weed into a viable, biodegradable "
    "building-insulation material."
)
rab = right.add_paragraph()
rab.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
rab.paragraph_format.line_spacing = 1.2
rab.paragraph_format.space_after = Pt(4)
rr = rab.add_run(ABSTRACT); rr.font.name = FONT; rr.font.size = Pt(9.5)

doc.add_paragraph().paragraph_format.space_after = Pt(2)
print("Abstract block added.")



# ============================================================ 1. INTRODUCTION
section(doc, "1.  Introduction", size=13, space_before=12)
P(doc, "Thermal insulation is among the most cost-effective strategies for reducing "
  "operational energy in buildings, which account for a substantial share of global "
  "energy use and greenhouse-gas emissions (Cosentino et al., 2023). The majority of "
  "commercial insulation products in service today \u2014 expanded polystyrene (EPS), "
  "extruded polystyrene (XPS), polyurethane (PU) foam, glass wool and mineral wool "
  "\u2014 are either petroleum-derived or energy-intensive to manufacture. These "
  "materials are non-biodegradable, accumulate in landfills at end of life, and, in "
  "the case of organic foams, are highly flammable and release toxic gases during "
  "combustion (Jaktorn and Jiajitsawat, 2021; Jeon et al., 2017). The emphasis on "
  "circular-economy principles and carbon-footprint reduction has therefore driven "
  "strong interest in bio-based insulation derived from renewable and waste "
  "lignocellulosic resources (Pawlowski et al., 2025).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "Natural-fibre bio-composites are attractive because they are renewable, "
  "biodegradable, lightweight and possess low thermal conductivity owing to their "
  "porous cellular structure, and can frequently be produced at low cost from "
  "agricultural or aquatic waste. A broad range of lignocellulosic feedstocks "
  "\u2014 rice straw (Zhou et al., 2022), corn cob (Pinto et al., 2021), hemp and "
  "flax (Kym\u00e4l\u00e4inen and Sj\u00f6berg, 2008), sugarcane bagasse (Oushabi et "
  "al., 2022) and pineapple-leaf fibre (Wang et al., 2023), among others \u2014 has "
  "been investigated for insulation. Among aquatic feedstocks, water hyacinth "
  "(Eichhornia crassipes) has attracted particular attention because of its "
  "abundance, rapid growth and high cellulose content (Salas-Ruiz et al., 2019).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "The thermal performance of an insulant is quantified primarily by its thermal "
  "conductivity, k (W m\u207b\u00b9 K\u207b\u00b9); a lower value denotes a better "
  "insulator. Lightweight insulants owe their low conductivity to a large volume "
  "fraction of immobilised air (k \u2248 0.026 W m\u207b\u00b9 K\u207b\u00b9) trapped "
  "within a cellular or fibrous solid. Any factor that displaces this air \u2014 "
  "densification, void filling or, most importantly, moisture ingress \u2014 raises "
  "the effective conductivity. Water (k \u2248 0.6 W m\u207b\u00b9 K\u207b\u00b9) "
  "conducts heat roughly 23 times more effectively than air, which is why moisture "
  "management is a recurring theme in bio-based insulation (Jeon et al., 2017).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "1.1.  The Dal Lake problem: an invasive weed as a resource")
P(doc, "Dal Lake in Srinagar, Jammu & Kashmir, is an iconic Himalayan urban lake and "
  "a designated Ramsar wetland of approximately 18 km\u00b2. Over recent decades it "
  "has suffered severe eutrophication driven by untreated sewage, catchment "
  "urbanisation and nutrient loading. A direct, highly visible consequence is the "
  "explosive proliferation of free-floating invasive weeds, principally water "
  "hyacinth and water lily (Nymphaea spp.), which form dense surface mats that "
  "deplete dissolved oxygen, block sunlight, destroy fish-breeding grounds and "
  "accelerate the conversion of open water to marshland. Mechanical de-weeding "
  "removes large quantities of this biomass each year, but the harvested material is "
  "generally treated as waste, dumped or left to decompose \u2014 releasing methane "
  "and returning nutrients to the lake. Converting this weed into a useful "
  "engineering product simultaneously addresses two problems: it provides an "
  "economic end-use for weed removal and supplies a renewable raw material for "
  "sustainable insulation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "1.2.  Research gap and objectives")
P(doc, "Reported water-hyacinth systems can reach insulation-grade conductivities, "
  "but most rely on synthetic matrices such as epoxy (Anjani et al., 2023) or "
  "polyester (Abral et al., 2014), on energy-intensive binders such as cement (Philip "
  "and Rakendu, 2022), or on binder-less boards that suffer very high water "
  "absorption and poor mechanical integrity (Salas-Ruiz et al., 2019). Studies that "
  "use natural binders are often directed at packaging foams rather than building "
  "panels (Chaireh et al., 2020), and few systematically optimise the "
  "biomass-to-binder ratio of a fully bio-based panel while reporting thermal, "
  "mechanical and moisture performance together. Moreover, the specific valorisation "
  "of Dal Lake aquatic weed for insulation has not been reported. The present work "
  "addresses this gap by fabricating fully bio-based panels from Dal Lake biomass "
  "bound with food-grade corn starch (no synthetic polymers or crosslinkers), "
  "systematically varying particle size and biomass:starch ratio, and characterising "
  "the resulting thermal, mechanical and moisture properties together to identify an "
  "optimal, insulation-grade formulation.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ====================================================== 2. MATERIALS & METHODS
section(doc, "2.  Materials and methods", size=13, space_before=12)

section(doc, "2.1.  Raw materials")
P(doc, "The primary raw material was waste aquatic biomass comprising water hyacinth "
  "(Eichhornia crassipes) and water lily (Nymphaea spp.), collected manually from "
  "Dal Lake, Srinagar (approximately 34.08\u00b0 N, 74.85\u00b0 E) during "
  "October\u2013November 2025. Petioles, stems and leaves were thoroughly washed to "
  "remove mud, debris and biological contaminants. The binder was commercial "
  "food-grade corn (maize) starch; tap/distilled water was used as the processing "
  "medium. No synthetic polymers, resins or chemical crosslinkers were employed at "
  "any stage, ensuring a fully bio-based and biodegradable composite. The petioles "
  "and stems possess a porous, fibrous, cellulose-rich internal structure that is "
  "advantageous both for thermal insulation and for structural reinforcement.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

section(doc, "2.2.  Biomass pre-treatment and size reduction")
P(doc, "The washed biomass was sun-dried for 5\u20137 days and then oven-dried at "
  "103 \u00b1 2 \u00b0C for about 24 h to standardise the moisture content prior to "
  "grinding (consistent with ASTM D4442). The dried biomass was reduced in size with "
  "a laboratory mixer-grinder and classified using ASTM E11 standard wire-mesh sieves "
  "into a coarse fraction (> 3 mm; finer than No. 7 mesh) and a fine fraction "
  "(1.0\u20131.5 mm; No. 12\u2013No. 18 mesh). Particle size is a critical variable: "
  "the coarse fraction produces an open, weakly interlocked structure, whereas the "
  "fine fraction offers a larger specific surface for inter-particle contact and "
  "binding, yielding more cohesive panels (Fig. 1).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img_row(doc, [(f'{PHOTOS}/fine_biomass.jpg', '(a) Fine fraction (1.0\u20131.5 mm)'),
              (f'{PHOTOS}/coarse_biomass.jpg', '(b) Coarse fraction (> 3 mm)')],
        total_width_in=5.0, label_bold="Fig. 1",
        cap_text="Ground and sieve-classified aquatic biomass: (a) fine fraction and "
                 "(b) coarse fraction.")

section(doc, "2.3.  Binder preparation and composite fabrication")
P(doc, "The corn-starch binder was prepared by dispersing starch in water and heating "
  "on a hot plate with magnetic stirring to 80\u201390 \u00b0C until a translucent, "
  "viscous gel formed, indicating complete gelatinisation. On cooling and drying the "
  "dispersed starch chains re-associate to form a continuous film that bridges and "
  "bonds the biomass particles, acting as a natural adhesive matrix. The classified "
  "biomass was mixed with the appropriate binder/water and packed into moulds: flat "
  "panels in 50 \u00d7 50 mm steel moulds and cylindrical specimens (30 mm diameter, "
  "25 mm height) in aluminium moulds. The filled moulds were cold-pressed by hand to "
  "consolidate the mixture, demoulded, and oven-dried at 103 \u00b1 2 \u00b0C to "
  "constant mass. Four formulations were produced (Table 1; Fig. 2).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 1", "Bio-composite sample fabrication details.")
make_table(doc,
    ["Sample", "Particle size", "Composition / binder", "Observation"],
    [["S1", "Coarse > 3 mm", "Biomass only + 15 % water",
      "Poor stability; fractured on demoulding"],
     ["S2", "Fine 1.0\u20131.5 mm", "Biomass only + 15 % water",
      "Good cohesion; intact panel"],
     ["S3", "Fine 1.0\u20131.5 mm", "Biomass:starch = 90:10",
      "Improved binding and surface finish"],
     ["S4", "Fine 1.0\u20131.5 mm", "Biomass:starch = 70:30",
      "Compact, denser panel; highest binder fraction"]],
    col_widths=[0.7, 1.2, 1.9, 2.4], font_size=9.5)

img_row(doc, [(f'{PHOTOS}/coarse_sample.jpg', '(a) S1'),
              (f'{PHOTOS}/sample1_nobinder.jpg', '(b) S2'),
              (f'{PHOTOS}/sample2_9010.jpg', '(c) S3'),
              (f'{PHOTOS}/sample3_7030.jpg', '(d) S4')],
        total_width_in=6.0, label_bold="Fig. 2",
        cap_text="Fabricated bio-composite panels: (a) S1 (coarse, no binder), "
                 "(b) S2 (fine, no binder), (c) S3 (90:10) and (d) S4 (70:30).")

img(doc, f'{CHARTS}/Fig06_Process_Flowsheet.png', 5.4, "Fig. 3",
    "Process flowsheet for fabrication of the bio-composite insulation panels from "
    "waste aquatic biomass.")

section(doc, "2.4.  Characterization methods")
P(doc, "Panels were characterized following recognised ASTM standards, with three "
  "replicates (n = 3) per property. Moisture content (MC) was determined "
  "gravimetrically as MC = (W\u1d62 \u2212 W\u0066)/W\u1d62 \u00d7 100, where W\u1d62 "
  "and W\u0066 are the masses before and after oven-drying to constant mass (ASTM "
  "D4442). Bulk density (\u03c1) was obtained from the oven-dry mass and the moulded "
  "volume measured with a vernier calliper (ASTM D1037). Water absorption (WA) was "
  "measured after a 2-h immersion as WA = (W_wet \u2212 W_dry)/W_dry \u00d7 100 "
  "(ASTM D570).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
P(doc, "Mechanical performance of the cylindrical specimens was evaluated with a "
  "Baker Type K12 proving-ring unconfined compression apparatus (calibration constant "
  "C = 2.256 N/div) fitted with an ASAHI displacement gauge (0.01 mm). Load and "
  "deformation were recorded at 0.5 mm intervals and reduced using the standard "
  "area-correction relations \u03b5 = \u0394L/H\u2080, A = A\u2080/(1 \u2212 \u03b5), "
  "\u03c3 = F/A, with the unconfined compressive strength q\u1d64 taken as the peak "
  "stress and the undrained shear strength as S\u1d64 = q\u1d64/2 (initial area "
  "A\u2080 = 706.86 mm\u00b2). Sample S1 was excluded as it fractured on demoulding. "
  "Thermal conductivity (k) was measured with a KD2 Pro Thermal Properties Analyzer "
  "(METER Group) fitted with a TR-3 three-needle probe applying the transient "
  "line-source method (ASTM D5334 / IEEE 442) in HIGH mode with 5-min read times at "
  "25\u201326 \u00b0C; the instrument goodness-of-fit parameter (Syx) was recorded, "
  "with values below 2.0 considered acceptable. Results are reported as the mean of "
  "three replicates with the sample standard deviation shown as error bars.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

print("Sections 1 and 2 added.")

# ================================================== 3. RESULTS AND DISCUSSION
section(doc, "3.  Results and discussion", size=13, space_before=12)

section(doc, "3.1.  Moisture content and bulk density")
P(doc, "Moisture content increased markedly with both decreasing particle size and "
  "increasing starch fraction (Table 2, Fig. 4a). The coarse binder-less panel S1 "
  "recorded the lowest MC (9.94 %) owing to its open, highly porous structure, while "
  "the high-starch panel S4 recorded the highest (41.18 %), reflecting the strongly "
  "hygroscopic nature of corn starch; S2 (18.67 %) and S3 (23.93 %) lay in between. "
  "The progressive rise in MC from S1 to S4 foreshadows the mechanical and thermal "
  "trends discussed below, since residual moisture both plasticises the binder and "
  "degrades insulation. Bulk density (Fig. 4b) was highest for the compact coarse "
  "panel S1 (1.087 g cm\u207b\u00b3) and lowest for the thinner, starch-rich S4 "
  "(0.868 g cm\u207b\u00b3); the 90:10 panel S3 (0.992 g cm\u207b\u00b3) was denser "
  "than the fine binder-less S2 (0.900 g cm\u207b\u00b3), indicating that a modest "
  "10 % starch addition fills inter-particle voids and improves compaction.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 2", "Physical properties of the bio-composite panels "
            "(mean of n = 3).")
make_table(doc,
    ["Sample", "Moisture content (%)", "Bulk density (g cm\u207b\u00b3)",
     "Water absorption (%)"],
    [["S1", "9.94", "1.087", "657.89"],
     ["S2", "18.67", "0.900", "522.86"],
     ["S3", "23.93", "0.992", "507.14"],
     ["S4", "41.18", "0.868", "280.00"]],
    col_widths=[0.9, 1.9, 1.9, 1.7], font_size=9.5)

img_row(doc, [(f'{CHARTS}/Fig14_Moisture_Content.png', '(a) Moisture content'),
              (f'{CHARTS}/Fig15_Bulk_Density.png', '(b) Bulk density')],
        total_width_in=6.2, label_bold="Fig. 4",
        cap_text="(a) Mean moisture content and (b) mean bulk density of the "
                 "bio-composite panels (error bars = \u00b1 1 SD).")

section(doc, "3.2.  Water absorption")
P(doc, "All panels exhibited high water uptake, characteristic of untreated "
  "lignocellulosic materials, but a clear decreasing trend with increasing binder "
  "fraction was observed (Table 2, Fig. 5). The binder-less coarse panel S1 absorbed "
  "the most (657.89 %), followed by S2 (522.86 %) and S3 (507.14 %); the high-starch "
  "panel S4 absorbed the least (280.00 %), because the larger starch fraction fills "
  "open pores and reduces accessible porosity despite starch itself being "
  "hydrophilic. These values are consistent with the 450\u2013555 % range reported by "
  "Salas-Ruiz et al. (2019) for binder-less water-hyacinth boards and confirm that "
  "moisture resistance remains the principal limitation, motivating a future "
  "hydrophobic treatment such as the beeswax coating of Chaireh et al. (2020).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img(doc, f'{CHARTS}/Fig16_Water_Absorption.png', 4.4, "Fig. 5",
    "Mean water absorption of the bio-composite panels after a 2-h soak "
    "(error bars = \u00b1 1 SD).")

section(doc, "3.3.  Unconfined compressive strength")
P(doc, "All tested panels failed at a consistent axial strain of about 16 % "
  "(Fig. 6a), indicating that the biomass skeleton governs the failure strain "
  "irrespective of binder content. The 90:10 panel S3 achieved the highest "
  "unconfined compressive strength (q\u1d64 = 186.0 kPa), a 97 % improvement over the "
  "fine binder-less panel S2 (94.4 kPa), confirming 10 wt% corn starch as the optimal "
  "binder content (Fig. 6b). Counter-intuitively, the high-starch panel S4 was the "
  "weakest (26.8 kPa): an excess of starch produces a continuous but weak, "
  "plasticised matrix that is further softened by its very high moisture content "
  "(41.18 %), so that load is carried by the soft binder rather than transferred "
  "efficiently between fibres. This optimum-loading behaviour mirrors that reported "
  "for other water-hyacinth composites, where intermediate fibre/binder ratios "
  "maximise strength (Abral et al., 2014; Zhou et al., 2022). The corresponding "
  "undrained shear strengths (S\u1d64 = q\u1d64/2) of 47.2, 93.0 and 13.4 kPa for S2, "
  "S3 and S4 follow the same ranking.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)
img_row(doc, [(f'{CHARTS}/Fig19_Stress_Strain.png', '(a) Stress\u2013strain response'),
              (f'{CHARTS}/Fig17_UCT_Strength.png', '(b) Compressive strength')],
        total_width_in=6.2, label_bold="Fig. 6",
        cap_text="Mechanical behaviour: (a) stress\u2013strain curves and (b) mean "
                 "unconfined compressive strength of panels S2\u2013S4 "
                 "(error bars = \u00b1 1 SD).")

section(doc, "3.4.  Thermal conductivity")
P(doc, "The fine binder-less panel S2 (k = 0.0577 W m\u207b\u00b9 K\u207b\u00b9) and "
  "the 90:10 panel S3 (k = 0.0608 W m\u207b\u00b9 K\u207b\u00b9) both fall below the "
  "0.065 W m\u207b\u00b9 K\u207b\u00b9 insulation-grade threshold of ASTM C168 and "
  "therefore qualify as insulation-grade materials (Table 3, Fig. 7). The high-starch "
  "panel S4 exhibited a three-fold higher conductivity (0.1846 W m\u207b\u00b9 "
  "K\u207b\u00b9), placing it outside the insulation grade. This is attributed to the "
  "dense, continuous starch matrix that fills the air voids responsible for "
  "insulation, compounded by its very high moisture content (41.18 %); because water "
  "conducts heat roughly 23 times better than air, moisture-filled pores sharply "
  "raise the effective conductivity \u2014 the same moisture-driven degradation "
  "reported for fibrous insulants by Jeon et al. (2017). All Syx values were below "
  "2.0, confirming acceptable measurement quality. The S2 and S3 values are "
  "competitive with the 0.047\u20130.065 W m\u207b\u00b9 K\u207b\u00b9 range reported "
  "by Salas-Ruiz et al. (2019) and approach conventional glass wool "
  "(\u2248 0.034 W m\u207b\u00b9 K\u207b\u00b9).",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 3", "Thermal conductivity of the bio-composite panels "
            "(KD2 Pro TR-3, ASTM D5334; mean of n = 3). S1 was not tested "
            "(fractured on demoulding).")
make_table(doc,
    ["Sample", "k (W m\u207b\u00b9 K\u207b\u00b9)", "Syx", "Classification"],
    [["S2", "0.0577", "0.4388", "Insulation grade (k < 0.065)"],
     ["S3", "0.0608", "0.3171", "Insulation grade (k < 0.065)"],
     ["S4", "0.1846", "1.2143", "Non-insulation grade"]],
    col_widths=[0.9, 1.8, 1.0, 2.6], font_size=9.5,
    highlight={(0, 1): 'E2EFDA', (0, 3): 'E2EFDA',
               (1, 1): 'E2EFDA', (1, 3): 'E2EFDA'})
img(doc, f'{CHARTS}/Fig18_Thermal_Conductivity.png', 4.6, "Fig. 7",
    "Mean thermal conductivity of the bio-composite panels relative to the ASTM C168 "
    "insulation threshold (error bars = \u00b1 1 SD; S1 excluded).")

section(doc, "3.5.  Optimal formulation and comparison with reported materials")
P(doc, "Considering the combined requirements of insulation-grade thermal "
  "conductivity and adequate mechanical strength, the 90:10 biomass:starch panel (S3) "
  "emerges as the best-balanced formulation: it delivers the highest compressive "
  "strength (186.0 kPa) while remaining insulation-grade (0.0608 W m\u207b\u00b9 "
  "K\u207b\u00b9). The fine binder-less panel S2 is also insulation-grade but only "
  "half as strong, while the high-starch panel S4 offers the best water resistance "
  "but fails on both thermal and mechanical criteria. Thus 10 wt% corn starch "
  "represents the optimal binder content for this system. As summarised in Table 4, "
  "the conductivities of S2 and S3 are higher than those of optimised commercial "
  "insulants such as glass wool (\u2248 0.034 W m\u207b\u00b9 K\u207b\u00b9) and EPS "
  "(0.035\u20130.040 W m\u207b\u00b9 K\u207b\u00b9) but fall within the same "
  "insulation-grade band, and are comparable to or better than several reported "
  "water-hyacinth composites such as the WH\u2013cement composite of Philip and "
  "Rakendu (2022) and the WH\u2013bagasse\u2013epoxy composite of Anjani et al. "
  "(2023). Crucially, the present panels achieve this using only a biodegradable "
  "corn-starch binder, without any synthetic resin or cement, giving a clear "
  "advantage in biodegradability and embodied carbon.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

tbl_caption(doc, "Table 4", "Comparison of the present bio-composites with "
            "conventional and reported insulation materials.")
make_table(doc,
    ["Material", "k (W m\u207b\u00b9 K\u207b\u00b9)", "Binder / matrix", "Reference"],
    [["Glass wool", "\u2248 0.034", "Inorganic (synthetic)", "Jeon et al. (2017)"],
     ["EPS foam", "0.035\u20130.040", "Polystyrene (synthetic)", "Typical"],
     ["WH binder-less board", "0.047\u20130.065", "None", "Salas-Ruiz et al. (2019)"],
     ["WH\u2013cement composite", "0.0765", "Cement", "Philip and Rakendu (2022)"],
     ["WH\u2013bagasse\u2013epoxy", "0.1987", "Epoxy (synthetic)", "Anjani et al. (2023)"],
     ["S2 (this work)", "0.0577", "None (water only)", "Present study"],
     ["S3 (this work)", "0.0608", "Corn starch (90:10)", "Present study"],
     ["S4 (this work)", "0.1846", "Corn starch (70:30)", "Present study"]],
    col_widths=[1.8, 1.3, 1.9, 1.7], font_size=9.0,
    highlight={(5, 0): 'E2EFDA', (5, 1): 'E2EFDA',
               (6, 0): 'E2EFDA', (6, 1): 'E2EFDA'})

print("Section 3 added.")

# ============================================================ 4. CONCLUSIONS
section(doc, "4.  Conclusions", size=13, space_before=12)
P(doc, "This study demonstrated the technical feasibility of converting invasive "
  "waste aquatic biomass from Dal Lake into fully bio-based insulation panels using a "
  "natural corn-starch binder. Particle size was decisive: the coarse binder-less "
  "panel fractured on demoulding, whereas all fine-particle panels formed intact "
  "specimens. An optimal binder content of 10 wt% corn starch was identified \u2014 "
  "the 90:10 panel (S3) achieved the highest unconfined compressive strength "
  "(186.0 kPa), a 97 % improvement over the binder-less fine panel (94.4 kPa). Two "
  "formulations qualified as insulation-grade per ASTM C168: S2 (0.0577 W m\u207b\u00b9 "
  "K\u207b\u00b9) and S3 (0.0608 W m\u207b\u00b9 K\u207b\u00b9). Excess starch was "
  "detrimental: the 70:30 panel showed the lowest strength (26.8 kPa) and a "
  "three-fold higher conductivity (0.1846 W m\u207b\u00b9 K\u207b\u00b9), caused by a "
  "weak plasticised matrix and high moisture content (41.18 %). Water absorption was "
  "high for all panels (280\u2013658 %) but decreased with binder fraction, confirming "
  "moisture resistance as the key property requiring improvement. Overall, the 90:10 "
  "biomass:starch panel (S3) is the best-balanced formulation, uniquely combining "
  "insulation-grade thermal conductivity with the highest mechanical strength. Future "
  "work should address hydrophobic surface treatment (e.g. beeswax or silicone), mild "
  "alkali fibre modification, eco-friendly flame retardants, SEM/FTIR "
  "characterisation, hot-pressing trials, and a comparative life-cycle assessment "
  "against EPS, PU foam and glass wool to advance the panels towards a deployable, "
  "circular-economy building product.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================== Declarations / Acknowledgements
section(doc, "CRediT authorship contribution statement", size=11, space_before=12)
P(doc, "Suchetan Karloopia: Conceptualization, Methodology, Investigation, Writing "
  "\u2013 original draft. Miran Haider: Investigation, Validation, Data curation. "
  "Akshita Sen: Investigation, Formal analysis, Visualization. Fasil Qayoom Mir: "
  "Supervision, Conceptualization, Writing \u2013 review & editing, Resources.",
  align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

section(doc, "Declaration of competing interest", size=11, space_before=8)
P(doc, "The authors declare that they have no known competing financial interests or "
  "personal relationships that could have appeared to influence the work reported in "
  "this paper.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

section(doc, "Acknowledgements", size=11, space_before=8)
P(doc, "The authors gratefully acknowledge the Department of Chemical Engineering, "
  "NIT Srinagar, for laboratory facilities; the Department of Chemistry for access to "
  "the mixer-grinder; the Department of Civil Engineering for the unconfined "
  "compression testing apparatus; and the staff who facilitated the KD2 Pro thermal "
  "conductivity measurements.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10)

print("Section 4 + declarations added.")

# ================================================================ REFERENCES
section(doc, "References", size=13, space_before=12)
references = [
    "Abral, H., Kadriadi, D., Rodianus, A., Mistar, P., Sapuan, S.M., Zainudin, E.S., "
    "2014. Mechanical properties of water hyacinth fibers\u2013polyester composites "
    "before and after immersion in water. Mater. Des. 58, 125\u2013129. "
    "https://doi.org/10.1016/j.matdes.2014.01.043.",
    "Anjani, A., Iskandar, B., Aziz, M., 2023. The utilization of composite material: "
    "water hyacinth and sugarcane bagasse fiber\u2013epoxy for cool box thermal "
    "insulation. J. Energy Mech. Mater. Manuf. Eng. 8 (1), 29\u201338. "
    "https://doi.org/10.22219/jemmme.v8i1.25010.",
    "Asdrubali, F., D'Alessandro, F., Schiavoni, S., 2015. A review of unconventional "
    "sustainable building insulation materials. Sustain. Mater. Technol. 4, "
    "1\u201317. https://doi.org/10.1016/j.susmat.2015.05.002.",
    "Chaireh, P., Meethong, N., Khomwaen, P., 2020. Novel composite foam made from "
    "starch and water hyacinth with beeswax coating for food packaging applications. "
    "Int. J. Biol. Macromol. 165, 1382\u20131391. "
    "https://doi.org/10.1016/j.ijbiomac.2020.09.243.",
    "Cosentino, L., Fernandes, P., Mateus, R., 2023. A review of natural bio-based "
    "insulation materials. Energies 16 (13), 4926. "
    "https://doi.org/10.3390/en16134926.",
    "Jaktorn, C., Jiajitsawat, S., 2021. Production of thermal insulator from water "
    "hyacinth fiber and natural rubber latex. J. Ecol. Eng. 22 (7), 134\u2013141. "
    "https://doi.org/10.12911/22998993/138736.",
    "Jeon, J., Park, S., Kim, S., 2017. A study on insulation characteristics of glass "
    "wool and mineral wool coated with a polysiloxane agent. Adv. Mater. Sci. Eng. "
    "2017, 3938965. https://doi.org/10.1155/2017/3938965.",
    "Kym\u00e4l\u00e4inen, H.-R., Sj\u00f6berg, A.-M., 2008. Flax and hemp fibres as "
    "raw materials for thermal insulation. Build. Environ. 43 (7), 1261\u20131269. "
    "https://doi.org/10.1016/j.buildenv.2007.03.006.",
    "Oushabi, A., Sair, S., Abboud, Y., Tanane, O., El Bouari, A., 2022. Thermal and "
    "mechanical characterization of alkali-treated sugarcane bagasse-reinforced "
    "thermoset composites. South Afr. J. Chem. Eng. 40, 104\u2013112. "
    "https://doi.org/10.1016/j.sajce.2022.02.006.",
    "Pawlowski, K., Strzalkowska, A., Chojnacka, B., 2025. Exploring advancements in "
    "bio-based composites for thermal insulation: a systematic review. Sustainability "
    "17 (3), 1143. https://doi.org/10.3390/su17031143.",
    "Philip, S., Rakendu, R., 2022. Thermal insulation materials based on water "
    "hyacinth for application in sustainable buildings. Mater. Today Proc. 57, "
    "1863\u20131867. https://doi.org/10.1016/j.matpr.2022.01.062.",
    "Pinto, J., Pereira, E., Tavares, A., Ferreira, V.M., 2021. Corn's cob as a "
    "potential ecological thermal insulation material. Constr. Build. Mater. 277, "
    "122282. https://doi.org/10.1016/j.conbuildmat.2021.122282.",
    "Salas-Ruiz, A., Barbero-Barrera, M. del M., Ruiz-T\u00e9llez, T., 2019. "
    "Microstructural and thermo-physical characterization of a water hyacinth petiole "
    "for thermal insulation particle board manufacture. Materials 12 (4), 560. "
    "https://doi.org/10.3390/ma12040560.",
    "Suwanniroj, P., Suppakarn, N., 2023. Water hyacinth fiber as a bio-based carbon "
    "source for intumescent flame-retardant poly(butylene succinate) composites. "
    "Polymers 15 (21), 4211. https://doi.org/10.3390/polym15214211.",
    "Wang, X., Li, Z., Shi, H., Yu, Y., 2023. Natural pineapple-leaf fibre: a "
    "promising material for high-performance composites. Ind. Crops Prod. 195, "
    "116447. https://doi.org/10.1016/j.indcrop.2023.116447.",
    "Zhou, Y., Trabelsi, A., El Mankibi, M., 2022. Hygrothermal properties of "
    "insulation materials from rice straw and natural binders for buildings. Polymers "
    "14 (9), 1735. https://doi.org/10.3390/polym14091735.",
]
for ref in references:
    pp = doc.add_paragraph()
    pp.paragraph_format.line_spacing = 1.2
    pp.paragraph_format.space_after = Pt(6)
    pp.paragraph_format.left_indent = Inches(0.4)
    pp.paragraph_format.first_line_indent = Inches(-0.4)
    pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = pp.add_run(ref); r.font.name = FONT; r.font.size = Pt(9.5)

print("References added.")

# ---------------------------------------------------------------- save -------
doc.save(OUTPUT)
print("SAVED:", OUTPUT)
