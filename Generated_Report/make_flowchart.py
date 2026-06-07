"""
Recreate the Process Flowsheet as a clean vertical flowchart (Figure 6)
matching the report — based on Process Flowsheet.pdf
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.font_manager as fm

plt.rcParams.update({'font.family': 'DejaVu Serif'})

fig, ax = plt.subplots(figsize=(8.2, 11.4))
ax.set_xlim(0, 10)
ax.set_ylim(0, 24)
ax.axis('off')

# colours
MAIN  = '#2E5A88'
MAINF = '#DCE6F1'
SIDE  = '#4E8542'
SIDEF = '#E2EFDA'
TXT   = '#1a1a1a'

def box(x, y, w, h, title, body, fc, ec, fontsize=8.4, title_fs=9.2):
    p = FancyBboxPatch((x - w/2, y - h/2), w, h,
                       boxstyle="round,pad=0.04,rounding_size=0.12",
                       linewidth=1.4, edgecolor=ec, facecolor=fc, zorder=2)
    ax.add_patch(p)
    ax.text(x, y + h/2 - 0.22, title, ha='center', va='top',
            fontsize=title_fs, fontweight='bold', color=ec, zorder=3)
    ax.text(x, y - 0.08, body, ha='center', va='center',
            fontsize=fontsize, color=TXT, zorder=3, wrap=True)

def arrow(x1, y1, x2, y2, color=MAIN):
    a = FancyArrowPatch((x1, y1), (x2, y2),
                        arrowstyle='-|>', mutation_scale=18,
                        linewidth=1.8, color=color, zorder=1)
    ax.add_patch(a)

# Main vertical chain (centre x = 3.4)
cx = 3.4
W, H = 4.5, 1.7
ys = [22.4, 19.9, 17.4, 14.9, 12.4, 9.9, 7.4, 4.9]

steps = [
    ("S1  RAW MATERIAL COLLECTION",
     "Water hyacinth (Eichhornia crassipes) &\nwater lily (Nymphaea spp.) harvested manually\nfrom Dal Lake, Srinagar. Washed to remove\nmud, debris & contaminants."),
    ("S2  SUN DRYING",
     "Washed biomass sun-dried 5–7 days.\nColour change green → tan/brown\nindicates moisture loss."),
    ("S3  OVEN DRYING",
     "Drying oven at 103 ± 2 °C for ~24 h to\nremove residual moisture and standardise\nmoisture content.  (ASTM D4442)"),
    ("S4  GRINDING",
     "Oven-dried biomass ground in a mixer-grinder\n(Dept. of Chemistry, NIT Srinagar) to\nreduce particle size."),
    ("S5  SIEVING & CLASSIFICATION",
     "Sieved on ASTM E11 sieve shaker:\n• Coarse  > 3 mm (< No. 7 mesh)\n• Fine  1.0–1.5 mm (No. 12–18 mesh)"),
    ("S7  COMPOSITE MIXING & MOULD FILLING",
     "Biomass mixed with binder/water in defined\nratios; packed into 50 × 50 mm steel moulds\n& cylindrical aluminium moulds (D30×H25)."),
    ("S8  COLD PRESSING",
     "Filled moulds cold-pressed by hand to\nconsolidate the mixture, reduce porosity and\nimprove inter-particle bonding."),
    ("S9  DEMOULDING & OVEN CURING",
     "Demoulded and oven-dried at 103 ± 2 °C to\nconstant mass → 4 composite samples\n(S1–S4) ready for characterization."),
]

for y, (t, b) in zip(ys, steps):
    box(cx, y, W, H, t, b, MAINF, MAIN)

for i in range(len(ys) - 1):
    arrow(cx, ys[i] - H/2, cx, ys[i+1] + H/2)

# Parallel side step: Binder preparation (S6) feeding into S7
sx = 7.9
box(sx, 12.4, 3.6, 1.9,
    "S6  BINDER PREPARATION\n(Parallel Step)",
    "Food-grade corn starch gelatinised on a\nhot plate with magnetic stirrer at 80–90 °C,\n150 rpm, until gel forms. Natural binder\nfor S3 (90:10) and S4 (70:30).",
    SIDEF, SIDE, fontsize=7.8, title_fs=8.6)
arrow(sx - 1.8, 12.4, cx + W/2, 9.9 + 0.3, color=SIDE)

# Title
ax.text(5, 23.6, "PROCESS FLOWSHEET",
        ha='center', fontsize=15, fontweight='bold', color=MAIN)
ax.text(5, 23.15,
        "Fabrication of Bio-Composite Insulation Panels from Waste Aquatic Biomass",
        ha='center', fontsize=9, style='italic', color=TXT)

# Footer note
ax.text(5, 3.4,
        "Four samples fabricated:  S1 (coarse, no binder) ·  S2 (fine, no binder) · "
        "S3 (fine 90:10 starch) ·  S4 (fine 70:30 starch)",
        ha='center', fontsize=7.8, style='italic', color='#444')

plt.tight_layout()
fig.savefig('/projects/sandbox/fyp_report/charts/Fig06_Process_Flowsheet.png',
            dpi=300, bbox_inches='tight', pad_inches=0.2)
plt.close()
print("Process flowchart saved")
