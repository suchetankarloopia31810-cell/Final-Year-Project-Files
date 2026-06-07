"""
Recreate the Process Flowsheet as a clean vertical flowchart (Figure 7)
matching the report — based on Process Flowsheet.pdf.
Redesigned for generous spacing so no text overlaps.
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

plt.rcParams.update({'font.family': 'DejaVu Serif'})

# ── canvas ────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(8.6, 12.6))
ax.set_xlim(0, 12)
ax.set_ylim(0, 30)
ax.axis('off')

MAIN, MAINF = '#2E5A88', '#DCE6F1'
SIDE, SIDEF = '#4E8542', '#E2EFDA'
TXT = '#1a1a1a'

# ── geometry ──────────────────────────────────────────────────────────────────
CX   = 4.3            # main column centre x
W    = 6.0            # box width
H    = 2.35           # box height
GAP  = 0.95           # vertical gap between boxes
TOP  = 26.6           # centre y of first box
STEP = H + GAP        # centre-to-centre spacing

def box(cx, cy, w, h, title, body, fc, ec, body_fs=8.6, title_fs=9.8):
    ax.add_patch(FancyBboxPatch(
        (cx - w/2, cy - h/2), w, h,
        boxstyle="round,pad=0.03,rounding_size=0.10",
        linewidth=1.5, edgecolor=ec, facecolor=fc, zorder=2))
    # title near top of box
    ax.text(cx, cy + h/2 - 0.30, title, ha='center', va='center',
            fontsize=title_fs, fontweight='bold', color=ec, zorder=3)
    # body below the title
    ax.text(cx, cy - 0.18, body, ha='center', va='center',
            fontsize=body_fs, color=TXT, zorder=3, linespacing=1.35)

def v_arrow(cy_top, cy_bot, cx=CX, color=MAIN):
    ax.add_patch(FancyArrowPatch(
        (cx, cy_top), (cx, cy_bot),
        arrowstyle='-|>', mutation_scale=20,
        linewidth=2.0, color=color, zorder=1))

# ── main chain (8 steps) ───────────────────────────────────────────────────────
steps = [
    ("S1   RAW MATERIAL COLLECTION",
     "Water hyacinth (Eichhornia crassipes) & water lily\n"
     "(Nymphaea spp.) harvested manually from Dal Lake,\n"
     "Srinagar; washed to remove mud, debris & contaminants."),
    ("S2   SUN DRYING",
     "Washed biomass sun-dried for 5\u20137 days.\n"
     "Colour change green \u2192 tan/brown indicates\nmoisture loss."),
    ("S3   OVEN DRYING   (ASTM D4442)",
     "Oven-dried at 103 \u00b1 2 \u00b0C for \u224824 h to remove\n"
     "residual moisture and standardise moisture content."),
    ("S4   GRINDING",
     "Oven-dried biomass ground in a mixer-grinder\n"
     "(Dept. of Chemistry, NIT Srinagar) to reduce\nparticle size."),
    ("S5   SIEVING & CLASSIFICATION  (ASTM E11)",
     "Sieved on an ASTM E11 sieve shaker into:\n"
     "Coarse > 3 mm (< No. 7)  &  Fine 1.0\u20131.5 mm (No. 12\u201318)."),
    ("S7   COMPOSITE MIXING & MOULD FILLING",
     "Biomass mixed with binder/water in defined ratios;\n"
     "packed into 50 \u00d7 50 mm steel and D30 \u00d7 H25 mm\n"
     "cylindrical aluminium moulds."),
    ("S8   COLD PRESSING",
     "Filled moulds cold-pressed by hand to consolidate\n"
     "the mixture, reduce porosity and improve\ninter-particle bonding."),
    ("S9   DEMOULDING & OVEN CURING",
     "Demoulded and oven-dried to constant mass \u2192\n"
     "4 composite samples (S1\u2013S4) ready for testing."),
]

ys = [TOP - i * STEP for i in range(len(steps))]
for cy, (t, b) in zip(ys, steps):
    box(CX, cy, W, H, t, b, MAINF, MAIN)
for i in range(len(ys) - 1):
    v_arrow(ys[i] - H/2, ys[i+1] + H/2)

# ── parallel binder-prep box (S6) feeding into S7 (mixing) ─────────────────────
# S7 is index 5
y_mix = ys[5]
SX, SW, SH = 9.7, 3.6, 2.7
box(SX, y_mix, SW, SH,
    "S6   BINDER PREPARATION\n(Parallel Step)",
    "Food-grade corn starch\ngelatinised on a hot plate with\n"
    "magnetic stirrer at 80\u201390 \u00b0C\nuntil a gel forms. Natural\n"
    "binder for S3 (90:10) & S4 (70:30).",
    SIDEF, SIDE, body_fs=8.0, title_fs=9.0)
# horizontal arrow from binder box into the mixing box
ax.add_patch(FancyArrowPatch(
    (SX - SW/2, y_mix), (CX + W/2, y_mix),
    arrowstyle='-|>', mutation_scale=18,
    linewidth=1.8, color=SIDE, zorder=1))

# ── titles & footer ────────────────────────────────────────────────────────────
ax.text(6, 29.2, "PROCESS FLOWSHEET", ha='center',
        fontsize=16, fontweight='bold', color=MAIN)
ax.text(6, 28.5,
        "Fabrication of Bio-Composite Insulation Panels from Waste Aquatic Biomass",
        ha='center', fontsize=9.5, style='italic', color=TXT)

foot_y = ys[-1] - H/2 - 0.9
ax.text(6, foot_y,
        "Four samples fabricated:   S1 (coarse, no binder)   \u00b7   "
        "S2 (fine, no binder)   \u00b7   S3 (fine 90:10 starch)   \u00b7   "
        "S4 (fine 70:30 starch)",
        ha='center', fontsize=8.4, style='italic', color='#444')

fig.savefig('/projects/sandbox/fyp_report/charts/Fig06_Process_Flowsheet.png',
            dpi=300, bbox_inches='tight', pad_inches=0.25)
plt.close()
print("Process flowchart regenerated (no overlap).")
