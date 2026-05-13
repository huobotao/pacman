#!/usr/bin/env python3
"""
BarCo Corporation Supply Chain Presentation Generator
Generates: PPTX + PDF presentation, and PDF speech script
"""

import os, io, textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 HRFlowable, Table, TableStyle, KeepTogether)
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib import colors as rl_colors

OUT  = "/home/user/pacman/presentation"
TMP  = "/tmp/barco_assets"
os.makedirs(OUT, exist_ok=True)
os.makedirs(TMP, exist_ok=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  COLOR PALETTE
# ═══════════════════════════════════════════════════════════════════════════════
NAVY      = RGBColor(0x08, 0x14, 0x29)
NAVY2     = RGBColor(0x0E, 0x22, 0x44)
BLUE      = RGBColor(0x1B, 0x4F, 0x8A)
GOLD      = RGBColor(0xF0, 0xAE, 0x1C)
TEAL      = RGBColor(0x0D, 0x92, 0x87)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE  = RGBColor(0xF4, 0xF7, 0xFD)
DARK      = RGBColor(0x12, 0x1C, 0x2C)
GREY      = RGBColor(0x5E, 0x6E, 0x86)
RED       = RGBColor(0xC6, 0x2B, 0x2B)
GREEN     = RGBColor(0x18, 0x87, 0x5C)
ORANGE    = RGBColor(0xE4, 0x7E, 0x10)
ROW_ODD   = RGBColor(0xEB, 0xF1, 0xFA)
ROW_EVEN  = RGBColor(0xFF, 0xFF, 0xFF)

NAVY_H    = '#081429'
BLUE_H    = '#1B4F8A'
GOLD_H    = '#F0AE1C'
TEAL_H    = '#0D9287'
RED_H     = '#C62B2B'
GREEN_H   = '#18875C'
ORANGE_H  = '#E47E10'
GREY_H    = '#5E6E86'
WHITE_H   = '#FFFFFF'
OFFWHITE_H= '#F4F7FD'

W = Inches(13.33)
H = Inches(7.5)

# ═══════════════════════════════════════════════════════════════════════════════
#  MATPLOTLIB CHART GENERATORS
# ═══════════════════════════════════════════════════════════════════════════════

def make_profit_per_gas_chart():
    products = ['AP\n(Ammonium\nPhosphate)', 'UR\n(Urea)', 'CL\n(Chlorine)',
                'AN\n(Ammonium\nNitrate)', 'CS\n(Caustic\nSoda)',
                'VCM\n(Vinyl Chloride\nMonomer)', 'AM\n(Ammonia)',
                'HF\n(Hydrofluoric\nAcid)']
    values   = [14.00, 10.36, 8.46, 6.67, 6.25, 5.94, 5.83, 5.59]
    colors_b = [GREEN_H, GREEN_H, TEAL_H, TEAL_H, ORANGE_H, RED_H, RED_H, RED_H]

    fig, ax = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor(OFFWHITE_H)
    ax.set_facecolor(OFFWHITE_H)

    bars = ax.bar(range(len(products)), values, color=colors_b,
                  edgecolor='white', linewidth=1.2, width=0.65)
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.12,
                f'${val:.2f}', ha='center', va='bottom',
                fontsize=9, fontweight='bold', color=NAVY_H)

    ax.set_xticks(range(len(products)))
    ax.set_xticklabels(products, fontsize=7.5, color=NAVY_H)
    ax.set_ylabel('Profit per 1,000 cu.ft of Gas ($)', fontsize=10, color=NAVY_H)
    ax.set_title('Product Ranking by Profit per Unit of Gas\n(LP Prioritisation Order)',
                 fontsize=12, fontweight='bold', color=NAVY_H)
    ax.set_ylim(0, 16)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cccccc')
    ax.spines['bottom'].set_color('#cccccc')
    ax.yaxis.set_tick_params(labelcolor=NAVY_H)

    # Shade zones
    ax.axhspan(0, 6.5, alpha=0.08, color=RED_H, label='Cut zone')
    ax.axhspan(6.5, 16, alpha=0.06, color=GREEN_H, label='Protected zone')
    ax.legend(loc='upper right', fontsize=8)

    plt.tight_layout()
    path = f"{TMP}/profit_per_gas.png"
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=OFFWHITE_H)
    plt.close()
    return path


def make_production_results_chart():
    scenarios = ['Baseline\n(0%)', '20%\nCurtailment', '30%\nCurtailment',
                 '40%\nCurtailment', '50%\nCurtailment']
    # tons/day stacked: AP, UR, CL, AN, CS, VCM, AM, HF
    AP  = [510, 510, 510, 510, 510]
    UR  = [150, 150, 150, 150, 150]
    CL  = [1350,1350,1350,1350,1350]
    AN  = [630, 630, 630, 630, 630]
    CS  = [1280,1280,1280,1280,720]
    VCM = [980, 980, 651, 115.5, 0]
    AM  = [1050,275.3,0,0,0]
    HF  = [461.2,0,0,0,0]

    fig, ax = plt.subplots(figsize=(11, 5.5))
    fig.patch.set_facecolor(OFFWHITE_H)
    ax.set_facecolor(OFFWHITE_H)

    x = np.arange(len(scenarios))
    bottom = np.zeros(len(scenarios))
    palette = [GREEN_H, '#2ECC71', TEAL_H, '#3498DB', ORANGE_H, RED_H, '#9B59B6', '#E74C3C']
    labels  = ['Ammonium Phosphate','Urea','Chlorine','Ammonium Nitrate',
               'Caustic Soda','Vinyl Chloride Monomer','Ammonia','Hydrofluoric Acid']
    for data, color, label in zip([AP,UR,CL,AN,CS,VCM,AM,HF], palette, labels):
        ax.bar(x, data, bottom=bottom, color=color, label=label,
               edgecolor='white', linewidth=0.8)
        bottom += np.array(data)

    # Profit labels
    profits = [611062, 513023, 462495, 411623, 358650]
    for i, (prof, tot) in enumerate(zip(profits, bottom)):
        ax.text(i, tot + 40, f'${prof:,.0f}/day',
                ha='center', va='bottom', fontsize=8, fontweight='bold', color=NAVY_H)

    ax.set_xticks(x)
    ax.set_xticklabels(scenarios, fontsize=9, color=NAVY_H)
    ax.set_ylabel('Production Volume (tons/day)', fontsize=10, color=NAVY_H)
    ax.set_title('Optimal Production Mix by Curtailment Scenario',
                 fontsize=12, fontweight='bold', color=NAVY_H)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.legend(loc='upper right', fontsize=7, ncol=2, framealpha=0.9)

    plt.tight_layout()
    path = f"{TMP}/production_results.png"
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=OFFWHITE_H)
    plt.close()
    return path


def make_profit_curve_chart():
    gas_pct  = [100, 90, 80, 70, 60, 50, 40]
    profits  = [611062, 579200, 547338, 515476, 462495, 411623, 358650]

    fig, ax = plt.subplots(figsize=(9, 4.5))
    fig.patch.set_facecolor(OFFWHITE_H)
    ax.set_facecolor(OFFWHITE_H)

    ax.plot(gas_pct, profits, color=NAVY_H, linewidth=2.5, marker='o',
            markersize=7, markerfacecolor=GOLD_H, markeredgecolor=NAVY_H)
    ax.fill_between(gas_pct, profits, alpha=0.12, color=BLUE_H)

    for x, y in zip(gas_pct, profits):
        ax.annotate(f'${y:,.0f}', (x, y), textcoords="offset points",
                    xytext=(4, 8), fontsize=7.5, color=NAVY_H)

    ax.set_xlabel('Gas Budget (% of Baseline)', fontsize=10, color=NAVY_H)
    ax.set_ylabel('Daily Profit ($)', fontsize=10, color=NAVY_H)
    ax.set_title('Profit Sensitivity to Gas Availability\n(Piecewise-Linear LP Curve)',
                 fontsize=11, fontweight='bold', color=NAVY_H)
    ax.set_xlim(38, 102)
    ax.invert_xaxis()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'${x:,.0f}'))
    ax.tick_params(colors=NAVY_H)

    plt.tight_layout()
    path = f"{TMP}/profit_curve.png"
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=OFFWHITE_H)
    plt.close()
    return path


def make_distribution_chart():
    fig, axes = plt.subplots(1, 3, figsize=(13, 4))
    fig.patch.set_facecolor(OFFWHITE_H)

    scenarios_data = [
        ('Baseline (0%)\nTotal: $113,952/day',
         [[0, 20000, 0], [20000, 20000, 0], [0, 10000, 15680]]),
        ('30% Curtailment\nTotal: $79,766/day',
         [[0, 14000, 0], [14000, 14000, 0], [0, 7000, 10976]]),
        ('50% Curtailment\nTotal: $56,976/day',
         [[0, 10000, 0], [10000, 10000, 0], [0, 5000, 7840]]),
    ]
    plants  = ['Plant A', 'Plant B', 'Plant C']
    sources = ['Source I', 'Source J', 'Source K']

    for ax, (title, data) in zip(axes, scenarios_data):
        ax.set_facecolor(OFFWHITE_H)
        mat = np.array(data, dtype=float)
        max_val = 20000 if mat.max() == 0 else mat.max()

        im = ax.imshow(mat, cmap='Blues', vmin=0, vmax=20000, aspect='auto')
        ax.set_xticks(range(3)); ax.set_xticklabels(plants, fontsize=8)
        ax.set_yticks(range(3)); ax.set_yticklabels(sources, fontsize=8)
        ax.set_title(title, fontsize=8.5, fontweight='bold', color=NAVY_H)

        for i in range(3):
            for j in range(3):
                val = int(mat[i, j])
                txt = f'{val:,}' if val > 0 else '—'
                color = 'white' if mat[i, j] > 10000 else NAVY_H
                ax.text(j, i, txt, ha='center', va='center',
                        fontsize=8.5, fontweight='bold', color=color)

    fig.suptitle('Optimal Gas Distribution Plan (Units: 10³ cu.ft/day)',
                 fontsize=11, fontweight='bold', color=NAVY_H, y=1.02)
    plt.tight_layout()
    path = f"{TMP}/distribution_heatmap.png"
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=OFFWHITE_H)
    plt.close()
    return path


def make_co2_chart():
    profits  = [611062, 549956, 488849, 427743, 366636, 305530]
    co2_vals = [5.74, 4.98, 4.26, 3.57, 2.89, 2.22]
    floors   = ['100%', '90%', '80%', '70%', '60%', '50%']

    fig, ax = plt.subplots(figsize=(9, 4.5))
    fig.patch.set_facecolor(OFFWHITE_H)
    ax.set_facecolor(OFFWHITE_H)

    scatter = ax.scatter(profits, co2_vals, c=co2_vals, cmap='RdYlGn_r',
                         s=120, zorder=5, edgecolors=NAVY_H, linewidths=0.8)
    ax.plot(profits, co2_vals, color=TEAL_H, linewidth=2, alpha=0.7, zorder=4)

    for x, y, label in zip(profits, co2_vals, floors):
        ax.annotate(f'Floor={label}', (x, y), textcoords="offset points",
                    xytext=(6, 6), fontsize=8, color=NAVY_H)

    ax.axhline(y=6.0, color=RED_H, linestyle='--', linewidth=1.5, alpha=0.7,
               label='Baseline CO₂ = 6.0 kg/day')
    ax.set_xlabel('Daily Profit ($)', fontsize=10, color=NAVY_H)
    ax.set_ylabel('Production CO₂ Emissions (kg/day)', fontsize=10, color=NAVY_H)
    ax.set_title('Profit–Emissions Pareto Frontier\n(10% Profit Sacrifice → 17% CO₂ Reduction)',
                 fontsize=11, fontweight='bold', color=NAVY_H)
    ax.legend(fontsize=9)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'${x/1000:.0f}K'))
    ax.tick_params(colors=NAVY_H)

    plt.tight_layout()
    path = f"{TMP}/co2_pareto.png"
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=OFFWHITE_H)
    plt.close()
    return path


# ═══════════════════════════════════════════════════════════════════════════════
#  PPTX HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_rect(slide, x, y, w, h, fill_rgb, line_rgb=None, line_w=Pt(0)):
    shp = slide.shapes.add_shape(1, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shp.line.color.rgb = line_rgb
        shp.line.width = line_w
    else:
        shp.line.fill.background()
    return shp


def add_text(slide, text, x, y, w, h,
             size=Pt(12), color=DARK, bold=False, italic=False,
             align=PP_ALIGN.LEFT, wrap=True, font="Calibri Light"):
    box = slide.shapes.add_textbox(x, y, w, h)
    box.word_wrap = wrap
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    return box, tf


def set_para_space(para, space_before=0, space_after=0):
    """Set paragraph spacing in points."""
    pPr = para._p.get_or_add_pPr()
    if space_before:
        spc = etree.SubElement(pPr, qn('a:spcBef'))
        spcPts = etree.SubElement(spc, qn('a:spcPts'))
        spcPts.set('val', str(int(space_before * 100)))
    if space_after:
        spc = etree.SubElement(pPr, qn('a:spcAft'))
        spcPts = etree.SubElement(spc, qn('a:spcPts'))
        spcPts.set('val', str(int(space_after * 100)))


def dark_header(slide, title_text, subtitle_text=None):
    """Add a dark navy full-width header bar with title."""
    # Background gradient: dark nav overlay at top
    add_rect(slide, 0, 0, W, Inches(1.15), NAVY)
    # Gold accent line at bottom of header
    add_rect(slide, 0, Inches(1.15), W, Inches(0.045), GOLD)

    title_box, tf = add_text(
        slide, title_text,
        Inches(0.45), Inches(0.13), W - Inches(1.0), Inches(0.78),
        size=Pt(28), color=WHITE, bold=True,
        align=PP_ALIGN.LEFT, font="Calibri"
    )
    if subtitle_text:
        add_text(
            slide, subtitle_text,
            Inches(0.45), Inches(0.83), W - Inches(1.0), Inches(0.35),
            size=Pt(13), color=RGBColor(0xB8, 0xC8, 0xE4),
            bold=False, align=PP_ALIGN.LEFT, font="Calibri Light"
        )


def slide_background(slide):
    """Light off-white background for content slides."""
    add_rect(slide, 0, 0, W, H, OFFWHITE)
    # Left accent stripe
    add_rect(slide, 0, Inches(1.2), Inches(0.06), H - Inches(1.2), GOLD)


def slide_number(slide, num, total=15):
    """Add slide number in bottom right."""
    add_text(slide, f"{num} / {total}",
             W - Inches(1.2), H - Inches(0.4), Inches(1.0), Inches(0.35),
             size=Pt(9), color=GREY, align=PP_ALIGN.RIGHT, font="Calibri Light")


def add_bullet_box(slide, items, x, y, w, h,
                   bullet="▸", color=DARK, size=Pt(13),
                   sub_items=None, font="Calibri Light"):
    """Add a text box with bullet points."""
    box = slide.shapes.add_textbox(x, y, w, h)
    box.word_wrap = True
    tf = box.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"{bullet}  {item}"
        run.font.size = size
        run.font.color.rgb = color
        run.font.name = font
        set_para_space(p, space_before=3)

        if sub_items and sub_items.get(i):
            for sub in sub_items[i]:
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.LEFT
                r2 = p2.add_run()
                r2.text = f"    –  {sub}"
                r2.font.size = Pt(size.pt - 2 if hasattr(size, 'pt') else 11)
                r2.font.color.rgb = GREY
                r2.font.name = font
    return box, tf


def add_kpi_box(slide, value, label, x, y, bg_color=NAVY, val_color=GOLD, lbl_color=WHITE):
    """Add a KPI/stat box."""
    add_rect(slide, x, y, Inches(2.5), Inches(1.2), bg_color)
    add_text(slide, value, x + Inches(0.1), y + Inches(0.08),
             Inches(2.3), Inches(0.65),
             size=Pt(28), color=val_color, bold=True,
             align=PP_ALIGN.CENTER, font="Calibri")
    add_text(slide, label, x + Inches(0.1), y + Inches(0.72),
             Inches(2.3), Inches(0.42),
             size=Pt(9.5), color=lbl_color, bold=False,
             align=PP_ALIGN.CENTER, font="Calibri Light")


def add_table_manual(slide, headers, rows, x, y, col_widths,
                     row_height=Inches(0.38)):
    """Render a table using colored rectangles and text."""
    # Header row
    cx = x
    for i, (hdr, cw) in enumerate(zip(headers, col_widths)):
        add_rect(slide, cx, y, cw, row_height, NAVY)
        add_text(slide, hdr, cx + Inches(0.06), y + Inches(0.07),
                 cw - Inches(0.1), row_height - Inches(0.1),
                 size=Pt(9.5), color=WHITE, bold=True,
                 align=PP_ALIGN.CENTER, font="Calibri")
        cx += cw

    # Data rows
    for r_idx, row in enumerate(rows):
        ry = y + row_height * (r_idx + 1)
        fill = ROW_ODD if r_idx % 2 == 0 else ROW_EVEN
        cx = x
        for c_idx, (cell, cw) in enumerate(zip(row, col_widths)):
            add_rect(slide, cx, ry, cw, row_height, fill,
                     line_rgb=RGBColor(0xD0, 0xD8, 0xE8), line_w=Pt(0.5))
            # Determine color for special cells
            cell_color = DARK
            cell_bold = False
            if str(cell).startswith('$') and c_idx > 0:
                cell_color = GREEN if '611' in str(cell) else (
                    ORANGE if '513' in str(cell) or '462' in str(cell) else RED)
                cell_bold = True
            elif str(cell) == '0' or str(cell) == '0.0':
                cell_color = RGBColor(0xAA, 0xAA, 0xAA)
            add_text(slide, str(cell), cx + Inches(0.05), ry + Inches(0.07),
                     cw - Inches(0.08), row_height - Inches(0.1),
                     size=Pt(9), color=cell_color, bold=cell_bold,
                     align=PP_ALIGN.CENTER, font="Calibri")
            cx += cw


# ═══════════════════════════════════════════════════════════════════════════════
#  INDIVIDUAL SLIDE BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def slide_title(prs):
    sl = blank_slide(prs)
    # Full dark background
    add_rect(sl, 0, 0, W, H, NAVY)
    # Gold top accent bar
    add_rect(sl, 0, 0, W, Inches(0.08), GOLD)
    # Teal bottom accent bar
    add_rect(sl, 0, H - Inches(0.08), W, Inches(0.08), TEAL)
    # Right decorative panel
    add_rect(sl, W - Inches(3.2), 0, Inches(3.2), H, NAVY2)
    add_rect(sl, W - Inches(3.2), 0, Inches(0.05), H, GOLD)

    # Decorative circles (subtle)
    add_rect(sl, W - Inches(2.8), Inches(1.2), Inches(2.0), Inches(2.0),
             RGBColor(0x10, 0x22, 0x44))

    # Main title
    add_text(sl, "BarCo Corporation",
             Inches(0.7), Inches(1.4), Inches(8.8), Inches(0.85),
             size=Pt(40), color=GOLD, bold=True,
             align=PP_ALIGN.LEFT, font="Calibri")

    add_text(sl, "Supply Chain Optimisation",
             Inches(0.7), Inches(2.22), Inches(8.8), Inches(0.75),
             size=Pt(34), color=WHITE, bold=True,
             align=PP_ALIGN.LEFT, font="Calibri")

    add_text(sl, "Under Natural Gas Curtailment",
             Inches(0.7), Inches(2.95), Inches(8.8), Inches(0.65),
             size=Pt(28), color=RGBColor(0xB8, 0xCE, 0xEB), bold=False,
             align=PP_ALIGN.LEFT, font="Calibri Light")

    # Gold separator line
    add_rect(sl, Inches(0.7), Inches(3.72), Inches(7.5), Inches(0.05), GOLD)

    add_text(sl, "A Linear Programming Approach to Production and Distribution Planning",
             Inches(0.7), Inches(3.85), Inches(8.8), Inches(0.45),
             size=Pt(13), color=RGBColor(0x90, 0xAC, 0xCC), bold=False,
             align=PP_ALIGN.LEFT, font="Calibri Light")

    add_text(sl, "49680 Value Chain Engineering Systems  ·  Autumn 2026  ·  UTS FEIT",
             Inches(0.7), Inches(5.5), Inches(8.8), Inches(0.35),
             size=Pt(11), color=RGBColor(0x70, 0x90, 0xB8),
             align=PP_ALIGN.LEFT, font="Calibri Light")

    add_text(sl, "Group Z  ·  Botao Huo (25104579)  ·  Shah Misbabun Nur Rupom (26042287)",
             Inches(0.7), Inches(5.92), Inches(8.8), Inches(0.35),
             size=Pt(11), color=RGBColor(0x70, 0x90, 0xB8),
             align=PP_ALIGN.LEFT, font="Calibri Light")

    # Right panel labels
    add_text(sl, "VCES\n2026",
             W - Inches(2.9), Inches(3.0), Inches(2.5), Inches(1.5),
             size=Pt(42), color=RGBColor(0x1A, 0x34, 0x5E), bold=True,
             align=PP_ALIGN.CENTER, font="Calibri")

    return sl


def slide_overview(prs):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Presentation Overview", "4 key sections — aligned with assessment criteria")
    slide_number(sl, 2)

    boxes = [
        ("01", "Business Problem", "Natural gas curtailment context,\nfederal priority rules & financial risk",
         NAVY, GOLD),
        ("02", "Analysis Approach", "Linear Programming formulation,\nmodel assumptions & methodology",
         BLUE, WHITE),
        ("03", "Key Results", "Production plan, distribution plan,\nsensitivity analysis & CO₂ scenario",
         TEAL, WHITE),
        ("04", "Recommendations", "Three evidence-based strategic\nrecommendations for management",
         RGBColor(0x8B, 0x2B, 0x8B), WHITE),
    ]

    for idx, (num, title, desc, bg_col, fg_col) in enumerate(boxes):
        col = idx % 2
        row = idx // 2
        bx = Inches(0.7) + col * Inches(6.2)
        by = Inches(1.6) + row * Inches(2.55)

        add_rect(sl, bx, by, Inches(5.8), Inches(2.3), bg_col)
        # Number accent
        # Slightly darker shade for the number column
        add_rect(sl, bx, by, Inches(0.9), Inches(2.3),
                 RGBColor(max(0, bg_col[0] - 20),
                          max(0, bg_col[1] - 20),
                          max(0, bg_col[2] - 20)))
        add_text(sl, num, bx + Inches(0.04), by + Inches(0.65), Inches(0.82), Inches(1.0),
                 size=Pt(30), color=GOLD if bg_col == NAVY else RGBColor(0xFF, 0xFF, 0xAA),
                 bold=True, align=PP_ALIGN.CENTER, font="Calibri")
        add_text(sl, title, bx + Inches(1.02), by + Inches(0.18),
                 Inches(4.6), Inches(0.52),
                 size=Pt(17), color=fg_col, bold=True,
                 align=PP_ALIGN.LEFT, font="Calibri")
        add_text(sl, desc, bx + Inches(1.02), by + Inches(0.72),
                 Inches(4.6), Inches(1.4),
                 size=Pt(11.5), color=RGBColor(0xCC, 0xDD, 0xEE) if fg_col == WHITE else RGBColor(0xDD, 0xEE, 0xFF),
                 bold=False, align=PP_ALIGN.LEFT, font="Calibri Light")

    return sl


def slide_business_problem(prs):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "The Business Problem",
                "BarCo Corporation — Natural Gas Curtailment Crisis")
    slide_number(sl, 3)

    # Left column
    add_text(sl, "Company Profile",
             Inches(0.6), Inches(1.45), Inches(5.8), Inches(0.38),
             size=Pt(14), color=NAVY, bold=True, font="Calibri")
    add_rect(sl, Inches(0.6), Inches(1.83), Inches(5.8), Inches(0.04), GOLD)

    bullets_left = [
        "8 chemical product lines",
        "3 manufacturing complexes (A, B, C)",
        "Gas supplier: MaxEnergy (nodes I, J, K)",
        "Current gas usage: 85,680 × 10³ cu.ft/day",
        "Ammonia uses gas as raw material (2nd Priority)",
        "All other products use gas as boiler fuel (3rd Priority)",
    ]
    add_bullet_box(sl, bullets_left,
                   Inches(0.72), Inches(1.92), Inches(5.7), Inches(2.8),
                   size=Pt(11.5), color=DARK)

    # Right column — Crisis
    add_text(sl, "The Crisis",
             Inches(7.0), Inches(1.45), Inches(5.9), Inches(0.38),
             size=Pt(14), color=NAVY, bold=True, font="Calibri")
    add_rect(sl, Inches(7.0), Inches(1.83), Inches(5.9), Inches(0.04), RED)

    bullets_right = [
        "MaxEnergy warns of natural gas shortages",
        "Cause: extreme heat wave — power plants at capacity",
        "Rolling brownouts: 20–40% curtailment expected",
        "Curtailment based on actual usage (85,680), NOT contract max (90,000)",
        "MaxEnergy will NOT specify which products to cut",
        "BarCo has flexibility to absorb cuts strategically",
    ]
    add_bullet_box(sl, bullets_right,
                   Inches(7.12), Inches(1.92), Inches(5.7), Inches(2.8),
                   size=Pt(11.5), color=DARK, bullet="!")

    # Priority rules box at bottom
    add_rect(sl, Inches(0.55), Inches(5.0), W - Inches(1.1), Inches(1.9),
             RGBColor(0xE8, 0xF0, 0xFA))
    add_rect(sl, Inches(0.55), Inches(5.0), Inches(0.06), Inches(1.9), BLUE)

    add_text(sl, "Federal Commission Priority Rules",
             Inches(0.72), Inches(5.08), Inches(4.0), Inches(0.36),
             size=Pt(12), color=NAVY, bold=True, font="Calibri")

    priority_items = [
        ("1st Priority", "Residential & commercial heating/cooling", GREEN),
        ("2nd Priority", "Industrial raw material users  ← BarCo Ammonia", ORANGE),
        ("3rd Priority", "Industrial boiler fuel users  ← BarCo (all other products)", RED),
    ]
    for i, (label, desc, col) in enumerate(priority_items):
        add_rect(sl, Inches(0.72) + i * Inches(4.1), Inches(5.5),
                 Inches(0.95), Inches(0.28), col)
        add_text(sl, label,
                 Inches(0.72) + i * Inches(4.1), Inches(5.5),
                 Inches(0.93), Inches(0.28),
                 size=Pt(8.5), color=WHITE, bold=True,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(sl, desc,
                 Inches(1.72) + i * Inches(4.1), Inches(5.5),
                 Inches(2.9), Inches(0.28),
                 size=Pt(9.5), color=DARK, font="Calibri Light")

    return sl


def slide_financial_risk(prs):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Financial Risk at a Glance",
                "What is the potential daily profit impact without management action?")
    slide_number(sl, 4)

    # KPI boxes
    kpis = [
        ("$611,062", "Baseline Daily Profit", NAVY, GOLD),
        ("$513,023", "Profit at 20% Curtailment\n(−16.0%)", BLUE, WHITE),
        ("$411,623", "Profit at 40% Curtailment\n(−32.6%)", ORANGE, WHITE),
        ("$358,650", "Profit at 50% Curtailment\n(−41.3%)", RED, WHITE),
    ]
    for i, (val, lbl, bg, fg) in enumerate(kpis):
        bx = Inches(0.55) + i * Inches(3.2)
        add_rect(sl, bx, Inches(1.5), Inches(3.0), Inches(1.5), bg)
        add_text(sl, val, bx + Inches(0.1), Inches(1.58),
                 Inches(2.8), Inches(0.7),
                 size=Pt(22), color=fg if i == 0 else (GOLD if bg != ORANGE else WHITE),
                 bold=True, align=PP_ALIGN.CENTER, font="Calibri")
        add_text(sl, lbl, bx + Inches(0.08), Inches(2.3),
                 Inches(2.85), Inches(0.65),
                 size=Pt(9.5), color=WHITE, align=PP_ALIGN.CENTER, font="Calibri Light")

    # Key insight box
    add_rect(sl, Inches(0.55), Inches(3.25), W - Inches(1.1), Inches(1.05),
             RGBColor(0xFD, 0xF0, 0xC2))
    add_rect(sl, Inches(0.55), Inches(3.25), Inches(0.07), Inches(1.05), GOLD)
    add_text(sl, "⚡  Key Insight:  Profit drops LESS than proportionally to gas "
             "— a 50% gas cut causes only a 41.3% profit drop. "
             "This is the value of smart prioritisation via LP.",
             Inches(0.72), Inches(3.32), W - Inches(1.3), Inches(0.9),
             size=Pt(12.5), color=RGBColor(0x5A, 0x3A, 0x00), bold=False,
             font="Calibri")

    # LP value callout
    add_rect(sl, Inches(0.55), Inches(4.5), W - Inches(1.1), Inches(2.55),
             RGBColor(0xE8, 0xF5, 0xF0))
    add_rect(sl, Inches(0.55), Inches(4.5), Inches(0.07), Inches(2.55), GREEN)

    add_text(sl, "Value of LP-Driven Prioritisation vs. Naive Proportional Cut",
             Inches(0.72), Inches(4.58), Inches(11.0), Inches(0.4),
             size=Pt(13), color=NAVY, bold=True, font="Calibri")
    add_bullet_box(sl, [
        "At 50% curtailment: LP saves ~$44,000/day vs. proportional cut",
        "Annualised value: ~$16 million per year",
        "The LP identifies exactly which products to protect and which to cut — the optimal answer, not a guess",
    ],
    Inches(0.72), Inches(5.05), W - Inches(1.3), Inches(1.8),
    size=Pt(12), color=DARK, bullet="✓")

    return sl


def slide_approach(prs):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Analysis Approach: Linear Programming",
                "Why LP? — The natural mathematical formalism for this problem")
    slide_number(sl, 5)

    # Why LP box
    add_rect(sl, Inches(0.55), Inches(1.45), Inches(5.9), Inches(5.6),
             RGBColor(0xF0, 0xF4, 0xFB))
    add_rect(sl, Inches(0.55), Inches(1.45), Inches(0.07), Inches(5.6), BLUE)

    add_text(sl, "Why Linear Programming?",
             Inches(0.72), Inches(1.55), Inches(5.6), Inches(0.4),
             size=Pt(14), color=NAVY, bold=True, font="Calibri")

    why_items = [
        "Linear objective: maximize profit (sum of profit × volume)",
        "Linear constraints: gas budget, capacity limits",
        "Continuous decision variables: production in tons/day",
        "Proven global optimal guarantee via Simplex algorithm",
        "Transparent, auditable, repeatable decisions",
        "Sensitivity analysis built in — shows impact of uncertainty",
    ]
    add_bullet_box(sl, why_items,
                   Inches(0.72), Inches(2.02), Inches(5.7), Inches(4.8),
                   size=Pt(11.5), color=DARK)

    # Two models
    add_rect(sl, Inches(6.7), Inches(1.45), Inches(6.2), Inches(2.6),
             RGBColor(0xE8, 0xF5, 0xF2))
    add_rect(sl, Inches(6.7), Inches(1.45), Inches(0.07), Inches(2.6), TEAL)
    add_text(sl, "Model 1: Production LP",
             Inches(6.88), Inches(1.55), Inches(5.9), Inches(0.4),
             size=Pt(13), color=NAVY, bold=True, font="Calibri")
    add_text(sl, "Maximise  Z = Σⱼ pⱼ · xⱼ\nSubject to:  Σⱼ gⱼ · xⱼ ≤ G  (gas budget)\n"
             "               0 ≤ xⱼ ≤ capⱼ  (capacity)",
             Inches(6.88), Inches(2.02), Inches(5.9), Inches(1.85),
             size=Pt(11.5), color=DARK, font="Courier New")

    add_rect(sl, Inches(6.7), Inches(4.25), Inches(6.2), Inches(2.8),
             RGBColor(0xF0, 0xE8, 0xF8))
    add_rect(sl, Inches(6.7), Inches(4.25), Inches(0.07), Inches(2.8),
             RGBColor(0x8B, 0x2B, 0x8B))
    add_text(sl, "Model 2: Distribution (Transportation) LP",
             Inches(6.88), Inches(4.35), Inches(5.9), Inches(0.4),
             size=Pt(13), color=NAVY, bold=True, font="Calibri")
    add_text(sl, "Minimise  T = ΣᵢΣⱼ cᵢⱼ · yᵢⱼ\nSubject to:  Σⱼ yᵢⱼ ≤ sᵢ  (supply)\n"
             "               Σᵢ yᵢⱼ = dⱼ  (demand)\n               yᵢⱼ ≥ 0",
             Inches(6.88), Inches(4.82), Inches(5.9), Inches(2.0),
             size=Pt(11.5), color=DARK, font="Courier New")

    return sl


def slide_assumptions(prs):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Key Model Assumptions",
                "Conditions under which the LP solution is valid")
    slide_number(sl, 6)

    assumptions = [
        ("Linearity", "Profit per ton, gas per ton, and transport cost per 10³ cu.ft "
         "are constant throughout the planning period — no economies of scale or diminishing returns.",
         TEAL),
        ("Divisibility", "Production volumes are continuous real numbers. "
         "Rounding to integer tons is operationally trivial at BarCo's scale of output.",
         BLUE),
        ("Certainty", "All profit, gas, and cost coefficients are known with certainty. "
         "Sensitivity analysis (Slide 11) relaxes this assumption for profit coefficients.",
         ORANGE),
        ("Independent Capacities", "Each product line operates independently up to its effective capacity. "
         "No cross-product machine-sharing or interlocks within a plant.",
         GREEN),
        ("Uniform Curtailment", "Gas reduction applies equally and proportionally "
         "to all three supply nodes (I, J, K), per MaxEnergy notification.",
         RGBColor(0x8B, 0x2B, 0x8B)),
        ("Single-Period", "The LP optimises a representative single day. "
         "Multi-period inventory dynamics are noted as a limitation.",
         RED),
    ]

    for i, (name, desc, color) in enumerate(assumptions):
        col = i % 2
        row = i // 2
        bx = Inches(0.55) + col * Inches(6.4)
        by = Inches(1.45) + row * Inches(1.9)

        add_rect(sl, bx, by, Inches(6.1), Inches(1.75), ROW_ODD)
        add_rect(sl, bx, by, Inches(0.07), Inches(1.75), color)
        add_text(sl, name, bx + Inches(0.18), by + Inches(0.1),
                 Inches(5.8), Inches(0.4),
                 size=Pt(12.5), color=NAVY, bold=True, font="Calibri")
        add_text(sl, desc, bx + Inches(0.18), by + Inches(0.52),
                 Inches(5.8), Inches(1.1),
                 size=Pt(10.5), color=DARK, font="Calibri Light")

    return sl


def slide_product_economics(prs, chart_path):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Product Economics: The LP's Prioritisation Logic",
                "Profit per unit of gas consumed — this ranking determines the optimal cut sequence")
    slide_number(sl, 7)

    # Embed chart
    sl.shapes.add_picture(chart_path, Inches(0.4), Inches(1.35), Inches(9.0), Inches(4.2))

    # Right column insight
    add_rect(sl, Inches(9.6), Inches(1.4), Inches(3.45), Inches(5.65),
             RGBColor(0xF0, 0xF4, 0xFB))
    add_rect(sl, Inches(9.6), Inches(1.4), Inches(0.06), Inches(5.65), GOLD)

    add_text(sl, "The Key Insight",
             Inches(9.76), Inches(1.5), Inches(3.2), Inches(0.4),
             size=Pt(13), color=NAVY, bold=True, font="Calibri")

    insights = [
        "Under a gas constraint, the LP ranks products by $/1,000 cu.ft",
        "Top 4 protected products (always run at full capacity):",
        "   ① Ammonium Phosphate  $14.00",
        "   ② Urea  $10.36",
        "   ③ Chlorine  $8.46",
        "   ④ Ammonium Nitrate  $6.67",
        "Products cut first (from lowest):",
        "   ⑧ Hydrofluoric Acid  $5.59",
        "   ⑦ Ammonia  $5.83",
        "   ⑥ Vinyl Chloride Monomer  $5.94",
        "   ⑤ Caustic Soda  $6.25",
    ]
    y_pos = Inches(2.0)
    for item in insights:
        color = GREEN if '①' in item or '②' in item or '③' in item or '④' in item else (
                RED if '⑧' in item or '⑦' in item or '⑥' in item else (
                ORANGE if '⑤' in item else DARK))
        bold = 'protected' in item.lower() or 'cut first' in item.lower() or 'key' in item.lower()
        add_text(sl, item, Inches(9.76), y_pos, Inches(3.1), Inches(0.38),
                 size=Pt(9.5 if item.startswith(' ') else 10),
                 color=color, bold=bold, font="Calibri" if bold else "Calibri Light")
        y_pos += Inches(0.42)

    # Bottom note
    add_rect(sl, Inches(0.4), Inches(5.7), Inches(9.0), Inches(0.55),
             RGBColor(0xFF, 0xF3, 0xCC))
    add_text(sl, "Note: This is not a heuristic — it is the exact LP solution. "
             "The Simplex algorithm discovers this ranking automatically.",
             Inches(0.55), Inches(5.72), Inches(8.7), Inches(0.48),
             size=Pt(10), color=RGBColor(0x5A, 0x3A, 0x00), font="Calibri Light")

    return sl


def slide_production_results(prs, chart_path):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Production Plan Results",
                "Optimal product mix under each curtailment scenario (LP-derived)")
    slide_number(sl, 8)

    # Table
    headers = ["Product", "Baseline", "20% Cut", "30% Cut", "40% Cut", "50% Cut"]
    col_w = [Inches(2.4), Inches(1.6), Inches(1.6), Inches(1.6), Inches(1.6), Inches(1.6)]
    rows = [
        ["Ammonium Phosphate", "510", "510", "510", "510", "510"],
        ["Urea",               "150", "150", "150", "150", "150"],
        ["Chlorine",          "1,350","1,350","1,350","1,350","1,350"],
        ["Ammonium Nitrate",   "630", "630", "630", "630", "630"],
        ["Caustic Soda",      "1,280","1,280","1,280","1,280", "720"],
        ["Vinyl Chloride Mon.","980", "980", "651", "116",   "0"],
        ["Ammonia",           "1,050","275",  "0",   "0",    "0"],
        ["Hydrofluoric Acid",  "461",  "0",   "0",   "0",    "0"],
        ["Daily Profit ($)","$611,062","$513,023","$462,495","$411,623","$358,650"],
        ["vs. Baseline (%)",   "—",  "−16.0%","−24.3%","−32.6%","−41.3%"],
    ]
    add_table_manual(sl, headers, rows, Inches(0.45), Inches(1.45), col_w,
                     row_height=Inches(0.39))

    # Right: chart
    sl.shapes.add_picture(chart_path, Inches(10.7), Inches(1.42), Inches(2.4), Inches(5.6))

    # Bottom callout
    add_rect(sl, Inches(0.45), Inches(5.68), Inches(10.1), Inches(0.62),
             RGBColor(0xE8, 0xF5, 0xF0))
    add_text(sl,
             "✓  The LP protects the 4 highest-margin products at full capacity even under 50% curtailment.  "
             "Cut sequence: HF → Ammonia → VCM → Caustic Soda.",
             Inches(0.6), Inches(5.72), Inches(9.8), Inches(0.52),
             size=Pt(10.5), color=GREEN, bold=True, font="Calibri")

    return sl


def slide_distribution(prs, chart_path):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Distribution Plan Results",
                "Optimal gas allocation from MaxEnergy nodes (I, J, K) to BarCo plants (A, B, C)")
    slide_number(sl, 9)

    sl.shapes.add_picture(chart_path, Inches(0.4), Inches(1.38), Inches(9.6), Inches(3.8))

    # Key findings
    add_rect(sl, Inches(0.4), Inches(5.3), Inches(9.6), Inches(1.9),
             RGBColor(0xEB, 0xF1, 0xFA))
    add_rect(sl, Inches(0.4), Inches(5.3), Inches(0.06), Inches(1.9), TEAL)
    findings = [
        "Five active lanes: I→B, J→A, J→B, K→B, K→C  |  Baseline total cost: $113,952/day",
        "Structure is INVARIANT under curtailment — only volumes shrink proportionally",
        "Average unit cost remains constant at $1.33 per 1,000 cu.ft regardless of curtailment level",
        "Unused lanes I→C ($1.60) and K→A ($1.80) — candidates for renegotiation",
    ]
    y_pos = Inches(5.4)
    for f in findings:
        add_text(sl, f"▸  {f}", Inches(0.58), y_pos, Inches(9.2), Inches(0.42),
                 size=Pt(10.5), color=DARK, font="Calibri Light")
        y_pos += Inches(0.43)

    # Right: cost table
    add_rect(sl, Inches(10.2), Inches(1.38), Inches(2.9), Inches(3.8),
             RGBColor(0xF0, 0xF4, 0xFB))
    add_text(sl, "Distribution Costs\n($ per 1,000 cu.ft)",
             Inches(10.3), Inches(1.48), Inches(2.7), Inches(0.55),
             size=Pt(11), color=NAVY, bold=True, align=PP_ALIGN.CENTER, font="Calibri")

    cost_headers = ["", "Plant A", "Plant B", "Plant C"]
    cost_rows = [
        ["Source I", "$1.30", "$1.20 ✓", "$1.60"],
        ["Source J", "$1.20 ✓", "$1.40", "$1.50"],
        ["Source K", "$1.80", "$1.60", "$1.40 ✓"],
    ]
    cost_cw = [Inches(0.78), Inches(0.68), Inches(0.68), Inches(0.68)]
    add_table_manual(sl, cost_headers, cost_rows, Inches(10.24), Inches(2.1),
                     cost_cw, row_height=Inches(0.37))
    add_text(sl, "✓ = cheapest lane per source",
             Inches(10.24), Inches(3.28), Inches(2.7), Inches(0.3),
             size=Pt(8), color=GREY, font="Calibri Light")

    return sl


def slide_sensitivity(prs, curve_path):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Sensitivity Analysis",
                "Shadow price of gas & profit robustness under coefficient changes")
    slide_number(sl, 10)

    # Left: shadow price table
    add_text(sl, "Shadow Price of Gas Constraint",
             Inches(0.55), Inches(1.45), Inches(5.8), Inches(0.4),
             size=Pt(13), color=NAVY, bold=True, font="Calibri")
    add_rect(sl, Inches(0.55), Inches(1.82), Inches(5.8), Inches(0.04), TEAL)

    headers = ["Scenario", "Gas Budget", "Marginal Product", "Shadow Price"]
    shadow_rows = [
        ["Baseline",  "85,680",  "Hydrofluoric Acid",     "$5.59/10³"],
        ["20% Cut",   "68,544",  "Ammonia",               "$5.83/10³"],
        ["30% Cut",   "59,976",  "Vinyl Chloride Mon.",   "$5.94/10³"],
        ["40% Cut",   "51,408",  "Vinyl Chloride Mon.",   "$5.94/10³"],
        ["50% Cut",   "42,840",  "Caustic Soda",          "$6.25/10³"],
    ]
    col_w = [Inches(1.3), Inches(1.2), Inches(1.95), Inches(1.25)]
    add_table_manual(sl, headers, shadow_rows, Inches(0.55), Inches(1.9), col_w,
                     row_height=Inches(0.41))

    # Managerial implication
    add_rect(sl, Inches(0.55), Inches(4.05), Inches(5.8), Inches(1.25),
             RGBColor(0xFD, 0xF0, 0xC2))
    add_rect(sl, Inches(0.55), Inches(4.05), Inches(0.06), Inches(1.25), GOLD)
    add_text(sl, "Managerial Implication",
             Inches(0.72), Inches(4.12), Inches(5.5), Inches(0.35),
             size=Pt(11.5), color=NAVY, bold=True, font="Calibri")
    add_text(sl, "Every 1,000 cu.ft of spot gas BarCo secures above the curtailed "
             "budget is worth $5.59–$6.25 to daily profit. "
             "This defines the maximum price BarCo should pay for emergency gas procurement.",
             Inches(0.72), Inches(4.5), Inches(5.5), Inches(0.72),
             size=Pt(10), color=DARK, font="Calibri Light")

    # Right: profit curve chart
    sl.shapes.add_picture(curve_path, Inches(6.6), Inches(1.38), Inches(6.5), Inches(4.0))

    # Coefficient sensitivity summary
    add_rect(sl, Inches(0.55), Inches(5.45), W - Inches(1.1), Inches(1.55),
             RGBColor(0xE8, 0xF0, 0xFA))
    add_rect(sl, Inches(0.55), Inches(5.45), Inches(0.06), Inches(1.55), BLUE)
    add_text(sl, "Profit Coefficient Stability",
             Inches(0.72), Inches(5.53), Inches(11.0), Inches(0.38),
             size=Pt(12), color=NAVY, bold=True, font="Calibri")
    add_bullet_box(sl, [
        "Protected 4 products: stable within ±20% margin movement — no change to optimal mix",
        "Marginal products (Ammonia, VCM, Caustic Soda, HF): a −20% price drop removes them from the optimal mix",
        "Rule: small change (<10%) → no action; medium (10–20%) → re-run LP monthly; large (>20%) → re-optimise immediately",
    ],
    Inches(0.72), Inches(5.93), W - Inches(1.3), Inches(0.95),
    size=Pt(10), color=DARK, bullet="•")

    return sl


def slide_co2(prs, co2_path):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Environmental Scenario: Minimising CO₂ Emissions",
                "What if BarCo pursues sustainability alongside profit?")
    slide_number(sl, 11)

    sl.shapes.add_picture(co2_path, Inches(0.4), Inches(1.38), Inches(7.8), Inches(4.0))

    # Right: trade-off table and insights
    add_text(sl, "Profit–Emissions Trade-off",
             Inches(8.4), Inches(1.45), Inches(4.7), Inches(0.38),
             size=Pt(13), color=NAVY, bold=True, font="Calibri")
    add_rect(sl, Inches(8.4), Inches(1.82), Inches(4.7), Inches(0.04), TEAL)

    headers = ["Profit Floor", "Daily Profit", "CO₂ (kg)", "CO₂ Cut"]
    trade_rows = [
        ["100%", "$611,062", "5.74 kg",  "4.4% ↓"],
        ["90%",  "$549,956", "4.98 kg", "17.0% ↓"],
        ["80%",  "$488,849", "4.26 kg", "29.0% ↓"],
        ["70%",  "$427,743", "3.57 kg", "40.5% ↓"],
    ]
    col_w = [Inches(1.1), Inches(1.35), Inches(1.05), Inches(1.1)]
    add_table_manual(sl, headers, trade_rows, Inches(8.4), Inches(1.9), col_w,
                     row_height=Inches(0.41))

    add_rect(sl, Inches(8.4), Inches(3.65), Inches(4.7), Inches(3.35),
             RGBColor(0xEB, 0xF5, 0xEE))
    add_rect(sl, Inches(8.4), Inches(3.65), Inches(0.06), Inches(3.35), GREEN)
    add_text(sl, "Key Findings",
             Inches(8.56), Inches(3.75), Inches(4.4), Inches(0.38),
             size=Pt(12), color=NAVY, bold=True, font="Calibri")

    insights = [
        "At profit-max baseline, BarCo's production CO₂ is already 4.4% below the stated 6.0 kg baseline — because ammonia runs at full capacity (lower emission use)",
        "A 10% profit sacrifice achieves 17% CO₂ reduction",
        "Break-even carbon price ≈ $80,000/tonne",
        "Australia's Safeguard Mechanism: ~$30/tonne today — profit-max remains optimal",
        "The CO₂-aware LP is ready to activate when carbon prices rise meaningfully",
    ]
    y_pos = Inches(4.2)
    for ins in insights:
        add_text(sl, f"▸  {ins}", Inches(8.56), y_pos, Inches(4.4), Inches(0.5),
                 size=Pt(9.5), color=DARK, font="Calibri Light")
        y_pos += Inches(0.5)

    # Bottom bar
    add_rect(sl, Inches(0.4), Inches(5.55), Inches(7.8), Inches(0.65),
             RGBColor(0xE0, 0xF5, 0xF2))
    add_text(sl, "Ammonia deserves protection: it is BarCo's lowest-emission gas use AND qualifies for "
             "2nd-Priority regulatory protection. Cutting ammonia for short-term profit is both financially "
             "and reputationally risky.",
             Inches(0.58), Inches(5.58), Inches(7.5), Inches(0.58),
             size=Pt(9.5), color=RGBColor(0x04, 0x50, 0x44), font="Calibri Light")

    return sl


def slide_recommendations(prs):
    sl = blank_slide(prs)
    slide_background(sl)
    dark_header(sl, "Strategic Recommendations",
                "Three evidence-based actions for BarCo management — each with quantified value")
    slide_number(sl, 12)

    recs = [
        (
            "R1", "Implement the LP-Derived Production Priority Rule",
            NAVY, GOLD,
            [
                "Written policy: always run AP, Urea, Chlorine, AN at full capacity",
                "Cutback sequence: HF → Ammonia → VCM → Caustic Soda",
                "Value: ~$16M/year over naive proportional cut at 50% scenario",
                "Shadow price defines max price for emergency spot gas: $5.59–$6.25/10³ cu.ft",
            ]
        ),
        (
            "R2", "Commit to 5-Lane Distribution Structure",
            TEAL, WHITE,
            [
                "Pre-negotiate volume-flexible contracts on lanes: I→B, J→A, J→B, K→B, K→C",
                "No need to re-tender each curtailment cycle — structure is invariant",
                "Renegotiate unused lanes I→C ($1.60) and K→A ($1.80) for cost optionality",
                "Average unit cost locked at $1.33/10³ cu.ft regardless of curtailment level",
            ]
        ),
        (
            "R3", "Run CO₂-Aware LP in Parallel",
            GREEN, WHITE,
            [
                "Report Pareto trade-off to the board each planning cycle",
                "Baseline production CO₂ already 4.4% below stated target — document this",
                "Activate CO₂-aware variant when AUS carbon prices clear ~$80/tonne",
                "Positions BarCo ahead of tightening carbon regulation",
            ]
        ),
    ]

    for i, (num, title, bg_col, fg_col, points) in enumerate(recs):
        by = Inches(1.45) + i * Inches(1.95)
        add_rect(sl, Inches(0.55), by, W - Inches(1.1), Inches(1.82), ROW_ODD)
        add_rect(sl, Inches(0.55), by, Inches(0.95), Inches(1.82), bg_col)

        add_text(sl, num, Inches(0.58), by + Inches(0.55),
                 Inches(0.88), Inches(0.72),
                 size=Pt(24), color=fg_col if bg_col != NAVY else GOLD,
                 bold=True, align=PP_ALIGN.CENTER, font="Calibri")

        add_text(sl, title, Inches(1.6), by + Inches(0.12),
                 Inches(11.0), Inches(0.42),
                 size=Pt(13.5), color=NAVY, bold=True, font="Calibri")

        pt_text = "  ·  ".join(points[:2]) + "\n" + "  ·  ".join(points[2:])
        add_text(sl, pt_text, Inches(1.6), by + Inches(0.58),
                 Inches(11.0), Inches(1.1),
                 size=Pt(10), color=DARK, font="Calibri Light")

    return sl


def slide_conclusion(prs):
    sl = blank_slide(prs)
    # Dark background
    add_rect(sl, 0, 0, W, H, NAVY)
    add_rect(sl, 0, 0, W, Inches(0.06), GOLD)
    add_rect(sl, 0, H - Inches(0.06), W, Inches(0.06), TEAL)
    slide_number(sl, 13)

    add_text(sl, "Conclusion",
             Inches(0.7), Inches(0.55), Inches(11.0), Inches(0.65),
             size=Pt(34), color=GOLD, bold=True,
             align=PP_ALIGN.LEFT, font="Calibri")

    add_rect(sl, Inches(0.7), Inches(1.22), Inches(7.0), Inches(0.045), GOLD)

    insights = [
        ("LP transforms curtailment decisions",
         "from high-pressure judgment calls into transparent, pre-committed policy"),
        ("Profit drops sub-proportionally to gas",
         "a 50% gas cut → only 41.3% profit drop  ≈  $16M/year in LP value"),
        ("Shadow price quantifies gas value",
         "$5.59–$6.25 per 1,000 cu.ft — the ceiling for emergency procurement"),
        ("Same model supports CO₂ planning",
         "ready to activate as carbon pricing enters Australian industrial policy"),
    ]

    for i, (bold_part, normal_part) in enumerate(insights):
        by = Inches(1.4) + i * Inches(0.9)
        add_rect(sl, Inches(0.7), by, Inches(0.06), Inches(0.7), TEAL)
        add_text(sl, bold_part, Inches(0.9), by + Inches(0.06),
                 Inches(11.0), Inches(0.38),
                 size=Pt(14), color=WHITE, bold=True, font="Calibri")
        add_text(sl, normal_part, Inches(0.9), by + Inches(0.44),
                 Inches(11.0), Inches(0.38),
                 size=Pt(12), color=RGBColor(0xA0, 0xB8, 0xD4),
                 font="Calibri Light")

    add_rect(sl, Inches(0.7), Inches(5.15), W - Inches(1.4), Inches(1.6),
             NAVY2)
    add_text(sl, "With these three recommendations implemented, BarCo will not only weather "
             "the 2026 brownouts — it will build a planning capability fit for the more "
             "constrained, more carbon-aware operating environment of the years ahead.",
             Inches(0.9), Inches(5.3), W - Inches(1.8), Inches(1.2),
             size=Pt(13.5), color=RGBColor(0xCC, 0xDD, 0xEE),
             align=PP_ALIGN.LEFT, font="Calibri Light")

    add_text(sl, "Thank you  ·  Questions welcome",
             Inches(0.7), H - Inches(0.7), Inches(8.0), Inches(0.45),
             size=Pt(15), color=GOLD, bold=True, font="Calibri")

    return sl


# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD PPTX
# ═══════════════════════════════════════════════════════════════════════════════

def build_pptx():
    print("Generating charts...")
    pg_chart   = make_profit_per_gas_chart()
    prod_chart = make_production_results_chart()
    curve      = make_profit_curve_chart()
    dist_chart = make_distribution_chart()
    co2_chart  = make_co2_chart()

    print("Building PPTX...")
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    slide_title(prs)
    slide_overview(prs)
    slide_business_problem(prs)
    slide_financial_risk(prs)
    slide_approach(prs)
    slide_assumptions(prs)
    slide_product_economics(prs, pg_chart)
    slide_production_results(prs, prod_chart)
    slide_distribution(prs, dist_chart)
    slide_sensitivity(prs, curve)
    slide_co2(prs, co2_chart)
    slide_recommendations(prs)
    slide_conclusion(prs)

    out_pptx = f"{OUT}/BarCo_Presentation.pptx"
    prs.save(out_pptx)
    print(f"  Saved: {out_pptx}")
    return out_pptx


# ═══════════════════════════════════════════════════════════════════════════════
#  SPEECH SCRIPT PDF
# ═══════════════════════════════════════════════════════════════════════════════

SPEECH_TEXT = """
BARCO CORPORATION SUPPLY CHAIN OPTIMISATION
Individual Video Presentation — Speech Script
49680 Value Chain Engineering Systems — Autumn 2026
Estimated duration: 5–6 minutes

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 1 — TITLE]  (~30 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Good morning, everyone. My name is Botao Huo, student ID 25104579, from Group Z. Today I'm presenting our group's analysis on how BarCo Corporation can best manage a natural gas curtailment [kɜːr-ˈteɪl-mənt] crisis using a technique called Linear Programming [ˈlɪn-i-ər ˈproʊ-ɡræm-ɪŋ].

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 2 — OVERVIEW]  (~15 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

I'll cover four areas: first, the business problem; second, our analysis approach; third, the key results; and finally, our strategic recommendations.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 3 — BUSINESS PROBLEM]  (~45 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BarCo Corporation operates eight chemical product lines across three manufacturing plants. The company's energy supplier, MaxEnergy, has warned of natural gas shortages caused by an extreme heat wave — electrical power plants are running at full capacity, and natural gas remains the dominant boiler fuel.

MaxEnergy plans rolling brownouts [ˈbraʊn-aʊts] — that means temporary, periodic reductions in gas supply — of between 20 and 40 percent. The Federal Commission has established priority rules: residential heating comes first; industrial raw material users come second; and industrial boiler fuel users come third. Most of BarCo's operations fall into the second and third categories, so BarCo must absorb a large portion of the cut itself.

The key question is: which products should BarCo reduce, and by how much, to protect as much profit as possible?

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 4 — FINANCIAL RISK]  (~30 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

At the current baseline, BarCo earns $611,062 per day. Without a smart plan, a 40 percent curtailment could cut profit by about 33 percent. However, with our LP approach, profit drops less than proportionally to gas — a key advantage worth approximately $16 million per year over doing nothing.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 5 — ANALYSIS APPROACH]  (~40 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

To answer this problem, we used Linear Programming — a mathematical optimisation [ˌɒp-tɪ-maɪ-ˈzeɪ-ʃən] technique designed exactly for this type of situation: maximise profit subject to constraints [kən-ˈstreɪnts].

We built two models. The first is a production LP that maximises daily profit under a gas budget. Each product has a known profit per ton, gas consumption per ton, and a daily capacity ceiling. The second is a distribution LP that minimises gas transport costs across the pipeline network connecting MaxEnergy's three supply nodes to BarCo's three plants. Both models were solved using the Simplex [ˈsɪm-pleks] algorithm [ˈæl-ɡə-rɪð-əm].

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 6 — ASSUMPTIONS]  (~20 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Key assumptions include: profits and gas consumption rates are constant per ton — we call this linearity; production volumes can be any real number — divisibility; all cost values are known with certainty; and curtailment applies equally to all three supply nodes. Sensitivity [ˌsen-sɪ-ˈtɪv-ɪ-ti] analysis later tests how robust our results are when these assumptions are relaxed.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 7 — PRODUCT ECONOMICS]  (~35 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This chart shows the key insight: when gas is scarce, what matters is not just profit per ton, but profit per unit of gas consumed. Ammonium [ə-ˈmoʊ-ni-əm] phosphate [ˈfɒs-feɪt] earns $14 per thousand cubic feet of gas — the highest in our portfolio. Urea [jʊə-ˈriː-ə] comes second at $10.36. At the bottom sits hydrofluoric [ˌhaɪ-drə-ˈflʊər-ɪk] acid at just $5.59. The LP will protect high-value products first, and cut low-value ones first. This is not a guess — it is the mathematically exact LP solution.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 8 — PRODUCTION RESULTS]  (~45 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This table shows the optimal production mix at each curtailment level. At baseline, we achieve $611,062 per day. Under a 20 percent cut, profit falls to $513,023 — a 16 percent drop despite a 20 percent gas reduction. Under 40 percent curtailment, profit falls to $411,623 — a 32.6 percent drop.

The cut sequence is: hydrofluoric acid first, then ammonia, then vinyl [ˈvaɪ-nəl] chloride [ˈklɔːr-aɪd] monomer [ˈmɒn-ə-mər], then caustic [ˈkɔːs-tɪk] soda. Our four protected products — ammonium phosphate, urea, chlorine, and ammonium nitrate [ˈnaɪ-treɪt] — remain at full capacity even under a 50 percent cut. Versus a simple proportional cut, this LP approach saves $44,000 per day — worth $16 million annualised [ˈæn-ju-ə-laɪzd].

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 9 — DISTRIBUTION RESULTS]  (~30 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

For gas distribution, the optimal plan uses five pipeline lanes: I to B, J to A, J to B, K to B, and K to C. The total transport cost at baseline is $113,952 per day. Remarkably, this structural pattern does not change under any level of curtailment — only the volumes shrink proportionally. This means BarCo can pre-negotiate the same five contracts and simply adjust volumes, eliminating the need to re-tender with each brownout.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 10 — SENSITIVITY ANALYSIS]  (~30 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Our sensitivity [ˌsen-sɪ-ˈtɪv-ɪ-ti] analysis shows that every extra 1,000 cubic feet of gas above the curtailed budget is worth between $5.59 and $6.25 in daily profit. This is called the shadow price. It gives management a clear maximum price to pay for emergency gas procurement or fuel-switching. On profit coefficients [ˌkoʊ-ɪ-ˈfɪʃ-ənts]: our four protected products can absorb price changes of plus or minus 20 percent without any change to the optimal mix. Marginal products are more sensitive — monitor them closely during volatile market periods.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 11 — ENVIRONMENTAL SCENARIO]  (~20 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

On sustainability: the same LP model, extended with one extra constraint, can minimise CO₂ emissions. If BarCo accepts a 10 percent reduction in profit, it achieves a 17 percent reduction in CO₂. The break-even carbon price is approximately $80,000 per tonne — well above Australia's current Safeguard Mechanism price of around $30. Today, profit maximisation remains optimal, but this model positions BarCo ahead of future carbon regulation.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 12 — RECOMMENDATIONS]  (~50 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Based on our analysis, we make three recommendations.

Recommendation One: Adopt the LP-derived production priority rule as a written policy. Always run ammonium phosphate, urea, chlorine, and ammonium nitrate at full capacity. The official cut sequence is: hydrofluoric acid first, then ammonia, then vinyl chloride monomer, then caustic soda. This rule alone saves approximately $16 million per year compared to a naive approach. The shadow price also tells management the exact price ceiling for emergency gas procurement.

Recommendation Two: Pre-commit to the five-lane distribution structure with volume-flexible contracts. There is no need to re-tender each cycle — the optimal routing is structurally stable. Also renegotiate the currently unused and expensive lanes — I to C and K to A — to retain optionality if curtailment patterns change.

Recommendation Three: Run the CO₂-aware LP model in parallel every planning cycle. Report the trade-off to the board. Monitor Australia's Safeguard Mechanism carbon price, and activate the emissions-aware production plan as soon as carbon prices become economically significant.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[SLIDE 13 — CONCLUSION]  (~20 seconds)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

In summary: Linear Programming transforms BarCo's curtailment challenge from a reactive, high-pressure guessing game into a proactive, quantified policy. With these three recommendations, BarCo is prepared not just for the 2026 brownouts, but for a more constrained and carbon-aware future ahead.

Thank you very much for listening. I'm happy to take any questions.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHONETIC GUIDE TO DIFFICULT TERMS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

curtailment         [kɜːr-ˈteɪl-mənt]
linear programming  [ˈlɪn-i-ər ˈproʊ-ɡræm-ɪŋ]
optimisation        [ˌɒp-tɪ-maɪ-ˈzeɪ-ʃən]
algorithm           [ˈæl-ɡə-rɪð-əm]
simplex             [ˈsɪm-pleks]
constraint          [kən-ˈstreɪnt]
coefficient         [ˌkoʊ-ɪ-ˈfɪʃ-ənt]
sensitivity         [ˌsen-sɪ-ˈtɪv-ɪ-ti]
annualised          [ˈæn-ju-ə-laɪzd]
ammonium            [ə-ˈmoʊ-ni-əm]
phosphate           [ˈfɒs-feɪt]
urea                [jʊə-ˈriː-ə]
nitrate             [ˈnaɪ-treɪt]
hydrofluoric acid   [ˌhaɪ-drə-ˈflʊər-ɪk ˈæs-ɪd]
vinyl               [ˈvaɪ-nəl]
chloride            [ˈklɔːr-aɪd]
monomer             [ˈmɒn-ə-mər]
caustic             [ˈkɔːs-tɪk]
brownout            [ˈbraʊn-aʊt]
formulation         [ˌfɔːr-mjʊ-ˈleɪ-ʃən]
scenario            [sɪ-ˈnær-i-oʊ]
prioritisation      [praɪ-ˌɒr-ɪ-tɪ-ˈzeɪ-ʃən]
"""


def build_speech_pdf():
    print("Building speech script PDF...")
    out_pdf = f"{OUT}/BarCo_Speech_Script.pdf"

    doc = SimpleDocTemplate(
        out_pdf, pagesize=A4,
        leftMargin=2.2*cm, rightMargin=2.2*cm,
        topMargin=2.0*cm, bottomMargin=2.0*cm
    )

    NAVY_RL  = HexColor('#081429')
    GOLD_RL  = HexColor('#F0AE1C')
    TEAL_RL  = HexColor('#0D9287')
    GREEN_RL = HexColor('#18875C')
    RED_RL   = HexColor('#C62B2B')
    GREY_RL  = HexColor('#5E6E86')
    DARK_RL  = HexColor('#121C2C')
    OFFWHITE_RL = HexColor('#F4F7FD')
    AMBER_RL = HexColor('#7A4500')

    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        'DocTitle', parent=styles['Normal'],
        fontSize=20, fontName='Helvetica-Bold',
        textColor=NAVY_RL, alignment=TA_CENTER,
        spaceAfter=4
    )
    style_subtitle = ParagraphStyle(
        'DocSubtitle', parent=styles['Normal'],
        fontSize=11, fontName='Helvetica',
        textColor=GREY_RL, alignment=TA_CENTER,
        spaceAfter=12
    )
    style_slide_header = ParagraphStyle(
        'SlideHeader', parent=styles['Normal'],
        fontSize=11, fontName='Helvetica-Bold',
        textColor=NAVY_RL, spaceBefore=10, spaceAfter=4,
        leftIndent=0, backColor=HexColor('#EBF1FA'),
        borderPad=6, borderRadius=2,
        leading=16
    )
    style_body = ParagraphStyle(
        'Body', parent=styles['Normal'],
        fontSize=11, fontName='Helvetica',
        textColor=DARK_RL, leading=17,
        spaceBefore=2, spaceAfter=4,
        alignment=TA_JUSTIFY
    )
    style_phonetic = ParagraphStyle(
        'Phonetic', parent=styles['Normal'],
        fontSize=10.5, fontName='Helvetica',
        textColor=DARK_RL, leading=16,
        spaceBefore=1, spaceAfter=1,
        leftIndent=12
    )
    style_guide_header = ParagraphStyle(
        'GuideHeader', parent=styles['Normal'],
        fontSize=12, fontName='Helvetica-Bold',
        textColor=NAVY_RL, spaceBefore=12, spaceAfter=6
    )
    style_guide_row = ParagraphStyle(
        'GuideRow', parent=styles['Normal'],
        fontSize=10.5, fontName='Courier',
        textColor=DARK_RL, leading=15, spaceAfter=1
    )
    style_note = ParagraphStyle(
        'Note', parent=styles['Normal'],
        fontSize=9.5, fontName='Helvetica-Oblique',
        textColor=GREY_RL, spaceBefore=2, spaceAfter=2
    )

    def phon(word, ipa, color=TEAL_RL):
        return (f'<font color="#{color.hexval()[2:]}">'
                f'<b>{word}</b></font> '
                f'<font color="#5E6E86" size="9">[{ipa}]</font>')

    story = []

    # Title block
    story.append(Paragraph("BarCo Corporation Supply Chain Optimisation", style_title))
    story.append(Paragraph("Individual Video Presentation — Speech Script", style_subtitle))
    story.append(Paragraph(
        "49680 Value Chain Engineering Systems · Autumn 2026 · UTS FEIT · Group Z",
        style_subtitle))
    story.append(HRFlowable(width="100%", thickness=2, color=GOLD_RL, spaceAfter=4))

    story.append(Paragraph(
        '<font color="#18875C"><b>Estimated speaking time: 5–6 minutes</b></font>  '
        '<font color="#5E6E86" size="9.5">(~700 words at 120–140 wpm)</font>',
        ParagraphStyle('Info', parent=styles['Normal'], fontSize=10,
                       alignment=TA_CENTER, spaceAfter=10)))
    story.append(Spacer(1, 4))

    # Parse and render speech sections
    sections = SPEECH_TEXT.strip().split('━' * 10)
    in_guide = False

    for section in sections:
        section = section.strip()
        if not section:
            continue

        if 'PHONETIC GUIDE' in section:
            story.append(HRFlowable(width="100%", thickness=1, color=GREY_RL,
                                    spaceBefore=8, spaceAfter=8))
            story.append(Paragraph("Phonetic Guide to Difficult Terms",
                                    style_guide_header))
            story.append(Paragraph(
                "Terms in bold appear in the script with their phonetic notation. "
                "Use this page as a quick reference before recording.",
                style_note))
            story.append(Spacer(1, 4))

            guide_pairs = [
                ("curtailment",        "kɜːr-ˈteɪl-mənt"),
                ("linear programming", "ˈlɪn-i-ər  ˈproʊ-ɡræm-ɪŋ"),
                ("optimisation",       "ˌɒp-tɪ-maɪ-ˈzeɪ-ʃən"),
                ("algorithm",          "ˈæl-ɡə-rɪð-əm"),
                ("simplex",            "ˈsɪm-pleks"),
                ("constraint",         "kən-ˈstreɪnt"),
                ("coefficient",        "ˌkoʊ-ɪ-ˈfɪʃ-ənt"),
                ("sensitivity",        "ˌsen-sɪ-ˈtɪv-ɪ-ti"),
                ("annualised",         "ˈæn-ju-ə-laɪzd"),
                ("ammonium",           "ə-ˈmoʊ-ni-əm"),
                ("phosphate",          "ˈfɒs-feɪt"),
                ("urea",               "jʊə-ˈriː-ə"),
                ("nitrate",            "ˈnaɪ-treɪt"),
                ("hydrofluoric acid",  "ˌhaɪ-drə-ˈflʊər-ɪk  ˈæs-ɪd"),
                ("vinyl",              "ˈvaɪ-nəl"),
                ("chloride",           "ˈklɔːr-aɪd"),
                ("monomer",            "ˈmɒn-ə-mər"),
                ("caustic",            "ˈkɔːs-tɪk"),
                ("brownout",           "ˈbraʊn-aʊt"),
                ("formulation",        "ˌfɔːr-mjʊ-ˈleɪ-ʃən"),
                ("scenario",           "sɪ-ˈnær-i-oʊ"),
                ("prioritisation",     "praɪ-ˌɒr-ɪ-tɪ-ˈzeɪ-ʃən"),
            ]
            tbl_data = [["Term", "Phonetic Notation"]]
            for term, ipa in guide_pairs:
                tbl_data.append([term, f"[{ipa}]"])

            tbl = Table(tbl_data, colWidths=[6.5*cm, 10*cm])
            tbl.setStyle(TableStyle([
                ('BACKGROUND',  (0, 0), (-1, 0), NAVY_RL),
                ('TEXTCOLOR',   (0, 0), (-1, 0), white),
                ('FONTNAME',    (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE',    (0, 0), (-1, 0), 10),
                ('FONTNAME',    (0, 1), (0, -1), 'Helvetica-Bold'),
                ('FONTNAME',    (1, 1), (1, -1), 'Courier'),
                ('FONTSIZE',    (0, 1), (-1, -1), 9.5),
                ('TEXTCOLOR',   (0, 1), (0, -1), NAVY_RL),
                ('TEXTCOLOR',   (1, 1), (1, -1), HexColor('#0D9287')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1),
                 [HexColor('#EBF1FA'), HexColor('#FFFFFF')]),
                ('GRID',        (0, 0), (-1, -1), 0.5, HexColor('#C0C8D8')),
                ('TOPPADDING',  (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ]))
            story.append(tbl)
            continue

        lines = section.split('\n')
        is_header = False

        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 3))
                continue

            if line.startswith('[SLIDE') and ']' in line:
                is_header = True
                slide_num = line.split(']')[0].replace('[', '')
                timing = line.split('(')[-1].replace(')', '').strip() if '(' in line else ''
                header_text = (f'<b>{slide_num}</b>'
                               + (f'  <font size="9" color="#5E6E86">({timing})</font>'
                                  if timing else ''))
                story.append(Paragraph(header_text, style_slide_header))
                continue

            # Inline phonetic annotations
            line = line.replace('curtailment [kɜːr-ˈteɪl-mənt]',
                phon('curtailment', 'kɜːr-ˈteɪl-mənt'))
            line = line.replace('Linear Programming [ˈlɪn-i-ər ˈproʊ-ɡræm-ɪŋ]',
                phon('Linear Programming', 'ˈlɪn-i-ər  ˈproʊ-ɡræm-ɪŋ'))
            line = line.replace('brownouts [ˈbraʊn-aʊts]',
                phon('brownouts', 'ˈbraʊn-aʊts'))
            line = line.replace('optimisation [ˌɒp-tɪ-maɪ-ˈzeɪ-ʃən]',
                phon('optimisation', 'ˌɒp-tɪ-maɪ-ˈzeɪ-ʃən'))
            line = line.replace('constraints [kən-ˈstreɪnts]',
                phon('constraints', 'kən-ˈstreɪnts'))
            line = line.replace('Simplex [ˈsɪm-pleks]',
                phon('Simplex', 'ˈsɪm-pleks'))
            line = line.replace('algorithm [ˈæl-ɡə-rɪð-əm]',
                phon('algorithm', 'ˈæl-ɡə-rɪð-əm'))
            line = line.replace('Sensitivity [ˌsen-sɪ-ˈtɪv-ɪ-ti]',
                phon('Sensitivity', 'ˌsen-sɪ-ˈtɪv-ɪ-ti'))
            line = line.replace('sensitivity [ˌsen-sɪ-ˈtɪv-ɪ-ti]',
                phon('sensitivity', 'ˌsen-sɪ-ˈtɪv-ɪ-ti'))
            line = line.replace('Ammonium [ə-ˈmoʊ-ni-əm]',
                phon('Ammonium', 'ə-ˈmoʊ-ni-əm'))
            line = line.replace('ammonium [ə-ˈmoʊ-ni-əm]',
                phon('ammonium', 'ə-ˈmoʊ-ni-əm'))
            line = line.replace('phosphate [ˈfɒs-feɪt]',
                phon('phosphate', 'ˈfɒs-feɪt'))
            line = line.replace('Urea [jʊə-ˈriː-ə]',
                phon('Urea', 'jʊə-ˈriː-ə'))
            line = line.replace('urea [jʊə-ˈriː-ə]',
                phon('urea', 'jʊə-ˈriː-ə'))
            line = line.replace('hydrofluoric [ˌhaɪ-drə-ˈflʊər-ɪk]',
                phon('hydrofluoric', 'ˌhaɪ-drə-ˈflʊər-ɪk'))
            line = line.replace('vinyl [ˈvaɪ-nəl]',
                phon('vinyl', 'ˈvaɪ-nəl'))
            line = line.replace('chloride [ˈklɔːr-aɪd]',
                phon('chloride', 'ˈklɔːr-aɪd'))
            line = line.replace('monomer [ˈmɒn-ə-mər]',
                phon('monomer', 'ˈmɒn-ə-mər'))
            line = line.replace('caustic [ˈkɔːs-tɪk]',
                phon('caustic', 'ˈkɔːs-tɪk'))
            line = line.replace('nitrate [ˈnaɪ-treɪt]',
                phon('nitrate', 'ˈnaɪ-treɪt'))
            line = line.replace('annualised [ˈæn-ju-ə-laɪzd]',
                phon('annualised', 'ˈæn-ju-ə-laɪzd'))
            line = line.replace('coefficients [ˌkoʊ-ɪ-ˈfɪʃ-ənts]',
                phon('coefficients', 'ˌkoʊ-ɪ-ˈfɪʃ-ənts'))

            story.append(Paragraph(line, style_body))

    doc.build(story)
    print(f"  Saved: {out_pdf}")
    return out_pdf


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    pptx_path = build_pptx()
    speech_pdf = build_speech_pdf()
    print("\nAll files generated:")
    print(f"  PPT:    {pptx_path}")
    print(f"  Speech: {speech_pdf}")
