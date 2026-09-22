"""
pdf_report.py
=============
Generates Kenya_Debt_Report.pdf — a professional 3-page executive briefing note.

Pages:
  1. Cover + Classification + Key Findings
  2. Scorecard Table + Stress Test Summary
  3. Watchpoints + Recommendations

Run:
    python pdf_report.py

Requires:
    pip install reportlab
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.platypus.flowables import HRFlowable
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics import renderPDF
from reportlab.graphics.widgets.markers import makeMarker
import datetime

OUTPUT = "Kenya_Debt_Report.pdf"
PAGE_W, PAGE_H = A4

# ── Colour constants (RGB 0-1)
NAVY   = colors.HexColor("#1F4E79")
BLUE   = colors.HexColor("#2E75B6")
RED    = colors.HexColor("#C00000")
AMBER  = colors.HexColor("#FFC000")
GREEN  = colors.HexColor("#70AD47")
LGREY  = colors.HexColor("#F2F2F2")
MGREY  = colors.HexColor("#BFBFBF")
DGREY  = colors.HexColor("#404040")
WHITE  = colors.white
BLACK  = colors.black
YELLOW = colors.HexColor("#FFD966")

# ── Data
FY       = ["FY22/23","FY23/24","FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
GDP      = [12_100, 13_200, 14_400, 15_800, 17_300, 18_900, 20_600, 22_400]
DEBT_GDP = [52.3,   57.1,   63.7,   62.5,   61.0,   59.8,   58.5,   57.2]
DS_REV   = [47.0,   50.2,   71.2,   66.5,   62.0,   58.8,   55.2,   52.0]
DS_GDP   = [7.7,    8.4,    9.0,    8.7,    8.5,    8.3,    8.0,    7.7]

PROJ_FY  = ["FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
SCENARIOS = {
    "Baseline":          [63.7, 62.5, 61.0, 59.8, 58.5, 57.2],
    "Lower Growth":      [63.7, 66.4, 66.9, 67.3, 67.5, 67.4],
    "FX Depreciation":   [63.7, 69.2, 67.8, 65.5, 63.0, 61.0],
    "Interest Shock":    [63.7, 64.0, 63.2, 62.1, 61.0, 60.0],
    "Revenue Shortfall": [63.7, 66.8, 68.4, 70.1, 71.5, 72.8],
    "Optimistic Reform": [63.7, 61.0, 58.8, 56.5, 54.2, 52.0],
}

# ── Styles
def make_styles():
    s = getSampleStyleSheet()
    styles = {}

    styles["cover_title"] = ParagraphStyle(
        "cover_title", fontSize=26, textColor=WHITE,
        fontName="Helvetica-Bold", alignment=TA_CENTER,
        spaceAfter=6, leading=32)

    styles["cover_sub"] = ParagraphStyle(
        "cover_sub", fontSize=13, textColor=WHITE,
        fontName="Helvetica", alignment=TA_CENTER,
        spaceAfter=4, leading=18)

    styles["cover_meta"] = ParagraphStyle(
        "cover_meta", fontSize=10, textColor=LGREY,
        fontName="Helvetica", alignment=TA_CENTER,
        spaceAfter=3, leading=14)

    styles["section_head"] = ParagraphStyle(
        "section_head", fontSize=12, textColor=WHITE,
        fontName="Helvetica-Bold", alignment=TA_LEFT,
        spaceAfter=4, leading=16,
        backColor=NAVY, leftIndent=-6, rightIndent=-6,
        borderPadding=(4,6,4,6))

    styles["body"] = ParagraphStyle(
        "body", fontSize=9.5, textColor=BLACK,
        fontName="Helvetica", alignment=TA_JUSTIFY,
        spaceAfter=4, leading=14)

    styles["body_bold"] = ParagraphStyle(
        "body_bold", fontSize=9.5, textColor=BLACK,
        fontName="Helvetica-Bold", alignment=TA_LEFT,
        spaceAfter=3, leading=14)

    styles["bullet"] = ParagraphStyle(
        "bullet", fontSize=9.5, textColor=DGREY,
        fontName="Helvetica", alignment=TA_LEFT,
        leftIndent=14, spaceAfter=3, leading=14,
        bulletIndent=4)

    styles["caption"] = ParagraphStyle(
        "caption", fontSize=8, textColor=DGREY,
        fontName="Helvetica-Oblique", alignment=TA_CENTER,
        spaceAfter=6, leading=11)

    styles["footer"] = ParagraphStyle(
        "footer", fontSize=7.5, textColor=MGREY,
        fontName="Helvetica", alignment=TA_CENTER, leading=10)

    styles["kpi_label"] = ParagraphStyle(
        "kpi_label", fontSize=8, textColor=WHITE,
        fontName="Helvetica-Bold", alignment=TA_CENTER, leading=10)

    styles["kpi_value"] = ParagraphStyle(
        "kpi_value", fontSize=16, textColor=BLACK,
        fontName="Helvetica-Bold", alignment=TA_CENTER, leading=20)

    styles["kpi_sub"] = ParagraphStyle(
        "kpi_sub", fontSize=7.5, textColor=DGREY,
        fontName="Helvetica-Oblique", alignment=TA_CENTER, leading=10)

    styles["toc"] = ParagraphStyle(
        "toc", fontSize=10, textColor=LGREY,
        fontName="Helvetica", alignment=TA_LEFT,
        spaceAfter=4, leading=16)

    styles["warn_box"] = ParagraphStyle(
        "warn_box", fontSize=9.5, textColor=colors.HexColor("#7F0000"),
        fontName="Helvetica-Bold", alignment=TA_CENTER,
        leading=13, backColor=YELLOW,
        borderPadding=(5,8,5,8))

    styles["rec_head"] = ParagraphStyle(
        "rec_head", fontSize=10, textColor=NAVY,
        fontName="Helvetica-Bold", alignment=TA_LEFT,
        spaceAfter=2, leading=14)

    return styles


# ── Page template callbacks
def _header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4

    # Top rule
    canvas.setStrokeColor(NAVY)
    canvas.setLineWidth(2)
    canvas.line(1.5*cm, h - 1.1*cm, w - 1.5*cm, h - 1.1*cm)

    # Header text
    canvas.setFont("Helvetica-Bold", 8)
    canvas.setFillColor(NAVY)
    canvas.drawString(1.5*cm, h - 0.85*cm,
                      "KENYA PUBLIC DEBT SUSTAINABILITY ANALYSIS")
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(DGREY)
    canvas.drawRightString(w - 1.5*cm, h - 0.85*cm,
                           "OFFICIAL — RESTRICTED  |  PDMO  |  Sep 2026")

    # Bottom rule
    canvas.setStrokeColor(MGREY)
    canvas.setLineWidth(0.5)
    canvas.line(1.5*cm, 1.3*cm, w - 1.5*cm, 1.3*cm)

    # Footer
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MGREY)
    canvas.drawCentredString(w/2, 0.9*cm,
        f"Page {doc.page}  |  Projections are model outputs. "
        "For official data refer to PDMO published bulletins.")
    canvas.restoreState()


def _cover_background(canvas, doc):
    """Navy background only on page 1."""
    if doc.page == 1:
        canvas.saveState()
        canvas.setFillColor(NAVY)
        canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
        # Accent bar
        canvas.setFillColor(BLUE)
        canvas.rect(0, PAGE_H * 0.38, PAGE_W, PAGE_H * 0.62, fill=1, stroke=0)
        # Gold bottom strip
        canvas.setFillColor(AMBER)
        canvas.rect(0, 0, PAGE_W, 0.8*cm, fill=1, stroke=0)
        canvas.restoreState()
    else:
        _header_footer(canvas, doc)


# ── Chart helper — DS/Revenue bar chart
def make_ds_rev_chart():
    drawing = Drawing(400, 160)

    bc = VerticalBarChart()
    bc.x = 40; bc.y = 20
    bc.width = 340; bc.height = 120
    bc.data = [DS_REV, [30.0]*8]
    bc.categoryAxis.categoryNames = FY
    bc.categoryAxis.labels.fontSize = 7
    bc.categoryAxis.labels.angle = 30
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = 85
    bc.valueAxis.valueStep = 10
    bc.valueAxis.labels.fontSize = 7
    bc.bars[0].fillColor = AMBER
    bc.bars[0].strokeColor = colors.HexColor("#CC9900")
    bc.bars[1].fillColor = RED
    bc.bars[1].strokeColor = RED
    bc.barWidth = 0.35 * cm
    bc.groupSpacing = 0.2 * cm
    drawing.add(bc)

    # Legend
    drawing.add(Rect(42, 148, 10, 8, fillColor=AMBER, strokeColor=None))
    drawing.add(String(56, 149, "DS/Revenue %", fontSize=7))
    drawing.add(Rect(130, 148, 10, 8, fillColor=RED, strokeColor=None))
    drawing.add(String(144, 149, "30% Threshold", fontSize=7))

    return drawing


# ── Chart helper — Debt/GDP line
def make_debt_gdp_chart():
    drawing = Drawing(400, 160)

    lc = HorizontalLineChart()
    lc.x = 40; lc.y = 20
    lc.width = 340; lc.height = 120
    lc.data = [DEBT_GDP, [55.0]*8]
    lc.categoryAxis.categoryNames = FY
    lc.categoryAxis.labels.fontSize = 7
    lc.categoryAxis.labels.angle = 30
    lc.valueAxis.valueMin = 40
    lc.valueAxis.valueMax = 75
    lc.valueAxis.valueStep = 5
    lc.valueAxis.labels.fontSize = 7
    lc.lines[0].strokeColor = NAVY
    lc.lines[0].strokeWidth = 2
    lc.lines[1].strokeColor = RED
    lc.lines[1].strokeWidth = 1.5
    lc.lines[1].strokeDashArray = [4, 2]
    drawing.add(lc)

    drawing.add(Rect(42, 148, 20, 5, fillColor=NAVY, strokeColor=None))
    drawing.add(String(66, 148, "Debt/GDP Baseline", fontSize=7))
    drawing.add(Rect(170, 148, 20, 5, fillColor=RED, strokeColor=None))
    drawing.add(String(194, 148, "55% Anchor", fontSize=7))

    return drawing


# ── KPI tile table
def make_kpi_table(styles):
    kpis = [
        ("Total Public Debt",       "KES 6,215 Bn", "55.0% of GDP",           AMBER),
        ("PV Total Debt / GDP",     "63.7%",        "Exceeds 55% anchor",     RED),
        ("Debt Service / Revenue",  "71.2%",        "2× above 30% threshold", RED),
        ("Primary Balance / GDP",   "–2.3%",        "Target ≥ 0% FY27/28",   AMBER),
    ]

    cell_data = []
    row1_labels = [Paragraph(k, styles["kpi_label"]) for k,v,s,c in kpis]
    row1_values = [Paragraph(v, styles["kpi_value"]) for k,v,s,c in kpis]
    row1_subs   = [Paragraph(s, styles["kpi_sub"])   for k,v,s,c in kpis]

    tbl = Table(
        [row1_labels, row1_values, row1_subs],
        colWidths=[4.2*cm]*4,
        rowHeights=[0.7*cm, 1.1*cm, 0.6*cm]
    )

    style_cmds = [
        ("BACKGROUND", (0,0), (-1,0), NAVY),
        ("BACKGROUND", (1,1), (1,1), RED),
        ("BACKGROUND", (2,1), (2,1), RED),
        ("BACKGROUND", (0,1), (0,1), AMBER),
        ("BACKGROUND", (3,1), (3,1), AMBER),
        ("BACKGROUND", (0,2), (-1,2), LGREY),
        ("BOX",       (0,0), (-1,-1), 0.5, MGREY),
        ("INNERGRID", (0,0), (-1,-1), 0.3, MGREY),
        ("VALIGN",    (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",     (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING",(0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]
    tbl.setStyle(TableStyle(style_cmds))
    return tbl


# ── Scorecard table
def make_scorecard_table():
    headers = ["Indicator", "FY24/25", "Threshold", "Status"]
    rows = [
        ("PV Ext. Debt / GDP",      "28.2%",  "40%",   "✓  SAFE",    GREEN),
        ("PV Ext. Debt / Exports",  "114.3%", "180%",  "✓  SAFE",    GREEN),
        ("PV Ext. Debt / Revenue",  "167.4%", "250%",  "✓  SAFE",    GREEN),
        ("Ext. DS / Exports",       "36.6%",  "15%",   "⚠  BREACH",  RED),
        ("Ext. DS / Revenue",       "53.7%",  "18%",   "⚠  BREACH",  RED),
        ("PV Total Debt / GDP",     "63.7%",  "55%",   "⚠  BREACH",  RED),
        ("DS / Revenue (Total)",    "71.2%",  "30%",   "⚠  BREACH",  RED),
        ("Primary Balance / GDP",   "–2.3%",  "≥ 0%",  "⚠  DEFICIT", AMBER),
    ]

    data = [[h for h in headers]]
    for label, val, thr, status, _ in rows:
        data.append([label, val, thr, status])

    col_w = [6.0*cm, 2.5*cm, 2.5*cm, 2.8*cm]
    tbl = Table(data, colWidths=col_w, rowHeights=0.55*cm)

    cmds = [
        ("BACKGROUND", (0,0), (-1,0), NAVY),
        ("TEXTCOLOR",  (0,0), (-1,0), WHITE),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 8.5),
        ("ALIGN",      (1,0), (-1,-1), "CENTER"),
        ("ALIGN",      (0,0), (0,-1),  "LEFT"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("INNERGRID",  (0,0), (-1,-1), 0.3, MGREY),
        ("BOX",        (0,0), (-1,-1), 0.5, MGREY),
        ("LEFTPADDING",(0,0), (-1,-1), 5),
    ]
    # Row colours
    for i, (_, _, _, _, fill_color) in enumerate(rows):
        r = i + 1
        alt = LGREY if i % 2 == 0 else WHITE
        cmds.append(("BACKGROUND", (0,r), (2,r), alt))
        cmds.append(("BACKGROUND", (3,r), (3,r), fill_color))
        if fill_color == RED:
            cmds.append(("TEXTCOLOR", (3,r), (3,r), WHITE))
        elif fill_color == GREEN:
            cmds.append(("TEXTCOLOR", (3,r), (3,r), WHITE))

    tbl.setStyle(TableStyle(cmds))
    return tbl


# ── Scenario summary table
def make_scenario_table():
    headers = ["Scenario", "FY24/25", "FY25/26", "FY26/27", "FY27/28", "FY28/29", "FY29/30"]
    data = [headers]
    for name, vals in SCENARIOS.items():
        data.append([name] + [f"{v:.1f}%" for v in vals])

    col_w = [4.2*cm] + [1.9*cm]*6
    tbl = Table(data, colWidths=col_w, rowHeights=0.52*cm)

    cmds = [
        ("BACKGROUND", (0,0), (-1,0), NAVY),
        ("TEXTCOLOR",  (0,0), (-1,0), WHITE),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 8),
        ("ALIGN",      (1,0), (-1,-1), "CENTER"),
        ("ALIGN",      (0,0), (0,-1),  "LEFT"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("INNERGRID",  (0,0), (-1,-1), 0.3, MGREY),
        ("BOX",        (0,0), (-1,-1), 0.5, MGREY),
        ("LEFTPADDING",(0,0), (-1,-1), 5),
    ]

    scenario_colors = [BLUE, colors.HexColor("#C55A11"), colors.HexColor("#833C00"),
                       colors.HexColor("#7030A0"), RED, GREEN]
    for i, (name, vals) in enumerate(SCENARIOS.items()):
        r = i + 1
        cmds.append(("BACKGROUND", (0,r), (0,r), scenario_colors[i]))
        cmds.append(("TEXTCOLOR",  (0,r), (0,r), WHITE))
        cmds.append(("FONTNAME",   (0,r), (0,r), "Helvetica-Bold"))
        alt = LGREY if i % 2 == 0 else WHITE
        cmds.append(("BACKGROUND", (1,r), (-1,r), alt))
        # Colour cells that breach 55%
        for j, v in enumerate(vals):
            col = j + 1
            if v > 65:
                cmds.append(("BACKGROUND", (col,r), (col,r), RED))
                cmds.append(("TEXTCOLOR",  (col,r), (col,r), WHITE))
            elif v > 55:
                cmds.append(("BACKGROUND", (col,r), (col,r), AMBER))

    tbl.setStyle(TableStyle(cmds))
    return tbl


# ── Build PDF
def build_pdf():
    styles = make_styles()
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.8*cm,  bottomMargin=1.8*cm,
        title="Kenya Debt Sustainability Analysis",
        author="PDMO",
        subject="Public Debt Management"
    )

    story = []

    # ═══════════════════════════════════════
    # PAGE 1 — COVER
    # ═══════════════════════════════════════
    story.append(Spacer(1, 3.5*cm))

    story.append(Paragraph("REPUBLIC OF KENYA", styles["cover_title"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Public Debt Sustainability Analysis", styles["cover_title"]))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(
        "FY2022/23 – FY2029/30  |  IMF LIC-DSF Framework",
        styles["cover_sub"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Public Debt Management Office (PDMO)  |  National Treasury  |  September 2026",
        styles["cover_meta"]))

    story.append(Spacer(1, 1.5*cm))

    # Classification badge
    badge = Table(
        [["OFFICIAL — RESTRICTED"]],
        colWidths=[7*cm], rowHeights=[0.8*cm]
    )
    badge.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AMBER),
        ("TEXTCOLOR",  (0,0), (-1,-1), BLACK),
        ("FONTNAME",   (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 11),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("BOX",        (0,0), (-1,-1), 1, BLACK),
    ]))
    story.append(badge)
    story.append(Spacer(1, 1.5*cm))

    # Headline finding box
    finding = Table(
        [[Paragraph(
            "⚠  KEY FINDING: Kenya's debt service/revenue ratio of <b>71.2%</b> is "
            "more than <b>twice</b> the IMF stress threshold of 30%. "
            "PV total debt/GDP at <b>63.7%</b> breaches the 55% anchor. "
            "Classification: <b>SUSTAINABLE BUT AT HIGH RISK</b>.",
            ParagraphStyle("fb", fontSize=10, textColor=colors.HexColor("#7F0000"),
                           fontName="Helvetica-Bold", leading=15,
                           alignment=TA_CENTER))]],
        colWidths=[16.5*cm], rowHeights=None
    )
    finding.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FFE699")),
        ("BOX",        (0,0), (-1,-1), 1.5, RED),
        ("TOPPADDING", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),12),
        ("RIGHTPADDING",(0,0),(-1,-1),12),
    ]))
    story.append(finding)
    story.append(Spacer(1, 1.2*cm))

    # Contents
    toc_data = [
        ["1.", "Executive Summary & KPI Dashboard"],
        ["2.", "IMF LIC-DSF Scorecard"],
        ["3.", "Stress Test Scenarios"],
        ["4.", "Watchpoints & Policy Recommendations"],
    ]
    for num, title in toc_data:
        story.append(Paragraph(
            f"<font color='#FFD966'>{num}</font>  {title}",
            styles["toc"]))

    story.append(PageBreak())

    # ═══════════════════════════════════════
    # PAGE 2 — EXECUTIVE SUMMARY + SCORECARD
    # ═══════════════════════════════════════

    story.append(Paragraph("1.  EXECUTIVE SUMMARY", styles["section_head"]))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(
        "Kenya's public debt stock reached approximately <b>KES 6,215 billion</b> in "
        "FY2024/25, equivalent to <b>55.0% of GDP</b> when measured at face value, "
        "and <b>63.7% on a present-value basis</b> — breaching the IMF LIC-DSF composite "
        "anchor of 55%. The primary driver of elevated risk is not the stock level per se "
        "but the <b>debt service burden</b>: Kenya is spending <b>71 cents of every "
        "revenue shilling</b> on debt service, leaving minimal fiscal space for "
        "development and social expenditure.",
        styles["body"]))
    story.append(Spacer(1, 0.25*cm))

    story.append(Paragraph(
        "The domestic portfolio carries the most acute near-term risk. Treasury bills "
        "outstanding of <b>KES 1.04 trillion</b> with an average time to maturity of "
        "only <b>2.7 years</b> create significant rollover exposure. Interest payments "
        "on the domestic portfolio alone consumed <b>KES 820 billion</b> in FY24/25. "
        "The external portfolio, while larger in nominal terms, benefits from longer "
        "maturities and concessional rates through IDA, AfDB, and IMF facilities.",
        styles["body"]))
    story.append(Spacer(1, 0.3*cm))

    # KPI tiles
    story.append(make_kpi_table(styles))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "KPI Dashboard — FY2024/25 actuals. Red = threshold breach, Amber = elevated risk.",
        styles["caption"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Debt / GDP Trend (FY2022/23 – FY2029/30)", styles["body_bold"]))
    story.append(make_debt_gdp_chart())
    story.append(Paragraph(
        "Baseline debt/GDP trajectory vs 55% IMF composite anchor. "
        "Breach began FY2023/24 and persists through FY2027/28 under the baseline.",
        styles["caption"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Debt Service / Revenue (FY2022/23 – FY2029/30)", styles["body_bold"]))
    story.append(make_ds_rev_chart())
    story.append(Paragraph(
        "DS/Revenue peaked at 71.2% in FY24/25 — more than double the 30% stress threshold. "
        "Gradual decline projected under baseline assuming fiscal consolidation.",
        styles["caption"]))

    story.append(PageBreak())

    # ═══════════════════════════════════════
    # PAGE 3 — SCORECARD + STRESS TESTS + WATCHPOINTS
    # ═══════════════════════════════════════

    story.append(Paragraph("2.  IMF LIC-DSF SCORECARD  (FY2024/25)", styles["section_head"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Kenya is assessed under the <b>Strong policy space</b> category of the LIC-DSF "
        "framework. Four of seven threshold indicators are in breach, confirming the "
        "<i>sustainable but at high risk</i> classification.",
        styles["body"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(make_scorecard_table())
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "✓ Green = within safe range  |  ⚠ Red = threshold breached  |  "
        "⚠ Amber = elevated / approaching threshold",
        styles["caption"]))

    story.append(Spacer(1, 0.35*cm))
    story.append(Paragraph("3.  STRESS TEST SCENARIOS — Debt/GDP (%)", styles["section_head"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Six scenarios are modelled through FY2029/30. The revenue shortfall scenario "
        "(Finance Bill analog) is the most adverse, pushing debt/GDP to <b>72.8%</b> "
        "by FY2029/30. The optimistic reform scenario demonstrates that the 55% anchor "
        "is achievable by FY2028/29 with sustained fiscal discipline.",
        styles["body"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(make_scenario_table())
    story.append(Spacer(1, 0.15*cm))
    story.append(Paragraph(
        "Amber = above 55% anchor  |  Red = above 65% (high risk zone)",
        styles["caption"]))

    story.append(Spacer(1, 0.35*cm))
    story.append(Paragraph("4.  WATCHPOINTS & POLICY RECOMMENDATIONS", styles["section_head"]))
    story.append(Spacer(1, 0.25*cm))

    watchpoints = [
        ("🔴 Rollover Risk",
         "T-bill stock of KES 1.04 Tn matures on a rolling short-term basis. "
         "PDMO should extend the maturity profile through pre-issuance of longer bonds "
         "and infrastructure bonds to reduce refinancing pressure."),
        ("🔴 Revenue Mobilisation",
         "KRA must sustain tax-to-GDP improvements. "
         "A 1pp increase in revenue/GDP reduces DS/Revenue by ~1.4pp. "
         "Broadening the tax base and improving compliance are priority levers."),
        ("🟠 FX Risk Management",
         "With external debt at ~72% of total stock, a 20% KES depreciation adds "
         "~5.5pp to debt/GDP immediately. Active hedging and preferring KES-denominated "
         "instruments for new issuance is recommended."),
        ("🟠 Eurobond Liability Management",
         "KES 780 Bn in Eurobonds outstanding. The 2027 maturity is manageable given "
         "current reserve coverage (4.2 months), but early liability management exercises "
         "should be explored to smooth the repayment profile."),
        ("🟡 Interest Expense Reduction",
         "Interest payments at KES 1.15 Tn (47% of revenue) are the single largest "
         "expenditure item. Concessional borrowing from IDA/AfDB should be maximised "
         "to displace commercial borrowing at the margin."),
        ("🟢 Primary Balance Path",
         "The model shows primary surplus is achievable by FY2026/27 under the baseline. "
         "Achieving and sustaining a primary surplus of +0.5% of GDP is the single most "
         "important medium-term fiscal anchor."),
    ]

    wp_colors = [RED, RED, AMBER, AMBER, YELLOW, GREEN]
    wp_text_colors = [WHITE, WHITE, BLACK, BLACK, BLACK, WHITE]

    for i, ((tag, text), bg, tc) in enumerate(zip(watchpoints, wp_colors, wp_text_colors)):
        row_data = [[
            Paragraph(tag, ParagraphStyle(
                f"wt{i}", fontSize=9, fontName="Helvetica-Bold",
                textColor=tc, leading=12, alignment=TA_CENTER)),
            Paragraph(text, ParagraphStyle(
                f"wb{i}", fontSize=9, fontName="Helvetica",
                textColor=BLACK, leading=13, alignment=TA_JUSTIFY))
        ]]
        wp_tbl = Table(row_data, colWidths=[3.2*cm, 13.3*cm])
        wp_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (0,0), bg),
            ("BACKGROUND",    (1,0), (1,0), LGREY if i%2==0 else WHITE),
            ("VALIGN",        (0,0), (-1,-1), "TOP"),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 6),
            ("BOX",           (0,0), (-1,-1), 0.3, MGREY),
        ]))
        story.append(wp_tbl)
        story.append(Spacer(1, 0.08*cm))

    story.append(Spacer(1, 0.4*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=MGREY))
    story.append(Spacer(1, 0.15*cm))
    story.append(Paragraph(
        "Source: National Treasury, PDMO Debt Bulletin, IMF Article IV Consultation (2025). "
        "Model projections are indicative and subject to revision. "
        "For official debt statistics refer to PDMO published bulletins and the "
        "National Treasury Medium-Term Debt Management Strategy.",
        styles["footer"]))

    # Build
    doc.build(story, onFirstPage=_cover_background, onLaterPages=_header_footer)
    print(f"✅  PDF saved: {OUTPUT}")


if __name__ == "__main__":
    build_pdf()
