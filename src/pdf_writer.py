# -*- coding: utf-8 -*-
"""
src/pdf_writer.py
=================
Generates Kenya_Debt_Report.pdf from a computed DebtModel.

Usage::

    from src.model import DebtModel
    from src.pdf_writer import PDFWriter

    model = DebtModel().compute()
    writer = PDFWriter(model)
    writer.write("outputs/Kenya_Debt_Report.pdf")

Requires: reportlab
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.shapes import Drawing, Rect, String

from utils.logger import get_logger, log_section

log = get_logger(__name__)

# ── Colours
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


def _style(name, **kw):
    return ParagraphStyle(name, **kw)


def _make_styles():
    return {
        "title": _style("title", fontSize=24, textColor=WHITE,
                        fontName="Helvetica-Bold", alignment=TA_CENTER,
                        spaceAfter=6, leading=30),
        "sub":   _style("sub", fontSize=12, textColor=WHITE,
                        fontName="Helvetica", alignment=TA_CENTER,
                        spaceAfter=4, leading=16),
        "meta":  _style("meta", fontSize=9, textColor=LGREY,
                        fontName="Helvetica", alignment=TA_CENTER,
                        spaceAfter=3, leading=13),
        "sec":   _style("sec", fontSize=12, textColor=WHITE,
                        fontName="Helvetica-Bold", alignment=TA_LEFT,
                        spaceAfter=4, leading=15, backColor=NAVY,
                        borderPadding=(4, 6, 4, 6)),
        "body":  _style("body", fontSize=9.5, textColor=BLACK,
                        fontName="Helvetica", alignment=TA_JUSTIFY,
                        spaceAfter=4, leading=14),
        "bold":  _style("bold", fontSize=9.5, textColor=BLACK,
                        fontName="Helvetica-Bold", alignment=TA_LEFT,
                        spaceAfter=3, leading=14),
        "cap":   _style("cap", fontSize=8, textColor=DGREY,
                        fontName="Helvetica-Oblique", alignment=TA_CENTER,
                        spaceAfter=6, leading=11),
        "foot":  _style("foot", fontSize=7.5, textColor=MGREY,
                        fontName="Helvetica", alignment=TA_CENTER, leading=10),
        "toc":   _style("toc", fontSize=10, textColor=LGREY,
                        fontName="Helvetica", alignment=TA_LEFT,
                        spaceAfter=4, leading=16),
        "warn":  _style("warn", fontSize=10, textColor=colors.HexColor("#7F0000"),
                        fontName="Helvetica-Bold", alignment=TA_CENTER,
                        leading=14),
    }


def _header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4
    canvas.setStrokeColor(NAVY); canvas.setLineWidth(2)
    canvas.line(1.5*cm, h - 1.1*cm, w - 1.5*cm, h - 1.1*cm)
    canvas.setFont("Helvetica-Bold", 8); canvas.setFillColor(NAVY)
    canvas.drawString(1.5*cm, h - 0.85*cm,
                      "KENYA PUBLIC DEBT SUSTAINABILITY ANALYSIS")
    canvas.setFont("Helvetica", 8); canvas.setFillColor(DGREY)
    canvas.drawRightString(w - 1.5*cm, h - 0.85*cm,
                           "OFFICIAL — RESTRICTED  |  PDMO  |  Sep 2026")
    canvas.setStrokeColor(MGREY); canvas.setLineWidth(0.5)
    canvas.line(1.5*cm, 1.3*cm, w - 1.5*cm, 1.3*cm)
    canvas.setFont("Helvetica", 7.5); canvas.setFillColor(MGREY)
    canvas.drawCentredString(w/2, 0.9*cm,
        f"Page {doc.page}  |  Projections are model outputs. "
        "For official data refer to PDMO bulletins.")
    canvas.restoreState()


def _cover_bg(canvas, doc):
    if doc.page == 1:
        canvas.saveState()
        w, h = A4
        canvas.setFillColor(NAVY); canvas.rect(0, 0, w, h, fill=1, stroke=0)
        canvas.setFillColor(BLUE); canvas.rect(0, h*0.38, w, h*0.62, fill=1, stroke=0)
        canvas.setFillColor(AMBER); canvas.rect(0, 0, w, 0.8*cm, fill=1, stroke=0)
        canvas.restoreState()
    else:
        _header_footer(canvas, doc)


def _bar_chart(data_series, categories, title, y_title, colors_list, y_max):
    d = Drawing(400, 160)
    bc = VerticalBarChart()
    bc.x = 40; bc.y = 20; bc.width = 340; bc.height = 120
    bc.data = data_series
    bc.categoryAxis.categoryNames = categories
    bc.categoryAxis.labels.fontSize = 7
    bc.categoryAxis.labels.angle = 30
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = y_max
    bc.valueAxis.valueStep = y_max // 8
    bc.valueAxis.labels.fontSize = 7
    for i, c in enumerate(colors_list):
        bc.bars[i].fillColor = c
        bc.bars[i].strokeColor = c
    bc.barWidth = 0.35 * cm
    d.add(bc)
    return d


def _line_chart(data_series, categories, y_min, y_max, line_colors):
    d = Drawing(400, 160)
    lc = HorizontalLineChart()
    lc.x = 40; lc.y = 20; lc.width = 340; lc.height = 120
    lc.data = data_series
    lc.categoryAxis.categoryNames = categories
    lc.categoryAxis.labels.fontSize = 7
    lc.categoryAxis.labels.angle = 30
    lc.valueAxis.valueMin = y_min
    lc.valueAxis.valueMax = y_max
    lc.valueAxis.valueStep = 5
    lc.valueAxis.labels.fontSize = 7
    for i, c in enumerate(line_colors):
        lc.lines[i].strokeColor = c
        lc.lines[i].strokeWidth = 2
    d.add(lc)
    return d


def _scorecard_table(model, base_i):
    checks = model.check_thresholds(base_i)
    data = [["Indicator", "FY24/25", "Threshold", "Status"]]
    rows_meta = [
        ("PV Ext. Debt / GDP",     f"{model.pv_ext_gdp[base_i]:.1f}%",    "40%",  checks["pv_ext_debt_gdp"]),
        ("PV Ext. Debt / Exports", f"{model.pv_ext_exports[base_i]:.1f}%", "180%", checks["pv_ext_debt_exports"]),
        ("PV Ext. Debt / Revenue", f"{model.pv_ext_revenue[base_i]:.1f}%", "250%", checks["pv_ext_debt_revenue"]),
        ("Ext. DS / Exports",      f"{model.ext_ds_exports[base_i]:.1f}%", "15%",  checks["ext_ds_exports"]),
        ("Ext. DS / Revenue",      f"{model.ext_ds_revenue[base_i]:.1f}%", "18%",  checks["ext_ds_revenue"]),
        ("PV Total Debt / GDP",    f"{model.pv_total_gdp[base_i]:.1f}%",   "55%",  checks["pv_total_debt_gdp"]),
        ("DS / Revenue (Total)",   f"{model.ds_revenue[base_i]:.1f}%",     "30%",  checks["ds_revenue"]),
        ("Primary Balance / GDP",  f"{model.primary_gdp[base_i]:.1f}%",    "≥ 0%", {"breach": model.primary_gdp[base_i] < 0, "severe": False}),
    ]
    for r in rows_meta:
        data.append([r[0], r[1], r[2],
                     "⚠ BREACH" if r[3]["breach"] else "✓ SAFE"])

    col_w = [5.8*cm, 2.4*cm, 2.4*cm, 2.7*cm]
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
    for i, (_, _, _, chk) in enumerate(rows_meta):
        r = i + 1
        alt = LGREY if i % 2 == 0 else WHITE
        cmds.append(("BACKGROUND", (0,r), (2,r), alt))
        sf = RED if chk.get("severe") else (AMBER if chk["breach"] else GREEN)
        cmds.append(("BACKGROUND", (3,r), (3,r), sf))
        if sf in (RED, GREEN):
            cmds.append(("TEXTCOLOR", (3,r), (3,r), WHITE))
    tbl.setStyle(TableStyle(cmds))
    return tbl


class PDFWriter:
    """Generates a 3-page executive PDF briefing from a computed DebtModel."""

    def __init__(self, model) -> None:
        self.m = model
        self.styles = _make_styles()

    def write(self, output_path: str | Path = "outputs/Kenya_Debt_Report.pdf") -> Path:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        log_section(log, "Writing PDF Report")

        doc = SimpleDocTemplate(
            str(path), pagesize=A4,
            leftMargin=1.5*cm, rightMargin=1.5*cm,
            topMargin=1.8*cm,  bottomMargin=1.8*cm,
            title="Kenya Debt Sustainability Analysis",
            author=self.m.meta["prepared_by"],
        )

        story = self._cover() + self._page2() + self._page3()
        doc.build(story, onFirstPage=_cover_bg, onLaterPages=_header_footer)

        log.info("PDF saved → %s  (%.0f KB)", path, path.stat().st_size / 1024)
        return path

    def _cover(self):
        s = self.styles
        story = [Spacer(1, 3.5*cm)]
        story.append(Paragraph("REPUBLIC OF KENYA", s["title"]))
        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph("Public Debt Sustainability Analysis", s["title"]))
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph(
            f"FY2022/23 – FY2029/30  |  {self.m.meta['framework']}", s["sub"]))
        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph(
            f"{self.m.meta['prepared_by']}  |  {self.m.meta['last_updated']}", s["meta"]))
        story.append(Spacer(1, 1.5*cm))

        badge = Table([["OFFICIAL — RESTRICTED"]], colWidths=[7*cm], rowHeights=[0.8*cm])
        badge.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), AMBER),
            ("FONTNAME",  (0,0),(-1,-1),"Helvetica-Bold"),
            ("FONTSIZE",  (0,0),(-1,-1), 11),
            ("ALIGN",     (0,0),(-1,-1),"CENTER"),
            ("VALIGN",    (0,0),(-1,-1),"MIDDLE"),
            ("BOX",       (0,0),(-1,-1), 1, BLACK),
        ]))
        story.append(badge); story.append(Spacer(1, 1.2*cm))

        base_i = 2
        finding = Table([[Paragraph(
            f"⚠  KEY FINDING: Kenya's debt service/revenue ratio of "
            f"<b>{self.m.ds_revenue[base_i]:.1f}%</b> is more than "
            f"<b>twice</b> the IMF threshold of 30%. "
            f"PV total debt/GDP at <b>{self.m.pv_total_gdp[base_i]:.1f}%</b> "
            f"breaches the 55% anchor. "
            f"Classification: <b>{self.m.meta['classification']}</b>.",
            ParagraphStyle("fb", fontSize=10,
                           textColor=colors.HexColor("#7F0000"),
                           fontName="Helvetica-Bold", leading=15,
                           alignment=TA_CENTER))]],
            colWidths=[16.5*cm])
        finding.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), YELLOW),
            ("BOX",       (0,0),(-1,-1), 1.5, RED),
            ("TOPPADDING",(0,0),(-1,-1), 10),
            ("BOTTOMPADDING",(0,0),(-1,-1),10),
            ("LEFTPADDING",(0,0),(-1,-1), 12),
            ("RIGHTPADDING",(0,0),(-1,-1),12),
        ]))
        story.append(finding); story.append(Spacer(1, 1.0*cm))

        for num, title in [
            ("1.", "Executive Summary & KPI Dashboard"),
            ("2.", "IMF LIC-DSF Scorecard & Stress Tests"),
            ("3.", "Watchpoints & Policy Recommendations"),
        ]:
            story.append(Paragraph(
                f"<font color='#FFD966'>{num}</font>  {title}", s["toc"]))
        story.append(PageBreak())
        return story

    def _page2(self):
        s = self.styles
        base_i = 2
        story = [Paragraph("1.  EXECUTIVE SUMMARY", s["sec"]), Spacer(1, 0.3*cm)]
        story.append(Paragraph(
            f"Kenya's public debt reached approximately "
            f"<b>KES {self.m.total_debt[base_i]:,.0f} billion</b> in "
            f"{self.m.meta['base_year']}, equivalent to "
            f"<b>{self.m.debt_gdp[base_i]:.1f}% of GDP</b> at face value and "
            f"<b>{self.m.pv_total_gdp[base_i]:.1f}% on a present-value basis</b> — "
            f"{'breaching' if self.m.pv_total_gdp[base_i] > 55 else 'remaining below'} "
            "the IMF LIC-DSF composite anchor of 55%. "
            f"Kenya is spending <b>{self.m.ds_revenue[base_i]/100:.2f} of every revenue "
            "shilling</b> on debt service, leaving minimal fiscal space for "
            "development expenditure.",
            s["body"]))
        story.append(Spacer(1, 0.25*cm))

        story.append(Paragraph("Debt / GDP — Baseline vs 55% Anchor", s["bold"]))
        story.append(_line_chart(
            [self.m.debt_gdp, [55.0]*self.m.n],
            self.m.years, 40, 75, [NAVY, RED]))
        story.append(Paragraph(
            "Debt/GDP breached the 55% IMF anchor in FY2023/24 and "
            "persists above threshold through FY2027/28 under the baseline.",
            s["cap"]))
        story.append(Spacer(1, 0.3*cm))

        story.append(Paragraph("Debt Service / Revenue vs 30% Threshold", s["bold"]))
        story.append(_bar_chart(
            [self.m.ds_revenue, [30.0]*self.m.n],
            self.m.years, "DS/Revenue %", "% of Revenue",
            [AMBER, RED], 90))
        story.append(Paragraph(
            f"DS/Revenue peaked at {self.m.ds_revenue[base_i]:.1f}% in "
            f"{self.m.meta['base_year']} — more than double the 30% stress threshold.",
            s["cap"]))
        story.append(PageBreak())
        return story

    def _page3(self):
        s = self.styles
        base_i = 2
        story = [Paragraph("2.  IMF LIC-DSF SCORECARD", s["sec"]),
                 Spacer(1, 0.3*cm)]
        story.append(Paragraph(
            "Kenya is assessed under the <b>Strong policy space</b> category. "
            "Four of seven threshold indicators are in breach.",
            s["body"]))
        story.append(Spacer(1, 0.2*cm))
        story.append(_scorecard_table(self.m, base_i))
        story.append(Spacer(1, 0.15*cm))
        story.append(Paragraph(
            "✓ Green = safe  |  ⚠ Amber = elevated  |  ⚠ Red = breach",
            s["cap"]))

        story.append(Spacer(1, 0.35*cm))
        story.append(Paragraph("3.  WATCHPOINTS & RECOMMENDATIONS", s["sec"]))
        story.append(Spacer(1, 0.2*cm))

        watchpoints = [
            (RED,    WHITE, "🔴 CRITICAL",
             f"Debt service = {self.m.ds_revenue[base_i]:.1f}% of revenue — IMF threshold is 30%. "
             "Fiscal space is severely compressed."),
            (RED if self.m.pv_total_gdp[base_i] > 55 else AMBER, WHITE, "🔴 CRITICAL",
             f"PV total debt/GDP at {self.m.pv_total_gdp[base_i]:.1f}% "
             f"{'breaches' if self.m.pv_total_gdp[base_i] > 55 else 'is below'} "
             "the 55% composite anchor."),
            (AMBER,  BLACK, "🟠 ELEVATED",
             "T-bill rollover risk: KES 1.04 Tn with domestic ATM of 2.7 years."),
            (AMBER,  BLACK, "🟠 ELEVATED",
             "Revenue shortfall scenario drives debt/GDP to 72.8% by FY2029/30."),
            (YELLOW, BLACK, "🟡 WATCH",
             "Eurobond KES 780 Bn — 2027 bullet manageable if reserves maintained."),
            (GREEN,  WHITE, "🟢 POSITIVE",
             "Optimistic reform achieves debt/GDP below 55% anchor by FY2028/29."),
        ]
        for i, (bg, fc, tag, text) in enumerate(watchpoints):
            row = [[
                Paragraph(tag, ParagraphStyle(
                    f"wt{i}", fontSize=9, fontName="Helvetica-Bold",
                    textColor=fc, leading=12, alignment=TA_CENTER)),
                Paragraph(text, ParagraphStyle(
                    f"wb{i}", fontSize=9, fontName="Helvetica",
                    textColor=BLACK, leading=13, alignment=TA_JUSTIFY))
            ]]
            tbl = Table(row, colWidths=[3.0*cm, 13.5*cm])
            tbl.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (0,0), bg),
                ("BACKGROUND",    (1,0), (1,0), LGREY if i%2==0 else WHITE),
                ("VALIGN",        (0,0), (-1,-1), "TOP"),
                ("TOPPADDING",    (0,0), (-1,-1), 5),
                ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                ("LEFTPADDING",   (0,0), (-1,-1), 6),
                ("BOX",           (0,0), (-1,-1), 0.3, MGREY),
            ]))
            story.append(tbl); story.append(Spacer(1, 0.06*cm))

        story.append(Spacer(1, 0.4*cm))
        story.append(HRFlowable(width="100%", thickness=0.5, color=MGREY))
        story.append(Spacer(1, 0.15*cm))
        story.append(Paragraph(
            "Source: National Treasury, PDMO Debt Bulletin, IMF Article IV (2025). "
            "Projections are indicative and subject to revision.",
            s["foot"]))
        return story
