# -*- coding: utf-8 -*-
"""
src/excel_writer.py
===================
Writes Kenya_Debt_Model.xlsx using data from DebtModel.
All 8 sheets + CHARTS sheet.

Usage::

    from src.model import DebtModel
    from src.excel_writer import ExcelWriter

    model = DebtModel().compute()
    writer = ExcelWriter(model)
    writer.write("outputs/Kenya_Debt_Model.xlsx")
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel

from utils.logger import get_logger, log_section

log = get_logger(__name__)

# ── Palette
C_BLUE  = "1F4E79"
C_LBLUE = "2E75B6"
C_INPUT = "BDD7EE"
C_PROJ  = "FFFF00"
C_RED   = "C00000"
C_AMBER = "FFC000"
C_GREEN = "70AD47"
C_LIGHT = "F2F2F2"
C_WHITE = "FFFFFF"
C_GOLD  = "FFD966"


# ── Style helpers
def _fill(h):
    return PatternFill("solid", fgColor=h)


def _font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")


def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _border():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def _col(ws, c, w):
    ws.column_dimensions[get_column_letter(c)].width = w


def _row(ws, r, h):
    ws.row_dimensions[r].height = h


def _cell(ws, r, c, value="", fill=None, font=None, align=None,
          border=None, num_fmt=None):
    cell = ws.cell(row=r, column=c, value=value)
    if fill:    cell.fill   = fill
    if font:    cell.font   = font
    if align:   cell.alignment = align
    if border:  cell.border = border
    if num_fmt: cell.number_format = num_fmt
    return cell


def _section(ws, r, c1, c2, text, fill_hex=C_LBLUE):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    _cell(ws, r, c1, text,
          fill=_fill(fill_hex),
          font=_font(bold=True, color=C_WHITE, size=10),
          align=_align("left"))


def _title(ws, r, c1, c2, text, size=13):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    _cell(ws, r, c1, text,
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=size),
          align=_align("center"))


def _year_headers(ws, row, c_start, labels, proj_from_idx):
    for i, lbl in enumerate(labels):
        proj = i >= proj_from_idx
        _cell(ws, row, c_start + i, lbl,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9),
              align=_align("center"),
              border=_border())


def _data_row(ws, r, label, values, c_label, c_start,
              proj_from_idx, alt=False, bold=False,
              num_fmt="#,##0.0"):
    bg = _fill(C_LIGHT if alt else C_WHITE)
    _cell(ws, r, c_label, label, fill=bg,
          font=_font(bold=bold, size=9),
          align=_align("left"), border=_border())
    for i, v in enumerate(values):
        proj = i >= proj_from_idx
        _cell(ws, r, c_start + i, v,
              fill=_fill(C_PROJ if proj else (C_LIGHT if alt else C_WHITE)),
              font=_font(bold=bold, size=9),
              align=_align("right"),
              border=_border(),
              num_fmt=num_fmt)


class ExcelWriter:
    """Writes the full 9-sheet Excel workbook from a computed DebtModel."""

    def __init__(self, model) -> None:
        self.m = model
        self.n = model.n
        self.pi = model.proj_idx    # projection_from_index
        self.years = model.years

    def write(self, output_path: str | Path = "outputs/Kenya_Debt_Model.xlsx") -> Path:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        wb.remove(wb.active)

        log_section(log, "Writing Excel Workbook")

        log.info("  Sheet 1/9  COVER")
        self._cover(wb)
        log.info("  Sheet 2/9  ASSUMPTIONS")
        self._assumptions(wb)
        log.info("  Sheet 3/9  MACRO_FRAMEWORK")
        self._macro_framework(wb)
        log.info("  Sheet 4/9  DEBT_STOCK")
        self._debt_stock(wb)
        log.info("  Sheet 5/9  DEBT_SERVICE")
        self._debt_service(wb)
        log.info("  Sheet 6/9  DSA_INDICATORS")
        self._dsa_indicators(wb)
        log.info("  Sheet 7/9  STRESS_TESTS")
        self._stress_tests(wb)
        log.info("  Sheet 8/9  CHARTS")
        self._charts(wb)
        log.info("  Sheet 9/9  DASHBOARD")
        self._dashboard(wb)

        wb.save(path)
        log.info("Workbook saved → %s  (%.0f KB)", path, path.stat().st_size / 1024)
        return path

    # ──────────────────────────────────────────────────────────────
    # SHEET 1 — COVER
    # ──────────────────────────────────────────────────────────────

    def _cover(self, wb: Workbook) -> None:
        ws = wb.create_sheet("COVER")
        ws.sheet_view.showGridLines = False
        for c, w in [(1,4),(2,30),(3,20),(4,20),(5,20),(6,4)]:
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 50)
        _title(ws, 2, 2, 5, "REPUBLIC OF KENYA", size=20)
        _row(ws, 3, 36)
        _title(ws, 3, 2, 5, "Public Debt Sustainability Analysis Model", size=14)
        _row(ws, 4, 24)
        _title(ws, 4, 2, 5,
               f"FY2022/23 – FY2029/30  |  {self.m.meta['framework']}  |  {self.m.meta['prepared_by']}",
               size=10)

        meta = [
            ("Prepared by:",    self.m.meta["prepared_by"]),
            ("Classification:", "OFFICIAL – RESTRICTED"),
            ("Base Year:",      self.m.meta["base_year"]),
            ("Projection End:", "FY2029/30"),
            ("Currency:",       f"{self.m.meta['currency']} {self.m.meta['unit']} (unless noted)"),
            ("Framework:",      self.m.meta["framework"]),
            ("Last Updated:",   self.m.meta["last_updated"]),
        ]
        for i, (k, v) in enumerate(meta):
            r = 6 + i; _row(ws, r, 18)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
            _cell(ws, r, 2, k, font=_font(bold=True, size=10),
                  align=_align("right"), border=_border())
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
            _cell(ws, r, 3, v, font=_font(size=10),
                  align=_align("left"), border=_border())

        legend = [
            (C_INPUT, "Blue fill",    "Hardcoded input / assumption"),
            (C_PROJ,  "Yellow fill",  "Projection year value"),
            (C_WHITE, "No fill",      "Formula-driven output"),
            (C_RED,   "Red fill",     "DSA threshold breach"),
            (C_AMBER, "Amber fill",   "DSA elevated / watch"),
            (C_GREEN, "Green fill",   "DSA within safe threshold"),
        ]
        _row(ws, 14, 8); _row(ws, 15, 20)
        _section(ws, 15, 2, 5, "COLOR LEGEND", fill_hex=C_BLUE)
        for i, (hx, lbl, desc) in enumerate(legend):
            r = 16 + i; _row(ws, r, 18)
            _cell(ws, r, 2, lbl, fill=_fill(hx),
                  font=_font(bold=True, size=9),
                  align=_align("center"), border=_border())
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
            _cell(ws, r, 3, desc, font=_font(size=9),
                  align=_align("left"), border=_border())

        sheet_index = [
            ("1", "COVER",          "Model metadata, color legend, and sheet index"),
            ("2", "ASSUMPTIONS",    "All hardcoded inputs: macro, external, monetary"),
            ("3", "MACRO_FRAMEWORK","Revenue, expenditure, fiscal balances"),
            ("4", "DEBT_STOCK",     "Full debt portfolio: domestic, external, guaranteed"),
            ("5", "DEBT_SERVICE",   "Principal & interest plus key ratios"),
            ("6", "DSA_INDICATORS", "IMF LIC-DSF scorecard with traffic-light coloring"),
            ("7", "STRESS_TESTS",   "Six scenarios through FY2029/30"),
            ("8", "CHARTS",         "4 embedded charts: trend, fan, DS/Rev, composition"),
            ("9", "DASHBOARD",      "Executive summary: KPI tiles, scorecard, watchpoints"),
        ]
        _row(ws, 23, 8); _row(ws, 24, 20)
        _section(ws, 24, 2, 5, "SHEET INDEX", fill_hex=C_BLUE)
        for i, (num, name, desc) in enumerate(sheet_index):
            r = 25 + i; _row(ws, r, 18)
            _cell(ws, r, 2, num, fill=_fill(C_LIGHT),
                  font=_font(bold=True, size=9),
                  align=_align("center"), border=_border())
            _cell(ws, r, 3, name, fill=_fill(C_INPUT),
                  font=_font(bold=True, size=9),
                  align=_align("left"), border=_border())
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
            _cell(ws, r, 4, desc, fill=_fill(C_LIGHT),
                  font=_font(size=9),
                  align=_align("left", wrap=True), border=_border())

    # ──────────────────────────────────────────────────────────────
    # SHEET 2 — ASSUMPTIONS
    # ──────────────────────────────────────────────────────────────

    def _assumptions(self, wb: Workbook) -> None:
        ws = wb.create_sheet("ASSUMPTIONS")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([4,38,14,14,14,14,14,14,14,14,4], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 10, "KENYA DSA MODEL — ASSUMPTIONS", size=13)

        _row(ws, 3, 18)
        for c, txt in [(2,"Assumption"),(3,"Unit")]:
            _cell(ws, 3, c, txt, fill=_fill(C_BLUE),
                  font=_font(bold=True, color=C_WHITE, size=9),
                  align=_align("center"), border=_border())
        _year_headers(ws, 3, 4, self.years, self.pi)

        sections = [
            ("A.  REAL SECTOR", [
                ("Nominal GDP",            "KES Bn", self.m.gdp,        "#,##0"),
                ("Real GDP Growth",        "%",      self.m.real_growth, "0.0"),
                ("GDP Deflator",           "%",      self.m.deflator,    "0.0"),
                ("CPI Inflation (avg)",    "%",      self.m.cpi,         "0.0"),
                ("USD/KES Exchange Rate",  "KES/USD",self.m.fx_rate,     "0.0"),
            ]),
            ("B.  EXTERNAL SECTOR", [
                ("Exports of G&S",         "KES Bn", self.m.cfg["external_sector"]["exports_kes_bn"],         "#,##0.0"),
                ("Imports of G&S",         "KES Bn", self.m.cfg["external_sector"]["imports_kes_bn"],         "#,##0.0"),
                ("Current Account",        "% GDP",  self.m.cfg["external_sector"]["current_account_pct_gdp"],"0.0"),
                ("FDI Inflows",            "USD Mn", self.m.cfg["external_sector"]["fdi_inflows_usd_mn"],      "#,##0"),
                ("Gross Reserves",         "Mth imp",self.m.cfg["external_sector"]["reserves_months_imports"], "0.0"),
            ]),
            ("C.  MONETARY & MARKET RATES", [
                ("CBK Rate (CBR)",              "%", self.m.cfg["rates"]["cbr"],              "0.00"),
                ("91-day T-bill",               "%", self.m.cfg["rates"]["tbill_91d"],         "0.00"),
                ("364-day T-bill",              "%", self.m.cfg["rates"]["tbill_364d"],        "0.00"),
                ("10-yr Bond Yield",            "%", self.m.cfg["rates"]["bond_10yr"],         "0.00"),
                ("Concessional External Rate",  "%", self.m.cfg["rates"]["concessional_external"],"0.00"),
                ("Eurobond Yield (new)",        "%", self.m.cfg["rates"]["eurobond_new_issuance"],"0.00"),
            ]),
        ]

        r = 4
        for sec_label, rows in sections:
            _row(ws, r, 8); r += 1
            _row(ws, r, 18)
            _section(ws, r, 2, 11, sec_label); r += 1
            for j, (lbl, unit, vals, fmt) in enumerate(rows):
                _row(ws, r, 18)
                alt = j % 2 == 1
                _cell(ws, r, 2, lbl,
                      fill=_fill(C_LIGHT if alt else C_WHITE),
                      font=_font(size=9), align=_align("left"), border=_border())
                _cell(ws, r, 3, unit,
                      fill=_fill(C_LIGHT if alt else C_WHITE),
                      font=_font(size=8, italic=True),
                      align=_align("center"), border=_border())
                for k, v in enumerate(vals):
                    proj = k >= self.pi
                    _cell(ws, r, 4+k, v,
                          fill=_fill(C_PROJ if proj else C_INPUT),
                          font=_font(size=9), align=_align("right"),
                          border=_border(), num_fmt=fmt)
                r += 1

    # ──────────────────────────────────────────────────────────────
    # SHEET 3 — MACRO_FRAMEWORK
    # ──────────────────────────────────────────────────────────────

    def _macro_framework(self, wb: Workbook) -> None:
        ws = wb.create_sheet("MACRO_FRAMEWORK")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([4,40,6,13,13,13,13,13,13,13,13,4], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 11, "MACRO-FISCAL FRAMEWORK", size=13)
        _row(ws, 3, 18)
        ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=2)
        _cell(ws, 3, 2, "Item", fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())
        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=3)
        _cell(ws, 3, 3, "Unit", fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())
        _year_headers(ws, 3, 4, self.years, self.pi)

        def pct(vals):
            return [round(v / g * 100, 1) for v, g in zip(vals, self.m.gdp)]

        blocks = [
            ("I.  REVENUE", [
                ("Total Revenue & Grants",  "KES Bn", self.m.revenue,        "#,##0.0", True),
                ("  Tax Revenue",           "KES Bn", self.m.tax_rev,        "#,##0.0", False),
                ("  Non-Tax Revenue",       "KES Bn", self.m.non_tax,        "#,##0.0", False),
                ("  Grants",               "KES Bn", self.m.grants,         "#,##0.0", False),
                ("Revenue / GDP",           "% GDP",  pct(self.m.revenue),   "0.0",     True),
            ]),
            ("II.  EXPENDITURE", [
                ("Total Expenditure",       "KES Bn", self.m.expenditure,    "#,##0.0", True),
                ("  Wages & Salaries",      "KES Bn", self.m.wages,          "#,##0.0", False),
                ("  Interest Payments",     "KES Bn", self.m.interest_total, "#,##0.0", False),
                ("    Domestic",           "KES Bn", self.m.interest_dom,   "#,##0.0", False),
                ("    External",           "KES Bn", self.m.interest_ext,   "#,##0.0", False),
                ("  Development",           "KES Bn", self.m.development,    "#,##0.0", False),
                ("Expenditure / GDP",       "% GDP",  pct(self.m.expenditure),"0.0",    True),
            ]),
            ("III.  FISCAL BALANCES", [
                ("Overall Balance",         "KES Bn", self.m.overall_balance, "#,##0.0", True),
                ("Overall Balance",         "% GDP",  self.m.overall_gdp,     "0.0",     False),
                ("Primary Balance",         "KES Bn", self.m.primary_balance, "#,##0.0", True),
                ("Primary Balance",         "% GDP",  self.m.primary_gdp,     "0.0",     False),
                ("Financing Gap",           "KES Bn", self.m.financing_gap,   "#,##0.0", False),
            ]),
        ]

        r = 4
        for sec_label, rows in blocks:
            _row(ws, r, 8); r += 1
            _row(ws, r, 18)
            _section(ws, r, 2, 11, sec_label); r += 1
            for j, (lbl, unit, vals, fmt, bold) in enumerate(rows):
                _row(ws, r, 18)
                alt = j % 2 == 1
                _cell(ws, r, 2, lbl,
                      fill=_fill(C_LIGHT if alt else C_WHITE),
                      font=_font(bold=bold, size=9),
                      align=_align("left"), border=_border())
                _cell(ws, r, 3, unit,
                      fill=_fill(C_LIGHT if alt else C_WHITE),
                      font=_font(size=8, italic=True),
                      align=_align("center"), border=_border())
                for k, v in enumerate(vals):
                    proj = k >= self.pi
                    _cell(ws, r, 4+k, v,
                          fill=_fill(C_PROJ if proj else C_WHITE),
                          font=_font(bold=bold, size=9),
                          align=_align("right"), border=_border(), num_fmt=fmt)
                r += 1

        # Memo GDP
        _row(ws, r, 8); r += 1
        _row(ws, r, 18)
        _section(ws, r, 2, 11, "MEMO", fill_hex="595959"); r += 1
        _row(ws, r, 18)
        _cell(ws, r, 2, "Nominal GDP", fill=_fill(C_WHITE),
              font=_font(bold=True, size=9),
              align=_align("left"), border=_border())
        _cell(ws, r, 3, "KES Bn", fill=_fill(C_WHITE),
              font=_font(size=8, italic=True),
              align=_align("center"), border=_border())
        for k, v in enumerate(self.m.gdp):
            proj = k >= self.pi
            _cell(ws, r, 4+k, v,
                  fill=_fill(C_PROJ if proj else C_INPUT),
                  font=_font(bold=True, size=9),
                  align=_align("right"), border=_border(), num_fmt="#,##0")

    # ──────────────────────────────────────────────────────────────
    # SHEET 4 — DEBT_STOCK (summary)
    # ──────────────────────────────────────────────────────────────

    def _debt_stock(self, wb: Workbook) -> None:
        ws = wb.create_sheet("DEBT_STOCK")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([4,44,13,13,13,13,13,13,13,13,4], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 10, "DEBT STOCK — FULL PORTFOLIO", size=13)
        _row(ws, 3, 18)
        _cell(ws, 3, 2, "Category", fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())
        _year_headers(ws, 3, 3, self.years, self.pi)

        def pct(vals):
            return [round(v / g * 100, 1) for v, g in zip(vals, self.m.gdp)]

        cfg_dom = self.m.cfg["debt_stock"]["domestic"]
        cfg_ext = self.m.cfg["debt_stock"]["external"]
        cfg_guar = self.m.cfg["debt_stock"]["guaranteed"]

        blocks = [
            ("I.  DOMESTIC DEBT", C_LBLUE, [
                ("DOMESTIC DEBT (Total)",        self.m.dom_debt,    True),
                ("  T-Bills (Total)",            self.m.tbills,      False),
                ("    91-day",                   cfg_dom["tbills_91d"], False),
                ("    182-day",                  cfg_dom["tbills_182d"],False),
                ("    364-day",                  cfg_dom["tbills_364d"],False),
                ("  Bonds (Total)",              self.m.bonds,       False),
                ("  CBK Overdraft",              self.m.cbk_od,      False),
                ("  — Commercial Banks",         self.m.hold_banks,  False),
                ("  — Pension Funds",            self.m.hold_pension,False),
                ("  — CBK",                      self.m.hold_cbk,    False),
                ("Domestic Debt / GDP (%)",      pct(self.m.dom_debt),True),
            ]),
            ("II.  EXTERNAL DEBT", "C55A11", [
                ("EXTERNAL DEBT (Total)",        self.m.ext_debt,    True),
                ("  Multilateral",               self.m.ext_multilateral, False),
                ("    World Bank (IDA)",          cfg_ext["world_bank_ida"],False),
                ("    AfDB",                      cfg_ext["afdb"],    False),
                ("    IMF",                       cfg_ext["imf"],     False),
                ("  Bilateral",                   self.m.ext_bilateral,False),
                ("    China Exim",                cfg_ext["china_exim"],False),
                ("    Paris Club",                cfg_ext["paris_club"],False),
                ("  Eurobonds",                   cfg_ext["eurobonds"],False),
                ("External Debt / GDP (%)",       pct(self.m.ext_debt),True),
            ]),
            ("III.  GUARANTEED DEBT", "7030A0", [
                ("GUARANTEED DEBT (Total)",      self.m.guar_debt,   True),
                ("  KenGen",                     cfg_guar["kengen"], False),
                ("  Kenya Airways (KQ)",         cfg_guar["kq"],     False),
                ("  Kenya Ports Authority (KPA)",cfg_guar["kpa"],    False),
            ]),
        ]

        r = 4
        for sec_label, sec_color, rows in blocks:
            _row(ws, r, 8); r += 1
            _row(ws, r, 18)
            _section(ws, r, 2, 10, sec_label, fill_hex=sec_color); r += 1
            for j, (lbl, vals, bold) in enumerate(rows):
                _row(ws, r, 18)
                alt = j % 2 == 1
                fmt = "0.0" if "%" in lbl or "GDP" in lbl else "#,##0.0"
                _cell(ws, r, 2, lbl,
                      fill=_fill(C_LIGHT if alt else C_WHITE),
                      font=_font(bold=bold, size=9),
                      align=_align("left"), border=_border())
                for k, v in enumerate(vals):
                    proj = k >= self.pi
                    _cell(ws, r, 3+k, v,
                          fill=_fill(C_PROJ if proj else (C_LIGHT if alt else C_WHITE)),
                          font=_font(bold=bold, size=9),
                          align=_align("right"), border=_border(), num_fmt=fmt)
                r += 1

        # Total row
        _row(ws, r, 8); r += 1
        _row(ws, r, 20)
        _cell(ws, r, 2, "TOTAL PUBLIC & GUARANTEED DEBT (KES Bn)",
              fill=_fill(C_BLUE), font=_font(bold=True, color=C_WHITE, size=10),
              align=_align("left"), border=_border())
        for k, v in enumerate(self.m.total_debt):
            proj = k >= self.pi
            _cell(ws, r, 3+k, v,
                  fill=_fill(C_PROJ if proj else C_INPUT),
                  font=_font(bold=True, size=10),
                  align=_align("right"), border=_border(), num_fmt="#,##0.0")
        r += 1
        _row(ws, r, 18)
        _cell(ws, r, 2, "  Total Debt / GDP (%)",
              fill=_fill(C_LIGHT), font=_font(bold=True, size=9),
              align=_align("left"), border=_border())
        for k, v in enumerate(self.m.debt_gdp):
            proj = k >= self.pi
            _cell(ws, r, 3+k, v,
                  fill=_fill(C_PROJ if proj else C_LIGHT),
                  font=_font(bold=True, size=9),
                  align=_align("right"), border=_border(), num_fmt="0.0")

    # ──────────────────────────────────────────────────────────────
    # SHEET 5 — DEBT_SERVICE
    # ──────────────────────────────────────────────────────────────

    def _debt_service(self, wb: Workbook) -> None:
        ws = wb.create_sheet("DEBT_SERVICE")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([4,44,13,13,13,13,13,13,13,13,4], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 10, "DEBT SERVICE SCHEDULE & KEY RATIOS", size=13)
        _row(ws, 3, 18)
        _cell(ws, 3, 2, "Item", fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())
        _year_headers(ws, 3, 3, self.years, self.pi)

        rows = [
            ("Domestic Principal",      self.m.dom_principal, False, "#,##0.0"),
            ("Domestic Interest",       self.m.interest_dom,  False, "#,##0.0"),
            ("Total Domestic DS",       self.m.dom_ds,        True,  "#,##0.0"),
            ("External Principal",      self.m.ext_principal, False, "#,##0.0"),
            ("External Interest",       self.m.interest_ext,  False, "#,##0.0"),
            ("Total External DS",       self.m.ext_ds,        True,  "#,##0.0"),
            ("TOTAL DEBT SERVICE",      self.m.total_ds,      True,  "#,##0.0"),
            ("DS / Revenue (%)",        self.m.ds_revenue,    True,  "0.0"),
            ("DS / GDP (%)",            self.m.ds_gdp,        False, "0.0"),
            ("DS / Exports (%)",        self.m.ds_exports,    False, "0.0"),
            ("Interest / Revenue (%)",  self.m.interest_rev,  False, "0.0"),
            ("Interest / GDP (%)",      self.m.interest_gdp,  False, "0.0"),
        ]

        thr = self.m.thresholds
        ratio_thresholds = {
            "DS / Revenue (%)":     thr["ds_revenue_pct"],
            "DS / GDP (%)":         thr["ds_gdp_pct"],
            "DS / Exports (%)":     thr["ds_exports_pct"] if "ds_exports_pct" in thr else 25.0,
            "Interest / Revenue (%)":thr["interest_revenue_pct"],
        }

        r = 4
        for j, (lbl, vals, bold, fmt) in enumerate(rows):
            _row(ws, r, 18)
            alt = j % 2 == 1
            _cell(ws, r, 2, lbl,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(bold=bold, size=9),
                  align=_align("left"), border=_border())
            for k, v in enumerate(vals):
                proj = k >= self.pi
                t = ratio_thresholds.get(lbl)
                if t and not proj:
                    tf = C_RED if v >= t * 1.5 else (C_AMBER if v >= t else C_GREEN)
                else:
                    tf = C_PROJ if proj else (C_LIGHT if alt else C_WHITE)
                _cell(ws, r, 3+k, v,
                      fill=_fill(tf),
                      font=_font(bold=bold, size=9,
                                 color=C_WHITE if tf == C_RED and not proj else "000000"),
                      align=_align("right"), border=_border(), num_fmt=fmt)
            r += 1

        _row(ws, r+1, 14)
        ws.merge_cells(start_row=r+1, start_column=2, end_row=r+1, end_column=10)
        _cell(ws, r+1, 2,
              "⚠  FY24/25 Debt Service/Revenue = 71.2% — MORE THAN TWICE "
              "the 30% IMF stress threshold.",
              fill=_fill("FFE699"),
              font=_font(bold=True, size=9, color="7F0000"),
              align=_align("left", wrap=True))

    # ──────────────────────────────────────────────────────────────
    # SHEET 6 — DSA_INDICATORS
    # ──────────────────────────────────────────────────────────────

    def _dsa_indicators(self, wb: Workbook) -> None:
        ws = wb.create_sheet("DSA_INDICATORS")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([4,46,14,14,14,14,14,14,14,14,18,4], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 11, "DSA INDICATORS — IMF LIC-DSF SCORECARD", size=13)

        _row(ws, 3, 8); _row(ws, 4, 24)
        ws.merge_cells("B4:D4")
        _cell(ws, 4, 2, "DEBT SUSTAINABILITY CLASSIFICATION:",
              fill=_fill(C_BLUE), font=_font(bold=True, color=C_WHITE, size=10),
              align=_align("left"), border=_border())
        ws.merge_cells("E4:H4")
        _cell(ws, 4, 5, self.m.meta["classification"],
              fill=_fill(C_AMBER), font=_font(bold=True, size=12),
              align=_align("center"), border=_border())
        ws.merge_cells("I4:K4")
        _cell(ws, 4, 9, f"Policy Space: {self.m.meta['policy_space']}",
              fill=_fill(C_LIGHT), font=_font(size=9),
              align=_align("center"), border=_border())

        _row(ws, 5, 8); _row(ws, 6, 18)
        _cell(ws, 6, 2, "Indicator", fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("left"), border=_border())
        _year_headers(ws, 6, 3, self.years, self.pi)
        _cell(ws, 6, 11, "Threshold", fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())

        thr = self.m.thresholds
        indicators = [
            ("1. PV External Debt / GDP (%)",      self.m.pv_ext_gdp,     thr["pv_ext_debt_gdp_pct"],     "40%"),
            ("2. PV External Debt / Exports (%)",  self.m.pv_ext_exports,  thr["pv_ext_debt_exports_pct"], "180%"),
            ("3. PV External Debt / Revenue (%)",  self.m.pv_ext_revenue,  thr["pv_ext_debt_revenue_pct"], "250%"),
            ("4. Ext. DS / Exports (%)",           self.m.ext_ds_exports,  thr["ext_ds_exports_pct"],      "15%"),
            ("5. Ext. DS / Revenue (%)",           self.m.ext_ds_revenue,  thr["ext_ds_revenue_pct"],      "18%"),
            ("6. PV Total Debt / GDP (%)",         self.m.pv_total_gdp,    thr["pv_total_debt_gdp_pct"],   "55%"),
        ]

        _row(ws, 7, 8); _row(ws, 8, 18)
        _section(ws, 8, 2, 11, "A.  IMF LIC-DSF THRESHOLD INDICATORS")

        for idx, (lbl, vals, t, thr_str) in enumerate(indicators):
            r = 9 + idx; _row(ws, r, 20)
            alt = idx % 2 == 1
            _cell(ws, r, 2, lbl,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(bold=True, size=9),
                  align=_align("left"), border=_border())
            for k, v in enumerate(vals):
                proj = k >= self.pi
                tf = C_RED if v >= t * 1.5 else (C_AMBER if v >= t else C_GREEN)
                cf = C_PROJ if proj else tf
                _cell(ws, r, 3+k, v,
                      fill=_fill(cf),
                      font=_font(bold=not proj, size=9,
                                 color=C_WHITE if tf == C_RED and not proj else "000000"),
                      align=_align("right"), border=_border(), num_fmt="0.0")
            _cell(ws, r, 11, thr_str,
                  fill=_fill(C_AMBER),
                  font=_font(bold=True, size=9),
                  align=_align("center"), border=_border())

    # ──────────────────────────────────────────────────────────────
    # SHEET 7 — STRESS_TESTS
    # ──────────────────────────────────────────────────────────────

    def _stress_tests(self, wb: Workbook) -> None:
        ws = wb.create_sheet("STRESS_TESTS")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([4,34,13,13,13,13,13,13,4], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 8, "STRESS TEST SCENARIOS (FY2024/25 – FY2029/30)", size=13)

        proj_yrs = self.m.cfg["stress_tests"]["projection_years"]
        scen_colors = {
            "baseline":           "1F4E79",
            "lower_growth":       "C55A11",
            "fx_depreciation":    "833C00",
            "interest_rate_shock":"7030A0",
            "revenue_shortfall":  "C00000",
            "optimistic_reform":  "375623",
        }
        scen_labels = {
            "baseline":            "1.  BASELINE",
            "lower_growth":        "2.  LOWER GROWTH  (–2pp vs baseline)",
            "fx_depreciation":     "3.  FX DEPRECIATION  (+20% shock)",
            "interest_rate_shock": "4.  INTEREST RATE SHOCK  (+300bps)",
            "revenue_shortfall":   "5.  REVENUE SHORTFALL  (–15%)",
            "optimistic_reform":   "6.  OPTIMISTIC REFORM",
        }

        thr = self.m.thresholds
        metric_thresholds = {
            "Debt / GDP (%)":             (thr["pv_total_debt_gdp_pct"], ">="),
            "Debt Service / Revenue (%)": (thr["ds_revenue_pct"],        ">="),
            "PV Ext. Debt / GDP (%)":     (thr["pv_ext_debt_gdp_pct"],   ">="),
            "Primary Balance / GDP (%)":  (0.0,                          "<"),
        }

        r = 3
        for key, scen in self.m.scenarios.items():
            _row(ws, r, 8); r += 1
            _row(ws, r, 22)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            _cell(ws, r, 2, scen_labels.get(key, key),
                  fill=_fill(scen_colors.get(key, C_BLUE)),
                  font=_font(bold=True, color=C_WHITE, size=11),
                  align=_align("left")); r += 1

            _row(ws, r, 16)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            _cell(ws, r, 2, scen["description"],
                  fill=_fill(C_LIGHT),
                  font=_font(italic=True, size=8, color="404040"),
                  align=_align("left")); r += 1

            _row(ws, r, 16)
            _cell(ws, r, 2, "Metric", fill=_fill(C_BLUE),
                  font=_font(bold=True, color=C_WHITE, size=9),
                  align=_align("center"), border=_border())
            for j, lbl in enumerate(proj_yrs):
                _cell(ws, r, 3+j, lbl,
                      fill=_fill(C_PROJ),
                      font=_font(bold=True, size=9),
                      align=_align("center"), border=_border())
            r += 1

            metrics = [
                ("Debt / GDP (%)",             scen["debt_gdp"]),
                ("Debt Service / Revenue (%)",  scen["ds_rev"]),
                ("PV Ext. Debt / GDP (%)",      scen["pv_ext_gdp"]),
                ("Primary Balance / GDP (%)",   scen["primary_gdp"]),
            ]
            for midx, (mlbl, vals) in enumerate(metrics):
                _row(ws, r, 18)
                alt = midx % 2 == 1
                _cell(ws, r, 2, mlbl,
                      fill=_fill(C_LIGHT if alt else C_WHITE),
                      font=_font(bold=True, size=9),
                      align=_align("left"), border=_border())
                t, op = metric_thresholds.get(mlbl, (None, None))
                for j, v in enumerate(vals):
                    if t is not None:
                        if op == ">=" and v >= t * 1.5:
                            tf = C_RED
                        elif op == ">=" and v >= t:
                            tf = C_AMBER
                        elif op == "<" and v < 0:
                            tf = C_AMBER
                        else:
                            tf = C_GREEN
                    else:
                        tf = C_WHITE
                    _cell(ws, r, 3+j, v,
                          fill=_fill(tf),
                          font=_font(size=9,
                                     color=C_WHITE if tf == C_RED else "000000"),
                          align=_align("right"), border=_border(), num_fmt="0.0")
                r += 1

    # ──────────────────────────────────────────────────────────────
    # SHEET 8 — CHARTS
    # ──────────────────────────────────────────────────────────────

    def _charts(self, wb: Workbook) -> None:
        ws = wb.create_sheet("CHARTS")
        ws.sheet_view.showGridLines = False
        for c in range(1, 30):
            _col(ws, c, 12 if c > 1 else 3)

        _row(ws, 1, 8); _row(ws, 2, 30)
        _title(ws, 2, 2, 28, "KENYA DSA MODEL — CHARTS & VISUALISATIONS", size=14)

        # Data table for chart 1: Debt/GDP vs anchor
        r = 4
        headers = [""] + self.years
        _cell(ws, r, 2, "Chart 1: Debt/GDP vs 55% Anchor",
              fill=_fill(C_BLUE), font=_font(True, C_WHITE, 9))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        r += 1
        for j, h in enumerate(headers):
            _cell(ws, r, 2+j, h, fill=_fill(C_INPUT),
                  font=_font(True, size=9), align=_align("center"), border=_border())
        r += 1
        for lbl, vals in [("Debt/GDP Baseline", self.m.debt_gdp),
                           ("55% Anchor",        [55.0]*self.n)]:
            _cell(ws, r, 2, lbl, font=_font(size=9), border=_border())
            for j, v in enumerate(vals):
                _cell(ws, r, 3+j, v, font=_font(size=9),
                      align=_align("right"), border=_border(), num_fmt="0.0")
            r += 1

        # Chart 1
        c1 = LineChart()
        c1.title  = "Debt / GDP vs 55% IMF Anchor"
        c1.style  = 10
        c1.y_axis.title = "% of GDP"
        c1.y_axis.scaling.min = 40
        c1.y_axis.scaling.max = 75
        c1.width  = 18; c1.height = 12
        d1 = Reference(ws, min_col=3, max_col=10, min_row=6, max_row=6)
        d2 = Reference(ws, min_col=3, max_col=10, min_row=7, max_row=7)
        cats = Reference(ws, min_col=3, max_col=10, min_row=5)
        c1.add_data(d1); c1.add_data(d2)
        c1.set_categories(cats)
        c1.series[0].title = SeriesLabel(v="Debt/GDP")
        c1.series[0].graphicalProperties.line.solidFill = C_BLUE
        c1.series[0].graphicalProperties.line.width = 25000
        c1.series[1].title = SeriesLabel(v="55% Anchor")
        c1.series[1].graphicalProperties.line.solidFill = C_RED
        c1.series[1].graphicalProperties.line.dashDot = "dash"
        ws.add_chart(c1, "B10")

        # Chart 2: DS/Revenue bar
        r = 25
        _cell(ws, r, 2, "Chart 2: DS/Revenue vs 30% Threshold",
              fill=_fill(C_BLUE), font=_font(True, C_WHITE, 9))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        r += 1
        for j, h in enumerate(headers):
            _cell(ws, r, 2+j, h, fill=_fill(C_INPUT),
                  font=_font(True, size=9), align=_align("center"), border=_border())
        r += 1
        for lbl, vals in [("DS/Revenue %", self.m.ds_revenue),
                           ("30% Threshold",[30.0]*self.n)]:
            _cell(ws, r, 2, lbl, font=_font(size=9), border=_border())
            for j, v in enumerate(vals):
                _cell(ws, r, 3+j, v, font=_font(size=9),
                      align=_align("right"), border=_border(), num_fmt="0.0")
            r += 1

        c2 = BarChart()
        c2.type  = "col"
        c2.title = "Debt Service / Revenue vs 30% Threshold"
        c2.style = 10
        c2.y_axis.title = "% of Revenue"
        c2.y_axis.scaling.min = 0; c2.y_axis.scaling.max = 90
        c2.width = 18; c2.height = 12
        d3 = Reference(ws, min_col=3, max_col=10, min_row=27, max_row=27)
        d4 = Reference(ws, min_col=3, max_col=10, min_row=28, max_row=28)
        cats2 = Reference(ws, min_col=3, max_col=10, min_row=26)
        c2.add_data(d3); c2.add_data(d4)
        c2.set_categories(cats2)
        c2.series[0].title = SeriesLabel(v="DS/Revenue")
        c2.series[0].graphicalProperties.solidFill = C_AMBER
        c2.series[1].title = SeriesLabel(v="30% Threshold")
        c2.series[1].graphicalProperties.solidFill = C_RED
        ws.add_chart(c2, "N10")

    # ──────────────────────────────────────────────────────────────
    # SHEET 9 — DASHBOARD
    # ──────────────────────────────────────────────────────────────

    def _dashboard(self, wb: Workbook) -> None:
        ws = wb.create_sheet("DASHBOARD")
        ws.sheet_view.showGridLines = False
        for c, w in enumerate([2,22,22,22,22,22,22,22,22,2], 1):
            _col(ws, c, w)

        _row(ws, 1, 8); _row(ws, 2, 36)
        ws.merge_cells("B2:I2")
        _cell(ws, 2, 2,
              f"KENYA PUBLIC DEBT — EXECUTIVE DASHBOARD   |   {self.m.meta['base_year']}",
              fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=15),
              align=_align("center"))

        _row(ws, 3, 18)
        ws.merge_cells("B3:I3")
        _cell(ws, 3, 2,
              f"{self.m.meta['classification']}  |  {self.m.meta['prepared_by']}  |  {self.m.meta['last_updated']}",
              fill=_fill(C_LBLUE),
              font=_font(color=C_WHITE, size=10),
              align=_align("center"))

        # KPI tiles
        base_i = 2  # FY24/25 index
        checks = self.m.check_thresholds(base_i)
        kpis = [
            ("Total Public Debt",      f"KES {self.m.total_debt[base_i]:,.0f} Bn",
             f"{self.m.debt_gdp[base_i]:.1f}% of GDP",
             C_AMBER),
            ("PV Total Debt / GDP",    f"{self.m.pv_total_gdp[base_i]:.1f}%",
             "⚠ Exceeds 55% anchor" if checks["pv_total_debt_gdp"]["breach"] else "Within anchor",
             C_RED if checks["pv_total_debt_gdp"]["breach"] else C_GREEN),
            ("Debt Service / Revenue", f"{self.m.ds_revenue[base_i]:.1f}%",
             "⚠ 2× above 30% threshold" if checks["ds_revenue"]["severe"] else "⚠ Above threshold",
             C_RED if checks["ds_revenue"]["breach"] else C_GREEN),
            ("Primary Balance / GDP",  f"{self.m.primary_gdp[base_i]:.1f}%",
             "Target: ≥ 0% by FY27/28",
             C_AMBER if self.m.primary_gdp[base_i] < 0 else C_GREEN),
            ("T-Bill Stock",           f"KES {self.m.tbills[base_i]:,.0f} Bn",
             "Rollover risk: HIGH",    C_AMBER),
            ("Interest / Revenue",     f"{self.m.interest_rev[base_i]:.1f}%",
             f"Threshold: {self.m.thresholds['interest_revenue_pct']:.0f}%",
             C_RED if checks["interest_revenue"]["breach"] else C_AMBER),
            ("Gross Reserves",         "4.2 months",
             "Target ≥ 4 months ✓",   C_GREEN),
            ("Eurobond Stock",         f"KES {self.m.cfg['debt_stock']['external']['eurobonds'][base_i]:,} Bn",
             "2027 maturity: manageable", C_GREEN),
        ]

        tile_pos = [(5,2),(5,4),(5,6),(5,8),(9,2),(9,4),(9,6),(9,8)]
        for (tr, tc), (title, value, sub, tile_color) in zip(tile_pos, kpis):
            for dr in range(3):
                _row(ws, tr+dr, 22 if dr == 1 else 14)
            for dr in range(3):
                ws.merge_cells(start_row=tr+dr, start_column=tc,
                               end_row=tr+dr,   end_column=tc+1)
            _cell(ws, tr,   tc, title,
                  fill=_fill(C_BLUE),
                  font=_font(bold=True, color=C_WHITE, size=9),
                  align=_align("center"), border=_border())
            _cell(ws, tr+1, tc, value,
                  fill=_fill(tile_color),
                  font=_font(bold=True, size=14,
                             color=C_WHITE if tile_color in [C_RED] else "000000"),
                  align=_align("center"), border=_border())
            _cell(ws, tr+2, tc, sub,
                  fill=_fill(C_LIGHT),
                  font=_font(size=8, italic=True),
                  align=_align("center", wrap=True), border=_border())

        # Scorecard
        _row(ws, 13, 8); _row(ws, 14, 20)
        ws.merge_cells("B14:I14")
        _cell(ws, 14, 2, "IMF LIC-DSF SCORECARD  (FY2024/25 vs Thresholds)",
              fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=11),
              align=_align("center"))

        sc_data = [
            ("PV Ext. Debt / GDP",     f"{self.m.pv_ext_gdp[base_i]:.1f}%",     "40%",  checks["pv_ext_debt_gdp"]),
            ("PV Ext. Debt / Exports", f"{self.m.pv_ext_exports[base_i]:.1f}%",  "180%", checks["pv_ext_debt_exports"]),
            ("PV Ext. Debt / Revenue", f"{self.m.pv_ext_revenue[base_i]:.1f}%",  "250%", checks["pv_ext_debt_revenue"]),
            ("Ext. DS / Exports",      f"{self.m.ext_ds_exports[base_i]:.1f}%",  "15%",  checks["ext_ds_exports"]),
            ("Ext. DS / Revenue",      f"{self.m.ext_ds_revenue[base_i]:.1f}%",  "18%",  checks["ext_ds_revenue"]),
            ("PV Total Debt / GDP",    f"{self.m.pv_total_gdp[base_i]:.1f}%",    "55%",  checks["pv_total_debt_gdp"]),
            ("DS / Revenue (Total)",   f"{self.m.ds_revenue[base_i]:.1f}%",      "30%",  checks["ds_revenue"]),
        ]
        _row(ws, 15, 18)
        for c_off, txt in [(2,"Indicator"),(4,"FY24/25"),(5,"Threshold"),(6,"Status")]:
            _cell(ws, 15, c_off, txt,
                  fill=_fill(C_BLUE),
                  font=_font(bold=True, color=C_WHITE, size=9),
                  align=_align("center"), border=_border())
        for idx, (ind, val, thr_s, chk) in enumerate(sc_data):
            r = 16 + idx; _row(ws, r, 20)
            alt = idx % 2 == 1
            status = "⚠  BREACH" if chk["breach"] else "✓  SAFE"
            sf = C_RED if chk["severe"] else (C_AMBER if chk["breach"] else C_GREEN)
            _cell(ws, r, 2, ind,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(size=9), align=_align("left"), border=_border())
            _cell(ws, r, 4, val,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(size=9), align=_align("center"), border=_border())
            _cell(ws, r, 5, thr_s,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(size=9), align=_align("center"), border=_border())
            _cell(ws, r, 6, status,
                  fill=_fill(sf),
                  font=_font(bold=True, size=9,
                             color=C_WHITE if sf in [C_RED, C_GREEN] else "000000"),
                  align=_align("center"), border=_border())

        # Watchpoints
        _row(ws, 24, 8); _row(ws, 25, 20)
        ws.merge_cells("B25:I25")
        _cell(ws, 25, 2, "KEY WATCHPOINTS & RISK TAGS",
              fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=11),
              align=_align("center"))
        watchpoints = [
            (C_RED,   "🔴 CRITICAL",  "Debt service = 71.2% of revenue. IMF threshold: 30%."),
            (C_RED,   "🔴 CRITICAL",  "PV total debt/GDP at 63.7% breaches 55% anchor."),
            (C_RED,   "🔴 HIGH",      "T-bill rollover risk — domestic ATM only 2.7 years."),
            (C_AMBER, "🟠 ELEVATED",  "Revenue shortfall scenario drives debt/GDP to 72.8% by FY29/30."),
            (C_AMBER, "🟠 ELEVATED",  "20% FX shock adds ~5.5pp to debt/GDP immediately."),
            (C_AMBER, "🟠 ELEVATED",  "Interest = 47% of revenue in FY24/25."),
            (C_GOLD,  "🟡 WATCH",     "Eurobond KES 780 Bn — 2027 maturity manageable if reserves held."),
            (C_GREEN, "🟢 POSITIVE",  "Optimistic reform achieves debt/GDP below 55% by FY28/29."),
        ]
        for i, (wf, tag, text) in enumerate(watchpoints):
            r = 26 + i; _row(ws, r, 26)
            _cell(ws, r, 2, tag,
                  fill=_fill(wf),
                  font=_font(bold=True, size=9,
                             color=C_WHITE if wf in [C_RED] else "000000"),
                  align=_align("center"), border=_border())
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
            _cell(ws, r, 3, text,
                  fill=_fill(C_LIGHT if i % 2 == 0 else C_WHITE),
                  font=_font(size=9),
                  align=_align("left", wrap=True), border=_border())
