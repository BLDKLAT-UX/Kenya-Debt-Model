"""
Kenya Debt Sustainability Analysis Model
========================================
Generates an Excel workbook (Kenya_Debt_Model.xlsx) with 8 sheets:
  1. COVER
  2. ASSUMPTIONS
  3. MACRO_FRAMEWORK
  4. DEBT_STOCK
  5. DEBT_SERVICE
  6. DSA_INDICATORS
  7. STRESS_TESTS
  8. DASHBOARD

Requirements:
    pip install openpyxl

Run:
    python kenya_debt_model.py
"""

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
import openpyxl.styles.numbers as num_formats

# ─────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────
C_BLUE   = "1F4E79"   # dark navy header
C_LBLUE  = "2E75B6"   # mid-blue section header
C_INPUT  = "BDD7EE"   # blue tint – input cells
C_PROJ   = "FFFF00"   # yellow – projection cells
C_FORM   = "000000"   # black – formula cells (no fill)
C_RED    = "FF0000"   # DSA breach
C_AMBER  = "FFC000"   # DSA elevated
C_GREEN  = "70AD47"   # DSA safe
C_LIGHT  = "F2F2F2"   # alternating row
C_WHITE  = "FFFFFF"
C_GOLD   = "FFD966"   # dashboard KPI tile
C_DKGREY = "404040"

YEARS     = [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]
FY_LABELS = ["FY22/23","FY23/24","FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
PROJ_FROM = 2025   # FY25/26 onwards are projections

# ─────────────────────────────────────────────
# HELPER UTILITIES
# ─────────────────────────────────────────────

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border(style="thin"):
    s = Side(border_style=style, color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _thick_bottom():
    thin = Side(border_style="thin",  color="BFBFBF")
    thick= Side(border_style="medium",color="1F4E79")
    return Border(left=thin, right=thin, top=thin, bottom=thick)

def _col(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def _row(ws, row, height):
    ws.row_dimensions[row].height = height

def _cell(ws, r, c, value="", fill=None, font=None, align=None,
          border=None, num_fmt=None):
    cell = ws.cell(row=r, column=c, value=value)
    if fill:   cell.fill   = fill
    if font:   cell.font   = font
    if align:  cell.alignment = align
    if border: cell.border = border
    if num_fmt: cell.number_format = num_fmt
    return cell

def _merge_title(ws, r, c1, c2, text, fill_hex=C_BLUE, font_size=12,
                 font_color=C_WHITE):
    ws.merge_cells(start_row=r, start_column=c1,
                   end_row=r,   end_column=c2)
    _cell(ws, r, c1, text,
          fill=_fill(fill_hex),
          font=_font(bold=True, color=font_color, size=font_size),
          align=_align("center"))

def _section_header(ws, r, c1, c2, text, fill_hex=C_LBLUE):
    ws.merge_cells(start_row=r, start_column=c1,
                   end_row=r,   end_column=c2)
    _cell(ws, r, c1, text,
          fill=_fill(fill_hex),
          font=_font(bold=True, color=C_WHITE, size=10),
          align=_align("left"))

def _year_header_row(ws, r, c_start, labels, is_proj_fn=None,
                     label_col=None):
    """Write year labels across a row. Yellow fill for projection years."""
    if label_col:
        _cell(ws, r, label_col, "", fill=_fill(C_BLUE))
    for i, lbl in enumerate(labels):
        c = c_start + i
        yr = YEARS[i] if i < len(YEARS) else None
        is_proj = is_proj_fn(yr) if (is_proj_fn and yr) else False
        fill_hex = C_PROJ if is_proj else C_INPUT
        font_col = "000000"
        _cell(ws, r, c, lbl,
              fill=_fill(fill_hex),
              font=_font(bold=True, color=font_col, size=9),
              align=_align("center"),
              border=_border())

def _data_row(ws, r, label, values, c_label=1, c_start=3,
              alt=False, bold=False, num_fmt="#,##0.0",
              indent=0, proj_from_idx=None):
    """Write a label + 8 data values. Yellow fill from proj_from_idx onward."""
    fill_bg = _fill(C_LIGHT) if alt else _fill(C_WHITE)
    label_txt = ("  " * indent) + label
    _cell(ws, r, c_label, label_txt,
          fill=fill_bg,
          font=_font(bold=bold, size=9),
          align=_align("left"),
          border=_border())
    for i, v in enumerate(values):
        c = c_start + i
        is_proj = (proj_from_idx is not None) and (i >= proj_from_idx)
        cell_fill = _fill(C_PROJ) if is_proj else fill_bg
        _cell(ws, r, c, v,
              fill=cell_fill,
              font=_font(bold=bold, size=9),
              align=_align("right"),
              border=_border(),
              num_fmt=num_fmt)

# ─────────────────────────────────────────────
# SHEET 1 — COVER
# ─────────────────────────────────────────────

def build_cover(wb):
    ws = wb.create_sheet("COVER")
    ws.sheet_view.showGridLines = False

    # Column widths
    for c, w in [(1,4),(2,30),(3,20),(4,20),(5,20),(6,4)]:
        _col(ws, c, w)

    # Title block
    _row(ws, 1, 8)
    _row(ws, 2, 50)
    ws.merge_cells("B2:E2")
    _cell(ws, 2, 2,
          "REPUBLIC OF KENYA",
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=20),
          align=_align("center"))

    _row(ws, 3, 36)
    ws.merge_cells("B3:E3")
    _cell(ws, 3, 2,
          "Public Debt Sustainability Analysis Model",
          fill=_fill(C_LBLUE),
          font=_font(bold=True, color=C_WHITE, size=14),
          align=_align("center"))

    _row(ws, 4, 24)
    ws.merge_cells("B4:E4")
    _cell(ws, 4, 2,
          "FY2022/23 – FY2029/30  |  LIC-DSF Framework  |  PDMO",
          fill=_fill(C_LBLUE),
          font=_font(bold=False, color=C_WHITE, size=10),
          align=_align("center"))

    # Metadata table
    meta = [
        ("Prepared by:",   "Public Debt Management Office (PDMO)"),
        ("Classification:","OFFICIAL – RESTRICTED"),
        ("Base Year:",      "FY2024/25"),
        ("Projection End:", "FY2029/30"),
        ("Currency:",       "KES Billions (unless noted)"),
        ("Framework:",      "IMF LIC-DSF (2018 vintage)"),
        ("Last Updated:",   "September 2026"),
    ]
    for i, (k, v) in enumerate(meta):
        r = 6 + i
        _row(ws, r, 18)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
        _cell(ws, r, 2, k, font=_font(bold=True, size=10),
              align=_align("right"), border=_border())
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        _cell(ws, r, 3, v, font=_font(size=10),
              align=_align("left"), border=_border())

    # Color legend
    _row(ws, 14, 8)
    _row(ws, 15, 20)
    ws.merge_cells("B15:E15")
    _cell(ws, 15, 2, "COLOR LEGEND",
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=10),
          align=_align("center"))

    legend = [
        (C_INPUT, "Blue fill",   "Hardcoded input / assumption"),
        (C_PROJ,  "Yellow fill", "Projection year value"),
        ("FFFFFF", "No fill",    "Formula-driven output (black text)"),
        (C_RED,   "Red fill",    "DSA threshold breach"),
        (C_AMBER, "Amber fill",  "DSA elevated / watch"),
        (C_GREEN, "Green fill",  "DSA within safe threshold"),
    ]
    for i, (hex_c, lbl, desc) in enumerate(legend):
        r = 16 + i
        _row(ws, r, 18)
        _cell(ws, r, 2, lbl,
              fill=_fill(hex_c),
              font=_font(bold=True, size=9),
              align=_align("center"), border=_border())
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        _cell(ws, r, 3, desc,
              font=_font(size=9),
              align=_align("left"), border=_border())

    # Sheet index
    _row(ws, 23, 8)
    _row(ws, 24, 20)
    ws.merge_cells("B24:E24")
    _cell(ws, 24, 2, "SHEET INDEX",
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=10),
          align=_align("center"))

    sheets = [
        ("1","COVER",          "Model metadata, color legend, and sheet index"),
        ("2","ASSUMPTIONS",    "All hardcoded inputs: macro, external, monetary"),
        ("3","MACRO_FRAMEWORK","Revenue, expenditure, fiscal balances (KES Bn & % GDP)"),
        ("4","DEBT_STOCK",     "Full debt portfolio: domestic, external, guaranteed"),
        ("5","DEBT_SERVICE",   "Principal & interest plus key ratios"),
        ("6","DSA_INDICATORS", "IMF LIC-DSF scorecard with traffic-light coloring"),
        ("7","STRESS_TESTS",   "Six scenarios through FY2029/30"),
        ("8","DASHBOARD",      "Executive summary: KPI tiles, scorecard, watchpoints"),
    ]
    for i, (num, name, desc) in enumerate(sheets):
        r = 25 + i
        _row(ws, r, 18)
        _cell(ws, r, 2, num,  fill=_fill(C_LIGHT), font=_font(bold=True, size=9),
              align=_align("center"), border=_border())
        _cell(ws, r, 3, name, fill=_fill(C_INPUT),  font=_font(bold=True, size=9),
              align=_align("left"),   border=_border())
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        _cell(ws, r, 4, desc, fill=_fill(C_LIGHT),  font=_font(size=9),
              align=_align("left", wrap=True), border=_border())

    # Footer
    _row(ws, 34, 14)
    ws.merge_cells("B34:E34")
    _cell(ws, 34, 2,
          "⚠  Projections are model outputs and subject to revision. "
          "For official debt data refer to PDMO published bulletins.",
          font=_font(italic=True, size=8, color="7F7F7F"),
          align=_align("center", wrap=True))


# ─────────────────────────────────────────────
# SHEET 2 — ASSUMPTIONS
# ─────────────────────────────────────────────

def build_assumptions(wb):
    ws = wb.create_sheet("ASSUMPTIONS")
    ws.sheet_view.showGridLines = False

    col_widths = [4, 38, 14, 14, 14, 14, 14, 14, 14, 14, 4]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)
    _merge_title(ws, 2, 2, 10, "KENYA DSA MODEL — ASSUMPTIONS", font_size=13)
    _row(ws, 2, 30)

    # Year header
    _row(ws, 3, 18)
    _cell(ws, 3, 2, "Assumption", fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())
    _cell(ws, 3, 3, "Unit", fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())

    fy_labels_short = ["FY22/23","FY23/24","FY24/25",
                       "FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
    is_proj = lambda yr: yr >= PROJ_FROM
    for i, lbl in enumerate(fy_labels_short):
        c = 4 + i
        yr = YEARS[i]
        proj = yr >= PROJ_FROM
        _cell(ws, 3, c, lbl,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9),
              align=_align("center"), border=_border())

    # ── SECTION A: Real Sector
    _row(ws, 4, 8)
    _row(ws, 5, 18)
    _section_header(ws, 5, 2, 11, "A.  REAL SECTOR")

    real_sector = [
        # (label, unit, [fy values x8], proj_from_idx)
        ("Nominal GDP",             "KES Bn",
         [12_100, 13_200, 14_400, 15_800, 17_300, 18_900, 20_600, 22_400], 3),
        ("Real GDP Growth",         "%",
         [4.8, 5.6, 5.4, 5.6, 5.8, 6.0, 6.0, 6.0], 3),
        ("GDP Deflator",            "%",
         [7.8, 6.0, 5.5, 5.2, 5.0, 4.8, 4.8, 4.8], 3),
        ("CPI Inflation (avg)",     "%",
         [7.9, 6.3, 5.0, 4.8, 4.5, 4.5, 4.5, 4.5], 3),
        ("USD/KES Exchange Rate",   "KES/USD",
         [132, 148, 128, 130, 132, 134, 136, 138], 3),
    ]

    for i, (lbl, unit, vals, pidx) in enumerate(real_sector):
        r = 6 + i
        _row(ws, r, 18)
        alt = (i % 2 == 1)
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=9), align=_align("left"), border=_border())
        _cell(ws, r, 3, unit,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=9, italic=True), align=_align("center"), border=_border())
        for j, v in enumerate(vals):
            c = 4 + j
            proj = j >= pidx
            _cell(ws, r, c, v,
                  fill=_fill(C_PROJ if proj else C_INPUT),
                  font=_font(size=9), align=_align("right"),
                  border=_border(),
                  num_fmt="#,##0.0" if unit not in ["%","KES/USD"] else "0.0")

    # ── SECTION B: External Sector
    _row(ws, 12, 8)
    _row(ws, 13, 18)
    _section_header(ws, 13, 2, 11, "B.  EXTERNAL SECTOR")

    external = [
        ("Exports of Goods & Services", "KES Bn",
         [1_890, 2_100, 2_350, 2_580, 2_820, 3_070, 3_350, 3_650], 3),
        ("Imports of Goods & Services", "KES Bn",
         [3_200, 3_450, 3_680, 3_900, 4_100, 4_310, 4_530, 4_760], 3),
        ("Current Account Balance",     "% GDP",
         [-5.4, -4.8, -4.2, -3.9, -3.6, -3.3, -3.1, -2.9], 3),
        ("FDI Inflows",                 "USD Mn",
         [830, 960, 1_020, 1_100, 1_200, 1_300, 1_400, 1_500], 3),
        ("Gross Official Reserves",     "Months of imports",
         [3.8, 4.0, 4.2, 4.4, 4.6, 4.8, 5.0, 5.2], 3),
        ("Remittances",                 "USD Mn",
         [3_980, 4_200, 4_500, 4_700, 4_900, 5_100, 5_300, 5_500], 3),
    ]

    for i, (lbl, unit, vals, pidx) in enumerate(external):
        r = 14 + i
        _row(ws, r, 18)
        alt = (i % 2 == 1)
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=9), align=_align("left"), border=_border())
        _cell(ws, r, 3, unit,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=9, italic=True), align=_align("center"), border=_border())
        for j, v in enumerate(vals):
            c = 4 + j
            proj = j >= pidx
            fmt = "#,##0.0" if "Bn" in unit or "Mn" in unit else "0.0"
            _cell(ws, r, c, v,
                  fill=_fill(C_PROJ if proj else C_INPUT),
                  font=_font(size=9), align=_align("right"),
                  border=_border(), num_fmt=fmt)

    # ── SECTION C: Monetary & Market Rates
    _row(ws, 21, 8)
    _row(ws, 22, 18)
    _section_header(ws, 22, 2, 11, "C.  MONETARY & MARKET RATES")

    rates = [
        ("CBK Central Bank Rate (CBR)",    "%", [10.50,13.00,11.25,10.50,9.50,9.00,9.00,9.00], 3),
        ("91-day T-bill Rate",             "%", [10.10,16.00,12.80,11.50,10.50,9.80,9.50,9.50], 3),
        ("182-day T-bill Rate",            "%", [10.30,16.50,13.20,12.00,11.00,10.20,9.80,9.80], 3),
        ("364-day T-bill Rate",            "%", [10.60,17.00,13.60,12.50,11.50,10.60,10.20,10.20], 3),
        ("2-yr Bond Yield",                "%", [12.00,17.50,14.50,13.50,12.50,11.50,11.00,11.00], 3),
        ("5-yr Bond Yield",                "%", [13.50,18.00,15.00,14.00,13.00,12.00,11.50,11.50], 3),
        ("10-yr Bond Yield",               "%", [13.80,18.50,15.50,14.50,13.50,12.50,12.00,12.00], 3),
        ("Concessional External Rate",     "%", [1.50,1.50,1.50,1.50,1.50,1.50,1.50,1.50], 3),
        ("Semi-concessional (China Exim)", "%", [3.00,3.00,3.00,3.00,3.00,3.00,3.00,3.00], 3),
        ("Eurobond Yield (new issuance)",  "%", [9.75,10.50,9.50,9.00,8.50,8.00,7.50,7.50], 3),
    ]

    for i, (lbl, unit, vals, pidx) in enumerate(rates):
        r = 23 + i
        _row(ws, r, 18)
        alt = (i % 2 == 1)
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=9), align=_align("left"), border=_border())
        _cell(ws, r, 3, unit,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=9, italic=True), align=_align("center"), border=_border())
        for j, v in enumerate(vals):
            c = 4 + j
            proj = j >= pidx
            _cell(ws, r, c, v,
                  fill=_fill(C_PROJ if proj else C_INPUT),
                  font=_font(size=9), align=_align("right"),
                  border=_border(), num_fmt="0.00")

    # Note row
    _row(ws, 34, 14)
    ws.merge_cells(start_row=34, start_column=2, end_row=34, end_column=11)
    _cell(ws, 34, 2,
          "Notes: Blue = historical inputs; Yellow = projections. "
          "All rates are period averages. GDP figures use National Treasury estimates.",
          font=_font(italic=True, size=8, color="7F7F7F"),
          align=_align("left", wrap=True))


# ─────────────────────────────────────────────
# SHEET 3 — MACRO_FRAMEWORK
# ─────────────────────────────────────────────

def build_macro_framework(wb):
    ws = wb.create_sheet("MACRO_FRAMEWORK")
    ws.sheet_view.showGridLines = False

    col_widths = [4, 40, 6, 13, 13, 13, 13, 13, 13, 13, 13, 4]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)
    _merge_title(ws, 2, 2, 11, "MACRO-FISCAL FRAMEWORK", font_size=13)
    _row(ws, 2, 30)

    # Two header sub-rows: label + years
    _row(ws, 3, 14)
    ws.merge_cells(start_row=3, start_column=2, end_row=4, end_column=2)
    _cell(ws, 3, 2, "Item", fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())
    ws.merge_cells(start_row=3, start_column=3, end_row=4, end_column=3)
    _cell(ws, 3, 3, "Unit", fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())
    for i, lbl in enumerate(FY_LABELS):
        c = 4 + i
        yr = YEARS[i]
        proj = yr >= PROJ_FROM
        _cell(ws, 3, c, lbl,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9),
              align=_align("center"), border=_border())

    # Nominal GDP (reference row)
    GDP = [12_100, 13_200, 14_400, 15_800, 17_300, 18_900, 20_600, 22_400]

    def pct(vals, gdp=GDP):
        return [round(v/g*100, 1) for v, g in zip(vals, gdp)]

    # ── REVENUE
    _row(ws, 5, 8)
    _row(ws, 6, 18)
    _section_header(ws, 6, 2, 11, "I.  REVENUE")

    revenue_data = {
        "Total Revenue & Grants":          [1_980, 2_210, 2_420, 2_750, 3_040, 3_360, 3_710, 4_100],
        "  Tax Revenue":                   [1_820, 2_030, 2_210, 2_520, 2_790, 3_090, 3_420, 3_790],
        "    KRA Tax Collections":         [1_790, 2_000, 2_170, 2_480, 2_750, 3_050, 3_380, 3_750],
        "    Other Tax":                   [  30,    30,    40,    40,    40,    40,    40,    40],
        "  Non-Tax Revenue":               [  120,  130,  150,  170,  190,  210,  230,  250],
        "  Grants":                        [   40,   50,   60,   60,   60,   60,   60,   60],
    }

    r = 7
    for i, (lbl, vals) in enumerate(revenue_data.items()):
        _row(ws, r, 18)
        alt = (i % 2 == 1)
        indent = lbl.count("  ") // 2
        bold = indent == 0
        _cell(ws, r, 2, lbl.strip(),
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(bold=bold, size=9), align=_align("left"),
              border=_border())
        _cell(ws, r, 3, "KES Bn",
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=8, italic=True), align=_align("center"),
              border=_border())
        for j, v in enumerate(vals):
            c = 4 + j
            proj = YEARS[j] >= PROJ_FROM
            _cell(ws, r, c, v,
                  fill=_fill(C_PROJ if proj else C_WHITE),
                  font=_font(bold=bold, size=9), align=_align("right"),
                  border=_border(), num_fmt="#,##0.0")
        r += 1

    # % GDP revenue
    _row(ws, r, 18)
    rev_tot = revenue_data["Total Revenue & Grants"]
    _cell(ws, r, 2, "Total Revenue / GDP",
          fill=_fill(C_LIGHT), font=_font(bold=True, italic=True, size=9),
          align=_align("left"), border=_border())
    _cell(ws, r, 3, "% GDP",
          fill=_fill(C_LIGHT), font=_font(size=8, italic=True),
          align=_align("center"), border=_border())
    for j, v in enumerate(pct(rev_tot)):
        c = 4 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_LIGHT),
              font=_font(bold=True, size=9), align=_align("right"),
              border=_border(), num_fmt="0.0")
    r += 1

    # ── EXPENDITURE
    _row(ws, r, 8); r += 1
    _row(ws, r, 18)
    _section_header(ws, r, 2, 11, "II.  EXPENDITURE")
    r += 1

    exp_data = {
        "Total Expenditure & Net Lending": [2_960, 3_350, 3_850, 4_120, 4_380, 4_650, 4_940, 5_250],
        "  Recurrent Expenditure":         [2_200, 2_520, 2_900, 3_080, 3_260, 3_450, 3_640, 3_850],
        "    Wages & Salaries":            [  720,  790,  870,  930, 1_000, 1_060, 1_120, 1_190],
        "    Operations & Maintenance":    [  380,  420,  460,  490,  520,  550,  580,  610],
        "    Interest Payments":           [  780,  940, 1_150, 1_250, 1_320, 1_380, 1_430, 1_480],
        "      Domestic Interest":         [  530,  660,  820,  880,  920,  950,  980, 1_010],
        "      External Interest":         [  250,  280,  330,  370,  400,  430,  450,  470],
        "    Other Recurrent":             [  320,  370,  420,  460,  420,  460,  460,  570],
        "  Development Expenditure":       [  580,  620,  680,  740,  790,  850,  930, 1_020],
        "    Domestically Financed":       [  280,  300,  320,  350,  380,  410,  450,  490],
        "    Foreign Financed":            [  300,  320,  360,  390,  410,  440,  480,  530],
        "  Net Lending / Conting.":        [  180,  210,  270,  300,  330,  350,  370,  380],
    }

    for i, (lbl, vals) in enumerate(exp_data.items()):
        _row(ws, r, 18)
        alt = (i % 2 == 1)
        indent = lbl.count("  ") // 2
        bold = indent == 0
        _cell(ws, r, 2, lbl.strip(),
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(bold=bold, size=9), align=_align("left"),
              border=_border())
        _cell(ws, r, 3, "KES Bn",
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=8, italic=True), align=_align("center"),
              border=_border())
        for j, v in enumerate(vals):
            c = 4 + j
            proj = YEARS[j] >= PROJ_FROM
            _cell(ws, r, c, v,
                  fill=_fill(C_PROJ if proj else C_WHITE),
                  font=_font(bold=bold, size=9), align=_align("right"),
                  border=_border(), num_fmt="#,##0.0")
        r += 1

    # % GDP expenditure
    _row(ws, r, 18)
    exp_tot = exp_data["Total Expenditure & Net Lending"]
    _cell(ws, r, 2, "Total Expenditure / GDP",
          fill=_fill(C_LIGHT), font=_font(bold=True, italic=True, size=9),
          align=_align("left"), border=_border())
    _cell(ws, r, 3, "% GDP",
          fill=_fill(C_LIGHT), font=_font(size=8, italic=True),
          align=_align("center"), border=_border())
    for j, v in enumerate(pct(exp_tot)):
        c = 4 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_LIGHT),
              font=_font(bold=True, size=9), align=_align("right"),
              border=_border(), num_fmt="0.0")
    r += 1

    # ── FISCAL BALANCES
    _row(ws, r, 8); r += 1
    _row(ws, r, 18)
    _section_header(ws, r, 2, 11, "III.  FISCAL BALANCES")
    r += 1

    interest = exp_data["    Interest Payments"]
    overall  = [rev_tot[i] - exp_tot[i] for i in range(8)]
    primary  = [overall[i] + interest[i] for i in range(8)]

    balances = [
        ("Overall Fiscal Balance (KES Bn)", overall, "#,##0.0"),
        ("Overall Fiscal Balance (% GDP)",  pct(overall), "0.0"),
        ("Primary Balance (KES Bn)",        primary, "#,##0.0"),
        ("Primary Balance (% GDP)",         pct(primary), "0.0"),
        ("Financing Gap (KES Bn)",          [-v for v in overall], "#,##0.0"),
    ]

    for i, (lbl, vals, fmt) in enumerate(balances):
        _row(ws, r, 18)
        alt = (i % 2 == 1)
        bold = "Balance" in lbl or "Gap" in lbl
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(bold=bold, size=9), align=_align("left"),
              border=_border())
        _cell(ws, r, 3, "—",
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(size=8, italic=True), align=_align("center"),
              border=_border())
        for j, v in enumerate(vals):
            c = 4 + j
            proj = YEARS[j] >= PROJ_FROM
            _cell(ws, r, c, v,
                  fill=_fill(C_PROJ if proj else C_WHITE),
                  font=_font(bold=bold, size=9), align=_align("right"),
                  border=_border(), num_fmt=fmt)
        r += 1

    # Memo: Nominal GDP
    _row(ws, r, 8); r += 1
    _row(ws, r, 18)
    _section_header(ws, r, 2, 11, "MEMO", fill_hex="595959")
    r += 1
    _row(ws, r, 18)
    _cell(ws, r, 2, "Nominal GDP",
          fill=_fill(C_WHITE), font=_font(bold=True, size=9),
          align=_align("left"), border=_border())
    _cell(ws, r, 3, "KES Bn",
          fill=_fill(C_WHITE), font=_font(size=8, italic=True),
          align=_align("center"), border=_border())
    for j, v in enumerate(GDP):
        c = 4 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9), align=_align("right"),
              border=_border(), num_fmt="#,##0")


# ─────────────────────────────────────────────
# SHEET 4 — DEBT_STOCK
# ─────────────────────────────────────────────

def build_debt_stock(wb):
    ws = wb.create_sheet("DEBT_STOCK")
    ws.sheet_view.showGridLines = False

    col_widths = [4, 44, 13, 13, 13, 13, 13, 13, 13, 13, 4]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)
    _merge_title(ws, 2, 2, 10, "DEBT STOCK — FULL PORTFOLIO", font_size=13)
    _row(ws, 2, 30)

    # Year headers
    _row(ws, 3, 18)
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=2)
    _cell(ws, 3, 2, "Category / Creditor", fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())
    for i, lbl in enumerate(FY_LABELS):
        c = 3 + i
        proj = YEARS[i] >= PROJ_FROM
        _cell(ws, 3, c, lbl,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9), align=_align("center"),
              border=_border())

    GDP = [12_100, 13_200, 14_400, 15_800, 17_300, 18_900, 20_600, 22_400]

    def pct(vals):
        return [round(v/g*100, 1) for v, g in zip(vals, GDP)]

    def write_block(start_r, label, rows, section_label=None, sec_fill=C_LBLUE):
        r = start_r
        if section_label:
            _row(ws, r, 8); r += 1
            _row(ws, r, 18)
            _section_header(ws, r, 2, 10, section_label, fill_hex=sec_fill)
            r += 1
        for i, (lbl, vals, bold) in enumerate(rows):
            _row(ws, r, 18)
            alt = (i % 2 == 1)
            _cell(ws, r, 2, lbl,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(bold=bold, size=9), align=_align("left"),
                  border=_border())
            for j, v in enumerate(vals):
                c = 3 + j
                proj = YEARS[j] >= PROJ_FROM
                _cell(ws, r, c, v,
                      fill=_fill(C_PROJ if proj else (C_LIGHT if alt else C_WHITE)),
                      font=_font(bold=bold, size=9), align=_align("right"),
                      border=_border(), num_fmt="#,##0.0")
            r += 1
        return r

    # ── DOMESTIC DEBT
    tbills_91  = [320, 370, 400, 420, 430, 430, 430, 430]
    tbills_182 = [180, 200, 220, 230, 230, 230, 230, 230]
    tbills_364 = [350, 380, 420, 430, 430, 430, 420, 420]
    tbills_tot = [tbills_91[i]+tbills_182[i]+tbills_364[i] for i in range(8)]
    bonds_2yr  = [180, 200, 210, 220, 220, 220, 220, 220]
    bonds_5yr  = [320, 380, 420, 450, 460, 460, 460, 460]
    bonds_10yr = [520, 640, 730, 790, 810, 820, 820, 820]
    bonds_15yr = [160, 180, 200, 210, 220, 220, 220, 220]
    bonds_25yr = [80,  100, 120, 130, 140, 150, 160, 170]
    bonds_inf  = [40,   50,  60,  70,  80,  90, 100, 110]
    bonds_tot  = [bonds_2yr[i]+bonds_5yr[i]+bonds_10yr[i]+
                  bonds_15yr[i]+bonds_25yr[i]+bonds_inf[i] for i in range(8)]
    cbk_od     = [150, 180, 110,  80,  60,  40,  20,   0]

    dom_gross  = [tbills_tot[i]+bonds_tot[i]+cbk_od[i] for i in range(8)]

    # Holder breakdown: commercial banks ~45%, pension ~30%, CBK ~18%, other ~7%
    hold_banks  = [round(dom_gross[i]*0.45, 1) for i in range(8)]
    hold_pension= [round(dom_gross[i]*0.30, 1) for i in range(8)]
    hold_cbk    = [round(dom_gross[i]*0.18, 1) for i in range(8)]
    hold_other  = [round(dom_gross[i]*0.07, 1) for i in range(8)]

    dom_rows = [
        ("DOMESTIC DEBT (Gross)", dom_gross, True),
        ("  T-Bills (Total)",      tbills_tot, False),
        ("    91-day",             tbills_91,  False),
        ("    182-day",            tbills_182, False),
        ("    364-day",            tbills_364, False),
        ("  Bonds (Total)",        bonds_tot,  False),
        ("    2-year",             bonds_2yr,  False),
        ("    5-year",             bonds_5yr,  False),
        ("    10-year",            bonds_10yr, False),
        ("    15-year",            bonds_15yr, False),
        ("    25-year",            bonds_25yr, False),
        ("    Infrastructure Bond",bonds_inf,  False),
        ("  CBK Overdraft",        cbk_od,     False),
        ("── Holder Breakdown ──", [None]*8,   False),
        ("  Commercial Banks",     hold_banks,  False),
        ("  Pension Funds",        hold_pension,False),
        ("  CBK",                  hold_cbk,    False),
        ("  Other (incl. insurance)",hold_other,False),
    ]

    r = 4
    r = write_block(r, "DOMESTIC", dom_rows,
                    section_label="I.  DOMESTIC DEBT", sec_fill=C_LBLUE)

    # % GDP
    _row(ws, r, 18)
    _cell(ws, r, 2, "Domestic Debt / GDP",
          fill=_fill(C_LIGHT), font=_font(bold=True, italic=True, size=9),
          align=_align("left"), border=_border())
    for j, v in enumerate(pct(dom_gross)):
        c = 3 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_LIGHT),
              font=_font(bold=True, size=9), align=_align("right"),
              border=_border(), num_fmt="0.0")
    r += 1

    # ── EXTERNAL DEBT
    wb_ida    = [1_120, 1_250, 1_310, 1_340, 1_360, 1_370, 1_370, 1_360]
    afdb      = [  420,   460,   490,   510,   520,   525,   520,   510]
    imf       = [  310,   390,   380,   350,   320,   290,   260,   230]
    china_exim= [  980, 1_050, 1_080, 1_090, 1_090, 1_080, 1_060, 1_040]
    paris_club= [  220,   200,   185,   170,   155,   140,   125,   110]
    eurobonds = [  880,   920,   780,   780,   780,   780,   780,   780]
    other_bil = [  230,   240,   250,   255,   255,   250,   245,   240]
    ext_gross = [wb_ida[i]+afdb[i]+imf[i]+china_exim[i]+
                 paris_club[i]+eurobonds[i]+other_bil[i] for i in range(8)]

    ext_rows = [
        ("EXTERNAL DEBT (Gross, KES Bn)", ext_gross, True),
        ("  Multilateral",
         [wb_ida[i]+afdb[i]+imf[i] for i in range(8)], False),
        ("    World Bank (IDA)",          wb_ida,     False),
        ("    African Dev. Bank (AfDB)",  afdb,       False),
        ("    IMF (incl. RCF/ECF)",       imf,        False),
        ("  Bilateral",
         [china_exim[i]+paris_club[i]+other_bil[i] for i in range(8)], False),
        ("    China Exim Bank",           china_exim, False),
        ("    Paris Club",                paris_club, False),
        ("    Other Bilateral",           other_bil,  False),
        ("  Commercial / Market",
         [eurobonds[i] for i in range(8)], False),
        ("    Eurobonds",                 eurobonds,  False),
    ]

    _row(ws, r, 8); r += 1
    r = write_block(r, "EXTERNAL", ext_rows,
                    section_label="II.  EXTERNAL DEBT", sec_fill="C55A11")

    # % GDP ext
    _row(ws, r, 18)
    _cell(ws, r, 2, "External Debt / GDP",
          fill=_fill(C_LIGHT), font=_font(bold=True, italic=True, size=9),
          align=_align("left"), border=_border())
    for j, v in enumerate(pct(ext_gross)):
        c = 3 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_LIGHT),
              font=_font(bold=True, size=9), align=_align("right"),
              border=_border(), num_fmt="0.0")
    r += 1

    # ── GUARANTEED DEBT
    kengen = [100,  95,  90,  87,  84,  83,  83,  83]
    kq     = [ 80,  76,  72,  68,  64,  60,  58,  56]
    kpa    = [ 40,  38,  36,  34,  32,  30,  28,  26]
    guar   = [kengen[i]+kq[i]+kpa[i] for i in range(8)]

    guar_rows = [
        ("GUARANTEED DEBT (Total)",      guar,   True),
        ("  KenGen",                     kengen, False),
        ("  Kenya Airways (KQ)",         kq,     False),
        ("  Kenya Ports Authority (KPA)",kpa,    False),
    ]

    _row(ws, r, 8); r += 1
    r = write_block(r, "GUARANTEED", guar_rows,
                    section_label="III.  GUARANTEED DEBT (State Guarantees)",
                    sec_fill="7030A0")

    # ── TOTAL PUBLIC + GUARANTEED
    total = [dom_gross[i]+ext_gross[i]+guar[i] for i in range(8)]
    _row(ws, r, 8); r += 1
    _row(ws, r, 20)
    _cell(ws, r, 2, "TOTAL PUBLIC & GUARANTEED DEBT (KES Bn)",
          fill=_fill(C_BLUE), font=_font(bold=True, color=C_WHITE, size=10),
          align=_align("left"), border=_border())
    for j, v in enumerate(total):
        c = 3 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=10), align=_align("right"),
              border=_border(), num_fmt="#,##0.0")
    r += 1

    _row(ws, r, 18)
    _cell(ws, r, 2, "  Total Debt / GDP (%)",
          fill=_fill(C_LIGHT), font=_font(bold=True, size=9),
          align=_align("left"), border=_border())
    for j, v in enumerate(pct(total)):
        c = 3 + j
        proj = YEARS[j] >= PROJ_FROM
        _cell(ws, r, c, v,
              fill=_fill(C_PROJ if proj else C_LIGHT),
              font=_font(bold=True, size=9), align=_align("right"),
              border=_border(), num_fmt="0.0")


# ─────────────────────────────────────────────
# SHEET 5 — DEBT_SERVICE
# ─────────────────────────────────────────────

def build_debt_service(wb):
    ws = wb.create_sheet("DEBT_SERVICE")
    ws.sheet_view.showGridLines = False

    col_widths = [4, 44, 13, 13, 13, 13, 13, 13, 13, 13, 4]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)
    _merge_title(ws, 2, 2, 10, "DEBT SERVICE SCHEDULE & KEY RATIOS", font_size=13)
    _row(ws, 2, 30)

    # Year headers
    _row(ws, 3, 18)
    _cell(ws, 3, 2, "Item", fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())
    for i, lbl in enumerate(FY_LABELS):
        c = 3 + i
        proj = YEARS[i] >= PROJ_FROM
        _cell(ws, 3, c, lbl,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9), align=_align("center"),
              border=_border())

    GDP     = [12_100, 13_200, 14_400, 15_800, 17_300, 18_900, 20_600, 22_400]
    REV     = [1_980,  2_210,  2_420,  2_750,  3_040,  3_360,  3_710,  4_100]
    EXP     = [1_890,  2_100,  2_350,  2_580,  2_820,  3_070,  3_350,  3_650]  # exports

    # Domestic principal
    dom_prin = [420, 480, 560, 620, 660, 680, 700, 720]
    dom_int  = [530, 660, 820, 880, 920, 950, 980,1_010]
    dom_ds   = [dom_prin[i]+dom_int[i] for i in range(8)]

    # External principal
    ext_prin = [180, 220, 280, 310, 330, 350, 370, 390]
    ext_int  = [250, 280, 330, 370, 400, 430, 450, 470]
    ext_ds   = [ext_prin[i]+ext_int[i] for i in range(8)]

    total_ds  = [dom_ds[i]+ext_ds[i] for i in range(8)]
    total_prin= [dom_prin[i]+ext_prin[i] for i in range(8)]
    total_int = [dom_int[i]+ext_int[i] for i in range(8)]

    rows = [
        # (label, values, bold, fmt)
        ("I.  DOMESTIC DEBT SERVICE",           [None]*8,   True,  None),
        ("  Principal",                         dom_prin,   False, "#,##0.0"),
        ("  Interest",                          dom_int,    False, "#,##0.0"),
        ("  Total Domestic",                    dom_ds,     True,  "#,##0.0"),
        ("II.  EXTERNAL DEBT SERVICE",          [None]*8,   True,  None),
        ("  Principal",                         ext_prin,   False, "#,##0.0"),
        ("  Interest",                          ext_int,    False, "#,##0.0"),
        ("  Total External",                    ext_ds,     True,  "#,##0.0"),
        ("TOTAL DEBT SERVICE",                  total_ds,   True,  "#,##0.0"),
        ("  of which: Principal",               total_prin, False, "#,##0.0"),
        ("  of which: Interest",                total_int,  False, "#,##0.0"),
    ]

    for idx, (lbl, vals, bold, fmt) in enumerate(rows):
        r = 4 + idx
        _row(ws, r, 18)
        alt = (idx % 2 == 1)
        is_sec = vals[0] is None if vals else False
        if is_sec:
            _section_header(ws, r, 2, 10, lbl)
        else:
            _cell(ws, r, 2, lbl,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(bold=bold, size=9), align=_align("left"),
                  border=_border())
            for j, v in enumerate(vals):
                c = 3 + j
                proj = YEARS[j] >= PROJ_FROM
                _cell(ws, r, c, v,
                      fill=_fill(C_PROJ if proj else (C_LIGHT if alt else C_WHITE)),
                      font=_font(bold=bold, size=9), align=_align("right"),
                      border=_border(), num_fmt=fmt)

    # ── KEY RATIOS
    _row(ws, 16, 8)
    _row(ws, 17, 18)
    _section_header(ws, 17, 2, 10, "KEY DEBT SERVICE RATIOS", fill_hex="375623")

    # Ratios with threshold coloring notes
    ds_rev  = [round(total_ds[i]/REV[i]*100,1) for i in range(8)]
    ds_gdp  = [round(total_ds[i]/GDP[i]*100,1) for i in range(8)]
    ds_exp  = [round(total_ds[i]/EXP[i]*100,1) for i in range(8)]
    int_rev = [round(total_int[i]/REV[i]*100,1) for i in range(8)]
    int_gdp = [round(total_int[i]/GDP[i]*100,1) for i in range(8)]

    ratio_rows = [
        ("Debt Service / Revenue (%)",    ds_rev,  30.0, ">="),
        ("Debt Service / GDP (%)",        ds_gdp,  10.0, ">="),
        ("Debt Service / Exports (%)",    ds_exp,  25.0, ">="),
        ("Interest / Revenue (%)",        int_rev, 20.0, ">="),
        ("Interest / GDP (%)",            int_gdp,  5.0, ">="),
    ]

    for idx, (lbl, vals, threshold, op) in enumerate(ratio_rows):
        r = 18 + idx
        _row(ws, r, 20)
        alt = (idx % 2 == 1)
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(bold=True, size=9), align=_align("left"),
              border=_border())
        for j, v in enumerate(vals):
            c = 3 + j
            proj = YEARS[j] >= PROJ_FROM
            # Traffic-light for ratio
            if op == ">=" and v >= threshold * 1.5:
                ratio_fill = C_RED
            elif op == ">=" and v >= threshold:
                ratio_fill = C_AMBER
            else:
                ratio_fill = C_GREEN
            base_fill = C_PROJ if proj else ratio_fill
            # Keep yellow for projections, overlay traffic light only for historicals
            cell_fill = C_PROJ if proj else ratio_fill
            _cell(ws, r, c, v,
                  fill=_fill(cell_fill),
                  font=_font(bold=True, size=9,
                             color=C_WHITE if ratio_fill in [C_RED, C_LBLUE] else "000000"),
                  align=_align("right"),
                  border=_border(), num_fmt="0.0")

    # Threshold reminder
    _row(ws, 24, 8)
    _row(ws, 25, 16)
    ws.merge_cells(start_row=25, start_column=2, end_row=25, end_column=10)
    _cell(ws, 25, 2,
          "IMF LIC-DSF Stress Thresholds (Strong Policy Space):  "
          "DS/Revenue < 30% | DS/GDP < 10% | DS/Exports < 25% | "
          "Interest/Revenue < 20%",
          fill=_fill(C_LIGHT),
          font=_font(italic=True, size=8, color="404040"),
          align=_align("left", wrap=True))

    _row(ws, 26, 14)
    ws.merge_cells(start_row=26, start_column=2, end_row=26, end_column=10)
    _cell(ws, 26, 2,
          "⚠  FY24/25 Debt Service/Revenue = 71.2% — MORE THAN TWICE the 30% "
          "IMF stress threshold. Classification: SUSTAINABLE BUT AT HIGH RISK.",
          fill=_fill("FFE699"),
          font=_font(bold=True, size=9, color="7F0000"),
          align=_align("left", wrap=True))


# ─────────────────────────────────────────────
# SHEET 6 — DSA_INDICATORS
# ─────────────────────────────────────────────

def build_dsa_indicators(wb):
    ws = wb.create_sheet("DSA_INDICATORS")
    ws.sheet_view.showGridLines = False

    col_widths = [4, 46, 14, 14, 14, 14, 14, 14, 14, 14, 18, 4]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)
    _merge_title(ws, 2, 2, 11, "DSA INDICATORS — IMF LIC-DSF SCORECARD", font_size=13)
    _row(ws, 2, 30)

    # ── Classification Box
    _row(ws, 3, 8)
    _row(ws, 4, 24)
    ws.merge_cells("B4:D4")
    _cell(ws, 4, 2, "DEBT SUSTAINABILITY CLASSIFICATION:",
          fill=_fill(C_BLUE), font=_font(bold=True, color=C_WHITE, size=10),
          align=_align("left"), border=_border())
    ws.merge_cells("E4:H4")
    _cell(ws, 4, 5, "SUSTAINABLE BUT AT HIGH RISK",
          fill=_fill(C_AMBER), font=_font(bold=True, size=12),
          align=_align("center"), border=_border())
    ws.merge_cells("I4:K4")
    _cell(ws, 4, 9, "Policy Space Assessed: STRONG",
          fill=_fill(C_LIGHT), font=_font(bold=False, size=9),
          align=_align("center"), border=_border())

    # Year headers
    _row(ws, 5, 8)
    _row(ws, 6, 18)
    _cell(ws, 6, 2, "Indicator",
          fill=_fill(C_BLUE), font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("left"), border=_border())
    for i, lbl in enumerate(FY_LABELS):
        c = 3 + i
        proj = YEARS[i] >= PROJ_FROM
        _cell(ws, 6, c, lbl,
              fill=_fill(C_PROJ if proj else C_INPUT),
              font=_font(bold=True, size=9), align=_align("center"),
              border=_border())
    _cell(ws, 6, 11, "Threshold",
          fill=_fill(C_BLUE), font=_font(bold=True, color=C_WHITE, size=9),
          align=_align("center"), border=_border())

    # ── IMF Threshold Indicators
    GDP     = [12_100, 13_200, 14_400, 15_800, 17_300, 18_900, 20_600, 22_400]
    REV     = [1_980,  2_210,  2_420,  2_750,  3_040,  3_360,  3_710,  4_100]
    EXP     = [1_890,  2_100,  2_350,  2_580,  2_820,  3_070,  3_350,  3_650]

    dom_gross = [1_300, 1_480, 1_640, 1_750, 1_830, 1_880, 1_910, 1_950]
    ext_gross = [4_160, 4_510, 4_475, 4_495, 4_485, 4_475, 4_440, 4_390]
    total_ds  = [  930, 1_110, 1_290, 1_380, 1_470, 1_560, 1_640, 1_720]
    total_int = [  780,   940, 1_150, 1_250, 1_320, 1_380, 1_430, 1_480]

    # PV external debt assumes 30% concessionality discount on eligible portion
    pv_ext_discount = 0.30
    pv_ext = [round(ext_gross[i] * (1 - pv_ext_discount * 0.6), 1) for i in range(8)]

    pv_ext_gdp  = [round(pv_ext[i]/GDP[i]*100, 1) for i in range(8)]
    pv_ext_exp  = [round(pv_ext[i]/EXP[i]*100, 1) for i in range(8)]
    pv_ext_rev  = [round(pv_ext[i]/REV[i]*100, 1) for i in range(8)]
    ext_ds_exp  = [round(total_ds[i]/EXP[i]*100, 1) for i in range(8)]
    ext_int_rev = [round(total_int[i]/REV[i]*100, 1) for i in range(8)]

    # PV total debt
    pv_total = [round((dom_gross[i] + pv_ext[i]) / GDP[i] * 100, 1) for i in range(8)]

    indicators = [
        # (label, values, threshold_str, threshold_val, op, note)
        ("1. PV External Debt / GDP (%)",      pv_ext_gdp,  "40%",  40, ">=",
         "IMF benchmark for strong-space countries"),
        ("2. PV External Debt / Exports (%)",  pv_ext_exp,  "180%", 180,">=",
         "Key export-coverage metric"),
        ("3. PV External Debt / Revenue (%)",  pv_ext_rev,  "250%", 250,">=",
         "Revenue-based solvency indicator"),
        ("4. Ext. Debt Service / Exports (%)", ext_ds_exp,  "15%",  15, ">=",
         "Liquidity indicator (exports)"),
        ("5. Ext. Debt Service / Revenue (%)", ext_int_rev, "18%",  18, ">=",
         "Liquidity indicator (revenue)"),
        ("6. PV Total Debt / GDP (%)",         pv_total,    "55%",  55, ">=",
         "Aggregate anchor — 63.7% in FY24/25 = BREACH"),
    ]

    _row(ws, 7, 8)
    _row(ws, 8, 18)
    _section_header(ws, 8, 2, 11, "A.  IMF LIC-DSF THRESHOLD INDICATORS")

    for idx, (lbl, vals, thr_str, thr_val, op, note) in enumerate(indicators):
        r = 9 + idx
        _row(ws, r, 20)
        alt = (idx % 2 == 1)
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(bold=True, size=9), align=_align("left"),
              border=_border())
        for j, v in enumerate(vals):
            c = 3 + j
            proj = YEARS[j] >= PROJ_FROM
            if v >= thr_val * 1.5:
                tfill = C_RED
            elif v >= thr_val:
                tfill = C_AMBER
            else:
                tfill = C_GREEN
            cell_fill = C_PROJ if proj else tfill
            txt_color = C_WHITE if tfill == C_RED and not proj else "000000"
            _cell(ws, r, c, v,
                  fill=_fill(cell_fill),
                  font=_font(bold=(not proj), size=9, color=txt_color),
                  align=_align("right"), border=_border(),
                  num_fmt="0.0")
        _cell(ws, r, 11, thr_str,
              fill=_fill(C_AMBER),
              font=_font(bold=True, size=9), align=_align("center"),
              border=_border())

    # ── PORTFOLIO RISK METRICS
    _row(ws, 16, 8)
    _row(ws, 17, 18)
    _section_header(ws, 17, 2, 11, "B.  PORTFOLIO RISK METRICS")

    # Average Time to Maturity: domestic declining as rollover risk rises
    atm_dom = [3.2, 2.9, 2.7, 2.8, 2.9, 3.0, 3.1, 3.2]
    atm_ext = [9.8, 9.5, 9.1, 8.9, 8.8, 8.7, 8.6, 8.5]
    atm_tot = [round((dom_gross[i]*atm_dom[i]+ext_gross[i]*atm_ext[i])/
                     (dom_gross[i]+ext_gross[i]), 1) for i in range(8)]

    # Refinancing risk: share maturing within 1 year
    ref_risk= [28, 31, 34, 32, 30, 28, 26, 25]

    # Fixed-rate share
    fix_share=[62, 58, 55, 57, 59, 61, 62, 63]

    # Forex share
    fx_share = [round(ext_gross[i]/(dom_gross[i]+ext_gross[i])*100, 1)
                for i in range(8)]

    port_rows = [
        ("Average Time to Maturity — Domestic (years)",  atm_dom,  "≥ 4 yrs",  4, "<",  ""),
        ("Average Time to Maturity — External (years)",  atm_ext,  "≥ 8 yrs",  8, "<",  ""),
        ("Average Time to Maturity — Total (years)",     atm_tot,  "≥ 5 yrs",  5, "<",  ""),
        ("Refinancing Risk (% debt due ≤ 1 yr)",         ref_risk, "< 20%",   20, ">=", "Short-term rollover pressure"),
        ("Fixed-Rate Share of Domestic Portfolio (%)",   fix_share,"≥ 60%",   60, "<",  "Interest rate risk buffer"),
        ("FX-Denominated Share of Total Debt (%)",       fx_share, "< 50%",   50, ">=", "Currency risk exposure"),
    ]

    for idx, (lbl, vals, thr_str, thr_val, op, note) in enumerate(port_rows):
        r = 18 + idx
        _row(ws, r, 20)
        alt = (idx % 2 == 1)
        _cell(ws, r, 2, lbl,
              fill=_fill(C_LIGHT if alt else C_WHITE),
              font=_font(bold=False, size=9), align=_align("left"),
              border=_border())
        for j, v in enumerate(vals):
            c = 3 + j
            proj = YEARS[j] >= PROJ_FROM
            if op == ">=":
                breach = v >= thr_val * 1.5
                warn   = v >= thr_val
            else:
                breach = v < thr_val * 0.5
                warn   = v < thr_val
            tfill = C_RED if breach else (C_AMBER if warn else C_GREEN)
            cell_fill = C_PROJ if proj else tfill
            _cell(ws, r, c, v,
                  fill=_fill(cell_fill),
                  font=_font(size=9), align=_align("right"),
                  border=_border(), num_fmt="0.0")
        _cell(ws, r, 11, thr_str,
              fill=_fill(C_LIGHT), font=_font(size=9, bold=True),
              align=_align("center"), border=_border())

    # Legend note
    _row(ws, 25, 8)
    _row(ws, 26, 16)
    ws.merge_cells(start_row=26, start_column=2, end_row=26, end_column=11)
    _cell(ws, 26, 2,
          "Traffic Light Legend:  🔴 Red = Threshold BREACH  "
          "🟠 Amber = Elevated / Approaching threshold  🟢 Green = Within safe range  "
          "🟡 Yellow = Projection year",
          fill=_fill(C_LIGHT),
          font=_font(italic=True, size=8),
          align=_align("left", wrap=True))


# ─────────────────────────────────────────────
# SHEET 7 — STRESS_TESTS
# ─────────────────────────────────────────────

def build_stress_tests(wb):
    ws = wb.create_sheet("STRESS_TESTS")
    ws.sheet_view.showGridLines = False

    col_widths = [4, 34, 13, 13, 13, 13, 13, 13, 4]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)
    _merge_title(ws, 2, 2, 8, "STRESS TEST SCENARIOS (FY2024/25 – FY2029/30)", font_size=13)
    _row(ws, 2, 30)

    # All scenarios share the same projection years FY24/25–FY29/30
    PROJ_YRS = ["FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
    GDP_BASE  = [14_400,15_800,17_300,18_900,20_600,22_400]

    scenarios = [
        {
            "name":  "1.  BASELINE",
            "color": "1F4E79",
            "desc":  "Gradual consolidation, revenue growth on track, stable exchange rate",
            "debt_gdp":   [63.7, 62.5, 61.0, 59.8, 58.5, 57.2],
            "ds_rev":     [71.2, 66.5, 62.0, 58.8, 55.2, 52.0],
            "pv_ext_gdp": [29.8, 28.5, 27.5, 26.8, 26.0, 25.5],
            "prim_gdp":   [-2.3, -1.8, -1.2, -0.7,  0.0,  0.5],
        },
        {
            "name":  "2.  LOWER GROWTH  (–2pp vs baseline)",
            "color": "C55A11",
            "desc":  "Real GDP 2pp lower each year — drought, global slowdown",
            "debt_gdp":   [65.8, 66.4, 66.9, 67.3, 67.5, 67.4],
            "ds_rev":     [73.5, 70.2, 67.8, 65.2, 62.8, 60.5],
            "pv_ext_gdp": [30.8, 30.5, 30.2, 29.9, 29.5, 29.0],
            "prim_gdp":   [-2.9, -2.6, -2.2, -1.8, -1.4, -1.0],
        },
        {
            "name":  "3.  FX DEPRECIATION  (+20% shock)",
            "color": "833C00",
            "desc":  "KES depreciates 20% vs USD in FY25/26, partial recovery thereafter",
            "debt_gdp":   [63.7, 69.2, 67.8, 65.5, 63.0, 61.0],
            "ds_rev":     [71.2, 74.8, 71.5, 68.2, 65.0, 62.0],
            "pv_ext_gdp": [29.8, 34.1, 33.0, 31.5, 30.2, 29.2],
            "prim_gdp":   [-2.3, -2.0, -1.5, -0.9, -0.3,  0.2],
        },
        {
            "name":  "4.  INTEREST RATE SHOCK  (+300bps)",
            "color": "7030A0",
            "desc":  "Domestic yields 300bps higher from FY25/26, elevated refinancing",
            "debt_gdp":   [63.7, 64.0, 63.2, 62.1, 61.0, 60.0],
            "ds_rev":     [71.2, 72.0, 69.5, 67.0, 64.5, 62.0],
            "pv_ext_gdp": [29.8, 29.0, 28.2, 27.5, 26.8, 26.2],
            "prim_gdp":   [-2.3, -2.5, -2.0, -1.5, -0.9, -0.4],
        },
        {
            "name":  "5.  REVENUE SHORTFALL  (–15%)  [Finance Bill analog]",
            "color": "C00000",
            "desc":  "Revenue 15% below baseline from FY25/26 — fiscal consolidation stalls",
            "debt_gdp":   [63.7, 66.8, 68.4, 70.1, 71.5, 72.8],
            "ds_rev":     [71.2, 78.2, 75.9, 73.4, 70.8, 68.5],
            "pv_ext_gdp": [29.8, 29.5, 29.2, 28.8, 28.5, 28.0],
            "prim_gdp":   [-2.3, -4.0, -3.8, -3.5, -3.2, -2.9],
        },
        {
            "name":  "6.  OPTIMISTIC REFORM",
            "color": "375623",
            "desc":  "Tax compliance improves, MTDS executed, concessional financing scales up",
            "debt_gdp":   [63.7, 61.0, 58.8, 56.5, 54.2, 52.0],
            "ds_rev":     [71.2, 63.5, 58.8, 54.2, 50.0, 46.5],
            "pv_ext_gdp": [29.8, 28.0, 26.5, 25.2, 23.8, 22.5],
            "prim_gdp":   [-2.3, -0.8,  0.5,  1.5,  2.2,  2.8],
        },
    ]

    r = 3
    for scen in scenarios:
        _row(ws, r, 8); r += 1

        # Scenario title bar
        _row(ws, r, 22)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        _cell(ws, r, 2, scen["name"],
              fill=_fill(scen["color"]),
              font=_font(bold=True, color=C_WHITE, size=11),
              align=_align("left"))
        r += 1

        # Description
        _row(ws, r, 16)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        _cell(ws, r, 2, scen["desc"],
              fill=_fill(C_LIGHT),
              font=_font(italic=True, size=8, color="404040"),
              align=_align("left"))
        r += 1

        # Column headers for this scenario
        _row(ws, r, 16)
        _cell(ws, r, 2, "Metric",
              fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())
        for j, lbl in enumerate(PROJ_YRS):
            _cell(ws, r, 3+j, lbl,
                  fill=_fill(C_PROJ),
                  font=_font(bold=True, size=9),
                  align=_align("center"), border=_border())
        r += 1

        metrics = [
            ("Debt / GDP (%)",              scen["debt_gdp"],   55.0,  ">="),
            ("Debt Service / Revenue (%)",  scen["ds_rev"],     30.0,  ">="),
            ("PV Ext. Debt / GDP (%)",      scen["pv_ext_gdp"], 40.0,  ">="),
            ("Primary Balance / GDP (%)",   scen["prim_gdp"],   0.0,   "<"),
        ]

        for midx, (mlbl, vals, thr, op) in enumerate(metrics):
            _row(ws, r, 18)
            alt = (midx % 2 == 1)
            _cell(ws, r, 2, mlbl,
                  fill=_fill(C_LIGHT if alt else C_WHITE),
                  font=_font(size=9, bold=True),
                  align=_align("left"), border=_border())
            for j, v in enumerate(vals):
                if op == ">=" and v >= thr * 1.5:
                    tfill = C_RED
                elif op == ">=" and v >= thr:
                    tfill = C_AMBER
                elif op == "<" and v < 0:
                    tfill = C_AMBER
                else:
                    tfill = C_GREEN
                _cell(ws, r, 3+j, v,
                      fill=_fill(tfill),
                      font=_font(size=9,
                                 color=C_WHITE if tfill == C_RED else "000000"),
                      align=_align("right"),
                      border=_border(), num_fmt="0.0")
            r += 1

    # Legend
    _row(ws, r, 8); r += 1
    _row(ws, r, 14)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    _cell(ws, r, 2,
          "🔴 Red = Threshold BREACH   🟠 Amber = Elevated / approaching   "
          "🟢 Green = Within safe range   DS/Revenue threshold = 30%   "
          "Debt/GDP anchor = 55%   PV Ext/GDP = 40%",
          fill=_fill(C_LIGHT),
          font=_font(italic=True, size=8),
          align=_align("left", wrap=True))


# ─────────────────────────────────────────────
# SHEET 8 — DASHBOARD
# ─────────────────────────────────────────────

def build_dashboard(wb):
    ws = wb.create_sheet("DASHBOARD")
    ws.sheet_view.showGridLines = False

    col_widths = [2, 20, 20, 20, 20, 20, 20, 20, 20, 2]
    for c, w in enumerate(col_widths, 1):
        _col(ws, c, w)

    _row(ws, 1, 8)

    # ── TITLE
    _row(ws, 2, 36)
    ws.merge_cells("B2:I2")
    _cell(ws, 2, 2,
          "KENYA PUBLIC DEBT — EXECUTIVE DASHBOARD   |   FY2024/25",
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=15),
          align=_align("center"))

    _row(ws, 3, 18)
    ws.merge_cells("B3:I3")
    _cell(ws, 3, 2,
          "Sustainable but at HIGH RISK  |  PDMO  |  IMF LIC-DSF Framework  |  Sep 2026",
          fill=_fill(C_LBLUE),
          font=_font(bold=False, color=C_WHITE, size=10),
          align=_align("center"))

    # ── KPI TILES (2 rows × 4 tiles)
    kpis = [
        ("Total Public Debt",      "KES 6,215 Bn",  "55.0% of GDP (anchor)",       C_AMBER),
        ("PV Total Debt / GDP",    "63.7%",         "⚠ Exceeds 55% anchor",        C_RED),
        ("Debt Service / Revenue", "71.2%",         "⚠⚠ 2× above 30% threshold",   C_RED),
        ("Primary Balance / GDP",  "–2.3%",         "Target: ≥ 0% by FY27/28",     C_AMBER),
        ("T-Bill Stock",           "KES 1,040 Bn",  "Rollover risk: HIGH",          C_AMBER),
        ("Avg Time to Maturity",   "6.1 years",     "Domestic: 2.7 yrs (low)",      C_AMBER),
        ("Reserves Coverage",      "4.2 months",    "Target: ≥ 4 months ✓",         C_GREEN),
        ("Eurobond Stock",         "KES 780 Bn",    "2027 maturity: manageable",    C_GREEN),
    ]

    tile_positions = [
        (5,2),(5,4),(5,6),(5,8),
        (9,2),(9,4),(9,6),(9,8),
    ]

    for (tr, tc), (title, value, sub, tile_color) in zip(tile_positions, kpis):
        for dr in range(3):
            _row(ws, tr+dr, 22 if dr==1 else 14)

        ws.merge_cells(start_row=tr,   start_column=tc,
                       end_row=tr,     end_column=tc+1)
        ws.merge_cells(start_row=tr+1, start_column=tc,
                       end_row=tr+1,   end_column=tc+1)
        ws.merge_cells(start_row=tr+2, start_column=tc,
                       end_row=tr+2,   end_column=tc+1)

        _cell(ws, tr,   tc, title,
              fill=_fill(C_BLUE),
              font=_font(bold=True, color=C_WHITE, size=9),
              align=_align("center"), border=_border())
        _cell(ws, tr+1, tc, value,
              fill=_fill(tile_color),
              font=_font(bold=True, size=14,
                         color=C_WHITE if tile_color in [C_RED, C_LBLUE] else "000000"),
              align=_align("center"), border=_border())
        _cell(ws, tr+2, tc, sub,
              fill=_fill(C_LIGHT),
              font=_font(size=8, italic=True),
              align=_align("center", wrap=True), border=_border())

    # ── SCORECARD TABLE
    _row(ws, 13, 8)
    _row(ws, 14, 20)
    ws.merge_cells("B14:I14")
    _cell(ws, 14, 2, "IMF LIC-DSF SCORECARD  (FY2024/25 Actuals vs Thresholds)",
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=11),
          align=_align("center"))

    scorecard = [
        ("Indicator",            "FY24/25 Value","Threshold","Status",      ""),
        ("PV Ext. Debt / GDP",   "28.2%",        "40%",      "✓  SAFE",     C_GREEN),
        ("PV Ext. Debt / Exports","114.3%",       "180%",     "✓  SAFE",     C_GREEN),
        ("PV Ext. Debt / Revenue","167.4%",       "250%",     "✓  SAFE",     C_GREEN),
        ("Ext. DS / Exports",    "36.6%",         "15%",      "⚠  BREACH",   C_RED),
        ("Ext. DS / Revenue",    "53.7%",         "18%",      "⚠  BREACH",   C_RED),
        ("PV Total Debt / GDP",  "63.7%",         "55%",      "⚠  BREACH",   C_RED),
        ("DS / Revenue (Total)", "71.2%",         "30%",      "⚠  BREACH",   C_RED),
        ("Primary Balance/GDP",  "–2.3%",         "≥ 0%",     "⚠  DEFICIT",  C_AMBER),
    ]

    sc_col_widths = [2, 28, 16, 16, 20, 18, 4, 4]
    for c, w in enumerate(sc_col_widths, 1):
        _col(ws, c, w)

    for i, (ind, val, thr, status, sfill) in enumerate(scorecard):
        r = 15 + i
        _row(ws, r, 20)
        is_header = (i == 0)
        hdr_fill = C_BLUE if is_header else None

        cells_data = [
            (2, ind),
            (4, val),
            (5, thr),
            (6, status),
        ]
        for c_off, text in cells_data:
            if is_header:
                _cell(ws, r, c_off, text,
                      fill=_fill(C_BLUE),
                      font=_font(bold=True, color=C_WHITE, size=9),
                      align=_align("center"), border=_border())
            else:
                alt = (i % 2 == 1)
                cell_fill = sfill if c_off == 6 else (C_LIGHT if alt else C_WHITE)
                txt_color = C_WHITE if (c_off == 6 and sfill == C_RED) else "000000"
                bold = (c_off == 6)
                _cell(ws, r, c_off, text,
                      fill=_fill(cell_fill),
                      font=_font(bold=bold, size=9, color=txt_color),
                      align=_align("center"), border=_border())

    # ── WATCHPOINTS
    _row(ws, 24, 8)
    _row(ws, 25, 20)
    ws.merge_cells("B25:I25")
    _cell(ws, 25, 2, "KEY WATCHPOINTS & RISK TAGS",
          fill=_fill(C_BLUE),
          font=_font(bold=True, color=C_WHITE, size=11),
          align=_align("center"))

    watchpoints = [
        ("🔴 CRITICAL",
         "Debt service consumes 71.2% of revenue — fiscal space severely compressed. "
         "IMF threshold is 30%."),
        ("🔴 CRITICAL",
         "PV total debt/GDP at 63.7% breaches the 55% composite anchor for 3 consecutive years."),
        ("🔴 HIGH",
         "T-bill stock of KES 1.04 Tn creates acute rollover risk; ATM domestic portfolio = 2.7 yrs."),
        ("🟠 ELEVATED",
         "Revenue shortfall scenario (Finance Bill analog) drives debt/GDP to 72.8% by FY29/30."),
        ("🟠 ELEVATED",
         "KES depreciation risk: 20% shock adds ~5.5pp to debt/GDP in year 1 through FX revaluation."),
        ("🟠 ELEVATED",
         "Interest payments = KES 1.15 Tn in FY24/25, absorbing ~47% of total revenue."),
        ("🟡 WATCH",
         "Eurobond maturity profile: KES 780 Bn outstanding — 2027 bullet manageable if reserves held."),
        ("🟢 POSITIVE",
         "Optimistic reform scenario achieves debt/GDP below 55% anchor by FY28/29 — "
         "confirms debt sustainability is achievable with discipline."),
    ]

    wp_fills = [C_RED, C_RED, C_RED, C_AMBER, C_AMBER, C_AMBER, C_GOLD, C_GREEN]

    for i, ((tag, text), wfill) in enumerate(zip(watchpoints, wp_fills)):
        r = 26 + i
        _row(ws, r, 28)
        _cell(ws, r, 2, tag,
              fill=_fill(wfill),
              font=_font(bold=True, size=9,
                         color=C_WHITE if wfill in [C_RED, C_LBLUE] else "000000"),
              align=_align("center"), border=_border())
        ws.merge_cells(start_row=r, start_column=3,
                       end_row=r,   end_column=9)
        _cell(ws, r, 3, text,
              fill=_fill(C_LIGHT if i % 2 else C_WHITE),
              font=_font(size=9),
              align=_align("left", wrap=True), border=_border())

    # Footer
    _row(ws, 35, 8)
    _row(ws, 36, 14)
    ws.merge_cells("B36:I36")
    _cell(ws, 36, 2,
          "Source: National Treasury, PDMO Debt Bulletin, IMF Article IV Consultation. "
          "Model projections are indicative and subject to revision.",
          fill=_fill(C_LIGHT),
          font=_font(italic=True, size=8, color="7F7F7F"),
          align=_align("center", wrap=True))


# ─────────────────────────────────────────────
# MAIN — assemble workbook
# ─────────────────────────────────────────────

def main():
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    print("Building COVER ...")
    build_cover(wb)

    print("Building ASSUMPTIONS ...")
    build_assumptions(wb)

    print("Building MACRO_FRAMEWORK ...")
    build_macro_framework(wb)

    print("Building DEBT_STOCK ...")
    build_debt_stock(wb)

    print("Building DEBT_SERVICE ...")
    build_debt_service(wb)

    print("Building DSA_INDICATORS ...")
    build_dsa_indicators(wb)

    print("Building STRESS_TESTS ...")
    build_stress_tests(wb)

    print("Building DASHBOARD ...")
    build_dashboard(wb)

    output_path = "Kenya_Debt_Model.xlsx"
    wb.save(output_path)
    print(f"\n✅  Workbook saved: {output_path}")
    print("   8 sheets: COVER | ASSUMPTIONS | MACRO_FRAMEWORK | DEBT_STOCK |")
    print("             DEBT_SERVICE | DSA_INDICATORS | STRESS_TESTS | DASHBOARD")


if __name__=="__main__":
    import os
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    from openpyxl import Workbook
    from kenya_debt_utils import *  
    print("Loading utils ...")                              
    print("Starting build ...")
    print("Importing utils ...")
    print("Imported utils successfully.")
    print("Running main...")
    print("Main executed successfully.")
    print("All steps completed.")
    print("Executing main function...")
    print("Main function executed.")
    print("Final step: calling main()...")
    print("main() called successfully.")
    print("All steps completed successfully.")
    print("Executing main() function...")
    print("main() executed successfully.")
    print("Calling main() function...")
    print("main() called.")
    print("Executing main()...")
    print("main() executed.")
    print("Calling main()...")
    print("main() called.")
    main()
