"""
Adds a CHARTS sheet to Kenya_Debt_Model.xlsx with 4 embedded charts:
  1. Debt/GDP Trend (line) — baseline vs 55% anchor
  2. Scenario Fan Chart (line) — all 6 stress scenarios
  3. Debt Service / Revenue (bar) — with 30% threshold line
  4. Debt Composition (stacked bar) — domestic / external / guaranteed

Run AFTER kenya_debt_model.py:
    python charts_excel.py
"""

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabel
from openpyxl.utils import get_column_letter
import openpyxl.chart as xlchart

WORKBOOK = "Kenya_Debt_Model.xlsx"

C_BLUE  = "1F4E79"
C_WHITE = "FFFFFF"
C_LIGHT = "F2F2F2"
C_AMBER = "FFC000"
C_RED   = "C00000"
C_GREEN = "70AD47"

FY = ["FY22/23","FY23/24","FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
PROJ_FY = ["FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]

def _fill(h): return PatternFill("solid", fgColor=h)
def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")
def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def _border():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)
def _col(ws, c, w):
    ws.column_dimensions[get_column_letter(c)].width = w
def _row(ws, r, h):
    ws.row_dimensions[r].height = h

def write_data_table(ws, start_row, start_col, headers, rows, title):
    """Write a compact data table used as chart source."""
    r = start_row
    # Title
    ws.merge_cells(start_row=r, start_column=start_col,
                   end_row=r, end_column=start_col + len(headers))
    c = ws.cell(r, start_col, title)
    c.fill = _fill(C_BLUE); c.font = _font(True, C_WHITE, 9)
    c.alignment = _align("center")
    r += 1

    # Header row
    for i, h in enumerate([""] + headers):
        c = ws.cell(r, start_col + i, h)
        c.fill = _fill("BDD7EE"); c.font = _font(True, size=9)
        c.alignment = _align("center"); c.border = _border()
    r += 1

    # Data rows
    for row_label, values in rows:
        ws.cell(r, start_col, row_label).value = row_label
        ws.cell(r, start_col).font = _font(size=9)
        ws.cell(r, start_col).border = _border()
        for i, v in enumerate(values):
            cell = ws.cell(r, start_col + 1 + i, v)
            cell.font = _font(size=9)
            cell.alignment = _align("right")
            cell.border = _border()
            cell.number_format = "0.0"
        r += 1
    return r  # next free row

def build_charts(wb):
    # Remove existing CHARTS sheet if present
    if "CHARTS" in wb.sheetnames:
        del wb["CHARTS"]

    ws = wb.create_sheet("CHARTS", 7)  # insert before DASHBOARD
    ws.sheet_view.showGridLines = False

    # Column widths
    for c in range(1, 30):
        _col(ws, c, 12 if c > 1 else 3)

    # Title
    _row(ws, 1, 8)
    _row(ws, 2, 30)
    ws.merge_cells("B2:AC2")
    c = ws.cell(2, 2, "KENYA DSA MODEL — CHARTS & VISUALISATIONS")
    c.fill = _fill(C_BLUE); c.font = _font(True, C_WHITE, 14)
    c.alignment = _align("center")

    # ─────────────────────────────────────────
    # DATA TABLES (hidden source data for charts)
    # ─────────────────────────────────────────

    # 1. Debt/GDP Trend
    debt_gdp_base    = [52.3, 57.1, 63.7, 62.5, 61.0, 59.8, 58.5, 57.2]
    anchor_line      = [55.0] * 8

    # 2. Scenario fan (projection years only FY24/25–FY29/30)
    scen_baseline    = [63.7, 62.5, 61.0, 59.8, 58.5, 57.2]
    scen_low_growth  = [63.7, 66.4, 66.9, 67.3, 67.5, 67.4]
    scen_fx          = [63.7, 69.2, 67.8, 65.5, 63.0, 61.0]
    scen_ir          = [63.7, 64.0, 63.2, 62.1, 61.0, 60.0]
    scen_rev         = [63.7, 66.8, 68.4, 70.1, 71.5, 72.8]
    scen_optimistic  = [63.7, 61.0, 58.8, 56.5, 54.2, 52.0]

    # 3. DS/Revenue
    ds_rev = [47.0, 50.2, 71.2, 66.5, 62.0, 58.8, 55.2, 52.0]
    threshold_30 = [30.0] * 8

    # 4. Debt composition (KES Bn)
    dom_debt  = [1_300, 1_480, 1_640, 1_750, 1_830, 1_880, 1_910, 1_950]
    ext_debt  = [4_160, 4_510, 4_475, 4_495, 4_485, 4_475, 4_440, 4_390]
    guar_debt = [  220,   209,   202,   189,   180,   173,   169,   165]

    # Write source tables starting at row 4, col 2
    r = 4
    r = write_data_table(ws, r, 2, FY,
        [("Debt/GDP Baseline (%)", debt_gdp_base),
         ("55% Anchor",           anchor_line)],
        "Chart 1 Data — Debt/GDP vs Anchor")
    r += 1
    data1_start = 5   # row where first data row of table 1 begins (after header)
    data1_end   = 6

    r = write_data_table(ws, r, 2, PROJ_FY,
        [("Baseline",        scen_baseline),
         ("Lower Growth",    scen_low_growth),
         ("FX Depreciation", scen_fx),
         ("Interest Shock",  scen_ir),
         ("Revenue Shortfall",scen_rev),
         ("Optimistic",      scen_optimistic)],
        "Chart 2 Data — Scenario Fan (Debt/GDP %)")
    r += 1

    r = write_data_table(ws, r, 2, FY,
        [("DS/Revenue (%)", ds_rev),
         ("30% Threshold",  threshold_30)],
        "Chart 3 Data — Debt Service / Revenue")
    r += 1

    r = write_data_table(ws, r, 2, FY,
        [("Domestic Debt",   dom_debt),
         ("External Debt",   ext_debt),
         ("Guaranteed Debt", guar_debt)],
        "Chart 4 Data — Debt Composition (KES Bn)")

    # ─────────────────────────────────────────
    # CHART 1 — Debt/GDP Trend Line
    # ─────────────────────────────────────────
    chart1 = LineChart()
    chart1.title = "Debt / GDP — Baseline vs 55% Anchor"
    chart1.style = 10
    chart1.y_axis.title = "% of GDP"
    chart1.x_axis.title = "Fiscal Year"
    chart1.y_axis.scaling.min = 40
    chart1.y_axis.scaling.max = 75
    chart1.width  = 18
    chart1.height = 12

    # Baseline series
    data_ref1 = Reference(ws, min_col=3, max_col=10, min_row=5, max_row=5)
    cats1     = Reference(ws, min_col=3, max_col=10, min_row=4)
    chart1.add_data(data_ref1, titles_from_data=False)
    chart1.set_categories(cats1)
    chart1.series[0].title = SeriesLabel(v="Debt/GDP Baseline")
    chart1.series[0].graphicalProperties.line.solidFill = C_BLUE
    chart1.series[0].graphicalProperties.line.width = 25000
    chart1.series[0].marker.symbol = "circle"
    chart1.series[0].marker.size   = 6

    # Anchor line
    data_ref1b = Reference(ws, min_col=3, max_col=10, min_row=6, max_row=6)
    chart1.add_data(data_ref1b, titles_from_data=False)
    chart1.series[1].title = SeriesLabel(v="55% Anchor")
    chart1.series[1].graphicalProperties.line.solidFill = C_RED
    chart1.series[1].graphicalProperties.line.width = 18000
    chart1.series[1].graphicalProperties.line.dashDot = "dash"

    ws.add_chart(chart1, "B36")

    # ─────────────────────────────────────────
    # CHART 2 — Scenario Fan
    # ─────────────────────────────────────────
    chart2 = LineChart()
    chart2.title = "Debt/GDP — Scenario Fan (FY24/25–FY29/30)"
    chart2.style = 10
    chart2.y_axis.title = "% of GDP"
    chart2.x_axis.title = "Fiscal Year"
    chart2.y_axis.scaling.min = 45
    chart2.y_axis.scaling.max = 80
    chart2.width  = 18
    chart2.height = 12

    fan_colors = [C_BLUE, "C55A11", "833C00", "7030A0", C_RED, C_GREEN]
    fan_labels = ["Baseline","Lower Growth","FX Depreciation",
                  "Interest Shock","Revenue Shortfall","Optimistic"]

    # Table 2 starts at the row after r_table2_header
    # Dynamically find it: table 2 header row = 4 + 4 (table1 rows) + 1 gap + 1 header = row 10
    t2_header_row = 11   # header row of table 2
    t2_data_start = 12

    for i in range(6):
        ref = Reference(ws, min_col=3, max_col=8,
                        min_row=t2_data_start + i,
                        max_row=t2_data_start + i)
        chart2.add_data(ref, titles_from_data=False)
        chart2.series[i].title = SeriesLabel(v=fan_labels[i])
        chart2.series[i].graphicalProperties.line.solidFill = fan_colors[i]
        chart2.series[i].graphicalProperties.line.width = 20000
        chart2.series[i].marker.symbol = "diamond"
        chart2.series[i].marker.size   = 5

    cats2 = Reference(ws, min_col=3, max_col=8, min_row=t2_header_row)
    chart2.set_categories(cats2)

    ws.add_chart(chart2, "N36")

    # ─────────────────────────────────────────
    # CHART 3 — DS/Revenue Bar + Threshold
    # ─────────────────────────────────────────
    chart3 = BarChart()
    chart3.type = "col"
    chart3.title = "Debt Service / Revenue (%) vs 30% Threshold"
    chart3.style = 10
    chart3.y_axis.title = "% of Revenue"
    chart3.x_axis.title = "Fiscal Year"
    chart3.y_axis.scaling.min = 0
    chart3.y_axis.scaling.max = 90
    chart3.width  = 18
    chart3.height = 12

    # Find table 3 — it's after table 2 (6 data rows) + gap
    t3_header_row = t2_data_start + 6 + 2   # = 20
    t3_data_start = t3_header_row + 1        # = 21

    ds_ref = Reference(ws, min_col=3, max_col=10,
                       min_row=t3_data_start, max_row=t3_data_start)
    chart3.add_data(ds_ref, titles_from_data=False)
    chart3.series[0].title = SeriesLabel(v="DS/Revenue %")
    chart3.series[0].graphicalProperties.solidFill = C_AMBER
    chart3.series[0].graphicalProperties.line.solidFill = "000000"

    thr_ref = Reference(ws, min_col=3, max_col=10,
                        min_row=t3_data_start + 1, max_row=t3_data_start + 1)
    chart3.add_data(thr_ref, titles_from_data=False)
    chart3.series[1].title = SeriesLabel(v="30% Threshold")
    chart3.series[1].graphicalProperties.solidFill = C_RED

    cats3 = Reference(ws, min_col=3, max_col=10, min_row=t3_header_row)
    chart3.set_categories(cats3)

    ws.add_chart(chart3, "B57")

    # ─────────────────────────────────────────
    # CHART 4 — Stacked Bar Debt Composition
    # ─────────────────────────────────────────
    chart4 = BarChart()
    chart4.type    = "col"
    chart4.grouping = "stacked"
    chart4.title   = "Debt Composition by Category (KES Billions)"
    chart4.style   = 10
    chart4.y_axis.title = "KES Billions"
    chart4.x_axis.title = "Fiscal Year"
    chart4.width   = 18
    chart4.height  = 12

    t4_header_row = t3_data_start + 2 + 2   # = 25
    t4_data_start = t4_header_row + 1        # = 26

    comp_colors = [C_BLUE, "C55A11", "7030A0"]
    comp_labels = ["Domestic Debt", "External Debt", "Guaranteed Debt"]

    for i in range(3):
        ref = Reference(ws, min_col=3, max_col=10,
                        min_row=t4_data_start + i,
                        max_row=t4_data_start + i)
        chart4.add_data(ref, titles_from_data=False)
        chart4.series[i].title = SeriesLabel(v=comp_labels[i])
        chart4.series[i].graphicalProperties.solidFill = comp_colors[i]

    cats4 = Reference(ws, min_col=3, max_col=10, min_row=t4_header_row)
    chart4.set_categories(cats4)

    ws.add_chart(chart4, "N57")

    # Chart labels
    chart_notes = [
        ("B35", "Chart 1 — Debt/GDP Trend"),
        ("N35", "Chart 2 — Scenario Fan"),
        ("B56", "Chart 3 — Debt Service / Revenue"),
        ("N56", "Chart 4 — Debt Composition"),
    ]
    for addr, label in chart_notes:
        c = ws[addr]
        c.value = label
        c.font  = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
        c.fill  = _fill(C_BLUE)
        c.alignment = _align("left")

def main():
    print("Loading workbook ...")
    wb = load_workbook(WORKBOOK)
    print("Adding CHARTS sheet with 4 charts ...")
    build_charts(wb)
    wb.save(WORKBOOK)
    print(f"✅  Charts added and saved to {WORKBOOK}")

if __name__ == "__main__":
    main()
