"""
sensitivity.py
==============
Generates Kenya_Sensitivity.xlsx with two sensitivity analysis tables:

  Table 1: Debt/GDP in FY2029/30 across a grid of
            Real GDP Growth (x-axis) vs USD/KES Exchange Rate shock (y-axis)

  Table 2: Debt Service/Revenue in FY2024/25 across a grid of
            Revenue Growth Rate vs Interest Rate Level

Traffic-light coloring on every cell:
  Green  = below threshold
  Amber  = at or above threshold
  Red    = severely above threshold (1.5×)

Run:
    python sensitivity.py

Requires:
    pip install openpyxl
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT = "Kenya_Sensitivity.xlsx"

# ── Colours
C_BLUE  = "1F4E79"
C_LBLUE = "2E75B6"
C_RED   = "C00000"
C_AMBER = "FFC000"
C_GREEN = "70AD47"
C_LIGHT = "F2F2F2"
C_WHITE = "FFFFFF"
C_GOLD  = "FFD966"

def _fill(h):  return PatternFill("solid", fgColor=h)
def _font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")
def _align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def _border():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)
def _col(ws, c, w):
    ws.column_dimensions[get_column_letter(c)].width = w
def _row(ws, r, h):
    ws.row_dimensions[r].height = h

def traffic_light(value, threshold, severe_multiple=1.5, reverse=False):
    """Return fill hex based on value vs threshold.
    reverse=True means LOWER is bad (e.g. primary balance, reserves).
    """
    if not reverse:
        if value >= threshold * severe_multiple:
            return C_RED, C_WHITE
        elif value >= threshold:
            return C_AMBER, "000000"
        else:
            return C_GREEN, C_WHITE
    else:
        if value <= threshold * (2 - severe_multiple):
            return C_RED, C_WHITE
        elif value <= threshold:
            return C_AMBER, "000000"
        else:
            return C_GREEN, C_WHITE


# ═══════════════════════════════════════════════════
# MODEL FUNCTIONS
# ═══════════════════════════════════════════════════

def compute_debt_gdp_2030(base_growth_pp, fx_shock_pct):
    """
    Simplified model: project debt/GDP to FY2029/30 given
    - base_growth_pp: real GDP growth adjustment (pp vs 5.8% baseline, e.g. -2 means 3.8%)
    - fx_shock_pct:   permanent % depreciation applied from FY25/26 (e.g. 20 = 20% weaker KES)

    Returns debt/GDP % in FY2029/30.
    """
    # Baseline FY24/25 starting point
    debt_kes     = 6_215.0   # KES Bn total debt
    gdp_kes      = 14_400.0  # KES Bn nominal GDP FY24/25
    ext_debt_usd = 4_475.0 / 128  # KES → USD at base rate (KES 128/USD)
    dom_debt     = 1_640.0   # KES Bn domestic
    guar_debt    = 202.0

    base_fx      = 128.0   # KES/USD baseline
    fx_rate      = base_fx * (1 + fx_shock_pct / 100)  # shocked rate

    base_nom_growth = 0.108   # ~10.8% nominal growth baseline
    adj_growth  = base_nom_growth - (base_growth_pp / 100) * (14_400 / 12_100)
    adj_growth  = max(adj_growth, 0.02)   # floor at 2%

    # Primary deficit converges: -2.3% → 0% linearly over 5 years
    primary_deficits = [-0.023, -0.015, -0.008, -0.003, 0.0]
    # Interest rate on domestic: proxy 13%
    dom_int_rate = 0.13 + max(0, -base_growth_pp * 0.005)

    gdp = gdp_kes
    dom = dom_debt
    ext_usd = ext_debt_usd
    guar = guar_debt

    for yr in range(5):   # FY25/26 through FY29/30
        gdp_new = gdp * (1 + adj_growth)
        # Domestic debt grows by net borrowing (primary deficit share) + interest
        dom_new = dom * (1 + dom_int_rate * 0.1)  # simplified net new issuance
        # External: principal repayments reduce stock ~3% p.a., convert at shocked rate
        ext_usd_new = ext_usd * 0.97
        ext_kes_new = ext_usd_new * fx_rate
        guar_new = guar * 0.97
        # Add primary deficit financing
        prim_def = primary_deficits[yr] * gdp
        total_new = dom_new + ext_kes_new + guar_new + abs(prim_def)
        # Update
        gdp = gdp_new
        dom = dom_new
        ext_usd = ext_usd_new
        guar = guar_new

    debt_2030 = dom + (ext_usd * fx_rate) + guar
    return round(debt_2030 / gdp * 100, 1)


def compute_ds_rev(revenue_growth_adj_pp, int_rate_adj_pp):
    """
    Simplified: DS/Revenue in FY24/25 given
    - revenue_growth_adj_pp: adjustment to revenue growth (pp vs baseline)
    - int_rate_adj_pp: domestic interest rate adjustment (bp as pp)

    Returns DS/Revenue %.
    """
    base_revenue   = 2_420.0   # KES Bn
    base_ds        = 1_290.0   # KES Bn
    base_dom_int   = 820.0     # KES Bn
    dom_debt       = 1_640.0

    # Revenue adjustment
    rev_adj = revenue_growth_adj_pp / 100
    revenue = base_revenue * (1 + rev_adj)

    # Interest adjustment: +1pp on 1.64T domestic portfolio
    extra_int = dom_debt * (int_rate_adj_pp / 100)
    ds = base_ds + extra_int

    return round(ds / revenue * 100, 1)


def compute_pv_ext_gdp(growth_adj_pp, fx_shock_pct):
    """PV External Debt / GDP at end of projection."""
    base_pv_ext_gdp = 28.2
    # FX shock increases numerator, lower growth increases ratio
    shock_effect = fx_shock_pct * 0.25 + (-growth_adj_pp) * 0.8
    return round(base_pv_ext_gdp + shock_effect, 1)


# ═══════════════════════════════════════════════════
# BUILD WORKBOOK
# ═══════════════════════════════════════════════════

def build_sensitivity():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Debt_GDP_Sensitivity"
    ws2 = wb.create_sheet("DS_Revenue_Sensitivity")
    ws3 = wb.create_sheet("PV_Ext_Sensitivity")

    # ─────────────────────────────────────────────────
    # SHEET 1: Debt/GDP 2030 — Growth × FX Shock
    # ─────────────────────────────────────────────────
    build_table(
        ws1,
        title="SENSITIVITY ANALYSIS — Debt/GDP in FY2029/30 (%)",
        subtitle="Rows = Real GDP Growth Shock (pp vs 5.8% baseline)   "
                 "|   Columns = USD/KES Depreciation Shock (%)",
        row_label="Growth Shock (pp)",
        col_label="FX Depreciation Shock (%)",
        row_values=[-2.0, -1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5, 2.0],
        col_values=[0, 5, 10, 15, 20, 25, 30, 35, 40],
        compute_fn=compute_debt_gdp_2030,
        threshold=55.0,
        severe_mult=65/55,
        note="Threshold = 55% (IMF composite anchor). "
             "Green < 55%  |  Amber 55–65%  |  Red ≥ 65%",
        row_fmt=lambda v: f"{v:+.1f}pp",
        col_fmt=lambda v: f"+{v:.0f}%",
    )

    # ─────────────────────────────────────────────────
    # SHEET 2: DS/Revenue — Revenue Growth × Interest Rate
    # ─────────────────────────────────────────────────
    build_table(
        ws2,
        title="SENSITIVITY ANALYSIS — Debt Service / Revenue in FY2024/25 (%)",
        subtitle="Rows = Revenue Growth Adjustment (pp vs baseline)   "
                 "|   Columns = Domestic Interest Rate Shock (pp)",
        row_label="Revenue Growth Adj (pp)",
        col_label="Interest Rate Shock (pp)",
        row_values=[-5.0, -4.0, -3.0, -2.0, -1.0, 0.0, 1.0, 2.0, 3.0],
        col_values=[-3.0, -2.0, -1.0, 0.0, 1.0, 2.0, 3.0, 4.0, 5.0],
        compute_fn=compute_ds_rev,
        threshold=30.0,
        severe_mult=1.5,
        note="Threshold = 30% (IMF DS/Revenue stress limit). "
             "Green < 30%  |  Amber 30–45%  |  Red ≥ 45%",
        row_fmt=lambda v: f"{v:+.1f}pp",
        col_fmt=lambda v: f"{v:+.1f}pp",
    )

    # ─────────────────────────────────────────────────
    # SHEET 3: PV Ext Debt/GDP — Growth × FX
    # ─────────────────────────────────────────────────
    build_table(
        ws3,
        title="SENSITIVITY ANALYSIS — PV External Debt/GDP (%)",
        subtitle="Rows = Real GDP Growth Shock (pp)   "
                 "|   Columns = USD/KES Depreciation Shock (%)",
        row_label="Growth Shock (pp)",
        col_label="FX Depreciation Shock (%)",
        row_values=[-2.0, -1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5, 2.0],
        col_values=[0, 5, 10, 15, 20, 25, 30, 35, 40],
        compute_fn=compute_pv_ext_gdp,
        threshold=40.0,
        severe_mult=1.5,
        note="Threshold = 40% (IMF PV External Debt/GDP benchmark). "
             "Green < 40%  |  Amber 40–60%  |  Red ≥ 60%",
        row_fmt=lambda v: f"{v:+.1f}pp",
        col_fmt=lambda v: f"+{v:.0f}%",
    )

    wb.save(OUTPUT)
    print(f"✅  Sensitivity tables saved: {OUTPUT}")
    print("   3 sheets: Debt_GDP_Sensitivity | DS_Revenue_Sensitivity | PV_Ext_Sensitivity")


def build_table(ws, title, subtitle, row_label, col_label,
                row_values, col_values, compute_fn,
                threshold, severe_mult, note,
                row_fmt, col_fmt):
    ws.sheet_view.showGridLines = False

    # Column widths
    _col(ws, 1, 3)
    _col(ws, 2, 26)
    for c in range(3, 3 + len(col_values)):
        _col(ws, c, 14)

    # Title
    _row(ws, 1, 8)
    _row(ws, 2, 28)
    ws.merge_cells(start_row=2, start_column=2,
                   end_row=2, end_column=2 + len(col_values))
    c = ws.cell(2, 2, title)
    c.fill = _fill(C_BLUE); c.font = _font(True, C_WHITE, 12)
    c.alignment = _align("center")

    _row(ws, 3, 18)
    ws.merge_cells(start_row=3, start_column=2,
                   end_row=3, end_column=2 + len(col_values))
    c = ws.cell(3, 2, subtitle)
    c.fill = _fill(C_LIGHT); c.font = _font(False, "404040", 9, italic=True)
    c.alignment = _align("center")

    # Corner / axis labels
    _row(ws, 4, 18)
    c = ws.cell(4, 2)
    c.value = f"{row_label}  ↓   {col_label}  →"
    c.fill = _fill(C_BLUE); c.font = _font(True, C_WHITE, 9)
    c.alignment = _align("center")
    c.border = _border()

    # Column headers
    for j, cv in enumerate(col_values):
        col = 3 + j
        _row(ws, 4, 18)
        cell = ws.cell(4, col, col_fmt(cv))
        cell.fill = _fill("BDD7EE")
        cell.font = _font(True, "000000", 9)
        cell.alignment = _align("center")
        cell.border = _border()

    # Data grid
    for i, rv in enumerate(row_values):
        r = 5 + i
        _row(ws, r, 20)

        # Row label
        cell = ws.cell(r, 2, row_fmt(rv))
        cell.fill = _fill("BDD7EE")
        cell.font = _font(True, "000000", 9)
        cell.alignment = _align("center")
        cell.border = _border()

        for j, cv in enumerate(col_values):
            col = 3 + j
            val = compute_fn(rv, cv)
            fill_hex, font_color = traffic_light(val, threshold, severe_mult)
            cell = ws.cell(r, col, val)
            cell.fill   = _fill(fill_hex)
            cell.font   = _font(True, font_color, 10)
            cell.alignment = _align("center")
            cell.border = _border()
            cell.number_format = "0.0"

    # Baseline highlight box
    base_r_idx = row_values.index(0.0) if 0.0 in row_values else None
    base_c_idx = col_values.index(0)   if 0   in col_values else None
    if base_r_idx is not None and base_c_idx is not None:
        br = 5 + base_r_idx
        bc = 3 + base_c_idx
        cell = ws.cell(br, bc)
        # Add bold border around baseline cell
        thick = Side(border_style="medium", color="000000")
        thin  = Side(border_style="thin",   color="BFBFBF")
        cell.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        # Add "BASE" annotation
        ws.cell(br, bc).comment = None

    # Note row
    note_row = 5 + len(row_values) + 1
    _row(ws, note_row, 16)
    ws.merge_cells(start_row=note_row, start_column=2,
                   end_row=note_row,   end_column=2 + len(col_values))
    c = ws.cell(note_row, 2, note)
    c.fill = _fill(C_LIGHT)
    c.font = _font(False, "404040", 8, italic=True)
    c.alignment = _align("left")

    # Legend
    leg_row = note_row + 1
    _row(ws, leg_row, 16)
    legend = [
        (2, "■ GREEN", C_GREEN, C_WHITE),
        (3, "■ AMBER", C_AMBER, "000000"),
        (4, "■ RED",   C_RED,   C_WHITE),
        (5, "■ [BASE] = Baseline scenario", C_LIGHT, "000000"),
    ]
    for col, text, bg, fc in legend:
        cell = ws.cell(leg_row, col, text)
        cell.fill = _fill(bg)
        cell.font = _font(True, fc, 8)
        cell.alignment = _align("center")
        cell.border = _border()


if __name__ == "__main__":
    build_sensitivity()
