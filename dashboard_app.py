# -*- coding: utf-8 -*-
"""
dashboard_app.py
================
Interactive web dashboard for the Kenya Debt Sustainability Analysis Model.
Runs a local Dash app at http://127.0.0.1:8050

Features:
  - 8 live KPI tiles
  - Debt/GDP trend chart with 55% anchor
  - Scenario fan chart (6 stress tests)
  - Debt Service/Revenue bar chart with 30% threshold
  - Debt composition stacked bar
  - DSA scorecard table with traffic-light badges
  - Sensitivity heatmap (growth × FX)
  - All charts interactive (hover, zoom, download)

Run:
    python dashboard_app.py
Then open: http://127.0.0.1:8050

Requires:
    pip install dash plotly pandas
"""

import dash
from dash import dcc, html, dash_table
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd

# ═══════════════════════════════════════════════════
# DATA
# ═══════════════════════════════════════════════════

FY       = ["FY22/23","FY23/24","FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]
PROJ_FY  = ["FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30"]

DEBT_GDP  = [52.3, 57.1, 63.7, 62.5, 61.0, 59.8, 58.5, 57.2]
DS_REV    = [47.0, 50.2, 71.2, 66.5, 62.0, 58.8, 55.2, 52.0]
DS_GDP    = [7.7,  8.4,  9.0,  8.7,  8.5,  8.3,  8.0,  7.7]
INT_REV   = [39.4, 42.5, 47.5, 45.5, 43.4, 41.1, 38.5, 36.1]
PRIMARY   = [-6.2, -8.6, -2.3, -1.8, -1.2, -0.7,  0.0,  0.5]

DOM_DEBT  = [1_300,1_480,1_640,1_750,1_830,1_880,1_910,1_950]
EXT_DEBT  = [4_160,4_510,4_475,4_495,4_485,4_475,4_440,4_390]
GUAR_DEBT = [  220,  209,  202,  189,  180,  173,  169,  165]
TOTAL_DEBT= [d+e+g for d,e,g in zip(DOM_DEBT,EXT_DEBT,GUAR_DEBT)]

SCENARIOS = {
    "Baseline":           [63.7,62.5,61.0,59.8,58.5,57.2],
    "Lower Growth (–2pp)":[63.7,66.4,66.9,67.3,67.5,67.4],
    "FX Depreciation +20%":[63.7,69.2,67.8,65.5,63.0,61.0],
    "Interest Shock +300bps":[63.7,64.0,63.2,62.1,61.0,60.0],
    "Revenue Shortfall –15%":[63.7,66.8,68.4,70.1,71.5,72.8],
    "Optimistic Reform":  [63.7,61.0,58.8,56.5,54.2,52.0],
}

SCEN_COLORS = ["#1F4E79","#C55A11","#833C00","#7030A0","#C00000","#70AD47"]

SCORECARD = [
    {"Indicator":"PV Ext. Debt / GDP",     "Value":"28.2%","Threshold":"40%", "Status":"✓ SAFE",   "Color":"#70AD47"},
    {"Indicator":"PV Ext. Debt / Exports", "Value":"114.3%","Threshold":"180%","Status":"✓ SAFE",  "Color":"#70AD47"},
    {"Indicator":"PV Ext. Debt / Revenue", "Value":"167.4%","Threshold":"250%","Status":"✓ SAFE",  "Color":"#70AD47"},
    {"Indicator":"Ext. DS / Exports",      "Value":"36.6%","Threshold":"15%", "Status":"⚠ BREACH", "Color":"#C00000"},
    {"Indicator":"Ext. DS / Revenue",      "Value":"53.7%","Threshold":"18%", "Status":"⚠ BREACH", "Color":"#C00000"},
    {"Indicator":"PV Total Debt / GDP",    "Value":"63.7%","Threshold":"55%", "Status":"⚠ BREACH", "Color":"#C00000"},
    {"Indicator":"DS / Revenue (Total)",   "Value":"71.2%","Threshold":"30%", "Status":"⚠ BREACH", "Color":"#C00000"},
    {"Indicator":"Primary Balance / GDP",  "Value":"–2.3%","Threshold":"≥ 0%","Status":"⚠ DEFICIT","Color":"#FFC000"},
]

# Sensitivity grid: Debt/GDP 2030 — growth shock × FX shock
GROWTH_SHOCKS = [-2.0,-1.5,-1.0,-0.5, 0.0, 0.5, 1.0, 1.5, 2.0]
FX_SHOCKS     = [  0,   5,  10,  15,  20,  25,  30,  35,  40]

def _debt_gdp_2030(g_shock, fx_shock):
    base = 57.2
    return round(base - g_shock * 1.8 + fx_shock * 0.22, 1)

SENS_GRID = [[_debt_gdp_2030(g, f) for f in FX_SHOCKS] for g in GROWTH_SHOCKS]

# ═══════════════════════════════════════════════════
# CHART BUILDERS
# ═══════════════════════════════════════════════════

NAVY   = "#1F4E79"
BLUE   = "#2E75B6"
RED    = "#C00000"
AMBER  = "#FFC000"
GREEN  = "#70AD47"
LGREY  = "#F2F2F2"
DGREY  = "#595959"

def chart_debt_gdp():
    fig = go.Figure()
    # Historical
    fig.add_trace(go.Scatter(
        x=FY[:3], y=DEBT_GDP[:3],
        mode="lines+markers",
        name="Debt/GDP (Historical)",
        line=dict(color=NAVY, width=3),
        marker=dict(size=8, color=NAVY),
        hovertemplate="%{x}<br>Debt/GDP: %{y:.1f}%<extra></extra>"
    ))
    # Projection
    fig.add_trace(go.Scatter(
        x=FY[2:], y=DEBT_GDP[2:],
        mode="lines+markers",
        name="Debt/GDP (Projection)",
        line=dict(color=BLUE, width=2.5, dash="dot"),
        marker=dict(size=7, color=BLUE, symbol="diamond"),
        hovertemplate="%{x}<br>Debt/GDP: %{y:.1f}%<extra></extra>"
    ))
    # 55% anchor
    fig.add_hline(y=55, line_dash="dash", line_color=RED, line_width=2,
                  annotation_text="55% IMF Anchor",
                  annotation_position="bottom right",
                  annotation_font_color=RED)
    # Breach zone
    fig.add_hrect(y0=55, y1=80, fillcolor=RED, opacity=0.05, layer="below",
                  line_width=0)

    fig.update_layout(
        title=dict(text="Debt / GDP — Baseline vs 55% IMF Anchor", font=dict(size=15, color=NAVY)),
        yaxis=dict(title="% of GDP", range=[40,75], ticksuffix="%",
                   gridcolor="#E5E5E5"),
        xaxis=dict(title="Fiscal Year", gridcolor="#E5E5E5"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
        margin=dict(l=50,r=20,t=60,b=40)
    )
    return fig


def chart_scenario_fan():
    fig = go.Figure()
    for i, (name, vals) in enumerate(SCENARIOS.items()):
        dash_style = "solid" if i == 0 else ("dot" if i == 5 else "dash")
        width = 3.5 if i == 0 else 2
        fig.add_trace(go.Scatter(
            x=PROJ_FY, y=vals,
            mode="lines+markers",
            name=name,
            line=dict(color=SCEN_COLORS[i], width=width, dash=dash_style),
            marker=dict(size=7),
            hovertemplate=f"{name}<br>%{{x}}: %{{y:.1f}}%<extra></extra>"
        ))
    fig.add_hline(y=55, line_dash="dash", line_color=RED, line_width=1.5,
                  annotation_text="55% Anchor", annotation_position="bottom right",
                  annotation_font_color=RED)
    fig.update_layout(
        title=dict(text="Scenario Fan — Debt/GDP Through FY2029/30",
                   font=dict(size=15, color=NAVY)),
        yaxis=dict(title="% of GDP", range=[45,80], ticksuffix="%",
                   gridcolor="#E5E5E5"),
        xaxis=dict(title="Fiscal Year", gridcolor="#E5E5E5"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=9)),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
        margin=dict(l=50,r=20,t=70,b=40)
    )
    return fig


def chart_ds_revenue():
    colors_bar = []
    for v in DS_REV:
        if v >= 45: colors_bar.append(RED)
        elif v >= 30: colors_bar.append(AMBER)
        else: colors_bar.append(GREEN)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=FY, y=DS_REV,
        name="DS / Revenue",
        marker_color=colors_bar,
        marker_line_color=DGREY,
        marker_line_width=0.5,
        hovertemplate="%{x}<br>DS/Revenue: %{y:.1f}%<extra></extra>"
    ))
    fig.add_hline(y=30, line_dash="dash", line_color=RED, line_width=2,
                  annotation_text="30% Threshold",
                  annotation_position="top left",
                  annotation_font_color=RED)
    fig.add_annotation(
        x="FY24/25", y=71.2,
        text="71.2%<br>⚠ 2× threshold",
        showarrow=True, arrowhead=2, arrowcolor=RED,
        font=dict(color=RED, size=10, family="Arial Black"),
        ay=-40, ax=0
    )
    fig.update_layout(
        title=dict(text="Debt Service / Revenue (%) vs 30% IMF Threshold",
                   font=dict(size=15, color=NAVY)),
        yaxis=dict(title="% of Revenue", range=[0,90], ticksuffix="%",
                   gridcolor="#E5E5E5"),
        xaxis=dict(title="Fiscal Year", gridcolor="#E5E5E5"),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=False,
        margin=dict(l=50,r=20,t=60,b=40)
    )
    return fig


def chart_debt_composition():
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=FY, y=DOM_DEBT, name="Domestic Debt",
        marker_color=NAVY,
        hovertemplate="%{x}<br>Domestic: KES %{y:,.0f} Bn<extra></extra>"
    ))
    fig.add_trace(go.Bar(
        x=FY, y=EXT_DEBT, name="External Debt",
        marker_color="#C55A11",
        hovertemplate="%{x}<br>External: KES %{y:,.0f} Bn<extra></extra>"
    ))
    fig.add_trace(go.Bar(
        x=FY, y=GUAR_DEBT, name="Guaranteed Debt",
        marker_color="#7030A0",
        hovertemplate="%{x}<br>Guaranteed: KES %{y:,.0f} Bn<extra></extra>"
    ))
    fig.update_layout(
        barmode="stack",
        title=dict(text="Debt Composition by Category (KES Billions)",
                   font=dict(size=15, color=NAVY)),
        yaxis=dict(title="KES Billions", gridcolor="#E5E5E5",
                   tickformat=",.0f"),
        xaxis=dict(title="Fiscal Year", gridcolor="#E5E5E5"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=50,r=20,t=70,b=40)
    )
    return fig


def chart_sensitivity_heatmap():
    z_text = [[f"{v:.1f}%" for v in row] for row in SENS_GRID]
    fig = go.Figure(go.Heatmap(
        z=SENS_GRID,
        x=[f"+{v:.0f}%" for v in FX_SHOCKS],
        y=[f"{v:+.1f}pp" for v in GROWTH_SHOCKS],
        text=z_text,
        texttemplate="%{text}",
        textfont=dict(size=10, color="white"),
        colorscale=[
            [0.0,  GREEN],
            [0.45, GREEN],
            [0.46, AMBER],
            [0.65, AMBER],
            [0.66, RED],
            [1.0,  RED],
        ],
        zmin=48, zmax=80,
        colorbar=dict(
            title="Debt/GDP %",
            ticksuffix="%",
            thickness=15,
            len=0.8
        ),
        hovertemplate="Growth: %{y}<br>FX Shock: %{x}<br>Debt/GDP: %{z:.1f}%<extra></extra>"
    ))
    # Anchor line
    fig.add_shape(
        type="line", x0=-0.5, x1=len(FX_SHOCKS)-0.5,
        y0=GROWTH_SHOCKS.index(0.0), y1=GROWTH_SHOCKS.index(0.0),
        line=dict(color="black", width=2, dash="dot")
    )
    fig.update_layout(
        title=dict(text="Sensitivity: Debt/GDP in FY2029/30 — Growth Shock × FX Shock",
                   font=dict(size=15, color=NAVY)),
        xaxis_title="USD/KES Depreciation Shock",
        yaxis_title="Real GDP Growth Shock (pp vs baseline)",
        margin=dict(l=90, r=20, t=60, b=50),
        paper_bgcolor="white"
    )
    return fig


def chart_primary_balance():
    bar_colors = [GREEN if v >= 0 else RED for v in PRIMARY]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=FY, y=PRIMARY,
        marker_color=bar_colors,
        marker_line_color=DGREY,
        marker_line_width=0.5,
        hovertemplate="%{x}<br>Primary Balance: %{y:.1f}% of GDP<extra></extra>"
    ))
    fig.add_hline(y=0, line_color=NAVY, line_width=1.5)
    fig.update_layout(
        title=dict(text="Primary Balance (% of GDP)",
                   font=dict(size=15, color=NAVY)),
        yaxis=dict(title="% of GDP", gridcolor="#E5E5E5", ticksuffix="%"),
        xaxis=dict(title="Fiscal Year", gridcolor="#E5E5E5"),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=False,
        margin=dict(l=50, r=20, t=60, b=40)
    )
    return fig


# ═══════════════════════════════════════════════════
# KPI TILES
# ═══════════════════════════════════════════════════

def kpi_tile(label, value, sub, color):
    return html.Div([
        html.Div(label, style={
            "background": NAVY, "color": "white",
            "padding": "6px 10px", "fontSize": "11px",
            "fontWeight": "bold", "textAlign": "center",
            "borderRadius": "6px 6px 0 0"
        }),
        html.Div(value, style={
            "background": color, "color": "white" if color in [RED,"#C00000"] else "#000",
            "padding": "10px", "fontSize": "22px",
            "fontWeight": "bold", "textAlign": "center",
        }),
        html.Div(sub, style={
            "background": LGREY, "color": DGREY,
            "padding": "5px 8px", "fontSize": "10px",
            "textAlign": "center", "borderRadius": "0 0 6px 6px",
            "fontStyle": "italic"
        }),
    ], style={
        "border": f"1.5px solid {color}",
        "borderRadius": "6px",
        "flex": "1",
        "minWidth": "160px",
        "margin": "6px",
        "boxShadow": "2px 2px 6px rgba(0,0,0,0.12)"
    })


KPIS = [
    ("Total Public Debt",      "KES 6,215 Bn", "55.0% of GDP",               AMBER),
    ("PV Total Debt / GDP",    "63.7%",        "⚠ Exceeds 55% anchor",       RED),
    ("Debt Service / Revenue", "71.2%",        "⚠⚠ 2× IMF threshold (30%)",  RED),
    ("Primary Balance / GDP",  "–2.3%",        "Target: ≥ 0% by FY27/28",   AMBER),
    ("T-Bill Stock",           "KES 1,040 Bn", "Rollover risk: HIGH",         AMBER),
    ("Avg Time to Maturity",   "6.1 years",    "Domestic ATM: 2.7 yrs",       AMBER),
    ("Gross Reserves",         "4.2 months",   "Target ≥ 4 months ✓",         GREEN),
    ("Eurobond Stock",         "KES 780 Bn",   "2027 maturity: manageable",   GREEN),
]


# ═══════════════════════════════════════════════════
# SCORECARD TABLE
# ═══════════════════════════════════════════════════

def scorecard_table():
    df = pd.DataFrame(SCORECARD)
    return dash_table.DataTable(
        data=df.to_dict("records"),
        columns=[
            {"name": "Indicator",  "id": "Indicator"},
            {"name": "FY24/25",    "id": "Value"},
            {"name": "Threshold",  "id": "Threshold"},
            {"name": "Status",     "id": "Status"},
        ],
        style_table={"overflowX": "auto", "borderRadius": "6px",
                     "border": "1px solid #BFBFBF"},
        style_header={
            "backgroundColor": NAVY, "color": "white",
            "fontWeight": "bold", "fontSize": "12px",
            "textAlign": "center", "padding": "10px"
        },
        style_cell={
            "fontFamily": "Calibri, Arial, sans-serif",
            "fontSize": "12px", "padding": "8px 12px",
            "textAlign": "center", "border": "1px solid #E5E5E5"
        },
        style_cell_conditional=[
            {"if": {"column_id": "Indicator"},
             "textAlign": "left", "fontWeight": "bold"}
        ],
        style_data_conditional=[
            {"if": {"filter_query": '{Status} contains "BREACH" || {Status} contains "DEFICIT"'},
             "backgroundColor": "#FFF0F0"},
            {"if": {"filter_query": '{Color} = "#C00000"', "column_id": "Status"},
             "backgroundColor": RED, "color": "white", "fontWeight": "bold"},
            {"if": {"filter_query": '{Color} = "#70AD47"', "column_id": "Status"},
             "backgroundColor": GREEN, "color": "white", "fontWeight": "bold"},
            {"if": {"filter_query": '{Color} = "#FFC000"', "column_id": "Status"},
             "backgroundColor": AMBER, "fontWeight": "bold"},
            {"if": {"row_index": "odd"}, "backgroundColor": LGREY},
        ],
        page_size=10,
        sort_action="native",
    )


# ═══════════════════════════════════════════════════
# WATCHPOINTS HELPER
# ═══════════════════════════════════════════════════

WATCHPOINTS_DATA = [
    ("🔴 CRITICAL", RED,   "Debt service = 71.2% of revenue — fiscal space severely compressed. IMF threshold: 30%."),
    ("🔴 CRITICAL", RED,   "PV total debt/GDP at 63.7% breaches the 55% composite anchor for 3 consecutive years."),
    ("🔴 HIGH",     RED,   "T-bill stock of KES 1.04 Tn — rollover risk acute; domestic ATM only 2.7 years."),
    ("🟠 ELEVATED", AMBER, "Revenue shortfall scenario drives debt/GDP to 72.8% by FY2029/30."),
    ("🟠 ELEVATED", AMBER, "KES depreciation risk: 20% shock adds ~5.5pp to debt/GDP immediately."),
    ("🟠 ELEVATED", AMBER, "Interest payments = KES 1.15 Tn, absorbing 47% of total revenue in FY24/25."),
    ("🟡 WATCH",    "#FFD966", "Eurobond maturity: KES 780 Bn outstanding — 2027 bullet manageable if reserves held."),
    ("🟢 POSITIVE", GREEN, "Optimistic reform scenario achieves debt/GDP below 55% anchor by FY2028/29."),
]

def watchpoint_rows():
    rows = []
    for i, (tag, bg, text) in enumerate(WATCHPOINTS_DATA):
        tag_color = "white" if bg in [RED, GREEN] else "#7F0000"
        rows.append(
            html.Div([
                html.Span(tag, style={
                    "fontWeight": "bold",
                    "color": tag_color,
                    "background": bg,
                    "padding": "3px 10px",
                    "borderRadius": "4px",
                    "marginRight": "10px",
                    "fontSize": "11px",
                    "whiteSpace": "nowrap"
                }),
                html.Span(text, style={"fontSize": "12px", "color": DGREY})
            ], style={
                "display": "flex",
                "alignItems": "center",
                "padding": "7px 10px",
                "background": LGREY if i % 2 == 0 else "white",
                "borderBottom": "1px solid #E5E5E5"
            })
        )
    return rows


# ═══════════════════════════════════════════════════
# LAYOUT
# ═══════════════════════════════════════════════════

app = dash.Dash(
    __name__,
    title="Kenya DSA Dashboard",
    meta_tags=[{"name": "viewport",
                "content": "width=device-width, initial-scale=1"}]
)

app.layout = html.Div([

    # ── Header
    html.Div([
        html.H1("🇰🇪  Kenya Public Debt Sustainability Dashboard",
                style={"color": "white", "margin": "0", "fontSize": "24px",
                       "fontWeight": "bold"}),
        html.P("FY2022/23 – FY2029/30  |  IMF LIC-DSF Framework  |  PDMO  |  Sep 2026",
               style={"color": "#BDD7EE", "margin": "4px 0 0", "fontSize": "12px"}),
    ], style={
        "background": f"linear-gradient(135deg, {NAVY} 0%, {BLUE} 100%)",
        "padding": "18px 28px",
        "borderBottom": f"4px solid {AMBER}",
        "boxShadow": "0 2px 8px rgba(0,0,0,0.3)"
    }),

    # ── Classification banner
    html.Div([
        html.Span("⚠  CLASSIFICATION: ", style={"fontWeight": "bold"}),
        html.Span("SUSTAINABLE BUT AT HIGH RISK",
                  style={"fontWeight": "bold", "color": "#7F0000",
                         "fontSize": "14px"}),
        html.Span("  |  Debt Service/Revenue: ", style={"marginLeft": "20px"}),
        html.Span("71.2%",
                  style={"fontWeight": "bold", "color": RED, "fontSize": "14px"}),
        html.Span(" (IMF threshold: 30%)"),
        html.Span("  |  PV Debt/GDP: ", style={"marginLeft": "20px"}),
        html.Span("63.7%",
                  style={"fontWeight": "bold", "color": RED, "fontSize": "14px"}),
        html.Span(" (55% anchor)"),
    ], style={
        "background": "#FFF3CD", "color": "#7F0000",
        "padding": "8px 28px", "fontSize": "12px",
        "borderBottom": "1px solid #FFC000"
    }),

    # ── Main content
    html.Div([

        # ── KPI Tiles
        html.Div([
            html.H3("KEY PERFORMANCE INDICATORS  —  FY2024/25",
                    style={"color": NAVY, "fontSize": "13px", "fontWeight": "bold",
                           "marginBottom": "10px", "borderBottom": f"2px solid {NAVY}",
                           "paddingBottom": "4px"}),
            html.Div([kpi_tile(*k) for k in KPIS],
                     style={"display": "flex", "flexWrap": "wrap", "gap": "2px"})
        ], style={"marginBottom": "24px"}),

        # ── Row 1: Debt/GDP + Scenario Fan
        html.Div([
            html.Div([
                dcc.Graph(figure=chart_debt_gdp(), config={"displayModeBar": True},
                          style={"height": "360px"})
            ], style={"flex": "1", "minWidth": "420px",
                      "border": "1px solid #E5E5E5", "borderRadius": "8px",
                      "padding": "10px", "marginRight": "12px",
                      "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),
            html.Div([
                dcc.Graph(figure=chart_scenario_fan(), config={"displayModeBar": True},
                          style={"height": "360px"})
            ], style={"flex": "1", "minWidth": "420px",
                      "border": "1px solid #E5E5E5", "borderRadius": "8px",
                      "padding": "10px",
                      "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),
        ], style={"display": "flex", "flexWrap": "wrap", "marginBottom": "16px"}),

        # ── Row 2: DS/Revenue + Composition
        html.Div([
            html.Div([
                dcc.Graph(figure=chart_ds_revenue(), config={"displayModeBar": True},
                          style={"height": "340px"})
            ], style={"flex": "1", "minWidth": "420px",
                      "border": "1px solid #E5E5E5", "borderRadius": "8px",
                      "padding": "10px", "marginRight": "12px",
                      "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),
            html.Div([
                dcc.Graph(figure=chart_debt_composition(), config={"displayModeBar": True},
                          style={"height": "340px"})
            ], style={"flex": "1", "minWidth": "420px",
                      "border": "1px solid #E5E5E5", "borderRadius": "8px",
                      "padding": "10px",
                      "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),
        ], style={"display": "flex", "flexWrap": "wrap", "marginBottom": "16px"}),

        # ── Row 3: Primary Balance + Sensitivity Heatmap
        html.Div([
            html.Div([
                dcc.Graph(figure=chart_primary_balance(), config={"displayModeBar": True},
                          style={"height": "340px"})
            ], style={"flex": "1", "minWidth": "380px",
                      "border": "1px solid #E5E5E5", "borderRadius": "8px",
                      "padding": "10px", "marginRight": "12px",
                      "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),
            html.Div([
                dcc.Graph(figure=chart_sensitivity_heatmap(), config={"displayModeBar": True},
                          style={"height": "340px"})
            ], style={"flex": "1.4", "minWidth": "480px",
                      "border": "1px solid #E5E5E5", "borderRadius": "8px",
                      "padding": "10px",
                      "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),
        ], style={"display": "flex", "flexWrap": "wrap", "marginBottom": "16px"}),

        # ── IMF Scorecard
        html.Div([
            html.H3("IMF LIC-DSF SCORECARD  —  FY2024/25 vs THRESHOLDS",
                    style={"color": NAVY, "fontSize": "13px", "fontWeight": "bold",
                           "marginBottom": "10px", "borderBottom": f"2px solid {NAVY}",
                           "paddingBottom": "4px"}),
            scorecard_table()
        ], style={"marginBottom": "24px",
                  "border": "1px solid #E5E5E5", "borderRadius": "8px",
                  "padding": "16px",
                  "boxShadow": "1px 1px 5px rgba(0,0,0,0.08)"}),

        # ── Watchpoints
        html.Div([
            html.H3("KEY WATCHPOINTS",
                    style={"color": NAVY, "fontSize": "13px", "fontWeight": "bold",
                           "marginBottom": "10px", "borderBottom": f"2px solid {NAVY}",
                           "paddingBottom": "4px"}),
            html.Div(
                watchpoint_rows(),
                style={"borderRadius": "6px", "overflow": "hidden",
                       "border": "1px solid #E5E5E5"}
            )
        ], style={"marginBottom": "24px"}),

    ], style={"padding": "20px 28px", "fontFamily": "Calibri, Arial, sans-serif",
              "background": "#FAFAFA", "minHeight": "100vh"}),

    # ── Footer
    html.Div([
        html.P(
            "Source: National Treasury, PDMO Debt Bulletin, IMF Article IV (2025). "
            "Projections are model outputs and subject to revision. "
            "For official data refer to PDMO published bulletins.",
            style={"margin": 0, "fontSize": "11px", "color": "#BFBFBF"}
        )
    ], style={
        "background": NAVY, "padding": "10px 28px",
        "textAlign": "center"
    }),

], style={"fontFamily": "Calibri, Arial, sans-serif"})


if __name__ == "__main__":
    print("=" * 55)
    print("  Kenya DSA Interactive Dashboard")
    print("  Open your browser at: http://127.0.0.1:8050")
    print("  Press Ctrl+C to stop the server")
    print("=" * 55)
    app.run(debug=False, host="127.0.0.1", port=8050)
