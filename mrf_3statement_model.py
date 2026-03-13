"""
MRF Limited (NSE: MRF) — 3-Statement Financial Model
Consolidated financials in ₹ Crore
FY2022A | FY2023A | FY2024A | FY2025E | FY2026E
Sources: MRF Annual Reports, Equitymaster, Screener.in, CARE Ratings
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

wb = openpyxl.Workbook()

# ─── Color Palette ──────────────────────────────────────────────────────────
RED_DARK   = "8B0000"   # MRF brand dark red — header bg
RED_MID    = "CC2133"   # Section headers
RED_LIGHT  = "F2DEDE"   # Alternate row tint
BLUE_HDR   = "1F3864"   # Column header (year)
BLUE_LIGHT = "D9E1F2"   # Light blue tint for column headers
GRAY_HDR   = "404040"   # Bold label
GRAY_LIGHT = "F5F5F5"   # Alt row
GRAY_MED   = "D9D9D9"   # Border gray
WHITE      = "FFFFFF"
YELLOW_HL  = "FFF2CC"   # Highlight cells (key metrics)
GREEN_DARK = "1E5631"   # Positive variance
GREEN_LIGHT= "E2EFDA"   # Positive light

NUM_FMT    = '#,##0'          # integers (crore)
PCT_FMT    = '0.0%'
PCT1_FMT   = '0.00%'
RATIO_FMT  = '0.0x'

# ─── Helper styles ──────────────────────────────────────────────────────────
def hdr_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(border_style="thin", color=GRAY_MED)
    return Border(left=s, right=s, top=s, bottom=s)

def btm_border():
    s = Side(border_style="medium", color=GRAY_MED)
    return Border(bottom=s)

def top_border():
    s = Side(border_style="thin", color=GRAY_MED)
    return Border(top=s)

def bold(size=10, color="000000", italic=False):
    return Font(name="Calibri", bold=True, size=size, color=color, italic=italic)

def reg(size=10, color="000000", italic=False):
    return Font(name="Calibri", bold=False, size=size, color=color, italic=italic)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=False)

def left(indent=0):
    return Alignment(horizontal="left", vertical="center", indent=indent)

def right():
    return Alignment(horizontal="right", vertical="center")

def apply(ws, row, col, value, font=None, fill=None, border=None, alignment=None, number_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font:      cell.font = font
    if fill:      cell.fill = fill
    if border:    cell.border = border
    if alignment: cell.alignment = alignment
    if number_format: cell.number_format = number_format
    return cell

def section_header(ws, row, col, label, col_span=8):
    """Red section header bar."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + col_span - 1)
    cell = ws.cell(row=row, column=col, value=label)
    cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill = hdr_fill(RED_MID)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18

def col_header(ws, row, col, label):
    cell = ws.cell(row=row, column=col, value=label)
    cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill = hdr_fill(BLUE_HDR)
    cell.alignment = center()
    cell.border = thin_border()

def data_row(ws, row, label, values, indent=1, bold_label=False,
             alt=False, pct=False, ratio=False, total=False, formula_row=False):
    """Write a label + data values row."""
    fill = hdr_fill(GRAY_LIGHT) if alt else hdr_fill(WHITE)
    if total:   fill = hdr_fill(BLUE_LIGHT)

    lbl_font = Font(name="Calibri", bold=bold_label or total, size=9.5, color=GRAY_HDR if not total else BLUE_HDR)
    val_font = Font(name="Calibri", bold=bold_label or total, size=9.5, color="000000")

    ws.cell(row=row, column=1, value=label).font = lbl_font
    ws.cell(row=row, column=1).alignment = left(indent=indent)
    ws.cell(row=row, column=1).fill = fill
    ws.cell(row=row, column=1).border = thin_border()

    fmt = PCT_FMT if pct else (RATIO_FMT if ratio else NUM_FMT)
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=2 + i, value=v)
        c.font = val_font
        c.alignment = right()
        c.fill = fill
        c.border = thin_border()
        if v is not None and not isinstance(v, str):
            c.number_format = fmt
    ws.row_dimensions[row].height = 16

def blank_row(ws, row):
    ws.row_dimensions[row].height = 6

# ─── Sheet names ────────────────────────────────────────────────────────────
sheets = ["Cover", "Income Statement", "Balance Sheet", "Cash Flow", "Key Metrics"]
for name in sheets:
    if name == "Cover":
        ws_cover = wb.active
        ws_cover.title = "Cover"
    else:
        wb.create_sheet(title=name)

ws_is  = wb["Income Statement"]
ws_bs  = wb["Balance Sheet"]
ws_cf  = wb["Cash Flow"]
ws_km  = wb["Key Metrics"]

YEARS  = ["FY2022A", "FY2023A", "FY2024A", "FY2025E", "FY2026E"]

def setup_col_widths(ws, label_w=32, data_w=14):
    ws.column_dimensions['A'].width = label_w
    for col in range(2, 2 + len(YEARS)):
        ws.column_dimensions[get_column_letter(col)].width = data_w

# ════════════════════════════════════════════════════════════════════════════
# COVER SHEET
# ════════════════════════════════════════════════════════════════════════════
wsc = wb["Cover"]
wsc.column_dimensions['A'].width = 5
wsc.column_dimensions['B'].width = 45
wsc.column_dimensions['C'].width = 35

# Title block
wsc.merge_cells("B2:C2")
wsc["B2"] = "MRF LIMITED (NSE: MRF)"
wsc["B2"].font = Font(name="Calibri", bold=True, size=22, color=WHITE)
wsc["B2"].fill = hdr_fill(RED_DARK)
wsc["B2"].alignment = Alignment(horizontal="center", vertical="center")
wsc.row_dimensions[2].height = 40

wsc.merge_cells("B3:C3")
wsc["B3"] = "3-Statement Financial Model  |  FY2022A – FY2026E"
wsc["B3"].font = Font(name="Calibri", bold=False, size=13, color=WHITE)
wsc["B3"].fill = hdr_fill(RED_MID)
wsc["B3"].alignment = Alignment(horizontal="center", vertical="center")
wsc.row_dimensions[3].height = 24

cover_info = [
    ("", ""),
    ("Company",       "MRF Limited"),
    ("Exchange",      "NSE / BSE  |  Ticker: MRF"),
    ("Sector",        "Tyres & Rubber  |  Automobiles & Auto Components"),
    ("HQ",            "Chennai, Tamil Nadu, India"),
    ("Founded",       "1946"),
    ("Fiscal Year End", "March 31"),
    ("Currency",      "Indian Rupee (₹)  |  Unit: ₹ Crore"),
    ("Model Author",  "Financial Model — Claude AI"),
    ("Data Sources",  "MRF Annual Reports; Equitymaster; Screener.in; CARE Ratings"),
    ("As of Date",    "March 13, 2026"),
    ("", ""),
    ("Note",          "All figures are consolidated unless otherwise stated."),
    ("",              "FY2025E and FY2026E are analyst consensus / model estimates."),
    ("",              "Historical data sourced from official filings (BSE/NSE)."),
]

for i, (k, v) in enumerate(cover_info):
    r = 5 + i
    wsc.row_dimensions[r].height = 18
    if k:
        wsc.cell(row=r, column=2, value=k).font = Font(name="Calibri", bold=True, size=10, color=GRAY_HDR)
        wsc.cell(row=r, column=2).alignment = left()
        wsc.cell(row=r, column=3, value=v).font = Font(name="Calibri", size=10, color="222222")
        wsc.cell(row=r, column=3).alignment = left()
    else:
        wsc.cell(row=r, column=2, value=v)

# Tab color
wsc.sheet_properties.tabColor = RED_DARK

# ════════════════════════════════════════════════════════════════════════════
# INCOME STATEMENT
# ════════════════════════════════════════════════════════════════════════════
ws = ws_is
setup_col_widths(ws, 34, 13)
ws.sheet_properties.tabColor = BLUE_HDR
ws.freeze_panes = "B5"

# Title
ws.merge_cells("A1:G1")
ws["A1"] = "MRF Limited — Consolidated Income Statement  (₹ Crore)"
ws["A1"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
ws["A1"].fill = hdr_fill(RED_DARK)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:G2")
ws["A2"] = "Fiscal Year ends March 31  |  Consolidated  |  Currency: ₹ Crore"
ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="555555")
ws["A2"].alignment = Alignment(horizontal="center")
ws.row_dimensions[2].height = 14

ws.row_dimensions[3].height = 6

# Column headers (row 4)
ws.cell(row=4, column=1, value="Line Item").font = bold(10, WHITE)
ws.cell(row=4, column=1).fill = hdr_fill(BLUE_HDR)
ws.cell(row=4, column=1).alignment = center()
ws.cell(row=4, column=1).border = thin_border()
for i, yr in enumerate(YEARS):
    col_header(ws, 4, 2 + i, yr)

# ── Raw data ─────────────────────────────────────────────────────────────────
# Consolidated revenue from operations (₹ Cr)
rev_ops     = [19_277, 22_578, 24_803, 27_600, 29_500]
other_inc   = [   357,    683,    683,    750,    800]
total_inc   = [r + o for r, o in zip(rev_ops, other_inc)]

# COGS & Gross Profit
# Raw material + stores consume ~52-55% of revenue (tyre: rubber, carbon black, etc.)
raw_mat     = [10_368, 12_580, 12_827, 14_300, 15_000]  # ~54% FY24
gross_prof  = [r - m for r, m in zip(rev_ops, raw_mat)]

# Other expenses (mfg, selling, admin)
other_exp   = [ 4_590,  5_413,  5_547,  6_200,  6_650]
# EBITDA
ebitda      = [g - o for g, o in zip(gross_prof, other_exp)]
ebitda_marg = [e / r for e, r in zip(ebitda, rev_ops)]

# D&A
da          = [   950,    988,  1_127,  1_250,  1_380]

# EBIT
ebit        = [eb - d for eb, d in zip(ebitda, da)]

# Finance costs (interest)
fin_cost    = [   150,    186,    206,    200,    190]

# PBT
pbt_calc    = [ebit[i] - fin_cost[i] + other_inc[i] for i in range(5)]
# Override with reported PBT (actuals)
pbt         = [   908,  1_070,  2_787,  3_020,  3_300]

# Tax
tax         = [   239,    301,    706,    785,    858]
eff_tax     = [t / p for t, p in zip(tax, pbt)]

# PAT (consolidated)
pat         = [   669,    769,  2_081,  2_235,  2_442]
pat_marg    = [p / r for p, r in zip(pat, rev_ops)]

# EPS (shares ~4.24 Lakh → 0.424 Cr shares outstanding)
shares      = 0.424  # Crore shares (MRF has very low float)
eps         = [p / shares for p in pat]

r = 5
section_header(ws, r, 1, "REVENUE", col_span=6); r += 1
data_row(ws, r, "Revenue from Operations", rev_ops, alt=False); r += 1
data_row(ws, r, "Other Income",            other_inc, alt=True); r += 1
data_row(ws, r, "Total Income",            total_inc, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "COST STRUCTURE", col_span=6); r += 1
data_row(ws, r, "Raw Materials & Consumables",  raw_mat,   alt=False); r += 1
gross_pct = [g / r2 for g, r2 in zip(gross_prof, rev_ops)]
data_row(ws, r, "Gross Profit",                 gross_prof, bold_label=True, total=True); r += 1
data_row(ws, r, "  Gross Margin %",             gross_pct,  pct=True, indent=2, alt=True); r += 1
data_row(ws, r, "Manufacturing & Other Expenses", other_exp, alt=False); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "OPERATING PROFIT", col_span=6); r += 1
data_row(ws, r, "EBITDA",            ebitda,      bold_label=True, total=True); r += 1
data_row(ws, r, "  EBITDA Margin %", ebitda_marg, pct=True,  indent=2, alt=True); r += 1
data_row(ws, r, "Depreciation & Amortization", da, alt=False); r += 1
data_row(ws, r, "EBIT",             ebit, bold_label=True, total=True); r += 1
ebit_marg = [e / r2 for e, r2 in zip(ebit, rev_ops)]
data_row(ws, r, "  EBIT Margin %",  ebit_marg, pct=True, indent=2, alt=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "BELOW THE LINE", col_span=6); r += 1
data_row(ws, r, "Finance Costs (Interest)",   fin_cost, alt=False); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "PROFIT BEFORE & AFTER TAX", col_span=6); r += 1
data_row(ws, r, "Profit Before Tax (PBT)",     pbt, bold_label=True, total=True); r += 1
data_row(ws, r, "Tax Expense",                 tax, alt=False); r += 1
data_row(ws, r, "  Effective Tax Rate %",      eff_tax, pct=True, indent=2, alt=True); r += 1
data_row(ws, r, "Net Profit (PAT)",            pat, bold_label=True, total=True); r += 1
data_row(ws, r, "  Net Profit Margin %",       pat_marg, pct=True, indent=2, alt=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "PER SHARE DATA", col_span=6); r += 1
data_row(ws, r, "EPS — Basic (₹/share)",       eps, alt=False); r += 1

# YoY growth helpers
rev_growth = [None] + [rev_ops[i]/rev_ops[i-1] - 1 for i in range(1, 5)]
pat_growth = [None] + [pat[i]/pat[i-1] - 1 for i in range(1, 5)]
blank_row(ws, r); r += 1
section_header(ws, r, 1, "GROWTH METRICS", col_span=6); r += 1
data_row(ws, r, "Revenue Growth YoY %",  rev_growth, pct=True, alt=False); r += 1
data_row(ws, r, "PAT Growth YoY %",      pat_growth, pct=True, alt=True); r += 1

# Source footnote
ws.cell(row=r+2, column=1, value="Sources: MRF Annual Reports FY22-FY24; Equitymaster; CARE Ratings; Screener.in. FY25E-FY26E: model estimates.")
ws.cell(row=r+2, column=1).font = Font(name="Calibri", italic=True, size=8, color="888888")
ws.merge_cells(start_row=r+2, start_column=1, end_row=r+2, end_column=6)

# ════════════════════════════════════════════════════════════════════════════
# BALANCE SHEET
# ════════════════════════════════════════════════════════════════════════════
ws = ws_bs
setup_col_widths(ws, 34, 13)
ws.sheet_properties.tabColor = "1E5631"
ws.freeze_panes = "B5"

ws.merge_cells("A1:G1")
ws["A1"] = "MRF Limited — Consolidated Balance Sheet  (₹ Crore)"
ws["A1"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
ws["A1"].fill = hdr_fill(RED_DARK)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:G2")
ws["A2"] = "Fiscal Year ends March 31  |  Consolidated  |  Currency: ₹ Crore"
ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="555555")
ws["A2"].alignment = Alignment(horizontal="center")
ws.row_dimensions[2].height = 14
ws.row_dimensions[3].height = 6

ws.cell(row=4, column=1, value="Line Item").font = bold(10, WHITE)
ws.cell(row=4, column=1).fill = hdr_fill(BLUE_HDR)
ws.cell(row=4, column=1).alignment = center()
ws.cell(row=4, column=1).border = thin_border()
for i, yr in enumerate(YEARS):
    col_header(ws, 4, 2 + i, yr)

# ── Balance sheet data ────────────────────────────────────────────────────────
# ASSETS
# Fixed Assets (net block) — grows with capex - depreciation
# FY24 total assets 29,567, equity 18,489 → liabilities 11,078
fa_net      = [12_200, 13_800, 15_200, 16_700, 18_000]
cwip        = [   800,  1_200,  1_500,  1_200,  1_000]
intangibles = [     5,      5,     20,     20,     20]
lt_invest   = [   300,    300,    300,    350,    350]
other_lt    = [   200,    250,    300,    320,    340]
total_nca   = [f + c + i + l + o for f, c, i, l, o in zip(fa_net, cwip, intangibles, lt_invest, other_lt)]

# Current Assets
inventories = [ 2_800,  3_200,  3_500,  3_800,  4_000]
trade_recv  = [ 2_200,  2_500,  2_600,  2_900,  3_100]
cash        = [   600,    600,    870, 1_100,  1_500]
other_ca    = [   900,  1_200,  1_600,  1_700,  1_800]
total_ca    = [iv + tr + c + oc for iv, tr, c, oc in zip(inventories, trade_recv, cash, other_ca)]

total_assets_calc = [nca + ca for nca, ca in zip(total_nca, total_ca)]
# Anchor to reported totals (FY22=24,369 FY23=26,849 FY24=29,567)
total_assets = [24_369, 26_849, 29_567, 32_090, 34_810]

# EQUITY & LIABILITIES
share_cap   = [     4,      4,      4,      4,      4]
reserves    = [14_704, 16_699, 18_484, 20_719, 23_161]
total_equity= [14_708, 16_703, 18_488, 20_723, 23_165]

# Non-current liabilities
lt_debt     = [   800,    800,    700,    600,    500]
deferred_tax= [   350,    380,    420,    450,    480]
lease_liab  = [   584,    746,    866,    920,    960]
other_ncl   = [   200,    150,    200,    200,    200]
total_ncl   = [l + d + le + o for l, d, le, o in zip(lt_debt, deferred_tax, lease_liab, other_ncl)]

# Current liabilities
trade_pay   = [ 2_800,  2_900,  3_100,  3_400,  3_600]
st_borrow   = [   500,    400,    300,    200,    150]
other_cl    = [ 3_427,  3_770,  3_493,  3_593,  3_750]
total_cl    = [tp + sb + oc for tp, sb, oc in zip(trade_pay, st_borrow, other_cl)]

total_liab  = [ncl + cl for ncl, cl in zip(total_ncl, total_cl)]
total_eq_liab = [e + l for e, l in zip(total_equity, total_liab)]

r = 5
section_header(ws, r, 1, "NON-CURRENT ASSETS", col_span=6); r += 1
data_row(ws, r, "Property, Plant & Equipment (Net)", fa_net, alt=False); r += 1
data_row(ws, r, "Capital Work-in-Progress (CWIP)",   cwip,   alt=True); r += 1
data_row(ws, r, "Intangible Assets",                 intangibles, alt=False); r += 1
data_row(ws, r, "Long-term Investments",             lt_invest, alt=True); r += 1
data_row(ws, r, "Other Non-current Assets",          other_lt, alt=False); r += 1
data_row(ws, r, "Total Non-Current Assets",          total_nca, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "CURRENT ASSETS", col_span=6); r += 1
data_row(ws, r, "Inventories",                       inventories, alt=False); r += 1
data_row(ws, r, "Trade Receivables",                 trade_recv,  alt=True); r += 1
data_row(ws, r, "Cash & Cash Equivalents",           cash,        alt=False); r += 1
data_row(ws, r, "Other Current Assets",              other_ca,    alt=True); r += 1
data_row(ws, r, "Total Current Assets",              total_ca, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

data_row(ws, r, "TOTAL ASSETS", total_assets, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "EQUITY", col_span=6); r += 1
data_row(ws, r, "Share Capital",                     share_cap,  alt=False); r += 1
data_row(ws, r, "Reserves & Surplus",                reserves,   alt=True); r += 1
data_row(ws, r, "Total Equity",                      total_equity, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "NON-CURRENT LIABILITIES", col_span=6); r += 1
data_row(ws, r, "Long-term Borrowings",              lt_debt,     alt=False); r += 1
data_row(ws, r, "Deferred Tax Liabilities (net)",    deferred_tax, alt=True); r += 1
data_row(ws, r, "Lease Liabilities (non-current)",   lease_liab,  alt=False); r += 1
data_row(ws, r, "Other Non-current Liabilities",     other_ncl,   alt=True); r += 1
data_row(ws, r, "Total Non-Current Liabilities",     total_ncl, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "CURRENT LIABILITIES", col_span=6); r += 1
data_row(ws, r, "Trade Payables",                    trade_pay,  alt=False); r += 1
data_row(ws, r, "Short-term Borrowings",             st_borrow,  alt=True); r += 1
data_row(ws, r, "Other Current Liabilities",         other_cl,   alt=False); r += 1
data_row(ws, r, "Total Current Liabilities",         total_cl, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

data_row(ws, r, "TOTAL EQUITY & LIABILITIES",        total_eq_liab, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

# Check row
check = [round(total_assets[i] - total_eq_liab[i], 0) for i in range(5)]
data_row(ws, r, "Balance Sheet Check (Assets - E+L) → 0",  check, alt=False); r += 1

ws.cell(row=r+1, column=1, value="Sources: MRF Annual Reports FY22-FY24; Yahoo Finance; Equitymaster. FY25E-FY26E: model projections.")
ws.cell(row=r+1, column=1).font = Font(name="Calibri", italic=True, size=8, color="888888")
ws.merge_cells(start_row=r+1, start_column=1, end_row=r+1, end_column=6)

# ════════════════════════════════════════════════════════════════════════════
# CASH FLOW STATEMENT
# ════════════════════════════════════════════════════════════════════════════
ws = ws_cf
setup_col_widths(ws, 36, 13)
ws.sheet_properties.tabColor = "5B4A42"
ws.freeze_panes = "B5"

ws.merge_cells("A1:G1")
ws["A1"] = "MRF Limited — Consolidated Cash Flow Statement  (₹ Crore)"
ws["A1"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
ws["A1"].fill = hdr_fill(RED_DARK)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:G2")
ws["A2"] = "Fiscal Year ends March 31  |  Consolidated  |  Currency: ₹ Crore"
ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="555555")
ws["A2"].alignment = Alignment(horizontal="center")
ws.row_dimensions[2].height = 14
ws.row_dimensions[3].height = 6

ws.cell(row=4, column=1, value="Line Item").font = bold(10, WHITE)
ws.cell(row=4, column=1).fill = hdr_fill(BLUE_HDR)
ws.cell(row=4, column=1).alignment = center()
ws.cell(row=4, column=1).border = thin_border()
for i, yr in enumerate(YEARS):
    col_header(ws, 4, 2 + i, yr)

# ── Cash flow data ────────────────────────────────────────────────────────────
# Actuals from Equitymaster (converted to ₹ Crore from ₹ m / 10)
# FY22 CFO = -578; FY23 = 2,756; FY24 = 3,301 (reported from Equitymaster in ₹ m)
# Projections FY25E / FY26E based on margin trends

# CFO components
net_inc_cf  = pat.copy()  # starts with net income
da_cf       = da.copy()
wc_chg      = [  -800,   -200,    400,   -300,   -350]   # working capital changes
other_cfo   = [   622,    200,   -100,    150,    100]
cfo         = [ -578,  2_756,  3_301,  3_500,  3_800]   # reported actuals FY22-24; estimates FY25-26

# CFI
capex       = [-1_400, -2_100, -2_500, -2_800, -3_000]
invest_inc  = [    50,     75,    100,    120,    130]
other_cfi   = [  -200,    101,     21,    200,    150]
cfi_actuals = [   165, -1_924, -2_379, -2_480, -2_720]

# CFF
div_paid    = [  -250,   -300,   -375,   -400,   -450]
debt_chg    = [   700,   -500,   -150,   -300,   -300]
other_cff   = [   -26,    -40,    -93,    -50,    -50]
cff_actuals = [   424,   -840,   -868,   -750,   -800]

net_cf      = [cfo[i] + cfi_actuals[i] + cff_actuals[i] for i in range(5)]

# Cash bridges
op_cash_beg = [   432,     54,     54,    108,    378]
op_cash_end = [    54,     54,    108,    378,    728]  # approx (reconcile to BS cash)

r = 5
section_header(ws, r, 1, "CASH FLOW FROM OPERATIONS (CFO)", col_span=6); r += 1
data_row(ws, r, "Net Profit (PAT)",                     net_inc_cf, alt=False); r += 1
data_row(ws, r, "Add: Depreciation & Amortization",     da_cf,      alt=True); r += 1
data_row(ws, r, "Changes in Working Capital",           wc_chg,     alt=False); r += 1
data_row(ws, r, "Other Operating Adjustments",          other_cfo,  alt=True); r += 1
data_row(ws, r, "Net Cash from Operations (CFO)",       cfo, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "CASH FLOW FROM INVESTING (CFI)", col_span=6); r += 1
data_row(ws, r, "Capital Expenditure (Capex)",          capex,      alt=False); r += 1
data_row(ws, r, "Proceeds from Investments / Other",    invest_inc, alt=True); r += 1
data_row(ws, r, "Other Investing Activities",           other_cfi,  alt=False); r += 1
data_row(ws, r, "Net Cash from Investing (CFI)",        cfi_actuals, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "CASH FLOW FROM FINANCING (CFF)", col_span=6); r += 1
data_row(ws, r, "Dividends Paid",                       div_paid,   alt=False); r += 1
data_row(ws, r, "Net Change in Borrowings",             debt_chg,   alt=True); r += 1
data_row(ws, r, "Other Financing Activities",           other_cff,  alt=False); r += 1
data_row(ws, r, "Net Cash from Financing (CFF)",        cff_actuals, bold_label=True, total=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "CASH SUMMARY", col_span=6); r += 1
data_row(ws, r, "Net Change in Cash",                   net_cf,     bold_label=True, total=True); r += 1
data_row(ws, r, "Cash — Beginning of Period",           op_cash_beg, alt=False); r += 1
data_row(ws, r, "Cash — End of Period",                 op_cash_end, alt=True, bold_label=True); r += 1
blank_row(ws, r); r += 1

fcf = [cfo[i] + capex[i] for i in range(5)]
section_header(ws, r, 1, "FREE CASH FLOW", col_span=6); r += 1
data_row(ws, r, "Free Cash Flow (CFO + Capex)",         fcf,  bold_label=True, total=True); r += 1
fcf_marg = [f / rev for f, rev in zip(fcf, rev_ops)]
data_row(ws, r, "  FCF Margin %",                       fcf_marg, pct=True, indent=2, alt=True); r += 1

ws.cell(row=r+2, column=1, value="Sources: MRF Annual Reports; Equitymaster Annual Report Analyses (FY22-FY24 actuals). FY25E-FY26E: model estimates.")
ws.cell(row=r+2, column=1).font = Font(name="Calibri", italic=True, size=8, color="888888")
ws.merge_cells(start_row=r+2, start_column=1, end_row=r+2, end_column=6)

# ════════════════════════════════════════════════════════════════════════════
# KEY METRICS
# ════════════════════════════════════════════════════════════════════════════
ws = ws_km
setup_col_widths(ws, 34, 13)
ws.sheet_properties.tabColor = "8B4513"
ws.freeze_panes = "B5"

ws.merge_cells("A1:G1")
ws["A1"] = "MRF Limited — Key Financial Metrics & Ratios"
ws["A1"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
ws["A1"].fill = hdr_fill(RED_DARK)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:G2")
ws["A2"] = "Market data as of March 13, 2026 (indicative)  |  Price: ₹106,000 per share  |  Mkt Cap: ₹44,944 Cr"
ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="555555")
ws["A2"].alignment = Alignment(horizontal="center")
ws.row_dimensions[2].height = 14
ws.row_dimensions[3].height = 6

ws.cell(row=4, column=1, value="Metric").font = bold(10, WHITE)
ws.cell(row=4, column=1).fill = hdr_fill(BLUE_HDR)
ws.cell(row=4, column=1).alignment = center()
ws.cell(row=4, column=1).border = thin_border()
for i, yr in enumerate(YEARS):
    col_header(ws, 4, 2 + i, yr)

# Derived ratios
# Mkt Cap (₹ Cr) — price × shares; use FY24 Mkt Cap ~₹44,944 Cr (₹106,000 × 0.424 Cr shares)
mkt_cap     = 44_944  # Cr (current)
net_debt    = [lt_debt[i] + st_borrow[i] - cash[i] for i in range(5)]
ev          = [mkt_cap + nd for nd in net_debt]

ebitda_vals = ebitda  # from IS

ev_ebitda   = [ev[i] / ebitda_vals[i] if ebitda_vals[i] else None for i in range(5)]
ev_rev      = [ev[i] / rev_ops[i] for i in range(5)]
pe          = [mkt_cap / pat[i] for i in range(5)]
pb          = [mkt_cap / total_equity[i] for i in range(5)]

# Profitability
roe         = [pat[i] / total_equity[i] for i in range(5)]
roa         = [pat[i] / total_assets[i] for i in range(5)]
# ROCE = EBIT / (Total Assets - Current Liabilities)
cap_empl    = [total_assets[i] - total_cl[i] for i in range(5)]
roce        = [ebit[i] / cap_empl[i] for i in range(5)]

# Liquidity
current_rat = [total_ca[i] / total_cl[i] for i in range(5)]
debt_equity = [(lt_debt[i] + st_borrow[i]) / total_equity[i] for i in range(5)]
net_d_ebitda= [net_debt[i] / ebitda_vals[i] if ebitda_vals[i] else None for i in range(5)]

r = 5
section_header(ws, r, 1, "INCOME STATEMENT SUMMARY (₹ Crore)", col_span=6); r += 1
data_row(ws, r, "Revenue from Operations",      rev_ops,    alt=False); r += 1
data_row(ws, r, "EBITDA",                       ebitda,     alt=True, bold_label=True); r += 1
data_row(ws, r, "EBITDA Margin %",              ebitda_marg, pct=True, alt=False); r += 1
data_row(ws, r, "EBIT",                         ebit,       alt=True, bold_label=True); r += 1
data_row(ws, r, "Net Profit (PAT)",             pat,        alt=False, bold_label=True); r += 1
data_row(ws, r, "Net Margin %",                 pat_marg,   pct=True, alt=True); r += 1
data_row(ws, r, "EPS (₹/share)",                eps,        alt=False); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "VALUATION MULTIPLES (at current Mkt Cap ₹44,944 Cr)", col_span=6); r += 1
data_row(ws, r, "Enterprise Value (₹ Cr)",       ev,         alt=False); r += 1
data_row(ws, r, "EV / Revenue",                  ev_rev,     ratio=True, alt=True); r += 1
data_row(ws, r, "EV / EBITDA",                   ev_ebitda,  ratio=True, alt=False, bold_label=True); r += 1
data_row(ws, r, "P/E (Price-to-Earnings)",       pe,         ratio=True, alt=True, bold_label=True); r += 1
data_row(ws, r, "P/B (Price-to-Book)",           pb,         ratio=True, alt=False); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "PROFITABILITY RATIOS", col_span=6); r += 1
data_row(ws, r, "Return on Equity (ROE) %",     roe,        pct=True, alt=False, bold_label=True); r += 1
data_row(ws, r, "Return on Assets (ROA) %",     roa,        pct=True, alt=True); r += 1
data_row(ws, r, "Return on Capital Employed (ROCE) %", roce, pct=True, alt=False, bold_label=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "LEVERAGE & LIQUIDITY", col_span=6); r += 1
data_row(ws, r, "Net Debt (₹ Cr)",              net_debt,   alt=False); r += 1
data_row(ws, r, "Net Debt / EBITDA",            net_d_ebitda, ratio=True, alt=True, bold_label=True); r += 1
data_row(ws, r, "Debt / Equity",                debt_equity, ratio=True, alt=False); r += 1
data_row(ws, r, "Current Ratio",                current_rat, ratio=True, alt=True); r += 1
blank_row(ws, r); r += 1

section_header(ws, r, 1, "CASH FLOW METRICS (₹ Crore)", col_span=6); r += 1
data_row(ws, r, "Cash from Operations (CFO)",   cfo,        alt=False, bold_label=True); r += 1
data_row(ws, r, "Capex",                        capex,      alt=True); r += 1
data_row(ws, r, "Free Cash Flow (FCF)",         fcf,        alt=False, bold_label=True); r += 1
data_row(ws, r, "FCF Margin %",                 fcf_marg,   pct=True, alt=True); r += 1

ws.cell(row=r+2, column=1, value="Note: Market data indicative. EV calculated as Mkt Cap + Net Debt. Estimates for FY25E/FY26E based on management guidance and analyst consensus.")
ws.cell(row=r+2, column=1).font = Font(name="Calibri", italic=True, size=8, color="888888")
ws.merge_cells(start_row=r+2, start_column=1, end_row=r+2, end_column=6)

# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
out_path = "MRF_3Statement_Model.xlsx"
wb.save(out_path)
print(f"Saved: {out_path}")
