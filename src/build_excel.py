"""Builds an Excel sales-trend + financial model from daily sales data.
Usage: python src/build_excel.py [sample_data/Sales by day.csv] [examples/Sample_Financial_Model.xlsx]
All numbers in the shipped example are SYNTHETIC; cost inputs are placeholders."""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
import datetime as dt

SRC = sys.argv[1] if len(sys.argv) > 1 else "sample_data/Sales by day.csv"
OUT = sys.argv[2] if len(sys.argv) > 2 else "examples/Sample_Financial_Model.xlsx"
d = pd.read_csv(SRC)
d["date"] = pd.to_datetime(d["yyyyMMdd"].astype(str))
d = d.rename(columns={"Net sales": "net", "Total orders": "orders", "Total guests": "guests"})
d = d[d.date <= "2026-10-07"].reset_index(drop=True)
F = "Arial"
blue = Font(name=F, color="0000FF"); black = Font(name=F); bold = Font(name=F, bold=True)
hdr = Font(name=F, bold=True, color="FFFFFF"); hfill = PatternFill("solid", fgColor="1F3A5F")
yellow = PatternFill("solid", fgColor="FFFF00"); grey = PatternFill("solid", fgColor="F2F2F2")
title = Font(name=F, bold=True, size=14)
USD = '$#,##0;($#,##0);-'; USD2 = '$#,##0.00;($#,##0.00);-'; PCT = '0.0%;(0.0%);-'; NUM = '#,##0;(#,##0);-'; DATE = 'yyyy-mm-dd'

def header(ws, row, labels, widths=None):
    for i, l in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=l); c.font = hdr; c.fill = hfill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w

wb = Workbook()

# ---------------- Legend ----------------
ws = wb.active; ws.title = "Legend"
ws["A1"] = "Restaurant Sales Trend & Financial Model - SAMPLE DATA (synthetic)"; ws["A1"].font = title
lines = [
 ("Purpose", "Quantify a late-summer sales decline (May 1 - Oct 7, 2026) and estimate profit under editable cost assumptions."),
 ("Data source", "SYNTHETIC sample data in the same format as a POS 'Sales by day' export. No real business data is included in this repository."),
 ("Blue text", "Hardcoded input or raw POS data."),
 ("Black text", "Formula - do not overwrite."),
 ("Yellow fill", "Key assumption you should edit. COST INPUTS ARE PLACEHOLDERS until the owners supply real figures."),
 ("Tabs", "Assumptions > Daily (raw) > Weekly > Decomposition (orders vs. avg check, day of week) > Scenario > Monthly P&L (estimate)."),
 ("Important", "Revenue-based findings (Weekly, Decomposition) use real POS data only. Anything involving cost (Scenario profit, Monthly P&L) depends on placeholder inputs."),
]
for i, (a, b) in enumerate(lines, 3):
    ws.cell(row=i, column=1, value=a).font = bold
    c = ws.cell(row=i, column=2, value=b); c.font = black; c.alignment = Alignment(wrap_text=True, vertical="top")
ws["A5"].font = Font(name=F, bold=True, color="0000FF"); ws["B6"].font = black
ws["A7"].fill = yellow
ws.column_dimensions["A"].width = 16; ws.column_dimensions["B"].width = 110
for r in range(3, 10): ws.row_dimensions[r].height = 32

# ---------------- Assumptions ----------------
wa = wb.create_sheet("Assumptions")
wa["A1"] = "Assumptions"; wa["A1"].font = title
wa.column_dimensions["A"].width = 46; wa.column_dimensions["B"].width = 16; wa.column_dimensions["C"].width = 80
rows = [
 (3, "Comparison periods (week-start dates, Mondays)", None, None, "h"),
 (4, "Base period - first week start", dt.date(2026,5,4), "First full Mon-Sun week of data", "d"),
 (5, "Base period - last week start", dt.date(2026,8,24), "Week of Aug 24-30; late summer peak weeks included", "d"),
 (6, "Recent period - first week start", dt.date(2026,9,7), "Aug 31-Sep 6 (Labor Day week) left out of both periods", "d"),
 (7, "Recent period - last week start", dt.date(2026,9,28), "Last full week of data (Sep 28 - Oct 4)", "d"),
 (8, "Base period - last day", "=B5+6", "Formula", "f"),
 (9, "Recent period - last day", "=B7+6", "Formula", "f"),
 (11, "Cost assumptions (PLACEHOLDERS - replace with owner figures)", None, None, "h"),
 (12, "Food & beverage cost (% of net sales)", 0.30, "PLACEHOLDER. Typical full-service range is roughly 28-35%; owners' actual figure needed.", "p"),
 (13, "Labor cost (% of net sales)", 0.30, "PLACEHOLDER. Owners' payroll figures needed.", "p"),
 (14, "Other operating costs (% of net sales)", 0.10, "PLACEHOLDER. Supplies, utilities, processing fees, etc.", "p"),
 (15, "Rent & fixed costs per month ($)", 8000, "PLACEHOLDER. Owners' actual figure needed.", "u"),
 (17, "Scenario lever", None, None, "h"),
 (18, "Share of lost weeknight sales recovered", 0.25, "Illustrative lever for the Scenario tab (not a forecast).", "p"),
]
for r, a, b, c, kind in rows:
    wa.cell(row=r, column=1, value=a).font = bold if kind == "h" else black
    if kind == "h":
        wa.cell(row=r, column=1).fill = grey; wa.cell(row=r, column=2).fill = grey; wa.cell(row=r, column=3).fill = grey
        continue
    cell = wa.cell(row=r, column=2, value=b)
    cell.font = black if kind == "f" else blue
    if kind in ("d", "f"): cell.number_format = DATE
    if kind == "p": cell.number_format = '0.0%'
    if kind == "u": cell.number_format = USD
    if kind in ("p", "u") or (kind == "d" and r in (4,5,6,7)): cell.fill = yellow if kind in ("p","u") else PatternFill()
    wa.cell(row=r, column=3, value=c).font = black

# ---------------- Daily ----------------
wd = wb.create_sheet("Daily")
header(wd, 1, ["Date", "Day of week", "Week start (Mon)", "Net sales ($)", "Orders", "Guests", "Avg check ($)"], [13, 14, 16, 14, 10, 10, 14])
for i, r in enumerate(d.itertuples(), 2):
    wd.cell(row=i, column=1, value=r.date.date()).font = blue; wd.cell(row=i, column=1).number_format = DATE
    wd.cell(row=i, column=2, value=f'=TEXT(A{i},"dddd")').font = black
    c = wd.cell(row=i, column=3, value=f"=A{i}-WEEKDAY(A{i},2)+1"); c.font = black; c.number_format = DATE
    c = wd.cell(row=i, column=4, value=float(r.net)); c.font = blue; c.number_format = USD2
    wd.cell(row=i, column=5, value=int(r.orders)).font = blue
    wd.cell(row=i, column=6, value=int(r.guests)).font = blue
    c = wd.cell(row=i, column=7, value=f"=IF(E{i}=0,0,D{i}/E{i})"); c.font = black; c.number_format = USD2
last = len(d) + 1
wd.freeze_panes = "A2"

# ---------------- Weekly ----------------
ww = wb.create_sheet("Weekly")
header(ww, 1, ["Week start (Mon)", "Net sales ($)", "Orders", "Guests", "Avg check ($)", "Spend / guest ($)", "4-week avg net sales ($)"], [16, 14, 10, 10, 14, 16, 22])
weeks = pd.date_range("2026-05-04", "2026-09-28", freq="7D")
for i, wk in enumerate(weeks, 2):
    c = ww.cell(row=i, column=1, value=wk.date()); c.font = blue; c.number_format = DATE
    c = ww.cell(row=i, column=2, value=f"=SUMIFS(Daily!$D$2:$D${last},Daily!$C$2:$C${last},A{i})"); c.number_format = USD
    ww.cell(row=i, column=3, value=f"=SUMIFS(Daily!$E$2:$E${last},Daily!$C$2:$C${last},A{i})").number_format = NUM
    ww.cell(row=i, column=4, value=f"=SUMIFS(Daily!$F$2:$F${last},Daily!$C$2:$C${last},A{i})").number_format = NUM
    ww.cell(row=i, column=5, value=f"=IF(C{i}=0,0,B{i}/C{i})").number_format = USD2
    ww.cell(row=i, column=6, value=f"=IF(D{i}=0,0,B{i}/D{i})").number_format = USD2
    if i >= 5: ww.cell(row=i, column=7, value=f"=AVERAGE(B{i-3}:B{i})").number_format = USD
    for col in range(2, 8): ww.cell(row=i, column=col).font = black
wlast = len(weeks) + 1
ww.freeze_panes = "A2"

# ---------------- Decomposition ----------------
wc = wb.create_sheet("Decomposition")
wc["A1"] = "Why did sales fall?  Net sales = Orders x Average check"; wc["A1"].font = title
header(wc, 3, ["Weekly average", "Base period", "Recent period", "Change", "% change"], [40, 16, 16, 16, 12])
W = f"Weekly!$A$2:$A${wlast}"
def avgw(col, lo, hi): return f"=AVERAGEIFS(Weekly!${col}$2:${col}${wlast},{W},\">=\"&Assumptions!{lo},{W},\"<=\"&Assumptions!{hi})"
wc["A4"] = "Net sales ($ / week)"; wc["B4"] = avgw("B", "$B$4", "$B$5"); wc["C4"] = avgw("B", "$B$6", "$B$7")
wc["A5"] = "Orders / week"; wc["B5"] = avgw("C", "$B$4", "$B$5"); wc["C5"] = avgw("C", "$B$6", "$B$7")
wc["A6"] = "Guests / week"; wc["B6"] = avgw("D", "$B$4", "$B$5"); wc["C6"] = avgw("D", "$B$6", "$B$7")
wc["A7"] = "Average check ($) = net sales / orders"; wc["B7"] = "=B4/B5"; wc["C7"] = "=C4/C5"
for r in range(4, 8):
    wc[f"D{r}"] = f"=C{r}-B{r}"; wc[f"E{r}"] = f"=IF(B{r}=0,0,D{r}/B{r})"
    wc[f"E{r}"].number_format = PCT
for r, fmt in ((4, USD), (5, '#,##0.0'), (6, '#,##0.0'), (7, USD2)):
    for col in "BCD": wc[f"{col}{r}"].number_format = fmt
wc["A9"] = "Where did the drop come from? (sums exactly to the total change)"; wc["A9"].font = bold
wc["A10"] = "Traffic effect = (recent orders - base orders) x base avg check"; wc["B10"] = "=(C5-B5)*B7"
wc["A11"] = "Average-check effect = recent orders x (recent avg check - base avg check)"; wc["B11"] = "=C5*(C7-B7)"
wc["A12"] = "Total change in weekly net sales"; wc["B12"] = "=B10+B11"
wc["A13"] = "Check: equals change in net sales row above"; wc["B13"] = '=IF(ABS(B12-D4)<0.01,"OK","MISMATCH")'
wc["A14"] = "Traffic effect as % of total change"; wc["B14"] = "=IF(B12=0,0,B10/B12)"
for r in (10, 11, 12): wc[f"B{r}"].number_format = USD
wc["B14"].number_format = '0%'
for row in wc.iter_rows(min_row=4, max_row=14):
    for c in row:
        if c.font != bold and c.column > 1: c.font = black
        if c.column == 1 and c.row not in (9,): c.font = black
wc["A9"].font = bold; wc["A12"].font = bold; wc["B12"].font = bold

wc["A17"] = "Average net sales by day of week"; wc["A17"].font = bold
header(wc, 18, ["Day", "Base ($/day)", "Recent ($/day)", "Change ($/day)", "% change"])
D = f"Daily!$B$2:$B${last}"; DD = f"Daily!$A$2:$A${last}"; DN = f"Daily!$D$2:$D${last}"
for i, day in enumerate(["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"], 19):
    wc[f"A{i}"] = day
    wc[f"B{i}"] = f'=AVERAGEIFS({DN},{D},A{i},{DD},">="&Assumptions!$B$4,{DD},"<="&Assumptions!$B$8)'
    wc[f"C{i}"] = f'=AVERAGEIFS({DN},{D},A{i},{DD},">="&Assumptions!$B$6,{DD},"<="&Assumptions!$B$9)'
    wc[f"D{i}"] = f"=C{i}-B{i}"; wc[f"E{i}"] = f"=IF(B{i}=0,0,D{i}/B{i})"
    for col in "BCD": wc[f"{col}{i}"].number_format = USD
    wc[f"E{i}"].number_format = PCT
    for col in "ABCDE": wc[f"{col}{i}"].font = black
wc["A26"] = "Total per week"; wc["A26"].font = bold
for col in "BCD":
    wc[f"{col}26"] = f"=SUM({col}19:{col}25)"; wc[f"{col}26"].number_format = USD; wc[f"{col}26"].font = bold
wc["E26"] = "=IF(B26=0,0,D26/B26)"; wc["E26"].number_format = PCT; wc["E26"].font = bold

# ---------------- Scenario ----------------
wsn = wb.create_sheet("Scenario")
wsn["A1"] = "Illustrative scenario: win back some weeknight sales"; wsn["A1"].font = title
wsn.column_dimensions["A"].width = 62; wsn.column_dimensions["B"].width = 18; wsn.column_dimensions["C"].width = 60
wsn["A3"] = "Weekly sales gap on days that fell (Mon-Thu + Sun), $ / week"; wsn["B3"] = '=-SUMIF(Decomposition!D19:D25,"<0")'
wsn["A4"] = "Share of that gap recovered (from Assumptions)"; wsn["B4"] = "=Assumptions!B18"
wsn["A5"] = "Added net sales per week"; wsn["B5"] = "=B3*B4"
wsn["A6"] = "Added gross profit per week (net of placeholder food cost)"; wsn["B6"] = "=B5*(1-Assumptions!B12)"
wsn["A7"] = "Added gross profit per 13 weeks (one quarter)"; wsn["B7"] = "=B6*13"
wsn["C7"] = "Quarter, not annualized: seasonality is unknown without prior-year data."
wsn["C6"] = "Depends on PLACEHOLDER food cost %."
for r, fmt in ((3, USD), (4, '0.0%'), (5, USD), (6, USD), (7, USD)): wsn[f"B{r}"].number_format = fmt
for row in wsn.iter_rows(min_row=3, max_row=7):
    for c in row: c.font = black
wsn["B4"].font = Font(name=F, color="008000")

# ---------------- Monthly P&L ----------------
wp = wb.create_sheet("Monthly P&L (est)")
wp["A1"] = "Estimated monthly operating profit  -  USES PLACEHOLDER COST INPUTS"; wp["A1"].font = Font(name=F, bold=True, size=14, color="C00000")
header(wp, 3, ["Month", "Start", "End", "Days", "Net sales ($)", "Net sales / day ($)", "Food cost ($)", "Labor ($)", "Other opex ($)", "Rent & fixed ($)", "Est. operating profit ($)", "Margin"],
       [14, 12, 12, 7, 15, 16, 14, 14, 14, 15, 20, 10])
months = [("May", "2026-05-01", "2026-05-31"), ("June", "2026-06-01", "2026-06-30"), ("July", "2026-07-01", "2026-07-31"),
          ("August", "2026-08-01", "2026-08-31"), ("September", "2026-09-01", "2026-09-30"), ("Oct 1-7 (partial)", "2026-10-01", "2026-10-07")]
for i, (m, s, e) in enumerate(months, 4):
    wp.cell(row=i, column=1, value=m).font = black
    c = wp.cell(row=i, column=2, value=dt.date.fromisoformat(s)); c.number_format = DATE; c.font = blue
    c = wp.cell(row=i, column=3, value=dt.date.fromisoformat(e)); c.number_format = DATE; c.font = blue
    wp.cell(row=i, column=4, value=f"=C{i}-B{i}+1")
    wp.cell(row=i, column=5, value=f'=SUMIFS({DN},{DD},">="&B{i},{DD},"<="&C{i})').number_format = USD
    wp.cell(row=i, column=6, value=f"=E{i}/D{i}").number_format = USD
    wp.cell(row=i, column=7, value=f"=E{i}*Assumptions!$B$12").number_format = USD
    wp.cell(row=i, column=8, value=f"=E{i}*Assumptions!$B$13").number_format = USD
    wp.cell(row=i, column=9, value=f"=E{i}*Assumptions!$B$14").number_format = USD
    wp.cell(row=i, column=10, value=f"=Assumptions!$B$15*D{i}/DAY(EOMONTH(B{i},0))").number_format = USD
    wp.cell(row=i, column=11, value=f"=E{i}-G{i}-H{i}-I{i}-J{i}").number_format = USD
    wp.cell(row=i, column=12, value=f"=IF(E{i}=0,0,K{i}/E{i})").number_format = PCT
    for col in range(4, 13): wp.cell(row=i, column=col).font = black
wp["A11"] = "Rent is prorated by days in the month. Every cost line is a % of sales or a flat placeholder - replace Assumptions B12:B15 with real figures before relying on any profit number."
wp["A11"].font = Font(name=F, italic=True); wp.freeze_panes = "B4"

wb.save(OUT)
print("saved")
