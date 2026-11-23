"""Build the Sinking Funds Tracker workbook (Google Sheets-compatible XLSX).

Run:  python3 build_tracker.py   -> writes Sinking_Funds_Tracker.xlsx next to this script.
Every number on the Dashboard is a live formula; only the cream "type here" cells hold data.
"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, Font as DFont
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ------------------------------------------------------------------ palette
SAGE = "A3B18A"
SAGE_DEEP = "7F9468"
SAGE_LIGHT = "E4EADB"
SAGE_PALE = "EEF2E8"
OLIVE = "6B705C"
CREAM = "FAF7F0"
CREAM_2 = "F5F0E6"
BEIGE = "E9E0D2"
BEIGE_DEEP = "D6C7A9"
CHARCOAL = "36393B"
GREY = "8A8D8F"
WHITE = "FFFFFF"
ROSE = "F2DCD5"
ROSE_TXT = "A0522D"
AMBER = "F6E7C8"
AMBER_TXT = "8A6A1F"
STONE = "ECEAE6"

FONT = "Inter"
CUR = '$#,##0.00'
CUR0 = '$#,##0'
PCT = '0%'
DATEF = 'mmm d, yyyy'

FUND_FIRST, FUND_LAST = 6, 25          # 20 sinking funds
CON_FIRST, CON_LAST = 6, 1005          # 1,000 contribution rows
PLAN_FIRST, PLAN_LAST = 6, 305         # 300 planner rows
GOAL_FIRST, GOAL_LAST = 6, 20          # 15 savings goals
CAT_FIRST, CAT_LAST = 6, 30            # 25 category slots
PAY_FIRST, PAY_LAST = 6, 20            # 15 payment methods

SF = "'Sinking Funds'"
CO = "Contributions"
MP = "'Monthly Planner'"
SG = "'Savings Goals'"
EC = "'Expense Categories'"
CD = "'Chart Data'"


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def font(size=10, bold=False, color=CHARCOAL, italic=False, name=FONT):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


thin = Side(style="thin", color=BEIGE)
hair = Side(style="hair", color=BEIGE_DEEP)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ROWLINE = Border(bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
RIGHT = Alignment(horizontal="right", vertical="center", indent=1)


def paint(ws, rng, color):
    for row in ws[rng]:
        for c in row:
            c.fill = fill(color)


def bar_formula(pct_ref):
    """10-block text progress bar that works in Google Sheets, Excel and LibreOffice."""
    n = f"ROUND(MIN(1,MAX(0,{pct_ref}))*10,0)"
    return f'REPT("█",{n})&REPT("░",10-{n})'


def sheet_header(ws, title, subtitle, last_col):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2.5
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 30
    paint(ws, f"A1:{last_col}4", CREAM)
    ws["B2"] = title
    ws["B2"].font = font(20, True, CHARCOAL)
    ws["B2"].alignment = Alignment(vertical="center")
    ws["B3"] = subtitle
    ws["B3"].font = font(9.5, False, OLIVE, italic=True)
    ws["B3"].alignment = Alignment(vertical="center")


def table_header(ws, row, headers, kinds):
    """headers: list of (col_letter, text). kinds: dict col->'in'|'auto' for the row-4 tag."""
    for col, text in headers:
        c = ws[f"{col}{row}"]
        c.value = text
        c.font = font(9.5, True, CREAM)
        c.fill = fill(OLIVE)
        c.alignment = CENTER
        c.border = Border(left=Side(style="thin", color=SAGE), right=Side(style="thin", color=SAGE))
        kind = kinds.get(col)
        if kind:
            t = ws[f"{col}{row - 1}"]
            t.value = "✎ TYPE HERE" if kind == "in" else "ƒ AUTOMATIC"
            t.font = font(7.5, True, OLIVE if kind == "in" else SAGE_DEEP)
            t.fill = fill(BEIGE if kind == "in" else SAGE_LIGHT)
            t.alignment = CENTER


def band(ws, first, last, cols, kinds):
    """Alternating rows; input cells cream/white, formula cells sage tints."""
    for r in range(first, last + 1):
        even = (r - first) % 2 == 0
        for col in cols:
            c = ws[f"{col}{r}"]
            if kinds.get(col) == "auto":
                c.fill = fill(SAGE_PALE if even else SAGE_LIGHT)
                c.font = font(10, False, OLIVE)
            else:
                c.fill = fill(WHITE if even else CREAM_2)
                c.font = font(10, False, CHARCOAL)
            c.border = Border(bottom=Side(style="thin", color=BEIGE))
            c.alignment = Alignment(vertical="center")
        ws.row_dimensions[r].height = 21


def add_dv(ws, dv, rng):
    dv.showErrorMessage = True
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def list_dv(src, title, msg):
    return DataValidation(type="list", formula1=src, allow_blank=True, errorStyle="stop",
                          errorTitle=title, error=msg, promptTitle=title, prompt=msg)


def date_dv():
    return DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)",
                          formula2="DATE(2100,12,31)", allow_blank=True, errorStyle="stop",
                          errorTitle="Date needed", error="Please enter a valid date, e.g. 11/15/2026.",
                          promptTitle="Date", prompt="Enter a date, e.g. 11/15/2026")


def money_dv(min_val=0, allow_negative=False):
    if allow_negative:
        return DataValidation(type="decimal", operator="between", formula1="-1000000000",
                              formula2="1000000000", allow_blank=True, errorStyle="stop",
                              errorTitle="Number needed",
                              error="Enter an amount like 150 or 49.99 (use a minus sign for withdrawals).",
                              promptTitle="Amount",
                              prompt="Deposit = positive number. Withdrawal / spending = negative number.")
    return DataValidation(type="decimal", operator="greaterThanOrEqual", formula1=str(min_val),
                          allow_blank=True, errorStyle="stop", errorTitle="Number needed",
                          error="Enter a positive amount like 1200 or 49.99.",
                          promptTitle="Amount", prompt="Enter a positive amount.")


def status_cf(ws, rng, col, first):
    ref = f"${col}{first}"
    styles = [
        ("Completed", SAGE, "FFFFFF"),
        ("Overdue", ROSE, ROSE_TXT),
        ("Due Soon", AMBER, AMBER_TXT),
        ("In Progress", SAGE_LIGHT, OLIVE),
        ("Not Started", STONE, GREY),
    ]
    for text, bg, fg in styles:
        ws.conditional_formatting.add(
            rng, FormulaRule(formula=[f'{ref}="{text}"'], fill=fill(bg),
                             font=Font(name=FONT, bold=True, color=fg), stopIfTrue=True))


def pct_scale(ws, rng):
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="num", start_value=0, start_color="F2DCD5",
        mid_type="num", mid_value=0.5, mid_color="F6E7C8",
        end_type="num", end_value=1, end_color="A3B18A"))


wb = Workbook()

# ====================================================================== SAMPLE DATA
TODAY = dt.date(2026, 11, 23)
categories = [
    ("Emergency Fund", "3–6 months of essential expenses for the unexpected"),
    ("Car Maintenance", "Oil changes, tires, repairs, registration"),
    ("Car Insurance", "Annual or semi-annual premium"),
    ("Christmas", "Gifts, food, decorations and travel"),
    ("Birthday Gifts", "Family and friends' birthdays"),
    ("Holidays & Travel", "Flights, hotels, spending money"),
    ("Home Repairs", "Roof, appliances, plumbing, upkeep"),
    ("Medical Expenses", "Dental, glasses, deductibles, co-pays"),
    ("Education", "Courses, books, school costs"),
    ("Annual Bills", "Subscriptions, memberships, licences"),
    ("Taxes", "Property tax, self-employment tax"),
    ("Clothing", "Seasonal wardrobe, kids' clothes"),
    ("Personal Care", "Haircuts, skincare, wellness"),
    ("Other", "Anything that doesn't fit above"),
]
payment_methods = ["Auto-Transfer", "Bank Transfer", "Cash", "Debit Card", "Mobile App",
                   "Opening Balance", "Withdrawal", "Other"]

# name, category, goal, monthly, target date, priority, note
funds = [
    ("Emergency Fund", "Emergency Fund", 6000, 300, dt.date(2027, 12, 31), "High", "Example – 3 months of essentials"),
    ("Christmas 2026", "Christmas", 800, 100, dt.date(2026, 12, 15), "High", "Example – fully funded!"),
    ("Car Insurance Renewal", "Car Insurance", 1200, 100, dt.date(2027, 3, 1), "High", "Example – annual premium"),
    ("Car Maintenance", "Car Maintenance", 900, 75, dt.date(2027, 6, 30), "Medium", "Example – tires + service"),
    ("Summer Trip – Portugal", "Holidays & Travel", 3000, 250, dt.date(2027, 7, 1), "Medium", "Example – flights + hotel"),
    ("New Roof", "Home Repairs", 5000, 150, dt.date(2028, 6, 30), "Medium", "Example – long-term fund"),
    ("Annual Subscriptions", "Annual Bills", 450, 40, dt.date(2026, 10, 31), "Low", "Example – shows an OVERDUE goal"),
    ("Dental Work", "Medical Expenses", 1500, 120, dt.date(2027, 4, 30), "High", "Example – crown + cleaning"),
    ("Birthday Gifts", "Birthday Gifts", 600, 50, dt.date(2026, 12, 31), "Low", "Example – shows DUE SOON"),
    ("Property Tax", "Taxes", 2400, 200, dt.date(2027, 1, 31), "High", "Example – paid each January"),
    ("Online Course", "Education", 350, 35, dt.date(2027, 9, 1), "Low", "Example – shows NOT STARTED"),
]

# fund, start month, end month (2026), amount, day, method, opening(balance, month)
schedules = [
    ("Emergency Fund", 1, 11, 300, 1, "Auto-Transfer"),
    ("Christmas 2026", 4, 11, 100, 5, "Auto-Transfer"),
    ("Car Insurance Renewal", 3, 11, 100, 10, "Bank Transfer"),
    ("Car Maintenance", 1, 11, 75, 12, "Bank Transfer"),
    ("Summer Trip – Portugal", 6, 11, 250, 15, "Mobile App"),
    ("New Roof", 2, 11, 150, 3, "Auto-Transfer"),
    ("Annual Subscriptions", 2, 10, 40, 8, "Debit Card"),
    ("Dental Work", 7, 11, 120, 18, "Bank Transfer"),
    ("Birthday Gifts", 1, 11, 50, 20, "Cash"),
    ("Property Tax", 1, 11, 200, 2, "Auto-Transfer"),
]
contribs = [
    (dt.date(2026, 1, 2), "Emergency Fund", 2000, "Opening Balance", "Example – money already saved"),
    (dt.date(2026, 2, 3), "New Roof", 1200, "Opening Balance", "Example – money already saved"),
    (dt.date(2026, 8, 22), "Car Maintenance", -180, "Withdrawal", "Example – new tires (spent from fund)"),
    (dt.date(2026, 5, 14), "Birthday Gifts", -60, "Withdrawal", "Example – Mom's birthday gift"),
    (dt.date(2026, 11, 8), "Annual Subscriptions", 20, "Debit Card", "Example – partial top-up"),
]
for name, m0, m1, amt, day, method in schedules:
    for m in range(m0, m1 + 1):
        contribs.append((dt.date(2026, m, day), name, amt, method, "Example – monthly deposit"))
contribs.sort(key=lambda x: (x[0], x[1]))

goals = [
    ("House Down Payment", 25000, 8000, 4200, dt.date(2028, 12, 31), "Example – long-term goal"),
    ("New Car (cash)", 12000, 3000, 2500, dt.date(2027, 12, 31), "Example"),
    ("New Laptop", 1800, 600, 1200, dt.date(2026, 9, 30), "Example – shows COMPLETED"),
    ("Family Reunion Trip", 4000, 500, 1100, dt.date(2026, 11, 1), "Example – shows OVERDUE"),
    ("Kitchen Renovation", 10000, 0, 0, dt.date(2029, 6, 30), "Example – shows NOT STARTED"),
]

# ====================================================================== START HERE
ws0 = wb.active
ws0.title = "Start Here"

# ====================================================================== DASHBOARD
wsd = wb.create_sheet("Dashboard")
# ====================================================================== SINKING FUNDS
ws = wb.create_sheet("Sinking Funds")
sheet_header(ws, "Sinking Funds Overview",
             "Your master list of funds. Type in the cream columns – sage columns calculate automatically.  "
             "★ Rows marked “Example” are sample data – delete them to start fresh.", "Q")
widths = {"B": 26, "C": 19, "D": 14, "E": 15, "F": 15, "G": 14, "H": 14, "I": 11, "J": 15,
          "K": 13, "L": 10, "M": 34, "N": 14, "O": 10, "P": 8, "Q": 8}
for k, v in widths.items():
    ws.column_dimensions[k].width = v
heads = [("B", "Fund Name"), ("C", "Category"), ("D", "Savings Goal"), ("E", "Current Saved Amount"),
         ("F", "Remaining Balance"), ("G", "Monthly Contribution"), ("H", "Target Date"),
         ("I", "Progress %"), ("J", "Progress Bar"), ("K", "Status"), ("L", "Priority"), ("M", "Notes"),
         ("N", "Needed / Month"), ("O", "Days Left"), ("P", "helper"), ("Q", "helper")]
kinds = {c: "in" for c in "BCDGHLM"}
kinds.update({c: "auto" for c in "EFIJKNO"})
table_header(ws, 5, heads, kinds)
band(ws, FUND_FIRST, FUND_LAST, "BCDEFGHIJKLMNO", kinds)
CON_D = f"{CO}!$D${CON_FIRST}:$D${CON_LAST}"
CON_C = f"{CO}!$C${CON_FIRST}:$C${CON_LAST}"
CON_B = f"{CO}!$B${CON_FIRST}:$B${CON_LAST}"
for i, r in enumerate(range(FUND_FIRST, FUND_LAST + 1)):
    if i < len(funds):
        n, cat, goal, mon, td, pri, note = funds[i]
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = n, cat, goal
        ws[f"G{r}"], ws[f"H{r}"], ws[f"L{r}"], ws[f"M{r}"] = mon, td, pri, note
    ws[f"E{r}"] = f'=IF(B{r}="","",SUMIFS({CON_D},{CON_C},B{r}))'
    ws[f"F{r}"] = f'=IF(B{r}="","",MAX(0,N(D{r})-N(E{r})))'
    ws[f"I{r}"] = f'=IF(OR(B{r}="",N(D{r})=0),"",MIN(1,MAX(0,E{r}/D{r})))'
    ws[f"J{r}"] = f'=IF(I{r}="","",{bar_formula(f"I{r}")})'
    ws[f"K{r}"] = (f'=IF(B{r}="","",IF(AND(N(D{r})>0,N(E{r})>=N(D{r})),"Completed",'
                   f'IF(AND(H{r}<>"",H{r}<TODAY()),"Overdue",IF(N(E{r})<=0,"Not Started",'
                   f'IF(AND(H{r}<>"",H{r}-TODAY()<=60),"Due Soon","In Progress")))))')
    ws[f"N{r}"] = (f'=IF(OR(B{r}="",H{r}=""),"",IF(N(F{r})=0,0,F{r}/MAX(1,'
                   f'(YEAR(H{r})-YEAR(TODAY()))*12+MONTH(H{r})-MONTH(TODAY()))))')
    ws[f"O{r}"] = f'=IF(OR(B{r}="",H{r}=""),"",H{r}-TODAY())'
    ws[f"P{r}"] = f'=IF(B{r}="","",ROW())'
    ws[f"Q{r}"] = f'=IF(OR(B{r}="",H{r}="",K{r}="Completed"),"",H{r}+ROW()/100000)'
    for col, fmt in (("D", CUR), ("E", CUR), ("F", CUR), ("G", CUR), ("N", CUR), ("H", DATEF),
                     ("I", PCT), ("O", '#,##0;[Color10]-#,##0')):
        ws[f"{col}{r}"].number_format = fmt
    for col in "DEFGN":
        ws[f"{col}{r}"].alignment = RIGHT
    for col in "HIKLO":
        ws[f"{col}{r}"].alignment = CENTER
    ws[f"J{r}"].font = Font(name="Arial", size=10, color=SAGE_DEEP)
    ws[f"J{r}"].alignment = CENTER
    ws[f"M{r}"].font = font(9, False, GREY, italic=True)
    ws[f"B{r}"].font = font(10, True, CHARCOAL)
    ws[f"B{r}"].alignment = LEFT
# totals row
tr = FUND_LAST + 1
ws.row_dimensions[tr].height = 26
ws[f"B{tr}"] = "TOTAL"
for col in "BCDEFGHIJKLMNO":
    c = ws[f"{col}{tr}"]
    c.fill = fill(BEIGE)
    c.font = font(10, True, CHARCOAL)
    c.border = Border(top=Side(style="medium", color=OLIVE))
    c.alignment = RIGHT
ws[f"B{tr}"].alignment = LEFT
rng = lambda col: f"{col}{FUND_FIRST}:{col}{FUND_LAST}"
for col in "DEFGN":
    ws[f"{col}{tr}"] = f"=SUM({rng(col)})"
    ws[f"{col}{tr}"].number_format = CUR
ws[f"I{tr}"] = f'=IFERROR(SUM({rng("E")})/SUM({rng("D")}),0)'
ws[f"I{tr}"].number_format = PCT
ws[f"I{tr}"].alignment = CENTER
ws[f"J{tr}"] = f'={bar_formula(f"I{tr}")}'
ws[f"J{tr}"].font = Font(name="Arial", size=10, bold=True, color=OLIVE)
ws[f"J{tr}"].alignment = CENTER
ws[f"K{tr}"] = f'=COUNTIF({rng("K")},"Completed")&" done"'
ws[f"K{tr}"].alignment = CENTER
ws.column_dimensions["P"].hidden = True
ws.column_dimensions["Q"].hidden = True
ws.freeze_panes = "C6"
ws.sheet_properties.tabColor = SAGE

add_dv(ws, list_dv(f"{EC}!$B${CAT_FIRST}:$B${CAT_LAST}", "Category",
                   "Pick a category from the list (edit the list on the Expense Categories tab)."),
       rng("C"))
add_dv(ws, list_dv('"High,Medium,Low"', "Priority", "Choose High, Medium or Low."), rng("L"))
add_dv(ws, money_dv(), rng("D"))
add_dv(ws, money_dv(), rng("G"))
add_dv(ws, date_dv(), rng("H"))
uniq = DataValidation(type="custom", formula1=f"COUNTIF($B${FUND_FIRST}:$B${FUND_LAST},B{FUND_FIRST})=1",
                      allow_blank=True, errorStyle="stop", errorTitle="Duplicate name",
                      error="Each fund needs a unique name.", promptTitle="Fund name",
                      prompt="Give each fund a unique name, e.g. 'Christmas 2026'.")
add_dv(ws, uniq, rng("B"))
status_cf(ws, rng("K"), "K", FUND_FIRST)
pct_scale(ws, rng("I"))
ws.conditional_formatting.add(rng("H"), FormulaRule(
    formula=[f'AND($H{FUND_FIRST}<>"",$H{FUND_FIRST}<TODAY(),$K{FUND_FIRST}<>"Completed")'],
    font=Font(name=FONT, bold=True, color=ROSE_TXT)))
ws.conditional_formatting.add(rng("L"), FormulaRule(
    formula=[f'$L{FUND_FIRST}="High"'], font=Font(name=FONT, bold=True, color=ROSE_TXT)))

# ====================================================================== CONTRIBUTIONS
wc = wb.create_sheet("Contributions")
sheet_header(wc, "Contribution Tracker",
             "Log every deposit (positive) or withdrawal (negative). Fund totals update everywhere automatically.  "
             "★ “Example” rows are sample data.", "N")
for k, v in {"B": 15, "C": 26, "D": 16, "E": 17, "F": 36, "G": 12, "H": 3, "I": 26, "J": 16, "K": 11,
             "L": 15, "M": 3}.items():
    wc.column_dimensions[k].width = v
ckinds = {c: "in" for c in "BCDEF"}
ckinds["G"] = "auto"
table_header(wc, 5, [("B", "Date"), ("C", "Fund Name"), ("D", "Contribution Amount"),
                     ("E", "Payment Method"), ("F", "Notes"), ("G", "Month")], ckinds)
band(wc, CON_FIRST, CON_LAST, "BCDEFG", ckinds)
for i, r in enumerate(range(CON_FIRST, CON_LAST + 1)):
    if i < len(contribs):
        d, n, a, m, note = contribs[i]
        wc[f"B{r}"], wc[f"C{r}"], wc[f"D{r}"], wc[f"E{r}"], wc[f"F{r}"] = d, n, a, m, note
    wc[f"G{r}"] = f'=IF(B{r}="","",DATE(YEAR(B{r}),MONTH(B{r}),1))'
    wc[f"B{r}"].number_format = DATEF
    wc[f"B{r}"].alignment = CENTER
    wc[f"D{r}"].number_format = '$#,##0.00;[Color9]-$#,##0.00'
    wc[f"D{r}"].alignment = RIGHT
    wc[f"E{r}"].alignment = CENTER
    wc[f"F{r}"].font = font(9, False, GREY, italic=True)
    wc[f"G{r}"].number_format = 'mmm yyyy'
    wc[f"G{r}"].alignment = CENTER
    wc[f"C{r}"].alignment = LEFT
add_dv(wc, date_dv(), f"B{CON_FIRST}:B{CON_LAST}")
add_dv(wc, list_dv(f"{SF}!$B${FUND_FIRST}:$B${FUND_LAST}", "Fund",
                   "Pick a fund. To add a new fund, type it on the Sinking Funds tab first."),
       f"C{CON_FIRST}:C{CON_LAST}")
add_dv(wc, money_dv(allow_negative=True), f"D{CON_FIRST}:D{CON_LAST}")
add_dv(wc, list_dv(f"{EC}!$I${PAY_FIRST}:$I${PAY_LAST}", "Payment method",
                   "Pick a payment method (edit the list on the Expense Categories tab)."),
       f"E{CON_FIRST}:E{CON_LAST}")
# side panel: totals by fund
wc["I4"] = "TOTALS BY FUND"
wc["I4"].font = font(9, True, OLIVE)
wc["I4"].fill = fill(CREAM)
table_header(wc, 5, [("I", "Fund"), ("J", "Total Contributed"), ("K", "# Entries"), ("L", "Last Deposit")],
             {})
for i, r in enumerate(range(FUND_FIRST, FUND_LAST + 1)):
    wc[f"I{r}"] = f'=IF({SF}!B{r}="","",{SF}!B{r})'
    wc[f"J{r}"] = f'=IF(I{r}="","",SUMIFS($D${CON_FIRST}:$D${CON_LAST},$C${CON_FIRST}:$C${CON_LAST},I{r}))'
    wc[f"K{r}"] = f'=IF(I{r}="","",COUNTIF($C${CON_FIRST}:$C${CON_LAST},I{r}))'
    wc[f"L{r}"] = (f'=IF(I{r}="","",IF(K{r}=0,"–",_xlfn.MAXIFS($B${CON_FIRST}:$B${CON_LAST},'
                   f'$C${CON_FIRST}:$C${CON_LAST},I{r})))')
    even = i % 2 == 0
    for col, fmt in (("I", None), ("J", CUR), ("K", "0"), ("L", DATEF)):
        c = wc[f"{col}{r}"]
        c.fill = fill(SAGE_PALE if even else SAGE_LIGHT)
        c.font = font(10, col == "I", OLIVE)
        c.border = ROWLINE
        c.alignment = LEFT if col == "I" else (RIGHT if col == "J" else CENTER)
        if fmt:
            c.number_format = fmt
tr2 = FUND_LAST + 1
wc[f"I{tr2}"] = "TOTAL SAVED"
wc[f"J{tr2}"] = f"=SUM(D{CON_FIRST}:D{CON_LAST})"
wc[f"K{tr2}"] = f"=COUNT(D{CON_FIRST}:D{CON_LAST})"
for col in "IJKL":
    c = wc[f"{col}{tr2}"]
    c.fill = fill(BEIGE)
    c.font = font(10, True)
    c.border = Border(top=Side(style="medium", color=OLIVE))
wc[f"J{tr2}"].number_format = CUR
wc[f"K{tr2}"].alignment = CENTER
wc[f"I{tr2 + 1}"] = "Unassigned / misspelled fund entries:"
wc[f"I{tr2 + 1}"].font = font(8.5, False, GREY, italic=True)
wc[f"J{tr2 + 1}"] = f"=J{tr2}-SUM(J{FUND_FIRST}:J{FUND_LAST})"
wc[f"J{tr2 + 1}"].number_format = CUR
wc[f"J{tr2 + 1}"].font = font(8.5, True, ROSE_TXT)
wc.freeze_panes = "B6"
wc.sheet_properties.tabColor = OLIVE

# ====================================================================== MONTHLY PLANNER
wp = wb.create_sheet("Monthly Planner")
sheet_header(wp, "Monthly Savings Planner",
             "Plan what you'll set aside each month per fund. “Actual” is pulled from the Contribution "
             "Tracker automatically.  ★ “Example” rows are sample data.", "R")
for k, v in {"B": 13, "C": 26, "D": 14, "E": 14, "F": 13, "G": 16, "H": 11, "I": 14, "J": 3,
             "K": 11, "L": 14, "M": 16, "N": 13, "O": 11, "P": 14}.items():
    wp.column_dimensions[k].width = v
pk = {"B": "in", "C": "in", "D": "in", "E": "auto", "F": "auto", "G": "auto", "H": "auto", "I": "auto"}
table_header(wp, 5, [("B", "Month"), ("C", "Fund Name"), ("D", "Planned Contribution"),
                     ("E", "Actual Contribution"), ("F", "Difference"), ("G", "Monthly Savings Total"),
                     ("H", "Monthly Progress"), ("I", "Progress Bar")], pk)
band(wp, PLAN_FIRST, PLAN_LAST, "BCDEFGHI", pk)
plan_rows = []
start_month = {sch[0]: sch[1] for sch in schedules}
for m in range(1, 13):
    for n, cat, goal, mon, td, pri, note in funds:
        if m < start_month.get(n, 9) or (n == "Christmas 2026" and m == 12):
            continue
        plan_rows.append((dt.date(2026, m, 1), n, mon))


def month_crit(rng_ref, cell):
    return (f'{rng_ref},">="&DATE(YEAR({cell}),MONTH({cell}),1),'
            f'{rng_ref},"<"&DATE(YEAR({cell}),MONTH({cell})+1,1)')


for i, r in enumerate(range(PLAN_FIRST, PLAN_LAST + 1)):
    if i < len(plan_rows):
        wp[f"B{r}"], wp[f"C{r}"], wp[f"D{r}"] = plan_rows[i]
    wp[f"E{r}"] = f'=IF(OR(B{r}="",C{r}=""),"",SUMIFS({CON_D},{CON_C},C{r},{month_crit(CON_B, f"B{r}")}))'
    wp[f"F{r}"] = f'=IF(E{r}="","",E{r}-N(D{r}))'
    wp[f"G{r}"] = f'=IF(B{r}="","",SUMIFS({CON_D},{month_crit(CON_B, f"B{r}")}))'
    wp[f"H{r}"] = f'=IF(OR(E{r}="",N(D{r})=0),"",E{r}/D{r})'
    wp[f"I{r}"] = f'=IF(H{r}="","",{bar_formula(f"H{r}")})'
    wp[f"B{r}"].number_format = "mmm yyyy"
    wp[f"B{r}"].alignment = CENTER
    wp[f"C{r}"].alignment = LEFT
    for col in "DEFG":
        wp[f"{col}{r}"].number_format = CUR if col != "F" else '+$#,##0.00;-$#,##0.00;$0.00'
        wp[f"{col}{r}"].alignment = RIGHT
    wp[f"H{r}"].number_format = PCT
    wp[f"H{r}"].alignment = CENTER
    wp[f"I{r}"].font = Font(name="Arial", size=10, color=SAGE_DEEP)
    wp[f"I{r}"].alignment = CENTER
add_dv(wp, date_dv(), f"B{PLAN_FIRST}:B{PLAN_LAST}")
add_dv(wp, list_dv(f"{SF}!$B${FUND_FIRST}:$B${FUND_LAST}", "Fund", "Pick a fund from your Sinking Funds list."),
       f"C{PLAN_FIRST}:C{PLAN_LAST}")
add_dv(wp, money_dv(), f"D{PLAN_FIRST}:D{PLAN_LAST}")
prng = f"F{PLAN_FIRST}:F{PLAN_LAST}"
wp.conditional_formatting.add(prng, FormulaRule(formula=[f'AND(F{PLAN_FIRST}<>"",F{PLAN_FIRST}<0)'],
                                                font=Font(name=FONT, bold=True, color=ROSE_TXT)))
wp.conditional_formatting.add(prng, FormulaRule(formula=[f'AND(F{PLAN_FIRST}<>"",F{PLAN_FIRST}>=0)'],
                                                font=Font(name=FONT, bold=True, color=SAGE_DEEP)))
pct_scale(wp, f"H{PLAN_FIRST}:H{PLAN_LAST}")
# monthly summary
wp["K3"] = "Summary year ✎"
wp["K3"].font = font(9, True, OLIVE)
wp["K3"].alignment = RIGHT
wp.merge_cells("K3:L3")
wp["M3"] = 2026
wp["M3"].font = font(12, True, CHARCOAL)
wp["M3"].fill = fill(WHITE)
wp["M3"].border = Border(left=Side(style="thin", color=OLIVE), right=Side(style="thin", color=OLIVE),
                         top=Side(style="thin", color=OLIVE), bottom=Side(style="thin", color=OLIVE))
wp["M3"].alignment = CENTER
yr_dv = DataValidation(type="whole", operator="between", formula1="2000", formula2="2100",
                       errorStyle="stop", errorTitle="Year", error="Enter a year such as 2027.",
                       promptTitle="Year", prompt="Type the year to summarise, e.g. 2027")
add_dv(wp, yr_dv, "M3")
wp["K4"] = "MONTHLY SUMMARY"
wp["K4"].font = font(9, True, OLIVE)
table_header(wp, 5, [("K", "Month"), ("L", "Planned Total"), ("M", "Monthly Savings Total"),
                     ("N", "Difference"), ("O", "Monthly Progress"), ("P", "Progress Bar")], {})
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
PL_D = f"$D${PLAN_FIRST}:$D${PLAN_LAST}"
PL_B = f"$B${PLAN_FIRST}:$B${PLAN_LAST}"
for i, mname in enumerate(months):
    r = 6 + i
    m = i + 1
    wp[f"K{r}"] = mname
    wp[f"L{r}"] = f'=SUMIFS({PL_D},{PL_B},">="&DATE($M$3,{m},1),{PL_B},"<"&DATE($M$3,{m + 1},1))'
    wp[f"M{r}"] = f'=SUMIFS({CON_D},{CON_B},">="&DATE($M$3,{m},1),{CON_B},"<"&DATE($M$3,{m + 1},1))'
    wp[f"N{r}"] = f"=M{r}-L{r}"
    wp[f"O{r}"] = f'=IF(L{r}=0,"",M{r}/L{r})'
    wp[f"P{r}"] = f'=IF(O{r}="","",{bar_formula(f"O{r}")})'
    even = i % 2 == 0
    for col in "KLMNOP":
        c = wp[f"{col}{r}"]
        c.fill = fill(SAGE_PALE if even else SAGE_LIGHT)
        c.font = font(10, col == "K", OLIVE)
        c.border = ROWLINE
        c.alignment = RIGHT
    wp[f"K{r}"].alignment = CENTER
    for col in "LM":
        wp[f"{col}{r}"].number_format = CUR
    wp[f"N{r}"].number_format = '+$#,##0.00;-$#,##0.00;$0.00'
    wp[f"O{r}"].number_format = PCT
    wp[f"O{r}"].alignment = CENTER
    wp[f"P{r}"].font = Font(name="Arial", size=10, color=SAGE_DEEP)
    wp[f"P{r}"].alignment = CENTER
r = 18
wp[f"K{r}"] = "YEAR"
wp[f"L{r}"] = "=SUM(L6:L17)"
wp[f"M{r}"] = "=SUM(M6:M17)"
wp[f"N{r}"] = f"=M{r}-L{r}"
wp[f"O{r}"] = f'=IF(L{r}=0,"",M{r}/L{r})'
wp[f"P{r}"] = f'=IF(O{r}="","",{bar_formula(f"O{r}")})'
for col in "KLMNOP":
    c = wp[f"{col}{r}"]
    c.fill = fill(BEIGE)
    c.font = font(10, True)
    c.border = Border(top=Side(style="medium", color=OLIVE))
    c.alignment = RIGHT
wp[f"K{r}"].alignment = CENTER
wp[f"L{r}"].number_format = CUR
wp[f"M{r}"].number_format = CUR
wp[f"N{r}"].number_format = '+$#,##0.00;-$#,##0.00;$0.00'
wp[f"O{r}"].number_format = PCT
wp[f"P{r}"].font = Font(name="Arial", size=10, bold=True, color=OLIVE)
wp.conditional_formatting.add("N6:N18", FormulaRule(formula=["N6<0"], font=Font(name=FONT, bold=True, color=ROSE_TXT)))
wp.conditional_formatting.add("N6:N18", FormulaRule(formula=["N6>0"], font=Font(name=FONT, bold=True, color=SAGE_DEEP)))
wp["K20"] = ("Monthly Savings Total = every contribution logged in that month (all funds), "
             "so it also counts deposits you didn't plan for.")
wp["K20"].font = font(8.5, False, GREY, italic=True)
wp["K20"].alignment = Alignment(wrap_text=True, vertical="top")
wp.merge_cells("K20:P22")
wp.freeze_panes = "B6"
wp.sheet_properties.tabColor = BEIGE_DEEP

# ====================================================================== SAVINGS GOALS
wg = wb.create_sheet("Savings Goals")
sheet_header(wg, "Savings Goals",
             "Bigger, longer-term goals you track by balance (house deposit, car, renovation). "
             "Update “Amount Saved” whenever you add money.  ★ “Example” rows are sample data.", "M")
for k, v in {"B": 26, "C": 15, "D": 15, "E": 15, "F": 14, "G": 16, "H": 15, "I": 15, "J": 11,
             "K": 15, "L": 32}.items():
    wg.column_dimensions[k].width = v
gk = {c: "in" for c in "BCDEFL"}
gk.update({c: "auto" for c in "GHIJK"})
table_header(wg, 5, [("B", "Goal Name"), ("C", "Target Amount"), ("D", "Starting Balance"),
                     ("E", "Amount Saved"), ("F", "Target Date"), ("G", "Monthly Savings Needed"),
                     ("H", "Remaining Amount"), ("I", "Completion Status"), ("J", "Progress %"),
                     ("K", "Progress Bar"), ("L", "Notes")], gk)
band(wg, GOAL_FIRST, GOAL_LAST, "BCDEFGHIJKL", gk)
for i, r in enumerate(range(GOAL_FIRST, GOAL_LAST + 1)):
    if i < len(goals):
        n, t, s, a, d, note = goals[i]
        wg[f"B{r}"], wg[f"C{r}"], wg[f"D{r}"], wg[f"E{r}"], wg[f"F{r}"], wg[f"L{r}"] = n, t, s, a, d, note
    wg[f"H{r}"] = f'=IF(B{r}="","",MAX(0,N(C{r})-N(D{r})-N(E{r})))'
    wg[f"G{r}"] = (f'=IF(OR(B{r}="",F{r}=""),"",IF(H{r}=0,0,H{r}/MAX(1,'
                   f'(YEAR(F{r})-YEAR(TODAY()))*12+MONTH(F{r})-MONTH(TODAY()))))')
    wg[f"J{r}"] = f'=IF(OR(B{r}="",N(C{r})=0),"",MIN(1,(N(D{r})+N(E{r}))/C{r}))'
    wg[f"I{r}"] = (f'=IF(B{r}="","",IF(AND(N(C{r})>0,H{r}=0),"Completed",'
                   f'IF(AND(F{r}<>"",F{r}<TODAY()),"Overdue",IF(N(D{r})+N(E{r})<=0,"Not Started",'
                   f'IF(AND(F{r}<>"",F{r}-TODAY()<=60),"Due Soon","In Progress")))))')
    wg[f"K{r}"] = f'=IF(J{r}="","",{bar_formula(f"J{r}")})'
    for col in "CDEGH":
        wg[f"{col}{r}"].number_format = CUR
        wg[f"{col}{r}"].alignment = RIGHT
    wg[f"F{r}"].number_format = DATEF
    wg[f"F{r}"].alignment = CENTER
    wg[f"I{r}"].alignment = CENTER
    wg[f"J{r}"].number_format = PCT
    wg[f"J{r}"].alignment = CENTER
    wg[f"K{r}"].font = Font(name="Arial", size=10, color=SAGE_DEEP)
    wg[f"K{r}"].alignment = CENTER
    wg[f"B{r}"].font = font(10, True)
    wg[f"B{r}"].alignment = LEFT
    wg[f"L{r}"].font = font(9, False, GREY, italic=True)
tr3 = GOAL_LAST + 1
wg[f"B{tr3}"] = "TOTAL"
for col in "BCDEFGHIJKL":
    c = wg[f"{col}{tr3}"]
    c.fill = fill(BEIGE)
    c.font = font(10, True)
    c.border = Border(top=Side(style="medium", color=OLIVE))
    c.alignment = RIGHT
wg[f"B{tr3}"].alignment = LEFT
for col in "CDEGH":
    wg[f"{col}{tr3}"] = f"=SUM({col}{GOAL_FIRST}:{col}{GOAL_LAST})"
    wg[f"{col}{tr3}"].number_format = CUR
wg[f"J{tr3}"] = f"=IFERROR((D{tr3}+E{tr3})/C{tr3},0)"
wg[f"J{tr3}"].number_format = PCT
wg[f"J{tr3}"].alignment = CENTER
wg[f"K{tr3}"] = f'={bar_formula(f"J{tr3}")}'
wg[f"K{tr3}"].font = Font(name="Arial", size=10, bold=True, color=OLIVE)
wg[f"K{tr3}"].alignment = CENTER
g = lambda col: f"{col}{GOAL_FIRST}:{col}{GOAL_LAST}"
for col in "CDE":
    add_dv(wg, money_dv(), g(col))
add_dv(wg, date_dv(), g("F"))
status_cf(wg, g("I"), "I", GOAL_FIRST)
pct_scale(wg, g("J"))
wg.conditional_formatting.add(g("F"), FormulaRule(
    formula=[f'AND($F{GOAL_FIRST}<>"",$F{GOAL_FIRST}<TODAY(),$I{GOAL_FIRST}<>"Completed")'],
    font=Font(name=FONT, bold=True, color=ROSE_TXT)))
wg.freeze_panes = "C6"
wg.sheet_properties.tabColor = SAGE_DEEP

# ====================================================================== EXPENSE CATEGORIES
we = wb.create_sheet("Expense Categories")
sheet_header(we, "Expense Categories & Lists",
             "Rename, add or delete categories and payment methods here – every dropdown in the "
             "workbook updates instantly.", "M")
for k, v in {"B": 22, "C": 42, "D": 10, "E": 15, "F": 15, "G": 11, "H": 3, "I": 22, "J": 3, "K": 34,
             "L": 3}.items():
    we.column_dimensions[k].width = v
ek = {"B": "in", "C": "in", "D": "auto", "E": "auto", "F": "auto", "G": "auto", "I": "in"}
table_header(we, 5, [("B", "Category"), ("C", "Description"), ("D", "# Funds"), ("E", "Total Goals"),
                     ("F", "Total Saved"), ("G", "Share"), ("I", "Payment Methods")], ek)
band(we, CAT_FIRST, CAT_LAST, "BCDEFG", ek)
band(we, PAY_FIRST, PAY_LAST, "I", ek)
SF_C = f"{SF}!$C${FUND_FIRST}:$C${FUND_LAST}"
for i, r in enumerate(range(CAT_FIRST, CAT_LAST + 1)):
    if i < len(categories):
        we[f"B{r}"], we[f"C{r}"] = categories[i]
    we[f"D{r}"] = f'=IF(B{r}="","",COUNTIF({SF_C},B{r}))'
    we[f"E{r}"] = f'=IF(B{r}="","",SUMIFS({SF}!$D${FUND_FIRST}:$D${FUND_LAST},{SF_C},B{r}))'
    we[f"F{r}"] = f'=IF(B{r}="","",SUMIFS({SF}!$E${FUND_FIRST}:$E${FUND_LAST},{SF_C},B{r}))'
    we[f"G{r}"] = f'=IF(OR(B{r}="",SUM($F${CAT_FIRST}:$F${CAT_LAST})=0),"",F{r}/SUM($F${CAT_FIRST}:$F${CAT_LAST}))'
    we[f"B{r}"].font = font(10, True)
    we[f"B{r}"].alignment = LEFT
    we[f"C{r}"].font = font(9, False, GREY, italic=True)
    we[f"D{r}"].alignment = CENTER
    for col in "EF":
        we[f"{col}{r}"].number_format = CUR
        we[f"{col}{r}"].alignment = RIGHT
    we[f"G{r}"].number_format = PCT
    we[f"G{r}"].alignment = CENTER
for i, r in enumerate(range(PAY_FIRST, PAY_LAST + 1)):
    if i < len(payment_methods):
        we[f"I{r}"] = payment_methods[i]
    we[f"I{r}"].alignment = LEFT
pct_scale(we, f"G{CAT_FIRST}:G{CAT_LAST}")
cat_uniq = DataValidation(type="custom", formula1=f"COUNTIF($B${CAT_FIRST}:$B${CAT_LAST},B{CAT_FIRST})=1",
                          allow_blank=True, errorStyle="stop", errorTitle="Duplicate",
                          error="This category already exists.")
add_dv(we, cat_uniq, f"B{CAT_FIRST}:B{CAT_LAST}")
# status legend
we["K5"] = "Status Guide"
we["K5"].font = font(9.5, True, CREAM)
we["K5"].fill = fill(OLIVE)
we["K5"].alignment = CENTER
legend = [("Completed", "Saved amount has reached the goal", SAGE, "FFFFFF"),
          ("In Progress", "Saving, target date more than 60 days away", SAGE_LIGHT, OLIVE),
          ("Due Soon", "Target date within the next 60 days", AMBER, AMBER_TXT),
          ("Overdue", "Target date passed, goal not yet reached", ROSE, ROSE_TXT),
          ("Not Started", "Nothing saved yet", STONE, GREY)]
for i, (s, desc, bg, fg) in enumerate(legend):
    r = 6 + i * 2
    we[f"K{r}"] = s
    we[f"K{r}"].fill = fill(bg)
    we[f"K{r}"].font = font(10, True, fg)
    we[f"K{r}"].alignment = CENTER
    we[f"K{r + 1}"] = desc
    we[f"K{r + 1}"].font = font(8.5, False, GREY, italic=True)
    we[f"K{r + 1}"].alignment = CENTER
we.freeze_panes = "B6"
we.sheet_properties.tabColor = BEIGE

# ====================================================================== CHART DATA (helper)
wx = wb.create_sheet("Chart Data")
wx.sheet_view.showGridLines = False
wx["A1"] = "⚙ Chart Data – helper tables that feed the Dashboard charts. No need to edit anything here."
wx["A1"].font = font(10, True, OLIVE)
hdr = ["#", "Row", "Fund", "Saved", "Goal", "Remaining", "Progress", "Bar", "Status", "Category"]
for j, h in enumerate(hdr):
    c = wx.cell(row=3, column=1 + j, value=h)
    c.font = font(9, True, CREAM)
    c.fill = fill(OLIVE)
for k in range(1, 21):
    r = 3 + k
    wx[f"A{r}"] = k
    wx[f"B{r}"] = f'=IFERROR(SMALL({SF}!$P${FUND_FIRST}:$P${FUND_LAST},A{r}),"")'
    idx = lambda col: f'IF($B{r}="","",INDEX({SF}!${col}:${col},$B{r}))'
    wx[f"C{r}"] = f"={idx('B')}"
    wx[f"D{r}"] = f'=IF($B{r}="",0,N(INDEX({SF}!$E:$E,$B{r})))'
    wx[f"E{r}"] = f'=IF($B{r}="",0,N(INDEX({SF}!$D:$D,$B{r})))'
    wx[f"F{r}"] = f'=IF($B{r}="",0,N(INDEX({SF}!$F:$F,$B{r})))'
    wx[f"G{r}"] = f"={idx('I')}"
    wx[f"H{r}"] = f"={idx('J')}"
    wx[f"I{r}"] = f"={idx('K')}"
    wx[f"J{r}"] = f"={idx('C')}"
    for col in "DEF":
        wx[f"{col}{r}"].number_format = CUR
    wx[f"G{r}"].number_format = PCT
# top categories (by saved) – L:N
for j, h in enumerate(["Rank", "Category", "Saved", "key", "row"]):
    c = wx.cell(row=3, column=12 + j, value=h)
    c.font = font(9, True, CREAM)
    c.fill = fill(OLIVE)
for i, r in enumerate(range(CAT_FIRST, CAT_LAST + 1)):
    xr = 4 + i
    wx[f"O{xr}"] = f'=IF(OR({EC}!B{r}="",N({EC}!F{r})<=0),"",{EC}!F{r}+ROW({EC}!F{r})/1000000000)'
for k in range(1, 9):
    r = 3 + k
    wx[f"L{r}"] = k
    wx[f"P{r}"] = f'=IFERROR(MATCH(LARGE($O$4:$O$28,L{r}),$O$4:$O$28,0),"")'
    wx[f"M{r}"] = f'=IF(P{r}="","",INDEX({EC}!$B${CAT_FIRST}:$B${CAT_LAST},P{r}))'
    wx[f"N{r}"] = f'=IF(P{r}="",0,INDEX({EC}!$F${CAT_FIRST}:$F${CAT_LAST},P{r}))'
    wx[f"N{r}"].number_format = CUR
# upcoming (Q:R)
for j, h in enumerate(["#", "Key", "Row"]):
    c = wx.cell(row=3, column=17 + j, value=h)
    c.font = font(9, True, CREAM)
    c.fill = fill(OLIVE)
for k in range(1, 6):
    r = 3 + k
    wx[f"Q{r}"] = k
    wx[f"R{r}"] = f'=IFERROR(SMALL({SF}!$Q${FUND_FIRST}:$Q${FUND_LAST},Q{r}),"")'
    wx[f"S{r}"] = f'=IF(R{r}="","",ROW({SF}!$Q${FUND_FIRST})-1+MATCH(R{r},{SF}!$Q${FUND_FIRST}:$Q${FUND_LAST},0))'
for col, w in zip("ABCDEFGHIJKLMNOPQRS", [4, 5, 24, 12, 12, 12, 9, 14, 12, 18, 3, 6, 20, 12, 14, 3, 4, 12, 6]):
    wx.column_dimensions[col].width = w
wx.sheet_properties.tabColor = "BFBFBF"

# ====================================================================== DASHBOARD build
d = wsd
d.sheet_view.showGridLines = False
d.sheet_view.zoomScale = 90
DCOLS = "ABCDEFGHIJKLMNOPQ"
dw = {"A": 2.5, "B": 13, "C": 13, "D": 13, "E": 2.5, "F": 13, "G": 13, "H": 13, "I": 2.5,
      "J": 13, "K": 13, "L": 13, "M": 2.5, "N": 13, "O": 13, "P": 13, "Q": 2.5}
for k, v in dw.items():
    d.column_dimensions[k].width = v
paint(d, "A1:Q112", CREAM)
d.row_dimensions[1].height = 12
d.row_dimensions[2].height = 40
d.row_dimensions[3].height = 20
d["B2"] = "Sinking Funds Tracker"
d["B2"].font = font(26, True, CHARCOAL)
d["B2"].alignment = Alignment(vertical="center")
d["B3"] = "Plan ahead · save calmly · spend guilt-free"
d["B3"].font = font(10, False, OLIVE, italic=True)
d["N2"] = "TODAY"
d["N2"].font = font(8, True, OLIVE)
d["N2"].alignment = Alignment(horizontal="right", vertical="bottom")
d.merge_cells("O2:P2")
d["O2"] = "=TODAY()"
d["O2"].number_format = "dddd, mmm d, yyyy"
d["O2"].font = font(10, True, CHARCOAL)
d["O2"].alignment = Alignment(horizontal="right", vertical="bottom")
d.merge_cells("N3:P3")
d["N3"] = '=IF(COUNTIF(' + f"{SF}!M{FUND_FIRST}:M{FUND_LAST}" + ',"Example*")>0,"★ Sample data loaded – see Start Here to clear it","")'
d["N3"].font = font(8, False, ROSE_TXT, italic=True)
d["N3"].alignment = Alignment(horizontal="right")
for c in "BCDEFGHIJKLMNOP":
    d[f"{c}4"].border = Border(bottom=Side(style="medium", color=SAGE))


def section(row, text, col="B", end="P"):
    d.row_dimensions[row].height = 26
    d[f"{col}{row}"] = text
    d[f"{col}{row}"].font = font(10, True, OLIVE)
    d[f"{col}{row}"].alignment = Alignment(vertical="bottom")
    for cc in DCOLS[DCOLS.index(col):DCOLS.index(end) + 1]:
        d[f"{cc}{row}"].border = Border(bottom=Side(style="thin", color=BEIGE_DEEP))


def card(col0, row0, label, formula, fmt, sub_formula, accent=SAGE):
    cols = DCOLS[DCOLS.index(col0):DCOLS.index(col0) + 3]
    c1, c3 = cols[0], cols[-1]
    for rr in range(row0, row0 + 3):
        d.merge_cells(f"{c1}{rr}:{c3}{rr}")
        for cc in cols:
            cell = d[f"{cc}{rr}"]
            cell.fill = fill(WHITE)
            left = Side(style="thin", color=BEIGE) if cc == c1 else None
            right = Side(style="thin", color=BEIGE) if cc == c3 else None
            top = Side(style="thick", color=accent) if rr == row0 else None
            bottom = Side(style="thin", color=BEIGE) if rr == row0 + 2 else None
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)
    d[f"{c1}{row0}"] = label
    d[f"{c1}{row0}"].font = font(8.5, True, GREY)
    d[f"{c1}{row0}"].alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    d[f"{c1}{row0 + 1}"] = formula
    d[f"{c1}{row0 + 1}"].number_format = fmt
    d[f"{c1}{row0 + 1}"].font = font(22, True, CHARCOAL)
    d[f"{c1}{row0 + 1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    d[f"{c1}{row0 + 2}"] = sub_formula
    d[f"{c1}{row0 + 2}"].font = font(8.5, False, OLIVE, italic=True)
    d[f"{c1}{row0 + 2}"].alignment = Alignment(horizontal="left", vertical="top", indent=1)


R = lambda col: f"{SF}!{col}{FUND_FIRST}:{col}{FUND_LAST}"
GR = lambda col: f"{SG}!{col}{GOAL_FIRST}:{col}{GOAL_LAST}"
n_funds = f'COUNTIF({R("B")},"?*")'
n_goals = f'COUNTIF({GR("B")},"?*")'
fund_saved = f'SUM({R("E")})'
goal_saved = f'(SUM({GR("D")})+SUM({GR("E")}))'
fund_goal = f'SUM({R("D")})'
goal_goal = f'SUM({GR("C")})'
remaining = f'(SUM({R("F")})+SUM({GR("H")}))'
this_month = (f'SUMIFS({CON_D},{CON_B},">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'
              f'{CON_B},"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))')
target_month = f'(SUM({R("N")})+SUM({GR("G")}))'

section(5, "AT A GLANCE")
for rr, h in ((6, 22), (7, 38), (8, 22), (9, 10), (10, 22), (11, 38), (12, 22)):
    d.row_dimensions[rr].height = h
card("B", 6, "TOTAL SAVINGS BALANCE", f"={fund_saved}+{goal_saved}", CUR0,
     f'="Funds "&TEXT({fund_saved},"$#,##0")&"  ·  Goals "&TEXT({goal_saved},"$#,##0")')
card("F", 6, "TOTAL SAVINGS GOALS", f"={fund_goal}+{goal_goal}", CUR0,
     f'={n_funds}&" sinking funds  ·  "&{n_goals}&" savings goals"', OLIVE)
card("J", 6, "TOTAL SAVED IN SINKING FUNDS", f"={fund_saved}", CUR0,
     f'="from "&COUNT({CON_D})&" logged contributions"', SAGE_DEEP)
card("N", 6, "REMAINING TO SAVE", f"={remaining}", CUR0,
     f'="Funds "&TEXT(SUM({R("F")}),"$#,##0")&"  ·  Goals "&TEXT(SUM({GR("H")}),"$#,##0")', BEIGE_DEEP)
card("B", 10, "OVERALL PROGRESS", f"=IFERROR(1-{remaining}/({fund_goal}+{goal_goal}),0)", "0.0%",
     f'=IFERROR({bar_formula(f"1-{remaining}/({fund_goal}+{goal_goal})")},"")', SAGE)
d["B12"].font = Font(name="Arial", size=10, color=SAGE_DEEP)
card("F", 10, "MONTHLY SAVINGS TARGET", f"={target_month}", CUR0,
     '="needed per month to hit every target date"', OLIVE)
card("J", 10, "COMPLETED FUNDS", f'=COUNTIF({R("K")},"Completed")', "0",
     f'="of "&{n_funds}&" funds  ·  "&COUNTIF({R("K")},"Overdue")&" overdue"', SAGE_DEEP)
card("N", 10, "SAVED THIS MONTH", f"={this_month}", CUR0,
     f'=IFERROR(TEXT({this_month}/{target_month},"0%")&" of monthly target","")', BEIGE_DEEP)

# charts area
section(14, "SAVINGS VISUALS")
for rr in range(15, 52):
    d.row_dimensions[rr].height = 18

CHART_COLORS = ["7F9468", "D6C7A9", "6B705C", "A3B18A", "C2A383", "36393B", "B7C4A3", "8B7E74",
                "E4D9C3", "556043", "CFB997", "9AA58A", "A89C8E", "DCE3D0", "766A5D", "BFAE8C",
                "5E6B4C", "E9E0D2", "8FA07A", "4A4D4F"]


def txt_props(size=900, color=CHARCOAL, bold=False):
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(
        sz=size, b=bold, solidFill=color, latin=DFont(typeface=FONT))), endParaRPr=CharacterProperties())])


def style_chart(ch, title):
    ch.title = title
    ch.roundedCorners = False
    ch.style = 2
    ch.title.overlay = False
    ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=CharacterProperties(
        sz=1200, b=True, solidFill=CHARCOAL, latin=DFont(typeface=FONT)))
    ch.graphical_properties = GraphicalProperties(solidFill=WHITE, ln=LineProperties(solidFill=BEIGE))
    if ch.legend is not None:
        ch.legend.overlay = False
        ch.legend.txPr = txt_props(850, CHARCOAL)


N = 20          # rows in the Fund Progress table
NC = 15         # funds plotted in the charts (keeps legends tidy; extend in the chart editor if needed)
cats = Reference(wx, min_col=3, min_row=4, max_row=3 + NC)
# 1. Donut – savings by fund
donut = DoughnutChart(holeSize=55)
donut.add_data(Reference(wx, min_col=4, min_row=3, max_row=3 + NC), titles_from_data=True)
donut.set_categories(cats)
style_chart(donut, "Savings by Fund")
s = donut.series[0]
for i in range(NC):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = CHART_COLORS[i]
    pt.graphicalProperties.line.solidFill = WHITE
    s.dPt.append(pt)
donut.legend.position = "r"
donut.width, donut.height = 15.6, 8.6
d.add_chart(donut, "B15")

# 2. Bar – goal vs saved
bar = BarChart()
bar.type = "bar"
bar.grouping = "clustered"
bar.add_data(Reference(wx, min_col=5, min_row=3, max_row=3 + NC), titles_from_data=True)
bar.add_data(Reference(wx, min_col=4, min_row=3, max_row=3 + NC), titles_from_data=True)
bar.set_categories(cats)
style_chart(bar, "Savings Goal vs Current Saved")
bar.series[0].graphicalProperties.solidFill = BEIGE_DEEP
bar.series[0].graphicalProperties.line.solidFill = BEIGE_DEEP
bar.series[1].graphicalProperties.solidFill = SAGE_DEEP
bar.series[1].graphicalProperties.line.solidFill = SAGE_DEEP
bar.gapWidth = 60
bar.overlap = -10
bar.x_axis.delete = False
bar.y_axis.delete = False
bar.x_axis.scaling.orientation = "maxMin"
bar.y_axis.numFmt = "$#,##0"
bar.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=BEIGE))
bar.x_axis.txPr = txt_props(800)
bar.y_axis.txPr = txt_props(800, GREY)
bar.legend.position = "b"
bar.width, bar.height = 15.6, 8.6
d.add_chart(bar, "J15")

# 3. Line – monthly trend
line = LineChart()
line.add_data(Reference(wp, min_col=12, min_row=5, max_row=17), titles_from_data=True)
line.add_data(Reference(wp, min_col=13, min_row=5, max_row=17), titles_from_data=True)
line.set_categories(Reference(wp, min_col=11, min_row=6, max_row=17))
style_chart(line, "Monthly Contributions – Planned vs Actual")
line.series[0].graphicalProperties.line.solidFill = BEIGE_DEEP
line.series[0].graphicalProperties.line.width = 28000
line.series[0].graphicalProperties.line.dashStyle = "dash"
line.series[0].marker.symbol = "none"
line.series[1].graphicalProperties.line.solidFill = OLIVE
line.series[1].graphicalProperties.line.width = 32000
line.series[1].marker.symbol = "circle"
line.series[1].marker.size = 6
line.series[1].marker.graphicalProperties = GraphicalProperties(solidFill=SAGE_DEEP)
line.series[1].marker.graphicalProperties.line.solidFill = WHITE
line.series[0].smooth = False
line.series[1].smooth = False
line.x_axis.delete = False
line.y_axis.delete = False
line.y_axis.numFmt = "$#,##0"
line.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=BEIGE))
line.x_axis.txPr = txt_props(800)
line.y_axis.txPr = txt_props(800, GREY)
line.legend.position = "b"
line.width, line.height = 15.6, 8.6
d.add_chart(line, "B33")

# 4. Column – savings distribution by category
col = BarChart()
col.type = "col"
col.add_data(Reference(wx, min_col=14, min_row=3, max_row=11), titles_from_data=True)
col.set_categories(Reference(wx, min_col=13, min_row=4, max_row=11))
style_chart(col, "Savings Distribution – Top Categories")
col.legend = None
col.series[0].graphicalProperties.solidFill = SAGE
col.series[0].graphicalProperties.line.solidFill = SAGE
col.gapWidth = 50
col.x_axis.delete = False
col.y_axis.delete = False
col.y_axis.numFmt = "$#,##0"
col.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=BEIGE))
col.x_axis.txPr = txt_props(750)
col.y_axis.txPr = txt_props(800, GREY)
col.dataLabels = DataLabelList()
col.dataLabels.showVal = True
col.dataLabels.showCatName = False
col.dataLabels.showSerName = False
col.dataLabels.showLegendKey = False
col.dataLabels.showPercent = False
col.dataLabels.numFmt = "$#,##0"
col.dataLabels.txPr = txt_props(750, OLIVE, True)
col.width, col.height = 15.6, 8.6
d.add_chart(col, "J33")

# fund progress table
P0 = 52
section(P0, "FUND PROGRESS")
hr = P0 + 1
d.row_dimensions[hr].height = 24
pcols = [("B", "D", "Fund"), ("F", "G", "Category"), ("H", "H", "Goal"), ("J", "J", "Saved"),
         ("K", "K", "Remaining"), ("L", "L", "Progress"), ("N", "O", "Progress Bar"), ("P", "P", "Status")]


def table_row(r, specs, fills, hdr_row=False):
    for cc in "BCDEFGHIJKLMNOP":
        cell = d[f"{cc}{r}"]
        cell.fill = fill(fills)
        cell.border = Border(bottom=Side(style="thin", color=BEIGE))
    for a, b, _ in specs:
        if a != b:
            d.merge_cells(f"{a}{r}:{b}{r}")


table_row(hr, pcols, OLIVE)
for a, b, t in pcols:
    d[f"{a}{hr}"] = t
    d[f"{a}{hr}"].font = font(9, True, CREAM)
    d[f"{a}{hr}"].alignment = CENTER
src = {"B": "C", "F": "J", "H": "E", "J": "D", "K": "F", "L": "G", "N": "H", "P": "I"}
for k in range(N):
    r = hr + 1 + k
    xr = 4 + k
    d.row_dimensions[r].height = 20
    table_row(r, pcols, WHITE if k % 2 == 0 else CREAM_2)
    for a, b, _ in pcols:
        sc = src[a]
        if sc in "DEF":
            d[f"{a}{r}"] = f'=IF({CD}!$C{xr}="","",{CD}!{sc}{xr})'
        else:
            d[f"{a}{r}"] = f"={CD}!{sc}{xr}"
        cell = d[f"{a}{r}"]
        cell.font = font(10, a == "B", CHARCOAL if a == "B" else OLIVE)
        cell.alignment = LEFT if a in "BF" else CENTER
        if sc in "DEF":
            cell.number_format = CUR0
            cell.alignment = RIGHT
    d[f"L{r}"].number_format = PCT
    d[f"N{r}"].font = Font(name="Arial", size=10, color=SAGE_DEEP)
last_p = hr + N
status_cf(d, f"P{hr + 1}:P{last_p}", "P", hr + 1)
pct_scale(d, f"L{hr + 1}:L{last_p}")

# upcoming + monthly overview
U0 = last_p + 2
section(U0, "UPCOMING TARGET DATES", "B", "H")
section(U0, "MONTHLY CONTRIBUTION OVERVIEW", "J", "P")
d[f"J{U0}"] = '="MONTHLY CONTRIBUTION OVERVIEW · "&' + f"{MP}!M3"
uh = U0 + 1
ucols = [("B", "D", "Fund"), ("F", "F", "Target Date"), ("G", "G", "Days Left"), ("H", "H", "Remaining")]
mcols = [("J", "J", "Month"), ("K", "K", "Planned"), ("L", "L", "Actual"), ("N", "N", "Difference"),
         ("O", "O", "Progress"), ("P", "P", "Bar")]


def mini_row(r, specs, color, cols):
    for cc in cols:
        d[f"{cc}{r}"].fill = fill(color)
        d[f"{cc}{r}"].border = Border(bottom=Side(style="thin", color=BEIGE))
    for a, b, _ in specs:
        if a != b:
            d.merge_cells(f"{a}{r}:{b}{r}")


mini_row(uh, ucols, OLIVE, "BCDEFGH")
mini_row(uh, mcols, OLIVE, "JKLMNOP")
for a, b, t in ucols + mcols:
    d[f"{a}{uh}"] = t
    d[f"{a}{uh}"].font = font(9, True, CREAM)
    d[f"{a}{uh}"].alignment = CENTER
for k in range(5):
    r = uh + 1 + k
    xr = 4 + k
    mini_row(r, ucols, WHITE if k % 2 == 0 else CREAM_2, "BCDEFGH")
    row_ref = f"{CD}!$S{xr}"
    d[f"B{r}"] = f'=IF({row_ref}="","",INDEX({SF}!$B:$B,{row_ref}))'
    d[f"F{r}"] = f'=IF({row_ref}="","",INDEX({SF}!$H:$H,{row_ref}))'
    d[f"G{r}"] = f'=IF({row_ref}="","",INDEX({SF}!$O:$O,{row_ref}))'
    d[f"H{r}"] = f'=IF({row_ref}="","",INDEX({SF}!$F:$F,{row_ref}))'
    d[f"B{r}"].font = font(10, True)
    d[f"B{r}"].alignment = LEFT
    d[f"F{r}"].number_format = "mmm d, yyyy"
    d[f"G{r}"].number_format = '0" days";[Color9]"overdue "0" d"'
    d[f"H{r}"].number_format = CUR0
    for cc in "FG":
        d[f"{cc}{r}"].alignment = CENTER
        d[f"{cc}{r}"].font = font(10, False, OLIVE)
    d[f"H{r}"].alignment = RIGHT
    d[f"H{r}"].font = font(10, False, OLIVE)
d.conditional_formatting.add(f"G{uh + 1}:G{uh + 5}", FormulaRule(
    formula=[f'AND(G{uh + 1}<>"",G{uh + 1}<0)'], fill=fill(ROSE), font=Font(name=FONT, bold=True, color=ROSE_TXT)))
d.conditional_formatting.add(f"G{uh + 1}:G{uh + 5}", FormulaRule(
    formula=[f'AND(G{uh + 1}<>"",G{uh + 1}>=0,G{uh + 1}<=60)'], fill=fill(AMBER),
    font=Font(name=FONT, bold=True, color=AMBER_TXT)))
d[f"B{uh + 7}"] = "Shows the 5 nearest target dates among funds that aren't completed yet (overdue first)."
d[f"B{uh + 7}"].font = font(8, False, GREY, italic=True)
for k in range(13):
    r = uh + 1 + k
    pr = 6 + k
    mini_row(r, mcols, (WHITE if k % 2 == 0 else CREAM_2) if k < 12 else BEIGE, "JKLMNOP")
    d[f"J{r}"] = f"={MP}!K{pr}"
    d[f"K{r}"] = f"={MP}!L{pr}"
    d[f"L{r}"] = f"={MP}!M{pr}"
    d[f"N{r}"] = f"={MP}!N{pr}"
    d[f"O{r}"] = f"={MP}!O{pr}"
    d[f"P{r}"] = f"={MP}!P{pr}"
    for cc in "JKLNOP":
        d[f"{cc}{r}"].font = font(9.5, k == 12 or cc == "J", OLIVE if cc != "J" else CHARCOAL)
        d[f"{cc}{r}"].alignment = CENTER if cc in "JOP" else RIGHT
    for cc in "KL":
        d[f"{cc}{r}"].number_format = CUR0
    d[f"N{r}"].number_format = '+$#,##0;-$#,##0;$0'
    d[f"O{r}"].number_format = PCT
    d[f"P{r}"].font = Font(name="Arial", size=7, color=SAGE_DEEP)
d.conditional_formatting.add(f"N{uh + 1}:N{uh + 13}", FormulaRule(
    formula=[f"N{uh + 1}<0"], font=Font(name=FONT, bold=True, color=ROSE_TXT)))
foot = uh + 15
d[f"B{foot}"] = ("Every number on this page is calculated from your other tabs – there is nothing to type here. "
                 "Change the summary year on the Monthly Planner tab (cell M3).")
d[f"B{foot}"].font = font(8.5, False, GREY, italic=True)
d.sheet_properties.tabColor = OLIVE
d.print_area = f"A1:Q{foot}"

# ====================================================================== START HERE build
s0 = ws0
s0.sheet_view.showGridLines = False
s0.column_dimensions["A"].width = 3
s0.column_dimensions["B"].width = 4
s0.column_dimensions["C"].width = 30
s0.column_dimensions["D"].width = 80
s0.column_dimensions["E"].width = 3
paint(s0, "A1:F140", CREAM)
s0.row_dimensions[2].height = 44
s0["B2"] = "Sinking Funds Tracker"
s0["B2"].font = font(28, True, CHARCOAL)
s0["B3"] = "START HERE  ·  a 5-minute guide to planning, saving and spending with confidence"
s0["B3"].font = font(10, True, OLIVE)
for cc in "BCD":
    s0[f"{cc}4"].border = Border(bottom=Side(style="medium", color=SAGE))
row = 6


def h(text):
    global row
    row += 1
    s0.row_dimensions[row].height = 26
    s0[f"B{row}"] = text
    s0[f"B{row}"].font = font(13, True, OLIVE)
    s0[f"B{row}"].alignment = Alignment(vertical="bottom")
    for cc in "BCD":
        s0[f"{cc}{row}"].border = Border(bottom=Side(style="thin", color=BEIGE_DEEP))
    row += 1


def step(num, title, body, height=None):
    global row
    s0[f"B{row}"] = num
    s0[f"B{row}"].font = font(10, True, SAGE_DEEP)
    s0[f"B{row}"].alignment = Alignment(horizontal="center", vertical="top")
    s0[f"C{row}"] = title
    s0[f"C{row}"].font = font(10, True, CHARCOAL)
    s0[f"C{row}"].alignment = Alignment(vertical="top", wrap_text=True)
    s0[f"D{row}"] = body
    s0[f"D{row}"].font = font(10, False, CHARCOAL)
    s0[f"D{row}"].alignment = Alignment(vertical="top", wrap_text=True)
    lines = max(1, sum(len(p) // 88 + 1 for p in body.split("\n")))
    s0.row_dimensions[row].height = height or max(20, 15 * lines + 6)
    row += 1


s0["B6"] = ("Welcome! A sinking fund is money you set aside a little at a time for a cost you know is coming – "
            "Christmas, car insurance, a holiday. This planner does all the maths for you: you just name your "
            "funds and log what you save.")
s0["B6"].font = font(10.5, False, CHARCOAL)
s0["B6"].alignment = Alignment(wrap_text=True, vertical="top")
s0.merge_cells("B6:D6")
s0.row_dimensions[6].height = 36

h("1 · Cell colour guide")
s0[f"C{row}"] = "✎ Cream / white cells"
s0[f"C{row}"].fill = fill(WHITE)
s0[f"C{row}"].font = font(10, True, CHARCOAL)
s0[f"C{row}"].border = BOX
s0[f"D{row}"] = "Editable – type, paste or pick from a dropdown. Column headers say ✎ TYPE HERE."
s0[f"D{row}"].font = font(10)
s0.row_dimensions[row].height = 22
row += 1
s0[f"C{row}"] = "ƒ Sage-green cells"
s0[f"C{row}"].fill = fill(SAGE_LIGHT)
s0[f"C{row}"].font = font(10, True, OLIVE)
s0[f"C{row}"].border = BOX
s0[f"D{row}"] = "Automatic formulas – please don't type over them. Column headers say ƒ AUTOMATIC."
s0[f"D{row}"].font = font(10)
s0.row_dimensions[row].height = 22
row += 1
s0[f"C{row}"] = "Dashboard"
s0[f"C{row}"].fill = fill(OLIVE)
s0[f"C{row}"].font = font(10, True, CREAM)
s0[f"C{row}"].border = BOX
s0[f"D{row}"] = "100% automatic – it reads everything from the other tabs. Nothing to type there."
s0[f"D{row}"].font = font(10)
s0.row_dimensions[row].height = 22
row += 1

h("2 · Quick start (first 5 minutes)")
step("①", "Clear the sample data", "Follow section 8 below (or keep the examples for a while to learn how it works).")
step("②", "Check your categories", "Expense Categories tab – rename or add categories and payment methods to suit you.")
step("③", "Add your funds", "Sinking Funds tab – one row per thing you're saving for (see section 3).")
step("④", "Log money as you save it", "Contributions tab – one row per deposit or withdrawal (section 4).")
step("⑤", "Check the Dashboard", "Watch your totals, progress bars and charts update by themselves.")

h("3 · How to create a new sinking fund")
step("1", "Go to the Sinking Funds tab", "Click the first empty cream row under Fund Name (up to 20 funds).")
step("2", "Fund Name", "Type a unique name, e.g. “Christmas 2027”. Duplicate names are blocked so totals stay correct.")
step("3", "Category & Priority", "Pick from the dropdowns. Priority High is shown in rust so it stands out.")
step("4", "Savings Goal", "The total you need, e.g. 800.")
step("5", "Monthly Contribution", "What you plan to put in each month (your own plan).")
step("6", "Target Date", "When you need the money. Used for Status, Days Left and Needed / Month.")
step("7", "Everything else is automatic",
     "Current Saved, Remaining, Progress %, Progress Bar, Status, Needed / Month and Days Left fill in by "
     "themselves.\nNeeded / Month = remaining balance ÷ months left until the target date – compare it with "
     "your Monthly Contribution to see if your plan is enough.")
step("8", "Already have money saved?",
     "Log it once on the Contributions tab with Payment Method “Opening Balance”.")

h("4 · How to record contributions (and spending)")
step("1", "Contributions tab", "Use the next empty row. Rows can be in any order – no need to sort.")
step("2", "Date", "Type a date like 11/15/2026 (or press Ctrl + ; for today).")
step("3", "Fund Name", "Pick from the dropdown – it lists the funds on your Sinking Funds tab.")
step("4", "Contribution Amount",
     "Positive number for money you save (e.g. 100).\nNegative number when you spend from the fund "
     "(e.g. -180 for new tyres). The fund balance goes down automatically.")
step("5", "Payment Method & Notes", "Optional, but handy for your records.")
step("6", "Totals by Fund panel",
     "On the right of the Contributions tab. If “Unassigned / misspelled” is not $0, an entry has a fund "
     "name that doesn't match any fund.")

h("5 · Monthly Savings Planner & Savings Goals")
step("Planner", "Plan each month",
     "Enter a Month (any date in that month, e.g. 12/1/2026), a Fund and the Planned Contribution. "
     "Actual, Difference and progress come from your logged contributions automatically.\nThe Monthly "
     "Summary on the right totals the whole year – change the year in cell M3. It also feeds the Dashboard "
     "trend chart.")
step("Goals", "Long-term goals",
     "Use the Savings Goals tab for big goals you track by balance (house deposit, car). Enter Target, "
     "Starting Balance and update Amount Saved as it grows. Monthly Savings Needed, Remaining and Status "
     "are automatic, and these goals are included in the Dashboard totals.")

h("6 · How to customise categories")
step("Rename", "Rename a category",
     "Type over it on the Expense Categories tab. Tip: also update any funds already using the old name "
     "(they'll show a red warning triangle until you do).")
step("Add", "Add a category / method", "Type it in the next empty cream row – it appears in every dropdown instantly.")
step("Delete", "Remove one", "Select the cell and press Delete. Blank rows are ignored.")
step("Currency", "Use £, €, or another currency",
     "Select the money columns ➜ Format ➜ Number ➜ Custom currency. (In Excel: Format Cells ➜ Currency.)")

h("7 · How to read the Dashboard")
step("Total Savings Balance", "Everything set aside", "Money in all sinking funds + all Savings Goals balances.")
step("Total Savings Goals", "Everything you're aiming for", "Sum of every fund goal and savings goal target.")
step("Total Saved in Funds", "Sinking funds only", "Net of all contributions minus withdrawals.")
step("Remaining to Save", "The gap", "How much is still needed to fully fund everything.")
step("Overall Progress", "Big picture %", "Share of all your goals that is already funded.")
step("Monthly Savings Target", "Your monthly number",
     "What you'd need to save each month to hit every target date on time.")
step("Completed Funds", "Wins!", "Funds where saved ≥ goal, plus how many are overdue.")
step("Saved This Month", "Momentum", "Everything logged in the current calendar month vs your monthly target.")
step("Status colours", "At a glance",
     "Completed (sage) · In Progress (light sage) · Due Soon = target within 60 days (gold) · "
     "Overdue = date passed, not fully funded (rust) · Not Started (grey).")
step("Charts", "Live visuals",
     "Savings by Fund (donut), Goal vs Saved (bars), Monthly Planned vs Actual (line), and Top Categories. "
     "They update as soon as you add funds or contributions. The fund charts plot your first 15 funds (the Fund Progress table shows all 20).")

h("8 · Start fresh – delete the sample data")
step("1", "Sinking Funds tab", "Select cream columns B–D and G–H, L–M in rows 6–25 ➜ press Delete.")
step("2", "Contributions tab", "Select B6:F1005 ➜ press Delete.")
step("3", "Monthly Planner tab", "Select B6:D305 ➜ press Delete. Set your year in M3.")
step("4", "Savings Goals tab", "Select B6:F20 and L6:L20 ➜ press Delete.")
step("!", "Only clear cream cells",
     "Never delete whole rows or the sage formula columns – the formulas are what make the tracker work. "
     "Tip: before clearing, use File ➜ Make a copy to keep an example version.")

h("9 · Sample scenario – how it all fits together")
step("Meet Sam", "The example data",
     "Sam has 11 sinking funds and logs money once a month. Look at these to see each feature in action:")
step("✓", "Christmas 2026", "Goal $800, saved $100/month since April = $800 ➜ Status Completed.")
step("⚠", "Annual Subscriptions", "Target date Oct 31 has passed with $380 of $450 ➜ Overdue (rust).")
step("◷", "Birthday Gifts", "Target Dec 31 is less than 60 days away ➜ Due Soon. A -$60 gift withdrawal is logged in May.")
step("↓", "Car Maintenance", "A -$180 withdrawal for new tyres in August reduces the balance automatically.")
step("○", "Online Course", "No contributions yet ➜ Not Started.")
step("＋", "Opening balances", "Emergency Fund and New Roof each start with an “Opening Balance” contribution.")
step("📅", "Planner", "Monthly plans for 2026 show Actual vs Planned (Online Course was planned from September but never funded, so it shows a shortfall). December is still ahead, so it shows a gap.")

h("10 · Good to know")
step("Google Sheets", "Best experience",
     "Upload the file to Google Drive ➜ open with Google Sheets ➜ File ➜ Save as Google Sheets. "
     "Everything uses standard formulas that work in Google Sheets and Excel.")
step("Capacity", "Room to grow",
     "20 sinking funds · 1,000 contributions · 300 planner rows · 15 savings goals · 25 categories.")
step("Dates", "TODAY()-based",
     "Status, Days Left and Saved This Month use today's date, so they update every day by themselves.")
s0.freeze_panes = "A5"
s0.sheet_properties.tabColor = CHARCOAL

# fonts default for blank cells
for wsx in wb.worksheets:
    wsx.sheet_format.defaultRowHeight = 18

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Sinking_Funds_Tracker.xlsx")
wb.save(out)
print("saved", out, "contributions:", len(contribs), "plan rows:", len(plan_rows))
