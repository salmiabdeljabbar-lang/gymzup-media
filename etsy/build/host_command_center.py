"""Build the Host Command Center workbook (Excel + Google Sheets compatible).

Usage: python3 host_command_center.py <output.xlsx>
"""
import random
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

YEAR = 2027
OUT = sys.argv[1] if len(sys.argv) > 1 else "Host-Command-Center.xlsx"
random.seed(7)

# ---------- palette & styles ----------
NAVY, OCEAN, SEA, FOAM, SAND, SHELL, CORAL, SUN = "14324A", "1F6F86", "3D97A6", "D6EEEB", "F2DFBF", "FCF6EC", "EC6F4F", "F5B041"
INPUT_FILL = PatternFill("solid", fgColor="FFF4D6")
F = "Arial"
fill = lambda c: PatternFill("solid", fgColor=c)
thin = Side(style="thin", color="E3D9C6")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
bottom_line = Border(bottom=Side(style="thin", color="E3D9C6"))
MONEY = '$#,##0.00;($#,##0.00);"-"'
MONEY0 = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = "mmm d, yyyy"
FIRST = 5          # first data row on log sheets
LAST = 504         # last data row (500 rows)

wb = Workbook()


def setup(ws, title, subtitle, widths, tab):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 34
    ws["A1"] = title
    ws["A1"].font = Font(name=F, size=20, bold=True, color=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=F, size=10, color="5B6F7D", italic=True)
    ws.row_dimensions[2].height = 18


def header(ws, row, labels, color=NAVY):
    for i, lab in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
        c.fill = fill(color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    ws.row_dimensions[row].height = 30


def body(ws, rows, cols, fmts=None, inputs=()):
    """Style a data grid: inputs get the input fill, formulas stay white."""
    fmts = fmts or {}
    for r in rows:
        for c in range(1, cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=F, size=10, color="1C2B36")
            cell.border = bottom_line
            cell.alignment = Alignment(vertical="center")
            if c in inputs:
                cell.fill = INPUT_FILL
            if c in fmts:
                cell.number_format = fmts[c]


def dv_list(ws, formula, rng):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "Please choose from the list (edit the lists on the Start tab)."
    dv.errorStyle = "warning"
    ws.add_data_validation(dv)
    dv.add(rng)


# =====================================================================
# START (instructions + settings)
# =====================================================================
st = wb.active
st.title = "Start"
setup(st, "Host Command Center " + str(YEAR), "Your short term rental business in one place. Start here.", [3, 26, 4, 26, 14, 4, 30, 4, 26], NAVY)
st.column_dimensions["A"].width = 3
steps = [
    ("1", "Fill in the yellow cells on this tab: year, your properties, booking platforms with their fee %, cleaners and expense categories."),
    ("2", "Log every reservation on the Bookings tab and every cost on the Expenses tab. Yellow cells are for you, white cells calculate automatically."),
    ("3", "Open the Dashboard. Pick a property (or All) in the yellow box to see revenue, profit, occupancy, ADR and RevPAR."),
    ("4", "Cleaning builds your turnover schedule from your bookings. Inventory flags anything below its minimum level in red."),
    ("5", "Use Pricing to find your break-even and target nightly rate, and Taxes for a yearly summary to share with your accountant."),
    ("!", "The file comes with sample data so you can see how it works. Delete the yellow sample rows on Bookings, Expenses, Inventory, Maintenance and Reviews before you start."),
]
st["B4"] = "HOW TO USE"
st["B4"].font = Font(name=F, size=10, bold=True, color=CORAL)
for i, (n, t) in enumerate(steps):
    r = 5 + i
    st.cell(row=r, column=2, value=f"{n}.  {t}").font = Font(name=F, size=10, color="1C2B36")
    st.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    st.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
    st.row_dimensions[r].height = 30

st["B12"] = "SETTINGS"
st["B12"].font = Font(name=F, size=10, bold=True, color=CORAL)
st["B13"] = "Year to track"
st["B13"].font = Font(name=F, size=10, bold=True)
st["D13"] = YEAR
st["D13"].fill = INPUT_FILL
st["D13"].font = Font(name=F, size=12, bold=True, color=NAVY)
st["D13"].alignment = Alignment(horizontal="left")
st["D13"].number_format = "0"
st["D13"].comment = Comment("All dashboards, cleaning and tax summaries use this year (based on check-in date and expense date).", "Host Command Center")
st["B14"] = "Legend"
st["B14"].font = Font(name=F, size=10, bold=True)
st["D14"] = "Yellow = you type here"
st["D14"].fill = INPUT_FILL
st["D14"].font = Font(name=F, size=9)
st["G14"] = "White = calculates automatically"
st["G14"].font = Font(name=F, size=9)
st["G14"].border = box

LIST_TOP, LIST_BOT = 17, 26
lists = [
    (2, "Properties", ["Salt & Sun House", "Pine Cabin", "", "", "", "", "", "", "", ""]),
    (7, "Expense categories", ["Cleaning & laundry", "Supplies & amenities", "Repairs & maintenance", "Utilities & internet", "Insurance", "Mortgage interest / rent", "Property taxes & licenses", "Software & subscriptions", "Furnishings & decor", "Other"]),
    (9, "Cleaners", ["Ana (Sparkle Co.)", "Mike", "", "", "", "", "", "", "", ""]),
]
for col, title, vals in lists:
    st.cell(row=16, column=col, value=title)
    header_cell = st.cell(row=16, column=col)
    header_cell.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    header_cell.fill = fill(OCEAN)
    for i, v in enumerate(vals):
        c = st.cell(row=LIST_TOP + i, column=col, value=v or None)
        c.fill = INPUT_FILL
        c.border = box
        c.font = Font(name=F, size=10)
# platforms with fee %
st.cell(row=16, column=4, value="Booking platforms")
st.cell(row=16, column=5, value="Host fee %")
for col in (4, 5):
    c = st.cell(row=16, column=col)
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    c.fill = fill(OCEAN)
platforms = [("Airbnb", 0.03), ("Vrbo", 0.05), ("Booking.com", 0.15), ("Direct booking", 0.0), ("Other", 0.0)]
for i in range(10):
    name, fee = platforms[i] if i < len(platforms) else (None, None)
    a = st.cell(row=LIST_TOP + i, column=4, value=name)
    b = st.cell(row=LIST_TOP + i, column=5, value=fee)
    for c in (a, b):
        c.fill = INPUT_FILL
        c.border = box
        c.font = Font(name=F, size=10)
    b.number_format = "0.0%"
st["E16"].comment = Comment("Example rates only. Enter the fee each platform actually deducts from your payout (check your platform's host fee settings).", "Host Command Center")
st.cell(row=28, column=2, value="Tip: change any list above and every dropdown in the workbook updates.").font = Font(name=F, size=9, italic=True, color="5B6F7D")

# hidden helper lists for dropdowns
st["K16"], st["L16"] = "helper: expense property", "helper: dashboard property"
st[f"K{LIST_TOP}"], st[f"L{LIST_TOP}"] = "Shared", "All"
for i in range(10):
    st[f"K{LIST_TOP + 1 + i}"] = f'=IF($B${LIST_TOP + i}="","",$B${LIST_TOP + i})'
    st[f"L{LIST_TOP + 1 + i}"] = f'=IF($B${LIST_TOP + i}="","",$B${LIST_TOP + i})'
st.column_dimensions["K"].hidden = True
st.column_dimensions["L"].hidden = True
PROPS = f"Start!$B${LIST_TOP}:$B${LIST_BOT}"
PLATS = f"Start!$D${LIST_TOP}:$D${LIST_BOT}"
FEES = f"Start!$E${LIST_TOP}:$E${LIST_BOT}"
CATS = f"Start!$G${LIST_TOP}:$G${LIST_BOT}"
CLEANERS = f"Start!$I${LIST_TOP}:$I${LIST_BOT}"
YR = "Start!$D$13"

# =====================================================================
# BOOKINGS
# =====================================================================
bk = wb.create_sheet("Bookings")
setup(bk, "Bookings", "One row per reservation. Yellow = you type. Nights, fees and payout calculate automatically.",
      [13, 13, 8, 18, 15, 18, 8, 11, 11, 12, 12, 12, 12, 12, 24], OCEAN)
cols = ["Check-in", "Check-out", "Nights", "Property", "Platform", "Guest name", "Guests", "Nightly rate", "Cleaning fee",
        "Gross booking", "Platform fee", "Taxes collected", "Payout", "Status", "Notes"]
header(bk, 4, cols)
for r in range(FIRST, LAST + 1):
    bk[f"C{r}"] = f'=IF(AND(A{r}<>"",B{r}<>""),B{r}-A{r},"")'
    bk[f"J{r}"] = f'=IF(C{r}="","",C{r}*H{r}+I{r})'
    bk[f"K{r}"] = f'=IF(J{r}="","",ROUND(J{r}*IFERROR(INDEX({FEES},MATCH(E{r},{PLATS},0)),0),2))'
    bk[f"M{r}"] = f'=IF(J{r}="","",J{r}-K{r})'
body(bk, range(FIRST, LAST + 1), 15, {1: DATE, 2: DATE, 3: "0", 8: MONEY, 9: MONEY, 10: MONEY, 11: MONEY, 12: MONEY, 13: MONEY},
     inputs=(1, 2, 4, 5, 6, 7, 8, 9, 12, 14, 15))
dv_list(bk, PROPS, f"D{FIRST}:D{LAST}")
dv_list(bk, PLATS, f"E{FIRST}:E{LAST}")
dv_list(bk, '"Confirmed,Completed,Cancelled"', f"N{FIRST}:N{LAST}")
bk.freeze_panes = "A5"
bk.auto_filter.ref = f"A4:O{LAST}"
bk.conditional_formatting.add(f"A{FIRST}:O{LAST}", FormulaRule(formula=[f'$N{FIRST}="Cancelled"'], font=Font(color="A0A0A0", strike=True)))
bk["C4"].comment = Comment("Check-out minus check-in.", "Host Command Center")
bk["K4"].comment = Comment("Gross booking x the platform's host fee % from the Start tab.", "Host Command Center")
bk["L4"].comment = Comment("Occupancy taxes the platform collected for you (for your records; not included in payout).", "Host Command Center")

# sample bookings (no overlaps per property)
guests = ["Emily R.", "Jason M.", "The Parkers", "Sofia L.", "Daniel K.", "Hannah W.", "Marco T.", "Priya S.", "The Nguyens", "Olivia B.",
          "Chris D.", "Ava G.", "Noah F.", "Lena H.", "Sam & Jo", "Grace P.", "Ethan C.", "Mia V.", "Lucas A.", "Zoe N."]
plat_names = [p for p, _ in platforms[:4]]
season = {1: 145, 2: 150, 3: 165, 4: 175, 5: 195, 6: 245, 7: 275, 8: 270, 9: 210, 10: 180, 11: 155, 12: 185}
rows = []
for prop, mult, clean in (("Salt & Sun House", 1.0, 120), ("Pine Cabin", 0.75, 95)):
    d = date(YEAR, 1, 3)
    while d < date(YEAR, 12, 20):
        gap = random.choice([0, 0, 1, 2, 3, 4, 6])
        d += timedelta(days=gap)
        n = random.choice([2, 2, 3, 3, 4, 5, 7])
        out = d + timedelta(days=n)
        if out > date(YEAR, 12, 31):
            break
        rate = round(season[d.month] * mult + random.choice([-10, 0, 0, 10, 15]))
        status = "Cancelled" if random.random() < 0.05 else ("Completed" if d < date(YEAR, 10, 1) else "Confirmed")
        rows.append([d, out, prop, random.choices(plat_names, weights=[5, 2, 1, 2])[0], random.choice(guests), random.choice([2, 2, 3, 4, 4, 5, 6]), rate, clean, status])
        d = out
rows.sort(key=lambda x: x[0])
rows = rows[:120]
for i, (cin, cout, prop, plat, g, ng, rate, clean, status) in enumerate(rows):
    r = FIRST + i
    bk[f"A{r}"], bk[f"B{r}"], bk[f"D{r}"], bk[f"E{r}"], bk[f"F{r}"] = cin, cout, prop, plat, g
    bk[f"G{r}"], bk[f"H{r}"], bk[f"I{r}"], bk[f"N{r}"] = ng, rate, clean, status
    bk[f"L{r}"] = round((rate * (cout - cin).days + clean) * 0.1, 2)
N_BOOK = len(rows)

# =====================================================================
# EXPENSES
# =====================================================================
ex = wb.create_sheet("Expenses")
setup(ex, "Expenses", "Log every cost. Choose 'Shared' for costs that cover all properties.", [13, 18, 24, 20, 30, 12, 10], CORAL)
header(ex, 4, ["Date", "Property", "Category", "Vendor", "Description", "Amount", "Receipt?"], CORAL)
body(ex, range(FIRST, LAST + 1), 7, {1: DATE, 6: MONEY}, inputs=(1, 2, 3, 4, 5, 6, 7))
dv_list(ex, f"Start!$K${LIST_TOP}:$K${LIST_BOT + 1}", f"B{FIRST}:B{LAST}")
dv_list(ex, CATS, f"C{FIRST}:C{LAST}")
dv_list(ex, '"Yes,No"', f"G{FIRST}:G{LAST}")
ex.freeze_panes = "A5"
ex.auto_filter.ref = f"A4:G{LAST}"
ex["B4"].comment = Comment("Pick a property, or type Shared for costs that cover all properties.", "Host Command Center")
monthly = [("Mortgage interest / rent", "Bank", 1450, "Salt & Sun House"), ("Mortgage interest / rent", "Bank", 980, "Pine Cabin"),
           ("Utilities & internet", "Coast Power & Fiber", 260, "Salt & Sun House"), ("Utilities & internet", "Valley Utilities", 190, "Pine Cabin"),
           ("Software & subscriptions", "Pricing & lock apps", 45, "Shared"), ("Insurance", "STR insurance", 140, "Shared")]
ex_rows = []
for m in range(1, 13):
    for cat, vend, amt, prop in monthly:
        ex_rows.append([date(YEAR, m, random.choice([1, 2, 3, 5])), prop, cat, vend, "Monthly", round(amt * random.uniform(.95, 1.05), 2), "Yes"])
    ex_rows.append([date(YEAR, m, 15), "Shared", "Supplies & amenities", "Costco", "Coffee, toiletries, paper goods", round(random.uniform(120, 260), 2), "Yes"])
    if m in (3, 7, 10):
        ex_rows.append([date(YEAR, m, 20), random.choice(["Salt & Sun House", "Pine Cabin"]), "Repairs & maintenance", "Handy Pro", random.choice(["Fix leaking faucet", "Deck stain", "AC service"]), round(random.uniform(150, 600), 2), "Yes"])
for m in range(1, 13):
    for prop, pay in (("Salt & Sun House", 95), ("Pine Cabin", 75)):
        turns = sum(1 for b in rows if b[2] == prop and b[8] != "Cancelled" and b[1].month == m)
        if turns:
            ex_rows.append([date(YEAR, m, 28), prop, "Cleaning & laundry", "Cleaner payout", f"{turns} turnovers", float(turns * pay), "Yes"])
ex_rows.append([date(YEAR, 4, 10), "Salt & Sun House", "Furnishings & decor", "Home store", "New patio chairs", 389.00, "Yes"])
ex_rows.append([date(YEAR, 1, 31), "Shared", "Property taxes & licenses", "City", "STR permit renewal", 250.00, "Yes"])
ex_rows.sort(key=lambda x: x[0])
for i, row in enumerate(ex_rows):
    for j, v in enumerate(row):
        ex.cell(row=FIRST + i, column=j + 1, value=v)

# =====================================================================
# CLEANING (auto schedule)
# =====================================================================
cl = wb.create_sheet("Cleaning")
setup(cl, "Cleaning Schedule", "Built from your bookings. Assign a cleaner, enter the fee and tick paid. Add bookings at the bottom of Bookings so rows stay lined up.",
      [14, 18, 16, 16, 10, 24, 18, 11, 9], SEA)
header(cl, 4, ["Turnover date", "Property", "Departing guest", "Next check-in", "Days free", "Alert", "Cleaner", "Cleaning pay", "Paid?"], SEA)
for r in range(FIRST, LAST + 1):
    b = f"Bookings!"
    ok = f'AND(Bookings!B{r}<>"",Bookings!N{r}<>"Cancelled")'
    cl[f"A{r}"] = f'=IF({ok},Bookings!B{r},"")'
    cl[f"B{r}"] = f'=IF(A{r}="","",Bookings!D{r})'
    cl[f"C{r}"] = f'=IF(A{r}="","",Bookings!F{r})'
    cl[f"D{r}"] = (f'=IF(A{r}="","",IF(COUNTIFS(Bookings!$D${FIRST}:$D${LAST},B{r},Bookings!$A${FIRST}:$A${LAST},">="&A{r},Bookings!$N${FIRST}:$N${LAST},"<>Cancelled")=0,"",'
                   f'_xlfn.MINIFS(Bookings!$A${FIRST}:$A${LAST},Bookings!$D${FIRST}:$D${LAST},B{r},Bookings!$A${FIRST}:$A${LAST},">="&A{r},Bookings!$N${FIRST}:$N${LAST},"<>Cancelled")))')
    cl[f"E{r}"] = f'=IF(OR(A{r}="",D{r}=""),"",D{r}-A{r})'
    cl[f"F{r}"] = f'=IF(A{r}="","",IF(D{r}="","No next booking yet",IF(E{r}=0,"SAME-DAY TURNOVER",IF(E{r}=1,"Next day",""))))'
body(cl, range(FIRST, LAST + 1), 9, {1: DATE, 4: DATE, 5: "0", 8: MONEY}, inputs=(7, 8, 9))
dv_list(cl, CLEANERS, f"G{FIRST}:G{LAST}")
dv_list(cl, '"Yes,No"', f"I{FIRST}:I{LAST}")
cl.conditional_formatting.add(f"F{FIRST}:F{LAST}", CellIsRule(operator="equal", formula=['"SAME-DAY TURNOVER"'], font=Font(bold=True, color="FFFFFF"), fill=fill(CORAL)))
cl.conditional_formatting.add(f"I{FIRST}:I{LAST}", CellIsRule(operator="equal", formula=['"No"'], font=Font(bold=True, color=CORAL)))
cl.freeze_panes = "A5"
cl.auto_filter.ref = f"A4:I{LAST}"
cleaners_sample = ["Ana (Sparkle Co.)", "Mike"]
for i in range(N_BOOK):
    r = FIRST + i
    if rows[i][8] == "Cancelled":
        continue
    cl[f"G{r}"] = cleaners_sample[0] if rows[i][2] == "Salt & Sun House" else cleaners_sample[1]
    cl[f"H{r}"] = 95 if rows[i][2] == "Salt & Sun House" else 75
    cl[f"I{r}"] = "Yes" if rows[i][8] == "Completed" else "No"

# =====================================================================
# INVENTORY
# =====================================================================
inv = wb.create_sheet("Inventory")
setup(inv, "Inventory & Restock", "Set a minimum level for each item. Anything below it turns red so you know what to buy.", [26, 18, 11, 11, 13, 11, 11, 12, 26], SUN)
header(inv, 4, ["Item", "Location", "Minimum", "In stock", "Status", "To buy", "Unit cost", "Restock cost", "Where to buy"], OCEAN)
for r in range(FIRST, 104):
    inv[f"E{r}"] = f'=IF(OR(A{r}="",D{r}=""),"",IF(D{r}<C{r},"RESTOCK","OK"))'
    inv[f"F{r}"] = f'=IF(E{r}="","",MAX(0,C{r}-D{r}))'
    inv[f"H{r}"] = f'=IF(OR(F{r}="",G{r}=""),"",F{r}*G{r})'
body(inv, range(FIRST, 104), 9, {7: MONEY, 8: MONEY}, inputs=(1, 2, 3, 4, 7, 9))
inv.conditional_formatting.add(f"E{FIRST}:E103", CellIsRule(operator="equal", formula=['"RESTOCK"'], font=Font(bold=True, color="FFFFFF"), fill=fill(CORAL)))
inv.conditional_formatting.add(f"E{FIRST}:E103", CellIsRule(operator="equal", formula=['"OK"'], font=Font(bold=True, color=OCEAN), fill=fill(FOAM)))
items = [("Toilet paper (rolls)", "Hall closet", 24, 30, .6), ("Paper towels", "Kitchen", 6, 4, 1.5), ("Coffee pods", "Kitchen drawer", 30, 12, .45),
         ("Dish soap", "Under sink", 2, 3, 3.5), ("Dishwasher tablets", "Under sink", 20, 25, .25), ("Trash bags", "Kitchen", 30, 18, .2),
         ("Shampoo (bottles)", "Bathroom", 4, 5, 4), ("Conditioner (bottles)", "Bathroom", 4, 2, 4), ("Body wash", "Bathroom", 4, 4, 4.5),
         ("Hand soap", "Bathrooms", 4, 6, 3), ("Laundry detergent", "Garage", 2, 1, 12), ("Bath towels", "Linen closet", 12, 12, 14),
         ("Sheet sets (queen)", "Linen closet", 6, 4, 38), ("Light bulbs", "Garage", 6, 8, 2.5), ("AA batteries", "Junk drawer", 8, 3, .8),
         ("Sunscreen", "Bathroom", 2, 2, 9), ("Welcome snacks", "Pantry", 10, 14, 2)]
for i, (it, loc, mn, stock, cost) in enumerate(items):
    r = FIRST + i
    inv[f"A{r}"], inv[f"B{r}"], inv[f"C{r}"], inv[f"D{r}"], inv[f"G{r}"], inv[f"I{r}"] = it, loc, mn, stock, cost, "Warehouse store"
inv["K4"] = "Items to restock"
inv["K4"].font = Font(name=F, size=10, bold=True, color=NAVY)
inv["K5"] = f'=COUNTIF(E{FIRST}:E103,"RESTOCK")'
inv["K5"].font = Font(name=F, size=22, bold=True, color=CORAL)
inv["K7"] = "Restock budget"
inv["K7"].font = Font(name=F, size=10, bold=True, color=NAVY)
inv["K8"] = f"=SUM(H{FIRST}:H103)"
inv["K8"].font = Font(name=F, size=22, bold=True, color=NAVY)
inv["K8"].number_format = MONEY
inv.column_dimensions["K"].width = 20
inv.freeze_panes = "A5"

# =====================================================================
# MAINTENANCE
# =====================================================================
mt = wb.create_sheet("Maintenance")
setup(mt, "Maintenance Log", "Track every fix so nothing slips through between guests.", [13, 18, 16, 32, 11, 12, 12, 13, 24], OCEAN)
header(mt, 4, ["Reported", "Property", "Area", "Issue", "Priority", "Status", "Cost", "Fixed on", "Notes"])
body(mt, range(FIRST, 205), 9, {1: DATE, 7: MONEY, 8: DATE}, inputs=range(1, 10))
dv_list(mt, PROPS, f"B{FIRST}:B204")
dv_list(mt, '"High,Medium,Low"', f"E{FIRST}:E204")
dv_list(mt, '"Open,In progress,Done"', f"F{FIRST}:F204")
mt.conditional_formatting.add(f"E{FIRST}:E204", CellIsRule(operator="equal", formula=['"High"'], font=Font(bold=True, color=CORAL)))
mt.conditional_formatting.add(f"A{FIRST}:I204", FormulaRule(formula=[f'$F{FIRST}="Done"'], font=Font(color="9AA5AD")))
for i, row in enumerate([(date(YEAR, 3, 12), "Salt & Sun House", "Bathroom", "Faucet dripping", "Medium", "Done", 180, date(YEAR, 3, 14), "Replaced cartridge"),
                         (date(YEAR, 7, 2), "Pine Cabin", "Deck", "Loose board near stairs", "High", "Done", 95, date(YEAR, 7, 3), ""),
                         (date(YEAR, 10, 18), "Salt & Sun House", "Bedroom 2", "AC remote missing", "Low", "Open", None, None, "Order universal remote")]):
    for j, v in enumerate(row):
        mt.cell(row=FIRST + i, column=j + 1, value=v)
mt.freeze_panes = "A5"

# =====================================================================
# REVIEWS
# =====================================================================
rv = wb.create_sheet("Reviews")
setup(rv, "Guest Reviews", "Log ratings to spot patterns and keep your average high.", [13, 18, 18, 15, 9, 40, 11], SUN)
header(rv, 4, ["Date", "Guest", "Property", "Platform", "Rating", "What they said", "Replied?"], OCEAN)
body(rv, range(FIRST, LAST + 1), 7, {1: DATE, 5: "0.0"}, inputs=range(1, 8))
dv_list(rv, PROPS, f"C{FIRST}:C{LAST}")
dv_list(rv, PLATS, f"D{FIRST}:D{LAST}")
dv_list(rv, '"Yes,No"', f"G{FIRST}:G{LAST}")
rdv = DataValidation(type="decimal", operator="between", formula1="1", formula2="5")
rv.add_data_validation(rdv)
rdv.add(f"E{FIRST}:E{LAST}")
rv.conditional_formatting.add(f"E{FIRST}:E{LAST}", CellIsRule(operator="lessThan", formula=["4"], font=Font(bold=True, color=CORAL)))
comments = ["Spotless and so cozy!", "Perfect location, will be back.", "Loved the welcome book and tips.", "Great stay, a bit noisy at night.", "Host was super responsive."]
k = 0
for i in range(N_BOOK):
    if rows[i][8] != "Completed" or i % 2:
        continue
    r = FIRST + k
    rv[f"A{r}"], rv[f"B{r}"], rv[f"C{r}"], rv[f"D{r}"] = rows[i][1] + timedelta(days=2), rows[i][4], rows[i][2], rows[i][3]
    rv[f"E{r}"] = random.choice([5, 5, 5, 5, 4])
    rv[f"F{r}"] = random.choice(comments)
    rv[f"G{r}"] = "Yes"
    k += 1
rv.freeze_panes = "A5"

# =====================================================================
# PRICING CALCULATOR
# =====================================================================
pr = wb.create_sheet("Pricing")
setup(pr, "Pricing Calculator", "Find the nightly rate that covers your costs and hits your profit goal. Yellow = your numbers.", [3, 34, 16, 4, 30, 14, 14], CORAL)
pr["B4"] = "MONTHLY FIXED COSTS"
pr["B4"].font = Font(name=F, size=10, bold=True, color=CORAL)
fixed = [("Mortgage / rent", 1450), ("Utilities & internet", 260), ("Insurance", 70), ("Software & apps", 25), ("HOA / other", 0)]
for i, (lab, v) in enumerate(fixed):
    pr[f"B{5 + i}"], pr[f"C{5 + i}"] = lab, v
pr["B10"] = "Total fixed per month"
pr["C10"] = "=SUM(C5:C9)"
pr["B12"] = "PER-STAY COSTS"
pr["B12"].font = Font(name=F, size=10, bold=True, color=CORAL)
var = [("Cleaning (what you pay)", 95), ("Laundry", 25), ("Supplies & amenities", 15)]
for i, (lab, v) in enumerate(var):
    pr[f"B{13 + i}"], pr[f"C{13 + i}"] = lab, v
pr["B16"] = "Total per stay"
pr["C16"] = "=SUM(C13:C15)"
pr["B18"] = "ASSUMPTIONS"
pr["B18"].font = Font(name=F, size=10, bold=True, color=CORAL)
assum = [("Target occupancy", 0.65, PCT), ("Average stay (nights)", 3, "0.0"), ("Platform fee", 0.03, PCT), ("Cleaning fee you charge guests", 120, MONEY), ("Profit goal per month", 1500, MONEY)]
for i, (lab, v, fmt) in enumerate(assum):
    pr[f"B{19 + i}"], pr[f"C{19 + i}"] = lab, v
    pr[f"C{19 + i}"].number_format = fmt
for r in list(range(5, 10)) + list(range(13, 16)) + list(range(19, 24)):
    pr[f"C{r}"].fill = INPUT_FILL
    pr[f"B{r}"].font = Font(name=F, size=10)
    pr[f"C{r}"].font = Font(name=F, size=10, color="0000FF")
    pr[f"C{r}"].border = box
    if r < 19:
        pr[f"C{r}"].number_format = MONEY
for r in (10, 16):
    pr[f"B{r}"].font = Font(name=F, size=10, bold=True)
    pr[f"C{r}"].font = Font(name=F, size=10, bold=True)
    pr[f"C{r}"].number_format = MONEY
pr["E4"] = "YOUR NUMBERS"
pr["E4"].font = Font(name=F, size=10, bold=True, color=CORAL)
out_rows = [
    ("Booked nights per month", "=30.4*C19", "0.0"),
    ("Stays per month", "=IF(C20=0,0,F5/C20)", "0.0"),
    ("Total monthly cost", "=C10+F6*C16", MONEY),
    ("Cleaning fees collected", "=F6*C22", MONEY),
    ("Room revenue needed to break even", "=MAX(0,(F7/(1-C21))-F8)", MONEY),
    ("Room revenue for your profit goal", "=MAX(0,((F7+C23)/(1-C21))-F8)", MONEY),
    ("Break-even nightly rate", "=IF(F5=0,0,F9/F5)", MONEY),
    ("Target nightly rate", "=IF(F5=0,0,F10/F5)", MONEY),
]
for i, (lab, fml, fmt) in enumerate(out_rows):
    r = 5 + i
    pr[f"E{r}"], pr[f"F{r}"] = lab, fml
    pr[f"F{r}"].number_format = fmt
    pr[f"E{r}"].font = Font(name=F, size=10)
    pr[f"F{r}"].font = Font(name=F, size=10)
    pr[f"F{r}"].border = box
for r, color in ((11, OCEAN), (12, CORAL)):
    pr[f"E{r}"].font = Font(name=F, size=12, bold=True, color=color)
    pr[f"F{r}"].font = Font(name=F, size=16, bold=True, color="FFFFFF")
    pr[f"F{r}"].fill = fill(color)
    pr.row_dimensions[r].height = 28
pr["E14"] = "SEASONAL RATES"
pr["E14"].font = Font(name=F, size=10, bold=True, color=CORAL)
for col, lab in zip("EFG", ["Season", "Multiplier", "Nightly rate"]):
    c = pr[f"{col}15"]
    c.value = lab
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    c.fill = fill(NAVY)
for i, (s, m) in enumerate([("Low season", 0.8), ("Shoulder", 1.0), ("High season", 1.35), ("Holidays & events", 1.6), ("Last-minute (< 3 days)", 0.9)]):
    r = 16 + i
    pr[f"E{r}"], pr[f"F{r}"], pr[f"G{r}"] = s, m, f"=F12*F{r}"
    pr[f"F{r}"].fill = INPUT_FILL
    pr[f"F{r}"].number_format = '0.00"x"'
    pr[f"G{r}"].number_format = MONEY
    for col in "EFG":
        pr[f"{col}{r}"].border = box
        pr[f"{col}{r}"].font = Font(name=F, size=10, color="0000FF" if col == "F" else "1C2B36")
pr["E22"] = "Multipliers are examples. Adjust them to your market."
pr["E22"].font = Font(name=F, size=9, italic=True, color="5B6F7D")
pr["C19"].comment = Comment("Share of nights you expect to be booked. 30.4 = average days in a month.", "Host Command Center")

# =====================================================================
# DASHBOARD
# =====================================================================
db = wb.create_sheet("Dashboard", 1)
setup(db, f"Dashboard", "Pick a property in the yellow box. Everything updates from your Bookings, Expenses and Reviews.", [3, 14, 13, 13, 13, 13, 13, 13, 3, 24, 14, 12, 12], NAVY)
db["J1"] = "Property"
db["J1"].font = Font(name=F, size=10, bold=True, color=NAVY)
db["J1"].alignment = Alignment(horizontal="right", vertical="center")
db["K1"] = "All"
db["K1"].fill = INPUT_FILL
db["K1"].font = Font(name=F, size=12, bold=True, color=NAVY)
db["K1"].alignment = Alignment(horizontal="center", vertical="center")
db.merge_cells("K1:M1")
dv_list(db, f"Start!$L${LIST_TOP}:$L${LIST_BOT + 1}", "K1")
db["K1"].comment = Comment("Choose All or one property. The list comes from your properties on the Start tab.", "Host Command Center")
db["J2"] = "Year"
db["J2"].font = Font(name=F, size=10, bold=True, color=NAVY)
db["J2"].alignment = Alignment(horizontal="right")
db["K2"] = f"={YR}"
db["K2"].font = Font(name=F, size=10, bold=True, color=NAVY)
db["K2"].number_format = "0"
# helper cells (hidden column O/P)
db["O1"] = "helper: property criteria"
db["P1"] = '=IF(K1="All","*",K1)'
db["O2"] = "helper: properties counted"
db["P2"] = f'=IF(K1="All",MAX(1,COUNTA({PROPS})),1)'
db["O3"] = "helper: year start"
db["P3"] = f"=DATE({YR},1,1)"
db["O4"] = "helper: year end"
db["P4"] = f"=DATE({YR},12,31)"
db.column_dimensions["O"].hidden = True
db.column_dimensions["P"].hidden = True
PC, NP, YS, YE = "$P$1", "$P$2", "$P$3", "$P$4"
BA, BB, BC, BD, BH, BJ, BM, BN = (f"Bookings!${c}${FIRST}:${c}${LAST}" for c in "ABCDHJMN")
EXA, EXB, EXF = (f"Expenses!${c}${FIRST}:${c}${LAST}" for c in "ABF")
prop_ok = f'((({BD})=$K$1)+($K$1="All")>0)'

# KPI cards: label row / value row / note row
cards = [
    ("B", "C", "GROSS REVENUE", f'=SUMIFS({BJ},{BA},">="&{YS},{BA},"<="&{YE},{BD},{PC},{BN},"<>Cancelled")', MONEY0, NAVY, "Bookings incl. cleaning fees"),
    ("D", "E", "NET PROFIT", f'=SUMIFS({BM},{BA},">="&{YS},{BA},"<="&{YE},{BD},{PC},{BN},"<>Cancelled")-SUMIFS({EXF},{EXA},">="&{YS},{EXA},"<="&{YE},{EXB},{PC})', MONEY0, CORAL, "Payouts minus expenses"),
    ("F", "G", "OCCUPANCY", "=IFERROR(SUM(F12:F23)/SUM(G12:G23),0)", PCT, OCEAN, "Booked ÷ available nights"),
    ("H", "H", "AVG RATING", f'=IFERROR(AVERAGEIFS(Reviews!$E${FIRST}:$E${LAST},Reviews!$C${FIRST}:$C${LAST},{PC}),"-")', "0.00", SEA, "From Reviews"),
]
row2 = [
    ("B", "C", "ADR", f"=IFERROR(SUMPRODUCT(({BA}>={YS})*({BA}<={YE})*({BN}<>\"Cancelled\")*{prop_ok},{BC},{BH})/SUMPRODUCT(({BA}>={YS})*({BA}<={YE})*({BN}<>\"Cancelled\")*{prop_ok},{BC}),0)", MONEY0, "Average nightly rate"),
    ("D", "E", "RevPAR", "=IFERROR(SUMPRODUCT(({BA}>={YS})*({BA}<={YE})*({BN}<>\"Cancelled\")*{prop_ok},{BC},{BH})/SUM(G12:G23),0)".format(BA=BA, YS=YS, YE=YE, BN=BN, prop_ok=prop_ok, BC=BC, BH=BH), MONEY0, "Revenue per available night"),
    ("F", "G", "NIGHTS BOOKED", "=SUM(F12:F23)", "#,##0", "Across the year"),
    ("H", "H", "BOOKINGS", f'=COUNTIFS({BA},">="&{YS},{BA},"<="&{YE},{BD},{PC},{BN},"<>Cancelled")', "#,##0", "Excl. cancelled"),
]
db.row_dimensions[4].height = 18
db.row_dimensions[5].height = 34
db.row_dimensions[6].height = 16
db.row_dimensions[7].height = 18
db.row_dimensions[8].height = 30
db.row_dimensions[9].height = 16


def card(c1, c2, r_lab, label, fml, fmt, color, note, light=False):
    if c1 != c2:
        for rr in (r_lab, r_lab + 1, r_lab + 2):
            db.merge_cells(f"{c1}{rr}:{c2}{rr}")
    lab, val, nt = db[f"{c1}{r_lab}"], db[f"{c1}{r_lab + 1}"], db[f"{c1}{r_lab + 2}"]
    lab.value, val.value, nt.value = label, fml, note
    bg = fill(color) if not light else fill(SHELL)
    fg = "FFFFFF" if not light else NAVY
    for col in range(ord(c1), ord(c2) + 1):
        for rr in (r_lab, r_lab + 1, r_lab + 2):
            db[f"{chr(col)}{rr}"].fill = bg
    lab.font = Font(name=F, size=8, bold=True, color=fg if light else "FFFFFF")
    val.font = Font(name=F, size=18 if not light else 15, bold=True, color=fg)
    nt.font = Font(name=F, size=8, color=("D6EEEB" if not light else "5B6F7D"))
    val.number_format = fmt
    for cc in (lab, val, nt):
        cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)


for c1, c2, lab, fml, fmt, color, note in cards:
    card(c1, c2, 4, lab, fml, fmt, color, note)
for c1, c2, lab, fml, fmt, note in row2:
    card(c1, c2, 7, lab, fml, fmt, NAVY, note, light=True)

# monthly table
db["B10"] = "MONTH BY MONTH"
db["B10"].font = Font(name=F, size=10, bold=True, color=CORAL)
mh = ["Month", "Revenue", "Expenses", "Profit", "Nights", "Available", "Occupancy"]
for i, lab in enumerate(mh):
    c = db.cell(row=11, column=2 + i, value=lab)
    c.font = Font(name=F, size=9, bold=True, color="FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="center")
for m in range(12):
    r = 12 + m
    S, E = f"DATE({YR},{m + 1},1)", f"EOMONTH(DATE({YR},{m + 1},1),0)"
    db[f"B{r}"] = f"=DATE({YR},{m + 1},1)"
    db[f"B{r}"].number_format = "mmm"
    db[f"C{r}"] = f'=SUMIFS({BM},{BA},">="&{S},{BA},"<="&{E},{BD},{PC},{BN},"<>Cancelled")'
    db[f"D{r}"] = f'=SUMIFS({EXF},{EXA},">="&{S},{EXA},"<="&{E},{EXB},{PC})'
    db[f"E{r}"] = f"=C{r}-D{r}"
    S1, E1 = f"$B{r}", f"(EOMONTH($B{r},0)+1)"
    db[f"F{r}"] = (f'=SUMPRODUCT(({BB}>{S1})*({BA}<{E1})*({BN}<>"Cancelled")*{prop_ok},'
                   f'({BB}-({BB}>{E1})*({BB}-{E1}))-({BA}+({BA}<{S1})*({S1}-{BA})))')
    db[f"G{r}"] = f"=DAY(EOMONTH($B{r},0))*{NP}"
    db[f"H{r}"] = f"=IFERROR(F{r}/G{r},0)"
    for col, fmt in zip("CDEFGH", [MONEY0, MONEY0, MONEY0, "0", "0", PCT]):
        db[f"{col}{r}"].number_format = fmt
    for col in "BCDEFGH":
        db[f"{col}{r}"].font = Font(name=F, size=10)
        db[f"{col}{r}"].border = bottom_line
db["B24"] = "Total"
for col in "CDEFG":
    db[f"{col}24"] = f"=SUM({col}12:{col}23)"
db["H24"] = "=IFERROR(F24/G24,0)"
for col, fmt in zip("CDEFGH", [MONEY0, MONEY0, MONEY0, "0", "0", PCT]):
    db[f"{col}24"].number_format = fmt
for col in "BCDEFGH":
    db[f"{col}24"].font = Font(name=F, size=10, bold=True, color=NAVY)
    db[f"{col}24"].fill = fill(SAND)
db.conditional_formatting.add("H12:H23", ColorScaleRule(start_type="num", start_value=0, start_color="FCF6EC", end_type="num", end_value=1, end_color="3D97A6"))
db.conditional_formatting.add("E12:E23", CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, color=CORAL)))

# platform table
db["J10"] = "BY PLATFORM"
db["J10"].font = Font(name=F, size=10, bold=True, color=CORAL)
for i, lab in enumerate(["Platform", "Payout", "Bookings", "Share"]):
    c = db.cell(row=11, column=10 + i, value=lab)
    c.font = Font(name=F, size=9, bold=True, color="FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="center")
for i in range(5):
    r = 12 + i
    db[f"J{r}"] = f'=IF(Start!$D${LIST_TOP + i}="","",Start!$D${LIST_TOP + i})'
    db[f"K{r}"] = f'=IF(J{r}="","",SUMIFS({BM},{BA},">="&{YS},{BA},"<="&{YE},{BD},{PC},{BN},"<>Cancelled",Bookings!$E${FIRST}:$E${LAST},J{r}))'
    db[f"L{r}"] = f'=IF(J{r}="","",COUNTIFS({BA},">="&{YS},{BA},"<="&{YE},{BD},{PC},{BN},"<>Cancelled",Bookings!$E${FIRST}:$E${LAST},J{r}))'
    db[f"M{r}"] = f'=IF(J{r}="","",IFERROR(K{r}/SUM($K$12:$K$16),0))'
    for col, fmt in zip("KLM", [MONEY0, "0", PCT]):
        db[f"{col}{r}"].number_format = fmt
    for col in "JKLM":
        db[f"{col}{r}"].font = Font(name=F, size=10)
        db[f"{col}{r}"].border = bottom_line

# expense breakdown
db["J18"] = "TOP EXPENSES"
db["J18"].font = Font(name=F, size=10, bold=True, color=CORAL)
for i, lab in enumerate(["Category", "Spent", "", "Share"]):
    c = db.cell(row=19, column=10 + i, value=lab or None)
    c.font = Font(name=F, size=9, bold=True, color="FFFFFF")
    c.fill = fill(NAVY)
db.merge_cells("K19:L19")
for i in range(10):
    r = 20 + i
    db[f"J{r}"] = f'=IF(Start!$G${LIST_TOP + i}="","",Start!$G${LIST_TOP + i})'
    db[f"K{r}"] = f'=IF(J{r}="","",SUMIFS({EXF},{EXA},">="&{YS},{EXA},"<="&{YE},{EXB},{PC},Expenses!$C${FIRST}:$C${LAST},J{r}))'
    db[f"M{r}"] = f'=IF(J{r}="","",IFERROR(K{r}/SUM($K$20:$K$29),0))'
    db[f"K{r}"].number_format = MONEY0
    db[f"M{r}"].number_format = PCT
    for col in "JKLM":
        db[f"{col}{r}"].font = Font(name=F, size=9)
        db[f"{col}{r}"].border = bottom_line
    db.merge_cells(f"K{r}:L{r}")

# charts
ch = BarChart()
ch.type = "col"
ch.grouping = "clustered"
ch.title = "Revenue vs expenses"
ch.y_axis.title = None
ch.add_data(Reference(db, min_col=3, max_col=4, min_row=11, max_row=23), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=12, max_row=23))
ch.series[0].graphicalProperties.solidFill = OCEAN
ch.series[1].graphicalProperties.solidFill = CORAL
ch.y_axis.numFmt = "$#,##0"
ch.height, ch.width = 7.5, 17
ch.roundedCorners = False
ch.legend.position = "b"
ch.legend.overlay = False
ch.title.overlay = False
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.majorGridlines = None
db.add_chart(ch, "B27")

oc = LineChart()
oc.title = "Occupancy by month"
oc.add_data(Reference(db, min_col=8, min_row=11, max_row=23), titles_from_data=True)
oc.set_categories(Reference(db, min_col=2, min_row=12, max_row=23))
oc.series[0].graphicalProperties.line.solidFill = SEA
oc.series[0].graphicalProperties.line.width = 28000
oc.y_axis.numFmt = "0%"
oc.y_axis.scaling.min = 0
oc.y_axis.scaling.max = 1
oc.height, oc.width = 7.5, 17
oc.roundedCorners = False
oc.legend = None
oc.title.overlay = False
oc.x_axis.delete = False
oc.y_axis.delete = False
db.add_chart(oc, "B43")

dn = DoughnutChart()
dn.title = "Payout by platform"
dn.add_data(Reference(db, min_col=11, min_row=11, max_row=16), titles_from_data=True)
dn.set_categories(Reference(db, min_col=10, min_row=12, max_row=16))
dn.height, dn.width = 7.5, 11
dn.roundedCorners = False
dn.legend.position = "r"
dn.legend.overlay = False
dn.title.overlay = False
from openpyxl.chart.series import DataPoint
for idx, color in enumerate([OCEAN, CORAL, SUN, SEA, NAVY]):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = color
    dn.series[0].dPt.append(pt)
db.add_chart(dn, "J31")

# =====================================================================
# TAXES
# =====================================================================
tx = wb.create_sheet("Taxes")
setup(tx, "Year-End Summary", "Income and expenses by category for the year on the Start tab. Share with your accountant.", [3, 44, 16, 16, 16, 4, 30], OCEAN)
props_cols = ["Salt & Sun House", "Pine Cabin"]
hdr = ["Category", "All properties", "=Start!$B$17", "=Start!$B$18"]
for i, lab in enumerate(hdr):
    c = tx.cell(row=4, column=2 + i, value=lab)
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="center", wrap_text=True)
tx.row_dimensions[4].height = 30
tx["B5"] = "INCOME"
tx["B5"].font = Font(name=F, size=10, bold=True, color=CORAL)
inc = [("Gross bookings (incl. cleaning fees)", BJ), ("Platform fees", f"Bookings!$K${FIRST}:$K${LAST}"), ("Net payouts", BM), ("Occupancy taxes collected by platforms", f"Bookings!$L${FIRST}:$L${LAST}")]
for i, (lab, rng) in enumerate(inc):
    r = 6 + i
    tx[f"B{r}"] = lab
    tx[f"C{r}"] = f'=SUMIFS({rng},{BA},">="&DATE({YR},1,1),{BA},"<="&DATE({YR},12,31),{BN},"<>Cancelled")'
    for j, col in enumerate("DE"):
        tx[f"{col}{r}"] = f'=IF({col}$4="","",SUMIFS({rng},{BA},">="&DATE({YR},1,1),{BA},"<="&DATE({YR},12,31),{BN},"<>Cancelled",{BD},{col}$4))'
tx["B11"] = "EXPENSES"
tx["B11"].font = Font(name=F, size=10, bold=True, color=CORAL)
for i in range(10):
    r = 12 + i
    tx[f"B{r}"] = f'=IF(Start!$G${LIST_TOP + i}="","",Start!$G${LIST_TOP + i})'
    tx[f"C{r}"] = f'=IF(B{r}="","",SUMIFS({EXF},{EXA},">="&DATE({YR},1,1),{EXA},"<="&DATE({YR},12,31),Expenses!$C${FIRST}:$C${LAST},B{r}))'
    for col in "DE":
        tx[f"{col}{r}"] = f'=IF(OR(B{r}="",{col}$4=""),"",SUMIFS({EXF},{EXA},">="&DATE({YR},1,1),{EXA},"<="&DATE({YR},12,31),Expenses!$C${FIRST}:$C${LAST},B{r},{EXB},{col}$4))'
tx["B22"] = "Shared costs (all properties)"
tx["C22"] = f'=SUMIFS({EXF},{EXA},">="&DATE({YR},1,1),{EXA},"<="&DATE({YR},12,31),{EXB},"Shared")'
tx["B23"] = "Total expenses"
tx["C23"] = "=SUM(C12:C21)"
tx["D23"] = "=SUM(D12:D21)"
tx["E23"] = "=SUM(E12:E21)"
tx["B25"] = "NET INCOME (payouts minus expenses)"
tx["C25"] = "=C8-C23"
tx["D25"] = '=IF(D4="","",D8-D23)'
tx["E25"] = '=IF(E4="","",E8-E23)'
for r in range(6, 26):
    for col in "BCDE":
        tx[f"{col}{r}"].font = Font(name=F, size=10, bold=(r in (23, 25)))
        if col != "B":
            tx[f"{col}{r}"].number_format = MONEY
        if r in (23, 25):
            tx[f"{col}{r}"].fill = fill(SAND)
tx["B25"].font = Font(name=F, size=11, bold=True, color=NAVY)
tx["B27"] = "Per-property columns show costs logged to that property; shared costs appear in the All properties column only. This summary organizes your records and is not tax advice."
tx["B27"].font = Font(name=F, size=9, italic=True, color="5B6F7D")
tx.merge_cells("B27:G28")
tx["B27"].alignment = Alignment(wrap_text=True, vertical="top")

# print setup for every sheet
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
db.print_area = "A1:M58"
st.print_area = "A1:J28"
pr.print_area = "A1:G24"
tx.print_area = "A1:G28"
import os
if os.environ.get("PREVIEW"):
    for ws, area in ((bk, "A1:O32"), (ex, "A1:G32"), (cl, "A1:I32"), (inv, "A1:K24"), (mt, "A1:I12"), (rv, "A1:G20")):
        ws.print_area = area

# order & finish
wb._sheets = [wb["Start"], wb["Dashboard"], wb["Bookings"], wb["Expenses"], wb["Cleaning"], wb["Inventory"], wb["Maintenance"], wb["Pricing"], wb["Taxes"], wb["Reviews"]]
wb.active = 0
wb.save(OUT)
print("saved", OUT, "bookings:", N_BOOK, "expenses:", len(ex_rows))
