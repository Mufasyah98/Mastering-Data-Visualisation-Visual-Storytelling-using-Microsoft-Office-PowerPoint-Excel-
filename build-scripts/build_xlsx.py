"""Build Day_2_Excel_Exercises.xlsx and Capstone_Excel_Dataset.xlsx."""
import json, random
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

D = json.load(open("/home/claude/work/data.json"))
TEAL, TEAL_L, ORANGE, GREY, INK = "0F766E", "CCFBF1", "EA580C", "F1F5F9", "1F2937"
F = lambda **k: Font(name="Arial", **{"size": 10, **k})
FILL = lambda c: PatternFill("solid", fgColor=c)
thin = Side(style="thin", color="CBD5E1")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
dash = Side(style="dashed", color="0F766E")
DBOX = Border(left=dash, right=dash, top=dash, bottom=dash)
WRAP = Alignment(wrap_text=True, vertical="top")
COMPANY = "Teratai Distribution Sdn Bhd"
HEAD = ["Order ID", "Order Date", "Region", "Product Category", "Order Value (MYR)",
        "Promised Delivery Date", "Actual Delivery Date", "Order Status"]
DATEF, MONEYF = "dd-mmm-yyyy", "#,##0.00"


def d(s):
    return date.fromisoformat(s) if s else None


def title(ws, text, sub=None, width_cols=10):
    ws["A1"] = text
    ws["A1"].font = F(size=16, bold=True, color=TEAL)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F(size=10, italic=True, color="475569")


def section(ws, r, text, ncols=8):
    for c in range(1, ncols + 1):
        ws.cell(r, c).fill = FILL(TEAL)
    ws.cell(r, 1, text).font = F(size=11, bold=True, color="FFFFFF")
    return r + 1


def table_rows(ws, r, header, rows, widths_ok=True, wrap=True, ncols=None):
    for j, h in enumerate(header, 1):
        c = ws.cell(r, j, h)
        c.font = F(bold=True, color=INK)
        c.fill = FILL(TEAL_L)
        c.border = BOX
        c.alignment = WRAP
    r += 1
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            c = ws.cell(r, j, v)
            c.font = F()
            c.border = BOX
            c.alignment = WRAP
            if i % 2:
                c.fill = FILL("F8FAFC")
        r += 1
    return r + 1


def para(ws, r, text, bold=False, merge_to=8, height=None, color=INK):
    ws.cell(r, 1, text).font = F(bold=bold, color=color)
    ws.cell(r, 1).alignment = WRAP
    if merge_to > 1:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_to)
    if height:
        import math
        ws.row_dimensions[r].height = max(15, 13.5 * math.ceil(len(text) / 130) + 3)
    return r + 1


FIELDS = [
    ("Order ID", "Text", "Unique order reference, format ORD-26xxxx. One row per order."),
    ("Order Date", "Date", "Date the customer order was received."),
    ("Region", "Text", "Delivery region: Central, Northern, Southern, East Coast, East Malaysia."),
    ("Product Category", "Text", "Beverages, Packaged Food, Personal Care, Household Care."),
    ("Order Value (MYR)", "Number", "Invoice value of the order in Malaysian Ringgit, before SST."),
    ("Promised Delivery Date", "Date", "Order Date + standard lead time (calendar days): Central 2, Northern 3, Southern 3, East Coast 4, East Malaysia 7."),
    ("Actual Delivery Date", "Date", "Date proof of delivery was signed. Blank if the order is still in transit."),
    ("Order Status", "Text", "Delivered (completed) or In Transit (not yet completed at data extract date)."),
]
KPIS = [
    ("Total Orders", "Count of unique Order IDs"),
    ("Total Order Value (MYR)", "Sum of Order Value (MYR)"),
    ("Average Order Value (MYR)", "Total Order Value ÷ Total Orders"),
    ("Completed Orders", "Orders with Order Status = Delivered (has an Actual Delivery Date)"),
    ("On-Time Delivery Rate", "On-time deliveries ÷ Completed Orders. On time = Actual Delivery Date on or before Promised Delivery Date. In-transit orders are excluded."),
]


# =====================================================================
# DAY 2 WORKBOOK
# =====================================================================
def build_day2():
    rows = [dict(r) for r in D["h1"]]
    rng = random.Random(55)
    n = len(rows)
    idx = list(range(n))
    rng.shuffle(idx)
    pool = iter(idx)
    take = lambda k: [next(pool) for _ in range(k)]

    region_variants = {"central": 3, "CENTRAL": 2, "Northern Region": 2, "southern": 2, "E. Coast": 2, "East Msia": 3}
    canon = {"central": "Central", "CENTRAL": "Central", "Northern Region": "Northern", "southern": "Southern",
             "E. Coast": "East Coast", "East Msia": "East Malaysia"}
    issues = {}
    for variant, k in region_variants.items():
        cands = [i for i in idx if rows[i]["region"] == canon[variant] and i not in issues]
        for i in cands[:k]:
            rows[i]["region_raw"] = variant
            issues[i] = "region"
    trail_reg = [i for i in idx if i not in issues][:10]
    for i in trail_reg:
        rows[i]["region_raw"] = rows[i]["region"] + " "
        issues[i] = "trail_region"
    trail_cat = [i for i in idx if i not in issues][:6]
    for i in trail_cat:
        rows[i]["category_raw"] = rows[i]["category"] + "  "
        issues[i] = "trail_cat"
    txt_od = [i for i in idx if i not in issues][:9]
    for i in txt_od:
        rows[i]["od_text"] = True
        issues[i] = "text_date"
    txt_ad = [i for i in idx if i not in issues and rows[i]["actual"]][:4]
    for i in txt_ad:
        rows[i]["ad_text"] = True
        issues[i] = "text_actual"
    dup_src = [i for i in idx if i not in issues][:6]

    raw = list(rows)
    # insert duplicate copies a few rows after their originals
    for i in sorted(dup_src, reverse=True):
        raw.insert(i + 1 + rng.randint(3, 25), dict(rows[i]))

    wb = Workbook()
    # ---------------- Start Here ----------------
    ws = wb.active
    ws.title = "Start Here"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [26, 16, 18, 18, 18, 18, 18, 18]):
        ws.column_dimensions[col].width = w
    title(ws, "Day 2 · Excel Exercises — Order & Delivery Performance Dashboard",
          f"Mastering Data Visualisation & Visual Storytelling using Microsoft Office · Trainer: Fakhrul Syahmi · Case company: {COMPANY} (fictional)")
    r = 4
    r = section(ws, r, "1. Business scenario")
    r = para(ws, r, f"{COMPANY} distributes fast-moving consumer goods to retailers across Malaysia from its distribution centre in Shah Alam. "
             "The Operations Director wants a one-page Excel dashboard that shows monthly order volume, order value and on-time delivery "
             "performance by region for January–June 2026. The system export in 'Raw Orders' has data-quality problems that must be fixed first.", height=58)
    r += 1
    r = section(ws, r, "2. How to use this workbook")
    for t in ["• Work through the sheets left to right: Raw Orders → Cleaned Orders → Pivot Practice → Dashboard.",
              "• Never edit 'Raw Orders' directly. Copy the data to 'Cleaned Orders' and clean it there, so you can always go back to the source.",
              "• Cells shaded light orange are for you to fill in. Dashed teal boxes are placeholders for charts, slicers and timelines.",
              "• Use the checkpoints below to confirm your numbers before moving on. If a checkpoint does not match, ask the trainer."]:
        r = para(ws, r, t, height=28)
    r += 1
    r = section(ws, r, "3. Field definitions (same in every workshop file)")
    r = table_rows(ws, r, ["Field", "Data type", "Definition"], [list(x) for x in FIELDS])
    for rr in range(r - len(FIELDS) - 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=8)
        ws.row_dimensions[rr].height = 28
    r = section(ws, r, "4. KPI definitions")
    r = table_rows(ws, r, ["KPI", "Definition"], [list(x) for x in KPIS])
    for rr in range(r - len(KPIS) - 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=8)
        ws.row_dimensions[rr].height = 28
    r = section(ws, r, "5. Exercise sequence")
    ex = [
        ("Ex 1", "Clean the source data", "Cleaned Orders", "Copy, remove duplicates, trim spaces, standardise regions, convert text dates, create an Excel Table and helper columns.", "40 min"),
        ("Ex 2", "Build PivotTables & PivotCharts", "Pivot Practice", "Answer six business questions with PivotTables, then turn two of them into PivotCharts.", "50 min"),
        ("Ex 3", "Add filters", "Pivot Practice / Dashboard", "Insert a Region slicer and an Order Date timeline and connect them to every PivotTable.", "20 min"),
        ("Ex 4", "Assemble the dashboard", "Dashboard", "Place KPI cards, two PivotCharts, the slicer and the timeline in the starter layout.", "40 min"),
        ("Ex 5", "Excel chart → PowerPoint", "Dashboard + Day 1 deck", "Paste one chart into PowerPoint twice (linked and embedded) and compare the behaviour.", "20 min"),
    ]
    r = table_rows(ws, r, ["Exercise", "Task", "Sheet", "What you will do", "Time"], [list(x) for x in ex])
    for rr in range(r - len(ex) - 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=7)
        ws.row_dimensions[rr].height = 30
    r = section(ws, r, "6. Exercise 1 — Cleaning steps (work in 'Cleaned Orders')")
    steps = [
        "1. In 'Raw Orders', select A1:H307 and copy. In 'Cleaned Orders', click A1 and Paste Values (Home ▸ Paste ▸ Values) so the headers land in row 1.",
        "2. Remove duplicate orders: Data ▸ Remove Duplicates ▸ tick Order ID only. Note how many duplicates Excel removed.",
        "3. Remove trailing spaces from Region and Product Category: use =TRIM() in a spare column, then Paste Values back over the original column (or use Find & Replace on the double space).",
        "4. Standardise Region labels: filter the Region column and use Find & Replace (Match entire cell contents) so only the five official region names remain.",
        "5. Convert dates stored as text: filter Order Date and Actual Delivery Date for left-aligned values, then use =DATEVALUE(cell) or Data ▸ Text to Columns ▸ Finish (date format YMD). Format all date columns as dd-mmm-yyyy.",
        "6. Click inside the data and press Ctrl+T (My table has headers). Name the table tblOrders (Table Design ▸ Table Name).",
        "7. Add helper columns to the table: Completed = IF([@[Order Status]]=\"Delivered\",1,0); On-Time Flag = IF([@Completed]=1,IF([@[Actual Delivery Date]]<=[@[Promised Delivery Date]],1,0),\"\"); Days Late = IF([@Completed]=1,MAX(0,[@[Actual Delivery Date]]-[@[Promised Delivery Date]]),\"\").",
        "Tip: Power Query (Data ▸ From Table/Range) can do steps 2–5 in one repeatable query: Remove Duplicates, Format ▸ Trim, Replace Values and Change Type ▸ Date.",
    ]
    for t in steps:
        r = para(ws, r, t, height=30)
    r += 1
    r = section(ws, r, "7. Checkpoints — confirm before moving on")
    cps = [
        ("After step 1", "Rows of data in Cleaned Orders (excluding header)", "306"),
        ("After step 2", "Duplicate rows removed / rows remaining", "6 removed · 300 remaining"),
        ("After step 4", "Distinct Region values", "5 (Central, Northern, Southern, East Coast, East Malaysia)"),
        ("After step 4", "Distinct Product Category values", "4"),
        ("After step 5", "Order Date cells that are still text (use =ISTEXT)", "0"),
        ("After step 7", "Completed orders (Order Status = Delivered)", "Check with the trainer"),
    ]
    r = table_rows(ws, r, ["When", "Check", "Expected"], [list(x) for x in cps])
    for rr in range(r - len(cps) - 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=8)
    r = section(ws, r, "8. Exercise 5 — Move an Excel chart into PowerPoint: linking vs embedding")
    for t in [
        "1. Copy your monthly PivotChart (or a normal chart) from the Dashboard sheet.",
        "2. In PowerPoint, on a blank slide: Home ▸ Paste ▸ Paste Special ▸ 'Use Destination Theme & Link Data'. Label it LINKED.",
        "3. On a second slide: Paste ▸ 'Use Destination Theme & Embed Workbook'. Label it EMBEDDED.",
        "4. Back in Excel, change one Order Value in Cleaned Orders (e.g. add 100,000 to one order), refresh the PivotTables, then check both slides (for the linked chart: Chart Design ▸ Refresh Data).",
        "5. Undo your test change. Complete the comparison table below.",
    ]:
        r = para(ws, r, t, height=28)
    r += 1
    cmp_rows = [["Does the chart update when the Excel data changes?", "", ""],
                ["What happens if you email only the .pptx file?", "", ""],
                ["Effect on .pptx file size", "", ""],
                ["Can you edit the chart data from inside PowerPoint?", "", ""],
                ["When would you use it at work?", "", ""]]
    r0 = r
    r = table_rows(ws, r, ["Question", "Linked", "Embedded"], cmp_rows)
    for rr in range(r0 + 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=8)
        ws.row_dimensions[rr].height = 30
        for c in range(2, 9):
            ws.cell(rr, c).fill = FILL("FFEDD5")
    ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=4)
    ws.merge_cells(start_row=r0, start_column=5, end_row=r0, end_column=8)
    ws.cell(r0, 5).value = "Embedded"

    # ---------------- Raw Orders ----------------
    ws = wb.create_sheet("Raw Orders")
    for j, h in enumerate(HEAD, 1):
        c = ws.cell(1, j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL("475569")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, o in enumerate(raw, 2):
        od = o["order_date"] if o.get("od_text") else d(o["order_date"])
        ad = (o["actual"] if o.get("ad_text") else d(o["actual"]))
        vals = [o["order_id"], od, o.get("region_raw", o["region"]), o.get("category_raw", o["category"]),
                o["value"], d(o["promised"]), ad, o["status"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            c.font = F()
            if j in (2, 6, 7) and isinstance(v, date):
                c.number_format = DATEF
            if j == 5:
                c.number_format = MONEYF
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{len(raw)+1}"
    for col, w in zip("ABCDEFGH", [14, 14, 16, 18, 18, 20, 20, 14]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 30
    ws["J1"] = "Source: order system export, extracted 30-Jun-2026. Do not edit this sheet — copy it to 'Cleaned Orders'."
    ws["J1"].font = F(italic=True, color="475569")

    # ---------------- Cleaned Orders ----------------
    ws = wb.create_sheet("Cleaned Orders")
    helper = ["Completed", "On-Time Flag", "Days Late"]
    for j, h in enumerate(HEAD + helper, 1):
        c = ws.cell(1, j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL(TEAL if j <= 8 else ORANGE)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30
    for col, w in zip("ABCDEFGHIJK", [14, 14, 16, 18, 18, 20, 20, 14, 12, 12, 11]):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["L"].width = 3
    ws.column_dimensions["M"].width = 70
    notes = [
        ("WORKSPACE — Exercise 1", True),
        ("Paste the Raw Orders data (values only) starting at cell A1, replacing these headers.", False),
        ("Teal headers = source fields. Orange headers = helper columns you add in step 7.", False),
        ("Convert the range to an Excel Table named tblOrders (Ctrl+T).", False),
        ("", False),
        ("Helper column formulas (structured references):", True),
        ('Completed:  =IF([@[Order Status]]="Delivered",1,0)', False),
        ('On-Time Flag:  =IF([@Completed]=1,IF([@[Actual Delivery Date]]<=[@[Promised Delivery Date]],1,0),"")', False),
        ('Days Late:  =IF([@Completed]=1,MAX(0,[@[Actual Delivery Date]]-[@[Promised Delivery Date]]),"")', False),
        ("", False),
        ("Self-check formulas (type them in a cell to the right of the table):", True),
        ("Rows:  =ROWS(tblOrders)   ·   Distinct regions:  =SUMPRODUCT(1/COUNTIF(tblOrders[Region],tblOrders[Region]))", False),
        ("Text dates left:  =SUMPRODUCT(--ISTEXT(tblOrders[Order Date]))", False),
        ("On-time rate:  =SUM(tblOrders[On-Time Flag])/SUM(tblOrders[Completed])", False),
    ]
    for k, (t, b) in enumerate(notes, 1):
        c = ws.cell(k, 13, t)
        c.font = F(bold=b, color=TEAL if b else INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if t:
            c.fill = FILL("F0FDFA")
    ws.freeze_panes = "A2"

    # ---------------- Pivot Practice ----------------
    ws = wb.create_sheet("Pivot Practice")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [6, 44, 22, 22, 26, 22, 24, 30]):
        ws.column_dimensions[col].width = w
    title(ws, "Exercise 2 & 3 — PivotTables, PivotCharts and filters",
          "Source for every PivotTable: tblOrders (Insert ▸ PivotTable ▸ Table/Range: tblOrders ▸ Existing Worksheet).")
    hdr = ["#", "Business question", "Rows", "Columns", "Values", "Filters", "Build hint", "Your answer"]
    qs = [
        ["Q1", "Which region generated the highest total order value in Jan–Jun 2026?", "Region", "—", "Sum of Order Value (MYR)", "—", "Sort descending. Number format #,##0.", ""],
        ["Q2", "How did the number of orders and order value change month by month?", "Order Date (grouped by Months)", "—", "Count of Order ID; Sum of Order Value (MYR)", "—", "Right-click a date ▸ Group ▸ Months. Make a PivotChart (line or combo).", ""],
        ["Q3", "Which product category contributes most value in each region?", "Region", "Product Category", "Sum of Order Value (MYR)", "—", "Show Values As ▸ % of Row Total.", ""],
        ["Q4", "What is the on-time delivery rate for each region?", "Region", "—", "Average of On-Time Flag", "Order Status = Delivered", "Format as 0.0%. Make a PivotChart (sorted bar).", ""],
        ["Q5", "Is on-time delivery improving or getting worse month by month?", "Order Date (Months)", "—", "Average of On-Time Flag; Count of Order ID", "Order Status = Delivered", "Count tells you how many completed orders sit behind each rate.", ""],
        ["Q6", "Which product category has the highest average order value?", "Product Category", "—", "Average of Order Value (MYR)", "—", "Sort descending.", ""],
        ["Q7", "How late are late deliveries, by region?", "Region", "—", "Average of Days Late; Max of Days Late", "Order Status = Delivered", "Optional stretch question.", ""],
    ]
    r = 4
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r, j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL(TEAL)
        c.border = BOX
        c.alignment = WRAP
    for q in qs:
        r += 1
        for j, v in enumerate(q, 1):
            c = ws.cell(r, j, v)
            c.font = F(bold=(j == 1))
            c.border = BOX
            c.alignment = WRAP
            if j == 8:
                c.fill = FILL("FFEDD5")
        ws.row_dimensions[r].height = 44
    r += 2
    ws.cell(r, 1, "Exercise 3 — Filters").font = F(bold=True, size=11, color=TEAL)
    for t in ["1. Click any PivotTable ▸ PivotTable Analyze ▸ Insert Slicer ▸ Region.",
              "2. PivotTable Analyze ▸ Insert Timeline ▸ Order Date. Set the timeline to MONTHS.",
              "3. Right-click the slicer ▸ Report Connections ▸ tick every PivotTable. Repeat for the timeline.",
              "4. Test: select East Malaysia only and Apr–Jun 2026. Every PivotTable should change together.",
              "5. Cut the slicer and timeline and paste them on the Dashboard sheet in Exercise 4."]:
        r += 1
        ws.cell(r, 2, t).font = F()
    r += 2
    ws.cell(r, 1, "PivotTable workspace — build Q1 at B" + str(r + 2) + ", Q2 at E" + str(r + 2) + ", and continue below. Leave 3 blank columns between PivotTables.").font = F(italic=True, color="475569")
    for k, (cell, lab) in enumerate([(f"B{r+2}", "Q1 PivotTable here"), (f"E{r+2}", "Q2 PivotTable here"),
                                     (f"B{r+14}", "Q3 PivotTable here"), (f"E{r+14}", "Q4 PivotTable here"),
                                     (f"B{r+26}", "Q5 PivotTable here"), (f"E{r+26}", "Q6 PivotTable here"),
                                     (f"B{r+38}", "Q7 PivotTable here (optional)")]):
        ws[cell] = lab
        ws[cell].font = F(bold=True, color=TEAL)
        ws[cell].fill = FILL("F0FDFA")
        ws[cell].border = DBOX

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    for c in range(1, 27):
        ws.column_dimensions[get_column_letter(c)].width = 6.5
    for rr in range(1, 45):
        ws.row_dimensions[rr].height = 18
    ws.merge_cells("B2:Y3")
    ws["B2"] = "Teratai Distribution · Order & Delivery Performance · Jan–Jun 2026"
    ws["B2"].font = F(size=18, bold=True, color="FFFFFF")
    ws["B2"].fill = FILL(TEAL)
    ws["B2"].alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("B4:Y4")
    ws["B4"] = "Headline (write your one-sentence takeaway here once the dashboard is built)"
    ws["B4"].font = F(italic=True, color=ORANGE, size=11)

    def box(rng_, label, sub="", fill="F0FDFA", border=DBOX, big=False):
        ws.merge_cells(rng_)
        tl = rng_.split(":")[0]
        c = ws[tl]
        c.value = label + ("\n" + sub if sub else "")
        c.font = F(bold=True, color=TEAL, size=12 if big else 10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        from openpyxl.utils.cell import range_boundaries
        c1, r1, c2, r2 = range_boundaries(rng_)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).fill = FILL(fill)
                b = Border(left=dash if cc == c1 else None, right=dash if cc == c2 else None,
                           top=dash if rr == r1 else None, bottom=dash if rr == r2 else None)
                ws.cell(rr, cc).border = b

    box("B6:G9", "KPI 1 · Total Orders", "=GETPIVOTDATA or link to a PivotTable", big=True)
    box("I6:N9", "KPI 2 · Total Order Value (MYR)", "Format: MYR #,##0", big=True)
    box("P6:S9", "KPI 3 · Average Order Value (MYR)", "Total value ÷ orders", big=True)
    box("U6:Y9", "KPI 4 · On-Time Delivery Rate", "Completed orders only · 0.0%", big=True)
    box("B11:E26", "SLICER", "Region\n(from Exercise 3)", fill="FFF7ED")
    box("G11:Y14", "TIMELINE", "Order Date · Months (from Exercise 3)", fill="FFF7ED")
    box("G16:P31", "CHART 1", "Monthly orders & order value\n(PivotChart from Q2 · line or combo)")
    box("R16:Y31", "CHART 2", "On-time delivery rate by region\n(PivotChart from Q4 · sorted bar)")
    box("B28:E31", "Notes", "Data: order system extract, 30-Jun-2026", fill=GREY)
    box("B33:Y36", "Optional CHART 3 — Order value by category (Q3 or Q6)", "")
    ws["B38"] = "Exercise 4 — Assemble the dashboard"
    ws["B38"].font = F(bold=True, size=11, color=TEAL)
    for k, t in enumerate([
        "1. KPI cards: in each box type = and click the matching PivotTable grand-total cell (Excel writes a GETPIVOTDATA formula). Format the number; font 20pt+.",
        "2. Move PivotCharts from Pivot Practice into the CHART 1 and CHART 2 boxes (cut & paste). Hide field buttons: PivotChart Analyze ▸ Field Buttons ▸ Hide All.",
        "3. Paste the Region slicer and Order Date timeline into their boxes. Resize to fit. Check Report Connections again.",
        "4. Replace each placeholder with your object, then delete the placeholder text. Write the headline in row 4. Test the slicer: do all KPIs and charts change?",
        "5. Final check: View ▸ uncheck Gridlines and Headings; File ▸ Print Preview fits on one landscape page.",
    ], 39):
        ws.cell(k, 2, t).font = F()

    for w_ in wb.worksheets:
        w_.page_setup.orientation = "landscape"; w_.page_setup.fitToWidth = 1; w_.page_setup.fitToHeight = 0; w_.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save("/home/claude/work/out/Day_2_Excel_Exercises.xlsx")
    return {"raw_rows": len(raw), "dups": len(dup_src), "region_fix": sum(region_variants.values()),
            "trail_region": len(trail_reg), "trail_cat": len(trail_cat), "text_od": len(txt_od), "text_ad": len(txt_ad),
            "dup_ids": [rows[i]["order_id"] for i in dup_src],
            "variants": region_variants}


# =====================================================================
# CAPSTONE WORKBOOK
# =====================================================================
def build_capstone():
    rows = D["q3"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Management Brief"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [26, 16, 18, 18, 18, 18, 18, 18]):
        ws.column_dimensions[col].width = w
    title(ws, "Capstone · Delivery Performance Review — Q3 2026 (July–September)",
          f"Group exercise · {COMPANY} (fictional) · Trainer: Fakhrul Syahmi")
    r = 4
    r = section(ws, r, "Management brief")
    r = para(ws, r, "From: Puan Rosnah Idris, Operations Director   ·   To: Analytics working groups", bold=True)
    r = para(ws, r, "Customer complaints about late deliveries rose during the quarter, and the Board meets in two weeks. "
             "I need to know where delivery performance needs attention, how big the problem is, and what one action we should take first. "
             "Please give me facts from the order data, not opinions. Show me a short dashboard and three slides.", height=58)
    r += 1
    r = section(ws, r, "Your task")
    for t in ["1. Identify where delivery performance needs attention (region, month, category — wherever the evidence points).",
              "2. Support your finding with evidence from the data: rates, counts, trend, and comparison with other regions.",
              "3. Recommend ONE action management should take, and say how they would know it worked."]:
        r = para(ws, r, t, height=20)
    r += 1
    r = section(ws, r, "Business questions to guide your analysis")
    for t in ["Q1. How many orders and how much order value (MYR) did each region generate in Q3, and how did volume move by month?",
              "Q2. What was the overall on-time delivery rate for completed orders in Q3? How does each region compare?",
              "Q3. Did on-time performance change from July to September? In which region(s)?",
              "Q4. When orders are late, how late are they (average days late) and does that differ by region?",
              "Q5. Is the problem concentrated in a particular product category or spread across all categories?",
              "Q6. What does the data NOT tell you? What would you need to investigate before acting?"]:
        r = para(ws, r, t, height=20)
    r += 1
    r = section(ws, r, "Submission requirements (per group)")
    reqs = [
        ("Mini dashboard (sheet 'Mini Dashboard')", "Three KPIs: Total Orders, Total Order Value (MYR), On-Time Delivery Rate. Two charts (PivotCharts). One Region slicer connected to all PivotTables."),
        ("Presentation (Capstone_PowerPoint_Template.pptx)", "Three slides — Context, Insight, Recommendation. The Insight slide must include a chart and a takeaway headline."),
        ("Time", "Analysis 75 minutes · presentation 5 minutes per group · Q&A and peer feedback 3 minutes."),
        ("Accuracy rule", "On-time rate = on-time deliveries ÷ completed orders (Order Status = Delivered). Exclude In Transit orders. State the number of orders behind every rate."),
    ]
    r = table_rows(ws, r, ["Deliverable", "Requirement"], [list(x) for x in reqs])
    for rr in range(r - len(reqs) - 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=8)
        ws.row_dimensions[rr].height = 32
    r = section(ws, r, "Field definitions (identical to the Day 2 workbook)")
    r = table_rows(ws, r, ["Field", "Data type", "Definition"], [list(x) for x in FIELDS])
    for rr in range(r - len(FIELDS) - 1, r - 1):
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=8)
        ws.row_dimensions[rr].height = 28
    r = section(ws, r, "Data notes")
    for t in [f"• 'Orders Q3' contains {len(rows)} orders dated 1-Jul-2026 to 30-Sep-2026, extracted 30-Sep-2026. The data has already been cleaned (no duplicates, standard labels, real dates).",
              "• The data is an Excel Table named tblQ3 with helper columns Completed, On-Time Flag and Days Late already calculated.",
              "• Order IDs continue from the Jan–Jun dataset (ORD-260301 onwards)."]:
        r = para(ws, r, t, height=28)

    # ---- Orders Q3 ----
    ws = wb.create_sheet("Orders Q3")
    head = HEAD + ["Completed", "On-Time Flag", "Days Late"]
    for j, h in enumerate(head, 1):
        c = ws.cell(1, j, h)
        c.font = F(bold=True, color="FFFFFF")
    n = len(rows)
    for i, o in enumerate(rows, 2):
        vals = [o["order_id"], d(o["order_date"]), o["region"], o["category"], o["value"],
                d(o["promised"]), d(o["actual"]), o["status"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            c.font = F()
            if j in (2, 6, 7):
                c.number_format = DATEF
            if j == 5:
                c.number_format = MONEYF
        ws.cell(i, 9, f'=IF(H{i}="Delivered",1,0)').font = F()
        ws.cell(i, 10, f'=IF(I{i}=1,IF(G{i}<=F{i},1,0),"")').font = F()
        ws.cell(i, 11, f'=IF(I{i}=1,MAX(0,G{i}-F{i}),"")').font = F()
    tab = Table(displayName="tblQ3", ref=f"A1:K{n+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHIJK", [14, 14, 16, 18, 18, 20, 20, 14, 12, 13, 11]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 30

    # ---- Analysis Workspace ----
    ws = wb.create_sheet("Analysis Workspace")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 60
    ws["B1"] = "Analysis workspace — build your PivotTables here (source: tblQ3)"
    ws["B1"].font = F(size=14, bold=True, color=TEAL)
    for k, t in enumerate(["Suggested PivotTables:",
                           "• Region × Count of Order ID, Sum of Order Value (MYR)",
                           "• Region (rows) × Order Date grouped by Months (columns) × Average of On-Time Flag  [Filter: Order Status = Delivered]",
                           "• Region × Average of Days Late  [Filter: Order Status = Delivered, On-Time Flag = 0]",
                           "• Product Category × Average of On-Time Flag, with the Region slicer set to your focus region",
                           "Always add Count of Order ID next to a rate so you know how many orders it is based on."], 3):
        ws.cell(k, 2, t).font = F(bold=(k == 3))
    ws["B11"] = "Start PivotTables from cell B12 →"
    ws["B11"].font = F(italic=True, color=ORANGE)

    # ---- Mini Dashboard ----
    ws = wb.create_sheet("Mini Dashboard")
    ws.sheet_view.showGridLines = False
    for c in range(1, 27):
        ws.column_dimensions[get_column_letter(c)].width = 6.5
    ws.merge_cells("B2:Y3")
    ws["B2"] = "Teratai Distribution · Delivery Performance · Q3 2026 · Group ____"
    ws["B2"].font = F(size=18, bold=True, color="FFFFFF")
    ws["B2"].fill = FILL(TEAL)
    ws["B2"].alignment = Alignment(vertical="center", indent=1)
    from openpyxl.utils.cell import range_boundaries

    def box(rng_, label, sub="", fill="F0FDFA"):
        ws.merge_cells(rng_)
        c1, r1, c2, r2 = range_boundaries(rng_)
        c = ws.cell(r1, c1)
        c.value = label + ("\n" + sub if sub else "")
        c.font = F(bold=True, color=TEAL, size=11)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).fill = FILL(fill)
                ws.cell(rr, cc).border = Border(left=dash if cc == c1 else None, right=dash if cc == c2 else None,
                                                top=dash if rr == r1 else None, bottom=dash if rr == r2 else None)
    ws.merge_cells("B4:Y4")
    ws["B4"] = "Takeaway headline: ______________________________________________"
    ws["B4"].font = F(italic=True, color=ORANGE, size=11)
    box("B6:H9", "KPI 1 · Total Orders")
    box("J6:Q9", "KPI 2 · Total Order Value (MYR)")
    box("S6:Y9", "KPI 3 · On-Time Delivery Rate", "completed orders only")
    box("B11:E28", "REGION SLICER", fill="FFF7ED")
    box("G11:P28", "CHART 1", "e.g. on-time rate by month and region")
    box("R11:Y28", "CHART 2", "e.g. orders or days late by region")
    ws["B30"] = "Checklist: KPIs link to PivotTables · slicer connected to all PivotTables · chart titles state the message · order counts shown next to rates"
    ws["B30"].font = F(italic=True, color="475569")

    for w_ in wb.worksheets:
        w_.page_setup.orientation = "landscape"; w_.page_setup.fitToWidth = 1; w_.page_setup.fitToHeight = 0; w_.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save("/home/claude/work/out/Capstone_Excel_Dataset.xlsx")


import os
os.makedirs("/home/claude/work/out", exist_ok=True)
meta = build_day2()
build_capstone()
json.dump(meta, open("/home/claude/work/issues.json", "w"), indent=1)
print(meta)
