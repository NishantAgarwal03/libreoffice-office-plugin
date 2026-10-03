"""Generate samples/office_plugin_test_data.xlsx: one sheet of messy test data per plugin feature.

Run: python tools/make_sample_workbook.py   (needs openpyxl; not part of the extension)
Each sheet starts at A1 with a header row so a lone click inside the table selects it.
The README sheet says which command to run on which sheet and what to expect.
"""

import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "samples", "office_plugin_test_data.xlsx")
BOLD = Font(bold=True)
HEAD = PatternFill("solid", fgColor="DDEBF7")


def sheet(wb, name, header, rows, widths=None):
    ws = wb.create_sheet(name)
    ws.append(header)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font, c.fill = BOLD, HEAD
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[chr(64 + i)].width = w
    return ws


def build():
    wb = Workbook()
    readme = wb.active
    readme.title = "README"

    # ---- Links (navigation: Alt+Left / Alt+Right) --------------------------------------------------------------
    ws = sheet(wb, "Links", ["Click a link, then press Alt+Left to come back", "Target"], [], [48, 30])
    targets = [("Go to the Dedupe sheet", "#Dedupe.A1"), ("Go to the far cell on this sheet's Far block", "#Far.A900"),
               ("Go to Gantt", "#Gantt.A1"), ("Go to Units", "#Units.A1")]
    for i, (text, ref) in enumerate(targets, 2):
        ws.cell(i, 1, text)
        ws.cell(i, 2, ref)
        ws.cell(i, 1).hyperlink = Hyperlink(ref=ws.cell(i, 1).coordinate, location=ref[1:].replace(".", "!"), display=text)
        ws.cell(i, 1).font = Font(color="0563C1", underline="single")
    far = wb.create_sheet("Far")
    far["A1"] = "Top of the Far sheet"
    far["A900"] = "You jumped to row 900 (Alt+Left returns)"
    far["A900"].font = BOLD

    # ---- Visible (Select visible cells only) -------------------------------------------------------------------
    ws = sheet(wb, "Visible", ["Item", "Qty", "Hidden?"],
               [["Item %d" % i, i * 10, "HIDDEN ROW" if i % 3 == 0 else ""] for i in range(1, 16)], [14, 8, 14])
    for i in range(1, 16):
        if i % 3 == 0:
            ws.row_dimensions[i + 1].hidden = True
    ws.auto_filter.ref = "A1:C16"

    # ---- Dedupe ------------------------------------------------------------------------------------------------
    sheet(wb, "Dedupe", ["Name", "City", "Amount"],
          [["Asha", "Pune", 100], ["Ravi", "Delhi", 250], ["asha", "pune", 100], ["Asha", "Pune", 100],
           ["Meera", "Mumbai", 75], ["Ravi", "Delhi", 250], ["Ravi", "Delhi", 999], ["Sunil", "Chennai", 40],
           ["Meera", "Mumbai", 75]], [14, 14, 10])

    # ---- Lists (Compare Lists + Highlight): List A = A1:B10, List B = D1:E10 --------------------------------------
    ws = sheet(wb, "Lists", ["ID (List A)", "Name (List A)", "", "ID (List B)", "Name (List B)"],
               [[101, "Asha", None, "101", "asha "],
                [102, "Ravi", None, 102, "Ravi"],
                [103, "Meera", None, 104, "Kiran"],
                ["104", "Kiran", None, 105, "Neha"],
                [105, "Neha  Rao", None, 106, "Omar"],
                [107, "Zoe", None, 107, "ZOE"],
                [108, "Liam", None, 109, "Noah"],
                [110, "Ivy", None, 110, "Ivy"]], [12, 16, 4, 12, 16])
    ws["A1"].fill = ws["B1"].fill = ws["D1"].fill = ws["E1"].fill = HEAD
    ws["C1"].fill = PatternFill()

    # ---- Transpose ---------------------------------------------------------------------------------------------
    ws = sheet(wb, "Transpose", ["Region", "Q1", "Q2", "Q3", "Total"],
               [["North", 10, 20, 30, "=SUM(B2:D2)"], ["South", 5, 15, 25, "=SUM(B3:D3)"],
                ["East", 7, 14, 21, "=SUM(B4:D4)"]], [10, 8, 8, 8, 8])
    ws["A2"].fill = PatternFill("solid", fgColor="FFF2CC")

    # ---- Sequence ----------------------------------------------------------------------------------------------
    ws = sheet(wb, "Sequence", ["Fill here (select A2:A21, run Sequence Numbers)", "", "Or here (B2:B6)"], [], [46, 4, 18])
    ws["C1"].font = BOLD

    # ---- MissingSeq --------------------------------------------------------------------------------------------
    inv = [1, 2, 3, 6, 7, 10, 11, 12, 15]
    ws = sheet(wb, "MissingSeq", ["Invoice", "Number", "Amount"],
               [["INV-%04d" % n, n, n * 100] for n in inv], [12, 10, 10])
    ws["B11"] = 12          # a duplicate (reported, never changed)
    ws["A11"] = "INV-0012"
    ws["C11"] = 1200

    # ---- FillBlanks --------------------------------------------------------------------------------------------
    ws = sheet(wb, "FillBlanks", ["Category", "Item", "Price"],
               [["Fruit", "Apple", 3], [None, "Banana", 2], [None, "Cherry", 9], ["Veg", "Carrot", 1],
                [None, "Leek", 4], ["Dairy", "Milk", 2], [None, "Cheese", 8], [None, "Butter", 6]], [12, 12, 8])
    ws.merge_cells("A10:A12")
    ws["A10"] = "Bakery"
    ws["B10"], ws["B11"], ws["B12"] = "Bread", "Bun", "Cake"
    ws["C10"], ws["C11"], ws["C12"] = 2, 1, 7

    # ---- RemoveBlanks ------------------------------------------------------------------------------------------
    ws = sheet(wb, "RemoveBlanks", ["A", "B", "C", "D", "E"],
               [[1, 2, None, 4, 5], [None, None, None, None, None], [6, None, None, 9, 10], [11, 12, None, 14, 15],
                [None, None, None, None, None], [16, 17, None, 19, 20], [21, None, None, 24, 25]],
               [8, 8, 8, 8, 8])
    ws["A9"] = "(a formula returning empty text is NOT a blank:)"
    ws["A10"] = '=""'
    ws["B10"] = 5

    # ---- Gantt -------------------------------------------------------------------------------------------------
    d0 = dt.date(2026, 11, 2)
    tasks = [(1, "Site survey", d0, 5, 1.0, None), (2, "Foundation", None, 10, 0.6, "1"),
             (3, "Structure", None, 20, 0.25, "2"), (4, "Electrical rough-in", None, 8, 0, "3"),
             (5, "Plumbing rough-in", None, 8, 0, "3"), (6, "Finishing", None, 15, 0, "4, 5"),
             (7, "Handover (milestone)", None, 0, 0, "6")]
    ws = sheet(wb, "Gantt", ["ID", "Task", "Start", "Duration", "% Complete", "Predecessors"],
               [[i, t, s, d, p, pr] for i, t, s, d, p, pr in tasks], [6, 24, 12, 10, 12, 14])
    for r in range(2, 9):
        ws.cell(r, 3).number_format = "DD-MMM-YY"
        ws.cell(r, 5).number_format = "0%"

    # ---- Decimals ----------------------------------------------------------------------------------------------
    ws = sheet(wb, "Decimals", ["Typed value", "Formula", "Percent", "Other (skipped)"],
               [[3.14159265, "=B$9*2", 0.12345, dt.date(2026, 5, 17)],
                [2.5, "=A3/3", 0.5, "text"], [-7.777, "=ROUND(A4*3;1)", 0.0333, True],
                [1234.5678, "=SUM(A2:A5)", 1.5, "#N/A"], [0.005, "=A6*100", 0.9999, None]], [14, 14, 10, 16])
    ws["B9"] = 1.23456789
    ws["A9"] = "helper used by B2"
    for r in range(2, 7):
        ws.cell(r, 3).number_format = "0.0%"
    ws["B3"] = "=A3/3"
    ws["B4"] = "=ROUND(A4*3,1)"
    ws["B6"] = "=A6*100"

    # ---- NumWords ----------------------------------------------------------------------------------------------
    ws = sheet(wb, "NumWords", ["Amount", "Notes"],
               [[0, ""], [7, ""], [123456, "1,23,456 Indian"], [1234567, ""], [1234.5, "Currency: Rupees / Paise"],
                [-250.75, "negative"], [1e12, "lakh crore"], [0.05, ""], [100000, "one lakh"],
                [dt.date(2026, 1, 15), "date cell: skipped, counted"], [True, "boolean: skipped"],
                ["12,345", "text that looks like a number"]], [18, 34])
    ws["A11"].number_format = "DD-MMM-YY"

    # ---- Units -------------------------------------------------------------------------------------------------
    ws = sheet(wb, "Units", ["Text", "Expected"],
               [["1 lakh", 100000], ["1L", 100000], ["2 cr", 20000000], ["2cr", 20000000], ["1.5k", 1500],
                ["3M", 3000000], ["Rs. 2 crore", 20000000], ["12,34,567", 1234567], ["1 crore 25 lakh", 12500000],
                ["(2.5L)", -250000], ["1,5L", "rejected as ambiguous"], ["abc", "unreadable, reported"],
                ["Rs 4.5 lac", 450000], ["one lakh twenty-three thousand", "use Words to Number instead"]], [34, 26])

    # ---- Words (reverse of NumWords / for the functions) ----------------------------------------------------------
    sheet(wb, "Words", ["Words", "Expected"],
          [["one lakh twenty-three thousand four hundred fifty-six", 123456],
           ["Rupees twelve and Paise fifty Only", 12.5], ["minus two hundred", -200],
           ["one lakh crore", 1e12], ["twenty one", 21], ["banana", "unreadable, reported"],
           ["five million", 5000000]], [54, 24])

    # ---- Functions ---------------------------------------------------------------------------------------------
    ws = sheet(wb, "Functions", ["Input", "=NUMTOWORDS(A)", "=NUMTOWORDS(A;Intl;USD;upper)", "=WORDSTONUM(E)", "Words", "=UNITSTONUM(G)", "Units"],
               [[5, None, None, None, "five lakh", None, "2cr"], [123456, None, None, None, "one lakh crore", None, "1.5k"],
                [1234567.89, None, None, None, "Rupees twelve Only", None, "3M"],
                ["abc", None, None, None, "banana", None, "abc"]], [12, 20, 30, 18, 22, 18, 10])
    ws["A7"] = "After Number Functions > Enable, type in B2: =NUMTOWORDS(A2)  C2: =NUMTOWORDS(A2;\"International\";\"USD\";\"upper\")"
    ws["A8"] = "D2: =WORDSTONUM(E2)   F2: =UNITSTONUM(G2)   then fill down. Row 5 shows the text #VALUE!."

    # ---- README ------------------------------------------------------------------------------------------------
    rows = [
        ("Sheet", "Feature", "What to do", "What you should see"),
        ("Links", "navigation", "Click a link (Ctrl+click), press Alt+Left, then Alt+Right", "Returns to where you were / goes forward again"),
        ("Far", "navigation / last_edit", "Edit A1, then A900, then jump away; press Alt+F5 repeatedly", "Cycles your last 3 edit places"),
        ("Visible", "calc_visible", "Select A1:C16 (rows 4,7,10... hidden), Alt+;, copy, paste elsewhere", "Only visible rows selected/copied"),
        ("Dedupe", "calc_dedupe", "Click in table > Remove Duplicates; tick Name+City (try Ignore case on/off)", "Name+City only: 9 rows -> 4; all three columns, ignore case: -> 5 (asha/pune = Asha/Pune)"),
        ("Lists", "calc_setops + calc_highlight", "A = A1:B9, B = D1:E9; Compare Lists with matching on ID with clean options; Highlight duplicates/differences", "Union/intersection/A-B/B-A on a NEW sheet; 101 matches \"101\"; 'asha ' matches 'Asha'"),
        ("Transpose", "calc_transpose", "Click in table > Transpose (in place; also try new sheet)", "Rows become columns, totals still formulas"),
        ("Sequence", "calc_sequence", "Select A2:A21 > Sequence Numbers, Start 1 Step 1; then B2:B6 Start 0 End 1 (step 0.25)", "1..20; 0, 0.25 ... 1"),
        ("MissingSeq", "calc_missing_seq", "Click in A or B column > Find Missing Numbers", "Gap list (4,5,8,9,13,14); placeholder rows inserted; duplicate 12 reported"),
        ("FillBlanks", "calc_fill_blanks", "Select A2:A12 > Fill Blanks (down)", "Fruit/Veg/Dairy repeat; merged Bakery block unmerged + filled"),
        ("RemoveBlanks", "calc_remove_blanks", "Select A1:E10 > Remove Blanks: try rows, columns, cells", "Empty rows/cols vanish; the formula returning \"\" row is kept"),
        ("Gantt", "calc_gantt", "Click in table > Gantt Chart; map columns (ID, Predecessors)", "Timeline with bars, 60% shading, milestone cell; edit a Duration and re-run"),
        ("Decimals", "calc_decimals", "Select A2:D6 > Set Decimals = 2", "Values rounded, formulas wrapped in ROUND, date/text/boolean/error skipped"),
        ("NumWords", "calc_numwords", "Select A2:A13 > Number to Words (try Indian, International, currency)", "Words in next column; date + TRUE skipped and counted"),
        ("Units", "calc_units", "Select A2:A15 > Units to Numbers", "Numbers in next column; '1,5L' and 'abc' reported"),
        ("Words", "calc_numwords (reverse)", "Select A2:A8 > Words to Number", "Numbers in next column; banana reported"),
        ("Functions", "calc_functions", "Number Functions... > Enable, then type the formulas shown on the sheet", "Live results; row 5 shows text #VALUE!"),
    ]
    for r in rows:
        readme.append(r)
    for c in readme[1]:
        c.font, c.fill = BOLD, HEAD
    for col, w in zip("ABCD", (14, 28, 70, 70)):
        readme.column_dimensions[col].width = w
    for row in readme.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    readme["A19"] = "Tip: work on a COPY of this file; save as .ods/.xlsx freely. Regenerate with tools/make_sample_workbook.py"

    order = ["README", "Links", "Far", "Visible", "Dedupe", "Lists", "Transpose", "Sequence", "MissingSeq", "FillBlanks",
             "RemoveBlanks", "Gantt", "Decimals", "NumWords", "Units", "Words", "Functions"]
    wb._sheets = [wb[n] for n in order]
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    return os.path.normpath(OUT)


if __name__ == "__main__":
    print(build())
