# -*- coding: utf-8 -*-
"""Rebuild Money Management Dashboard with linked dummy data."""
from datetime import date
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart.legend import Legend
from copy import deepcopy

wb = Workbook()
wb.remove(wb.active)

# --- styles ---
NAVY = "1F4E78"
TEAL = "0D7377"
GOLD = "C65911"
GREEN = "548235"
PURPLE = "5B2C6F"
SLATE = "334155"

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_green = PatternFill("solid", fgColor=GREEN)
fill_purple = PatternFill("solid", fgColor=PURPLE)
fill_kpi = PatternFill("solid", fgColor="E8F1FB")
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_calc = PatternFill("solid", fgColor="E2EFDA")
fill_alt = PatternFill("solid", fgColor="F8FAFC")
fill_white = PatternFill("solid", fgColor="FFFFFF")
fill_soft_red = PatternFill("solid", fgColor="FCE4EC")
fill_title_bar = PatternFill("solid", fgColor="0F2744")
fill_note = PatternFill("solid", fgColor="FFF8E7")

font_white_b = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
font_title = Font(name="Calibri", color="FFFFFF", bold=True, size=20)
font_section = Font(name="Calibri", color=NAVY, bold=True, size=14)
font_label = Font(name="Calibri", color="475569", bold=True, size=10)
font_input = Font(name="Calibri", color="9A5700", bold=True, size=12)
font_kpi = Font(name="Calibri", color=NAVY, bold=True, size=16)
font_blue = Font(name="Calibri", color="1D4ED8", size=11)  # formula
font_black = Font(name="Calibri", color="1E293B", size=11)
font_small = Font(name="Calibri", color="64748B", size=9, italic=True)
font_bar = Font(name="Consolas", color=NAVY, size=10)
font_header_note = Font(name="Calibri", color="7C2D12", size=10, italic=True)

thin = Side(style="thin", color="CBD5E1")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_c = Alignment(horizontal="left", vertical="center", wrap_text=True)

EXPENSE_CATS = [
    "Home Rent",
    "Utilities",
    "Food",
    "Supplies",
    "Transportation",
    "Healthcare",
    "Debt",
    "Shopping",
    "Gifts",
]
INCOME_CATS = ["Salary", "Bonus", "Other Income"]
ALL_CATS = EXPENSE_CATS + INCOME_CATS
BANKS = ["Bank 1", "Bank 2", "Bank 3"]
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
YEARS = [2024, 2025, 2026]

# Monthly budget target (same every month) so over/under is obvious
ALLOC = {
    "Home Rent": 5_000_000,
    "Utilities": 1_500_000,
    "Food": 3_000_000,
    "Supplies": 1_000_000,
    "Transportation": 1_200_000,
    "Healthcare": 800_000,
    "Debt": 2_000_000,
    "Shopping": 1_500_000,
    "Gifts": 500_000,
}

# Deterministic dummy expenses per month index 1-12, per bank, per category (rupiah)
# Tuned so December Bank 1 Supplies exceeds 100% (about 170%)
def expense_amount(month, bank_i, cat):
    base = {
        "Home Rent": 5_000_000,
        "Utilities": 900_000 + month * 40_000,
        "Food": 2_200_000 + (month % 3) * 250_000,
        "Supplies": 600_000 + month * 20_000,
        "Transportation": 700_000 + (bank_i * 80_000),
        "Healthcare": 250_000 if month % 2 == 0 else 450_000,
        "Debt": 2_000_000,
        "Shopping": 400_000 + (month % 4) * 300_000,
        "Gifts": 0 if month not in (6, 12) else 750_000,
    }[cat]
    # bank 2 spends a bit less, bank 3 a bit more on food/shopping
    if bank_i == 1:
        base = int(base * 0.85)
    elif bank_i == 2:
        base = int(base * 1.1)
    if cat == "Supplies" and month == 12 and bank_i == 0:
        base = 1_700_000  # 170% of 1,000,000 allocation
    if cat == "Gifts" and month == 12 and bank_i == 0:
        base = 900_000  # over 500k budget
    if cat == "Food" and month == 12 and bank_i == 0:
        base = 3_400_000  # slightly over
    return base


SALARY = {"Bank 1": 25_000_000, "Bank 2": 18_000_000, "Bank 3": 12_000_000}
OPENING = {"Bank 1": 40_000_000, "Bank 2": 22_000_000, "Bank 3": 8_500_000}


def style_header_row(ws, row, start_col, end_col, fill):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font_white_b
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[row].height = 22


def paint_body(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = 20
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.font = font_black
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if (r - r1) % 2 == 1:
                if cell.fill is None or cell.fill.fgColor is None or (
                    getattr(cell.fill.fgColor, "rgb", None) in (None, "00000000")
                ):
                    cell.fill = fill_alt


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def add_list_dv(ws, name_ref, cells):
    dv = DataValidation(type="list", formula1=name_ref, allow_blank=False)
    dv.error = "Pilih dari daftar."
    dv.errorTitle = "Input tidak valid"
    dv.prompt = "Pilih dari dropdown"
    dv.promptTitle = "Pilihan"
    dv.showErrorMessage = True
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(cells)
    return dv


# ============================================================
# SETUP
# ============================================================
setup = wb.create_sheet("Setup")
setup.sheet_properties.tabColor = "64748B"
setup["A1"] = "MASTER LIST — sumber dropdown. Jangan hapus baris ini."
setup["A1"].font = font_section
setup.merge_cells("A1:D1")

headers_s = ["Years", "Months", "Banks", "Categories"]
for i, h in enumerate(headers_s, 1):
    setup.cell(row=3, column=i, value=h)
style_header_row(setup, 3, 1, 4, fill_navy)

for i, y in enumerate(YEARS):
    setup.cell(row=4 + i, column=1, value=y)
for i, m in enumerate(MONTHS):
    setup.cell(row=4 + i, column=2, value=m)
for i, b in enumerate(BANKS):
    setup.cell(row=4 + i, column=3, value=b)
for i, c in enumerate(ALL_CATS):
    setup.cell(row=4 + i, column=4, value=c)

paint_body(setup, 4, 15, 1, 4)
set_widths(setup, {"A": 16, "B": 18, "C": 16, "D": 20})
setup.row_dimensions[1].height = 24
setup.freeze_panes = "A4"
setup.auto_filter.ref = "A3:D15"
setup["A17"] = "Category filter di sheet Data membaca kolom D (Categories) lewat named range CatList."
setup["A17"].font = font_small
setup.merge_cells("A17:D17")

# Named ranges — dropdown tetap jalan meski user pindah sheet
wb.defined_names.add(DefinedName(name="YearList", attr_text="Setup!$A$4:$A$6"))
wb.defined_names.add(DefinedName(name="MonthList", attr_text="Setup!$B$4:$B$15"))
wb.defined_names.add(DefinedName(name="BankList", attr_text="Setup!$C$4:$C$6"))
wb.defined_names.add(DefinedName(name="CatList", attr_text="Setup!$D$4:$D$15"))
wb.defined_names.add(DefinedName(name="TypeList", attr_text='"Income,Expense"'))

# ============================================================
# DATA  (raw transactions)
# ============================================================
data = wb.create_sheet("Data")
data.sheet_properties.tabColor = GREEN
data["A1"] = "DATA MENTAH — satu baris = satu transaksi. Report / Summary / Spending / Assets membaca sheet ini."
data["A1"].font = Font(name="Calibri", color=GREEN, bold=True, size=12)
data.merge_cells("A1:I1")
data.row_dimensions[1].height = 22

data_headers = [
    "Date", "Bank", "Category", "Description", "Amount", "Type",
    "Year", "MonthNum", "MonthName",
]
for i, h in enumerate(data_headers, 1):
    data.cell(row=2, column=i, value=h)
style_header_row(data, 2, 1, 9, PatternFill("solid", fgColor=GREEN))

rows = []
# Income: gaji tgl 1 + bonus Juni & Desember
for m in range(1, 13):
    for bank in BANKS:
        rows.append((date(2025, m, 1), bank, "Salary", f"Gaji {MONTHS[m-1]}", SALARY[bank], "Income"))
    if m in (6, 12):
        rows.append((date(2025, m, 20), "Bank 1", "Bonus", f"Bonus {MONTHS[m-1]}", 5_000_000, "Income"))
# Expenses: 1 baris per kategori per bank per bulan (nilai 0 dilewati, kecuali kita tetap tulis gifts 0? skip 0)
for m in range(1, 13):
    for bi, bank in enumerate(BANKS):
        for cat in EXPENSE_CATS:
            amt = expense_amount(m, bi, cat)
            if amt <= 0:
                continue
            day = 3 + (EXPENSE_CATS.index(cat) % 25)
            rows.append((
                date(2025, m, min(day, 28)),
                bank,
                cat,
                f"{cat} {MONTHS[m-1]}",
                amt,
                "Expense",
            ))

rows.sort(key=lambda r: (r[0], r[1], r[2]))

for i, (dt, bank, cat, desc, amt, typ) in enumerate(rows):
    r = 3 + i
    data.cell(row=r, column=1, value=dt).number_format = "DD-MMM-YYYY"
    data.cell(row=r, column=2, value=bank)
    data.cell(row=r, column=3, value=cat)
    data.cell(row=r, column=4, value=desc)
    data.cell(row=r, column=5, value=amt).number_format = '#,##0'
    data.cell(row=r, column=6, value=typ)
    # helper — formula, bukan ketikan, biar filter Year/Month selalu ikut tanggal
    data.cell(row=r, column=7, value=f"=YEAR(A{r})")
    data.cell(row=r, column=8, value=f"=MONTH(A{r})")
    data.cell(row=r, column=9, value=f'=INDEX(MonthList,H{r})')
    for c in range(1, 10):
        cell = data.cell(row=r, column=c)
        cell.border = border
        cell.alignment = Alignment(horizontal="center" if c != 4 else "left", vertical="center")
        if c >= 7:
            cell.font = font_blue
        else:
            cell.font = font_black
        if i % 2 == 1:
            cell.fill = fill_alt
        if c == 6:
            if typ == "Income":
                cell.font = Font(name="Calibri", color=GREEN, bold=True, size=11)
            else:
                cell.font = Font(name="Calibri", color="B91C1C", size=11)

last_data = 2 + len(rows)
# spare blank rows so user can keep typing with dropdown + helper formulas
for r in range(last_data + 1, last_data + 31):
    for c in range(1, 10):
        cell = data.cell(row=r, column=c)
        cell.border = border
        cell.fill = fill_input if c <= 6 else fill_calc
    data.cell(row=r, column=7, value=f'=IF(A{r}="","",YEAR(A{r}))')
    data.cell(row=r, column=8, value=f'=IF(A{r}="","",MONTH(A{r}))')
    data.cell(row=r, column=9, value=f'=IF(H{r}="","",INDEX(MonthList,H{r}))')
    for c in (7, 8, 9):
        data.cell(row=r, column=c).font = font_blue
    data.cell(row=r, column=5).number_format = '#,##0'
    data.cell(row=r, column=1).number_format = "DD-MMM-YYYY"

add_list_dv(data, "=BankList", f"B3:B{last_data+30}")
add_list_dv(data, "=CatList", f"C3:C{last_data+30}")
add_list_dv(data, "=TypeList", f"F3:F{last_data+30}")

set_widths(data, {
    "A": 16, "B": 14, "C": 18, "D": 28, "E": 16, "F": 12, "G": 10, "H": 12, "I": 14,
})
data.freeze_panes = "A3"
data.auto_filter.ref = f"A2:I{last_data+30}"
data.auto_filter.add_sort_condition("C2")
data.sheet_view.showGridLines = False
data.page_setup.orientation = "landscape"
data.page_setup.fitToPage = True
data.page_setup.fitToWidth = 1
data.page_setup.fitToHeight = 0
data.sheet_view.zoomScale = 110
data.oddHeader.left.text = "Data Mentah"
data.oddFooter.center.text = "Input kuning = ketik manual. Kolom Year/MonthNum/MonthName = rumus, jangan dihapus."

# ============================================================
# BUDGETING
# ============================================================
bud = wb.create_sheet("Budgeting")
bud.sheet_properties.tabColor = "C65911"
bud.merge_cells("A1:F1")
bud["A1"] = "BUDGETING — target anggaran (INPUT). Report kolom Allocation membaca sheet ini."
bud["A1"].font = Font(name="Calibri", color="FFFFFF", bold=True, size=16)
bud["A1"].fill = fill_gold
bud["A1"].alignment = Alignment(horizontal="left", vertical="center")
bud.row_dimensions[1].height = 28

bud.merge_cells("A2:F2")
bud["A2"] = (
    "Cara pakai: ubah angka di kolom Allocation (kuning). "
    "Satu baris = target 1 kategori untuk 1 bulan. "
    "Jangan ubah Year / Month / Category kalau tidak perlu — itu kunci yang dicari Report."
)
bud["A2"].font = font_header_note
bud["A2"].fill = fill_note
bud["A2"].alignment = left_c
bud.row_dimensions[2].height = 32

bud_headers = ["Year", "MonthNum", "Month", "Category", "Allocation (Rp)", "Catatan"]
for i, h in enumerate(bud_headers, 1):
    bud.cell(row=4, column=i, value=h)
style_header_row(bud, 4, 1, 6, fill_gold)

notes = {
    "Home Rent": "Sewa tetap tiap bulan",
    "Utilities": "Listrik air internet",
    "Food": "Makan harian + belanja dapur",
    "Supplies": "Kebutuhan rumah / ATK",
    "Transportation": "Bensin dan transport",
    "Healthcare": "Obat dan check-up",
    "Debt": "Cicilan tetap",
    "Shopping": "Belanja non-rutin",
    "Gifts": "Hadiah / donasi",
}
br = 5
for year in (2025,):
    for mi, month in enumerate(MONTHS, 1):
        for cat in EXPENSE_CATS:
            bud.cell(row=br, column=1, value=year)
            bud.cell(row=br, column=2, value=mi)
            bud.cell(row=br, column=3, value=month)
            bud.cell(row=br, column=4, value=cat)
            cell_alloc = bud.cell(row=br, column=5, value=ALLOC[cat])
            cell_alloc.number_format = '"Rp "#,##0'
            cell_alloc.fill = fill_input
            cell_alloc.font = font_input
            bud.cell(row=br, column=6, value=notes[cat])
            for c in range(1, 7):
                cell = bud.cell(row=br, column=c)
                cell.border = border
                cell.alignment = center if c != 6 else left_c
                if c != 5:
                    cell.font = font_black
                if (br % 2 == 0) and c != 5:
                    cell.fill = fill_alt
            br += 1
last_bud = br - 1

add_list_dv(bud, "=YearList", f"A5:A{last_bud+20}")
add_list_dv(bud, "=MonthList", f"C5:C{last_bud+20}")
add_list_dv(bud, "=CatList", f"D5:D{last_bud+20}")

# blank input rows
for r in range(last_bud + 1, last_bud + 13):
    for c in range(1, 7):
        cell = bud.cell(row=r, column=c)
        cell.border = border
        cell.fill = fill_input
    bud.cell(row=r, column=5).number_format = '"Rp "#,##0'
    bud.cell(row=r, column=2, value=f'=IF(C{r}="","",MATCH(C{r},MonthList,0))')
    bud.cell(row=r, column=2).font = font_blue

set_widths(bud, {"A": 12, "B": 12, "C": 16, "D": 20, "E": 22, "F": 36})
bud.freeze_panes = "A5"
bud.auto_filter.ref = f"A4:F{last_bud}"
bud.sheet_view.showGridLines = False
bud.auto_filter.add_sort_condition("D4")

bud["H4"] = "BACA CEPAT"
bud["H4"].font = font_white_b
bud["H4"].fill = fill_gold
bud["H4"].alignment = center
bud.merge_cells("H4:I4")
bud["H5"] = "Total anggaran / bulan (semua kategori)"
bud["I5"] = f"=SUMIF(A5:A{last_bud},Report!B3,E5:E{last_bud})/12"
bud["I5"].number_format = '"Rp "#,##0'
bud["H6"] = "Artinya"
bud["I6"] = "Jumlah target yang Report pakai sebagai Allocation bulan terpilih (per kategori, bukan total ini)."
bud["I6"].alignment = left_c
bud.merge_cells("I6:I8")
for r in range(5, 9):
    bud.cell(row=r, column=8).font = font_label
    bud.cell(row=r, column=8).alignment = left_c
bud["I5"].font = font_blue
bud["I5"].fill = fill_calc
set_widths(bud, {"H": 42, "I": 28})
bud.row_dimensions[6].height = 20

# ============================================================
# SUMMARY
# ============================================================
sm = wb.create_sheet("Summary")
sm.sheet_properties.tabColor = TEAL
sm.merge_cells("A1:G1")
sm["A1"] = "SUMMARY — ringkasan 12 bulan. Semua angka rumus dari sheet Data. Bukan input."
sm["A1"].font = font_title
sm["A1"].fill = PatternFill("solid", fgColor=TEAL)
sm["A1"].alignment = Alignment(horizontal="left", vertical="center")
sm.row_dimensions[1].height = 32

sm["A2"] = "Year ikut Report"
sm["B2"] = "=Report!B3"
sm["D2"] = "Bank ikut Report"
sm["E2"] = "=Report!B5"
for addr in ("B2", "E2"):
    sm[addr].font = font_input
    sm[addr].fill = fill_calc
    sm[addr].alignment = center
    sm[addr].border = border
for addr in ("A2", "D2"):
    sm[addr].font = font_label
sm.merge_cells("A3:G3")
sm["A3"] = "Ganti Year / Bank di sheet Report. Tabel ini ikut berubah. Income = Type Income. Expense = Type Expense. Savings = Income − Expense."
sm["A3"].font = font_small
sm.row_dimensions[3].height = 20

sum_headers = ["Month", "MonthNum", "Income (Rp)", "Expense (Rp)", "Savings (Rp)", "Savings %", "Status"]
for i, h in enumerate(sum_headers, 1):
    sm.cell(row=5, column=i, value=h)
style_header_row(sm, 5, 1, 7, PatternFill("solid", fgColor=TEAL))

for i, month in enumerate(MONTHS, 1):
    r = 5 + i
    sm.cell(row=r, column=1, value=month)
    sm.cell(row=r, column=2, value=i)
    sm.cell(row=r, column=3, value=(
        f'=SUMIFS(Data!E:E,Data!F:F,"Income",Data!B:B,$E$2,Data!G:G,$B$2,Data!H:H,B{r})'
    ))
    sm.cell(row=r, column=4, value=(
        f'=SUMIFS(Data!E:E,Data!F:F,"Expense",Data!B:B,$E$2,Data!G:G,$B$2,Data!H:H,B{r})'
    ))
    sm.cell(row=r, column=5, value=f"=C{r}-D{r}")
    sm.cell(row=r, column=6, value=f'=IF(C{r}=0,0,E{r}/C{r})')
    sm.cell(row=r, column=7, value=f'=IF(E{r}<0,"Defisit","Aman")')
    for c in range(1, 8):
        cell = sm.cell(row=r, column=c)
        cell.border = border
        cell.font = font_blue
        cell.alignment = center
        if i % 2 == 0:
            cell.fill = fill_alt
    for c in (3, 4, 5):
        sm.cell(row=r, column=c).number_format = '"Rp "#,##0'
    sm.cell(row=r, column=6).number_format = "0.0%"

# totals
tr = 18
sm.cell(row=tr, column=1, value="TOTAL TAHUN")
sm.cell(row=tr, column=3, value="=SUM(C6:C17)")
sm.cell(row=tr, column=4, value="=SUM(D6:D17)")
sm.cell(row=tr, column=5, value="=SUM(E6:E17)")
sm.cell(row=tr, column=6, value="=IF(C18=0,0,E18/C18)")
for c in range(1, 8):
    cell = sm.cell(row=tr, column=c)
    cell.fill = PatternFill("solid", fgColor=TEAL)
    cell.font = font_white_b
    cell.border = border
    cell.alignment = center
for c in (3, 4, 5):
    sm.cell(row=tr, column=c).number_format = '"Rp "#,##0'
sm.cell(row=tr, column=6).number_format = "0.0%"

sm.conditional_formatting.add("G6:G17", CellIsRule(
    operator="equal", formula=['"Defisit"'],
    fill=PatternFill("solid", fgColor="FECACA"),
    font=Font(name="Calibri", color="991B1B", bold=True),
))
sm.conditional_formatting.add("E6:E17", CellIsRule(
    operator="lessThan", formula=["0"],
    font=Font(name="Calibri", color="991B1B", bold=True),
))

# mini column chart
chart_sum = BarChart()
chart_sum.type = "col"
chart_sum.grouping = "clustered"
chart_sum.title = "Income vs Expense per Bulan"
chart_sum.y_axis.title = "Rp"
chart_sum.x_axis.title = None
data_ref = Reference(sm, min_col=3, max_col=4, min_row=5, max_row=17)
cats_ref = Reference(sm, min_col=1, min_row=6, max_row=17)
chart_sum.add_data(data_ref, titles_from_data=True)
chart_sum.set_categories(cats_ref)
chart_sum.shape = 4
chart_sum.legend.position = "b"
chart_sum.y_axis.numFmt = '#,##0'
chart_sum.width = 18
chart_sum.height = 8
sm.add_chart(chart_sum, "A20")

set_widths(sm, {"A": 16, "B": 12, "C": 20, "D": 20, "E": 20, "F": 14, "G": 14})
sm.freeze_panes = "A6"
sm.sheet_view.showGridLines = False
sm.row_dimensions[2].height = 22

# ============================================================
# SPENDING
# ============================================================
sp = wb.create_sheet("Spending")
sp.sheet_properties.tabColor = "B91C1C"
sp.merge_cells("A1:H1")
sp["A1"] = "SPENDING — rincian pengeluaran vs anggaran bulan yang dipilih di Report."
sp["A1"].font = font_title
sp["A1"].fill = PatternFill("solid", fgColor="B91C1C")
sp["A1"].alignment = Alignment(horizontal="left", vertical="center")
sp.row_dimensions[1].height = 32

sp["A2"] = "Year"
sp["B2"] = "=Report!B3"
sp["C2"] = "Month"
sp["D2"] = "=Report!E3"
sp["E2"] = "Bank"
sp["F2"] = "=Report!B5"
sp["G2"] = "MonthNum"
sp["H2"] = "=MATCH(D2,MonthList,0)"
for addr in ("B2", "D2", "F2", "H2"):
    sp[addr].font = font_blue
    sp[addr].fill = fill_calc
    sp[addr].alignment = center
    sp[addr].border = border
for addr in ("A2", "C2", "E2", "G2"):
    sp[addr].font = font_label
    sp[addr].alignment = center

sp.merge_cells("A3:H3")
sp["A3"] = (
    "This Month = uang keluar bulan+bank itu. YTD = uang keluar Januari sampai bulan itu di tahun yang sama. "
    "Budget = Allocation di sheet Budgeting. Sisa = Budget − This Month. Negatif = boros."
)
sp["A3"].font = font_small
sp.row_dimensions[3].height = 32

sp_headers = [
    "Category", "Budget Bulan Ini", "Spending Bulan Ini", "Sisa",
    "% Terpakai", "Bar", "YTD Spending", "Share YTD",
]
for i, h in enumerate(sp_headers, 1):
    sp.cell(row=5, column=i, value=h)
style_header_row(sp, 5, 1, 8, PatternFill("solid", fgColor="B91C1C"))

for i, cat in enumerate(EXPENSE_CATS):
    r = 6 + i
    sp.cell(row=r, column=1, value=cat).font = Font(name="Calibri", bold=True, size=11)
    sp.cell(row=r, column=2, value=(
        f'=SUMIFS(Budgeting!E:E,Budgeting!A:A,$B$2,Budgeting!B:B,$H$2,Budgeting!D:D,A{r})'
    ))
    sp.cell(row=r, column=3, value=(
        f'=SUMIFS(Data!E:E,Data!C:C,A{r},Data!F:F,"Expense",Data!B:B,$F$2,Data!G:G,$B$2,Data!H:H,$H$2)'
    ))
    sp.cell(row=r, column=4, value=f"=B{r}-C{r}")
    sp.cell(row=r, column=5, value=f'=IF(B{r}=0,0,C{r}/B{r})')
    sp.cell(row=r, column=6, value=(
        f'=REPT("█",MIN(20,ROUND(IF(E{r}="",0,E{r})*20,0)))&REPT("░",MAX(0,20-MIN(20,ROUND(IF(E{r}="",0,E{r})*20,0))))'
    ))
    sp.cell(row=r, column=7, value=(
        f'=SUMIFS(Data!E:E,Data!C:C,A{r},Data!F:F,"Expense",Data!B:B,$F$2,Data!G:G,$B$2,Data!H:H,"<="&$H$2)'
    ))
    sp.cell(row=r, column=8, value=f'=IF(SUM($G$6:$G$14)=0,0,G{r}/SUM($G$6:$G$14))')
    for c in range(1, 9):
        cell = sp.cell(row=r, column=c)
        cell.border = border
        cell.alignment = center
        if c != 1:
            cell.font = font_blue
        if i % 2 == 1 and c != 6:
            cell.fill = fill_alt
    sp.cell(row=r, column=6).font = font_bar
    sp.cell(row=r, column=6).alignment = Alignment(horizontal="left", vertical="center")
    for c in (2, 3, 4, 7):
        sp.cell(row=r, column=c).number_format = '"Rp "#,##0'
    sp.cell(row=r, column=5).number_format = "0.0%"
    sp.cell(row=r, column=8).number_format = "0.0%"

sp.conditional_formatting.add("E6:E14", DataBarRule(
    start_type="num", start_value=0, end_type="num", end_value=1,
    color="DC2626", showValue=True, minLength=0, maxLength=100,
))
sp.conditional_formatting.add("D6:D14", CellIsRule(
    operator="lessThan", formula=["0"],
    fill=PatternFill("solid", fgColor="FECACA"),
    font=Font(name="Calibri", color="991B1B", bold=True, size=11),
))
sp.conditional_formatting.add("F6:F14", FormulaRule(
    formula=["$E6>1"],
    font=Font(name="Consolas", color="991B1B", bold=True, size=10),
))

sp.cell(row=15, column=1, value="TOTAL")
sp.cell(row=15, column=2, value="=SUM(B6:B14)")
sp.cell(row=15, column=3, value="=SUM(C6:C14)")
sp.cell(row=15, column=4, value="=SUM(D6:D14)")
sp.cell(row=15, column=5, value="=IF(B15=0,0,C15/B15)")
sp.cell(row=15, column=7, value="=SUM(G6:G14)")
sp.cell(row=15, column=8, value="=SUM(H6:H14)")
for c in range(1, 9):
    cell = sp.cell(row=15, column=c)
    cell.fill = PatternFill("solid", fgColor="7F1D1D")
    cell.font = font_white_b
    cell.border = border
    cell.alignment = center
for c in (2, 3, 4, 7):
    sp.cell(row=15, column=c).number_format = '"Rp "#,##0'
sp.cell(row=15, column=5).number_format = "0.0%"
sp.cell(row=15, column=8).number_format = "0.0%"

# bar chart budget vs spending
bar_sp = BarChart()
bar_sp.type = "bar"  # horizontal
bar_sp.grouping = "clustered"
bar_sp.title = "Budget vs Spending Bulan Ini"
bar_sp.y_axis.title = None
bar_sp.x_axis.title = "Rp"
ref = Reference(sp, min_col=2, max_col=3, min_row=5, max_row=14)
cats = Reference(sp, min_col=1, min_row=6, max_row=14)
bar_sp.add_data(ref, titles_from_data=True)
bar_sp.set_categories(cats)
bar_sp.shape = 4
bar_sp.legend.position = "b"
bar_sp.x_axis.numFmt = '#,##0'
bar_sp.width = 18
bar_sp.height = 8
sp.add_chart(bar_sp, "A17")

set_widths(sp, {
    "A": 20, "B": 22, "C": 22, "D": 18, "E": 14, "F": 28, "G": 20, "H": 14,
})
sp.freeze_panes = "A6"
sp.sheet_view.showGridLines = False

# ============================================================
# ASSETS
# ============================================================
assets = wb.create_sheet("Assets")
assets.sheet_properties.tabColor = PURPLE
assets.merge_cells("A1:H1")
assets["A1"] = "ASSETS — saldo. Opening = input. Income & Expense = rumus dari Data. Saldo akhir = Opening + Income − Expense."
assets["A1"].font = font_title
assets["A1"].fill = PatternFill("solid", fgColor=PURPLE)
assets["A1"].alignment = Alignment(horizontal="left", vertical="center")
assets.row_dimensions[1].height = 32

assets["A2"] = "Year"
assets["B2"] = "=Report!B3"
assets["B2"].font = font_blue
assets["B2"].fill = fill_calc
assets["B2"].border = border
assets["B2"].alignment = center
assets["A2"].font = font_label
assets.merge_cells("C2:H2")
assets["C2"] = "Ubah Opening (kuning) kalau saldo awal tahun beda. Jangan ketik di kolom biru."
assets["C2"].font = font_small

as_headers = [
    "Asset", "Tipe", "Terhubung ke Bank", "Opening (input)",
    "Income tahun ini", "Expense tahun ini", "Saldo akhir", "Catatan",
]
for i, h in enumerate(as_headers, 1):
    assets.cell(row=4, column=i, value=h)
style_header_row(assets, 4, 1, 8, PatternFill("solid", fgColor=PURPLE))

asset_rows = [
    ("Bank 1", "Rekening", "Bank 1", OPENING["Bank 1"], "Rekening gaji utama"),
    ("Bank 2", "Rekening", "Bank 2", OPENING["Bank 2"], "Rekening operasional"),
    ("Bank 3", "Rekening", "Bank 3", OPENING["Bank 3"], "Rekening cadangan"),
    ("Cash on Hand", "Tunai", "", 1_500_000, "Tidak masuk sheet Data. Ubah manual bila perlu."),
    ("Emergency Fund", "Tabungan", "", 15_000_000, "Dana darurat. Tidak gerak otomatis."),
    ("Investasi Reksa", "Investasi", "", 10_000_000, "Nilai manual, bukan dari transaksi harian."),
]

for i, (name, tipe, bank, opening, note) in enumerate(asset_rows):
    r = 5 + i
    assets.cell(row=r, column=1, value=name).font = Font(name="Calibri", bold=True, size=11)
    assets.cell(row=r, column=2, value=tipe)
    assets.cell(row=r, column=3, value=bank if bank else "—")
    cell_o = assets.cell(row=r, column=4, value=opening)
    cell_o.number_format = '"Rp "#,##0'
    cell_o.fill = fill_input
    cell_o.font = font_input
    if bank:
        assets.cell(row=r, column=5, value=(
            f'=SUMIFS(Data!E:E,Data!F:F,"Income",Data!B:B,C{r},Data!G:G,$B$2)'
        ))
        assets.cell(row=r, column=6, value=(
            f'=SUMIFS(Data!E:E,Data!F:F,"Expense",Data!B:B,C{r},Data!G:G,$B$2)'
        ))
    else:
        assets.cell(row=r, column=5, value=0)
        assets.cell(row=r, column=6, value=0)
        assets.cell(row=r, column=5).fill = fill_input
        assets.cell(row=r, column=6).fill = fill_input
        assets.cell(row=r, column=5).font = font_input
        assets.cell(row=r, column=6).font = font_input
    assets.cell(row=r, column=7, value=f"=D{r}+E{r}-F{r}")
    assets.cell(row=r, column=8, value=note)
    for c in range(1, 9):
        cell = assets.cell(row=r, column=c)
        cell.border = border
        cell.alignment = center if c != 8 else left_c
        if c in (5, 6, 7) and bank:
            cell.font = font_blue
        elif c not in (4, 5, 6):
            cell.font = font_black
        if i % 2 == 1 and c not in (4, 5, 6):
            cell.fill = fill_alt
    for c in (5, 6, 7):
        assets.cell(row=r, column=c).number_format = '"Rp "#,##0'

assets.cell(row=11, column=1, value="TOTAL ASET")
assets.cell(row=11, column=4, value="=SUM(D5:D10)")
assets.cell(row=11, column=5, value="=SUM(E5:E10)")
assets.cell(row=11, column=6, value="=SUM(F5:F10)")
assets.cell(row=11, column=7, value="=SUM(G5:G10)")
for c in range(1, 9):
    cell = assets.cell(row=11, column=c)
    cell.fill = PatternFill("solid", fgColor=PURPLE)
    cell.font = font_white_b
    cell.border = border
    cell.alignment = center
for c in (4, 5, 6, 7):
    assets.cell(row=11, column=c).number_format = '"Rp "#,##0'

assets.merge_cells("A13:H13")
assets["A13"] = (
    "Arti value: Opening = saldo 1 Januari sebelum transaksi tahun ini. "
    "Income tahun ini = semua Type Income di Data untuk bank itu. "
    "Expense tahun ini = semua Type Expense. "
    "Saldo akhir = uang yang seharusnya masih ada di rekening pada akhir tahun (atau sampai transaksi terakhir yang diinput)."
)
assets["A13"].font = font_small
assets["A13"].alignment = left_c
assets.row_dimensions[13].height = 36

# asset composition pie uses saldo akhir
pie_as = PieChart()
pie_as.title = "Komposisi Saldo Akhir"
labels = Reference(assets, min_col=1, min_row=5, max_row=10)
vals = Reference(assets, min_col=7, min_row=4, max_row=10)
pie_as.add_data(vals, titles_from_data=True)
pie_as.set_categories(labels)
pie_as.dataLabels = DataLabelList()
pie_as.dataLabels.showPercent = True
pie_as.dataLabels.showVal = False
pie_as.dataLabels.showCatName = False
pie_as.width = 15
pie_as.height = 8
assets.add_chart(pie_as, "A15")

set_widths(assets, {
    "A": 20, "B": 14, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 55,
})
assets.freeze_panes = "A5"
assets.sheet_view.showGridLines = False

# ============================================================
# REPORT
# ============================================================
rp = wb.create_sheet("Report", 0)
rp.sheet_properties.tabColor = "E11D48"
rp.sheet_view.showGridLines = False
rp.sheet_view.zoomScale = 110
rp.page_setup.orientation = "landscape"
rp.page_setup.fitToPage = True
rp.page_setup.fitToWidth = 1
rp.page_setup.fitToHeight = 1
rp.page_setup.paperSize = rp.PAPERSIZE_A4
rp.sheet_view.view = "pageBreakPreview"
rp.sheet_view.view = "normal"
rp.print_title_rows = "1:2"
rp.page_setup.horizontalCentered = True
rp.sheet_properties.pageSetUpPr.fitToPage = True

# banner
rp.merge_cells("A1:L1")
rp["A1"] = "Money Management Dashboard 💸"
rp["A1"].font = font_title
rp["A1"].fill = fill_title_bar
rp["A1"].alignment = Alignment(horizontal="left", vertical="center")
rp.row_dimensions[1].height = 36

rp.merge_cells("A2:L2")
rp["A2"] = "Kuning = filter / input.  Biru = rumus, jangan dioverwrite.  Bar merah = pemakaian anggaran."
rp["A2"].font = Font(name="Calibri", color="E2E8F0", italic=True, size=10)
rp["A2"].fill = PatternFill("solid", fgColor="1E3A5F")
rp["A2"].alignment = Alignment(horizontal="left", vertical="center")
rp.row_dimensions[2].height = 18

# filters
rp["A4"] = "YEAR"
rp["B4"] = 2025
rp["D4"] = "MONTH"
rp["E4"] = "December"
rp["G4"] = "BANK"
rp["H4"] = "Bank 1"
for addr in ("A4", "D4", "G4"):
    rp[addr].font = font_label
    rp[addr].alignment = Alignment(horizontal="right", vertical="center")
for addr in ("B4", "E4", "H4"):
    rp[addr].font = Font(name="Calibri", color="9A5700", bold=True, size=14)
    rp[addr].fill = fill_input
    rp[addr].alignment = center
    rp[addr].border = Border(
        left=Side(style="medium", color="F59E0B"),
        right=Side(style="medium", color="F59E0B"),
        top=Side(style="medium", color="F59E0B"),
        bottom=Side(style="medium", color="F59E0B"),
    )
rp.row_dimensions[4].height = 26

add_list_dv(rp, "=YearList", "B4")
add_list_dv(rp, "=MonthList", "E4")
add_list_dv(rp, "=BankList", "H4")

rp["J4"] = "MonthNum"
rp["K4"] = "=MATCH(E4,MonthList,0)"
rp["J4"].font = font_label
rp["K4"].font = font_blue
rp["K4"].fill = fill_calc
rp["K4"].alignment = center
rp["K4"].border = border

rp.merge_cells("A5:L5")
rp["A5"] = 'Monthly Report 📊    Filter Year + Month + Bank menggerakkan KPI, tabel, bar, dan chart di sheet ini. Summary, Spending, Assets ikut filter yang sama.'
rp["A5"].font = font_section
rp.row_dimensions[5].height = 22

# KPI cards occupying two columns each
kpis = [
    (1, "TOTAL INCOME", f'=SUMIFS(Data!E:E,Data!F:F,"Income",Data!B:B,$H$4,Data!G:G,$B$4,Data!H:H,$K$4)', "E8F5E9", "166534"),
    (3, "SAVINGS", "=B8-F8", "E0F2FE", "075985"),
    (5, "SPENDING", f'=SUMIFS(Data!E:E,Data!F:F,"Expense",Data!B:B,$H$4,Data!G:G,$B$4,Data!H:H,$K$4)', "FEF2F2", "B91C1C"),
    (7, "SALDO BANK (YTD)", f'=SUMIF(Assets!A:A,$H$4,Assets!G:G)', "F5F3FF", "5B2C6F"),
]
for col, label, formula, bg, fg in kpis:
    rp.merge_cells(start_row=7, start_column=col, end_row=7, end_column=col + 1)
    lab = rp.cell(row=7, column=col, value=label)
    lab.font = Font(name="Calibri", color="64748B", bold=True, size=9)
    lab.alignment = center
    lab.fill = PatternFill("solid", fgColor=bg)
    rp.cell(row=7, column=col + 1).fill = PatternFill("solid", fgColor=bg)
    rp.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col + 1)
    val = rp.cell(row=8, column=col, value=formula)
    val.font = Font(name="Calibri", color=fg, bold=True, size=16)
    val.number_format = '"Rp "#,##0'
    val.alignment = center
    val.fill = PatternFill("solid", fgColor=bg)
    rp.cell(row=8, column=col + 1).fill = PatternFill("solid", fgColor=bg)
    for rr in (7, 8):
        for cc in (col, col + 1):
            rp.cell(row=rr, column=cc).border = Border(
                left=Side(style="medium", color=fg),
                right=Side(style="medium", color=fg),
                top=Side(style="medium", color=fg),
                bottom=Side(style="medium", color=fg),
            )
rp.row_dimensions[7].height = 18
rp.row_dimensions[8].height = 28

rp.merge_cells("I7:L7")
rp["I7"] = "Saldo Bank (YTD) = saldo akhir di sheet Assets untuk bank filter, bukan cuma bulan ini."
rp["I7"].font = font_small
rp["I7"].alignment = left_c

# Expense table
rp.merge_cells("A10:E10")
rp["A10"] = "EXPENSES LIST"
rp["A10"].font = Font(name="Calibri", color="FFFFFF", bold=True, size=12)
rp["A10"].fill = fill_navy
rp["A10"].alignment = Alignment(horizontal="left", vertical="center")
for col in range(1, 6):
    rp.cell(row=10, column=col).fill = fill_navy
rp.row_dimensions[10].height = 20

exp_headers = ["Expenses List", "Allocation 💰", "Realization 💸", "Budget Usage Progress", "% Usage"]
for i, h in enumerate(exp_headers, 1):
    rp.cell(row=11, column=i, value=h)
style_header_row(rp, 11, 1, 5, PatternFill("solid", fgColor="16365C"))

for i, cat in enumerate(EXPENSE_CATS):
    r = 12 + i
    rp.cell(row=r, column=1, value=cat).font = Font(name="Calibri", bold=True, color="1E293B", size=11)
    rp.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    # Allocation from Budgeting
    rp.cell(row=r, column=2, value=(
        f'=SUMIFS(Budgeting!E:E,Budgeting!A:A,$B$4,Budgeting!B:B,$K$4,Budgeting!D:D,A{r})'
    ))
    # Realization from Data
    rp.cell(row=r, column=3, value=(
        f'=SUMIFS(Data!E:E,Data!C:C,A{r},Data!F:F,"Expense",Data!B:B,$H$4,Data!G:G,$B$4,Data!H:H,$K$4)'
    ))
    # in-cell bar capped at 20 blocks; over-budget still full bar + red via CF
    rp.cell(row=r, column=4, value=(
        f'=REPT("█",MIN(20,ROUND(IF(E{r}="",0,E{r})*20,0)))&REPT("░",MAX(0,20-MIN(20,ROUND(IF(E{r}="",0,E{r})*20,0))))'
    ))
    rp.cell(row=r, column=5, value=f'=IF(B{r}=0,0,C{r}/B{r})')
    for c in range(1, 6):
        cell = rp.cell(row=r, column=c)
        cell.border = border
        cell.alignment = center if c != 1 else Alignment(horizontal="left", vertical="center")
        if i % 2 == 1 and c != 4:
            cell.fill = fill_alt
    rp.cell(row=r, column=2).font = font_blue
    rp.cell(row=r, column=3).font = font_blue
    rp.cell(row=r, column=4).font = font_bar
    rp.cell(row=r, column=4).alignment = Alignment(horizontal="left", vertical="center")
    rp.cell(row=r, column=5).font = font_blue
    rp.cell(row=r, column=2).number_format = '"Rp "#,##0'
    rp.cell(row=r, column=3).number_format = '"Rp "#,##0'
    rp.cell(row=r, column=5).number_format = "0.0%"
    rp.row_dimensions[r].height = 22

# total row
rp.cell(row=21, column=1, value="TOTAL")
rp.cell(row=21, column=2, value="=SUM(B12:B20)")
rp.cell(row=21, column=3, value="=SUM(C12:C20)")
rp.cell(row=21, column=5, value="=IF(B21=0,0,C21/B21)")
rp.cell(row=21, column=4, value='=REPT("█",MIN(20,ROUND(IF(E21="",0,E21)*20,0)))&REPT("░",MAX(0,20-MIN(20,ROUND(IF(E21="",0,E21)*20,0))))')
for c in range(1, 6):
    cell = rp.cell(row=21, column=c)
    cell.fill = fill_navy
    cell.font = font_white_b if c != 4 else Font(name="Consolas", color="FFFFFF", size=10)
    cell.border = border
    cell.alignment = center if c != 4 else Alignment(horizontal="left", vertical="center")
rp.cell(row=21, column=2).number_format = '"Rp "#,##0'
rp.cell(row=21, column=3).number_format = '"Rp "#,##0'
rp.cell(row=21, column=5).number_format = "0.0%"
rp.cell(row=21, column=4).font = Font(name="Consolas", color="FFFFFF", size=10)

# data bar on % usage + red when over 100%
rp.conditional_formatting.add("E12:E20", DataBarRule(
    start_type="num", start_value=0, end_type="num", end_value=1,
    color="1D4ED8", showValue=True, minLength=0, maxLength=100,
))
rp.conditional_formatting.add("E12:E21", CellIsRule(
    operator="greaterThan", formula=["1"],
    fill=PatternFill("solid", fgColor="FECACA"),
    font=Font(name="Calibri", color="991B1B", bold=True, size=11),
))
rp.conditional_formatting.add("D12:D20", FormulaRule(
    formula=["$E12>1"],
    fill=PatternFill("solid", fgColor="FECACA"),
    font=Font(name="Consolas", color="991B1B", bold=True, size=10),
))
rp.conditional_formatting.add("A12:C20", FormulaRule(
    formula=["$E12>1"],
    fill=PatternFill("solid", fgColor="FFF1F2"),
))

# Pie — realization
pie = PieChart()
pie.title = "Expense Realization Perce"
labels = Reference(rp, min_col=1, min_row=12, max_row=20)
pie_data = Reference(rp, min_col=3, min_row=11, max_row=20)
pie.add_data(pie_data, titles_from_data=True)
pie.set_categories(labels)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.dataLabels.showCatName = True
pie.dataLabels.showSerName = False
pie.width = 15
pie.height = 8
rp.add_chart(pie, "G10")

# Horizontal bar Allocation vs Realization
hbar = BarChart()
hbar.type = "bar"
hbar.grouping = "clustered"
hbar.title = "Allocation vs Realization"
hbar.x_axis.title = "Rp"
hdata = Reference(rp, min_col=2, max_col=3, min_row=11, max_row=20)
hcats = Reference(rp, min_col=1, min_row=12, max_row=20)
hbar.add_data(hdata, titles_from_data=True)
hbar.set_categories(hcats)
hbar.shape = 4
hbar.legend.position = "b"
hbar.x_axis.numFmt = '#,##0'
hbar.width = 15
hbar.height = 8
rp.add_chart(hbar, "G26")

# Monthly trend source (all 12 months, selected year+bank) — feeds column chart
rp.merge_cells("A23:B23")
rp["A23"] = "MONTHLY EXPENSE TREND  (bank + year filter, semua bulan)"
rp["A23"].font = Font(name="Calibri", color=NAVY, bold=True, size=11)
rp.row_dimensions[23].height = 18

rp["A24"] = "Month"
rp["B24"] = "Expense (Rp)"
rp["C24"] = "Income (Rp)"
style_header_row(rp, 24, 1, 3, PatternFill("solid", fgColor="16365C"))

for i, month in enumerate(MONTHS, 1):
    r = 24 + i
    rp.cell(row=r, column=1, value=month).font = font_black
    rp.cell(row=r, column=2, value=(
        f'=SUMIFS(Data!E:E,Data!F:F,"Expense",Data!B:B,$H$4,Data!G:G,$B$4,Data!H:H,{i})'
    ))
    rp.cell(row=r, column=3, value=(
        f'=SUMIFS(Data!E:E,Data!F:F,"Income",Data!B:B,$H$4,Data!G:G,$B$4,Data!H:H,{i})'
    ))
    for c in range(1, 4):
        cell = rp.cell(row=r, column=c)
        cell.border = border
        cell.alignment = center
        if c > 1:
            cell.font = font_blue
            cell.number_format = '"Rp "#,##0'
        if i % 2 == 0 and c == 1:
            cell.fill = fill_alt
        if i % 2 == 0 and c > 1:
            cell.fill = fill_alt
    rp.row_dimensions[r].height = 18

col = BarChart()
col.type = "col"
col.grouping = "clustered"
col.title = "Monthly Expense Trend"
col.y_axis.title = "Rp"
col.x_axis.title = "Month"
cdata = Reference(rp, min_col=2, max_col=3, min_row=24, max_row=36)
ccats = Reference(rp, min_col=1, min_row=25, max_row=36)
col.add_data(cdata, titles_from_data=True)
col.set_categories(ccats)
col.shape = 4
col.legend.position = "b"
col.y_axis.numFmt = '#,##0'
col.width = 18
col.height = 8
rp.add_chart(col, "E23")

rp.merge_cells("A39:E39")
rp["A39"] = (
    "Bar di kolom Progress: 1 blok = 5% anggaran. Penuh 20 blok = 100%. "
    "Lebih dari 100% bar tetap penuh dan sel jadi merah (contoh Supplies Desember Bank 1 ≈ 170%). "
    "Data bar biru di kolom % Usage mengukur hal yang sama. "
    "Chart batang horizontal = Allocation vs Realization. Chart kolom = tren Jan–Des."
)
rp["A39"].font = font_small
rp["A39"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
rp.row_dimensions[39].height = 48

set_widths(rp, {
    "A": 22, "B": 18, "C": 18, "D": 30, "E": 14,
    "F": 3, "G": 14, "H": 16, "I": 14, "J": 12, "K": 12, "L": 18,
})
rp.freeze_panes = "A10"
rp.auto_filter.ref = "A11:E20"
rp.print_options.horizontalCentered = True
rp.page_margins.left = 0.4
rp.page_margins.right = 0.4
rp.page_margins.top = 0.5
rp.page_margins.bottom = 0.5
rp.oddFooter.center.text = "Money Management Dashboard  |  filter Year / Month / Bank di baris 4"
rp.oddFooter.center.font = "Calibri"
rp.oddFooter.center.size = 9

# sheet order
order = ["Report", "Budgeting", "Summary", "Spending", "Assets", "Data", "Setup"]
for i, name in enumerate(order):
    wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True

out = r"D:\AGI\PROJECT\COBA AI AGENT\Money_Management_Dashboard.xlsx"
wb.save(out)
print(f"rows_data={len(rows)} last_data={last_data} last_bud={last_bud}")
print("saved", out)
