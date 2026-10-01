import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.comments import Comment

FONT = 'Arial'
CLEAN_CSV = '/tmp/claude-0/-home-claude/85218b8e-007c-5ba2-bfe2-a5a971ad410a/scratchpad/retail-sales-project/build/clean_data.csv'
OUT = '/mnt/user-data/outputs/Retail_Sales_Inventory_Performance.xlsx'

df = pd.read_csv(CLEAN_CSV, parse_dates=['Order Date', 'Ship Date'])
n = len(df)
last_row = n + 1

HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
HEADER_FONT = Font(name=FONT, bold=True, color='FFFFFF', size=11)
TITLE_FONT = Font(name=FONT, bold=True, size=14, color='1F4E78')
SUBTITLE_FONT = Font(name=FONT, italic=True, size=10, color='595959')
LABEL_FONT = Font(name=FONT, bold=True, size=10)
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
INPUT_FONT = Font(name=FONT, bold=True, color='0000FF')
INPUT_FILL = PatternFill('solid', fgColor='FFFFCC')
NEG_FILL = PatternFill('solid', fgColor='FCE4E4')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()

# =========================================================================
# Sheet 1: Raw Data
# =========================================================================
ws = wb.active
ws.title = 'Raw Data'
cols = list(df.columns)
ws.append(cols)
for c in range(1, len(cols) + 1):
    cell = ws.cell(row=1, column=c)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL

for row in df.itertuples(index=False):
    vals = []
    for v, colname in zip(row, cols):
        vals.append(v.to_pydatetime() if colname in ('Order Date', 'Ship Date') else v)
    ws.append(vals)

col_letter = {name: get_column_letter(i + 1) for i, name in enumerate(cols)}
for c, colname in enumerate(cols, start=1):
    letter = col_letter[colname]
    if colname in ('Order Date', 'Ship Date'):
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c).number_format = 'mm/dd/yyyy'
        ws.column_dimensions[letter].width = 12
    elif colname in ('Sales', 'Profit'):
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c).number_format = '$#,##0.00'
        ws.column_dimensions[letter].width = 12
    elif colname == 'Discount':
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c).number_format = '0.0%'
        ws.column_dimensions[letter].width = 10
    else:
        ws.column_dimensions[letter].width = 14

last_col_letter = get_column_letter(len(cols))
tbl = Table(displayName='RawData', ref=f'A1:{last_col_letter}{last_row}')
tbl.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
ws.add_table(tbl)
ws.freeze_panes = 'A2'

RD = "'Raw Data'!"
SALES_RNG = f"{RD}${col_letter['Sales']}$2:${col_letter['Sales']}${last_row}"
PROFIT_RNG = f"{RD}${col_letter['Profit']}$2:${col_letter['Profit']}${last_row}"
QTY_RNG = f"{RD}${col_letter['Quantity']}$2:${col_letter['Quantity']}${last_row}"
DISC_RNG = f"{RD}${col_letter['Discount']}$2:${col_letter['Discount']}${last_row}"
CAT_RNG = f"{RD}${col_letter['Category']}$2:${col_letter['Category']}${last_row}"
SUBCAT_RNG = f"{RD}${col_letter['Sub-Category']}$2:${col_letter['Sub-Category']}${last_row}"
REGION_RNG = f"{RD}${col_letter['Region']}$2:${col_letter['Region']}${last_row}"
YEAR_RNG = f"{RD}${col_letter['Year']}$2:${col_letter['Year']}${last_row}"

# =========================================================================
# Hidden helper sheet: distinct lists (for data validation + iteration order)
# =========================================================================
subcats_by_profit = (df.groupby('Sub-Category')['Profit'].sum().sort_values().index.tolist())
regions = sorted(df['Region'].dropna().unique().tolist())
categories = sorted(df['Category'].dropna().unique().tolist())

lst = wb.create_sheet('Lists')
lst['A1'] = 'Sub-Category (worst profit first)'
for i, s in enumerate(subcats_by_profit, start=2):
    lst.cell(row=i, column=1, value=s)
lst.sheet_state = 'hidden'

print("Sheet 1 (Raw Data) + Lists done.", n, "rows,", len(subcats_by_profit), "sub-categories")
wb.save(OUT)

# =========================================================================
# Sheet 2: Category Summary
# =========================================================================
ws2 = wb.create_sheet('Category Summary')
ws2['A1'] = 'Retail Sales & Inventory Performance — Category Summary'
ws2['A1'].font = TITLE_FONT
ws2['A2'] = 'Source: Raw Data tab (9,994 cleaned order-line records, 2015-2018). All figures below are formulas (SUMIFS/AVERAGEIFS) that recalculate if Raw Data changes.'
ws2['A2'].font = SUBTITLE_FONT
ws2.merge_cells('A1:G1')
ws2.merge_cells('A2:G2')

# --- Region x Category matrix (Sales) ---
ws2['A4'] = 'Sales by Region x Category'
ws2['A4'].font = LABEL_FONT
r0 = 5
ws2.cell(row=r0, column=1, value='Region').font = HEADER_FONT
ws2.cell(row=r0, column=1).fill = HEADER_FILL
for j, cat in enumerate(categories, start=2):
    c = ws2.cell(row=r0, column=j, value=cat)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
tot_col = len(categories) + 2
ws2.cell(row=r0, column=tot_col, value='Total').font = HEADER_FONT
ws2.cell(row=r0, column=tot_col).fill = HEADER_FILL

for i, reg in enumerate(regions, start=r0 + 1):
    ws2.cell(row=i, column=1, value=reg).font = BOLD
    for j, cat in enumerate(categories, start=2):
        col = get_column_letter(j)
        f = f'=SUMIFS({SALES_RNG},{REGION_RNG},$A{i},{CAT_RNG},{col}${r0})'
        cell = ws2.cell(row=i, column=j, value=f)
        cell.number_format = '$#,##0'
    col = get_column_letter(tot_col)
    ws2.cell(row=i, column=tot_col, value=f'=SUM(B{i}:{get_column_letter(tot_col-1)}{i})').number_format = '$#,##0'
last_reg_row = r0 + len(regions)
ws2.cell(row=last_reg_row + 1, column=1, value='Total').font = BOLD
for j in range(2, tot_col + 1):
    col = get_column_letter(j)
    ws2.cell(row=last_reg_row + 1, column=j,
              value=f'=SUM({col}{r0+1}:{col}{last_reg_row})').number_format = '$#,##0'

# --- Sub-Category ranked table ---
sc_start = last_reg_row + 4
ws2.cell(row=sc_start - 1, column=1, value='Sub-Category Performance (worst profit first)').font = LABEL_FONT
headers = ['Sub-Category', 'Sales', 'Profit', 'Quantity', 'Avg Discount', 'Profit Margin %']
for j, h in enumerate(headers, start=1):
    c = ws2.cell(row=sc_start, column=j, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
for i, sc in enumerate(subcats_by_profit, start=sc_start + 1):
    ws2.cell(row=i, column=1, value=sc)
    ws2.cell(row=i, column=2, value=f'=SUMIFS({SALES_RNG},{SUBCAT_RNG},$A{i})').number_format = '$#,##0'
    ws2.cell(row=i, column=3, value=f'=SUMIFS({PROFIT_RNG},{SUBCAT_RNG},$A{i})').number_format = '$#,##0'
    ws2.cell(row=i, column=4, value=f'=SUMIFS({QTY_RNG},{SUBCAT_RNG},$A{i})').number_format = '#,##0'
    ws2.cell(row=i, column=5, value=f'=AVERAGEIFS({DISC_RNG},{SUBCAT_RNG},$A{i})').number_format = '0.0%'
    ws2.cell(row=i, column=6, value=f'=C{i}/B{i}').number_format = '0.0%'
sc_end = sc_start + len(subcats_by_profit)

neg_rule = CellIsRule(operator='lessThan', formula=['0'], fill=NEG_FILL)
ws2.conditional_formatting.add(f'C{sc_start+1}:C{sc_end}', neg_rule)
ws2.conditional_formatting.add(f'F{sc_start+1}:F{sc_end}', neg_rule)

for col, w in zip('ABCDEF', [16, 13, 13, 12, 13, 15]):
    ws2.column_dimensions[col].width = w

ws2['A' + str(sc_end + 2)] = ('Note: Tables, Bookcases and Supplies show negative margin — driven by average '
                                'discount rates well above the category norm. See Discount Scenario tab.')
ws2['A' + str(sc_end + 2)].font = SUBTITLE_FONT
ws2.merge_cells(f'A{sc_end+2}:F{sc_end+2}')

print("Sheet 2 (Category Summary) done.")
wb.save(OUT)

# =========================================================================
# Sheet 3: ABC Pareto Analysis
# =========================================================================
subcats_desc = list(reversed(subcats_by_profit))
ws3 = wb.create_sheet('ABC Pareto Analysis')
ws3['A1'] = 'ABC Classification — Sub-Category Profit Contribution'
ws3['A1'].font = TITLE_FONT
ws3['A2'] = ('Sub-categories ranked by total profit, highest first. Class A = top contributors up to 80% of '
             'cumulative profit, B = next tier to 95%, C = remainder (includes loss-making sub-categories).')
ws3['A2'].font = SUBTITLE_FONT
ws3.merge_cells('A1:F1')
ws3.merge_cells('A2:F2')

headers = ['Sub-Category', 'Profit', 'Cumulative Profit', 'Cumulative %', 'ABC Class']
r0 = 4
for j, h in enumerate(headers, start=1):
    c = ws3.cell(row=r0, column=j, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL

start = r0 + 1
end = start + len(subcats_desc) - 1
for i, sc in enumerate(subcats_desc, start=start):
    ws3.cell(row=i, column=1, value=sc)
    ws3.cell(row=i, column=2, value=f'=SUMIFS({PROFIT_RNG},{SUBCAT_RNG},$A{i})').number_format = '$#,##0'
    if i == start:
        ws3.cell(row=i, column=3, value=f'=B{i}')
    else:
        ws3.cell(row=i, column=3, value=f'=C{i-1}+B{i}')
    ws3.cell(row=i, column=3).number_format = '$#,##0'
    ws3.cell(row=i, column=4, value=f'=C{i}/$C${end}').number_format = '0.0%'
    ws3.cell(row=i, column=5, value=f'=IF(D{i}<=0.8,"A",IF(D{i}<=0.95,"B","C"))')

neg_rule = CellIsRule(operator='lessThan', formula=['0'], fill=NEG_FILL)
ws3.conditional_formatting.add(f'B{start}:B{end}', neg_rule)

for col, w in zip('ABCDE', [16, 13, 16, 13, 10]):
    ws3.column_dimensions[col].width = w

# Pareto chart: bar (Profit) + line (Cumulative %) on secondary axis
bar = BarChart()
bar.title = 'Profit by Sub-Category with Cumulative % (Pareto)'
bar.y_axis.title = 'Profit ($)'
bar.x_axis.title = 'Sub-Category'
data = Reference(ws3, min_col=2, min_row=r0, max_row=end)
cats = Reference(ws3, min_col=1, min_row=start, max_row=end)
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.height, bar.width = 10, 22

line = LineChart()
line_data = Reference(ws3, min_col=4, min_row=r0, max_row=end)
line.add_data(line_data, titles_from_data=True)
line.y_axis.axId = 200
line.y_axis.title = 'Cumulative %'
line.y_axis.crosses = 'max'
bar.y_axis.crosses = 'autoZero'
line.x_axis.delete = True  # hide the line chart's own category axis so only
# the bar chart's "Sub-Category" axis title shows (combo charts otherwise render two)
bar += line
ws3.add_chart(bar, f'G{r0}')

print("Sheet 3 (ABC Pareto Analysis) done.")
wb.save(OUT)

# =========================================================================
# Sheet 4: Discount Scenario (Goal Seek)
# =========================================================================
ws4 = wb.create_sheet('Discount Scenario')
ws4['A1'] = 'Discount Breakeven Scenario (What-If / Goal Seek)'
ws4['A1'].font = TITLE_FONT
ws4.merge_cells('A1:D1')
ws4['A2'] = ('Model: Cost and pre-discount (list) revenue for the selected sub-category are held constant; '
             'only the discount rate varies. This isolates the pricing lever from volume/cost changes — a '
             'simplification stated here, not hidden in the formulas.')
ws4['A2'].font = SUBTITLE_FONT
ws4.merge_cells('A2:F2')
ws4.row_dimensions[2].height = 28
ws4['A2'].alignment = Alignment(wrap_text=True, vertical='top')

ws4['A4'] = 'Sub-Category'
ws4['A4'].font = LABEL_FONT
ws4['B4'] = subcats_by_profit[0]  # worst performer, e.g. Tables
ws4['B4'].font = INPUT_FONT
ws4['B4'].fill = INPUT_FILL
dv = DataValidation(type='list', formula1=f"=Lists!$A$2:$A${len(subcats_by_profit)+1}", allow_blank=False)
ws4.add_data_validation(dv)
dv.add(ws4['B4'])

rows = [
    ('Actual Sales', f'=SUMIFS({SALES_RNG},{SUBCAT_RNG},$B$4)', '$#,##0'),
    ('Actual Profit', f'=SUMIFS({PROFIT_RNG},{SUBCAT_RNG},$B$4)', '$#,##0'),
    ('Actual Quantity', f'=SUMIFS({QTY_RNG},{SUBCAT_RNG},$B$4)', '#,##0'),
    ('Actual Avg Discount', f'=AVERAGEIFS({DISC_RNG},{SUBCAT_RNG},$B$4)', '0.0%'),
    ('Implied Cost (Sales - Profit)', '=B5-B6', '$#,##0'),
    ('Implied Pre-Discount (List) Revenue', '=B5/(1-B8)', '$#,##0'),
]
r = 5
for label, formula, fmt in rows:
    ws4.cell(row=r, column=1, value=label).font = NORMAL
    cell = ws4.cell(row=r, column=2, value=formula)
    cell.number_format = fmt
    r += 1
# rows now at 5..10 -> B5 Sales, B6 Profit, B7 Qty, B8 AvgDiscount, B9 Cost, B10 ListRevenue

ws4['A12'] = 'SCENARIO — set Discount Rate, or use Goal Seek to zero out Scenario Profit'
ws4['A12'].font = LABEL_FONT
ws4.merge_cells('A12:D12')

default_subcat = subcats_by_profit[0]
default_discount = float(df.loc[df['Sub-Category'] == default_subcat, 'Discount'].mean())

ws4['A13'] = 'Discount Rate (Goal Seek: changing cell)'
ws4['B13'] = round(default_discount, 4)  # hardcoded starting value, NOT a formula —
# Goal Seek's "changing cell" must hold a constant, not a formula, or Excel refuses it
# with "Cell must contain a value". Starts equal to the sub-category's actual avg
# discount; edit it directly (or via Goal Seek) to test other rates.
ws4['B13'].font = INPUT_FONT
ws4['B13'].fill = INPUT_FILL
ws4['B13'].number_format = '0.0%'
ws4['B13'].comment = Comment(
    'Fixed input (not a formula) — Goal Seek requires that for the changing cell. '
    'Starts at the actual avg discount for the selected Sub-Category. '
    'In Excel: Data > What-If Analysis > Goal Seek > Set cell B15 To value 0 '
    'By changing cell B13.', 'Model')

ws4['A14'] = 'Scenario Sales'
ws4['B14'] = '=B10*(1-B13)'
ws4['B14'].number_format = '$#,##0'

ws4['A15'] = 'Scenario Profit (Goal Seek: set cell, target 0)'
ws4['B15'] = '=B14-B9'
ws4['B15'].number_format = '$#,##0'
ws4['B15'].font = BOLD
ws4['B15'].comment = Comment('Goal Seek target cell — set this to 0 by changing B13.', 'Model')

ws4['A16'] = 'Scenario Margin %'
ws4['B16'] = '=B15/B14'
ws4['B16'].number_format = '0.0%'

ws4['A17'] = ('Tip: B13 is a fixed value, not linked to the Sub-Category dropdown above. If you '
              'change B4, copy B8 (Actual Avg Discount) into B13 before running Goal Seek again.')
ws4['A17'].font = SUBTITLE_FONT
ws4.merge_cells('A17:F17')
ws4.row_dimensions[17].height = 26
ws4['A17'].alignment = Alignment(wrap_text=True, vertical='top')

neg_rule = CellIsRule(operator='lessThan', formula=['0'], fill=NEG_FILL)
ws4.conditional_formatting.add('B15:B16', neg_rule)
ws4.conditional_formatting.add('B6', neg_rule)

# --- Sensitivity table: Profit at discount 0% to 60% ---
ws4['A19'] = 'Sensitivity: Scenario Profit at Discount Rate 0%–60%'
ws4['A19'].font = LABEL_FONT
ws4.merge_cells('A19:D19')
ws4.cell(row=20, column=1, value='Discount Rate').font = HEADER_FONT
ws4.cell(row=20, column=1).fill = HEADER_FILL
ws4.cell(row=20, column=2, value='Scenario Profit').font = HEADER_FONT
ws4.cell(row=20, column=2).fill = HEADER_FILL
sens_start = 21
for i, pct in enumerate([x / 100 for x in range(0, 65, 5)], start=sens_start):
    ws4.cell(row=i, column=1, value=pct).number_format = '0%'
    ws4.cell(row=i, column=2, value=f'=$B$10*(1-A{i})-$B$9').number_format = '$#,##0'
sens_end = sens_start + 12
ws4.conditional_formatting.add(f'B{sens_start}:B{sens_end}', neg_rule)

for col, w in zip('ABCDEF', [34, 15, 12, 12, 12, 12]):
    ws4.column_dimensions[col].width = w

line = LineChart()
line.title = 'Scenario Profit vs. Discount Rate (breakeven = where line crosses zero)'
line.y_axis.title = 'Scenario Profit ($)'
line.x_axis.title = 'Discount Rate'
data = Reference(ws4, min_col=2, min_row=20, max_row=sens_end)
cats = Reference(ws4, min_col=1, min_row=sens_start, max_row=sens_end)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
line.height, line.width = 9, 18
ws4.add_chart(line, 'D4')

print("Sheet 4 (Discount Scenario) done.")
wb.save(OUT)

# =========================================================================
# Sheet 5: Dashboard
# =========================================================================
ws5 = wb.create_sheet('Dashboard')
ws5['A1'] = 'Retail Sales & Inventory Performance — Dashboard'
ws5['A1'].font = TITLE_FONT
ws5.merge_cells('A1:H1')
ws5['A2'] = 'Sample Superstore dataset (US, 2015-2018) — 9,994 cleaned order-line records across 3 categories, 17 sub-categories, 4 regions, 49 states.'
ws5['A2'].font = SUBTITLE_FONT
ws5.merge_cells('A2:H2')

# KPI cards
kpis = [
    ('Total Sales', f'=SUM({SALES_RNG})', '$#,##0'),
    ('Total Profit', f'=SUM({PROFIT_RNG})', '$#,##0'),
    ('Overall Margin %', '=C4/A4', '0.0%'),
    ('Total Orders', f"=SUMPRODUCT(1/COUNTIF({RD}${col_letter['Order ID']}$2:${col_letter['Order ID']}${last_row},{RD}${col_letter['Order ID']}$2:${col_letter['Order ID']}${last_row}))", '#,##0'),
]
for i, (label, formula, fmt) in enumerate(kpis):
    col = 1 + i * 2
    lc = get_column_letter(col)
    ws5.cell(row=3, column=col, value=label).font = LABEL_FONT
    c = ws5.cell(row=4, column=col, value=formula)
    c.font = Font(name=FONT, bold=True, size=13, color='1F4E78')
    c.number_format = fmt
    ws5.column_dimensions[lc].width = 18

# Yearly trend
ws5['A7'] = 'Sales & Profit by Year'
ws5['A7'].font = LABEL_FONT
ws5.cell(row=8, column=1, value='Year').font = HEADER_FONT
ws5.cell(row=8, column=1).fill = HEADER_FILL
ws5.cell(row=8, column=2, value='Sales').font = HEADER_FONT
ws5.cell(row=8, column=2).fill = HEADER_FILL
ws5.cell(row=8, column=3, value='Profit').font = HEADER_FONT
ws5.cell(row=8, column=3).fill = HEADER_FILL
years = sorted(df['Year'].unique().tolist())
for i, yr in enumerate(years, start=9):
    ws5.cell(row=i, column=1, value=yr)
    ws5.cell(row=i, column=2, value=f'=SUMIFS({SALES_RNG},{YEAR_RNG},$A{i})').number_format = '$#,##0'
    ws5.cell(row=i, column=3, value=f'=SUMIFS({PROFIT_RNG},{YEAR_RNG},$A{i})').number_format = '$#,##0'
yr_end = 9 + len(years) - 1

# Region totals
reg_start = yr_end + 3
ws5.cell(row=reg_start - 1, column=1, value='Profit by Region').font = LABEL_FONT
ws5.cell(row=reg_start, column=1, value='Region').font = HEADER_FONT
ws5.cell(row=reg_start, column=1).fill = HEADER_FILL
ws5.cell(row=reg_start, column=2, value='Profit').font = HEADER_FONT
ws5.cell(row=reg_start, column=2).fill = HEADER_FILL
for i, reg in enumerate(regions, start=reg_start + 1):
    ws5.cell(row=i, column=1, value=reg)
    ws5.cell(row=i, column=2, value=f'=SUMIFS({PROFIT_RNG},{REGION_RNG},$A{i})').number_format = '$#,##0'
reg_end = reg_start + len(regions)

for col, w in zip('ABC', [12, 14, 14]):
    ws5.column_dimensions[col].width = w

line = LineChart()
line.title = 'Sales Trend by Year'
data = Reference(ws5, min_col=2, min_row=8, max_row=yr_end)
cats = Reference(ws5, min_col=1, min_row=9, max_row=yr_end)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
line.height, line.width = 8, 14
ws5.add_chart(line, 'E7')

bar = BarChart()
bar.title = 'Profit by Region'
data2 = Reference(ws5, min_col=2, min_row=reg_start, max_row=reg_end)
cats2 = Reference(ws5, min_col=1, min_row=reg_start + 1, max_row=reg_end)
bar.add_data(data2, titles_from_data=True)
bar.set_categories(cats2)
bar.height, bar.width = 8, 14
ws5.add_chart(bar, 'E22')

wb.move_sheet('Dashboard', offset=-4)  # put Dashboard first
wb.active = 0

print("Sheet 5 (Dashboard) done.")
wb.save(OUT)
print("ALL SHEETS BUILT. Saved:", OUT)
