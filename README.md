# Retail Sales & Inventory Performance Dashboard

Excel and Power BI analysis of retail sales, profitability, and discounting
using the Sample Superstore dataset (9,994 orders, 2015-2018). Built to
demonstrate formula-driven Excel modelling and a star-schema Power BI
dashboard with DAX measures.

## Business questions

- Which sub-categories drive the most profit, and which are losing money?
- How much of total profit comes from a small number of sub-categories
  (ABC/Pareto analysis)?
- At what discount level does a sub-category stop being profitable?
- How do sales and profit trend over time, and by category/region?

## Data

Source: [Sample Superstore dataset](https://github.com/leonism/sample-superstore)
(`data/clean_data.csv`), cleaned from the raw export by dropping 806 fully
blank rows. 9,994 clean order-line records remain, matching the canonical
Superstore dataset size.

## Repo structure

```
data/
  clean_data.csv          # cleaned source data (9,994 rows)
scripts/
  01_clean_data.py        # raw -> clean_data.csv
  02_build_excel_workbook.py   # builds the Excel workbook below (openpyxl)
excel/
  Retail_Sales_Inventory_Performance.xlsx
powerbi/
  Retail_Sales_Inventory_Performance.pbix   # add this file, see note below
```

## Excel workbook

Built entirely with formulas (SUMIFS/AVERAGEIFS, no hardcoded values), with
native charts and data validation. Sheets:

- **Dashboard** - KPI cards (Total Sales, Total Profit, Overall Margin %,
  Order Count) plus summary charts
- **Raw Data** - the cleaned dataset as an Excel Table
- **Category Summary** - PivotTable-style rollup by Category and Sub-Category
- **ABC Pareto Analysis** - sub-categories ranked by profit with cumulative
  profit % and ABC classification, combo bar/line chart
- **Discount Scenario** - Goal Seek-ready what-if model: finds the discount
  rate at which a chosen sub-category's scenario profit hits zero

## Power BI dashboard

Star schema: `RawData` (fact) + `DimDate` + `DimProduct` (dimensions, built
from `RawData` via Power Query/DAX). Key DAX measures: Total Sales, Total
Profit, Profit Margin %, YoY Sales Growth %, Rolling 3-Month Sales,
Sub-Category Profit Rank, Cumulative Profit %, ABC Class.

Four report pages:

1. **Overview** - KPI cards, sales trend over time, YoY growth
2. **Category & Region** - matrix and clustered bar by category/region
3. **ABC Pareto** - combo chart (profit bars + cumulative % line) and detail
   table with ABC classification, conditional formatting on negative profit
4. **Discount & Margin** - scatter of discount vs. profit by sub-category,
   sized by sales, highlighting over-discounted loss-makers



## Key finding

Tables, Bookcases, and Supplies are the three loss-making sub-categories,
driven by discount rates well above their breakeven point (confirmed via
Excel Goal Seek: ~19.8% breakeven discount vs. actual average discounts in
the 25-30%+ range for these sub-categories).

## Tools

Excel (formulas, PivotTables, Goal Seek, conditional formatting, native
charts), Power BI (Power Query, DAX, star schema data modelling).
