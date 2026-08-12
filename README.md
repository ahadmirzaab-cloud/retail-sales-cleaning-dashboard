# Retail Sales Analysis — Data Cleaning + Dashboard (Excel)

## Overview
A realistic retail sales analysis project that starts with **messy, real-world-style data** (missing values, duplicate entries, inconsistent formatting, data entry errors) and walks through a documented cleaning process before building a KPI dashboard.

## Problem Statement
Raw transaction data (635 rows) was inconsistent and unreliable: missing quantities/prices, duplicate order entries, inconsistent text casing, mixed date formats, and negative quantity errors. This project cleans the data end-to-end and turns it into a decision-ready dashboard.

## Workbook Structure
1. **Raw Data (Messy)** — original 635-row dataset exactly as received, problem cells highlighted
2. **Data Cleaning Log** — every cleaning step documented (what was wrong, how it was fixed) + a Before vs After comparison table
3. **Clean Data** — final 620-row clean dataset with a formula-based `TotalSale` column
4. **Summary Dashboard** — KPIs, category/region/rep/product revenue breakdown (via `SUMIFS`), 3 charts
5. **Monthly Trend** — 12-month revenue trend line chart

## Data Issues Found & Fixed
| Issue | Rows Affected | Fix Applied |
|---|---|---|
| Missing Quantity | 18 | Filled with category median |
| Missing Unit Price | 12 | Filled with product-level median price |
| Missing Sales Rep | 8 | Labeled "Unassigned" (revenue kept) |
| Duplicate transactions | 15 | Removed (kept first occurrence) |
| Negative quantities | 6 | Corrected to absolute value |
| Inconsistent casing/spacing | 40+ | Standardized (title case + canonical product names) |
| Mixed date formats | 25 | Standardized to YYYY-MM-DD |

## Tools & Techniques
- Data cleaning logic documented step-by-step (reproducible, auditable)
- `SUMIFS`, `SUMPRODUCT` + `TEXT()` for dynamic aggregation — no hardcoded numbers
- Bar, pie, and line charts for visual storytelling
- (See note below on adding a live Excel PivotTable)

## Want a live PivotTable too?
Programmatically-generated PivotTables in `.xlsx` files are unreliable across Excel versions, so this workbook uses `SUMIFS`-based summary tables instead — they update live and are considered equally (often more) professional by clients. If you want to practice the actual PivotTable skill (useful for client calls/demos):
1. Open `Clean Data` sheet → select all data
2. Insert → PivotTable → New Worksheet
3. Drag `Category` or `Region` to Rows, `TotalSale` to Values
Takes 30 seconds and is a good skill to demo live to clients.

## Files
- `Retail_Sales_Analysis_v2.xlsx` — full workbook (5 sheets)
- `raw_messy_sales_data.csv` — original messy source data
- `clean_sales_data.csv` — cleaned output data

---
*Built as a freelance portfolio project — Data Cleaning & Analysis (Excel).*
