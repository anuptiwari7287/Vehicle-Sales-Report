# 🚗 Vehicle Sales Dashboard (Excel)

An interactive Excel dashboard that analyses vehicle / collectible-model sales orders by month, product line, customer, salesperson, country and order status. It is built entirely with **PivotTables, PivotCharts and Slicers**, with no macros or external add-ins.

---

## 📌 Project Overview

| | |
|---|---|
| **File** | `Vehicle_sales_Dashboard.xlsx` |
| **Records** | 2,823 order lines (Jan 2003 – Dec 2005) |
| **Total sales** | ≈ $8.29 M |
| **Tools** | Microsoft Excel (Tables, PivotTables, PivotCharts, Slicers, formulas) |

**Goal:** give a sales manager a one-page view of *what sells, where, when, and through whom*, and let them filter every chart at once.

---

## 📂 Workbook Structure

| Sheet | Purpose |
|---|---|
| **Dashboard** | Final report page: 9 charts, 3 slicers and 4 KPI cards, all linked to the pivots |
| **Worksheet** | Raw data in an Excel Table (`Table1`), 2,823 rows × 17 columns |
| **Order Vs Sales** | Pivot: sales and order count by month, plus total sales / territory KPI source |
| **Product line Vs Sales** | Pivot: sales by product line |
| **Top Customers by sales** | Pivot: top 5 customers by sales |
| **Monthly sales** | Pivot: month × product line sales matrix |
| **Top Salespersons** | Pivot: top 10 salespersons by sales |
| **Order sales across countries** | Pivot: order count and sales by country |
| **Order status vs Quantity** | Pivot: orders by status, plus total vs shipped order counts |
| **Orders according to month** | Pivot: order count by month |

---

## 🗃️ Dataset Fields

| Column | Description |
|---|---|
| Order Number | Order identifier |
| Quantity | Units ordered |
| Price | Unit price |
| **Sales** | Calculated column: `Quantity × Price` (`=PRODUCT(Table1[[#This Row],[Quantity]:[Price]])`) |
| Order Date | Date of the order |
| Year / Month | Derived from the date with `TEXT()` (see *Known Issues*) |
| Order Status | Shipped, Cancelled, Disputed, In Process, On Hold, Resolved |
| Product Line | Classic Cars, Vintage Cars, Motorcycles, Trucks and Buses, Planes, Ships, Trains |
| Product Code | Product identifier |
| Customers | Customer name |
| Address Line / City / Country | Customer location (19 countries) |
| Territory | EMEA, North America, Asia Pacific, Japan |
| Salesperson | Sales representative |
| Deal Size | Small / Medium / Large |

---

## 📊 Dashboard Contents

**Charts (9)**

| Chart | Type | Insight |
|---|---|---|
| Sales Vs Orders | Clustered column + line | Monthly sales against order volume |
| Product line Vs Sales | 3-D clustered column | Which product lines earn the most |
| Top Customers by sales | Pie | Share of the top 5 customers |
| Monthly sales across product line | Stacked line | Seasonality for each product line |
| Top Salesperson | Clustered column | Best-performing reps |
| Order sales across countries | Column + line | Orders and sales by country |
| Order status vs Quantity | Pie | Shipped vs cancelled / on hold / disputed |
| Territory split | Pie | Sales share by territory |
| Orders according to month | Line | Monthly order trend |

**Interactivity**

- 3 **Slicers**: *Month*, *Country* and *Product Line*, connected to all the pivots so one click filters the whole dashboard
- **KPI cards**: total sales, total orders, shipped orders and territory sales

---

## 🔍 Key Findings

*(Figures as currently shown in the pivot tables. See the Known Issues note on month-level figures.)*

- **Total sales ≈ $8.29 M** from **2,823 orders**; **2,618 (~92.7 %)** were shipped.
- **Classic Cars** is the top product line at ≈ **$2.97 M (~36 %)**, followed by **Vintage Cars** (≈ $1.64 M). **Trains** is the smallest (≈ $0.20 M).
- **USA** is the largest market (≈ $2.99 M, 1,004 orders), then **Spain** (≈ $1.02 M) and **France** (≈ $0.92 M).
- **EMEA** has the most order lines (1,407), followed by North America (1,074).
- **Euro Shopping Channel** is the top customer (≈ $766 K), followed by **Mini Gifts Distributors Ltd.** (≈ $531 K). Their salespersons, **Diego Freyre** and **Valarie Nelson**, lead the salesperson ranking.
- **Q4 peak:** October and November are the strongest months (≈ $1.05 M and ≈ $1.21 M); December is weak by comparison.

---

## ▶️ How to Use

1. Open `Vehicle_sales_Dashboard.xlsx` in **Excel 2013 or later** (slicers need desktop Excel; Excel Online has limited slicer support).
2. Go to the **Dashboard** sheet.
3. Click values in the **Month**, **Country** or **Product Line** slicers to filter all charts. Hold `Ctrl` to select several, and use the clear-filter icon to reset.
4. To refresh after editing the data: **Data → Refresh All**.

---

## ⚠️ Known Issues / Suggested Fixes

While documenting the workbook I found a few data problems you may want to fix before publishing:

1. **Year and Month formulas are misaligned.** In `Worksheet`, `Year` and `Month` point at the wrong rows (row 2 uses `E5` and `E3` instead of `E2`). About **2,714 of 2,823 rows** show a different year or month from their Order Date. Month-based charts (*Sales Vs Orders*, *Monthly sales*, *Orders according to month*) and any year filter are therefore affected.
   **Fix:** use `=TEXT([@[Order Date]],"yyyy")` and `=TEXT([@[Order Date]],"mmmm")`, then refresh the pivots.
2. **Corrupted text values.** `CaNorth Americada` appears in the Country column (should be `Canada`). The same pattern shows up in a salesperson name (`Eric North Americatividad`, which should be `Eric Natividad`) and a city (`North Americantes`, which should be `Nantes`). This looks like a Find & Replace of "na" with "North America" that hit other words.
   **Fix:** re-correct these values in `Worksheet`.
3. The three slicer defined names show `#N/A` when read outside Excel. This doesn't affect normal use in desktop Excel.

---

## 🚀 Possible Improvements

- Add Year as a slicer and a year-over-year comparison
- Add a Deal Size analysis (Small / Medium / Large)
- Add profit margin analysis if cost data becomes available
- Replace 3-D charts with 2-D versions for easier reading

---

## 👤 Author

*Your Name*: built as a data-analytics portfolio project in Excel. (Add your GitHub / LinkedIn links here.)

*Dataset appears to be the widely used "Sales Data Sample" of vehicle and model orders.*
