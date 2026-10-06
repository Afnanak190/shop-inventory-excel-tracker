# Shop Inventory & Sales Performance Tracker

**Tools:** Microsoft Excel

## What This Project Does

Built an Excel-based inventory and sales tracker for a retail shop carrying 75 products across 5 categories. The goal was to flag which products need reordering, calculate profit per product, and summarize performance by category — the kind of tracker a small retail business would actually use.

## Dataset

75 products across 5 categories (Electronics, Groceries, Stationery, Home Care, Personal Care), with fields for stock quantity, reorder level, cost price, selling price, and units sold.

## What Was Done

**1. Data Cleaning**
- Used `TRIM` and `PROPER` to fix inconsistent spacing and casing in product names, categories, and supplier names (e.g., "ELECTRONICS" / "electronics" → "Electronics").

**2. Profit Calculation**
```excel
=(Selling_Price - Cost_Price) * Units_Sold_This_Month
```
Calculated profit per product for the month.

**3. Stock Status Flag**
```excel
=IF(Stock_Quantity < Reorder_Level, "Reorder Needed", "Sufficient Stock")
```
Automatically flags products that have dropped below their reorder threshold.

**4. Category Summary**
Used `UNIQUE` to pull the distinct category list automatically, then `SUMIFS` to total profit per category:
```excel
=SUMIFS(Total_Profit, Category, [category_name])
```

**5. Pivot Table & Chart**
Built a Pivot Table (Category × Stock Status) summarizing profit, and a Pivot Chart to visualize which categories carry the most reorder risk.

**6. Conditional Formatting**
Highlighted "Reorder Needed" products in red and "Sufficient Stock" products in green for quick visual scanning.

## Key Findings

- Total profit across all products: **₹3,01,995**
- **Stationery** stands out — a relatively large share of its profit (₹28,735 of ₹56,995) comes from products currently flagged "Reorder Needed," meaning stockouts there would hit profit disproportionately hard.
- **Home Care** products are almost entirely well-stocked, with minimal reorder risk.
- Category-wise profit breakdown:

| Category | Total Profit |
|---|---|
| Groceries | ₹79,060 |
| Home Care | ₹71,460 |
| Stationery | ₹56,995 |
| Personal Care | ₹48,630 |
| Electronics | ₹45,850 |

## Skills Demonstrated

Excel formulas (TRIM, PROPER, IF, SUMIFS, UNIQUE), Pivot Tables, Pivot Charts, Conditional Formatting, data cleaning, business analysis# shop-inventory-excel-tracker
