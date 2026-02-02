# Retail Supplier Performance & Inventory Optimization Dashboard

**Flagship project:** P&G supplies materials to **Walmart**, **Sam's Club**, and **Costco** across regions. Overstock = cash locked + storage cost; Understock = lost sales + retailer penalties. Pricing & promotions directly affect sell-through.

---

## Business Context

| Trade-off | Impact |
|-----------|--------|
| **Overstock** | Cash locked, storage cost, markdown pressure |
| **Understock** | Lost sales, retailer penalties, scorecard impact |
| **Pricing & promos** | Direct effect on sell-through and margin |

Retailers: **Walmart** (EDLP, broad SKU), **Sam's Club** (bulk, promos), **Costco** (limited SKU, margin stability).

---

## Project Structure

```
Retail_Supplier_Performance_Dashboard/
├── data/                           # Simulated datasets (80k–120k rows)
│   ├── Sales.csv                   # ~105k rows
│   ├── Inventory.csv               # ~12k rows
│   └── Pricing_Promo.csv           # ~527 rows
├── scripts/
│   └── generate_data.py            # Regenerate data
├── PowerQuery/
│   └── Merge_Sales_Inventory_Pricing.pq
├── VBA/
│   └── DashboardAutomation.bas
├── Excel/
│   ├── KPI_Formulas_Reference.txt
│   ├── (Pivot suggestions in KPI_Formulas_Reference.txt)
│   ├── PowerQuery_Setup.txt
│   └── VBA_Button_And_Alerts.txt
├── requirements.txt
└── README.md
```

---

## Data Schema

### Sales Table
| Column | Type |
|--------|------|
| Date | Date |
| SKU | Text |
| Product Category | household, personal care, packaged foods, beverages, baby care, cleaning |
| Retailer | Walmart, Sam's Club, Costco |
| Store | Text |
| Region | Midwest, South, West, Northeast |
| Units Sold | Integer |
| Revenue | Number |
| Cost | Number |
| Gross Margin % | Number |

### Inventory Table
| Column | Type |
|--------|------|
| SKU | Text |
| Retailer | Text |
| Store | Text |
| Region | Text |
| Beginning Inventory | Integer |
| Ending Inventory | Integer |
| Reorder Point | Integer |
| Lead Time (days) | Integer |

### Pricing & Promo Table
| Column | Type |
|--------|------|
| SKU | Text |
| Retailer | Text |
| List Price | Number |
| Promo Price | Number |
| Promo Start | Date |
| Promo End | Date |
| Marketing Spend | Number |

---

## Quick Start

### 1. Generate data (80k–120k rows)
```bash
python scripts/generate_data.py
```
Output: `data/Sales.csv`, `data/Inventory.csv`, `data/Pricing_Promo.csv`.

### 2. Excel setup
1. **Power Query:** Load `Sales`, `Inventory`, `Pricing_Promo` as connections (see `Excel/PowerQuery_Setup.txt`).
2. **Merge:** Create query `Merged` using M code from `PowerQuery/Merge_Sales_Inventory_Pricing.pq`.
3. **VBA:** Import `VBA/DashboardAutomation.bas`; add button → assign `RefreshAndReport` (see `Excel/VBA_Button_And_Alerts.txt`).
4. **KPIs & Pivots:** Use `Excel/KPI_Formulas_Reference.txt`.

### 3. One-click automation
- **RefreshAndReport:** Refresh data, update pivots, highlight stockout risks, run alerts.
- **ExportSummaryReport:** Export Executive Summary/KPIs sheet to PDF.

---

## Excel Skills Demonstrated

| Skill | Use |
|-------|-----|
| **Power Query** | Merge Sales + Inventory + Pricing; clean SKUs, dates, nulls; trim; dedupe |
| **Formulas** | SUMIFS, XLOOKUP, FILTER; gross margin %, inventory turnover, sell-through % |
| **KPIs** | Inventory Turnover, Stockout Rate, Gross Margin %, Promo Lift, Revenue per SKU |
| **VBA** | One-click refresh, pivot refresh, stockout highlighting, conditional formatting, alerts, PDF export |

---

## VBA Alerts

1. **Low inventory risk for top SKUs** – Ending Inventory ≤ Reorder Point (5+ combinations).
2. **Margin erosion due to discounting** – Gross Margin % below 15% target (50+ transactions).
3. **Stockout risk** – Zero ending inventory (3+ SKU/Store).

---


