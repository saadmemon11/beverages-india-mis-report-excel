# 📊 Beverages India — Monthly Category MIS Report
### Microsoft Excel | FMCG Analytics | Beverages Category | India | FY 2023

---

## 📌 Project Overview

This project simulates a **Monthly Category MIS (Management Information System) Report** — the type of client-ready report that FMCG analytics companies deliver to CPG manufacturers every week.

The report tracks **6 beverage SKU categories** across **4 India regions**, **12 months**, and **4 sales channels** — covering revenue performance, market share, KPI tracking, and business insights — all built in Microsoft Excel with live formulas, pivot analysis, and an interactive dashboard.

---

## 🎯 Business Context

> A brand manager at a company like **Coca-Cola, Monster, or Tropicana** would use this MIS report to answer:
> - Are we hitting our monthly revenue targets?
> - Which SKU category is growing fastest and in which region?
> - How is our market share trending vs competitors?
> - Which months are underperforming and why?
> - Are our promotional campaigns actually driving incremental revenue?

This is the exact type of weekly deliverable that FMCG analytics analysts produce for their CPG clients.

---

## 📁 Workbook Structure — 5 Sheets

| Sheet | Name | What's Inside |
|---|---|---|
| 📦 | **Raw Data** | 290 rows of transactional sales data across 6 SKUs, 4 regions, 12 months |
| 📊 | **Pivot Analysis** | Revenue by SKU × Region and Monthly Revenue by SKU pivot-style tables |
| 🎯 | **KPI Tracker** | Monthly Actual vs Target, Attainment %, MoM Growth %, YTD tracking, Market Share |
| 📈 | **Dashboard** | 5 charts, 5 KPI cards, full visual summary of FY 2023 performance |
| 💡 | **Insight Summary** | 5 analyst-written business observations with recommendations |

---

## 📦 Dataset Details

| Attribute | Details |
|---|---|
| Total Rows | 290 |
| Time Period | Jan 2023 – Dec 2023 (12 months) |
| SKU Categories | Cola, Energy Drink, Juice, Water, Sports Drink, Tea |
| Brands | ColaMax, Monster, Tropika, AquaPure, PowerX, TeaLeaf |
| Regions | North India, South India, East India, West India |
| Key Columns | Month, SKU Category, Brand, Region, Units Sold, Unit Price, Gross Revenue, Discount %, Net Revenue, Promotion Flag, Customer Rating |

Dataset is synthetic and generated using Python for portfolio purposes, modeled after real FMCG retail data patterns.

---

## 📈 Dashboard — 5 Charts

| Chart | Type | What It Shows |
|---|---|---|
| Monthly Net Revenue vs Target + MoM Growth % | Combo (Column + Line) | Revenue bars vs target bars with growth % line on secondary axis |
| Market Share by SKU Category | Doughnut | Revenue share split across all 6 SKU categories |
| Net Revenue by Region | Horizontal Bar | Which India region drives the most revenue |
| Monthly Revenue Trend by SKU | Multi-Line | How each SKU category performed month by month across FY 2023 |
| Promotion Impact | Clustered Column | Promo vs Non-Promo revenue comparison for each SKU |

---

## 🔢 Excel Formulas Used

```excel
-- Pull actual monthly revenue from Raw Data
=SUMIFS('Raw Data'!K:K, 'Raw Data'!B:B, A5)

-- Revenue vs Target gap
=C5 - D5

-- Attainment percentage with error handling
=IFERROR(C5 / D5, 0)

-- Month on Month growth rate
=IFERROR((C6 - C5) / C5, 0)

-- Running YTD total (expanding range)
=SUM($C$5:C5)

-- Market share per SKU from Raw Data
=SUMIF('Raw Data'!D:D, A22, 'Raw Data'!K:K)

-- Market share percentage
=IFERROR(C22 / SUM($C$22:$C$27), 0)

-- Brand lookup using modern XLOOKUP
=IFERROR(XLOOKUP(A22, 'Raw Data'!D:D, 'Raw Data'!E:E, "N/A", 0, 1), "N/A")

-- SKU revenue rank
=RANK(C22, $C$22:$C$27, 0)

-- Dynamic status with nested IF
=IF(F5>=1, "✅ ON TRACK", IF(F5>=0.92, "⚠️ AT RISK", "❌ BELOW"))

-- FY Total revenue
=SUM(C5:C16)

-- Total revenue KPI card on Dashboard
=SUMIF('Raw Data'!D:D, "Energy Drink", 'Raw Data'!K:K) + SUMIF(...)

-- YTD Attainment linked from KPI Tracker
='KPI Tracker'!F17
```

**Features demonstrated:** SUMIFS, SUMIF, XLOOKUP, IFERROR, RANK, nested IF, SUM with expanding range, cross-sheet references, dynamic formulas

---

## 📊 Key Findings — FY 2023

- **Energy Drink leads revenue** at ₹57.35L — highest revenue per unit at ₹125 (5x higher than Water)
- **June–July 2023 were peak months** — combined +34% MoM surge driven by summer seasonality
- **FY 2023 closed at 96.2% of annual target** — Q4 shortfall driven by post-monsoon demand slowdown
- **West India contributed 29.8%** of total revenue — highest regional share
- **East India is underpenetrated** at 22% share — largest growth opportunity in FY 2024
- **Promotional campaigns show positive lift** across Energy Drink and Sports Drink categories

---

## 🛠️ Tools & Skills

- **Microsoft Excel / WPS Office** — Workbook design, formulas, charts, formatting
- **Excel Formulas** — SUMIFS, SUMIF, XLOOKUP, IFERROR, RANK, nested IF, expanding SUM range
- **Data Visualization** — 5 chart types (Combo, Doughnut, Bar, Line, Clustered Column)
- **Conditional Formatting** — Traffic lights, color scales, custom rules
- **Pivot Analysis** — Manual pivot-style tables with grand totals
- **Python** — Synthetic dataset generation (used only for data creation)

---

## 📁 Files in This Repository

```
├── Beverages_India_MIS_Report_FINAL.xlsx    # Main Excel workbook
├── Screenshots/
│   ├── Sheet1_Pivot_Analysis.png
│   ├── Sheet2_KPI_Tracker.png
│   ├── Sheet3_Dashboard.png
│   └── Sheet4_Insight_Summary.png
└── README.md
```

---

## 🚀 How to Use

1. Download **Beverages_India_MIS_Report_FINAL.xlsx**
2. Open in **Microsoft Excel desktop** or **WPS Office** (free)
3. Enable editing if prompted
4. All formulas calculate automatically from Raw Data sheet
5. KPI Tracker, Dashboard, and charts all update based on Raw Data

---

## 👤 About

Built by **Saad** — CS Engineering Student at Parul University transitioning into Data Analytics.

This project is part of a **targeted FMCG analytics portfolio** covering Power BI, Excel, MySQL, and Python — simulating real-world data work in the consumer goods and retail analytics space.

🔗 [LinkedIn](https://linkedin.com/in/saad-memon-49aa75350) | 📧 saadmemon1104@gmail.com | 🐙 [GitHub](https://github.com/saadmemon11)
