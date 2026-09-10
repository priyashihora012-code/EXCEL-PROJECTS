<div align="center">

# -- ! Sales Performance Dashboard & Analyzer ! --
### *Interactive Excel-Based Sales Analytics, Regression & Pivot Reporting*

[![Excel](https://img.shields.io/badge/Excel-365%2B-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Formulas](https://img.shields.io/badge/Formulas-SLOPE%2FINTERCEPT%2FRSQ-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-Auto%20Updating-4CAF50?style=for-the-badge&logo=googlesheets&logoColor=white)](https://www.microsoft.com/excel)
[![Analytics](https://img.shields.io/badge/Analytics-Regression%20%26%20Pivot-9C27B0?style=for-the-badge&logo=googleanalytics&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"Data tells a story — a good dashboard just makes sure everyone can read it."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📊 Part A — Dashboard & Dataset](#-part-a--dashboard--dataset)
- [📈 Part B — Regression, Trend & Pivot Analysis](#-part-b--regression-trend--pivot-analysis)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Sales Performance Dashboard & Analyzer (PR2 Analyzer)** is a fully formula-driven Excel workbook that transforms 2,000 rows of raw sales transactions into a live, auto-updating business dashboard. It combines **KPI cards**, **charts**, **linear regression**, **month-over-month trend tracking**, and a **region × product pivot table** — all built using native Excel formulas, without any macros or add-ins.

This project is designed to:
- Turn raw transactional data into decision-ready KPIs
- Demonstrate statistical analysis (regression) using built-in Excel functions
- Track monthly sales growth with conditional arrow formatting
- Summarize multi-dimensional data using a pivot-style matrix

---

## 🎯 Problem Statement

> **Objective:** Build an Excel workbook that automatically analyzes sales data and presents it as a live, easy-to-read performance dashboard.

You are building an analytics workbook for a retail/electronics business. The workbook must ingest raw customer-level sales data and automatically compute totals, profit margins, regional performance, top customers, monthly growth trends, a profit-vs-sales regression model, and a region-by-product breakdown — updating instantly whenever the underlying `Sales_Data` sheet changes.

| 📂 Feature | 📄 Sheet | 🔍 Description |
|------------|----------|----------------|
| Executive Dashboard | `Dashboard` | KPI cards, charts & key insights |
| Raw Transactions | `Sales_Data` | 2,000-row customer sales dataset |
| Descriptive Stats | `Descriptive_Stats` | Summary statistics of the dataset |
| Regression Model | `Regression_Analysis` | Profit vs Sales linear regression |
| Scenario Testing | `WhatIf_Analysis` | What-if projections |
| Top Customers | `High_Value_Customers` | Ranked high-value customer list |
| Monthly Trend | `Monthly_Trend` | Month-over-month growth with arrows |
| Region × Product | `Pivot_Summary` | Cross-tab sales pivot table |

The goal is to demonstrate **applied Excel analytics** — formulas, charts, and conditional formatting working together as a self-updating business tool.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **Auto-Updating Dashboard** | KPI cards & charts refresh instantly when `Sales_Data` changes |
| 💰 **KPI Cards** | Total Sales, Total Profit, Total Orders, Avg Order Value, Profit Margin % |
| 📈 **Linear Regression** | Profit vs Sales modeled using `SLOPE`, `INTERCEPT`, `RSQ`, `CORREL` |
| 📅 **Monthly Trend Tracking** | Month-by-month sales & profit with growth % and arrow icons |
| 🧮 **Pivot-Style Summary** | Region × Product breakdown using formula-based cross-tabulation |
| 🏆 **High-Value Customer List** | Customers ranked by total purchase value |
| 🎯 **What-If Analysis** | Scenario projections built on top of the regression model |
| 🖥️ **Chart Visuals** | Bar chart (region), line chart (monthly trend), pie chart (product share) |
| ⚠️ **Key Insights Panel** | Auto-generated text summary of top region, product & customer |

---

## 🏗️ Project Structure

```
📦 project-2/
│
├── 📊 PR2_Analyzer.xlsx      ← Main Excel workbook (entry point)
├── 🖼️ 1_DASHBOARD.png        ← Screenshot: Dashboard sheet
├── 🖼️ 2_DATASET.png          ← Screenshot: Sales_Data sheet
├── 🖼️ 3_LINEAR-REG.png       ← Screenshot: Regression_Analysis sheet
├── 🖼️ 4_MONTHLY-SALES.png    ← Screenshot: Monthly_Trend sheet
├── 🖼️ 5_PIVOT-TABLE.png      ← Screenshot: Pivot_Summary sheet
│
└── 📄 README.md              ← Project documentation
```

---

## 🔄 Project Workflow

```
Raw Data Entry
      │
      ▼
┌─────────────────────────────┐
│      Sales_Data Sheet       │  ← 2,000 rows: Customer, Date, Region,
└────────────┬────────────────┘     Product, Qty, Price, Discount, Profit
             │
     ┌───────┴────────────────────┬──────────────────────┐
     ▼                            ▼                       ▼
┌─────────────┐        ┌──────────────────┐   ┌─────────────────────┐
│ Descriptive │        │   Regression      │   │  Monthly_Trend /    │
│ Stats       │        │   Analysis        │   │  Pivot_Summary      │
└──────┬──────┘        └────────┬──────────┘   └──────────┬──────────┘
       │                        │                          │
       └────────────┬───────────┴──────────────┬───────────┘
                     ▼                          ▼
           ┌──────────────────┐      ┌────────────────────┐
           │ High_Value_Cust.  │      │  WhatIf_Analysis    │
           └─────────┬─────────┘      └──────────┬──────────┘
                     │                            │
                     └─────────────┬──────────────┘
                                   ▼
                     ┌─────────────────────────────┐
                     │      Dashboard Sheet        │
                     │  KPIs · Charts · Insights   │
                     └─────────────────────────────┘
```

---

## 📊 Part A — Dashboard & Dataset

### 📝 1. What is the Dashboard Sheet?

The `Dashboard` sheet is the executive summary layer of the workbook. It pulls live figures from `Sales_Data` using aggregate formulas (`SUM`, `SUMIFS`, `AVERAGE`) and presents them as KPI cards, a key-insights panel, and three charts — bar, line, and pie.

**Live KPI Cards:**

| KPI | Value |
|-----|-------|
| 💰 Total Sales | ₹13,84,13,167 |
| 📈 Total Profit | ₹2,70,99,906 |
| 🧾 Total Orders | 2,000 |
| 🛒 Avg Order Value | ₹69,207 |
| 📐 Profit Margin % | 19.6% |

**Key Insights (auto-generated):**
- 🏆 Top performing region: **Central** — Total sales ₹2,91,89,150
- 💻 Best-selling product: **Laptop** — Total sales ₹4,43,63,713
- 👤 Top customer by value: **Rahul Desai** — ₹18,59,128 in purchases
- 📊 Profit vs Sales correlation (R²): **74.7%**
- 📅 Strongest sales month: **Feb-2024** with ₹71,99,783

**Screenshot:**

![Sales Performance Dashboard](1_DASHBOARD.png)

---

### 🗂️ 2. Sales_Data Sheet — Overview

> The raw dataset that powers every other sheet in the workbook. Each row is one customer transaction.

| Column | Field | Description |
|--------|-------|-------------|
| A | `Customer_ID` | Unique customer identifier |
| B | `Customer_Name` | Customer's full name |
| C | `Date` | Transaction date |
| D | `Region` | North / South / East / West / Central |
| E | `Product` | Laptop, Smartphone, Tablet, Camera, etc. |
| F | `Quantity` | Units purchased |
| G | `Unit_Price` | Price per unit |
| H | `Discount_%` | Discount applied |
| I | `Total_Purchase` | Quantity × Unit Price (net of discount) |
| J | `Total_Cost` | Cost basis of the order |
| K | `Total_Profit` | `Total_Purchase − Total_Cost` |
| L | `Report_Timestamp` | Auto-updated recalculation timestamp |

**Screenshot:**

![Sales_Data Dataset](2_DATASET.png)

---

## 📈 Part B — Regression, Trend & Pivot Analysis

### 🔍 3. Linear Regression — Profit vs Sales

> Models the relationship between total sales and total profit using native Excel regression formulas (an equivalent to Data Analysis ToolPak → Regression).

**Formulas Used:**
```
Slope (b1)     = SLOPE(Profit_Range, Sales_Range)
Intercept (b0) = INTERCEPT(Profit_Range, Sales_Range)
R Square       = RSQ(Profit_Range, Sales_Range)
Multiple R     = CORREL(Profit_Range, Sales_Range)
```

**Regression Statistics:**

| Statistic | Value |
|-----------|-------|
| Slope (b1) | 0.1966 |
| Intercept (b0) | -53.2722 |
| R Square | 0.7468 |
| Multiple R (Correlation) | 0.8642 |
| Standard Error (Estimate) | 9754.5465 |
| Observations | 2,000 |

**Regression Equation:**
```
Profit = 0.1966 × Sales + (-53.2722)
```

**Interpretation:** For every extra ₹1 in Sales, Profit changes by approximately ₹0.1966, and the model explains 74.7% of the variation in Profit (R²).

**Screenshot:**

![Linear Regression: Profit vs Sales](3_LINEAR-REG.png)

---

### 📅 4. Monthly Sales Trend & Growth

> Tracks Total Sales and Total Profit month by month, with a `Growth %` column that compares each month to the previous one — highlighted using custom conditional arrow formatting (🔼 up / 🔽 down / ➡️ flat).

**Logic:**
```
Growth % = (Current_Month_Sales − Previous_Month_Sales) / Previous_Month_Sales
```

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 📊 `SUMIFS` | Aggregates sales/profit per month from `Sales_Data` |
| ➗ Percentage Change | `(current − previous) / previous` |
| 🎨 Conditional Formatting | Custom icon set — green up-arrow, red down-arrow |
| 📐 Date Grouping | `EOMONTH` / `TEXT` for month bucketing |

**Screenshot:**

![Monthly Sales Trend & Growth](4_MONTHLY-SALES.png)

---

### 🧮 5. Pivot Table — Region × Product

> A formula-built cross-tabulation summarizing Total Sales for every Region–Product combination, with row and column totals.

**Logic:**
```
Cell Value = SUMIFS(Total_Purchase, Region_Range, Row_Region, Product_Range, Column_Product)
Region Total = SUM(all products for that region)
Product Total = SUM(all regions for that product)
```

**Sample Output:**

| Region \ Product | Laptop | Smartphone | ... | Region Total |
|---|---|---|---|---|
| North | ₹87,05,085 | ₹44,88,789 | ... | ₹2,64,48,617 |
| Central | ₹80,66,958 | ₹61,23,453 | ... | ₹2,91,89,150 |
| **Product Total** | **₹4,43,63,713** | **₹2,24,03,766** | ... | **₹13,84,13,167** |

**Screenshot:**

![Pivot Table: Total Sales by Region and Product](5_PIVOT-TABLE.png)

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📊 **Microsoft Excel** | 365 / 2019+ | Core spreadsheet application |
| 🧮 **SLOPE / INTERCEPT / RSQ / CORREL** | Built-in | Linear regression statistics |
| 🔢 **SUMIFS / AVERAGEIFS** | Built-in | Conditional aggregation across sheets |
| 📈 **Native Charts** | Built-in | Bar, Line & Pie chart visuals |
| 🎨 **Conditional Formatting** | Built-in | Growth-arrow icon sets, color scales |
| 📐 **Structured Tables** | Excel Tables | Auto-expanding `Sales_Data` with filters |

---

## 📈 Results & Insights

After opening the workbook, the following outputs are produced:

- ✅ **Live KPI Dashboard** — Total Sales, Profit, Orders, AOV & Margin update instantly
- 📊 **3 Chart Visuals** — Region bar chart, monthly trend line chart, product-share pie chart
- 📈 **Regression Model** — Profit vs Sales relationship quantified (R² = 74.7%)
- 📅 **12+ Month Trend** — Growth % with color-coded up/down arrows
- 🧮 **Region × Product Pivot** — Full cross-tab of ₹13,84,13,167 total sales
- 🏆 **Top Performers Identified** — Central region, Laptop product, Rahul Desai (customer)

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **No Macros Needed** | Fully built with native Excel formulas |
| 🔄 **Auto-Refreshing** | Dashboard updates the moment `Sales_Data` changes |
| 📚 **Educational** | Demonstrates regression, pivoting & KPI design in one file |
| 🖥️ **Single Workbook** | Everything — data, analysis & dashboard — in one `.xlsx` |
| ⚡ **Lightweight** | No external plugins, add-ins or Data Analysis ToolPak required |
| 🧪 **Extensible** | Easy to add more regions, products, or new KPI cards |
| 📖 **Readable Layout** | Clearly labeled sheets and color-coded headers |
| 🛡️ **Data Integrity** | Structured Excel Tables keep formulas consistent as rows are added |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Priya Shihora

[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/priya)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/priya-shihora -686533312/)

> *"A dashboard is only as good as the story it tells at a glance."*

**🎓 Role:** Junior Data/Excel Analyst | Programming Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel · Data Analysis · Dashboards · Regression · Pivot Tables

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support Docs](https://support.microsoft.com/excel) — Official Excel function reference
- 🧮 [Microsoft — SLOPE/INTERCEPT/RSQ Functions](https://support.microsoft.com/en-us/office/slope-function-11fb8f97-3117-4813-94b4-af0000c99a11) — Regression formula reference
- 📊 [ExcelJet](https://exceljet.net/) — Formula patterns & dashboard tips
- 📐 [Chandoo.org](https://chandoo.org/wp/) — Dashboard design best practices
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support
- 📖 [Kaggle Learn](https://www.kaggle.com/learn) — Data analysis courses

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: 10 Sep, 2026*

</div>
