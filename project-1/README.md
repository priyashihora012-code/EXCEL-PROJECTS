<div align="center">

# -- ! Excel Fundamentals Booster ! --
### *A Multi-Sheet Excel Workbook Demonstrating Core & Advanced Spreadsheet Formulas*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-365%2F2019%2B-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-Lookup%20%26%20Logical-FF6F00?style=for-the-badge&logo=googlesheets&logoColor=white)](https://support.microsoft.com/excel)
[![Data](https://img.shields.io/badge/Data-Multi%20Sheet%20Workbook-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Math](https://img.shields.io/badge/Math-Conditional%20%26%20Date%20Logic-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)

<br/>

> *"Formulas are the heartbeat of a spreadsheet — master them, and raw data becomes insight."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🎓 Part A — Students Sheet](#-part-a--students-sheet)
- [🛒 Part B — Sales Sheet](#-part-b--sales-sheet)
- [👔 Part C — Employees Sheet](#-part-c--employees-sheet)
- [📚 Part D — Lookup\_Data Sheet](#-part-d--lookup_data-sheet)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Fundamentals Booster** is a beginner-to-intermediate friendly, multi-sheet Excel workbook that demonstrates core and advanced spreadsheet concepts such as **lookup functions**, **conditional logic**, **aggregation formulas**, and **date/number handling**. The workbook is organized into four connected sheets, each built around a real-world scenario.

This project is designed to:
- Strengthen understanding of lookup formulas (`VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `XMATCH`)
- Practice conditional aggregation (`COUNTIFS`, `SUMIFS`, `AVERAGEIFS`)
- Apply logical and text formulas (`IF`, `AND`, `OR`, `LEFT`, `UPPER`, `LOWER`, `FIND`)
- Work with dates, rounding, and dynamic ranges (`DATEDIF`, `ROUND`, `CEILING`, `FLOOR`, `OFFSET`, `INDIRECT`)

---

## 🎯 Problem Statement

> **Objective:** Build a single Excel workbook that showcases practical formula usage across student, sales, and employee datasets, backed by shared reference tables.

You are building a self-contained practice workbook for anyone learning Excel. The workbook must hold raw records on one sheet, compute derived columns using formulas (never hardcoded values), and include a dedicated "Analysis / Formula Demos" block per sheet that answers real questions using lookup and aggregation functions.

| 📂 Sheet | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| Students | Data + Formulas | Grades, averages, grading logic, and text functions |
| Sales | Data + Formulas | Transactions, tax, discount, and lookup-driven pricing |
| Employees | Data + Formulas | Age, tenure, and bonus calculations with date logic |
| Lookup\_Data | Reference Tables | Product master and monthly salesperson performance |

The goal is to demonstrate **fundamental-to-advanced Excel skills** through a clean, formula-driven, multi-sheet workbook.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **4 Connected Sheets** | Students, Sales, Employees, and Lookup\_Data working together |
| 🔍 **Lookup Function Suite** | `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `XMATCH` side by side |
| ➕ **Conditional Aggregation** | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` over real datasets |
| 🧮 **Logical Grading Engine** | Nested `IF` and `AND`/`OR` for grades, bonuses, and flags |
| 🔤 **Text Function Demos** | `LEFT`, `FIND`, `UPPER`, `LOWER` for name parsing |
| 📅 **Date & Tenure Logic** | `DATEDIF` and `TODAY()` for live age and days-since-joining |
| 🎯 **Dynamic Referencing** | `OFFSET` and `INDIRECT` for flexible, movable ranges |
| ✅ **Fallback-Safe Formulas** | `IFERROR` pairs modern functions with classic equivalents |

---

## 🏗️ Project Structure

```
📦 project-1/
│
├── 📊 Fundamentals_Booster.xlsx    ← Main Excel workbook (Students, Sales, Employees, Lookup_Data)
│
└── 📄 README.md                    ← Project documentation
```

---

## 🔄 Project Workflow

```
Workbook Open
      │
      ▼
┌─────────────────────────────┐
│   Lookup_Data (Reference)   │  ← Product Master + Salesperson Performance
└────────────┬────────────────┘
             │ feeds
     ┌───────┼────────────┬────────────────┐
     ▼                    ▼                 ▼
┌─────────────┐   ┌──────────────┐   ┌──────────────┐
│  Students   │   │    Sales     │   │  Employees   │
│  Sheet      │   │    Sheet     │   │  Sheet       │
└──────┬──────┘   └──────┬───────┘   └──────┬───────┘
       │                 │                  │
       ▼                 ▼                  ▼
┌─────────────┐   ┌──────────────┐   ┌──────────────┐
│ Grades /    │   │ Totals /     │   │ Age / Tenure │
│ Averages    │   │ Discounts    │   │ / Bonus      │
└──────┬──────┘   └──────┬───────┘   └──────┬───────┘
       │                 │                  │
       └────────┬────────┴──────────┬───────┘
                 ▼                   ▼
       ┌─────────────────────────────────┐
       │  Analysis / Formula Demos Block  │  ← COUNTIFS, SUMIFS, VLOOKUP,
       │  (per sheet)                     │     XLOOKUP, INDEX/MATCH, etc.
       └─────────────────────────────────┘
```

---

## 🎓 Part A — Students Sheet

### 📝 1. What does this sheet cover?

The **Students** sheet holds 15 student records and computes grade-related metrics using nested logical formulas and text functions — building intuition for conditional grading and string manipulation.

---

### 🗺️ 2. Derived Columns — Overview

| Column | Purpose | Logic Used |
|--------|---------|------------|
| Total | Sum of Math + Science | `=C+D` |
| Average | Total ÷ 2 | `=E/2` |
| Grade | A–F band from Average | Nested `IF` |
| Above80\_Both | Both subjects above 80? | `AND` |
| First Name | Text before the first space | `LEFT` + `FIND` |
| Name (UPPER/LOWER) | Case conversion | `UPPER` / `LOWER` |

---

### 🎯 3. Grading Logic

> Nested `IF` bands the average score into a letter grade.

**Logic:**
```excel
=IF(F4>=90,"A",IF(F4>=75,"B",IF(F4>=60,"C",IF(F4>=40,"D","F"))))
```

**Sample Output:**
```
Aarav Sharma  | Avg 88.5 | Grade B
Anaya Joshi   | Avg 92   | Grade A
Myra Nair     | Avg 36.5 | Grade F
```

---

### 🔎 4. Formula Demos on this Sheet

**Logic:**
```excel
=COUNTIFS(C4:C18,">60")                         ' Students scoring above 60 in Math
=AVERAGEIFS(F4:F18,C4:C18,">60")                 ' Average score of that group
=IFERROR(XMATCH("S105",A4:A18),MATCH("S105",A4:A18,0))   ' Position lookup, modern + safe fallback
=VLOOKUP("S107",A4:B18,2,FALSE())                ' Name lookup by Student ID
```

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🔁 `COUNTIFS` / `AVERAGEIFS` | Conditional counting and averaging |
| 🎯 `XMATCH` with `IFERROR` fallback | Modern lookup with a `MATCH`-based safety net |
| 🔍 `VLOOKUP` | Classic exact-match lookup by ID |
| ✂️ `LEFT` + `FIND` | Extracts first name up to the first space |

---

## 🛒 Part B — Sales Sheet

### 📝 5. What does this sheet cover?

The **Sales** sheet tracks 10 transactions and calculates pricing, discounts, and tax by pulling live prices from the `Lookup_Data` sheet — demonstrating cross-sheet references and multi-condition logic.

**Logic:**
```excel
=VLOOKUP(B4,Lookup_Data!$A$5:$C$12,3,FALSE())    ' Unit price via product code
=E4*F4                                            ' Total = Qty × Unit Price
=IF(OR(G4>10000,E4>10),0.1,0)                     ' 10% discount if total or qty is high
=G4*(1-H4)                                        ' Final amount after discount
=G4*(1+$H$1)                                      ' Total with tax (absolute reference)
```

**Sample Output:**
```
O1002 | 27-inch Monitor | Qty 1 | Total ₹12,999 | Discount 10% | Final ₹11,699.10
```

---

### 🔎 6. Formula Demos on this Sheet

| Formula Demo | Concept |
|--------------|---------|
| `SUMIFS(G4:G13,C4:C13,"North",B4:B13,"P001")` | Multi-criteria sales total |
| `INDEX/MATCH` on `Lookup_Data` | Two-way lookup (Salesperson × Month) |
| `XLOOKUP` with `IFERROR` fallback | Modern flexible search, same result as `INDEX/MATCH` |
| `XMATCH` with `MATCH` fallback | Position lookup for a Product Code |
| `OFFSET(Lookup_Data!$E$5,0,2,1,3)` | Dynamic 3-month rolling sum |
| `INDIRECT("Lookup_Data!F6:F10")` | Range built dynamically from text |

---

## 👔 Part C — Employees Sheet

### 📝 7. What does this sheet cover?

The **Employees** sheet holds 8 employee records and computes live age, tenure, and bonus figures — reinforcing date arithmetic and rounding functions.

**Logic:**
```excel
=DATEDIF(D4,TODAY(),"y")     ' Current age in completed years
=TODAY()-E4                   ' Days since joining
=F4*0.1                       ' Bonus at 10% of salary
=ROUND(I4,0)                  ' Bonus rounded to nearest whole number
=CEILING(I4,100)              ' Bonus rounded up to nearest 100
=FLOOR(I4,100)                ' Bonus rounded down to nearest 100
```

**Sample Output:**
```
Rahul Bhatt | Age 40 | Days Since Joining 4805 | Bonus ₹8,800 (Ceiling & Floor: ₹8,800)
```

A small **XLOOKUP Demo** search box lets you type any Employee ID and instantly see their salary, using `XLOOKUP` with an `INDEX`/`MATCH` fallback for older Excel versions.

---

## 📚 Part D — Lookup\_Data Sheet

### 📝 8. What does this sheet cover?

The **Lookup\_Data** sheet is the shared reference layer — a **Product Master** table and a **Salesperson Performance** table (monthly sales, Jan–Jun) — feeding the `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `OFFSET`, and `INDIRECT` formulas used across the other sheets.

| Table | Columns | Feeds Into |
|-------|---------|------------|
| Product Master | ProductCode, ProductName, Price | Sales sheet pricing |
| Salesperson Performance | Salesperson, Jan–Jun sales | Sales sheet analysis demos |

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📊 **Microsoft Excel** | 365 / 2019+ | Core spreadsheet application |
| 🔍 **Lookup Functions** | `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `XMATCH` | Data retrieval across sheets |
| ➕ **Aggregate Functions** | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` | Conditional totals and averages |
| 🧮 **Logical Functions** | `IF`, `AND`, `OR`, `IFERROR` | Grading, flags, and safe fallbacks |
| 📅 **Date Functions** | `DATEDIF`, `TODAY()` | Live age and tenure calculations |
| 🔢 **Rounding Functions** | `ROUND`, `CEILING`, `FLOOR` | Bonus and numeric formatting |
| 🎯 **Dynamic Reference Functions** | `OFFSET`, `INDIRECT` | Flexible, movable range building |
| ✂️ **Text Functions** | `LEFT`, `FIND`, `UPPER`, `LOWER` | Name parsing and case conversion |

---

## 📈 Results & Insights

After opening the workbook, the following outputs are available:

- ✅ **4 Linked Sheets** — Students, Sales, Employees, and Lookup\_Data working together
- 🔢 **Live Formula Results** — Every derived column recalculates automatically on edit
- 🔍 **Modern vs Classic Lookups** — `XLOOKUP`/`XMATCH` shown side by side with `VLOOKUP`/`INDEX`-`MATCH` fallbacks
- 📅 **Live Date Logic** — Age and days-since-joining update automatically with `TODAY()`
- ⚠️ **Fallback Safety** — `IFERROR` ensures older Excel versions still return correct results

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | Core-to-advanced formulas across one compact workbook |
| 🔄 **Reusability** | Sheet structure can be duplicated for new datasets |
| 📚 **Educational** | Each sheet reinforces a different formula category |
| 🖥️ **No Add-ins Needed** | Works with built-in Excel functions only |
| ⚡ **Lightweight** | Single workbook, instantly usable in Excel or Google Sheets |
| 🧪 **Extensible** | Easy to add new sheets (Inventory, Attendance, etc.) |
| 📖 **Readable Layout** | Clear headers and an "Analysis / Formula Demos" block per sheet |
| 🛡️ **Version-Safe** | Modern functions paired with `IFERROR` fallbacks for compatibility |

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

> *"Every insight starts with a single formula — just like every workbook starts with a single cell."*

**🎓 Role:** Excel Enthusiast | Data & Formula Learner \
**📍 Location:** India\
**🛠️ Skills:** Microsoft Excel · Lookup Functions · Data Analysis · Logic Building · Spreadsheet Design

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support Docs](https://support.microsoft.com/excel) — Official Excel function reference
- 🔍 [Exceljet — Formulas](https://exceljet.net/formulas) — In-depth formula tutorials
- 📐 [Excel Easy](https://www.excel-easy.com/) — Beginner Excel reference
- 🧮 [Chandoo.org](https://chandoo.org/wp/) — Excel tips and formula guides
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support
- 📖 [Kaggle Learn](https://www.kaggle.com/learn) — Data analysis and spreadsheet courses

---

<div align="center">

---

*Made with ❤️ and 📊 — Last updated: 27 August, 2026*

</div>
