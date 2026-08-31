<div align="center">

# -- ! Excel Fundamentals Booster ! --
### *A Multi-Sheet Excel Workbook Mastering Lookup, Logical & Date Functions*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Lookup Functions](https://img.shields.io/badge/Functions-VLOOKUP%20%7C%20XLOOKUP%20%7C%20INDEX%2FMATCH-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Logical](https://img.shields.io/badge/Logic-IF%20%7C%20OR%20%7C%20SUMIFS-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Dynamic Range](https://img.shields.io/badge/Dynamic-INDIRECT%20%7C%20OFFSET-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"A spreadsheet is only as smart as the formulas behind it — master the functions, and data becomes insight."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🧑‍🎓 Sheet A — Students](#-sheet-a--students)
- [💰 Sheet B — Sales](#-sheet-b--sales)
- [🧑‍💼 Sheet C — Employees](#-sheet-c--employees)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Fundamentals Booster** is a beginner-to-intermediate practice workbook that demonstrates core Microsoft Excel concepts such as **lookup functions**, **logical conditions**, **date & time calculations**, **text manipulation**, and **dynamic range referencing**. The workbook is organized into three real-world themed sheets, each targeting a different family of Excel functions.

This project is designed to:
- Strengthen understanding of `VLOOKUP`, `XLOOKUP`, `XMATCH`, and `INDEX`/`MATCH`
- Practice conditional aggregation using `SUMIFS`, `COUNTIF`, and `AVERAGEIF`
- Apply logical and text functions (`IF`, `OR`, `LEFT`, `FIND`, `UPPER`, `LOWER`)
- Explore dynamic and volatile functions like `INDIRECT` and `OFFSET`
- Work with date arithmetic (`age`, `years of service`, `days since joining`)

---

## 🎯 Problem Statement

> **Objective:** Build a single Excel workbook that showcases fundamental-to-advanced spreadsheet formulas across multiple realistic datasets.

You are building a self-practice workbook for someone learning Excel. Each sheet must hold a small real-world dataset (students, sales, employees) and use formulas to derive computed columns, classifications, and quick lookup panels — instead of hardcoded values.

| 📂 Sheet | 📄 Type | 🔍 Description |
|----------|---------|-----------------|
| Students | Academic Data | Grades, averages, name parsing, and score-based lookups |
| Sales | Transactional Data | Product pricing, discount eligibility, and dynamic sales analysis |
| Employees | HR Data | Salary calculations, tenure tracking, and employee lookups |

The goal is to demonstrate **fundamental-to-intermediate Excel skills** through a clean, formula-driven, multi-sheet workbook.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **3 Themed Sheets** | Students, Sales, and Employees — each with a distinct dataset |
| 🔍 **Lookup Family** | `VLOOKUP`, `XLOOKUP`, `XMATCH`, and `INDEX`/`MATCH` used side-by-side |
| ➕ **Conditional Aggregation** | `SUMIFS` for region + product-wise totals |
| 🔤 **Text Extraction** | `LEFT` + `FIND` to pull first names; `UPPER`/`LOWER` for casing |
| 📅 **Date Arithmetic** | Age, years of service, and days-since-joining calculations |
| 🧮 **Logical Classification** | `IF`/`OR` combinations for grades and discount eligibility |
| 🔗 **Dynamic Ranges** | `INDIRECT` for text-based ranges and `OFFSET` for rolling trends |
| 🔒 **Absolute Referencing** | Fixed bonus-rate cell reused across every employee row |

---

## 🏗️ Project Structure

```
📦 project-1/
│
├── 📄 pr-1.xlsx   ← Main Excel workbook (entry point)
│      ├── 🧑‍🎓 Students
│      ├── 💰 Sales
│      └── 🧑‍💼 Employees
│
└── 📄 README.md                    ← Project documentation
```

---

## 🔄 Project Workflow

```
Open Workbook
      │
      ▼
┌─────────────────────────────┐
│   Select a Sheet Tab        │  ← Students / Sales / Employees
└────────────┬────────────────┘
             │
   ┌─────────┼─────────────┐
   ▼         ▼              ▼
┌─────────┐ ┌─────────┐ ┌──────────┐
│Students │ │  Sales  │ │Employees │
└────┬────┘ └────┬────┘ └────┬─────┘
     │            │            │
     ▼            ▼            ▼
┌──────────┐ ┌───────────┐ ┌────────────┐
│ Grades & │ │ Pricing & │ │ Salary &   │
│ Names    │ │ Discounts │ │ Tenure     │
└────┬─────┘ └─────┬─────┘ └─────┬──────┘
     │              │              │
     ▼              ▼              ▼
┌─────────────────────────────────────┐
│   Formula-Driven Lookup Panels      │
│   (VLOOKUP / XLOOKUP / INDEX-MATCH) │
└──────────────────┬───────────────────┘
                    ▼
            Computed Results ✅
```

---

## 🧑‍🎓 Sheet A — Students

### 📝 1. What This Sheet Covers

The **Students** sheet tracks academic performance for 10 students and demonstrates scoring, grading, name-parsing, and quick-lookup logic — all driven by formulas.

---

### 🗺️ 2. Columns & Logic — Overview

| Column | Logic Used |
|--------|------------|
| Age | Calculated from **Date of Birth** |
| Total / Average | `SUM` and `AVERAGE` of Math + Science |
| Grade | `IF` ladder on the Average score |
| Both > 80 | `IF`/`AND` check across Math and Science |
| First Name | `LEFT` + `FIND` to isolate the first word |
| Upper / Lower Name | `UPPER()` and `LOWER()` text conversion |

---

### 🔺 3. Grading Logic

> Classifies each student into a letter grade based on their Average.

**Logic:**
```
=IF(Average>=90,"A",IF(Average>=80,"B",IF(Average>=70,"C","D")))
```

**Sample Output:**
```
Riya Shah   → Average 92.5 → Grade A
Aarav Sharma→ Average 86.5 → Grade B
Rahul Mehta → Average 57.5 → Grade D
```

---

### 🏔️ 4. First Name Extraction

> Pulls the first name out of a "Full Name" cell using position logic.

**Logic:**
```
=LEFT(FullName, FIND(" ", FullName) - 1)
```

**Output:**
```
"Aarav Sharma" → "Aarav"
"Priya Patel"  → "Priya"
```

---

### 🔍 5. Quick Lookup Panels

> Side panels answer specific questions using `COUNTIF`, `AVERAGEIF`, and lookup formulas.

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🔁 `COUNTIF` | Students scoring above 50 in Math / above 60 overall |
| ➗ `AVERAGEIF` | Average score of students above a threshold |
| 🔎 Lookup by ID | Retrieve a student's name from their Student ID |
| 📋 Filtered List | Names of all students with grades above 80 |

---

## 💰 Sheet B — Sales

### 📝 6. What This Sheet Covers

The **Sales** sheet logs 10 transactions across regions and products, and demonstrates pricing lookups, discount rules, and several ways to summarize sales data dynamically.

---

### 🗺️ 7. Columns & Logic — Overview

| Column / Panel | Logic Used |
|-----------------|------------|
| Product Price | `VLOOKUP` against the Product Master table |
| Total Sale | Units Sold × Product Price |
| Discount Eligible | `IF`/`OR` on quantity and sale-value thresholds |
| Product Master | Reference table for `VLOOKUP`/`XMATCH` |

---

### 🔺 8. Region + Product Sales Total

> Sums total sales filtered by both Region and Product Code simultaneously.

**Logic:**
```
=SUMIFS(TotalSale, Region, "South", ProductCode, "P101")
```

**Output (example):**
```
Region: South | Product: P101 → Total Sale: 2250
```

---

### 🏔️ 9. Salesperson Lookup (XLOOKUP)

> Retrieves the total sale value for a chosen salesperson.

**Logic:**
```
=XLOOKUP("Karan", Salesperson, TotalSale)
```

---

### 🔻 10. Salesperson & Month Lookup (INDEX/MATCH)

> Finds the sale value for a specific salesperson in a specific month using a two-way `INDEX`/`MATCH`.

**Logic:**
```
=INDEX(SalesGrid, MATCH(Salesperson, RowHeaders, 0), MATCH(Month, ColHeaders, 0))
```

**Sample Output:**
```
Salesperson: Karan | Month: February → Sale Value: 45600
```

---

### 🔗 11. Dynamic Range Functions

> Demonstrates two ways to reference ranges without hardcoding them.

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🧩 `INDIRECT` | Sums a range built from text, e.g. `"H2:H11"` |
| ⏱️ `OFFSET` | Rolling total of the last **N** sales, shifting automatically |

---

## 🧑‍💼 Sheet C — Employees

### 📝 12. What This Sheet Covers

The **Employees** sheet manages HR data for 10 employees, focusing on tenure calculations, salary rounding rules, and absolute-reference-based bonus computation.

---

### 🗺️ 13. Columns & Logic — Overview

| Column | Logic Used |
|--------|------------|
| Age / Years of Service | Derived from Date of Birth and Date of Joining |
| Days Since Joining | Current date minus Date of Joining |
| Bonus Amount (10%) | Salary × **absolute-referenced** bonus rate |
| Salary Rounded | `ROUND()` to the nearest rupee |
| Salary Ceiling / Floor | `CEILING()` / `FLOOR()` to the nearest 500 |

---

### 🔺 14. Bonus Calculation (Absolute Reference)

> Every employee's bonus is calculated by multiplying their salary with a single fixed bonus-rate cell, locked with `$` so it never shifts when copied down.

**Logic:**
```
=Salary * $BonusRate$
```

**Sample Output:**
```
Salary: 68500 | Bonus Rate: 10% → Bonus Amount: 6850
```

---

### 🏔️ 15. Employee Lookup Panels

> Two lookup methods retrieve employee details from an ID.

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🔎 `XLOOKUP` | Find Salary directly by Emp ID |
| 🧭 `INDEX`/`MATCH` | Retrieve Department and Date of Joining by Emp ID |

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📊 **Microsoft Excel** | 2019+ / 365 | Core spreadsheet application |
| 🔍 **Lookup Functions** | Built-in | `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX`/`MATCH` |
| ➕ **Aggregate Functions** | Built-in | `SUMIFS`, `COUNTIF`, `AVERAGEIF` |
| 🔤 **Text Functions** | Built-in | `LEFT`, `FIND`, `UPPER`, `LOWER` |
| 📅 **Date Functions** | Built-in | Age, tenure, and days-since-joining formulas |
| 🔗 **Dynamic Functions** | Built-in | `INDIRECT`, `OFFSET` |

---

## 📈 Results & Insights

After exploring the workbook, the following outputs are produced:

- ✅ **3 Fully Formula-Driven Sheets** — Students, Sales, and Employees
- 🔢 **Automatic Grading & Classification** — every row updates instantly with changed inputs
- 💰 **Accurate Sales Totals** — region-and-product-wise sums via `SUMIFS`
- 🔎 **Instant Lookups** — Employee, Student, and Salesperson details fetched in one formula
- 🔗 **Dynamic Trend Tracking** — rolling sales totals via `OFFSET`, no manual range edits needed

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | Core Excel functions covered in one compact workbook |
| 🔄 **Reusability** | Formula patterns can be reused across any dataset |
| 📚 **Educational** | Each sheet reinforces a different function family |
| 🖥️ **No Add-ins Needed** | Runs with pure native Excel functions |
| ⚡ **Lightweight** | Single-file workbook, instantly usable |
| 🧪 **Extensible** | Easy to add new sheets (Inventory, Finance, etc.) |
| 📖 **Well Organized** | Clear column headers and labeled lookup panels |
| 🛡️ **Formula Safety** | No hardcoded results — every value recalculates |

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

[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/)

> *"Every insight starts with a single formula — just like every workbook starts with a single cell."*

**🎓 Role:** Junior Excel Analyst | Spreadsheet Enthusiast \
**📍 Location:** India\
**🛠️ Skills:** Microsoft Excel · Lookup Functions · Data Analysis · Formula Building · Spreadsheet Design

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support Docs](https://support.microsoft.com/excel) — Official Excel function reference
- 🔍 [ExcelJet — Function Guides](https://exceljet.net/) — In-depth formula tutorials
- 📐 [GeeksForGeeks — Excel Functions](https://www.geeksforgeeks.org/) — Formula examples and use cases
- 🖥️ [W3Schools](https://www.w3schools.com/) — Beginner-friendly reference material
- 🧮 [Chandoo.org](https://chandoo.org/) — Excel tips and dashboard techniques
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support
- 📖 [Kaggle Learn](https://www.kaggle.com/learn) — Data analysis fundamentals

---

<div align="center">

---

*Made with ❤️ and 📊 — Last updated: 30 August, 2026*

</div>
