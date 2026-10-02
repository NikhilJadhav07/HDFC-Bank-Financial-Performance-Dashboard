# 🏦 HDFC Bank Financial Performance Dashboard

![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?logo=powerbi&logoColor=black)
![DAX](https://img.shields.io/badge/DAX-Measures-0078D4)
![Power Query](https://img.shields.io/badge/Power%20Query-ETL-217346)
![Excel](https://img.shields.io/badge/Excel-Data%20Source-217346?logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)

An interactive, multi-page **Power BI dashboard** that analyses **HDFC Bank's financial performance and banking health over five financial years (FY2020-21 to FY2024-25)**: profitability, business growth, asset quality, efficiency and capital, and cash flow.

![Executive Overview](images/executive-overview.png)

---

## 📑 Table of Contents
- [Project Overview](#-project-overview)
- [Objectives](#-objectives)
- [Dashboard Pages](#-dashboard-pages)
- [Dataset](#-dataset)
- [Tech Stack](#-tech-stack)
- [Data Preparation](#-data-preparation)
- [DAX Measures](#-dax-measures)
- [Key Insights](#-key-insights)
- [How to Open the Dashboard](#-how-to-open-the-dashboard)
- [Repository Structure](#-repository-structure)
- [About the Author](#-about-the-author)
- [Contact](#-contact)

---

## 📌 Project Overview

Bank annual data is spread across many ratios and statements, which makes it hard to judge overall health at a glance. This project turns five years of HDFC Bank's reported financials into a single dashboard that answers the questions a banking analyst or executive would ask:

- Is the bank growing, and is that growth profitable?
- Are advances growing faster than deposits, and is funding becoming tighter?
- Is asset quality holding up?
- How efficient is the bank, and how strong is its capital position?
- Where is the cash coming from and going to?

A **Financial Year slicer** and a **side navigation panel** let the viewer filter any year and move between six report pages.

## 🎯 Objectives

1. Clean and model the bank's multi-year financial data in Excel and Power Query.
2. Build reusable **DAX measures** for growth, ratios and cash-flow metrics.
3. Design a consistent, executive-friendly dashboard with KPI cards, trend charts and navigation.
4. Extract business insights from profitability, balance-sheet, asset-quality and capital metrics.

## 🖥️ Dashboard Pages

| # | Page | What it covers |
|---|------|----------------|
| 1 | **Executive Overview** | Headline KPIs (Net Profit, Net Profit Growth %, Total Advances, Total Deposits, ROE %); Net Profit trend; Advances vs Deposits; Credit-Deposit Ratio; CASA Ratio; Asset Quality (GNPA % vs Net NPA %) |
| 2 | **Revenue & Profitability** | Sales, operating profit, net profit, net profit growth and margin trends |
| 3 | **Banking Business Growth** | Total advances and deposits, their YoY growth, and credit-deposit ratio |
| 4 | **Asset Quality & Risk** | GNPA %, Net NPA %, NPA spread and provision coverage ratio |
| 5 | **Efficiency, Returns & Capital** | NIM, ROA, ROE, cost-to-income ratio and capital adequacy ratio |
| 6 | **Cash Flow & Ad-Hoc Analysis** | Operating, investing and financing cash flow and net cash flow |

Every page has the same layout: logo and title, FY slicer, left navigation buttons, KPI cards and charts.

## 📊 Dataset

- **File:** [`data/HDFC_Bank_Final_Analysis.xlsx`](data/HDFC_Bank_Final_Analysis.xlsx) (sheet `Sheet1`)
- **Grain:** one row per financial year (5 rows), 22 columns
- **Period:** FY2020-21 to FY2024-25
- **Units:** amounts in ₹ Crore; ratios in percent

<details>
<summary><b>Data dictionary (click to expand)</b></summary>

| Column | Description | Type |
|--------|-------------|------|
| `financial_year` | Financial year (e.g. FY2024-25) | Text |
| `sales` | Total revenue / sales | Number |
| `operating_profit` | Operating profit | Number |
| `net_profit` | Net profit | Number |
| `net_profit_yoy_pct` | Net profit growth, year over year (%) | Number |
| `net_profit_margin_pct` | Net profit as % of sales | Number |
| `total_advances` | Total loans and advances | Whole number |
| `total_deposits` | Total deposits | Whole number |
| `credit_deposit_ratio_pct` | Advances ÷ deposits (%) | Number |
| `casa_ratio` | Current & savings account share of deposits (%) | Number |
| `gnpa_ratio` | Gross NPA ratio (%) | Number |
| `net_npa_ratio` | Net NPA ratio (%) | Number |
| `provision_coverage_ratio` | Provision coverage ratio (%) | Number |
| `nim` | Net interest margin (%) | Number |
| `roa` | Return on assets (%) | Number |
| `roe` | Return on equity (%) | Number |
| `car` | Capital adequacy ratio (%) | Number |
| `cost_to_income_ratio` | Cost-to-income ratio (%) | Number |
| `cash_from_operating_activity` | Operating cash flow | Number |
| `cash_from_investing_activity` | Investing cash flow | Number |
| `cash_from_financing_activity` | Financing cash flow | Number |
| `net_cash_flow` | Net cash flow | Number |

</details>

> **Note:** `net_profit_yoy_pct` has no value for FY2020-21 (no prior year in the dataset) and `casa_ratio` has no value for FY2020-21.

## 🛠️ Tech Stack

| Tool | Used for |
|------|----------|
| **Microsoft Power BI Desktop** | Data modelling, visuals, navigation, report design |
| **DAX** | Measures for growth, ratios and cash-flow metrics |
| **Power Query (M)** | Importing the Excel sheet, promoting headers, setting data types |
| **Microsoft Excel** | Cleaned source dataset |

## 🧹 Data Preparation

Done in Power Query:

1. Connected to the Excel workbook and loaded `Sheet1`.
2. Promoted the first row to headers.
3. Set data types: text for `financial_year`, decimal numbers for ratios and amounts, whole numbers for advances and deposits.
4. Loaded the table to the data model as `HDFC_Bank_Final_Analysis`.

## 🧮 DAX Measures

The model contains 20+ measures, including:

| Category | Measures |
|----------|----------|
| **Growth** | Net Profit Growth %, Advances YoY Growth %, Deposits YoY Growth % |
| **Profitability** | Net Profit Margin %, Operating Profit Margin %, NIM %, ROA %, ROE % |
| **Business health** | Credit Deposit Ratio %, CASA Ratio %, Cost to Income % |
| **Asset quality** | GNPA %, Net NPA %, NPA Spread %, Provision Coverage Ratio |
| **Capital** | Capital Adequacy Ratio |
| **Cash flow** | Operating Cash Flow, Investing Cash Flow, Financing Cash Flow, Net Cash Flow |

Example: year-over-year growth driven by the selected financial year

```DAX
Advances YoY Growth % =
VAR CurrentYear = MAX ( HDFC_Bank_Final_Analysis[financial_year] )
VAR CurrentAdvances =
    CALCULATE (
        SUM ( HDFC_Bank_Final_Analysis[total_advances] ),
        HDFC_Bank_Final_Analysis[financial_year] = CurrentYear
    )
VAR PreviousAdvances =
    SWITCH (
        CurrentYear,
        "FY2021-22", CALCULATE ( SUM ( HDFC_Bank_Final_Analysis[total_advances] ), HDFC_Bank_Final_Analysis[financial_year] = "FY2020-21" ),
        "FY2022-23", CALCULATE ( SUM ( HDFC_Bank_Final_Analysis[total_advances] ), HDFC_Bank_Final_Analysis[financial_year] = "FY2021-22" ),
        "FY2023-24", CALCULATE ( SUM ( HDFC_Bank_Final_Analysis[total_advances] ), HDFC_Bank_Final_Analysis[financial_year] = "FY2022-23" ),
        "FY2024-25", CALCULATE ( SUM ( HDFC_Bank_Final_Analysis[total_advances] ), HDFC_Bank_Final_Analysis[financial_year] = "FY2023-24" )
    )
RETURN
    IF ( ISBLANK ( PreviousAdvances ), BLANK (), DIVIDE ( CurrentAdvances - PreviousAdvances, PreviousAdvances ) )
```

## 💡 Key Insights

*(All figures from the dataset; amounts in ₹ Crore.)*

**Growth and profitability**
- Net profit more than doubled from **31,833 to 70,792** (+122%), about **22% CAGR** over four years.
- Sales grew faster (+162%, from 128,552 to 336,367), so the **net profit margin fell from a peak of 27.99% (FY22) to 21.05% (FY25)**.
- FY2023-24 shows a sharp step-up (sales 170,754 → 283,649; advances +55%), which coincides with the HDFC Ltd. merger.

**Business growth and funding**
- Advances grew **2.3x** (11.3 to 26.2 lakh crore) against **2.0x** for deposits (13.4 to 27.1 lakh crore).
- The **credit-deposit ratio peaked at 104.42% in FY2023-24** (advances exceeded deposits) and eased to 96.50% in FY2024-25.
- **CASA ratio declined steadily from 48.2% (FY22) to 34.8% (FY25)**, pointing to a costlier deposit mix.

**Asset quality**
- Gross NPA improved to a low of **1.12% in FY2022-23**, then drifted up to **1.33% in FY2024-25**; Net NPA followed the same path (0.27% to 0.43%).
- Provision coverage peaked at **75.8% (FY23)** and fell to **67.86% (FY25)**.

**Efficiency, returns and capital**
- **NIM compressed from 4.10% to 3.48%**, and **ROE fell from 17.39% (FY23) to 14.56% (FY25)**.
- Cost-to-income rose from 36.3% to 40.5%.
- Capital adequacy stayed strong and stable (**18.8% to 19.6%**).

**Cash flow**
- Net cash flow was positive in all five years. In FY2024-25, operating cash flow was **+127,242** and financing cash flow **-102,478**.

## ▶️ How to Open the Dashboard

1. Install [Power BI Desktop](https://powerbi.microsoft.com/desktop/) (free, Windows).
2. Download or clone this repository:
   ```bash
   git clone https://github.com/NikhilJadhav07/hdfc-bank-financial-analysis-powerbi.git
   ```
3. Open `dashboard/Financial_Dashboard.pbix`.
4. If Power BI asks about the data source, go to **Home → Transform data → Data source settings → Change Source** and select `data/HDFC_Bank_Final_Analysis.xlsx`, then **Close & Apply**.
5. Use the **Financial Year** slicer and the left navigation buttons to explore.

## 📁 Repository Structure

```
hdfc-bank-financial-analysis-powerbi/
├── dashboard/
│   └── Financial_Dashboard.pbix
├── data/
│   └── HDFC_Bank_Final_Analysis.xlsx
├── images/
│   └── executive-overview.png
├── README.md
├── LICENSE
└── .gitignore
```

## 👤 About the Author

**Nikhil Jadhav**: B.E. in Artificial Intelligence & Data Science (2026), Guru Gobind Singh College of Engineering and Research Center, Nashik. Currently a **Data Analyst Intern at Code B, Nashik** (since July 2026), and looking for software and data roles.

## 📬 Contact

| | |
|---|---|
| 📧 Email | [nikhiljadhav42135@gmail.com](mailto:nikhiljadhav42135@gmail.com) |
| 📱 Phone | +91 7588404338 |
| 💼 LinkedIn | [linkedin.com/in/nikhil-jadhav-347520423](https://www.linkedin.com/in/nikhil-jadhav-347520423) |
| 🐙 GitHub | [github.com/NikhilJadhav07](https://github.com/NikhilJadhav07) |

---

⭐ If you found this project useful, consider giving it a star.

*Disclaimer: This project is for learning and portfolio purposes only and is not investment advice. Figures come from the dataset in this repository.*
