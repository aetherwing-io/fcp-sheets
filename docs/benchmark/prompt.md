# Sheets Showdown v2 — Executive Portfolio Review (3-Phase)

## Overview

Build a private equity portfolio review workbook across three phases:
1. **Create** — Full build from scratch (5 sheets, charts, cross-sheet formulas, styling)
2. **Update** — Cold-open the saved file and apply CFO feedback (fix, change, add)
3. **Polish** — Final LP-meeting-ready refinements

Each phase simulates a separate session — no prior context carries over.

---

## PHASE 1: Build

### Task: Meridian Capital Partners — Q4 2025 Portfolio Review

Build a 5-sheet workbook for the quarterly LP meeting.

---

#### Sheet 1: "Executive Summary"

Merged title A1:F1: "Meridian Capital Partners" (20pt, bold, #1a1a2e fill, white text)
Merged subtitle A2:F2: "Fund III — Q4 2025 Portfolio Review" (13pt, italic, #1a1a2e fill, #B0B0B0 text)

**Fund Overview (rows 4-11):**
- A4: "Fund Overview" (14pt, bold, underline)
- Data in A5:B11, two columns (Metric | Value):

| Metric | Value | Format |
|---|---|---|
| Fund Size | 500000000 | $#,##0,,"M" |
| Vintage Year | 2021 | 0 |
| Investment Period | "2021–2024" | text |
| Invested Capital | 385000000 | $#,##0,,"M" |
| Realized Value | 142000000 | $#,##0,,"M" |
| Unrealized Value | 498000000 | $#,##0,,"M" |
| Total Value | =B9+B10 | $#,##0,,"M" |

- A5:A11 bold (metric labels)
- Borders: thin inner gridlines + medium outline on A5:B11

**Performance Metrics (rows 13-19):**
- A13: "Performance" (14pt, bold, underline)
- Data in A14:B19:

| Metric | Value | Format |
|---|---|---|
| Net IRR | 0.218 | 0.0% |
| Gross IRR | 0.267 | 0.0% |
| Net MOIC | 1.66 | 0.00x (custom: `0.00"x"`) |
| DPI (Distributions/Paid-In) | =B9/B8 | 0.00x |
| RVPI (Residual/Paid-In) | =B10/B8 | 0.00x |
| TVPI (Total/Paid-In) | =B11/B8 | 0.00x |

- Conditional formatting on B14 (Net IRR): green fill (#C6EFCE) if > 0.15, yellow fill (#FFF2CC) if 0.08–0.15, red fill (#FFC7CE) if < 0.08
- Borders: thin inner + medium outline on A14:B19

**Portfolio Snapshot (rows 21-23):**
- A21: "Portfolio Snapshot" (14pt, bold, underline)
- A22: "Active Companies", B22: 8, format #,##0
- A23: "Fully Exited", B23: 2, format #,##0

- Column widths: A=28, B=20
- Protect sheet (no password)
- Page setup: landscape, fit to 1 page wide
- Named range: `FundSize` → B5
- Named range: `InvestedCapital` → B8
- Named range: `TotalValue` → B11

---

#### Sheet 2: "Fund Performance"

Merged title A1:H1: "Fund Performance & Benchmarks" (16pt, bold, #1a1a2e fill, white text)

Header row 2: bold, #4472C4 fill, white text, frozen at A3

| Year | Contributions | Distributions | Net Cash Flow | Cumulative Net | Net IRR | MOIC | Benchmark IRR |
|---|---|---|---|---|---|---|---|
| 2021 | -125000000 | 0 | =B3+C3 | =D3 | -0.08 | 0.92 | -0.05 |
| 2022 | -140000000 | 12000000 | =B4+C4 | =E3+D4 | -0.03 | 0.95 | 0.02 |
| 2023 | -85000000 | 58000000 | =B5+C5 | =E4+D5 | 0.11 | 1.18 | 0.09 |
| 2024 | -35000000 | 72000000 | =B6+C6 | =E5+D6 | 0.19 | 1.48 | 0.14 |
| 2025 | 0 | 0 | =B7+C7 | =E6+D7 | 0.218 | 1.66 | 0.16 |
| **Total** | =SUM(B3:B7) | =SUM(C3:C7) | =SUM(D3:D7) | | | | |

Data starts row 3, Total in row 8. (Contributions are negative, distributions positive.)

Requirements:
- Contributions/Distributions/Net Cash Flow/Cumulative in `$#,##0,,"M"` (millions display — custom format `$#,##0,,\\"M\\"` or just `$#,##0,,"M"`)
- IRR/Benchmark in `0.0%`; MOIC in `0.00"x"`
- Conditional formatting on F3:F7 (Net IRR): green text (#006100) if > 0.15, red text (#9C0006) if < 0
- Conditional formatting on G3:G7 (MOIC): green fill (#C6EFCE) if > 1.5, yellow fill (#FFF2CC) if 1.0–1.5, red fill (#FFC7CE) if < 1.0
- Row labels (A column) bold; Total row (row 8) bold entire row
- Column widths: A=10, B-E=18, F-H=14
- **Chart 1:** Stacked bar chart titled "Annual Cash Flows" — Contributions (B3:B7) and Distributions (C3:C7) stacked, categories years (A3:A7), placed near A11, size 700x350
- **Chart 2:** Line chart titled "IRR vs Benchmark" — Net IRR (F3:F7) and Benchmark IRR (H3:H7), categories years (A3:A7), placed near A28, size 700x300
- Page setup: landscape

---

#### Sheet 3: "Holdings"

Merged title A1:K1: "Portfolio Holdings Detail" (16pt, bold, #1a1a2e fill, white text)

Header row 2: bold, #1a1a2e fill, white text, frozen at A3

| Company | Sector | Investment Date | Cost Basis | Fair Value | MOIC | Revenue (LTM) | EBITDA (LTM) | EBITDA Margin | Rev Growth YoY | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Helios Energy | Clean Energy | 2021-03-15 | 65000000 | 142000000 | =E3/D3 | 89000000 | 22000000 | =H3/G3 | 0.42 | Active |
| NovaBio Sciences | Healthcare | 2021-06-22 | 45000000 | 38000000 | =E4/D4 | 31000000 | -2000000 | =H4/G4 | 0.08 | Active |
| Axion Robotics | Industrials | 2021-11-01 | 55000000 | 98000000 | =E5/D5 | 72000000 | 18000000 | =H5/G5 | 0.35 | Active |
| Prism Analytics | Software | 2022-02-14 | 40000000 | 88000000 | =E6/D6 | 54000000 | 16000000 | =H6/G6 | 0.55 | Active |
| Stratos Logistics | Logistics | 2022-07-30 | 35000000 | 42000000 | =E7/D7 | 48000000 | 8000000 | =H7/G7 | 0.18 | Active |
| Cirrus Cloud | Software | 2022-09-18 | 50000000 | 105000000 | =E8/D8 | 68000000 | 21000000 | =H8/G8 | 0.48 | Active |
| Verdant Agriculture | AgTech | 2023-01-10 | 30000000 | 27000000 | =E9/D9 | 19000000 | 1000000 | =H9/G9 | 0.12 | Active |
| Quantum Mesh | Telecom | 2023-05-20 | 25000000 | 36000000 | =E10/D10 | 22000000 | 5000000 | =H10/G10 | 0.28 | Active |
| Bolt Payments | Fintech | 2021-08-05 | 20000000 | 0 | =E11/D11 | 0 | 0 | 0 | 0 | Written Off |
| Apex Dynamics | SaaS | 2022-04-12 | 20000000 | 62000000 | =E12/D12 | 41000000 | 12000000 | =H12/G12 | 0.62 | Exited |
| **Totals** | | | =SUM(D3:D12) | =SUM(E3:E12) | =E13/D13 | =SUM(G3:G12) | =SUM(H3:H12) | =H13/G13 | | |

Data starts row 3, Totals in row 13.

Requirements:
- Cost Basis, Fair Value, Revenue, EBITDA in `$#,##0,,"M"`
- MOIC in `0.00"x"`; EBITDA Margin in `0.0%`; Rev Growth in `0.0%`
- Investment Date in `yyyy-mm-dd`
- Data validation on Status column (K3:K12): list ["Active", "Exited", "Written Off"]
- Conditional formatting on F3:F12 (MOIC): color scale — min red (#F8696B), mid yellow (#FFEB84), max green (#63BE7B)
- Conditional formatting on J3:J12 (Rev Growth): data bars, green (#70AD47)
- Bold row labels (A column); bold Totals row (row 13)
- Alternating row fill: even rows #F2F2F2, odd rows white (rows 3-12)
- Column widths: A=18, B=14, C=16, D-E=16, F=10, G-H=16, I=14, J=14, K=12
- **Chart:** Bubble chart titled "Portfolio Map" — X=Revenue (G3:G10, active only), Y=EBITDA Margin (I3:I10), Bubble size=Fair Value (E3:E10). Place near A16, size 800x400. (If bubble chart not supported, use scatter with data labels.)
- Filter enabled on row 2 (A2:K2)
- Page setup: landscape

---

#### Sheet 4: "Cash Flows"

Merged title A1:F1: "Quarterly Cash Flow Detail" (16pt, bold, #1a1a2e fill, white text)

Header row 2: bold, #70AD47 fill, white text, frozen at A3

| Quarter | Contributions | Distributions | Net Cash Flow | Cumulative Cash Flow | NAV |
|---|---|---|---|---|---|
| Q1 2021 | -50000000 | 0 | =B3+C3 | =D3 | 48000000 |
| Q2 2021 | -75000000 | 0 | =B4+C4 | =E3+D4 | 118000000 |
| Q3 2021 | 0 | 0 | =B5+C5 | =E4+D5 | 122000000 |
| Q4 2021 | 0 | 0 | =B6+C6 | =E5+D6 | 115000000 |
| Q1 2022 | -60000000 | 0 | =B7+C7 | =E6+D7 | 168000000 |
| Q2 2022 | -45000000 | 0 | =B8+C8 | =E7+D8 | 210000000 |
| Q3 2022 | -35000000 | 12000000 | =B9+C9 | =E8+D9 | 235000000 |
| Q4 2022 | 0 | 0 | =B10+C10 | =E9+D10 | 228000000 |
| Q1 2023 | -40000000 | 18000000 | =B11+C11 | =E10+D11 | 278000000 |
| Q2 2023 | -15000000 | 0 | =B12+C12 | =E11+D12 | 295000000 |
| Q3 2023 | -30000000 | 20000000 | =B13+C13 | =E12+D13 | 340000000 |
| Q4 2023 | 0 | 20000000 | =B14+C14 | =E13+D14 | 365000000 |
| Q1 2024 | -20000000 | 25000000 | =B15+C15 | =E14+D15 | 410000000 |
| Q2 2024 | -15000000 | 0 | =B16+C16 | =E15+D16 | 438000000 |
| Q3 2024 | 0 | 27000000 | =B17+C17 | =E16+D17 | 485000000 |
| Q4 2024 | 0 | 20000000 | =B18+C18 | =E17+D18 | 520000000 |
| Q1 2025 | 0 | 0 | =B19+C19 | =E18+D19 | 548000000 |
| Q2 2025 | 0 | 0 | =B20+C20 | =E19+D20 | 575000000 |
| Q3 2025 | 0 | 0 | =B21+C21 | =E20+D21 | 610000000 |
| Q4 2025 | 0 | 0 | =B22+C22 | =E21+D22 | 640000000 |
| **Total** | =SUM(B3:B22) | =SUM(C3:C22) | =SUM(D3:D22) | | |

Data starts row 3, Total in row 23.

Requirements:
- All dollar values in `$#,##0,,"M"`
- Bold row labels (A column); bold Total row
- Conditional formatting on D3:D22 (Net Cash Flow): green fill (#C6EFCE) if > 0, red fill (#FFC7CE) if < 0
- Column widths: A=12, B-F=18
- **Chart 1:** Area chart titled "NAV Growth" — NAV data (F3:F22), categories quarters (A3:A22), placed near A26, size 700x300, blue fill (#4472C4) with 50% transparency
- **Chart 2:** Clustered bar chart titled "Contributions & Distributions" — Contributions (B3:B22) and Distributions (C3:C22), categories quarters (A3:A22), placed near A44, size 700x300
- Named range: `LatestNAV` → F22
- Page setup: landscape

---

#### Sheet 5: "Scenario Analysis"

Merged title A1:E1: "Portfolio Scenario Analysis" (16pt, bold, #1a1a2e fill, white text)

**Section 1 — Assumptions (rows 3-9):**
- A3: "Assumptions" (14pt, bold, underline)
- Header row 4: bold, #4472C4 fill, white text

| Assumption | Base | Bull | Bear |
|---|---|---|---|
| Revenue Growth | 0.25 | 0.40 | 0.10 |
| EBITDA Margin Expansion | 0.02 | 0.05 | -0.03 |
| Multiple Expansion | 0 | 0.03 | -0.05 |
| Exit Timeline (years) | 3 | 2 | 5 |
| Loss Rate | 0.05 | 0.02 | 0.15 |

Data rows 5-9.

**Section 2 — Projected Outcomes (rows 11-18):**
- A11: "Projected Outcomes" (14pt, bold, underline)
- Header row 12: bold, #70AD47 fill, white text

| Metric | Base | Bull | Bear |
|---|---|---|---|
| Portfolio Revenue | =Holdings!G13*(1+B5) | =Holdings!G13*(1+C5) | =Holdings!G13*(1+D5) |
| Portfolio EBITDA | =B13*(Holdings!I13+B6) | =C13*(Holdings!I13+C6) | =D13*(Holdings!I13+D6) |
| Implied EV (12x EBITDA) | =B14*12 | =C14*12 | =D14*12 |
| Loss Adjustment | =B15*(-B9) | =C15*(-C9) | =D15*(-D9) |
| Net Portfolio Value | =B15+B16 | =C15+C16 | =D15+D16 |
| Implied MOIC | =B17/'Executive Summary'!B8 | =C17/'Executive Summary'!B8 | =D17/'Executive Summary'!B8 |

Data rows 13-18.

Requirements:
- Revenue, EBITDA, EV, Loss Adjustment, Net Value in `$#,##0,,"M"`
- MOIC in `0.00"x"`; percentages in `0.0%`; years in `0`
- Conditional formatting on B18:D18 (Implied MOIC): green fill if > 2.0, yellow if 1.0–2.0, red if < 1.0
- Borders: medium outline around assumptions table (A4:E9), medium outline around outcomes table (A12:E18)
- Column widths: A=24, B-D=18, E=2 (spacer)
- **Chart:** Clustered bar chart titled "Scenario MOIC Comparison" — MOIC data (B18:D18), categories ["Base", "Bull", "Bear"], placed near A21, size 500x300
- Page setup: landscape

---

### Global Requirements (Phase 1)

- All sheets: landscape page setup
- No hardcoded values where formulas are specified
- The default "Sheet" sheet (if created) must be removed — only the 5 named sheets should exist

---

## PHASE 2: Update (CFO Feedback)

**Context for agent:** You are opening an existing spreadsheet for the first time. You have no prior context about how it was built. The CFO has reviewed the Q4 Portfolio Review and has the following feedback.

**File to open:**
- FCP: `/Users/scottmeyer/projects/fcp/test/showdown/v2-fcp-result.xlsx`
- Raw: `/Users/scottmeyer/projects/fcp/test/showdown/v2-raw-result.xlsx`

### FIX — Cash Flow Cumulative Error
"The Cumulative Cash Flow column on the Cash Flows sheet doesn't account for NAV changes — it only tracks cash in/out. Add a new column G called 'Total Return' that calculates Cumulative Cash Flow + NAV for each quarter. Format as `$#,##0,,"M"`. Also, the Total row should show the final Total Return value (=G22), not a SUM."

### CHANGE — Holdings Table Restructure
"On the Holdings sheet:
1. Rename the sheet tab from 'Holdings' to 'Portfolio Detail'
2. Add a new column after Status called 'Investment Thesis' (column L) — this should be a free-text column, populate it with one-line thesis for each company:
   - Helios Energy: 'Grid-scale battery storage leader'
   - NovaBio Sciences: 'Novel drug delivery platform'
   - Axion Robotics: 'Warehouse automation at scale'
   - Prism Analytics: 'AI-native business intelligence'
   - Stratos Logistics: 'Last-mile delivery optimization'
   - Cirrus Cloud: 'Multi-cloud infrastructure orchestration'
   - Verdant Agriculture: 'Precision agriculture IoT'
   - Quantum Mesh: 'Private 5G network solutions'
   - Bolt Payments: 'Real-time cross-border payments'
   - Apex Dynamics: 'Vertical SaaS for construction'
3. Add a new column M called 'Watchlist' with data validation: list ['Green', 'Yellow', 'Red']. Set values: Helios=Green, NovaBio=Red, Axion=Green, Prism=Green, Stratos=Yellow, Cirrus=Green, Verdant=Yellow, Quantum=Green, Bolt=Red, Apex=Green.
4. Conditional formatting on M3:M12: fill green (#C6EFCE) if 'Green', yellow (#FFF2CC) if 'Yellow', red (#FFC7CE) if 'Red'"

### ADD — Sector Allocation Sheet
"Add a new sheet called 'Sector Allocation' after the Scenario Analysis sheet. It should contain:

**Allocation Table (rows 1-10):**
- Merged title A1:E1: 'Sector Allocation Analysis' (16pt, bold, #1a1a2e fill, white text)
- Header row 2: bold, #9673A6 fill (purple), white text

| Sector | # Companies | Total Cost Basis | Total Fair Value | Sector MOIC |
|---|---|---|---|---|
| Software | 2 | =SUM of software companies' cost | =SUM of software fair values | =D/C |
| Clean Energy | 1 | (from Holdings) | (from Holdings) | =D/C |
| Healthcare | 1 | ... | ... | =D/C |
| Industrials | 1 | ... | ... | =D/C |
| Logistics | 1 | ... | ... | =D/C |
| AgTech | 1 | ... | ... | =D/C |
| Telecom | 1 | ... | ... | =D/C |
| Fintech | 1 | ... | ... | =D/C |

NOTE: The cost basis and fair value should be hardcoded aggregates from the Holdings sheet (since SUMIF cross-sheet is fragile). Calculate them manually from the Holdings data.

- Amounts in `$#,##0,,"M"`; MOIC in `0.00"x"`
- Conditional formatting on E3:E10 (MOIC): color scale min-red, mid-yellow, max-green
- Doughnut chart titled 'Sector Allocation by Fair Value' — Fair Value (D3:D10), categories (A3:A10), near A13, size 600x400
- Column widths: A=16, B=14, C-D=18, E=14"

---

## PHASE 3: Polish (LP Meeting Prep)

**Context for agent:** The workbook is nearly final. Apply these finishing touches before the LP meeting.

**File to open:**
- Same file as modified in Phase 2

### Polish Items

1. **Table of Contents:** On the Executive Summary sheet, add a "Navigation" section starting at row 26:
   - A26: "Quick Navigation" (14pt, bold, underline)
   - A27: "Fund Performance" — hyperlink to 'Fund Performance'!A1
   - A28: "Portfolio Detail" — hyperlink to 'Portfolio Detail'!A1
   - A29: "Cash Flows" — hyperlink to 'Cash Flows'!A1
   - A30: "Scenario Analysis" — hyperlink to 'Scenario Analysis'!A1
   - A31: "Sector Allocation" — hyperlink to 'Sector Allocation'!A1

2. **Header Consistency:** Ensure ALL sheets have the same title bar style: #1a1a2e fill, white text, bold. Verify the new Sector Allocation sheet matches.

3. **Data Bars:** On the Fund Performance sheet, add data bars (blue, #4472C4) to the Distributions column (C3:C7).

4. **Footer:** On every sheet, set the page footer to: Left="Meridian Capital Partners", Center="Confidential", Right="Q4 2025"

5. **Print Titles:** On the Holdings/Portfolio Detail sheet and Cash Flows sheet, set rows 1-2 as print title rows (repeat at top of each page).

6. **Final Named Range:** Add named range `LatestMOIC` pointing to the Net MOIC cell on Executive Summary (B16).

---

## Evaluation Criteria

Each phase is audited separately. Total score across all phases.

### Phase 1 Audit (Build)
1. Sheet structure (5 sheets, correct names, no default "Sheet")
2. Data values (all hardcoded numbers match)
3. Formulas (SUM, cross-sheet, calculated MOIC/margins/cumulative)
4. Number formats ($M display, percentages, MOIC, dates)
5. Styling (title bars, header fills, bold, frozen panes, merges)
6. Conditional formatting (IRR thresholds, MOIC color scales, data bars, cash flow coloring)
7. Charts (6 charts: stacked bar, line, bubble/scatter, area, clustered bar, clustered bar)
8. Named ranges (FundSize, InvestedCapital, TotalValue, LatestNAV)
9. Data validation (Holdings Status column)
10. Borders (Executive Summary sections, Scenario Analysis tables)
11. Alternating row fills (Holdings)
12. Filter (Holdings)
13. Protection (Executive Summary)
14. Column widths
15. Page setup (landscape on all sheets)

### Phase 2 Audit (Update)
1. Cash Flows Total Return column (G) added with correct formulas
2. Holdings renamed to Portfolio Detail
3. Investment Thesis column (L) with correct text
4. Watchlist column (M) with data validation and conditional formatting
5. Sector Allocation sheet with table, formulas, chart
6. All existing data/formulas preserved (no regression)

### Phase 3 Audit (Polish)
1. Navigation hyperlinks on Executive Summary
2. Header consistency across all sheets
3. Data bars on Fund Performance distributions
4. Page footers on all sheets
5. Print title rows on Portfolio Detail and Cash Flows
6. LatestMOIC named range
