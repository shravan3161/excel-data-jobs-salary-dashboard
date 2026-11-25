# Interactive Data Jobs Salary Dashboard (Microsoft Excel)

An interactive, formula-driven **Microsoft Excel 365** salary dashboard analyzing **32,672 global data-role job postings** across **10 job titles**, **111 countries**, **5 schedule types**, and **593 hiring platforms**. Built using Excel Tables, **Dynamic Array formulas** (`UNIQUE`, `SORT`, `FILTER`, `XLOOKUP`, `ANCHORARRAY`), **Named Ranges**, **Data Validation**, conditional `NA()` chart highlighting, and a geographic **Filled Map** chart.

---

## Table of Contents
- [Executive Summary](#executive-summary)
- [Key Questions Answered](#key-questions-answered)
- [Dashboard Architecture](#dashboard-architecture)
- [Workbook Structure](#workbook-structure)
- [Excel Formulas & Techniques](#excel-formulas--techniques)
- [Dashboard Visuals & KPIs](#dashboard-visuals--kpis)
- [Sample Insights (US Full-Time Baseline)](#sample-insights-us-full-time-baseline)
- [Technical Challenges & Future Improvements](#technical-challenges--future-improvements)
- [Repository Structure](#repository-structure)
- [How to Use](#how-to-use)
- [Author & Acknowledgments](#author--acknowledgments)

---

## Executive Summary

Navigating compensation across data careers requires comparing pay across titles, geographies, and contract types simultaneously. Instead of relying on static summary tables, this workbook creates a live **Salary Calculator & Dashboard** powered entirely by native Excel formulas and dynamic arrays.

Selecting any **Job Title**, **Country**, or **Job Type** from the dropdown controls automatically recalculates all backend sheets, re-sorts rankings, highlights the active selection inside the bar charts, updates the global salary map, and refreshes the KPI summary cards.

---

## Key Questions Answered

1. **Role Pay Hierarchy:** How do median salaries compare across data roles (e.g., Data Analyst vs. Data Engineer vs. Data Scientist vs. Senior & ML roles) for a given country and schedule type?
2. **Geographic Compensation Disparities:** Which countries offer the highest median pay for a selected role and employment type?
3. **Employment Type Impact:** How does median compensation vary across Full-time, Part-time, Contractor, Internship, and Temp work?
4. **Hiring Volume & Platform Dominance:** How many matching postings exist, and which job board (e.g., Indeed, Ladders, LinkedIn) lists the most opportunities for the selected criteria?

---

## Dashboard Architecture

```text
+----------------------------------+
|     Data Sheet (`jobs` Table)    |  32,672 job postings x 16 attributes
+----------------+-----------------+
                 |
                 v
+----------------------------------+
|   Named Ranges & Dropdown Lists  |  `title`, `country`, `type`
|     (`Data_Validation` Sheet)    |  Sorted UNIQUE / FILTER lists
+----------------+-----------------+
                 |
                 v
+----------------------------------+
|     Hidden Calculation Engine    |
|  - `Title`   (Role medians)      |  MEDIAN(IF(...)), SORT(FILTER(...)),
|  - `Country` (Geo medians)       |  COUNTIFS, XLOOKUP, SUBSTITUTE,
|  - `Type`    (Schedule medians)  |  and IF(..., NA()) chart series
|  - `Platform`(Job board counts)  |
+----------------+-----------------+
                 |
                 v
+----------------------------------+
|    `Salary_Calculator` Dashboard |  - 3 Data Validation Dropdowns
|         (Visible UI Sheet)       |  - Stacked Bar Charts w/ Highlighting
|                                  |  - Regional Map Chart
|                                  |  - 3 Formula-Linked KPI Cards
+----------------------------------+
```

---

## Workbook Structure

Only **`Salary_Calculator`** is visible to end users; all data and calculation sheets are hidden to keep the interface clean.

| Worksheet | Visibility | Purpose |
| :--- | :--- | :--- |
| **`Salary_Calculator`** | Visible | Interactive dashboard UI with 3 dropdown controls (`C4`, `G4`, `Q4`), 3 charts, and 3 KPI cards. |
| **`Data`** | Hidden | Raw dataset structured as Excel Table `jobs` (`A1:P32673`: 32,672 records across 16 columns). |
| **`Data_Validation`** | Hidden | Generates sorted `UNIQUE` job titles (by count), sorted `UNIQUE` countries, filtered `UNIQUE` schedule types, and active `count` lookup (`M2`). |
| **`Title`** | Hidden | Computes multi-criteria `MEDIAN(IF(...))` salaries per job title, sorts results, splits active vs. inactive series via `NA()`, and computes `median_salary` (`I2`). |
| **`Country`** | Hidden | Computes multi-criteria `MEDIAN(IF(...))` salaries across 111 countries and filters out `#NUM!` entries (`ISNUMBER`) for the Map chart. |
| **`Type`** | Hidden | Computes `MEDIAN(IF(...))` salaries across the 5 schedule types, sorts ascending, and splits active vs. inactive series via `NA()`. |
| **`Platform`** | Hidden | Ranks 593 hiring platforms using `COUNTIFS`, sorts descending, and strips `"via"` prefixes (`G2`) for the `platform` KPI card. |

---

## Excel Formulas & Techniques

* **Multi-Criteria Median Array Logic:**
  Calculates median annual salary while filtering out zero/blank salary records and matching job title, country, and schedule substring:
  ```excel
  =MEDIAN(
    IF(
      (jobs[job_title_short]=A2)*
      (jobs[salary_year_avg]<>0)*
      (jobs[job_country]=country)*
      (ISNUMBER(SEARCH(type,jobs[job_schedule_type]))),
      jobs[salary_year_avg]
    )
  )
  ```
* **Dynamic Arrays (`UNIQUE`, `SORT`, `FILTER`, `#`):**
  * Filters out countries with no matching salary records (`#NUM!`) and ranks them:
    ```excel
    =SORT(FILTER(A2:B112, ISNUMBER(B2:B112)), 2, -1)
    ```
  * Extracts single-category schedule types (excluding combined `"and"` entries and zeros) for dropdown validation:
    ```excel
    =FILTER(J2#, NOT(ISNUMBER(SEARCH("and", J2#)))*(J2#<>0))
    ```
* **Conditional Chart Highlighting via `NA()`:**
  Splits chart values into two stacked series so the user's selected Job Title or Job Type is emphasized in a darker shade while non-selected bars remain muted:
  ```excel
  =IF($D2<>title, $E2, NA())   ' Unselected bar series
  =IF($D2=title, $E2, NA())    ' Highlighted bar series
  ```
* **Named Ranges & Lookups:**
  Named inputs (`title`, `country`, `type`) and outputs (`median_salary`, `count`, `platform`) connect formulas (`XLOOKUP`, `COUNTIFS`, `SUBSTITUTE`) directly to the dashboard's Data Validation dropdowns and KPI text boxes.

---

## Dashboard Visuals & KPIs

* **Median Salary by Job Title (Stacked Bar Chart):** Compares all 10 data roles sorted by median compensation with the selected role highlighted.
* **Global Median Salary Distribution (Filled Map Chart):** Visualizes country-level median salaries for the active role and job type.
* **Median Salary by Schedule Type (Stacked Bar Chart):** Benchmarks pay across `Full-time`, `Part-time`, `Contractor`, `Internship`, and `Temp work` with the active selection highlighted.
* **KPI Summary Cards:** Dynamically display **Median Salary** (`Title!$I$2`), **Top Job Platform** (`Platform!$G$2`), and **Job Count** (`Data_Validation!$M$2`).

---

## Sample Insights (US Full-Time Baseline)

For **United States + Full-time** postings in the dataset:
* **Data Scientist:** **$130,000** median annual salary across **5,638** postings (top platform: **Indeed** with 961 listings, followed by Ladders with 886 and LinkedIn with 878).
* **Senior & ML Engineering Roles Lead Pay:** **Senior Data Scientist** ($155,000), **Senior Data Engineer** ($150,000), and **Machine Learning Engineer** ($150,000) command the highest median salaries, compared to **Data Analyst** ($90,000) and **Business Analyst** ($90,000).
* **Schedule Type Variance (US Data Scientist):** Full-time leads at **$130,000**, followed by Part-time (**$128,583**), Contractor (**$122,500**), Internship (**$97,013**), and Temp work (**$87,500**).

---

## Technical Challenges & Future Improvements

### 1. Formula Readability & Optimization
* **Challenge:** Repeating nested `MEDIAN(IF(...))` array expressions across 100+ country rows makes debugging harder.
* **Improvement:** Refactor formulas using Excel 365's `LET()` and `FILTER()` functions for cleaner variable naming and single-pass filtering:
  ```excel
  =LET(
    mask, (jobs[job_country]=A2)*(jobs[salary_year_avg]<>0)*(jobs[job_title_short]=title)*(ISNUMBER(SEARCH(type,jobs[job_schedule_type]))),
    vals, FILTER(jobs[salary_year_avg], mask),
    IFERROR(MEDIAN(vals), NA())
  )
  ```

### 2. Map Chart Recalculation Lag
* **Challenge:** Updating dropdowns triggers simultaneous array recalculations across 111 countries and Bing geocoding in the Map chart.
* **Improvement:** Pre-aggregate metrics in Power Pivot / DAX or toggle Manual Calculation mode during layout edits.

---

## Repository Structure

```text
excel-data-jobs-salary-dashboard/
├── Salary_DashBoard.xlsx   # Interactive Excel 365 Dashboard Workbook
├── .gitignore              # Ignores Excel lock files (~$*.xlsx) and OS artifacts
└── README.md               # Project documentation, architecture, and formulas
```

---

## How to Use

1. **Download or clone** this repository (GitHub's web preview does not execute Excel formulas or render Map charts; open the file locally in **Microsoft Excel 365**).
   ```bash
   git clone https://github.com/shravan3161/excel-data-jobs-salary-dashboard.git
   ```
2. Open **`Salary_DashBoard.xlsx`** in Microsoft Excel 365.
3. On the **`Salary_Calculator`** sheet, use the dropdown menus for **Job Title** (`C4`), **Country** (`G4`), and **Type** (`Q4`).
4. All charts, rankings, and KPI cards update automatically.

---
