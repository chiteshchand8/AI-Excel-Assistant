# Online AI Excel Assistant — V3.6.5 Beta

A Windows-based Python assistant that translates natural-language requests into structured Excel operations using Google Gemini and Microsoft Excel COM automation. It creates native PivotTables, charts, KPI summaries and interactive slicer-driven dashboards inside the selected workbook.

## Features
- Dataset profiling, descriptive summaries and KPI calculations
- Filtering, sorting, blank-row and duplicate removal
- Excel formulas and data validation
- Native PivotTables with mixed aggregations and six chart types
- Executive dashboards with up to four KPIs, four charts and three connected slicers
- Plan preview and confirmation, backups, staged writes, replacement retries and manual recovery

## Requirements
Windows, installed desktop Microsoft Excel (tested on Office 2021), Python 3.10 (tested on 3.10.11), internet access and a Gemini API key. Excel Online and macOS are not supported by this COM implementation. Model access and quotas depend on Google; the configured model IDs may need updating over time.

## Setup (Windows CMD)
Open CMD in the extracted project directory.

```cmd
python3.10 -m pip install -r requirements.txt
copy .env.example .env
notepad .env
```
Replace the placeholder with your own Gemini API key and save. Never commit `.env`.
If your Python installation uses `python` instead of `python3.10`, substitute that command.

Place exactly one trusted `.xlsx` file in `workbooks/`, with unique nonempty text headers in row 1 starting at A1. Keep an independent backup. Close the workbook in Excel before running:

```cmd
python3.10 assistant.py
```

Review the JSON plan carefully before typing `y`. Type `exit` to quit. Successful write operations update the selected workbook itself. Backups are stored locally in `backups/`.

## Example requests
Adapt sheet and column names to your workbook.
- Inspect Sales_Data and report missing values and duplicates.
- Calculate total Sales_Amount and Profit from Sales_Data in a separate KPI sheet.
- Filter Sales_Data where Sales_Amount is greater than 100000 into a separate sheet.
- Create an executive dashboard from Sales_Data with sales, profit, order count and quantity KPIs, meaningful category charts and slicers for Product_Category, State and Payment_Mode.

## Architecture
Natural-language request + workbook profile → Gemini JSON plan → normalization and confirmation → local pandas / Excel COM execution on a staged copy → save and reopen check → original-file comparison and replacement.

## Privacy and safety
The request and workbook metadata (including sheet and column names, types, row counts, missing counts, duplicate counts and unique counts) are sent to Google Gemini. The engine does not intentionally send raw worksheet rows in its profiling prompt. Do not put confidential information in requests or metadata. Use trusted workbooks only. Formula filtering is not a complete security sandbox. Cleaning, sorting, formulas and validation can modify existing data or formatting; approve only intended changes. Do not run concurrent writers.

If replacement fails, the engine attempts to retain the completed staged workbook in `backups/RECOVERY_*.xlsx` (or staging if moving fails). Inspect the recovered workbook and compare it with the current original before manually replacing anything.

## Limitations
Beta portfolio software, not production-certified. Tested with Sales and HR datasets on Windows. No automatic monthly grouping or top-N support; dashboard charts use one measure and existing grouped fields. Dashboard categories are limited to 50 distinct values and some identifier names are rejected. Filtered outputs contain values rather than original formulas, and dates may become text. KPI count counts non-null entries, not distinct records. Dashboard mean alias can mismatch GETPIVOTDATA captions; use average. Chart formatting and color consistency have minor issues. Cleaning's reported removal count may exclude blank rows. Formula restrictions and plan validation are incomplete. Backups and staging consume disk space. Concurrent changes cannot be ruled out by a non-atomic comparison and replacement. See TEST_REPORT.md.

## License
No open-source license has been selected yet. The owner should choose one before granting reuse rights.
---

## Dashboard Demonstrations

### Executive Sales Dashboard

![Sales Dashboard](screenshots/sales_dashboard.png)

**Key highlights:**

- Total Sales: 5,005,259
- Total Profit: 893,913.62
- Total Orders: 100
- Total Quantity: 489
- Four native Excel charts
- Three interactive slicers
- Dynamic KPI calculations

### HR Executive Dashboard

![HR Dashboard](screenshots/hr_dashboard.png)

**Key highlights:**

- Total Employees: 200
- Total Monthly Salary: 12,378,000
- Average Monthly Salary: 61,890
- Average Performance Rating: 2.94
- Department-wise employee and salary analysis
- Interactive department, location and employment status slicers

## Multi-Dataset Compatibility

Both dashboards were automatically generated using the same Python engine, without requiring dataset-specific code modifications.

The project demonstrates the integration of Generative AI, Python, data analytics and native Microsoft Excel automation.

---
