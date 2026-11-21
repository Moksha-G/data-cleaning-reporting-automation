# Data Cleaning & Reporting Automation

This project automates the cleanup demonstrated in the tutorial using the supplied `Data Cleaning Start.xlsx` workbook.

## Files

- `Data Cleaning Start.xlsx` — unchanged source workbook.
- `clean_data.py` — repeatable cleaning and reporting workflow.
- `Cleaned_Data_Report_Final.xlsx` — formatted report with a summary dashboard, cleaned data, and the raw source fields for comparison.
- `cleaned_data.csv` — cleaned transaction rows for reuse in other tools.
- `profit_by_department.png` and `payment_method_counts.png` — chart images embedded in the workbook summary.
- `PROJECT_REPORT.md` — concise summary of method and results.
- `requirements.txt` — Python package list.
- `cleaned_data.csv` — clean transaction data for reuse in other tools.

## Cleaning steps

The script removes exchange tickers in parentheses from client names and lowercases them; trims extra spaces and standardizes contact names; splits the combined department and region; removes duplicate records after text normalization; fills blank payment methods with `NA`; leaves missing revenue values blank rather than inventing amounts; and calculates profit margin safely, returning `NA` when revenue is missing or zero.

The report includes row counts, financial totals, department/region summaries, payment method summaries, and charts. The `Raw Data` tab preserves the source records for comparison.

For the supplied workbook, the script processed 31 transaction rows, removed 3 duplicates, and produced 28 cleaned rows. It filled 2 missing payment methods with `NA`; it kept 2 missing revenue values blank and marked the corresponding profit margins `NA` rather than imputing revenue.

## Run

From this folder, install the dependencies and run the script:

```powershell
py -m pip install -r requirements.txt
py clean_data.py
```

The script recreates `Cleaned_Data_Report_Final.xlsx`, `cleaned_data.csv`, and the two chart images in this folder.
