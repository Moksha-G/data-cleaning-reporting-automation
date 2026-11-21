# Data Cleaning & Reporting Automation

## Objective

Automate the data preparation steps shown in the tutorial and produce a reviewable transaction report.

## Source and cleaning

The supplied workbook contains 31 transaction rows dated May 30 through June 2, 2023. The script removes exchange tickers from client names and lowercases them, standardizes contacts by trimming and collapsing spaces and applying proper case, and splits the combined department and region field. It removes 3 repeated records after text normalization, leaving 28 rows.

Two blank payment methods are labeled `NA`. Two missing revenue amounts remain blank because no reliable value is available. Their profit margins are labeled `NA`; the raw sheet uses an `IFERROR` formula so missing or zero revenue does not produce a spreadsheet error.

## Report results

- Revenue with values: **129,332.50 source units**.
- Profit across all cleaned rows: **28,845.00 source units**.
- Weighted profit margin for rows with known revenue: **20.9%**.

The source file does not identify a currency, so amounts are reported as source units. The workbook includes department and region totals, payment method counts and revenue, and two charts. It keeps a raw transaction sheet alongside the cleaned output for comparison.

## Reproduction

Run `py -m pip install -r requirements.txt` followed by `py clean_data.py` from this folder. The script regenerates the Excel report and cleaned CSV from `Data Cleaning Start.xlsx`.

## Limitation

Revenue is not estimated for the two incomplete rows. Revenue totals and margin calculations use only rows with recorded revenue; total profit includes all cleaned rows.
