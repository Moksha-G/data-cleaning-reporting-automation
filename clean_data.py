"""Clean the tutorial workbook and produce a formatted Excel report."""

from pathlib import Path
import re

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as SheetImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "Data Cleaning Start.xlsx"
OUTPUT = ROOT / "Cleaned_Data_Report_Final.xlsx"
CSV_OUTPUT = ROOT / "cleaned_data.csv"

NAVY = "17365D"
BLUE = "4472C4"
PALE_BLUE = "D9EAF7"
PALE_YELLOW = "FFF2CC"
GRAY = "E7E6E6"
WHITE = "FFFFFF"
TEXT = "202124"
THIN_BLUE = Side(style="thin", color="B4C7E7")


def load_source(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Place the source workbook next to this script: {path.name}")
    source_book = load_workbook(path, data_only=False, read_only=True)
    if "Sheet1 (2)" not in source_book.sheetnames:
        raise ValueError("Expected a worksheet named 'Sheet1 (2)'.")
    source_sheet = source_book["Sheet1 (2)"]
    headers = [cell.value for cell in source_sheet[2][1:9]]
    expected = ["Date", "Client", "Contact", "Department", "Payment", "Revenue", "Profit", "Profit Margin"]
    if headers != expected:
        raise ValueError(f"Unexpected source headers: {headers}")

    records = []
    for row in source_sheet.iter_rows(min_row=3, min_col=2, max_col=8, values_only=True):
        # Source includes formatted empty rows below the actual records.
        if any(value is not None for value in row):
            records.append(row)
    source_book.close()
    return pd.DataFrame(records, columns=expected[:7])


def clean_data(raw: pd.DataFrame):
    df = raw.copy()
    source_rows = len(df)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce")
    df["Profit"] = pd.to_numeric(df["Profit"], errors="coerce")

    # Normalize company labels and contact names.
    original_clients = df["Client"].fillna("").astype(str)
    ticker_count = int(original_clients.str.contains(r"\([^)]*\)", regex=True).sum())
    df["Client"] = (original_clients
                    .str.replace(r"\s*\([^)]*\)", "", regex=True)
                    .str.strip().str.lower())

    original_contacts = df["Contact"].fillna("").astype(str)
    normalized_contacts = original_contacts.map(lambda value: re.sub(r"\s+", " ", value).strip().title())
    contact_count = int((original_contacts != normalized_contacts).sum())
    df["Contact"] = normalized_contacts

    # Split the tutorial's Department_Region field into two fields.
    department_region = df["Department"].fillna("").astype(str).str.split("_", n=1, expand=True)
    department = department_region[0].str.strip().str.title()
    region = department_region[1].fillna("").str.strip().str.title()
    df["Department"] = department.where(department.ne(""), "NA")
    df["Region"] = region.where(region.ne(""), "NA")

    # Remove repeated records after text normalization, excluding the source formula column.
    duplicate_keys = ["Date", "Client", "Contact", "Department", "Region", "Payment", "Revenue", "Profit"]
    duplicate_rows = int(df.duplicated(subset=duplicate_keys, keep="first").sum())
    df = df.drop_duplicates(subset=duplicate_keys, keep="first").reset_index(drop=True)

    missing_payment = int(df["Payment"].isna().sum())
    missing_revenue = int(df["Revenue"].isna().sum())
    missing_profit = int(df["Profit"].isna().sum())
    df["Payment"] = df["Payment"].fillna("NA").astype(str).str.strip()
    df.loc[df["Payment"].eq(""), "Payment"] = "NA"

    # IFERROR equivalent: margins with missing/zero revenue are labeled NA.
    def safe_margin(row):
        revenue, profit = row["Revenue"], row["Profit"]
        if pd.isna(revenue) or revenue == 0 or pd.isna(profit):
            return "NA"
        return float(profit / revenue)

    df["Profit Margin"] = df.apply(safe_margin, axis=1)
    df = df[["Date", "Client", "Contact", "Department", "Region", "Payment", "Revenue", "Profit", "Profit Margin"]]
    df = df.sort_values(["Date", "Department", "Client"], kind="stable", na_position="last").reset_index(drop=True)

    actions = {
        "source_rows": source_rows,
        "duplicate_rows": duplicate_rows,
        "clean_rows": len(df),
        "ticker_count": ticker_count,
        "contact_count": contact_count,
        "missing_payment": missing_payment,
        "missing_revenue": missing_revenue,
        "missing_profit": missing_profit,
        "margin_na": int((df["Profit Margin"] == "NA").sum()),
    }
    return df, actions


def amount(value):
    return None if pd.isna(value) else float(value)


def summaries(df: pd.DataFrame):
    known = df["Revenue"].notna() & df["Profit"].notna() & df["Revenue"].ne(0)
    known_revenue = float(df.loc[known, "Revenue"].sum())
    matched_profit = float(df.loc[known, "Profit"].sum())
    weighted_margin = matched_profit / known_revenue if known_revenue else None

    department_rows = []
    for (dept, region), group in df.groupby(["Department", "Region"], dropna=False, sort=True):
        valid = group["Revenue"].notna() & group["Revenue"].ne(0) & group["Profit"].notna()
        revenue = float(group["Revenue"].sum())
        profit = float(group["Profit"].sum())
        valid_revenue = float(group.loc[valid, "Revenue"].sum())
        valid_profit = float(group.loc[valid, "Profit"].sum())
        margin = valid_profit / valid_revenue if valid_revenue else None
        department_rows.append([dept, region, int(len(group)), revenue, profit, margin])

    payment_rows = []
    for payment, group in df.groupby("Payment", dropna=False, sort=True):
        payment_rows.append([payment, int(len(group)), float(group["Revenue"].sum())])
    metrics = {
        "known_revenue": known_revenue,
        "profit_all_records": float(df["Profit"].sum()),
        "weighted_margin": weighted_margin,
    }
    return metrics, department_rows, payment_rows


def style_title(ws, title, subtitle, end_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    c = ws.cell(1, 1, title)
    c.font = Font(name="Aptos Display", size=16, bold=True, color=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 27
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    c = ws.cell(2, 1, subtitle)
    c.font = Font(name="Aptos", size=10, italic=True, color="5B6573")
    ws.row_dimensions[2].height = 20


def format_table_header(ws, row, start_col, end_col):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN_BLUE)
    ws.row_dimensions[row].height = 28


def add_table(ws, ref, name):
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                          showLastColumn=False, showRowStripes=True,
                                          showColumnStripes=False)
    ws.add_table(table)


def create_charts(department_rows, payment_rows):
    regular_font = Path("C:/Windows/Fonts/arial.ttf")
    bold_font = Path("C:/Windows/Fonts/arialbd.ttf")
    try:
        label_font = ImageFont.truetype(str(regular_font), 20)
        title_font = ImageFont.truetype(str(bold_font), 27)
        small_font = ImageFont.truetype(str(regular_font), 16)
    except OSError:
        label_font = ImageFont.load_default()
        title_font = ImageFont.load_default()
        small_font = ImageFont.load_default()

    def horizontal_bars(title, labels, values, axis_title, color, destination):
        width = 960
        row_height = 62
        top = 100
        bottom = 82
        height = top + row_height * len(labels) + bottom
        left = 260
        right = 100
        plot_width = width - left - right
        maximum = max(values) if values else 1
        axis_max = maximum * 1.2
        canvas = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(canvas)
        draw.text((36, 28), title, font=title_font, fill=f"#{NAVY}")
        grid_top = top - 12
        grid_bottom = top + row_height * len(labels) - 8
        for tick in range(6):
            x = left + int(plot_width * tick / 5)
            draw.line((x, grid_top, x, grid_bottom), fill="#E5EAF1", width=2)
            tick_value = axis_max * tick / 5
            draw.text((x, grid_bottom + 10), f"{tick_value:,.0f}", font=small_font,
                      fill="#536273", anchor="mt")
        for idx, (label, value) in enumerate(zip(labels, values)):
            y_center = top + idx * row_height + row_height // 2 - 8
            draw.text((left - 18, y_center), label, font=label_font,
                      fill="#263548", anchor="rm")
            bar_width = int(plot_width * value / axis_max)
            draw.rounded_rectangle((left, y_center - 15, left + bar_width, y_center + 15),
                                   radius=7, fill=color)
            draw.text((left + bar_width + 10, y_center), f"{value:,.0f}", font=label_font,
                      fill="#263548", anchor="lm")
        draw.text((left + plot_width // 2, height - 24), axis_title, font=small_font,
                  fill="#536273", anchor="mm")
        canvas.save(destination, format="PNG", optimize=True)

    dept = sorted(department_rows, key=lambda row: row[4], reverse=True)
    horizontal_bars(
        "Profit by Department",
        [f"{row[0]} / {row[1]}" for row in dept],
        [row[4] for row in dept],
        "Profit (source units)",
        "#4472C4",
        ROOT / "profit_by_department.png",
    )

    payments = sorted(payment_rows, key=lambda row: row[1], reverse=True)
    horizontal_bars(
        "Records by Payment Method",
        [str(row[0]) for row in payments],
        [row[1] for row in payments],
        "Number of records",
        "#69A88D",
        ROOT / "payment_method_counts.png",
    )
    dept_path = ROOT / "profit_by_department.png"
    payment_path = ROOT / "payment_method_counts.png"
    return dept_path, payment_path


def build_workbook(df, actions, metrics, department_rows, payment_rows, raw, chart_paths):
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    clean = wb.create_sheet("Cleaned Data")
    raw_ws = wb.create_sheet("Raw Data")

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Summary output.
    style_title(summary, "Data Cleaning & Reporting", "Automated cleanup of the supplied transaction dataset | financial values in source units", 14)
    kpis = [
        (1, "Source rows", actions["source_rows"]),
        (4, "Duplicates removed", actions["duplicate_rows"]),
        (7, "Rows after cleaning", actions["clean_rows"]),
        (10, "Payment blanks filled", actions["missing_payment"]),
    ]
    for col, label, value in kpis:
        summary.cell(4, col, label)
        summary.cell(4, col).font = Font(name="Aptos", size=10, bold=True, color=NAVY)
        summary.cell(4, col).fill = PatternFill("solid", fgColor=PALE_BLUE)
        summary.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        summary.cell(5, col, value)
        summary.cell(5, col).font = Font(name="Aptos Display", size=16, bold=True, color=TEXT)
        summary.cell(5, col).fill = PatternFill("solid", fgColor="F3F7FC")
        summary.cell(5, col).alignment = Alignment(horizontal="left")
        summary.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)

    summary.cell(7, 1, "Financial summary")
    summary.cell(7, 1).font = Font(name="Aptos", size=11, bold=True, color=NAVY)
    financial = [
        (8, "Revenue with values", metrics["known_revenue"], '#,##0.00'),
        (9, "Profit, all cleaned rows", metrics["profit_all_records"], '#,##0.00'),
        (10, "Profit margin on rows with revenue", metrics["weighted_margin"], '0.0%'),
    ]
    for row, label, value, number_format in financial:
        summary.cell(row, 1, label)
        summary.cell(row, 1).font = Font(name="Aptos", size=10, color=TEXT)
        summary.cell(row, 2, value)
        summary.cell(row, 2).number_format = number_format
        summary.cell(row, 2).font = Font(name="Aptos", size=10, bold=True, color=TEXT)

    summary.cell(12, 1, "Cleaning actions")
    summary.cell(12, 1).font = Font(name="Aptos", size=11, bold=True, color=NAVY)
    actions_header = ["Field", "Automated action", "Rows / cells", "Handling note"]
    for col, value in enumerate(actions_header, 1):
        summary.cell(13, col, value)
    format_table_header(summary, 13, 1, 4)
    action_rows = [
        ["Client", "Removed ticker; lowercased client", actions["ticker_count"], "Kept the client name"],
        ["Contact", "Trimmed spaces; proper case", actions["contact_count"], "Collapsed repeated spaces"],
        ["Department", "Split department and region", actions["clean_rows"], "Both values kept in separate columns"],
        ["Duplicates", "Removed repeated normalized rows", actions["duplicate_rows"], "Kept the first occurrence"],
        ["Payment", "Filled blank methods with NA", actions["missing_payment"], "No payment method was inferred"],
        ["Revenue", "Kept missing amounts blank", actions["missing_revenue"], "No revenue values were imputed"],
        ["Profit Margin", "Used IFERROR for margin calculation", actions["margin_na"], "Missing or zero revenue returns NA"],
    ]
    for r, values in enumerate(action_rows, 14):
        for col, value in enumerate(values, 1):
            summary.cell(r, col, value)
            summary.cell(r, col).font = Font(name="Aptos", size=9, color=TEXT)
            summary.cell(r, col).alignment = Alignment(vertical="center", wrap_text=(col in (2, 4)))
            if r % 2 == 0:
                summary.cell(r, col).fill = PatternFill("solid", fgColor="F5F8FC")

    dept_header_row = 24
    summary.cell(dept_header_row - 1, 1, "Department and region summary")
    summary.cell(dept_header_row - 1, 1).font = Font(name="Aptos", size=11, bold=True, color=NAVY)
    dept_headers = ["Department", "Region", "Records", "Revenue (known)", "Profit (all rows)", "Margin (known revenue)"]
    for col, value in enumerate(dept_headers, 1):
        summary.cell(dept_header_row, col, value)
    format_table_header(summary, dept_header_row, 1, len(dept_headers))
    for r, values in enumerate(department_rows, dept_header_row + 1):
        for col, value in enumerate(values, 1):
            summary.cell(r, col, value)
            summary.cell(r, col).font = Font(name="Aptos", size=9, color=TEXT)
        summary.cell(r, 4).number_format = '#,##0.00'
        summary.cell(r, 5).number_format = '#,##0.00'
        summary.cell(r, 6).number_format = '0.0%'
    dept_end = dept_header_row + len(department_rows)
    add_table(summary, f"A{dept_header_row}:F{dept_end}", "DepartmentSummary")

    pay_header_row = 24
    pay_start_col = 8
    summary.cell(pay_header_row - 1, pay_start_col, "Payment method summary")
    summary.cell(pay_header_row - 1, pay_start_col).font = Font(name="Aptos", size=11, bold=True, color=NAVY)
    pay_headers = ["Payment", "Records", "Revenue (known)"]
    for col, value in enumerate(pay_headers, pay_start_col):
        summary.cell(pay_header_row, col, value)
    format_table_header(summary, pay_header_row, pay_start_col, pay_start_col + 2)
    for r, values in enumerate(payment_rows, pay_header_row + 1):
        for col, value in enumerate(values, pay_start_col):
            summary.cell(r, col, value)
            summary.cell(r, col).font = Font(name="Aptos", size=9, color=TEXT)
        summary.cell(r, pay_start_col + 2).number_format = '#,##0.00'
    pay_end = pay_header_row + len(payment_rows)
    add_table(summary, f"H{pay_header_row}:J{pay_end}", "PaymentSummary")

    for chart_path, anchor, width, height in [
        (chart_paths[0], "H7", 760, 400),
        (chart_paths[1], "H32", 760, 380),
    ]:
        chart_image = SheetImage(str(chart_path))
        chart_image.width = width
        chart_image.height = height
        summary.add_image(chart_image, anchor)
    summary.freeze_panes = "A4"
    for col, width in {"A": 29, "B": 26, "C": 15, "D": 32, "E": 19, "F": 23,
                       "G": 4, "H": 17, "I": 13, "J": 19, "K": 4, "L": 13, "M": 13, "N": 13}.items():
        summary.column_dimensions[col].width = width
    for r in range(14, 21):
        summary.row_dimensions[r].height = 34

    # Cleaned table.
    clean_headers = list(df.columns)
    clean.append(clean_headers)
    for row in df.itertuples(index=False, name=None):
        values = list(row)
        values[0] = values[0].to_pydatetime() if hasattr(values[0], "to_pydatetime") else values[0]
        values[6] = amount(values[6])
        values[7] = amount(values[7])
        clean.append(values)
    format_table_header(clean, 1, 1, len(clean_headers))
    add_table(clean, f"A1:I{len(df) + 1}", "CleanedTransactions")
    clean.freeze_panes = "A2"
    for row in range(2, len(df) + 2):
        clean.cell(row, 1).number_format = "yyyy-mm-dd"
        clean.cell(row, 7).number_format = '#,##0.00'
        clean.cell(row, 8).number_format = '#,##0.00'
        clean.cell(row, 9).number_format = '0.0%'
        if clean.cell(row, 7).value is None:
            clean.cell(row, 7).fill = PatternFill("solid", fgColor=PALE_YELLOW)
        if clean.cell(row, 9).value == "NA":
            clean.cell(row, 9).fill = PatternFill("solid", fgColor=PALE_YELLOW)
    clean_widths = [14, 27, 23, 20, 17, 16, 16, 16, 20]
    for col, width in enumerate(clean_widths, 1):
        clean.column_dimensions[get_column_letter(col)].width = width

    # Raw source values remain unmodified for auditability.
    raw_headers = ["Date", "Client", "Contact", "Department", "Payment", "Revenue", "Profit", "Profit Margin"]
    raw_ws.append(raw_headers)
    for row_index, source_row in enumerate(raw.itertuples(index=False, name=None), start=2):
        values = list(source_row)
        values[0] = values[0].to_pydatetime() if hasattr(values[0], "to_pydatetime") else values[0]
        values[5] = amount(values[5])
        values[6] = amount(values[6])
        raw_ws.append(values[:7] + [f'=IFERROR(G{row_index}/F{row_index},"NA")'])
    format_table_header(raw_ws, 1, 1, len(raw_headers))
    add_table(raw_ws, f"A1:H{len(raw) + 1}", "RawTransactions")
    raw_ws.freeze_panes = "A2"
    for row in range(2, len(raw) + 2):
        raw_ws.cell(row, 1).number_format = "yyyy-mm-dd"
        raw_ws.cell(row, 6).number_format = '#,##0.00'
        raw_ws.cell(row, 7).number_format = '#,##0.00'
        raw_ws.cell(row, 8).number_format = '0.0%'
    raw_widths = [14, 38, 26, 24, 16, 16, 16, 20]
    for col, width in enumerate(raw_widths, 1):
        raw_ws.column_dimensions[get_column_letter(col)].width = width

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(OUTPUT)


def main():
    raw = load_source(SOURCE)
    cleaned, actions = clean_data(raw)
    metrics, department_rows, payment_rows = summaries(cleaned)
    chart_paths = create_charts(department_rows, payment_rows)
    build_workbook(cleaned, actions, metrics, department_rows, payment_rows, raw, chart_paths)
    export = cleaned.copy()
    export["Date"] = export["Date"].dt.strftime("%Y-%m-%d")
    export.to_csv(CSV_OUTPUT, index=False, na_rep="")
    print(f"Input rows: {actions['source_rows']}")
    print(f"Duplicates removed: {actions['duplicate_rows']}")
    print(f"Clean rows: {actions['clean_rows']}")
    print(f"Missing payment values filled: {actions['missing_payment']}")
    print(f"Missing revenue values retained: {actions['missing_revenue']}")
    print(f"Profit margin unavailable: {actions['margin_na']}")
    print(f"Saved workbook: {OUTPUT}")
    print(f"Saved cleaned CSV: {CSV_OUTPUT}")


if __name__ == "__main__":
    main()
