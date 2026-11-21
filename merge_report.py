"""
Excel Merge & Auto-Report
-------------------------
Merges every Excel/CSV file in a folder (e.g. daily or per-branch sales sheets)
into one master sheet, cleans the data, and builds a ready-to-send report with
summary tables and charts.

Usage:
    python merge_report.py --input sample_data --out sales_report.xlsx
    python make_sample_data.py   # generates demo files first
"""

import argparse
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

REQUIRED = ["Date", "Branch", "Product", "Quantity", "Unit Price"]


def load_files(folder: Path) -> pd.DataFrame:
    frames = []
    files = sorted([*folder.glob("*.xlsx"), *folder.glob("*.xls"), *folder.glob("*.csv")])
    if not files:
        raise SystemExit(f"No Excel/CSV files found in {folder}")

    for f in files:
        df = pd.read_csv(f) if f.suffix == ".csv" else pd.read_excel(f)
        df.columns = [str(c).strip().title() for c in df.columns]
        missing = [c for c in REQUIRED if c not in df.columns]
        if missing:
            print(f"  ! Skipped {f.name}: missing columns {missing}")
            continue
        df["Source File"] = f.name
        frames.append(df)
        print(f"  + {f.name}: {len(df)} rows")

    if not frames:
        raise SystemExit("No valid files to merge.")
    return pd.concat(frames, ignore_index=True)


def clean(df: pd.DataFrame) -> pd.DataFrame:
    before = len(df)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce")
    for col in ("Branch", "Product"):
        df[col] = df[col].astype(str).str.strip()
    df = df.dropna(subset=["Date", "Quantity", "Unit Price"])
    df = df.drop_duplicates(subset=REQUIRED)
    df["Total"] = (df["Quantity"] * df["Unit Price"]).round(2)
    print(f"  Cleaned: {before} -> {len(df)} rows (removed invalid/duplicates)")
    return df.sort_values("Date").reset_index(drop=True)


def style_sheet(ws, money_cols=()):
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    for i, col in enumerate(ws.columns, start=1):
        width = max(len(str(c.value if c.value is not None else "")) for c in col)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 3, 40)
        if ws.cell(1, i).value in money_cols:
            for c in col[1:]:
                c.number_format = "#,##0.00"
    ws.freeze_panes = "A2"


def build_report(df: pd.DataFrame, out: Path) -> None:
    by_branch = df.groupby("Branch", as_index=False).agg(Orders=("Total", "count"), Revenue=("Total", "sum"))
    by_branch = by_branch.sort_values("Revenue", ascending=False)
    by_product = df.groupby("Product", as_index=False).agg(Units=("Quantity", "sum"), Revenue=("Total", "sum"))
    by_product = by_product.sort_values("Revenue", ascending=False)
    daily = df.groupby(df["Date"].dt.date, as_index=False).agg(Revenue=("Total", "sum"))

    kpis = pd.DataFrame({
        "Metric": ["Total Revenue", "Total Orders", "Units Sold", "Average Order Value",
                   "Top Branch", "Top Product", "Period"],
        "Value": [
            round(df["Total"].sum(), 2), len(df), int(df["Quantity"].sum()),
            round(df["Total"].mean(), 2), by_branch.iloc[0]["Branch"], by_product.iloc[0]["Product"],
            f"{df['Date'].min():%Y-%m-%d} to {df['Date'].max():%Y-%m-%d}",
        ],
    })

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        kpis.to_excel(w, sheet_name="Dashboard", index=False)
        by_branch.to_excel(w, sheet_name="By Branch", index=False)
        by_product.to_excel(w, sheet_name="By Product", index=False)
        daily.to_excel(w, sheet_name="Daily", index=False)
        df.assign(Date=df["Date"].dt.date).to_excel(w, sheet_name="All Data", index=False)

        wb = w.book
        for ws in wb.worksheets:
            style_sheet(ws, money_cols=("Revenue", "Total", "Unit Price"))

        # Charts
        ws = wb["By Branch"]
        bar = BarChart(); bar.title = "Revenue by Branch"; bar.height, bar.width = 8, 16
        bar.add_data(Reference(ws, min_col=3, min_row=1, max_row=ws.max_row), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
        wb["Dashboard"].add_chart(bar, "D2")

        ws = wb["Daily"]
        line = LineChart(); line.title = "Daily Revenue"; line.height, line.width = 8, 16
        line.add_data(Reference(ws, min_col=2, min_row=1, max_row=ws.max_row), titles_from_data=True)
        line.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
        wb["Dashboard"].add_chart(line, "D19")


def main():
    p = argparse.ArgumentParser(description="Merge Excel/CSV files and build a report")
    p.add_argument("--input", default="sample_data", help="folder with Excel/CSV files")
    p.add_argument("--out", default="sales_report.xlsx", help="output report file")
    args = p.parse_args()

    print(f"Reading files from {args.input}/")
    df = clean(load_files(Path(args.input)))
    build_report(df, Path(args.out))
    print(f"\nReport saved: {args.out}")


if __name__ == "__main__":
    main()
