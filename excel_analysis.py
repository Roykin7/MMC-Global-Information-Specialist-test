#!/usr/bin/env python3
"""Analyze an Excel workbook and generate a markdown report with per-sheet summaries."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import pandas as pd


DEFAULT_FILE = "IMS_Test_Data 1.xlsx"
OUTPUT_DIR = "analysis_output"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyze an Excel workbook")
    parser.add_argument(
        "--file",
        default=DEFAULT_FILE,
        help="Path to Excel file (default: IMS_Test_Data 1.xlsx)",
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help="Directory where reports are written",
    )
    return parser.parse_args()


def _safe_name(name: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(name))
    return cleaned.strip("_") or "sheet"


def _coerce_possible_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Attempt date parsing for object columns when conversion success is high."""
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == "object":
            parsed = pd.to_datetime(out[col], errors="coerce")
            parse_ratio = parsed.notna().mean()
            if parse_ratio >= 0.7:
                out[col] = parsed
    return out


def _section(title: str) -> str:
    return f"\n## {title}\n"


def _to_markdown_or_string(df: pd.DataFrame, index: bool = True) -> str:
    if df.empty:
        return "(none)"
    try:
        return df.to_markdown(index=index)
    except Exception:
        return df.to_string(index=index)


def _format_list(items: Iterable[str]) -> str:
    items = list(items)
    if not items:
        return "(none)"
    return "\n".join(f"- {item}" for item in items)


def analyze_sheet(sheet_name: str, df: pd.DataFrame, output_dir: Path) -> str:
    df = _coerce_possible_dates(df)

    safe_sheet = _safe_name(sheet_name)
    report_parts: list[str] = []

    report_parts.append(f"\n# Sheet: {sheet_name}\n")
    report_parts.append(f"Rows: {len(df):,}")
    report_parts.append(f"Columns: {len(df.columns):,}")

    dtypes = df.dtypes.astype(str).rename("dtype").to_frame()
    report_parts.append(_section("Column Types"))
    report_parts.append(_to_markdown_or_string(dtypes))

    missing = pd.DataFrame(
        {
            "missing_count": df.isna().sum(),
            "missing_pct": (df.isna().mean() * 100).round(2),
        }
    ).sort_values("missing_count", ascending=False)

    report_parts.append(_section("Missing Values"))
    report_parts.append(_to_markdown_or_string(missing))
    missing.to_csv(output_dir / f"{safe_sheet}_missing_values.csv")

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    datetime_cols = df.select_dtypes(include=["datetime64[ns]", "datetimetz"]).columns.tolist()
    categorical_cols = df.select_dtypes(include=["object", "category", "bool"]).columns.tolist()

    report_parts.append(_section("Column Groups"))
    report_parts.append("Numeric columns:\n" + _format_list(numeric_cols))
    report_parts.append("\nDatetime columns:\n" + _format_list(datetime_cols))
    report_parts.append("\nCategorical columns:\n" + _format_list(categorical_cols))

    if numeric_cols:
        num_summary = df[numeric_cols].describe().T.round(3)
        report_parts.append(_section("Numeric Summary"))
        report_parts.append(_to_markdown_or_string(num_summary))
        num_summary.to_csv(output_dir / f"{safe_sheet}_numeric_summary.csv")

        if len(numeric_cols) >= 2:
            corr = df[numeric_cols].corr(numeric_only=True).round(3)
            report_parts.append(_section("Correlation Matrix (Numeric)"))
            report_parts.append(_to_markdown_or_string(corr))
            corr.to_csv(output_dir / f"{safe_sheet}_correlation.csv")

    if categorical_cols:
        report_parts.append(_section("Top Category Values"))
        for col in categorical_cols:
            top_vals = (
                df[col]
                .astype("string")
                .fillna("<NA>")
                .value_counts(dropna=False)
                .head(10)
                .rename_axis(col)
                .rename("count")
                .to_frame()
            )
            report_parts.append(f"\n### {col}\n")
            report_parts.append(_to_markdown_or_string(top_vals))
            top_vals.to_csv(output_dir / f"{safe_sheet}_{_safe_name(col)}_top_values.csv")

    if datetime_cols:
        report_parts.append(_section("Datetime Ranges"))
        ranges = pd.DataFrame(
            {
                "min": [df[col].min() for col in datetime_cols],
                "max": [df[col].max() for col in datetime_cols],
                "non_null": [df[col].notna().sum() for col in datetime_cols],
            },
            index=datetime_cols,
        )
        report_parts.append(_to_markdown_or_string(ranges))
        ranges.to_csv(output_dir / f"{safe_sheet}_datetime_ranges.csv")

    return "\n".join(report_parts)


def main() -> None:
    args = parse_args()
    excel_path = Path(args.file)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    workbook = pd.ExcelFile(excel_path)

    report = [
        "# Excel Analysis Report",
        "",
        f"Workbook: {excel_path.name}",
        f"Sheets: {', '.join(workbook.sheet_names)}",
    ]

    for sheet_name in workbook.sheet_names:
        sheet_df = pd.read_excel(excel_path, sheet_name=sheet_name)
        report.append(analyze_sheet(sheet_name, sheet_df, output_dir))

    report_path = output_dir / "excel_analysis_report.md"
    report_path.write_text("\n".join(report), encoding="utf-8")

    print(f"Analysis complete. Report: {report_path}")
    print(f"Additional sheet outputs written to: {output_dir}")


if __name__ == "__main__":
    main()
