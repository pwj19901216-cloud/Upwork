#!/usr/bin/env python3
"""Merge and clean many Excel / CSV files into one formatted report.

Handles the problems real client files usually have:
  * different column names for the same field ("Order ID", "order_no", "OrderNo")
  * title rows above the real header
  * blank rows, stray spaces, currency symbols and thousands separators
  * mixed date formats
  * duplicate rows across files

Usage:
    python merge_excel.py INPUT_DIR -o merged_report.xlsx
    python merge_excel.py INPUT_DIR -o out.xlsx --mapping columns.json --key order_id
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SUPPORTED = {".xlsx", ".xlsm", ".csv"}
SOURCE_COLS = ["source_file", "source_sheet"]

# Canonical name -> spellings seen in the wild (compared after normalising).
DEFAULT_ALIASES: dict[str, list[str]] = {
    "order_id": ["order id", "order no", "order number", "orderno", "order #", "id"],
    "order_date": ["date", "order date", "purchase date", "orderdate"],
    "customer": ["customer name", "client", "client name", "buyer"],
    "product": ["item", "product name", "sku name"],
    "quantity": ["qty", "units", "quantity sold"],
    "unit_price": ["price", "unit price", "price per unit"],
    "amount": ["total", "total amount", "revenue", "sales", "amount usd"],
    "region": ["area", "territory"],
}


def normalise_name(name: object) -> str:
    """'  Order No. ' -> 'order no'"""
    text = str(name).strip().lower()
    text = re.sub(r"[_\-.]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_alias_lookup(aliases: dict[str, list[str]]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for canonical, spellings in aliases.items():
        lookup[normalise_name(canonical)] = canonical
        for s in spellings:
            lookup[normalise_name(s)] = canonical
    return lookup


def canonical_column(name: object, lookup: dict[str, str]) -> str:
    norm = normalise_name(name)
    if norm in lookup:
        return lookup[norm]
    squashed = norm.replace(" ", "")
    for key, canonical in lookup.items():
        if key.replace(" ", "") == squashed:
            return canonical
    return norm.replace(" ", "_")


def find_header_row(raw: pd.DataFrame, scan_rows: int = 10) -> int:
    """Return the index of the first row that looks like a header.

    Title rows ("Q3 Sales Report - West") usually fill one cell, while the
    header fills as many cells as the data does.
    """
    head = raw.head(scan_rows)
    counts = head.notna().sum(axis=1)
    if counts.empty or counts.max() == 0:
        return 0
    return int(counts[counts == counts.max()].index[0])


def read_tables(path: Path) -> list[tuple[str, pd.DataFrame]]:
    """Read every sheet of a workbook (or a CSV) with header auto-detection."""
    if path.suffix.lower() == ".csv":
        raws = {"csv": pd.read_csv(path, header=None, dtype=object, skip_blank_lines=False)}
    else:
        raws = pd.read_excel(path, sheet_name=None, header=None, dtype=object)

    tables = []
    for sheet, raw in raws.items():
        raw = raw.dropna(how="all").dropna(axis=1, how="all")
        if raw.empty:
            continue
        header_idx = find_header_row(raw)
        header = raw.loc[header_idx]
        body = raw.loc[header_idx + 1 :].copy()
        body.columns = [str(h) if pd.notna(h) else f"unnamed_{i}" for i, h in enumerate(header)]
        tables.append((str(sheet), body.reset_index(drop=True)))
    return tables


_NUMBER_JUNK = re.compile(r"[^\d.\-]")


def to_number(value: object) -> float | None:
    """' $1,200.50 ' -> 1200.5, '(300)' -> -300, '' -> None"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    cleaned = _NUMBER_JUNK.sub("", text)
    if cleaned in {"", "-", ".", "-."}:
        return None
    try:
        number = float(cleaned)
    except ValueError:
        return None
    return -number if negative else number


@dataclass
class Settings:
    aliases: dict[str, list[str]] = field(default_factory=lambda: dict(DEFAULT_ALIASES))
    date_columns: set[str] = field(default_factory=lambda: {"order_date"})
    numeric_columns: set[str] = field(default_factory=lambda: {"quantity", "unit_price", "amount"})
    key_columns: list[str] = field(default_factory=list)
    dayfirst: bool = False


@dataclass
class FileStat:
    file: str
    sheet: str
    rows_read: int


@dataclass
class MergeResult:
    merged: pd.DataFrame
    duplicates: pd.DataFrame
    stats: list[FileStat]
    skipped: list[tuple[str, str]]


def clean_frame(df: pd.DataFrame, settings: Settings) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if col in SOURCE_COLS:
            continue
        if col in settings.numeric_columns:
            df[col] = pd.to_numeric(df[col].map(to_number), errors="coerce")
        elif col in settings.date_columns:
            df[col] = pd.to_datetime(
                df[col].map(lambda v: v.strip() if isinstance(v, str) else v),
                format="mixed",
                dayfirst=settings.dayfirst,
                errors="coerce",
            )
        else:
            df[col] = df[col].map(lambda v: re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v)
            df[col] = df[col].map(lambda v: None if isinstance(v, str) and v == "" else v)
    data_cols = [c for c in df.columns if c not in SOURCE_COLS]
    return df.dropna(how="all", subset=data_cols)


def merge_files(paths: list[Path], settings: Settings) -> MergeResult:
    lookup = build_alias_lookup(settings.aliases)
    frames, stats, skipped = [], [], []

    for path in paths:
        try:
            tables = read_tables(path)
        except Exception as exc:  # unreadable / corrupt file: report and continue
            skipped.append((path.name, str(exc)))
            continue
        for sheet, df in tables:
            df.columns = [canonical_column(c, lookup) for c in df.columns]
            df = df.loc[:, ~pd.Index(df.columns).duplicated()]
            df.insert(0, "source_file", path.name)
            df.insert(1, "source_sheet", sheet)
            df = clean_frame(df, settings)
            stats.append(FileStat(path.name, sheet, len(df)))
            frames.append(df)

    if not frames:
        empty = pd.DataFrame()
        return MergeResult(empty, empty, stats, skipped)

    merged = pd.concat(frames, ignore_index=True, sort=False)

    # Known columns first, in a sensible order; anything unexpected after.
    known = [c for c in settings.aliases if c in merged.columns]
    extra = [c for c in merged.columns if c not in known and c not in SOURCE_COLS]
    merged = merged[known + extra + SOURCE_COLS]

    subset = settings.key_columns or [c for c in merged.columns if c not in SOURCE_COLS]
    missing = [c for c in subset if c not in merged.columns]
    if missing:
        raise SystemExit(f"Key column(s) not found after merging: {', '.join(missing)}")
    dup_mask = merged.duplicated(subset=subset, keep="first")
    duplicates = merged[dup_mask].reset_index(drop=True)
    merged = merged[~dup_mask].reset_index(drop=True)

    if "order_date" in merged.columns:
        merged = merged.sort_values("order_date", kind="stable", na_position="last").reset_index(drop=True)

    return MergeResult(merged, duplicates, stats, skipped)


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _style_sheet(ws, df: pd.DataFrame, settings: Settings) -> None:
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    if ws.max_row > 1 and ws.max_column > 0:
        ws.auto_filter.ref = ws.dimensions

    for idx, col in enumerate(df.columns, start=1):
        letter = get_column_letter(idx)
        values = [str(col)] + [str(v) for v in df[col].head(500).tolist() if pd.notna(v)]
        ws.column_dimensions[letter].width = min(max(len(v) for v in values) + 2, 45)
        fmt = None
        if col in settings.date_columns:
            fmt = "yyyy-mm-dd"
        elif col in {"unit_price", "amount"}:
            fmt = "#,##0.00"
        elif col == "quantity":
            fmt = "#,##0"
        if fmt:
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                row[0].number_format = fmt
                if col in settings.date_columns:
                    ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width, 12)


def write_report(result: MergeResult, out_path: Path, settings: Settings) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)

    summary_rows = [
        {"file": s.file, "sheet": s.sheet, "rows_after_cleaning": s.rows_read} for s in result.stats
    ]
    summary = pd.DataFrame(summary_rows, columns=["file", "sheet", "rows_after_cleaning"])
    overview = pd.DataFrame(
        [
            ("Files / sheets read", len(result.stats)),
            ("Rows after cleaning", int(summary["rows_after_cleaning"].sum()) if len(summary) else 0),
            ("Duplicate rows removed", len(result.duplicates)),
            ("Rows in final report", len(result.merged)),
            ("Files skipped (unreadable)", len(result.skipped)),
        ],
        columns=["metric", "value"],
    )

    sheets: list[tuple[str, pd.DataFrame]] = [("Merged", result.merged), ("Overview", overview), ("Per file", summary)]
    if "amount" in result.merged.columns and "region" in result.merged.columns:
        by_region = (
            result.merged.groupby("region", dropna=False)["amount"]
            .agg(orders="count", total_amount="sum")
            .reset_index()
            .sort_values("total_amount", ascending=False)
        )
        sheets.append(("By region", by_region))
    sheets.append(("Duplicates removed", result.duplicates))
    if result.skipped:
        sheets.append(("Skipped files", pd.DataFrame(result.skipped, columns=["file", "error"])))

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets:
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            _style_sheet(ws, df, settings)
            if name == "By region":
                for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
                    row[0].number_format = "#,##0.00"


def collect_inputs(input_dir: Path, output: Path) -> list[Path]:
    files = []
    for p in sorted(input_dir.iterdir()):
        if not p.is_file() or p.suffix.lower() not in SUPPORTED:
            continue
        if p.name.startswith("~$"):  # Excel lock files
            continue
        if p.resolve() == output.resolve():
            continue
        files.append(p)
    return files


def load_settings(args: argparse.Namespace) -> Settings:
    settings = Settings(dayfirst=args.dayfirst, key_columns=args.key or [])
    if args.mapping:
        extra = json.loads(Path(args.mapping).read_text(encoding="utf-8"))
        for canonical, spellings in extra.get("aliases", {}).items():
            settings.aliases.setdefault(canonical, [])
            settings.aliases[canonical] = list(settings.aliases[canonical]) + list(spellings)
        settings.date_columns |= set(extra.get("date_columns", []))
        settings.numeric_columns |= set(extra.get("numeric_columns", []))
    return settings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Merge and clean Excel/CSV files into one report.")
    parser.add_argument("input_dir", type=Path, help="folder containing .xlsx / .xlsm / .csv files")
    parser.add_argument("-o", "--output", type=Path, default=Path("merged_report.xlsx"))
    parser.add_argument("--mapping", help="JSON file with extra column aliases / date / numeric columns")
    parser.add_argument("--key", nargs="+", help="columns that identify a unique row (default: all columns)")
    parser.add_argument("--dayfirst", action="store_true", help="read ambiguous dates as DD/MM/YYYY")
    args = parser.parse_args(argv)

    if not args.input_dir.is_dir():
        parser.error(f"{args.input_dir} is not a folder")

    settings = load_settings(args)
    paths = collect_inputs(args.input_dir, args.output)
    if not paths:
        print(f"No .xlsx/.xlsm/.csv files found in {args.input_dir}", file=sys.stderr)
        return 1

    result = merge_files(paths, settings)
    write_report(result, args.output, settings)

    print(f"Read {len(result.stats)} sheet(s) from {len(paths)} file(s)")
    print(f"Removed {len(result.duplicates)} duplicate row(s)")
    for name, err in result.skipped:
        print(f"Skipped {name}: {err}", file=sys.stderr)
    print(f"Wrote {len(result.merged)} row(s) to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
