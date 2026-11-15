#!/usr/bin/env python3
"""Generate deliberately messy sample files that look like real client data.

    python make_sample_data.py sample_data
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook


def west(folder: Path) -> None:
    """Title rows above the header, '$1,200.50' text amounts, a blank row."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Q3"
    ws.append(["Q3 Sales Report - West Region"])
    ws.append([])
    ws.append(["Order ID", "Order Date", "Customer Name", "Item", "Qty", "Price", "Total", "Region"])
    rows = [
        ["W-1001", "07/03/2026", "  Acme Corp ", "Desk Lamp", 4, "$25.00", "$100.00", "West"],
        ["W-1002", "07/09/2026", "Blue Harbor LLC", "Office Chair", 2, "$149.50", "$299.00", "West"],
        [None, None, None, None, None, None, None, None],
        ["W-1003", "2026-07-15", "Northwind", "Monitor 27in", 3, "$1,200.50", "$3,601.50", "West"],
        ["W-1004", "July 21, 2026", "acme corp", "Desk Lamp", 10, "25", "250", "West"],
    ]
    for r in rows:
        ws.append(r)
    wb.save(folder / "sales_west.xlsx")


def east(folder: Path) -> None:
    """Different column names, real Excel dates, two sheets, a duplicate of West."""
    wb = Workbook()
    ws = wb.active
    ws.title = "July"
    ws.append(["order_no", "date", "client", "product_name", "units", "unit price", "revenue", "territory"])
    ws.append(["E-2001", datetime(2026, 7, 2), "Contoso", "Office Chair", 1, 149.5, 149.5, "East"])
    ws.append(["E-2002", datetime(2026, 7, 18), "Fabrikam Inc.", "Standing Desk", 2, 420, 840, "East"])
    ws2 = wb.create_sheet("August")
    ws2.append(["order_no", "date", "client", "product_name", "units", "unit price", "revenue", "territory"])
    ws2.append(["E-2003", datetime(2026, 8, 5), "Contoso", "Monitor 27in", 2, 1200.5, 2401, "East"])
    ws2.append(["E-2004", datetime(2026, 8, 11), "Tailspin Toys", "Desk Lamp", 6, 25, 150, "East"])
    # Exact duplicate of E-2002 pasted twice by mistake
    ws2.append(["E-2002", datetime(2026, 7, 18), "Fabrikam Inc.", "Standing Desk", 2, 420, 840, "East"])
    wb.save(folder / "sales_east.xlsx")


def online(folder: Path) -> None:
    """A CSV export with yet another naming style and an extra column."""
    df = pd.DataFrame(
        {
            "OrderNo": ["O-3001", "O-3002", "O-3003", "O-3003"],
            "Purchase Date": ["2026-08-01", "2026-08-03 14:22", "2026-08-09", "2026-08-09"],
            "Buyer": ["Litware", "Wide World Importers", "Proseware ", "Proseware "],
            "SKU Name": ["Standing Desk", "Desk Lamp", "Office Chair", "Office Chair"],
            "Quantity Sold": ["1", "12", "3", "3"],
            "Price per Unit": ["420.00", "25.00", "149.50", "149.50"],
            "Amount USD": ["420.00", "300.00", "448.50", "448.50"],
            "Area": ["Online", "Online", "Online", "Online"],
            "Coupon": ["", "SUMMER10", "", ""],
        }
    )
    df.to_csv(folder / "sales_online.csv", index=False)


def main() -> None:
    folder = Path(sys.argv[1] if len(sys.argv) > 1 else "sample_data")
    folder.mkdir(parents=True, exist_ok=True)
    west(folder)
    east(folder)
    online(folder)
    print(f"Sample files written to {folder}/")


if __name__ == "__main__":
    main()
