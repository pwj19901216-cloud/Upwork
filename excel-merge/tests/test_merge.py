import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))

import merge_excel as m  # noqa: E402


@pytest.fixture()
def sample_dir(tmp_path):
    subprocess.run([sys.executable, str(HERE / "make_sample_data.py"), str(tmp_path)], check=True)
    return tmp_path


def test_to_number():
    assert m.to_number(" $1,200.50 ") == 1200.5
    assert m.to_number("(300)") == -300
    assert m.to_number("") is None
    assert m.to_number(7) == 7.0
    assert m.to_number("n/a") is None


def test_canonical_column():
    lookup = m.build_alias_lookup(m.DEFAULT_ALIASES)
    assert m.canonical_column("Order No.", lookup) == "order_id"
    assert m.canonical_column("OrderNo", lookup) == "order_id"
    assert m.canonical_column("  Amount_USD ", lookup) == "amount"
    assert m.canonical_column("Coupon Code", lookup) == "coupon_code"


def test_merge_end_to_end(sample_dir, tmp_path):
    out = tmp_path / "out" / "report.xlsx"
    assert m.main([str(sample_dir), "-o", str(out)]) == 0

    merged = pd.read_excel(out, sheet_name="Merged")
    # 4 west + 4 east (one duplicate removed) + 3 online (one duplicate removed)
    assert len(merged) == 11
    assert merged["order_id"].is_unique
    assert {"order_id", "order_date", "customer", "amount", "region", "coupon"} <= set(merged.columns)

    w3 = merged.set_index("order_id").loc["W-1003"]
    assert w3["amount"] == pytest.approx(3601.5)
    assert pd.Timestamp(w3["order_date"]) == pd.Timestamp("2026-07-15")
    assert merged.set_index("order_id").loc["W-1001", "customer"] == "Acme Corp"

    dups = pd.read_excel(out, sheet_name="Duplicates removed")
    assert sorted(dups["order_id"]) == ["E-2002", "O-3003"]

    by_region = pd.read_excel(out, sheet_name="By region").set_index("region")
    assert by_region.loc["East", "orders"] == 4
    assert by_region.loc["West", "total_amount"] == pytest.approx(100 + 299 + 3601.5 + 250)


def test_key_columns_and_bad_file(sample_dir, tmp_path):
    (sample_dir / "broken.xlsx").write_bytes(b"not really an excel file")
    out = tmp_path / "report.xlsx"
    assert m.main([str(sample_dir), "-o", str(out), "--key", "order_id"]) == 0
    skipped = pd.read_excel(out, sheet_name="Skipped files")
    assert list(skipped["file"]) == ["broken.xlsx"]


def test_empty_folder(tmp_path):
    assert m.main([str(tmp_path), "-o", str(tmp_path / "x.xlsx")]) == 1
