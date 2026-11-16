"""Offline tests: parse saved pages and run a full crawl against a fake site."""
import sys
from pathlib import Path

import pandas as pd
import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

import scrape_books as s  # noqa: E402

LIST_HTML = (HERE / "fixtures" / "list_page1.html").read_text(encoding="utf-8")
DETAIL_HTML = (HERE / "fixtures" / "detail.html").read_text(encoding="utf-8")
PAGE1 = "https://books.toscrape.com/catalogue/page-1.html"
DETAIL = "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"


def test_parse_list_page():
    links, nxt = s.parse_list_page(LIST_HTML, PAGE1)
    assert len(links) == 20
    assert links[0] == DETAIL
    assert nxt == "https://books.toscrape.com/catalogue/page-2.html"


def test_parse_detail_page():
    book = s.parse_detail_page(DETAIL_HTML, DETAIL)
    assert book.title == "A Light in the Attic"
    assert book.category == "Poetry"
    assert book.price == pytest.approx(51.77)
    assert book.rating == 3
    assert book.in_stock == 22
    assert book.upc == "a897fe39b1053632"
    assert book.description.startswith("It's hard to imagine")
    assert book.image_url.startswith("https://books.toscrape.com/media/")


def test_parse_price():
    assert s.parse_price("£1,051.77") == pytest.approx(1051.77)
    assert s.parse_price("n/a") is None


class FakeFetcher(s.Fetcher):
    """Serves page 1 then a last page; every detail URL returns the fixture (one is broken)."""

    def __init__(self):
        super().__init__(delay=0)

    def get(self, url):
        self.requests_made += 1
        if url == PAGE1:
            return LIST_HTML
        if url.endswith("page-2.html"):
            return LIST_HTML.replace('<li class="next">', '<li class="gone">')
        if "/olio_984/" in url:
            return "<html>broken</html>"
        return DETAIL_HTML


def test_crawl_and_report(tmp_path):
    books, errors = s.crawl(FakeFetcher(), PAGE1, log=lambda *_: None)
    # 2 list pages x same 20 links -> 20 unique detail pages, 1 broken
    assert len(books) == 19
    assert len(errors) == 1 and "/olio_984/" in errors[0][0]

    out = tmp_path / "books.xlsx"
    s.write_excel(books, errors, out)
    sheets = pd.read_excel(out, sheet_name=None)
    assert set(sheets) == {"Overview", "Books", "By category", "Top rated (cheapest)", "Errors"}
    assert len(sheets["Books"]) == 19
    assert sheets["By category"].iloc[0]["books"] == 19


def test_cache(tmp_path, monkeypatch):
    f = s.Fetcher(delay=0, cache_dir=tmp_path)
    calls = []

    class Resp:
        text, encoding = "<html>ok</html>", None
        def raise_for_status(self):
            pass

    monkeypatch.setattr(f.session, "get", lambda url, timeout: calls.append(url) or Resp())
    assert f.get("https://example.com/a") == "<html>ok</html>"
    assert f.get("https://example.com/a") == "<html>ok</html>"
    assert len(calls) == 1
