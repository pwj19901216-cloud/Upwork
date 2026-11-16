#!/usr/bin/env python3
"""Scrape a product catalogue into a clean Excel report.

Target: https://books.toscrape.com, a public sandbox site built for scraping practice.
The same structure (list pages -> detail pages -> clean table) applies to most
e-commerce, directory and listings sites.

Usage:
    python scrape_books.py -o output/books.xlsx
    python scrape_books.py -o output/books.xlsx --max-pages 2 --delay 1
"""
from __future__ import annotations

import argparse
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

START_URL = "https://books.toscrape.com/catalogue/page-1.html"
USER_AGENT = "Mozilla/5.0 (compatible; portfolio-scraper/1.0)"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


@dataclass
class Book:
    title: str
    category: str
    price: float | None
    rating: int | None
    in_stock: int
    upc: str
    description: str
    url: str
    image_url: str


# ---------------------------------------------------------------- fetching


class Fetcher:
    """HTTP session with retries, a polite delay shared across threads, and an optional HTML cache."""

    def __init__(self, delay: float = 0.5, cache_dir: Path | None = None, timeout: float = 20):
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT
        retry = Retry(total=4, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
        self.session.mount("https://", HTTPAdapter(max_retries=retry, pool_maxsize=16))
        self.session.mount("http://", HTTPAdapter(max_retries=retry, pool_maxsize=16))
        self.delay = delay
        self.timeout = timeout
        self.cache_dir = cache_dir
        self._lock = threading.Lock()
        self._next_time = 0.0
        self.requests_made = 0

    def _cache_path(self, url: str) -> Path | None:
        if not self.cache_dir:
            return None
        name = re.sub(r"[^A-Za-z0-9._-]+", "_", url.split("://", 1)[-1])
        return self.cache_dir / f"{name}.html"

    def _wait_turn(self) -> None:
        with self._lock:
            now = time.monotonic()
            wait = self._next_time - now
            self._next_time = max(now, self._next_time) + self.delay
        if wait > 0:
            time.sleep(wait)

    def get(self, url: str) -> str:
        path = self._cache_path(url)
        if path and path.exists():
            return path.read_text(encoding="utf-8")
        self._wait_turn()
        resp = self.session.get(url, timeout=self.timeout)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        html = resp.text
        with self._lock:
            self.requests_made += 1
        if path:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(html, encoding="utf-8")
        return html


# ---------------------------------------------------------------- parsing


def parse_list_page(html: str, page_url: str) -> tuple[list[str], str | None]:
    """Return (detail page URLs, next page URL or None)."""
    soup = BeautifulSoup(html, "html.parser")
    links = [urljoin(page_url, a["href"]) for a in soup.select("article.product_pod h3 a[href]")]
    nxt = soup.select_one("li.next a[href]")
    return links, (urljoin(page_url, nxt["href"]) if nxt else None)


def parse_price(text: str) -> float | None:
    match = re.search(r"\d+(?:[.,]\d+)?", text.replace(",", ""))
    return float(match.group()) if match else None


def parse_detail_page(html: str, url: str) -> Book:
    soup = BeautifulSoup(html, "html.parser")
    main = soup.select_one("div.product_main")
    if main is None:
        raise ValueError("product block not found")

    title = main.h1.get_text(strip=True)
    crumbs = [li.get_text(strip=True) for li in soup.select("ul.breadcrumb li")]
    category = crumbs[2] if len(crumbs) > 3 else ""

    rating_el = main.select_one("p.star-rating")
    rating = None
    if rating_el:
        rating = next((RATINGS[c] for c in rating_el.get("class", []) if c in RATINGS), None)

    table = {
        row.th.get_text(strip=True): row.td.get_text(strip=True)
        for row in soup.select("table.table-striped tr")
        if row.th and row.td
    }
    stock_match = re.search(r"\((\d+) available\)", table.get("Availability", ""))

    desc_el = soup.select_one("#product_description ~ p")
    img = soup.select_one("#product_gallery img[src]")

    return Book(
        title=title,
        category=category,
        price=parse_price(table.get("Price (incl. tax)", "") or main.select_one("p.price_color").get_text()),
        rating=rating,
        in_stock=int(stock_match.group(1)) if stock_match else 0,
        upc=table.get("UPC", ""),
        description=desc_el.get_text(strip=True) if desc_el else "",
        url=url,
        image_url=urljoin(url, img["src"]) if img else "",
    )


# ---------------------------------------------------------------- crawling


def crawl(fetcher: Fetcher, start_url: str = START_URL, max_pages: int | None = None,
          workers: int = 4, log=print) -> tuple[list[Book], list[tuple[str, str]]]:
    detail_urls: list[str] = []
    page_url: str | None = start_url
    pages = 0
    while page_url and (max_pages is None or pages < max_pages):
        links, page_url = parse_list_page(fetcher.get(page_url), page_url)
        detail_urls.extend(links)
        pages += 1
        log(f"list page {pages}: {len(links)} items")
    detail_urls = list(dict.fromkeys(detail_urls))  # de-duplicate, keep order

    books: dict[str, Book] = {}
    errors: list[tuple[str, str]] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(fetcher.get, u): u for u in detail_urls}
        for i, fut in enumerate(as_completed(futures), 1):
            url = futures[fut]
            try:
                books[url] = parse_detail_page(fut.result(), url)
            except Exception as exc:  # keep going; report at the end
                errors.append((url, str(exc)))
            if i % 100 == 0 or i == len(futures):
                log(f"detail pages: {i}/{len(futures)}")
    ordered = [books[u] for u in detail_urls if u in books]
    return ordered, errors


# ---------------------------------------------------------------- report

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
LINK_FONT = Font(color="0563C1", underline="single")


def _style(ws, widths: dict[str, int], formats: dict[str, str]) -> None:
    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    for idx, name in enumerate(headers, 1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = widths.get(name, 14)
        if name in formats:
            for (cell,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                cell.number_format = formats[name]


def build_frames(books: list[Book]) -> dict[str, pd.DataFrame]:
    df = pd.DataFrame([asdict(b) for b in books],
                      columns=list(Book.__dataclass_fields__))
    by_cat = (
        df.groupby("category")
        .agg(books=("title", "count"), avg_price=("price", "mean"), min_price=("price", "min"),
             max_price=("price", "max"), avg_rating=("rating", "mean"), units_in_stock=("in_stock", "sum"))
        .reset_index()
        .sort_values("books", ascending=False)
    ) if len(df) else pd.DataFrame()
    top = df.sort_values(["rating", "price"], ascending=[False, True]).head(20)[
        ["title", "category", "rating", "price", "in_stock", "url"]] if len(df) else pd.DataFrame()
    return {"Books": df, "By category": by_cat, "Top rated (cheapest)": top}


def write_excel(books: list[Book], errors: list[tuple[str, str]], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    frames = build_frames(books)
    df = frames["Books"]
    overview = pd.DataFrame(
        [
            ("Books scraped", len(df)),
            ("Categories", df["category"].nunique() if len(df) else 0),
            ("Average price (GBP)", round(df["price"].mean(), 2) if len(df) else 0),
            ("Total units in stock", int(df["in_stock"].sum()) if len(df) else 0),
            ("Pages failed", len(errors)),
            ("Source", "https://books.toscrape.com (scraping sandbox)"),
            ("Scraped at (UTC)", time.strftime("%Y-%m-%d %H:%M", time.gmtime())),
        ],
        columns=["metric", "value"],
    )
    sheets = [("Overview", overview), *frames.items()]
    if errors:
        sheets.append(("Errors", pd.DataFrame(errors, columns=["url", "error"])))

    money = "£#,##0.00"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, frame in sheets:
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            _style(
                ws,
                widths={"title": 50, "category": 20, "description": 60, "url": 40, "image_url": 40,
                        "upc": 18, "metric": 24, "value": 44},
                formats={"price": money, "avg_price": money, "min_price": money, "max_price": money,
                         "avg_rating": "0.00"},
            )
            headers = [c.value for c in ws[1]]
            if "url" in headers:
                col = headers.index("url") + 1
                for (cell,) in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                    if cell.value:
                        cell.hyperlink, cell.font = cell.value, LINK_FONT
            if "description" in headers:
                col = headers.index("description") + 1
                for (cell,) in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                    if isinstance(cell.value, str) and len(cell.value) > 300:
                        cell.value = cell.value[:297] + "..."


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Scrape books.toscrape.com into an Excel report.")
    p.add_argument("-o", "--output", type=Path, default=Path("output/books.xlsx"))
    p.add_argument("--csv", type=Path, help="also save the full table as CSV")
    p.add_argument("--max-pages", type=int, help="stop after N list pages (20 books each)")
    p.add_argument("--delay", type=float, default=0.5, help="seconds between requests (default 0.5)")
    p.add_argument("--workers", type=int, default=4, help="parallel downloads (default 4)")
    p.add_argument("--cache", type=Path, help="folder to cache HTML so re-runs don't hit the site again")
    p.add_argument("--start-url", default=START_URL)
    args = p.parse_args(argv)

    fetcher = Fetcher(delay=args.delay, cache_dir=args.cache)
    started = time.monotonic()
    books, errors = crawl(fetcher, args.start_url, args.max_pages, args.workers)
    write_excel(books, errors, args.output)
    if args.csv:
        args.csv.parent.mkdir(parents=True, exist_ok=True)
        build_frames(books)["Books"].to_csv(args.csv, index=False, encoding="utf-8-sig")

    print(f"Scraped {len(books)} books with {fetcher.requests_made} requests "
          f"in {time.monotonic() - started:.0f}s; {len(errors)} failed")
    print(f"Wrote {args.output}" + (f" and {args.csv}" if args.csv else ""))
    for url, err in errors[:10]:
        print(f"  failed: {url}: {err}", file=sys.stderr)
    return 0 if books else 1


if __name__ == "__main__":
    raise SystemExit(main())
