#!/usr/bin/env python3
"""Update prices and view counts in SREALITY.xlsx from sreality.cz listings."""

import copy
import json
import re
import sys
import time
import unicodedata
from datetime import date, datetime, timezone
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

WORKBOOK = Path(__file__).resolve().parent / "SREALITY.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
)
NEXT_DATA_RE = re.compile(
    r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', re.S
)
TIMEOUT = 30
MAX_ATTEMPTS = 3
RETRY_DELAY = 2
REQUEST_DELAY = 0.5

GREEN_FILL = PatternFill(fill_type="solid", fgColor="FFC6EFCE")
RED_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
GREEN_FONT = "FF006100"
RED_FONT = "FF9C0006"


class FetchError(RuntimeError):
    pass


def as_int(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).replace("\u00a0", "").replace(" ", "").replace(",", "")
    return int(text) if re.fullmatch(r"-?\d+", text) else None


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def is_url(value):
    return isinstance(value, str) and value.strip().lower().startswith("https://")


def normalize(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(c for c in text if not unicodedata.combining(c)).strip().lower()


def find_column(sheet, header):
    for cell in sheet[1]:
        if normalize(cell.value) == normalize(header):
            return cell.column
    raise RuntimeError(f"Column {header!r} not found in sheet {sheet.title!r}")


def parse_listing(html, url):
    match = NEXT_DATA_RE.search(html)
    if not match:
        raise FetchError("__NEXT_DATA__ not found")
    try:
        next_data = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        raise FetchError(f"invalid __NEXT_DATA__: {exc}") from exc

    queries = (
        next_data.get("props", {})
        .get("pageProps", {})
        .get("dehydratedState", {})
        .get("queries", [])
    )
    url_id = re.search(r"/(\d+)\s*$", url)
    estate = None
    for query in queries:
        key = query.get("queryKey") or []
        if not key or key[0] != "estate":
            continue
        payload = (query.get("state") or {}).get("data")
        if not isinstance(payload, dict):
            continue
        if url_id and isinstance(key[1], dict) and str(key[1].get("id")) == url_id.group(1):
            estate = payload
            break
        if estate is None:
            estate = payload
    if estate is None:
        raise FetchError("estate data not found")

    price = None
    for key in ("price", "priceCzk", "priceSummaryCzk"):
        price = as_int(estate.get(key))
        if price is not None:
            break
    views = as_int((estate.get("params") or {}).get("stats"))
    return {"price": price, "views": views}


def fetch_listing(session, url):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = session.get(url, timeout=TIMEOUT)
            if response.status_code == 200:
                return parse_listing(response.text, url)
            if response.status_code in (404, 410):
                raise FetchError(f"HTTP {response.status_code} (listing unavailable)")
            error = f"HTTP {response.status_code}"
        except requests.RequestException as exc:
            error = str(exc)
        if attempt < MAX_ATTEMPTS:
            time.sleep(RETRY_DELAY * attempt)
    raise FetchError(f"failed after {MAX_ATTEMPTS} attempts: {error}")


def find_history_row(sheet, today):
    for row in range(3, sheet.max_row + 1):
        if as_date(sheet.cell(row, 1).value) == today:
            return row
    return None


def next_history_row(sheet):
    last = 2
    for row in range(3, sheet.max_row + 1):
        if any(sheet.cell(row, col).value is not None for col in range(1, sheet.max_column + 1)):
            last = row
    return last + 1


def set_price_style(cell, direction):
    if direction == "down":
        cell.fill = GREEN_FILL
        color = GREEN_FONT
    elif direction == "up":
        cell.fill = RED_FILL
        color = RED_FONT
    else:
        cell.fill = PatternFill(fill_type=None)
        color = None
    font = copy.copy(cell.font)
    font.color = color
    cell.font = font


def main():
    workbook_path = Path(sys.argv[1]) if len(sys.argv) > 1 else WORKBOOK
    if not workbook_path.is_file():
        raise SystemExit(f"Workbook not found: {workbook_path}")

    workbook = load_workbook(workbook_path)
    prehled = workbook["prehled"]
    historicky = workbook["historicky"]

    price_col = find_column(prehled, "cena")
    views_col = find_column(prehled, "shlednuti")
    link_col = find_column(prehled, "odkaz")

    prehled_rows = []
    for row in range(2, prehled.max_row + 1):
        url = prehled.cell(row, link_col).value
        if is_url(url):
            prehled_rows.append((row, url.strip()))

    history_cols = []
    for col in range(2, historicky.max_column + 1):
        url = historicky.cell(2, col).value
        if is_url(url):
            history_cols.append((col, url.strip()))

    urls = sorted({url for _, url in prehled_rows} | {url for _, url in history_cols})
    print(f"Fetching {len(urls)} unique listings...", flush=True)

    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "cs,en;q=0.9"})
    results = {}
    for index, url in enumerate(urls, 1):
        try:
            results[url] = fetch_listing(session, url)
            print(f"[{index}/{len(urls)}] OK   {url} -> {results[url]}", flush=True)
        except FetchError as exc:
            results[url] = {"price": None, "views": None}
            print(f"[{index}/{len(urls)}] FAIL {url}: {exc}", flush=True)
        if index < len(urls):
            time.sleep(REQUEST_DELAY)

    updated_prices = updated_views = 0
    for row, url in prehled_rows:
        result = results[url]
        if result["price"] is not None:
            prehled.cell(row, price_col).value = result["price"]
            updated_prices += 1
        if result["views"] is not None:
            prehled.cell(row, views_col).value = result["views"]
            updated_views += 1

    today = datetime.now(timezone.utc).date()
    history_row = find_history_row(historicky, today) or next_history_row(historicky)
    date_cell = historicky.cell(history_row, 1)
    date_cell.value = today
    date_cell.number_format = "dd.mm.yyyy"

    colored_down = colored_up = 0
    for col, url in history_cols:
        price = results[url]["price"]
        cell = historicky.cell(history_row, col)
        cell.value = price

        previous = None
        if history_row > 3:
            previous = as_int(historicky.cell(history_row - 1, col).value)
        direction = None
        if price is not None and previous is not None:
            if price < previous:
                direction = "down"
                colored_down += 1
            elif price > previous:
                direction = "up"
                colored_up += 1
        set_price_style(cell, direction)

    workbook.save(workbook_path)
    print(
        f"Updated {updated_prices} prices and {updated_views} view counts in 'prehled'; "
        f"wrote row {history_row} ({today:%d.%m.%Y}) in 'historicky' "
        f"({colored_down} down/green, {colored_up} up/red).",
        flush=True,
    )
    print(f"Saved {workbook_path}", flush=True)


if __name__ == "__main__":
    main()
