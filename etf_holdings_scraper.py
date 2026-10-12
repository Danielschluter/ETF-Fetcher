#!/usr/bin/env python3
"""Scrape the latest holdings of every US-listed ETF.

Holdings come from the fund issuers' own websites, which publish them daily,
plus a catch-all source for every ETF without a dedicated issuer scraper:

  ishares      iShares holdings CSV (BlackRock)
  ssga         SPDR daily holdings workbook (State Street)
  invesco      Invesco holdings API
  first_trust  First Trust holdings page
  global_x     Global X daily holdings CSV
  zacks        Zacks ETF holdings page: used for all other ETFs, and as the
               fallback whenever an issuer source fails. Zacks does not say
               which date its holdings are as of, so as_of_date is left blank.

The ETF universe is every security flagged as an ETF in the Nasdaq Trader
symbol directory, which covers all US exchanges (NYSE, Arca, Nasdaq, Cboe...).

Outputs (in --output-dir)
  holdings.csv     one row per position, all ETFs combined
  etf_summary.csv  one row per ETF: source used, as-of date, number of holdings,
                   total weight, or why no holdings were found

Examples
  python etf_holdings_scraper.py
  python etf_holdings_scraper.py --tickers SPY IVV QQQ
  python etf_holdings_scraper.py --sources ishares ssga --no-fallback
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import gzip
import hashlib
import html
import io
import json
import logging
import os
import re
import sys
import threading
import time
import warnings
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Sequence, TypeVar
from urllib.parse import urlsplit

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

log = logging.getLogger("etf_holdings")
T = TypeVar("T")

# SSGA workbooks have no default style, which openpyxl warns about on every file.
warnings.filterwarnings("ignore", message="Workbook contains no default style", module="openpyxl")

NASDAQ_SYMBOL_DIRECTORY = "https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqtraded.txt"

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
)
HOLDINGS_TTL = 6 * 3600  # issuers publish once per business day
REFERENCE_TTL = 24 * 3600

EXCHANGES = {
    "A": "NYSE American",
    "N": "NYSE",
    "P": "NYSE Arca",
    "Q": "Nasdaq",
    "V": "IEX",
    "Z": "Cboe BZX",
}

HOLDING_FIELDS = [
    "holding_name",
    "ticker",
    "cusip",
    "isin",
    "sedol",
    "other_id",
    "asset_class",
    "sector",
    "country",
    "currency",
    "shares",
    "price",
    "market_value",
    "weight_pct",
    "maturity_date",
    "coupon_rate",
]
HOLDING_COLUMNS = ["etf_ticker", "etf_name", "source", "as_of_date", *HOLDING_FIELDS]
NUMERIC_FIELDS = {"shares", "price", "market_value", "weight_pct", "coupon_rate"}

SUMMARY_COLUMNS = [
    "ticker",
    "name",
    "exchange",
    "status",
    "source",
    "as_of_date",
    "holdings_count",
    "weight_sum_pct",
    "source_url",
    "detail",
]


# --------------------------------------------------------------------------- HTTP


class NotAvailable(Exception):
    """A source has no holdings for this fund; the next source should be tried."""


class NotFound(NotAvailable):
    """HTTP 404/410."""


class RateLimiter:
    """Spaces requests at least 1/rate seconds apart across all threads."""

    def __init__(self, rate: float):
        self.interval = 1.0 / rate
        self.lock = threading.Lock()
        self.next_slot = time.monotonic()

    def wait(self) -> None:
        with self.lock:
            now = time.monotonic()
            slot = max(now, self.next_slot)
            self.next_slot = slot + self.interval
        if slot > now:
            time.sleep(slot - now)


class HttpClient:
    """GET with a per-host rate limit and retries; one requests.Session per thread."""

    RETRY_STATUSES = {429, 500, 502, 503, 504}

    def __init__(self, user_agent: str, rate: float, timeout: float = 60, retries: int = 3):
        self.headers = {
            "User-Agent": user_agent,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        self.rate = rate
        self.timeout = timeout
        self.retries = retries
        self.local = threading.local()
        self.limiters: dict[str, RateLimiter] = {}
        self.limiters_lock = threading.Lock()

    def _session(self) -> requests.Session:
        session = getattr(self.local, "session", None)
        if session is None:
            session = requests.Session()
            session.headers.update(self.headers)
            self.local.session = session
        return session

    def _limiter(self, url: str) -> RateLimiter:
        host = urlsplit(url).hostname or ""
        with self.limiters_lock:
            if host not in self.limiters:
                self.limiters[host] = RateLimiter(self.rate)
            return self.limiters[host]

    def get(self, url: str) -> bytes:
        limiter = self._limiter(url)
        for attempt in range(self.retries + 1):
            limiter.wait()
            delay = 2.0 ** (attempt + 1)
            try:
                resp = self._session().get(url, timeout=self.timeout)
            except requests.RequestException as exc:
                error: Exception = exc
            else:
                if resp.status_code in (404, 410):
                    raise NotFound(f"HTTP {resp.status_code}")
                if resp.status_code not in self.RETRY_STATUSES:
                    resp.raise_for_status()
                    return resp.content
                error = requests.HTTPError(f"HTTP {resp.status_code}", response=resp)
                retry_after = resp.headers.get("Retry-After", "")
                if retry_after.isdigit():
                    delay = max(delay, float(retry_after))
            if attempt == self.retries:
                raise error
            log.debug("Retrying %s in %.0fs after: %s", url, delay, error)
            time.sleep(delay)
        raise AssertionError("unreachable")


class DiskCache:
    """Gzipped file cache. Writes are atomic so an interrupted run never leaves a bad entry."""

    def __init__(self, root: Path | None):
        self.root = root

    def _path(self, url: str) -> Path:
        host = urlsplit(url).hostname or "other"
        return self.root / host / f"{hashlib.sha1(url.encode()).hexdigest()}.gz"

    def read(self, url: str, max_age: float) -> bytes | None:
        if self.root is None:
            return None
        path = self._path(url)
        try:
            if time.time() - path.stat().st_mtime > max_age:
                return None
            return gzip.decompress(path.read_bytes())
        except (OSError, EOFError):
            return None

    def write(self, url: str, data: bytes) -> None:
        if self.root is None:
            return
        path = self._path(url)
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(f"{path.name}.{os.getpid()}.{threading.get_ident()}.tmp")
        tmp.write_bytes(gzip.compress(data, compresslevel=6))
        os.replace(tmp, path)

    def delete(self, url: str) -> None:
        if self.root is not None:
            self._path(url).unlink(missing_ok=True)


class Fetcher:
    """HTTP access with a cache in front of it, so reruns and resumed runs are fast."""

    def __init__(self, http: HttpClient, cache: DiskCache, refresh: bool = False):
        self.http = http
        self.cache = cache
        self.refresh = refresh

    def get(self, url: str, max_age: float = HOLDINGS_TTL) -> bytes:
        if not self.refresh:
            cached = self.cache.read(url, max_age)
            if cached is not None:
                return cached
        data = self.http.get(url)
        self.cache.write(url, data)
        return data

    def forget(self, url: str) -> None:
        """Drop a cached response, e.g. an error page that should not be served again."""
        self.cache.delete(url)


# ----------------------------------------------------------------------- Universe


@dataclass(frozen=True)
class Etf:
    ticker: str
    name: str
    exchange: str


def parse_nasdaq_directory(text: str) -> list[Etf]:
    """ETFs from nasdaqtraded.txt (pipe-delimited, one row per US-listed security)."""
    etfs = []
    for row in csv.DictReader(io.StringIO(text), delimiter="|"):
        if row.get("ETF") != "Y" or row.get("Test Issue") == "Y":
            continue
        ticker = (row.get("Symbol") or "").strip()
        if not ticker:
            continue
        code = (row.get("Listing Exchange") or "").strip()
        name = (row.get("Security Name") or "").strip()
        etfs.append(Etf(ticker, name, EXCHANGES.get(code, code)))
    return sorted(etfs, key=lambda e: e.ticker)


def load_universe(fetcher: Fetcher) -> list[Etf]:
    raw = fetcher.get(NASDAQ_SYMBOL_DIRECTORY, max_age=REFERENCE_TTL)
    etfs = parse_nasdaq_directory(raw.decode("utf-8", errors="replace"))
    if not etfs:
        raise RuntimeError(f"No ETFs found in {NASDAQ_SYMBOL_DIRECTORY}; has its format changed?")
    return etfs


# ------------------------------------------------------------- Holdings tables

_MISSING = {"", "-", "--", "na", "n/a", "nan", "none", "null"}


def clean_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    text = html.unescape(str(value)).replace("\xa0", " ").strip()
    return "" if text.lower() in _MISSING else text


def parse_number(value: object) -> float | str:
    """'1,234.5', '$1,234.50', '2.66%', '(12)' -> float; blanks and dashes -> ''."""
    if isinstance(value, bool) or value is None:
        return ""
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace(",", "").replace("$", "").replace("%", "").strip()
    negative = text.startswith("(") and text.endswith(")")
    try:
        number = float(text.strip("()"))
    except ValueError:
        return ""
    return -number if negative else number


_DATE_FORMATS = [
    (re.compile(r"\b\d{4}-\d{2}-\d{2}\b"), ("%Y-%m-%d",)),
    (re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b"), ("%m/%d/%Y",)),
    (re.compile(r"\b\d{1,2}-[A-Za-z]{3}-\d{4}\b"), ("%d-%b-%Y",)),
    (re.compile(r"\b[A-Za-z]{3,9}\.? \d{1,2}, \d{4}\b"), ("%b %d, %Y", "%B %d, %Y")),
    (re.compile(r"\b\d{1,2} [A-Za-z]{3,9} \d{4}\b"), ("%d %b %Y", "%d %B %Y")),
]


def parse_date(text: str) -> str:
    """First recognizable date in text, as YYYY-MM-DD ('' if none)."""
    for pattern, formats in _DATE_FORMATS:
        for match in pattern.finditer(text):
            candidate = match.group(0).replace(".", "")
            for fmt in formats:
                try:
                    return dt.datetime.strptime(candidate, fmt).date().isoformat()
                except ValueError:
                    pass
    return ""


def find_as_of(text: str) -> str:
    """The date following 'as of' in text, as YYYY-MM-DD ('' if none)."""
    for match in re.finditer(r"as of", text, re.IGNORECASE):
        found = parse_date(text[match.end():match.end() + 40])
        if found:
            return found
    return ""


# Issuer column headers, normalized by _header_key, mapped to output fields.
COLUMN_ALIASES = {
    "holding_name": ["name", "securityname", "security", "description", "securitydescription",
                     "holding", "holdingname", "issuername", "issuer"],
    "ticker": ["ticker", "symbol", "tickersymbol"],
    "cusip": ["cusip"],
    "isin": ["isin"],
    "sedol": ["sedol"],
    "other_id": ["identifier", "securityid"],
    "asset_class": ["assetclass", "securitytype", "assettype", "type"],
    "sector": ["sector", "gicssector", "industry"],
    "country": ["location", "country", "countryofrisk", "locationofrisk"],
    "currency": ["currency", "localcurrency", "marketcurrency", "basecurrency"],
    "shares": ["shares", "sharesheld", "quantity", "sharesquantity", "parvalue", "par", "units",
               "sharespar", "nominal", "principal"],
    "price": ["price", "marketprice", "closingprice"],
    "market_value": ["marketvalue", "marketvalueusd", "basemarketvalue", "value"],
    "weight_pct": ["weightpct", "weight", "weighting", "pctofnetassets", "percentofnetassets",
                   "pctnetassets", "netassetspct", "portfolioweight", "pctoffund"],
    "maturity_date": ["maturity", "maturitydate"],
    "coupon_rate": ["couponpct", "coupon", "couponrate"],
}
_ALIAS_TO_FIELD = {alias: f for f, aliases in COLUMN_ALIASES.items() for alias in aliases}


def _header_key(cell: object) -> str:
    """'Weight (%)' -> 'weightpct', 'Market Value ($)' -> 'marketvalue'."""
    return re.sub(r"[^a-z0-9]", "", clean_text(cell).lower().replace("%", "pct"))


def _map_header(cells: list[str], overrides: dict[str, str]) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for i, cell in enumerate(cells):
        key = _header_key(cell)
        target = overrides.get(key) or _ALIAS_TO_FIELD.get(key)
        if target and target not in mapping.values():
            mapping[i] = target
    return mapping


def _is_header(mapping: dict[int, str]) -> bool:
    fields = set(mapping.values())
    return bool(fields & {"holding_name", "ticker"}) and bool(
        fields & {"weight_pct", "shares", "market_value"}
    )


def normalize_holding(raw: dict) -> dict:
    holding = {}
    for name in HOLDING_FIELDS:
        value = raw.get(name)
        holding[name] = parse_number(value) if name in NUMERIC_FIELDS else clean_text(value)
    holding["maturity_date"] = parse_date(holding["maturity_date"]) or holding["maturity_date"]
    return holding


def rows_to_holdings(
    rows: Iterable[Sequence[object]], overrides: dict[str, str] | None = None
) -> tuple[list[dict], str]:
    """Holdings from a spreadsheet-like table, plus the text found above its header.

    Issuer files put a few lines of fund details (name, as-of date) above the
    table and disclaimers below it. The header row is the first row whose
    columns name a security and a weight, quantity or value; the table ends at
    the first row with fewer than two filled cells. ``overrides`` remaps
    ambiguous headers, e.g. SSGA's "Identifier" column holds CUSIPs.
    """
    overrides = overrides or {}
    preamble: list[str] = []
    mapping: dict[int, str] | None = None
    holdings = []
    for row in rows:
        cells = [clean_text(c) for c in row]
        if mapping is None:
            candidate = _map_header(cells, overrides)
            if _is_header(candidate):
                mapping = candidate
            else:
                preamble.append(" ".join(c for c in cells if c))
            continue
        if sum(1 for c in cells if c) < 2:
            break
        holdings.append(normalize_holding(
            {name: cells[i] for i, name in mapping.items() if i < len(cells)}
        ))
    if mapping is None:
        raise NotAvailable("no holdings table found")
    return holdings, "\n".join(preamble)


def read_csv_rows(data: bytes) -> list[list[str]]:
    return list(csv.reader(io.StringIO(data.decode("utf-8-sig", errors="replace"))))


def read_xlsx_rows(data: bytes) -> list[list[object]]:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        return [list(row) for row in workbook.worksheets[0].iter_rows(values_only=True)]
    finally:
        workbook.close()


# ---------------------------------------------------------------------- Sources


@dataclass
class Holdings:
    source: str
    url: str
    as_of: str
    rows: list[dict]


class Source:
    name = ""

    def __init__(self, fetcher: Fetcher):
        self.fetcher = fetcher

    def prepare(self) -> None:
        """One-time setup before scraping starts, e.g. loading a fund list."""

    def handles(self, etf: Etf) -> bool:
        raise NotImplementedError

    def fetch(self, etf: Etf) -> Holdings:
        raise NotImplementedError

    def download(self, url: str, parse: Callable[[bytes], T]) -> T:
        """Fetch url and parse it. A response that can't be parsed is usually an error
        or bot-check page: it is dropped from the cache, and the error says what it was."""
        data = self.fetcher.get(url)
        try:
            return parse(data)
        except Exception as exc:
            self.fetcher.forget(url)
            reason = str(exc) if isinstance(exc, NotAvailable) else f"{type(exc).__name__}: {exc}"
            raise NotAvailable(f"{reason}; {describe_response(data)}") from None


def describe_response(data: bytes) -> str:
    """'got an HTML page titled ...' or the first line of an unexpected response."""
    text = data[:50_000].decode("utf-8", errors="replace")
    title = re.search(r"<title[^>]*>(.*?)</title>", text, re.IGNORECASE | re.DOTALL)
    if title:
        return f"got a {len(data):,}-byte HTML page titled {clean_text(title.group(1))[:80]!r}"
    first = next((line.strip() for line in text.splitlines() if line.strip()), "")
    return f"got {len(data):,} bytes starting {first[:80]!r}"


class IsharesSource(Source):
    name = "ishares"
    FUND_LIST = (
        "https://www.ishares.com/us/product-screener/product-screener-v3.1.jsn?"
        "dcrPath=/templatedata/config/product-screener-v3/data/en/us-ishares/"
        "ishares-product-screener-backend-config&siteEntryPassthrough=true"
    )
    # 1467271812596 is the id of the holdings download component on every fund page.
    HOLDINGS = ("https://www.ishares.com{page}/1467271812596.ajax"
                "?fileType=csv&fileName={ticker}_holdings&dataType=fund")

    def __init__(self, fetcher: Fetcher):
        super().__init__(fetcher)
        self.pages: dict[str, str] = {}

    def prepare(self) -> None:
        try:
            funds = json.loads(self.fetcher.get(self.FUND_LIST, max_age=REFERENCE_TTL))
        except Exception as exc:
            log.warning("iShares fund list unavailable (%s); iShares ETFs will use the "
                        "fallback source", exc)
            return
        self.pages = parse_ishares_fund_list(funds)
        log.info("iShares: %d funds", len(self.pages))

    def handles(self, etf: Etf) -> bool:
        return etf.ticker in self.pages

    def fetch(self, etf: Etf) -> Holdings:
        url = self.HOLDINGS.format(page=self.pages[etf.ticker], ticker=etf.ticker)
        holdings, preamble = self.download(url, lambda d: rows_to_holdings(read_csv_rows(d)))
        return Holdings(self.name, url, find_as_of(preamble), holdings)


def parse_ishares_fund_list(funds: dict) -> dict[str, str]:
    """{product id: {...}} from the iShares product screener -> {ticker: product page path}."""
    pages = {}
    for fund in funds.values():
        ticker = (fund.get("localExchangeTicker") or "").strip().upper()
        page = (fund.get("productPageUrl") or "").strip()
        if ticker and page.startswith("/"):
            pages[ticker] = page.rstrip("/")
    return pages


class SsgaSource(Source):
    name = "ssga"
    HOLDINGS = ("https://www.ssga.com/library-content/products/fund-data/etfs/us/"
                "holdings-daily-us-en-{ticker}.xlsx")
    NAME = re.compile(r"\bSPDR\b|^State Street\b", re.IGNORECASE)

    def handles(self, etf: Etf) -> bool:
        return bool(self.NAME.search(etf.name))

    def fetch(self, etf: Etf) -> Holdings:
        url = self.HOLDINGS.format(ticker=etf.ticker.lower())
        holdings, preamble = self.download(
            url, lambda d: rows_to_holdings(read_xlsx_rows(d), {"identifier": "cusip"}))
        return Holdings(self.name, url, find_as_of(preamble), holdings)


class InvescoSource(Source):
    name = "invesco"
    HOLDINGS = ("https://dng-api.invesco.com/cache/v1/accounts/en_US/shareclasses/{ticker}"
                "/holdings/fund?idType=ticker&interval=monthly&productType=ETF")

    def handles(self, etf: Etf) -> bool:
        return etf.name.lower().startswith("invesco")

    def fetch(self, etf: Etf) -> Holdings:
        url = self.HOLDINGS.format(ticker=etf.ticker)
        as_of, holdings = self.download(url, lambda d: parse_invesco(json.loads(d)))
        return Holdings(self.name, url, as_of, holdings)


def parse_invesco(data: dict) -> tuple[str, list[dict]]:
    holdings = [
        normalize_holding({
            "holding_name": h.get("issuerName") or h.get("securityName"),
            "ticker": h.get("ticker"),
            "cusip": h.get("cusip"),
            "isin": h.get("isin"),
            "sedol": h.get("sedol"),
            "asset_class": h.get("securityTypeName"),
            "sector": h.get("sectorName"),
            "country": h.get("countryName"),
            "currency": h.get("currency"),
            "shares": h.get("units"),
            "price": h.get("price"),
            "market_value": h.get("marketValue"),
            "weight_pct": h.get("percentageOfTotalNetAssets"),
            "maturity_date": h.get("maturityDate"),
            "coupon_rate": h.get("couponRate"),
        })
        for h in data.get("holdings") or []
    ]
    as_of = parse_date(str(data.get("effectiveBusinessDate") or data.get("effectiveDate") or ""))
    return as_of, holdings


class FirstTrustSource(Source):
    name = "first_trust"
    HOLDINGS = "https://www.ftportfolios.com/Retail/Etf/EtfHoldings.aspx?Ticker={ticker}"
    NAME = re.compile(r"^(First Trust|FT Vest)\b", re.IGNORECASE)

    def handles(self, etf: Etf) -> bool:
        return bool(self.NAME.search(etf.name))

    def fetch(self, etf: Etf) -> Holdings:
        url = self.HOLDINGS.format(ticker=etf.ticker)
        as_of, holdings = self.download(
            url, lambda d: parse_first_trust(d.decode("utf-8", errors="replace")))
        return Holdings(self.name, url, as_of, holdings)


def parse_first_trust(page: str) -> tuple[str, list[dict]]:
    soup = BeautifulSoup(page, "html.parser")
    title = soup.find(id=re.compile(r"lblHoldingsTitle$"))
    as_of = find_as_of(title.get_text(" ", strip=True)) if title else ""
    holdings = []
    for table in soup.select("table.fundSilverGrid"):
        rows = [[cell.get_text(" ", strip=True) for cell in tr.find_all(["td", "th"])]
                for tr in table.find_all("tr")]
        try:
            # Here "Identifier" is the exchange ticker; CUSIPs have their own column.
            table_holdings, _ = rows_to_holdings(rows, {"identifier": "ticker"})
        except NotAvailable:
            continue  # e.g. a futures exposure summary table
        holdings.extend(table_holdings)
    if not holdings:
        raise NotAvailable("no holdings table found")
    return as_of, holdings


class GlobalXSource(Source):
    name = "global_x"
    HOLDINGS = "https://assets.globalxetfs.com/funds/holdings/{ticker}_full-holdings_{date:%Y%m%d}.csv"
    LOOKBACK_BUSINESS_DAYS = 5

    def __init__(self, fetcher: Fetcher, today: dt.date | None = None):
        super().__init__(fetcher)
        self.today = today

    def handles(self, etf: Etf) -> bool:
        return etf.name.lower().startswith("global x")

    def fetch(self, etf: Etf) -> Holdings:
        # Files are named by date, so look for the most recent one.
        for day in recent_business_days(self.today or dt.date.today(), self.LOOKBACK_BUSINESS_DAYS):
            url = self.HOLDINGS.format(ticker=etf.ticker.lower(), date=day)
            try:
                holdings, preamble = self.download(
                    url, lambda d: rows_to_holdings(read_csv_rows(d)))
            except NotFound:
                continue
            except requests.HTTPError as exc:
                # The file store answers 403 for a file that doesn't exist (yet).
                if exc.response is not None and exc.response.status_code == 403:
                    continue
                raise
            return Holdings(self.name, url, find_as_of(preamble) or day.isoformat(), holdings)
        raise NotAvailable(
            f"no holdings file in the last {self.LOOKBACK_BUSINESS_DAYS} business days")


def recent_business_days(today: dt.date, count: int) -> list[dt.date]:
    days, day = [], today
    while len(days) < count:
        if day.weekday() < 5:
            days.append(day)
        day -= dt.timedelta(days=1)
    return days


class ZacksSource(Source):
    name = "zacks"
    HOLDINGS = "https://www.zacks.com/funds/etf/{ticker}/holding"

    def handles(self, etf: Etf) -> bool:
        return True

    def fetch(self, etf: Etf) -> Holdings:
        url = self.HOLDINGS.format(ticker=etf.ticker)
        holdings = self.download(url, lambda d: parse_zacks(d.decode("utf-8", errors="replace")))
        return Holdings(self.name, url, "", holdings)


_ZACKS_DATA = re.compile(r"etf_holdings\.formatted_data\s*=\s*(\[.*?\])\s*;", re.DOTALL)
_TITLE_ATTR = re.compile(r'title="([^"]*)"')
_REL_ATTR = re.compile(r'rel="([^"]+)"')
_TAGS = re.compile(r"<[^>]+>")
_CUSIP = re.compile(r"^[0-9A-Z]{8}[0-9]$")


def parse_zacks(page: str) -> list[dict]:
    """Holdings from the JavaScript array embedded in a Zacks ETF holdings page.

    Each row is [name, symbol, shares, weight %, 52-week change, report link];
    long names are truncated in the text but complete in a title attribute.
    """
    match = _ZACKS_DATA.search(page)
    if not match:
        raise NotAvailable("no holdings on page")
    holdings = []
    for row in json.loads(match.group(1)):
        if len(row) < 4:
            continue
        name_cell, symbol_cell, shares, weight = (str(c) for c in row[:4])
        title = _TITLE_ATTR.search(name_cell)
        name = title.group(1) if title else _TAGS.sub("", name_cell)
        linked = _REL_ATTR.search(symbol_cell)
        symbol = clean_text(linked.group(1) if linked else _TAGS.sub("", symbol_cell))
        fields = {"holding_name": name, "shares": shares, "weight_pct": weight}
        if linked:
            fields["ticker"] = symbol
        elif _CUSIP.match(symbol):
            fields["cusip"] = symbol
        else:
            fields["other_id"] = symbol
        holdings.append(normalize_holding(fields))
    return holdings


ISSUER_SOURCES: dict[str, type[Source]] = {
    "ishares": IsharesSource,
    "ssga": SsgaSource,
    "invesco": InvescoSource,
    "first_trust": FirstTrustSource,
    "global_x": GlobalXSource,
}
FALLBACK_SOURCE = ZacksSource


# ---------------------------------------------------------------------- Scraper


class HoldingsScraper:
    """Tries the ETF's issuer source first, then the fallback, and keeps the first success."""

    def __init__(self, issuers: list[Source], fallback: Source | None):
        self.issuers = issuers
        self.fallback = fallback

    def sources_for(self, etf: Etf) -> list[Source]:
        chain = [s for s in self.issuers if s.handles(etf)]
        if self.fallback is not None:
            chain.append(self.fallback)
        return chain

    def scrape(self, etf: Etf) -> tuple[dict, list[dict]]:
        summary = {col: "" for col in SUMMARY_COLUMNS}
        summary.update(ticker=etf.ticker, name=etf.name, exchange=etf.exchange)
        failures = []
        for source in self.sources_for(etf):
            try:
                result = source.fetch(etf)
                if not result.rows:
                    raise NotAvailable("no holdings listed")
            except NotAvailable as exc:
                failures.append(f"{source.name}: {exc}")
                continue
            except Exception as exc:  # one bad page must not stop the run
                log.debug("%s failed for %s", source.name, etf.ticker, exc_info=True)
                failures.append(f"{source.name}: {type(exc).__name__}: {exc}")
                continue

            weights = [h["weight_pct"] for h in result.rows if isinstance(h["weight_pct"], float)]
            summary.update(
                status="ok",
                source=result.source,
                as_of_date=result.as_of,
                holdings_count=len(result.rows),
                weight_sum_pct=round(sum(weights), 2) if weights else "",
                source_url=result.url,
                detail="; ".join(failures),
            )
            prefix = {"etf_ticker": etf.ticker, "etf_name": etf.name,
                      "source": result.source, "as_of_date": result.as_of}
            return summary, [{**prefix, **h} for h in result.rows]

        summary.update(status="no_data",
                       detail="; ".join(failures) or "no source covers this ETF")
        return summary, []


# -------------------------------------------------------------------------- Main


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape the latest holdings of every US-listed ETF from issuer websites.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--tickers", nargs="+", metavar="TICKER",
                        help="Only these tickers (space or comma separated) instead of all ETFs")
    parser.add_argument("--limit", type=int, help="Only the first N ETFs, alphabetically")
    parser.add_argument("--sources", nargs="+", choices=list(ISSUER_SOURCES),
                        default=list(ISSUER_SOURCES), help="Issuer sources to use")
    parser.add_argument("--no-fallback", action="store_true",
                        help="Skip Zacks: only ETFs covered by an issuer source get holdings")
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument("--cache-dir", type=Path, default=Path(".cache"),
                        help="Downloads are kept here for a few hours so reruns resume quickly")
    parser.add_argument("--no-cache", action="store_true", help="Do not read or write the cache")
    parser.add_argument("--refresh", action="store_true", help="Ignore cached downloads")
    parser.add_argument("--workers", type=int, default=8, help="Parallel download threads")
    parser.add_argument("--rate", type=float, default=2.0,
                        help="Max requests per second to any one website")
    parser.add_argument("--user-agent", default=DEFAULT_USER_AGENT)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    if args.rate <= 0:
        parser.error("--rate must be positive")
    if args.workers < 1:
        parser.error("--workers must be at least 1")
    if args.tickers:
        args.tickers = [
            t.strip().upper() for arg in args.tickers for t in arg.split(",") if t.strip()]
    return args


def select_etfs(universe: list[Etf], tickers: list[str] | None, limit: int | None) -> list[Etf]:
    etfs = universe
    if tickers:
        by_ticker = {e.ticker: e for e in universe}
        missing = [t for t in tickers if t not in by_ticker]
        if missing:
            log.warning("Not in the US ETF list, trying anyway: %s", ", ".join(missing))
        etfs = [by_ticker.get(t, Etf(t, "", "")) for t in dict.fromkeys(tickers)]
    if limit:
        etfs = etfs[:limit]
    return etfs


def write_csv(path: Path, columns: list[str], rows: Iterable[dict]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def run(args: argparse.Namespace, http: HttpClient | None = None) -> int:
    http = http or HttpClient(args.user_agent, args.rate)
    fetcher = Fetcher(http, DiskCache(None if args.no_cache else args.cache_dir), args.refresh)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    universe = load_universe(fetcher)
    log.info("Found %d US-listed ETFs", len(universe))
    etfs = select_etfs(universe, args.tickers, args.limit)

    issuers = [ISSUER_SOURCES[name](fetcher) for name in args.sources]
    for source in issuers:
        source.prepare()
    scraper = HoldingsScraper(issuers, None if args.no_fallback else FALLBACK_SOURCE(fetcher))
    planned = Counter(scraper.sources_for(etf)[0].name if scraper.sources_for(etf) else "none"
                      for etf in etfs)
    log.info("Primary source per ETF: %s",
             ", ".join(f"{k}={v}" for k, v in planned.most_common()))

    holdings_path = args.output_dir / "holdings.csv"
    summary_path = args.output_dir / "etf_summary.csv"
    summaries: list[dict] = []
    statuses: Counter = Counter()
    holdings_total = 0
    interrupted = False
    started = time.monotonic()

    with holdings_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=HOLDING_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        pool = ThreadPoolExecutor(max_workers=args.workers)
        try:
            futures = [pool.submit(scraper.scrape, etf) for etf in etfs]
            for done, future in enumerate(as_completed(futures), 1):
                summary, rows = future.result()
                writer.writerows(rows)  # only this thread writes, so no lock is needed
                summaries.append(summary)
                statuses[summary["source"] or summary["status"]] += 1
                holdings_total += len(rows)
                if done % 50 == 0 or done == len(futures):
                    log.info("[%d/%d] %s | %d holdings | %.0fs", done, len(futures),
                             " ".join(f"{k}={v}" for k, v in sorted(statuses.items())),
                             holdings_total, time.monotonic() - started)
        except KeyboardInterrupt:
            interrupted = True
            log.warning("Interrupted; writing what has been collected so far")
            pool.shutdown(wait=False, cancel_futures=True)
        else:
            pool.shutdown()

    summaries.sort(key=lambda s: s["ticker"])
    write_csv(summary_path, SUMMARY_COLUMNS, summaries)
    found = sum(1 for s in summaries if s["status"] == "ok")
    log.info("Wrote %d holdings for %d of %d ETFs to %s",
             holdings_total, found, len(summaries), holdings_path)
    log.info("Wrote per-ETF status to %s", summary_path)
    return 130 if interrupted else 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )
    if not args.verbose:
        logging.getLogger("urllib3").setLevel(logging.WARNING)
    try:
        return run(args)
    except requests.RequestException as exc:  # per-ETF failures are handled inside run()
        log.error("Could not download the ETF list: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
