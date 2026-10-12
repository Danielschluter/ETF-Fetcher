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
               fallback whenever an issuer source fails. Zacks turns away plain
               HTTP clients, so its pages are loaded in headless Chrome (via
               Playwright). Zacks does not say which date its holdings are as
               of, so as_of_date is left blank.

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
import queue
import re
import sys
import threading
import time
import warnings
from collections import Counter
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
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


class Blocked(NotAvailable):
    """A website answered with a bot-check page instead of the content."""


# Pages that bot-protection services (Imperva, Cloudflare) serve instead of the content.
_BOT_CHECK = re.compile(
    r"Pardon Our Interruption|<title>\s*Just a moment\.\.\.\s*</title>", re.IGNORECASE)


def looks_like_bot_check(resp: requests.Response) -> bool:
    if "html" not in resp.headers.get("Content-Type", "") or len(resp.content) > 100_000:
        return False
    return bool(_BOT_CHECK.search(resp.content.decode("utf-8", errors="replace")))


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

    def pause(self, seconds: float) -> bool:
        """Hold every request for a while. False if they were already held about as long."""
        with self.lock:
            until = time.monotonic() + seconds
            if until <= self.next_slot + 1:
                return False
            self.next_slot = until
            return True


class HttpClient:
    """GET with a per-host rate limit and retries; one requests.Session per thread.

    Bot protection (on Zacks, for one) lets most requests through but answers
    some with a check page. Then all requests to that site pause and the page
    is retried; a site that keeps blocking is skipped for the rest of the run.
    """

    RETRY_STATUSES = {429, 500, 502, 503, 504}
    BOT_CHECK_PAUSES = (30, 60, 120)  # seconds
    GIVE_UP_AFTER = 3  # pages in a row still blocked after every pause

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
        self.lock = threading.Lock()
        self.limiters: dict[str, RateLimiter] = {}
        self.blocked_streak: Counter = Counter()
        self.given_up: set[str] = set()

    def _session(self) -> requests.Session:
        session = getattr(self.local, "session", None)
        if session is None:
            session = requests.Session()
            session.headers.update(self.headers)
            self.local.session = session
        return session

    def _limiter(self, host: str) -> RateLimiter:
        with self.lock:
            if host not in self.limiters:
                self.limiters[host] = RateLimiter(self.rate)
            return self.limiters[host]

    def get(self, url: str) -> bytes:
        host = urlsplit(url).hostname or ""
        limiter = self._limiter(host)
        for pause in (*self.BOT_CHECK_PAUSES, None):
            if host in self.given_up:
                raise Blocked(f"skipped: {host} kept answering with a bot check")
            resp = self._get_with_retries(url, limiter)
            if not looks_like_bot_check(resp):
                with self.lock:
                    self.blocked_streak[host] = 0
                return resp.content
            if pause is not None and limiter.pause(pause):
                log.warning("%s answered with a bot check; pausing requests to it for %ds",
                            host, pause)
        with self.lock:
            self.blocked_streak[host] += 1
            if self.blocked_streak[host] >= self.GIVE_UP_AFTER and host not in self.given_up:
                self.given_up.add(host)
                log.error("%s keeps answering with a bot check; skipping it for the rest of "
                          "this run. Run again later to fill in the gaps.", host)
        raise Blocked(f"{host} answered with a bot check")

    def _get_with_retries(self, url: str, limiter: RateLimiter) -> requests.Response:
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
                    return resp
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


class SetupError(RuntimeError):
    """Something the run needs could not be started."""


class BrowserPool:
    """Loads pages in headless Chrome, for websites that turn scripts away but serve browsers.

    Playwright objects can only be used by the thread that created them, so each
    worker thread runs its own browser and takes URLs from a shared queue.
    """

    BOT_CHECK_WAIT = 30  # seconds for a bot check to clear itself and reload the page
    GIVE_UP_AFTER = 5  # pages in a row that never got past the bot check

    def __init__(self, workers: int, rate: float, browser: str = "chrome",
                 headless: bool = True, timeout: float = 60):
        self.workers = workers
        self.browser = browser
        self.headless = headless
        self.timeout_ms = timeout * 1000
        self.limiter = RateLimiter(rate)
        self.jobs: queue.Queue = queue.Queue()
        self.threads: list[threading.Thread] = []
        self.lock = threading.Lock()
        self.blocked_streak = 0
        self.given_up = False

    def start(self) -> None:
        try:
            import playwright.sync_api  # noqa: F401
        except ImportError:
            raise SetupError("Playwright is not installed: pip install playwright") from None
        for i in range(self.workers):
            ready = threading.Event()
            failure: list[Exception] = []
            thread = threading.Thread(target=self._work, args=(ready, failure),
                                      name=f"browser-{i}", daemon=True)
            thread.start()
            ready.wait()
            if failure:
                self.close()
                raise SetupError(f"could not start the browser ({self.browser}): {failure[0]}")
            self.threads.append(thread)

    def get(self, url: str) -> bytes:
        if self.given_up:
            raise Blocked("skipped: Zacks kept answering with a bot check")
        future: Future = Future()
        self.jobs.put((url, future))
        return future.result()

    def close(self) -> None:
        while True:  # cancel the pages that haven't started loading
            try:
                job = self.jobs.get_nowait()
            except queue.Empty:
                break
            if job is not None:
                job[1].cancel()
        for _ in self.threads:
            self.jobs.put(None)
        for thread in self.threads:
            thread.join(timeout=60)
        self.threads = []

    def _launch_options(self) -> dict:
        options: dict = {"headless": self.headless,
                         "args": ["--disable-blink-features=AutomationControlled"]}
        if self.browser == "chrome":
            options["channel"] = "chrome"  # the installed Google Chrome
        elif self.browser != "chromium":  # "chromium" is Playwright's own download
            options["executable_path"] = self.browser
        return options

    def _work(self, ready: threading.Event, failure: list[Exception]) -> None:
        from playwright.sync_api import sync_playwright
        try:
            with sync_playwright() as playwright:
                try:
                    browser = playwright.chromium.launch(**self._launch_options())
                    page = browser.new_page()
                except Exception as exc:
                    failure.append(exc)
                    return
                finally:
                    ready.set()
                self._serve(browser, page)
                browser.close()
        except Exception as exc:
            if not ready.is_set():
                failure.append(exc)
                ready.set()
            log.debug("Browser worker stopped", exc_info=True)

    def _serve(self, browser, page) -> None:
        while (job := self.jobs.get()) is not None:
            url, future = job
            if not future.set_running_or_notify_cancel():
                continue
            try:
                if self.given_up:
                    raise Blocked("skipped: Zacks kept answering with a bot check")
                if page.is_closed():
                    page = browser.new_page()
                future.set_result(self._load(page, url))
            except Exception as exc:
                future.set_exception(exc)

    def _load(self, page, url: str) -> bytes:
        self.limiter.wait()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
        except Exception:
            if page.url in ("", "about:blank"):
                raise  # nothing loaded at all
            # Otherwise the bot check's own reload interrupted it; keep watching the page.
        deadline = time.monotonic() + self.BOT_CHECK_WAIT
        while True:
            try:
                html = page.content()
            except Exception:  # the page is reloading
                html = ""
            if html and not _BOT_CHECK.search(html):
                self._record(blocked=False)
                return html.encode("utf-8")
            if time.monotonic() > deadline:
                self._record(blocked=True)
                raise Blocked("answered with a bot check that did not clear")
            page.wait_for_timeout(500)

    def _record(self, blocked: bool) -> None:
        with self.lock:
            self.blocked_streak = self.blocked_streak + 1 if blocked else 0
            if self.blocked_streak >= self.GIVE_UP_AFTER and not self.given_up:
                self.given_up = True
                log.error("Zacks keeps answering with a bot check, even in a browser; skipping "
                          "it for the rest of this run. Run again later to fill in the gaps.")


class Fetcher:
    """A cache in front of an HttpClient or BrowserPool, so reruns and resumed runs are fast."""

    def __init__(self, http: HttpClient | BrowserPool, cache: DiskCache, refresh: bool = False):
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
    HOLDINGS = "https://www.ishares.com{page}/latest-holdings.csv"

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
        url = self.HOLDINGS.format(page=self.pages[etf.ticker])
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
    parser.add_argument("--browser", default="chrome",
                        help='Browser for Zacks pages: "chrome" (your installed Google Chrome), '
                             '"chromium" (after `playwright install chromium`), or a path')
    parser.add_argument("--browser-workers", type=int, default=4,
                        help="Browsers loading Zacks pages in parallel")
    parser.add_argument("--show-browser", action="store_true",
                        help="Show the browser windows instead of running them in the background")
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
    if args.workers < 1 or args.browser_workers < 1:
        parser.error("--workers and --browser-workers must be at least 1")
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


def run(args: argparse.Namespace, http: HttpClient | None = None,
        browser: BrowserPool | None = None) -> int:
    http = http or HttpClient(args.user_agent, args.rate)
    cache = DiskCache(None if args.no_cache else args.cache_dir)
    fetcher = Fetcher(http, cache, args.refresh)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    universe = load_universe(fetcher)
    log.info("Found %d US-listed ETFs", len(universe))
    etfs = select_etfs(universe, args.tickers, args.limit)

    issuers = [ISSUER_SOURCES[name](fetcher) for name in args.sources]
    for source in issuers:
        source.prepare()
    fallback = None
    if not args.no_fallback:
        # Zacks turns away plain HTTP clients, so its pages are loaded in a browser.
        browser = browser or BrowserPool(args.browser_workers, args.rate, args.browser,
                                         headless=not args.show_browser)
        browser.start()
        fallback = ZacksSource(Fetcher(browser, cache, args.refresh))
    try:
        return scrape_all(args, etfs, HoldingsScraper(issuers, fallback))
    finally:
        if browser is not None:
            browser.close()


def scrape_all(args: argparse.Namespace, etfs: list[Etf], scraper: HoldingsScraper) -> int:
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
    except SetupError as exc:
        log.error("Zacks fallback unavailable: %s. Install Google Chrome, or run "
                  "`playwright install chromium` and pass --browser chromium, or pass "
                  "--no-fallback to use issuer websites only.", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
