import csv
import datetime as dt
import io
import json
import sys
import time
from pathlib import Path

import pytest
import requests
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import etf_holdings_scraper as s  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures"


def fixture(name: str) -> bytes:
    return (FIXTURES / name).read_bytes()


def ssga_workbook() -> bytes:
    """Rows as found in SSGA's SPY holdings workbook (holdings trimmed)."""
    rows = [
        [None] * 8,
        ["Fund Name:", "State Street® SPDR® S&P 500® ETF Trust"],
        ["Ticker Symbol:", "SPY"],
        ["Holdings:", "As of 08-Oct-2026"],
        [None] * 8,
        ["Name", "Ticker", "Identifier", "SEDOL", "Weight", "Sector", "Shares Held",
         "Local Currency"],
        ["NVIDIA CORP", "NVDA", "67066G104", "2379504", 8.336249, "-", 295403453.0, "USD"],
        ["APPLE INC", "AAPL", "037833100", "2046251", 7.425376, "-", 178148225.0, "USD"],
        ["TPG INC", "2602335D", "436CVR021", "-", 0.000003, "-", 2578626.0, "USD"],
        [None] * 8,
        ["Past performance is not a reliable indicator of future performance."],
    ]
    workbook = Workbook()
    for row in rows:
        workbook.active.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


class FakeFetcher:
    """Serves canned responses by URL; anything else is a 404."""

    def __init__(self, responses: dict):
        self.responses = responses
        self.requested: list[str] = []

    def get(self, url, max_age=s.HOLDINGS_TTL):
        self.requested.append(url)
        response = self.responses.get(url)
        if response is None:
            raise s.NotFound("HTTP 404")
        if isinstance(response, Exception):
            raise response
        return response


def http_error(status: int) -> requests.HTTPError:
    resp = requests.Response()
    resp.status_code = status
    return requests.HTTPError(f"HTTP {status}", response=resp)


# --------------------------------------------------------------------- helpers


def test_parse_nasdaq_directory_keeps_only_listed_etfs():
    etfs = s.parse_nasdaq_directory(fixture("nasdaqtraded_sample.txt").decode())
    tickers = [e.ticker for e in etfs]
    # AAPL is not an ETF and ZXZZT is a test issue
    assert tickers == ["FPE", "GLD", "IVV", "MCHI", "QQQ", "QYLD", "SPY", "VOO"]
    spy = next(e for e in etfs if e.ticker == "SPY")
    assert spy == s.Etf("SPY", "State Street SPDR S&P 500 ETF Trust", "NYSE Arca")


@pytest.mark.parametrize("raw, expected", [
    ("1,234.50", 1234.5),
    ("$156,125,130.29", 156125130.29),
    ("2.66%", 2.66),
    ("(12)", -12.0),
    ("-0.01", -0.01),
    (7.5, 7.5),
    ("-", ""),
    ("NA", ""),
    ("", ""),
    (None, ""),
])
def test_parse_number(raw, expected):
    assert s.parse_number(raw) == expected


@pytest.mark.parametrize("text, expected", [
    ("Holdings: As of 08-Oct-2026", "2026-10-08"),
    ("Holdings of the Fund as of 10/9/2026", "2026-10-09"),
    ('Fund Holdings as of Oct 09, 2026', "2026-10-09"),
    ("as of September 30, 2026", "2026-09-30"),
    ("Inception Date May 15, 2000\nFund Holdings as of 2026-10-09", "2026-10-09"),
    ("no date here", ""),
])
def test_find_as_of(text, expected):
    assert s.find_as_of(text) == expected


def test_rows_to_holdings_finds_header_and_stops_at_footer():
    rows = [
        ["Fund Name:", "Example ETF"],
        ["Ticker Symbol:", "EXM"],  # names a column alias, but is not a header
        ["Security", "Weight (%)", "Shares"],
        ["Acme Corp", "60.5", "1,000"],
        ["Cash", "39.5", ""],
        ["", "", ""],
        ["Acme Corp", "99", "1"],  # past the end of the table
    ]
    holdings, preamble = s.rows_to_holdings(rows)
    assert [h["holding_name"] for h in holdings] == ["Acme Corp", "Cash"]
    assert holdings[0]["weight_pct"] == 60.5 and holdings[0]["shares"] == 1000.0
    assert holdings[1]["shares"] == ""
    assert "Example ETF" in preamble


def test_rows_to_holdings_without_table_is_not_available():
    with pytest.raises(s.NotAvailable):
        s.rows_to_holdings([["Fund Name:", "Example"], ["nothing", "here"]])


def test_recent_business_days_skips_weekends():
    sunday = dt.date(2026, 10, 11)
    assert s.recent_business_days(sunday, 3) == [
        dt.date(2026, 10, 9), dt.date(2026, 10, 8), dt.date(2026, 10, 7)]


def test_rate_limiter_spaces_requests():
    limiter = s.RateLimiter(rate=20)
    start = time.monotonic()
    for _ in range(5):
        limiter.wait()
    assert time.monotonic() - start >= 4 / 20 * 0.9


# --------------------------------------------------------------------- sources

IVV = s.Etf("IVV", "iShares Core S&P 500 ETF", "NYSE Arca")
SPY = s.Etf("SPY", "State Street SPDR S&P 500 ETF Trust", "NYSE Arca")
QQQ = s.Etf("QQQ", "Invesco QQQ Trust, Series 1", "Nasdaq")
FPE = s.Etf("FPE", "First Trust Preferred Securities and Income ETF ETF", "NYSE Arca")
QYLD = s.Etf("QYLD", "Global X NASDAQ 100 Covered Call ETF", "Nasdaq")
VOO = s.Etf("VOO", "Vanguard S&P 500 ETF", "NYSE Arca")

IVV_URL = ("https://www.ishares.com/us/products/239726/ishares-core-sp-500-etf/1467271812596.ajax"
           "?fileType=csv&fileName=IVV_holdings&dataType=fund")
SPY_URL = ("https://www.ssga.com/library-content/products/fund-data/etfs/us/"
           "holdings-daily-us-en-spy.xlsx")
QQQ_URL = ("https://dng-api.invesco.com/cache/v1/accounts/en_US/shareclasses/QQQ/holdings/fund"
           "?idType=ticker&interval=monthly&productType=ETF")
FPE_URL = "https://www.ftportfolios.com/Retail/Etf/EtfHoldings.aspx?Ticker=FPE"
QYLD_URL = "https://assets.globalxetfs.com/funds/holdings/qyld_full-holdings_{}.csv"
ZACKS_URL = "https://www.zacks.com/funds/etf/{}/holding"


def test_ishares_fund_list_maps_tickers_to_pages():
    pages = s.parse_ishares_fund_list(json.loads(fixture("ishares_fund_list.json")))
    assert pages == {
        "IVV": "/us/products/239726/ishares-core-sp-500-etf",
        "MCHI": "/us/products/239619/ishares-msci-china-etf",
    }


def test_ishares_source():
    fetcher = FakeFetcher({
        s.IsharesSource.FUND_LIST: fixture("ishares_fund_list.json"),
        IVV_URL: fixture("ishares_ivv_holdings.csv"),
    })
    source = s.IsharesSource(fetcher)
    source.prepare()
    assert source.handles(IVV) and not source.handles(SPY)
    result = source.fetch(IVV)
    assert result.as_of == "2026-10-09"
    assert [h["ticker"] for h in result.rows] == ["NVDA", "AAPL", "ESZ6", "USD"]
    nvda = result.rows[0]
    assert nvda["holding_name"] == "NVIDIA CORP"
    assert nvda["weight_pct"] == 8.34
    assert nvda["market_value"] == 75912345678.12
    assert nvda["shares"] == 419364521.0
    assert nvda["sector"] == "Information Technology"
    assert nvda["asset_class"] == "Equity"
    assert nvda["country"] == "United States"
    assert result.rows[3]["weight_pct"] == -0.01


def test_ishares_without_fund_list_handles_nothing():
    source = s.IsharesSource(FakeFetcher({}))
    source.prepare()
    assert not source.handles(IVV)


def test_ssga_source():
    source = s.SsgaSource(FakeFetcher({SPY_URL: ssga_workbook()}))
    assert source.handles(SPY)
    assert source.handles(s.Etf("XLK", "Technology Select Sector SPDR Fund", ""))
    assert not source.handles(IVV)
    result = source.fetch(SPY)
    assert result.as_of == "2026-10-08"
    assert len(result.rows) == 3
    nvda = result.rows[0]
    assert (nvda["ticker"], nvda["cusip"], nvda["sedol"]) == ("NVDA", "67066G104", "2379504")
    assert nvda["weight_pct"] == 8.336249 and nvda["shares"] == 295403453.0
    assert nvda["sector"] == "" and nvda["currency"] == "USD"
    assert result.rows[2]["sedol"] == ""


def test_invesco_source():
    source = s.InvescoSource(FakeFetcher({QQQ_URL: fixture("invesco_qqq.json")}))
    assert source.handles(QQQ) and not source.handles(SPY)
    result = source.fetch(QQQ)
    assert result.as_of == "2026-10-09"
    assert len(result.rows) == 5
    nvda = result.rows[0]
    assert nvda["holding_name"] == "NVIDIA Corp"
    assert (nvda["ticker"], nvda["cusip"], nvda["shares"], nvda["weight_pct"]) == (
        "NVDA", "67066G104", 182116420.0, 8.250119)
    assert result.rows[3]["holding_name"] == "CASH & EQUIVALENTS"
    assert result.rows[4]["weight_pct"] == ""


def test_first_trust_source():
    source = s.FirstTrustSource(FakeFetcher({FPE_URL: fixture("first_trust_fpe.html")}))
    assert source.handles(FPE)
    assert source.handles(s.Etf("DJAN", "FT Vest U.S. Equity Buffer ETF - January", ""))
    assert not source.handles(QQQ)
    result = source.fetch(FPE)
    assert result.as_of == "2026-10-09"
    assert len(result.rows) == 4
    wfc = result.rows[1]
    assert wfc["holding_name"] == "Wells Fargo & Company, Series L, 7.500%"
    assert (wfc["ticker"], wfc["cusip"]) == ("WFC.L", "949746804")
    assert (wfc["shares"], wfc["market_value"], wfc["weight_pct"]) == (102786.0, 110146505.46, 1.88)
    assert result.rows[0]["ticker"] == ""


def test_global_x_source_looks_back_for_latest_file():
    friday = QYLD_URL.format("20261009")
    fetcher = FakeFetcher({
        QYLD_URL.format("20261012"): http_error(403),  # Monday's file not published yet
        friday: fixture("global_x_qyld_holdings.csv"),
    })
    source = s.GlobalXSource(fetcher, today=dt.date(2026, 10, 12))
    assert source.handles(QYLD) and not source.handles(SPY)
    result = source.fetch(QYLD)
    assert fetcher.requested == [QYLD_URL.format("20261012"), friday]
    assert result.url == friday
    assert result.as_of == "2026-10-09"
    assert [h["ticker"] for h in result.rows] == ["NVDA", "AAPL", "NDX US 10/16/26 C25900", "CASH"]
    nvda = result.rows[0]
    assert (nvda["weight_pct"], nvda["shares"], nvda["price"], nvda["market_value"]) == (
        8.39, 3512880.0, 181.02, 635911537.6)


def test_global_x_source_gives_up_after_lookback():
    source = s.GlobalXSource(FakeFetcher({}), today=dt.date(2026, 10, 12))
    with pytest.raises(s.NotAvailable):
        source.fetch(QYLD)


def test_zacks_parser():
    holdings = s.parse_zacks(fixture("zacks_spy.html").decode())
    assert [h["holding_name"] for h in holdings] == [
        "NVIDIA CORP", "APPLE INC", "MICROSOFT CORP", "CONTRA HOLOGIC INCORPO"]
    nvda = holdings[0]
    assert (nvda["ticker"], nvda["shares"], nvda["weight_pct"]) == ("NVDA", 295192685.0, 8.39)
    contra = holdings[3]
    assert (contra["ticker"], contra["other_id"], contra["weight_pct"]) == ("", "2602335D", "")


def test_zacks_parser_reads_cusips_of_bonds():
    holdings = s.parse_zacks(fixture("zacks_agg.html").decode())
    assert holdings[0]["holding_name"] == "TREASURY NOTE-4.63%-02-15-35"
    assert (holdings[0]["cusip"], holdings[0]["ticker"]) == ("91282CMM0", "")


def test_zacks_parser_without_holdings():
    with pytest.raises(s.NotAvailable):
        s.parse_zacks("<html>no data</html>")


# --------------------------------------------------------------------- scraper


def test_scraper_falls_back_when_issuer_source_fails():
    fetcher = FakeFetcher({
        SPY_URL: http_error(500),
        ZACKS_URL.format("SPY"): fixture("zacks_spy.html"),
    })
    scraper = s.HoldingsScraper([s.SsgaSource(fetcher)], s.ZacksSource(fetcher))
    summary, rows = scraper.scrape(SPY)
    assert summary["status"] == "ok"
    assert summary["source"] == "zacks"
    assert summary["as_of_date"] == ""
    assert summary["holdings_count"] == 4
    assert "ssga: HTTPError: HTTP 500" in summary["detail"]
    assert rows[0]["etf_ticker"] == "SPY" and rows[0]["source"] == "zacks"


def test_scraper_reports_when_nothing_has_data():
    scraper = s.HoldingsScraper([], s.ZacksSource(FakeFetcher({})))
    summary, rows = scraper.scrape(VOO)
    assert summary["status"] == "no_data"
    assert summary["detail"] == "zacks: HTTP 404"
    assert rows == []


def test_scraper_without_fallback():
    scraper = s.HoldingsScraper([s.SsgaSource(FakeFetcher({}))], None)
    summary, _ = scraper.scrape(VOO)
    assert summary["detail"] == "no source covers this ETF"


class FakeHttp:
    def __init__(self, responses):
        self.fetcher = FakeFetcher(responses)

    def get(self, url):
        return self.fetcher.get(url)


def test_run_writes_holdings_and_summary(tmp_path):
    responses = {
        s.NASDAQ_SYMBOL_DIRECTORY: fixture("nasdaqtraded_sample.txt"),
        s.IsharesSource.FUND_LIST: fixture("ishares_fund_list.json"),
        IVV_URL: fixture("ishares_ivv_holdings.csv"),
        SPY_URL: ssga_workbook(),
        QQQ_URL: fixture("invesco_qqq.json"),
        FPE_URL: fixture("first_trust_fpe.html"),
        ZACKS_URL.format("VOO"): fixture("zacks_spy.html"),
    }
    args = s.parse_args([
        "--tickers", "SPY,IVV", "QQQ", "FPE", "VOO", "GLD",
        "--sources", "ishares", "ssga", "invesco", "first_trust",
        "--output-dir", str(tmp_path), "--no-cache", "--workers", "2",
    ])
    assert s.run(args, http=FakeHttp(responses)) == 0

    with (tmp_path / "etf_summary.csv").open() as fh:
        summary = {row["ticker"]: row for row in csv.DictReader(fh)}
    assert list(summary) == ["FPE", "GLD", "IVV", "QQQ", "SPY", "VOO"]
    assert {t: summary[t]["source"] for t in summary} == {
        "FPE": "first_trust", "GLD": "", "IVV": "ishares", "QQQ": "invesco",
        "SPY": "ssga", "VOO": "zacks"}
    assert summary["SPY"]["as_of_date"] == "2026-10-08"
    assert summary["IVV"]["holdings_count"] == "4"
    # GLD is a SPDR fund: SSGA has no holdings file for it and Zacks has nothing either.
    assert summary["GLD"]["status"] == "no_data"
    assert summary["GLD"]["detail"] == "ssga: HTTP 404; zacks: HTTP 404"

    with (tmp_path / "holdings.csv").open() as fh:
        holdings = list(csv.DictReader(fh))
    assert holdings and list(holdings[0]) == s.HOLDING_COLUMNS
    counts = {}
    for row in holdings:
        counts[row["etf_ticker"]] = counts.get(row["etf_ticker"], 0) + 1
    assert counts == {"FPE": 4, "IVV": 4, "QQQ": 5, "SPY": 3, "VOO": 4}
    spy_nvda = next(r for r in holdings if r["etf_ticker"] == "SPY" and r["ticker"] == "NVDA")
    assert spy_nvda["weight_pct"] == "8.336249" and spy_nvda["as_of_date"] == "2026-10-08"


def test_disk_cache_round_trip(tmp_path):
    cache = s.DiskCache(tmp_path)
    url = "https://example.com/a?b=1"
    assert cache.read(url, max_age=60) is None
    cache.write(url, b"payload")
    assert cache.read(url, max_age=60) == b"payload"
    assert cache.read(url, max_age=-1) is None
    assert s.DiskCache(None).read(url, max_age=60) is None


def test_parse_args_splits_tickers():
    args = s.parse_args(["--tickers", "spy,ivv", "qqq"])
    assert args.tickers == ["SPY", "IVV", "QQQ"]
