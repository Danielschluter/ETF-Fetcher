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

    def forget(self, url):
        pass


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


# ------------------------------------------------------------------------ http


def html_response(body: bytes, content_type: str = "text/html; charset=utf-8"):
    resp = requests.Response()
    resp.status_code = 200
    resp._content = body
    resp.headers["Content-Type"] = content_type
    return resp


class FakeSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.requested = []

    def get(self, url, timeout):
        self.requested.append(url)
        return self.responses.pop(0)


def client_serving(responses, monkeypatch):
    client = s.HttpClient("test", rate=1000)
    client.BOT_CHECK_PAUSES = (0, 0, 0)
    session = FakeSession(responses)
    monkeypatch.setattr(client, "_session", lambda: session)
    pauses = []
    original_pause = s.RateLimiter.pause

    def record_pause(limiter, seconds):
        pauses.append(seconds)
        return original_pause(limiter, seconds)

    monkeypatch.setattr(s.RateLimiter, "pause", record_pause)
    return client, session, pauses


def test_bot_check_page_is_recognized():
    assert s.looks_like_bot_check(html_response(fixture("zacks_bot_check.html")))
    assert not s.looks_like_bot_check(html_response(fixture("zacks_spy.html")))
    assert not s.looks_like_bot_check(
        html_response(b"Pardon Our Interruption", content_type="text/csv"))


def test_bot_check_pauses_the_site_and_retries(monkeypatch):
    holdings_page = fixture("zacks_spy.html")
    client, session, pauses = client_serving(
        [html_response(fixture("zacks_bot_check.html")), html_response(holdings_page)],
        monkeypatch)
    assert client.get(ZACKS_URL.format("SPY")) == holdings_page
    assert len(session.requested) == 2
    assert pauses == [0]


def test_site_that_keeps_blocking_is_skipped(monkeypatch):
    block = fixture("zacks_bot_check.html")
    client, session, pauses = client_serving([html_response(block)] * 12, monkeypatch)
    for ticker in ("A", "B", "C"):
        with pytest.raises(s.Blocked, match="answered with a bot check"):
            client.get(ZACKS_URL.format(ticker))
    assert len(session.requested) == 12  # each page: first try plus one retry per pause
    with pytest.raises(s.Blocked, match="skipped: www.zacks.com"):
        client.get(ZACKS_URL.format("D"))
    assert len(session.requested) == 12


def test_blocked_response_is_not_cached(tmp_path, monkeypatch):
    client, _, _ = client_serving([html_response(fixture("zacks_bot_check.html"))] * 4,
                                  monkeypatch)
    cache = s.DiskCache(tmp_path)
    url = ZACKS_URL.format("SPY")
    with pytest.raises(s.Blocked):
        s.Fetcher(client, cache).get(url)
    assert cache.read(url, max_age=60) is None


# --------------------------------------------------------------------- sources

IVV = s.Etf("IVV", "iShares Core S&P 500 ETF", "NYSE Arca")
SPY = s.Etf("SPY", "State Street SPDR S&P 500 ETF Trust", "NYSE Arca")
QQQ = s.Etf("QQQ", "Invesco QQQ Trust, Series 1", "Nasdaq")
FPE = s.Etf("FPE", "First Trust Preferred Securities and Income ETF ETF", "NYSE Arca")
QYLD = s.Etf("QYLD", "Global X NASDAQ 100 Covered Call ETF", "Nasdaq")
VOO = s.Etf("VOO", "Vanguard S&P 500 ETF", "NYSE Arca")

IVV_URL = "https://www.ishares.com/us/products/239726/ishares-core-sp-500-etf/latest-holdings.csv"
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
    assert [h["ticker"] for h in result.rows] == ["NVDA", "AAPL", "ETD_USD", "ESZ6"]
    nvda = result.rows[0]
    assert nvda["holding_name"] == "NVIDIA"
    assert nvda["weight_pct"] == 8.23
    assert nvda["market_value"] == 73792336834.24
    assert nvda["shares"] == 321843758.0
    assert nvda["price"] == 229.28
    assert nvda["sector"] == "Information Technology"
    assert nvda["asset_class"] == "Equity"
    assert nvda["country"] == "United States"
    assert result.rows[3]["asset_class"] == "Futures" and result.rows[3]["country"] == ""


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


def test_unparseable_response_is_described_and_dropped_from_cache(tmp_path):
    url = ZACKS_URL.format("IVV")
    block_page = b"<html><head><title>Access Denied</title></head><body>No.</body></html>"
    cache = s.DiskCache(tmp_path)
    fetcher = s.Fetcher(FakeHttp({url: block_page}), cache)
    with pytest.raises(s.NotAvailable) as error:
        s.ZacksSource(fetcher).fetch(IVV)
    assert str(error.value) == (
        f"no holdings on page; got a {len(block_page)}-byte HTML page titled 'Access Denied'")
    assert cache.read(url, max_age=60) is None


def test_describe_response_without_html_title():
    data = b"\n\nSorry, try later\nmore"
    assert s.describe_response(data) == f"got {len(data)} bytes starting 'Sorry, try later'"


def test_zacks_parser_without_holdings():
    with pytest.raises(s.NotAvailable):
        s.parse_zacks("<html>no data</html>")


# --------------------------------------------------------------- value weights


def holding(name, market_value, weight_pct):
    return s.normalize_holding(
        {"holding_name": name, "market_value": market_value, "weight_pct": weight_pct})


def test_value_weights_restore_weights_rounded_to_zero():
    # 1,000 positions of 0.004% each are published as 0.00%.
    rows = [holding("A", 500_000.0, 50.0), holding("B", 460_000.0, 46.0)]
    rows += [holding(f"tiny {i}", 40.0, 0.0) for i in range(1000)]
    assert s.add_value_weights(rows) == ""
    assert rows[0]["value_weight_pct"] == 50.0
    assert rows[2]["value_weight_pct"] == 0.004
    assert sum(r["value_weight_pct"] for r in rows) == pytest.approx(100)


def test_value_weights_with_short_positions_and_futures():
    as_of, rows = s.parse_first_trust(fixture("first_trust_ftls.html").decode())
    assert as_of == "2026-10-09" and len(rows) == 518
    assert s.add_value_weights(rows) == ""
    assert sum(r["value_weight_pct"] for r in rows if r["value_weight_pct"] != "") == \
        pytest.approx(100)
    # Short sales: negative values and weights, each within rounding of the published weight.
    shorts = [r for r in rows if isinstance(r["market_value"], float) and r["market_value"] < 0]
    assert len(shorts) == 247
    assert all(abs(r["value_weight_pct"] - r["weight_pct"]) <= 0.005 for r in shorts)
    snps = next(r for r in rows if r["ticker"] == "SNPS")
    assert (snps["shares"], snps["weight_pct"]) == (-11283.0, -0.22)
    assert snps["sector"] == "Information Technology"
    # First Trust lists the short S&P 500 futures at their notional amount.
    future = next(r for r in rows if r["holding_name"].startswith("S&P500 EMINI FUT"))
    assert (future["market_value"], future["notional_value"]) == ("", -269589425.0)
    assert future["value_weight_pct"] == ""


def test_value_weights_with_written_options():
    rows, _ = s.rows_to_holdings(s.read_csv_rows(fixture("ishares_ivvw_holdings.csv")))
    assert [r["ticker"] for r in rows] == ["IVV", "XTSLA", "USD", "SPX"]  # not the look-through
    assert s.add_value_weights(rows) == ""
    call = rows[3]
    assert call["holding_name"] == "OCT26 SPX C @ 7625.000000"
    assert (call["market_value"], call["notional_value"]) == (-8847792.03, -334100467.3)
    assert call["value_weight_pct"] == pytest.approx(-2.5475, abs=1e-4)
    assert call["weight_pct"] == -2.55


def test_value_weights_need_market_values():
    rows = s.parse_zacks(fixture("zacks_spy.html").decode())
    assert s.add_value_weights(rows) == ""
    assert all(r["value_weight_pct"] == "" for r in rows)


def test_value_weights_skipped_when_market_values_do_not_add_up():
    # Published weights put net assets at 1,000,000, but only 700,000 is listed.
    rows = [holding("A", 500_000.0, 50.0), holding("B", 200_000.0, 20.0)]
    note = s.add_value_weights(rows)
    assert note == ("value weights not computed: market values add up to 70.0% of the net "
                    "assets implied by the published weights")
    assert rows[0]["value_weight_pct"] == ""


def test_value_weights_skipped_when_some_market_values_are_missing():
    rows = [holding("A", 500_000.0, 50.0), holding("B", None, 50.0)]
    assert s.add_value_weights(rows) == (
        "value weights not computed: some holdings have no market value")


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


def test_scraper_adds_value_weights_to_rows():
    ftls = s.Etf("FTLS", "First Trust Long/Short Equity ETF", "NYSE Arca")
    url = "https://www.ftportfolios.com/Retail/Etf/EtfHoldings.aspx?Ticker=FTLS"
    fetcher = FakeFetcher({url: fixture("first_trust_ftls.html")})
    summary, rows = s.HoldingsScraper([s.FirstTrustSource(fetcher)], None).scrape(ftls)
    assert summary["status"] == "ok" and summary["detail"] == ""
    nvda = next(r for r in rows if r["ticker"] == "NVDA")
    assert nvda["value_weight_pct"] == pytest.approx(6.49, abs=0.005)


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


class FakeBrowser(FakeHttp):
    started = closed = False

    def start(self):
        self.started = True

    def close(self):
        self.closed = True


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
    browser = FakeBrowser(responses)
    assert s.run(args, http=FakeHttp(responses), browser=browser) == 0
    assert browser.started and browser.closed
    # Only ETFs whose issuer source failed or is missing reach Zacks.
    assert sorted(browser.fetcher.requested) == [ZACKS_URL.format(t) for t in ("GLD", "VOO")]

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
