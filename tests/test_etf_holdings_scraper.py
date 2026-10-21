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
    """Serves canned responses by URL (by s.request_key for POSTs); anything else is a 404."""

    def __init__(self, responses: dict):
        self.responses = responses
        self.requested: list[str] = []
        self.posted: list[tuple] = []

    def post_json(self, url, payload, headers=None, max_age=s.HOLDINGS_TTL):
        self.posted.append((url, payload, headers))
        return self.get(s.request_key(url, payload))

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
    ("As of 10/13/26", "2026-10-13"),
    ("no date here", ""),
])
def test_find_as_of(text, expected):
    assert s.find_as_of(text) == expected


@pytest.mark.parametrize("text, expected", [
    ("01/31/2028", "2028-01-31"),
    ("11/01/2027-10/01/2056", "2027-11-01/2056-10-01"),  # a pool of mortgages
    ("2027-11-01", "2027-11-01"),
    ("Perpetual", "Perpetual"),
    ("", ""),
])
def test_parse_maturity(text, expected):
    assert s.parse_maturity(text) == expected


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


def test_rows_to_holdings_reads_the_date_column_as_the_as_of_date():
    rows = [["Date", "Ticker", "Weight"], ["10/14/2026", "AAA", "60"], ["10/14/2026", "BBB", "40"]]
    holdings, preamble = s.rows_to_holdings(rows)
    assert s.find_as_of(preamble) == "2026-10-14"
    assert [h["ticker"] for h in holdings] == ["AAA", "BBB"]


def test_split_by_fund():
    rows = [["Date", "Account", "StockTicker", "Weightings"],
            ["10/15/2026", "AAA", "X", "60%"], ["10/15/2026", "BBB", "Y", "100%"],
            ["10/15/2026", "aaa", "Z", "40%"]]
    funds = s.split_by_fund(rows, "account")
    assert funds == {"AAA": [rows[0], rows[1], rows[3]], "BBB": [rows[0], rows[2]]}
    with pytest.raises(s.NotAvailable):
        s.split_by_fund(rows, "fund")


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


def test_sites_asking_for_a_crawl_delay_get_it():
    client = s.HttpClient("test", rate=2)
    assert client._limiter("www.vaneck.com").interval == 25
    assert client._limiter("www.ishares.com").interval == 0.5


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
        self.posted = []

    def get(self, url, timeout):
        self.requested.append(url)
        return self.responses.pop(0)

    def post(self, url, json, headers, timeout):
        self.posted.append((url, json, headers))
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


def test_post_json_sends_the_payload(monkeypatch):
    client, session, _ = client_serving(
        [html_response(b'{"data": {}}', content_type="application/json")], monkeypatch)
    payload = {"query": "query q { x }"}
    assert client.post_json("https://example.com/graphql", payload, {"Origin": "x"}) == (
        b'{"data": {}}')
    assert session.posted == [("https://example.com/graphql", payload, {"Origin": "x"})]
    assert session.requested == []


def test_posted_responses_are_cached_per_payload(tmp_path):
    class Http:
        calls = 0

        def post_json(self, url, payload, headers=None):
            self.calls += 1
            return json.dumps(payload).encode()

    http = Http()
    fetcher = s.Fetcher(http, s.DiskCache(tmp_path))
    url = "https://example.com/graphql"
    assert fetcher.post_json(url, {"id": 1}) == b'{"id": 1}'
    assert fetcher.post_json(url, {"id": 2}) == b'{"id": 2}'
    assert fetcher.post_json(url, {"id": 1}) == b'{"id": 1}'
    assert http.calls == 2
    assert s.request_key(url, {"id": 1}) != s.request_key(url, {"id": 2})
    assert s.request_key(url) == url


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
INVESCO_PAGE = ("https://www.invesco.com/us/financial-products/etfs/product-detail"
                "?audienceType=Investor&ticker={}")
# A fund keyed by ticker, as QQQ's holdings are.
QQQ_PAGE = (b'<meta name="cusip" content="46090E103">'
            b'<meta name="uniqueIdentifier" content="ticker"><meta name="ticker" content="QQQ">')
RSP = s.Etf("RSP", "Invesco S&P 500 Equal Weight ETF", "NYSE Arca")
RSP_URL = ("https://dng-api.invesco.com/cache/v1/accounts/en_US/shareclasses/46137V357/holdings/"
           "fund?idType=cusip&interval=monthly&productType=ETF")
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


def test_ishares_market_weight_layout():
    # Some iShares funds name the column "Market Weight" and add a "Type" column.
    rows, preamble = s.rows_to_holdings(s.read_csv_rows(fixture("ishares_iai_holdings.csv")))
    assert s.find_as_of(preamble) == "2026-10-13"
    gs = rows[0]
    assert (gs["ticker"], gs["weight_pct"], gs["asset_class"]) == ("GS", 17.84, "Equity")
    assert rows[2]["asset_class"] == "Futures" and rows[2]["weight_pct"] == ""


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


def test_invesco_source_by_cusip():
    source = s.InvescoSource(FakeFetcher({
        INVESCO_PAGE.format("RSP"): fixture("invesco_rsp_page.html"),
        RSP_URL: fixture("invesco_rsp.json"),
    }))
    assert source.handles(RSP)
    result = source.fetch(RSP)
    assert result.url == RSP_URL and result.as_of == "2026-10-13"
    first = result.rows[0]
    assert (first["ticker"], first["cusip"], first["weight_pct"]) == ("P", "74624M102", 0.324309)
    assert first["market_value"] == 319721832.98


def test_invesco_fund_page_without_identifier():
    with pytest.raises(s.NotAvailable):
        s.parse_invesco_fund_id("<html><head><title>Not found</title></head></html>")


def test_invesco_source_by_ticker():
    source = s.InvescoSource(FakeFetcher({
        INVESCO_PAGE.format("QQQ"): QQQ_PAGE,
        QQQ_URL: fixture("invesco_qqq.json"),
    }))
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


VANGUARD_URL = "https://investor.vanguard.com/vmf/api/{}/portfolio-holding/{}?start={}&count={}"
# What the API answers for a kind of holding the fund doesn't have.
VANGUARD_NONE = (b'{"size":0,"asOfDate":"2026-08-31T00:00:00-04:00","self":{"type":'
                 b'"application/xml,application/json","href":"http://api.vanguard.com/rs/ire/01/'
                 b'ind/fund/VOO/portfolio-holding/bond?start=1&count=2","ref":"self"},"fund":{}}')


def test_vanguard_source_pages_through_holdings():
    # The VOO pages are real responses trimmed to 2 + 1 holdings, with "size" set to match.
    fetcher = FakeFetcher({
        VANGUARD_URL.format("VOO", "stock", 1, 2): fixture("vanguard_voo_stock_1.json"),
        VANGUARD_URL.format("VOO", "stock", 3, 2): fixture("vanguard_voo_stock_2.json"),
        VANGUARD_URL.format("VOO", "bond", 1, 2): VANGUARD_NONE,
    })
    source = s.VanguardSource(fetcher)
    source.PAGE_SIZE = 2
    assert source.handles(VOO) and not source.handles(SPY)
    result = source.fetch(VOO)
    assert len(fetcher.requested) == 3
    assert result.url == VANGUARD_URL.format("VOO", "stock", 1, 2)
    assert result.as_of == "2026-08-31"
    assert [h["ticker"] for h in result.rows] == ["NVDA", "AAPL", "BF.B"]
    nvda = result.rows[0]
    assert nvda["holding_name"] == "NVIDIA Corp."
    assert (nvda["cusip"], nvda["isin"], nvda["sedol"]) == ("67066G104", "US67066G1040", "2379504")
    assert (nvda["shares"], nvda["market_value"], nvda["weight_pct"]) == (
        643435730.0, 142057740469.4, 8.08)
    assert nvda["asset_class"] == "Stock" and nvda["notional_value"] == ""


def test_vanguard_source_reads_bonds():
    bnd = s.Etf("BND", "Vanguard Total Bond Market ETF", "Nasdaq")
    fetcher = FakeFetcher({
        VANGUARD_URL.format("BND", "stock", 1, 3000): VANGUARD_NONE,
        VANGUARD_URL.format("BND", "bond", 1, 3000): fixture("vanguard_bnd_bond.json"),
    })
    result = s.VanguardSource(fetcher).fetch(bnd)
    assert result.url == VANGUARD_URL.format("BND", "bond", 1, 3000)
    pool, note = result.rows
    assert pool["holding_name"] == "Federal National Mortgage Assn."
    assert pool["maturity_date"] == "2027-11-01/2056-10-01"
    assert (pool["shares"], pool["coupon_rate"], pool["asset_class"]) == (6743672547.0, 2.0, "Bond")
    assert (note["cusip"], note["maturity_date"], note["weight_pct"]) == (
        "91282CGH8", "2028-01-31", 0.48)


PROSHARES_PAGE = "https://www.proshares.com/our-etfs/{}/{}"
TQQQ = s.Etf("TQQQ", "ProShares UltraPro QQQ", "Nasdaq")


def test_proshares_source():
    url = PROSHARES_PAGE.format("leveraged-and-inverse", "tqqq")
    source = s.ProSharesSource(FakeFetcher({url: fixture("proshares_tqqq.html")}))
    assert source.handles(TQQQ) and not source.handles(QQQ)
    result = source.fetch(TQQQ)
    assert result.url == url
    assert result.as_of == "2026-10-13"  # not the month-end date above the holdings
    assert [h["holding_name"] for h in result.rows] == [
        "NVIDIA CORP", "APPLE INC", "NASDAQ 100 INDEX SWAP BANK OF AMERICA NA",
        "NASDAQ 100 E-MINI EQUITY INDEX 18/DEC/2026 NQZ6 INDEX",
        "PROSHARES GENIUS MONEY MARKET ETF", "TREASURY BILL", "NET OTHER ASSETS (LIABILITIES)"]
    nvda, swap = result.rows[0], result.rows[2]
    assert (nvda["ticker"], nvda["sedol"], nvda["weight_pct"]) == ("NVDA", "2379504", 2.86)
    assert (nvda["market_value"], nvda["shares"]) == (1094849832.0, 4791046.0)
    assert (swap["ticker"], swap["weight_pct"], swap["notional_value"], swap["market_value"]) == (
        "", 32.54, 12465194582.0, "")


def test_proshares_source_tries_the_other_fund_page_path():
    onln = s.Etf("ONLN", "ProShares Online Retail ETF", "NYSE Arca")
    fetcher = FakeFetcher(
        {PROSHARES_PAGE.format("strategic", "onln"): fixture("proshares_tqqq.html")})
    result = s.ProSharesSource(fetcher).fetch(onln)
    assert fetcher.requested == [PROSHARES_PAGE.format("leveraged-and-inverse", "onln"),
                                 PROSHARES_PAGE.format("strategic", "onln")]
    assert result.url == PROSHARES_PAGE.format("strategic", "onln")
    with pytest.raises(s.NotAvailable, match="no fund page"):
        s.ProSharesSource(FakeFetcher({})).fetch(onln)


AAPR = s.Etf("AAPR", "Innovator Equity Defined Protection ETF - 2 Yr to April 2028", "Cboe BZX")
ZSEP = s.Etf("ZSEP", "Innovator Equity Defined Protection ETF - 1 Yr September", "Cboe BZX")


def test_innovator_source_reads_one_file_for_all_funds():
    fetcher = FakeFetcher({s.InnovatorSource.HOLDINGS: fixture("innovator_xt_holdings.csv")})
    source = s.InnovatorSource(fetcher)
    assert source.handles(AAPR) and not source.handles(SPY)
    result = source.fetch(AAPR)
    assert len(source.fetch(ZSEP).rows) == 7
    assert fetcher.requested == [s.InnovatorSource.HOLDINGS]
    assert result.as_of == "2026-10-15"
    assert len(result.rows) == 7
    voo = next(r for r in result.rows if r["ticker"] == "VOO")
    assert (voo["cusip"], voo["holding_name"], voo["shares"], voo["price"]) == (
        "922908363", "Vanguard S&P 500 ETF", 110921.0, 707.43)
    assert (voo["market_value"], voo["weight_pct"]) == (78468843.03, 110.45)
    written_call = result.rows[1]
    assert (written_call["shares"], written_call["weight_pct"]) == (-1158.0, -15.27)
    assert s.add_value_weights(result.rows) == ""
    assert voo["value_weight_pct"] == pytest.approx(110.45, abs=0.005)
    with pytest.raises(s.NotAvailable, match="not in the holdings file"):
        source.fetch(s.Etf("NEWX", "Innovator New ETF", ""))


def test_innovator_source_downloads_a_failing_file_once():
    fetcher = FakeFetcher({s.InnovatorSource.HOLDINGS: http_error(500)})
    source = s.InnovatorSource(fetcher)
    for etf in (AAPR, ZSEP):
        with pytest.raises(s.NotAvailable, match="HTTP 500"):
            source.fetch(etf)
    assert fetcher.requested == [s.InnovatorSource.HOLDINGS]


def test_pacer_source():
    cowz = s.Etf("COWZ", "Pacer US Cash Cows 100 ETF", "Cboe BZX")
    url = "https://www.paceretfs.com/products/holdings_download/COWZ"
    source = s.PacerSource(FakeFetcher({url: fixture("pacer_cowz_holdings.csv")}))
    assert source.handles(cowz) and not source.handles(SPY)
    result = source.fetch(cowz)
    assert result.as_of == "2026-10-14"
    assert [h["ticker"] for h in result.rows] == ["ABNB", "ACN", "MPC", "USBFS03", "Cash&Other"]
    acn = result.rows[1]
    assert (acn["holding_name"], acn["cusip"], acn["shares"], acn["price"]) == (
        "Accenture PLC", "G1151C101", 2107313.0, 213.56)
    assert (acn["market_value"], acn["weight_pct"]) == (450037764.28, 2.35)


SCHWAB_URL = "https://www.schwabassetmanagement.com/allholdings/SCHD?page={}"


def test_schwab_source_reads_every_page():
    # Real pages trimmed to two holdings each, with "Displaying 1 - 2 of 4" set to match.
    schd = s.Etf("SCHD", "Schwab U.S. Dividend Equity ETF", "NYSE Arca")
    fetcher = FakeFetcher({SCHWAB_URL.format(0): fixture("schwab_schd_page0.html"),
                           SCHWAB_URL.format(1): fixture("schwab_schd_page1.html")})
    source = s.SchwabSource(fetcher)
    assert source.handles(schd) and not source.handles(SPY)
    result = source.fetch(schd)
    assert fetcher.requested == [SCHWAB_URL.format(0), SCHWAB_URL.format(1)]
    assert result.as_of == "2026-10-13"
    assert [h["ticker"] for h in result.rows] == ["TXN", "QCOM", "FAZ6", "DMZ6"]
    txn = result.rows[0]
    assert (txn["holding_name"], txn["cusip"], txn["shares"], txn["weight_pct"]) == (
        "TEXAS INSTRUMENT INC", "882508104", 18224064.0, 4.71)
    assert txn["market_value"] == ""  # shown as "$5.2B"
    assert result.rows[2]["holding_name"] == "S&P MID 400 EMINI DEC 26"


def vaneck_workbook() -> bytes:
    """Rows as found in VanEck's SMH holdings workbook (holdings trimmed)."""
    rows = [
        ["Daily Holdings (%)  10/13/2026"],
        [None] * 9,
        ["Number", "Ticker", "Holding Name", "Identifier (FIGI)", "Shares", "Asset Class",
         "Market Value (US$)", "Notional Value", "% of Net Assets"],
        [1, "NVDA", "Nvidia Corp", "BBG000BBJQV0", "62,977,808", "Stock", "$14,391,688,684.16",
         "--", "20.04%"],
        [2, "TSM", "Taiwan Semiconductor Manufacturing Co L", "BBG000BD8ZK0", "15,048,413", "Stock",
         "$6,698,349,594.56", "--", "9.33%"],
        [26, "-USD CASH-", None, None, "42,568,818", "Cash Bal", "$42,568,817.82", "--", "0.06%"],
        [27, "--", "Other/Cash", "--", "--", "Cash", "$-6,814,694.57", "--", "-0.01%"],
        ["These are not recommendations to buy or to sell any security. Securities and holdings "
         "may vary."],
    ]
    workbook = Workbook()
    for row in rows:
        workbook.active.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


VANECK_URL = "https://www.vaneck.com/us/en/etf/{}/{}/holdings/download/xlsx/"
SMH = s.Etf("SMH", "VanEck Semiconductor ETF", "Nasdaq")


def test_vaneck_source():
    source = s.VanEckSource(FakeFetcher({VANECK_URL.format("equity", "smh"): vaneck_workbook()}))
    assert source.handles(SMH) and not source.handles(SPY)
    result = source.fetch(SMH)
    assert result.as_of == "2026-10-13"
    assert [h["ticker"] for h in result.rows] == ["NVDA", "TSM", "-USD CASH-", ""]
    nvda = result.rows[0]
    assert (nvda["holding_name"], nvda["other_id"], nvda["asset_class"]) == (
        "Nvidia Corp", "BBG000BBJQV0", "Stock")
    assert (nvda["shares"], nvda["market_value"], nvda["weight_pct"]) == (
        62977808.0, 14391688684.16, 20.04)
    assert result.rows[3]["market_value"] == -6814694.57


def test_vaneck_source_tries_the_other_asset_class():
    # With the wrong asset class in the path, the site redirects to the fund's overview page.
    overview = b"<html><head><title>SMH - VanEck Semiconductor ETF</title></head></html>"
    fetcher = FakeFetcher({VANECK_URL.format("equity", "smh"): overview,
                           VANECK_URL.format("income", "smh"): vaneck_workbook()})
    result = s.VanEckSource(fetcher).fetch(SMH)
    assert result.url == VANECK_URL.format("income", "smh")
    angl = s.Etf("ANGL", "VanEck Fallen Angel High Yield Bond ETF", "Nasdaq")
    fetcher = FakeFetcher({})
    with pytest.raises(s.NotAvailable, match="^income: HTTP 404; equity: HTTP 404$"):
        s.VanEckSource(fetcher).fetch(angl)


ARKK = s.Etf("ARKK", "ARK Innovation ETF", "NYSE Arca")
ARKK_CSV = ("https://assets.ark-funds.com/fund-documents/funds-etf-csv/"
            "ARK_INNOVATION_ETF_ARKK_HOLDINGS.csv")


def test_ark_source_finds_the_file_in_the_fund_documents():
    source = s.ArkSource(FakeFetcher({
        "https://www.ark-funds.com/funds/arkk": fixture("ark_arkk_page.html"),
        s.ArkSource.DOCUMENTS.format(fund_id=1004): fixture("ark_arkk_documents.html"),
        ARKK_CSV: fixture("ark_arkk_holdings.csv"),
    }))
    assert source.handles(ARKK) and not source.handles(SPY)
    result = source.fetch(ARKK)
    assert result.url == ARKK_CSV and result.as_of == "2026-10-14"
    assert [h["ticker"] for h in result.rows] == ["TSLA", "SPCX", "", "RKLB UQ", ""]
    tsla = result.rows[0]
    assert (tsla["holding_name"], tsla["cusip"], tsla["shares"]) == (
        "TESLA INC", "88160R101", 1716830.0)
    assert (tsla["market_value"], tsla["weight_pct"]) == (644206120.9, 9.83)
    assert result.rows[2]["cusip"] == "PPOPENAIC"


def test_ark_fund_page_without_holdings_link():
    with pytest.raises(s.NotAvailable, match="links no holdings file"):
        s.find_ark_csv("<html><body>No holdings here</body></html>")
    with pytest.raises(s.NotAvailable, match="no list of documents"):
        s.find_ark_fund_id("<html><body>No documents here</body></html>")


def test_ark_holdings_link_relative_or_in_script_data():
    page_url = "https://www.ark-funds.com/funds/arkk"
    relative = '<a href="/csv/ARK_INNOVATION_ETF_ARKK_HOLDINGS.csv">'
    assert s.find_ark_csv(relative, page_url) == (
        "https://www.ark-funds.com/csv/ARK_INNOVATION_ETF_ARKK_HOLDINGS.csv")
    script = ('{"etfs":[{"csv":"https:\\/\\/assets.ark-funds.com\\/ARK_FINTECH_ARKF_HOLDINGS.csv"},'
              '{"csv":"https:\\/\\/assets.ark-funds.com\\/ARK_INNOVATION_ETF_ARKK_HOLDINGS.csv"}]}')
    assert s.find_ark_csv(script, page_url, "ARKK") == (
        "https://assets.ark-funds.com/ARK_INNOVATION_ETF_ARKK_HOLDINGS.csv")


def test_direxion_source():
    # Representative: columns as in a copy of SOXL's file; Direxion's site blocked the capture tool.
    soxl = s.Etf("SOXL", "Direxion Daily Semiconductor Bull 3X Shares", "NYSE Arca")
    url = "https://www.direxion.com/holdings/SOXL.csv"
    source = s.DirexionSource(FakeFetcher({url: fixture("direxion_soxl_holdings.csv")}))
    assert source.handles(soxl) and not source.handles(SPY)
    result = source.fetch(soxl)
    assert [h["ticker"] for h in result.rows] == ["NVDA", "AVGO", "NXPI"]  # not AccountTicker
    nvda = result.rows[0]
    assert (nvda["holding_name"], nvda["cusip"], nvda["shares"], nvda["weight_pct"]) == (
        "NVIDIA CORP", "67066G104", 661671.0, 6.32)


DIMENSIONAL_URL = "https://tools-blob.dimensional.com/etf/{}/DFAC.csv"


def test_dimensional_source_looks_back_for_latest_file():
    dfac = s.Etf("DFAC", "U.S. Core Equity 2 Portfolio ETF Class Shares", "NYSE Arca")
    fetcher = FakeFetcher(
        {DIMENSIONAL_URL.format("20261013"): fixture("dimensional_dfac_holdings.csv")})
    source = s.DimensionalSource(fetcher, today=dt.date(2026, 10, 15))
    assert source.handles(dfac) and not source.handles(SPY)
    result = source.fetch(dfac)
    assert fetcher.requested == [
        DIMENSIONAL_URL.format(day) for day in ("20261015", "20261014", "20261013")]
    assert result.as_of == "2026-10-13"
    assert [h["ticker"] for h in result.rows] == ["AIR", "", "", "AAPL", "NVDA", "LPG"]
    nvda = result.rows[4]
    assert (nvda["holding_name"], nvda["cusip"], nvda["isin"], nvda["sedol"]) == (
        "NVIDIA CORP COMMON STOCK USD.001", "67066G104", "US67066G1040", "2379504")
    assert (nvda["shares"], nvda["market_value"]) == (22553641.0, 5153958041.32)
    assert nvda["weight_pct"] == 5.81701798  # published as the fraction 0.0581701798
    assert result.rows[0]["weight_pct"] == 0.01619189


@pytest.mark.parametrize("name, handled", [
    ("Dimensional US Equity Market ETF", True),
    ("DFA Real Estate Securities Portfolio ETF Class Shares", True),
    ("U.S. Small Cap Portfolio: ETF Class Shares", True),
    ("Nuveen Dividend Growth Fund: ETF Class Shares", False),
    ("Vanguard Short-Term Inflation-Protected Securities Index Fund ETF Shares", False),
])
def test_dimensional_funds_by_name(name, handled):
    assert s.DimensionalSource(FakeFetcher({})).handles(s.Etf("X", name, "")) is handled


AVANTIS_HOLDINGS = ("https://www.avantisinvestors.com/avantis-investments/total-holdings/{}/"
                    "?type=etf")
AVUS = s.Etf("AVUS", "Avantis U.S. Equity ETF", "NYSE Arca")


def test_avantis_source_finds_fund_numbers_on_the_fund_list():
    fetcher = FakeFetcher({
        s.AvantisSource.FUND_LIST: fixture("avantis_fund_list.html"),
        AVANTIS_HOLDINGS.format(114): fixture("avantis_avus_holdings.html"),
    })
    source = s.AvantisSource(fetcher)
    source.prepare()
    assert source.numbers == {"AVUS": "114", "AVUV": "119"}  # not the AVUSX mutual fund
    assert source.handles(AVUS) and not source.handles(SPY)
    result = source.fetch(AVUS)
    assert result.url == AVANTIS_HOLDINGS.format(114) and result.as_of == "2026-10-13"
    assert [h["ticker"] for h in result.rows] == ["NVDA", "AAPL", "MSFT", "SKYD/WS", ""]
    nvda = result.rows[0]
    assert (nvda["holding_name"], nvda["cusip"], nvda["isin"], nvda["sedol"]) == (
        "NVIDIA CORP COMMON STOCK USD.001", "67066G104", "US67066G1040", "2379504")
    assert (nvda["shares"], nvda["market_value"], nvda["weight_pct"]) == (
        3737604.0, 854117266.08, 5.75)
    assert (nvda["asset_class"], nvda["sector"], nvda["country"]) == (
        "COMMON STOCK", "INFORMATION TECHNOLOGY", "UNITED STATES")
    warrant, dollars = result.rows[3:]
    assert (warrant["market_value"], warrant["weight_pct"]) == ("", "")
    assert (dollars["holding_name"], dollars["asset_class"], dollars["shares"]) == (
        "US DOLLAR", "", -253137.0)


def test_american_century_source():
    sdsi = s.Etf("SDSI", "American Century Short Duration Strategic Income ETF", "NYSE Arca")
    holdings_url = "https://www.americancentury.com/invest/funds/total-holdings/724/?type=etf"
    source = s.AmericanCenturySource(FakeFetcher({
        s.AmericanCenturySource.FUND_LIST: fixture("american_century_fund_list.html"),
        holdings_url: fixture("american_century_sdsi_holdings.html"),
    }))
    source.prepare()
    assert source.numbers == {"SDSI": "724", "QGRO": "495"}
    result = source.fetch(sdsi)
    assert result.as_of == "2026-10-13"
    note, future, swap = result.rows[:3]
    assert (note["holding_name"], note["cusip"], note["coupon_rate"], note["maturity_date"]) == (
        "US TREASURY N/B 09/28 4.75", "91282CRP8", 4.75, "2028-09-30")
    assert (note["shares"], note["weight_pct"], note["asset_class"]) == (
        53440000.0, 15.92, "U.S. TREASURY NOTES")
    assert (future["ticker"], future["market_value"], future["weight_pct"]) == (
        "TUZ26", 8879.54, 8.87)
    assert swap["asset_class"] == "CREDIT DEFAULT SWAPS"


def test_american_century_without_fund_list_handles_nothing():
    source = s.AmericanCenturySource(FakeFetcher({}))
    source.prepare()
    assert not source.handles(
        s.Etf("SDSI", "American Century Short Duration Strategic Income ETF", ""))


def test_js_array_items_ignore_brackets_inside_strings():
    text = 'x:[{name:"A ]} \\"B\\"",ticker:"A"},{name:"C",ticker:null}],y:1'
    items = s.js_array_items(text, text.index("["))
    assert [s._js_fields(i) for i in items] == [
        {"name": 'A ]} "B"', "ticker": "A"}, {"name": "C", "ticker": ""}]
    with pytest.raises(s.NotAvailable):
        s.js_array_items(text[:20], text.index("["))


def test_fidelity_source():
    ftec = s.Etf("FTEC", "Fidelity MSCI Information Technology Index ETF", "NYSE Arca")
    url = ("https://research2.fidelity.com/fidelity/screeners/etf/etfholdings.asp"
           "?symbol=FTEC&view=Holdings")
    source = s.FidelitySource(FakeFetcher({url: fixture("fidelity_ftec_holdings.html")}))
    assert source.handles(ftec) and not source.handles(SPY)
    result = source.fetch(ftec)
    assert result.as_of == "2026-09-30"
    assert [(h["ticker"], h["holding_name"], h["weight_pct"]) for h in result.rows] == [
        ("NVDA", "NVIDIA Corp", 17.61), ("AAPL", "Apple Inc", 15.95),
        ("", "Net Other Assets", 0.79),
        ("XAKZ26", "E-mini Technology Select Sector Future Dec 26", 0.0),
        ("", "Pivotal Software Inc Cvr", 0.0)]


def test_fidelity_page_without_holdings():
    page = ('<h3 class="num-results">Basket Holdings: 0<span class="timestamp">AS OF 09/30/2026'
            '</span></h3><table class="results-table"><thead><tr><th>Symbol</th><th>Company</th>'
            '<th>Weight</th></tr></thead><tbody><tr class="highlight"><td colspan="3">The search '
            'entry provided does not yield any results</td></tr></tbody></table>')
    assert s.parse_fidelity(page) == ("2026-09-30", [])


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


def test_value_weights_for_equal_weight_funds():
    # 500 positions of 0.2% each: none is large enough for the usual check.
    rows = [holding(f"stock {i}", 2000.0, 0.2) for i in range(500)]
    assert s.add_value_weights(rows) == ""
    assert rows[0]["value_weight_pct"] == 0.2


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


SCHD_DIRECTORY = (fixture("nasdaqtraded_sample.txt").decode().splitlines()[0] + "\n"
                  "Y|SCHD|Schwab U.S. Dividend Equity ETF|P| |Y|100|N||SCHD|SCHD|N\n"
                  "Y|SPY|State Street SPDR S&P 500 ETF Trust|P| |Y|100|N||SPY|SPY|N\n").encode()


def test_run_loads_schwab_pages_in_the_browser_without_zacks(tmp_path):
    http = FakeHttp({s.NASDAQ_SYMBOL_DIRECTORY: SCHD_DIRECTORY})
    browser = FakeBrowser({SCHWAB_URL.format(0): fixture("schwab_schd_page0.html"),
                           SCHWAB_URL.format(1): fixture("schwab_schd_page1.html")})
    args = s.parse_args(["--tickers", "SCHD", "--sources", "schwab", "--no-fallback",
                         "--output-dir", str(tmp_path), "--no-cache"])
    assert s.run(args, http=http, browser=browser) == 0
    assert browser.started and browser.closed
    assert browser.fetcher.requested == [SCHWAB_URL.format(0), SCHWAB_URL.format(1)]
    assert http.fetcher.requested == [s.NASDAQ_SYMBOL_DIRECTORY]
    with (tmp_path / "etf_summary.csv").open() as fh:
        summary = next(csv.DictReader(fh))
    assert (summary["status"], summary["source"], summary["holdings_count"]) == (
        "ok", "schwab", "4")


def test_run_without_a_browser_skips_sources_that_need_one(tmp_path, caplog):
    class NoBrowser(FakeBrowser):
        def start(self):
            raise s.SetupError("could not start the browser (chrome): not installed")

    http = FakeHttp({s.NASDAQ_SYMBOL_DIRECTORY: SCHD_DIRECTORY, SPY_URL: ssga_workbook()})
    args = s.parse_args(["--tickers", "SCHD", "SPY", "--sources", "schwab", "ssga",
                         "--no-fallback", "--output-dir", str(tmp_path), "--no-cache"])
    assert s.run(args, http=http, browser=NoBrowser({})) == 0
    assert "schwab ETFs will have no holdings" in caplog.text
    with (tmp_path / "etf_summary.csv").open() as fh:
        summary = {row["ticker"]: row for row in csv.DictReader(fh)}
    assert summary["SCHD"]["detail"] == "no source covers this ETF"
    assert summary["SPY"]["source"] == "ssga"


def test_run_writes_holdings_and_summary(tmp_path):
    responses = {
        s.NASDAQ_SYMBOL_DIRECTORY: fixture("nasdaqtraded_sample.txt"),
        s.IsharesSource.FUND_LIST: fixture("ishares_fund_list.json"),
        IVV_URL: fixture("ishares_ivv_holdings.csv"),
        SPY_URL: ssga_workbook(),
        QQQ_URL: fixture("invesco_qqq.json"),
        INVESCO_PAGE.format("QQQ"): QQQ_PAGE,
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


def test_sector_lookup_fills_blanks_from_ishares(tmp_path):
    def row(source, ticker, sector, **extra):
        base = {c: "" for c in s.HOLDING_COLUMNS}
        base.update(source=source, ticker=ticker, sector=sector, asset_class="Equity",
                    country="United States", sector_source=source if sector else "")
        base.update(extra)
        return base

    lookup = s.SectorLookup()
    for r in [row("ishares", "BRKB", "Financials"),
              row("ishares", "NVDA", "Information Technology"),
              row("ishares", "", "Health Care", isin="US0000000001"),
              row("ishares", "USD", "Cash and/or Derivatives", asset_class="Cash"),
              row("first_trust", "XOM", "Energy")]:
        lookup.learn(r)
    path = tmp_path / "holdings.csv"
    s.write_csv(path, s.HOLDING_COLUMNS, [
        row("zacks", "BRK.B", ""), row("zacks", "NVDA", ""),
        row("ssga", "", "", isin="US0000000001"), row("zacks", "XOM", ""),
        row("zacks", "USD", ""), row("invesco", "AAPL", "Technology")])
    assert s.backfill_sectors(path, lookup) == (3, 2)
    with path.open() as fh:
        out = list(csv.DictReader(fh))
    assert [(r["sector"], r["sector_source"]) for r in out] == [
        ("Financials", "ishares_lookup"), ("Information Technology", "ishares_lookup"),
        ("Health Care", "ishares_lookup"), ("", ""), ("", ""), ("Technology", "invesco")]


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


GSLC = s.Etf("GSLC", "Goldman Sachs ActiveBeta U.S. Large Cap Equity ETF", "NYSE Arca")
# The request a browser sends for GSLC's fund page, with only the fields read here.
GSLC_REQUEST = {
    "operationName": "getFundsDetail",
    "query": s.GoldmanSachsSource.QUERY,
    "variables": {"fundDetailRequest": {"country": "us", "language": "en", "audience": "individual",
                                        "pvNumber": "PV102394", "shareClassId": "381430503"}},
}


def goldman_fetcher(response: bytes) -> FakeFetcher:
    return FakeFetcher({s.request_key(s.GoldmanSachsSource.URL, GSLC_REQUEST): response})


def test_goldman_sachs_source_posts_the_fund_query():
    fetcher = goldman_fetcher(fixture("goldman_gslc_holdings.json"))
    source = s.GoldmanSachsSource(fetcher)
    assert source.handles(GSLC) and not source.handles(SPY)
    result = source.fetch(GSLC)
    assert fetcher.posted == [
        ("https://am.gs.com/services/funds", GSLC_REQUEST, {"Origin": "https://am.gs.com"})]
    assert result.as_of == "2026-10-19"
    assert [h["ticker"] for h in result.rows] == ["MMM", "ABBV", "GOOGL", "GOOG", "AMZN"]
    googl = result.rows[2]
    assert (googl["holding_name"], googl["isin"], googl["sedol"], googl["cusip"]) == (
        "ALPHABET INC.", "US02079K3059", "BYVY8G0", "")
    assert (googl["shares"], googl["market_value"], googl["weight_pct"]) == (
        1277549.0, 439706814.82, 2.9)


def test_goldman_sachs_query_errors_are_reported():
    refused = b'{"errors": [{"message": "Validation error of type FieldUndefined"}]}'
    with pytest.raises(s.NotAvailable, match="the query was refused: Validation error"):
        s.GoldmanSachsSource(goldman_fetcher(refused)).fetch(GSLC)
    with pytest.raises(s.NotAvailable, match="no such fund"):
        s.GoldmanSachsSource(goldman_fetcher(b'{"data": {"fundsDetail": null}}')).fetch(GSLC)
    other = b'{"data": {"fundsDetail": {"ticker": "GSUS", "allHoldings": null}}}'
    with pytest.raises(s.NotAvailable, match="belong to GSUS"):
        s.GoldmanSachsSource(goldman_fetcher(other)).fetch(GSLC)
