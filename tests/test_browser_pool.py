"""BrowserPool against a local stand-in for Zacks, in a real headless browser.

Skipped when no browser can be started. Set ETF_TEST_BROWSER to "chrome",
"chromium" or a browser executable path to choose one.
"""

import os
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import etf_holdings_scraper as s  # noqa: E402

HOLDINGS_PAGE = (Path(__file__).parent / "fixtures" / "zacks_spy.html").read_bytes()

# Like Imperva's check: sets a cookie from JavaScript, then reloads the page.
BOT_CHECK_PAGE = b"""<!DOCTYPE html><html><head>
<noscript><title>Pardon Our Interruption</title></noscript>
<script>setTimeout(function () {
  document.cookie = "reese84=passed; path=/"; location.reload(); }, 300);</script>
</head><body></body></html>"""

STUCK_PAGE = b"""<!DOCTYPE html><html><head>
<noscript><title>Pardon Our Interruption</title></noscript></head><body></body></html>"""


class FakeZacks(BaseHTTPRequestHandler):
    def do_GET(self):
        passed = "reese84=passed" in (self.headers.get("Cookie") or "")
        if self.path.startswith("/funds/etf/SPY/"):
            body = HOLDINGS_PAGE if passed else BOT_CHECK_PAGE
        elif self.path.startswith("/funds/etf/STUCK"):
            body = STUCK_PAGE
        else:
            body = b"<html><head><title>No holdings</title></head><body></body></html>"
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass


def browser_for_tests() -> str:
    if os.environ.get("ETF_TEST_BROWSER"):
        return os.environ["ETF_TEST_BROWSER"]
    if Path("/opt/pw-browsers/chromium").exists():
        return "/opt/pw-browsers/chromium"
    return "chrome"


@pytest.fixture(scope="module")
def server():
    httpd = ThreadingHTTPServer(("127.0.0.1", 0), FakeZacks)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{httpd.server_port}"
    httpd.shutdown()


@pytest.fixture
def pool():
    pool = s.BrowserPool(workers=2, rate=50, browser=browser_for_tests())
    pool.BOT_CHECK_WAIT = 3
    try:
        pool.start()
    except s.SetupError as exc:
        pytest.skip(f"no browser available: {exc}")
    yield pool
    pool.close()


def test_page_behind_bot_check_loads_once_it_clears(server, pool):
    page = pool.get(f"{server}/funds/etf/SPY/holding").decode()
    assert len(s.parse_zacks(page)) == 4


def test_page_without_holdings_returns_without_waiting(server, pool):
    page = pool.get(f"{server}/funds/etf/NONE/holding").decode()
    with pytest.raises(s.NotAvailable):
        s.parse_zacks(page)


def test_bot_check_that_never_clears(server, pool):
    pool.GIVE_UP_AFTER = 2
    for _ in range(2):
        with pytest.raises(s.Blocked, match="did not clear"):
            pool.get(f"{server}/funds/etf/STUCK/holding")
    with pytest.raises(s.Blocked, match="skipped"):
        pool.get(f"{server}/funds/etf/SPY/holding")


def test_pages_load_in_parallel(server, pool):
    from concurrent.futures import ThreadPoolExecutor

    with ThreadPoolExecutor(4) as executor:
        pages = list(executor.map(pool.get, [f"{server}/funds/etf/SPY/holding"] * 6))
    assert all(len(s.parse_zacks(p.decode())) == 4 for p in pages)


def test_missing_browser_is_a_setup_error():
    pool = s.BrowserPool(workers=1, rate=1, browser="/nonexistent/browser")
    with pytest.raises(s.SetupError, match="could not start the browser"):
        pool.start()
