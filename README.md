# ETF-Fetcher

`etf_holdings_scraper.py` downloads the latest holdings of every US-listed ETF
(about 5,800 tickers) and writes them to CSV.

## How it works

1. **ETF list**: every security flagged as an ETF in the
   [Nasdaq Trader symbol directory](https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqtraded.txt),
   which covers all US exchanges.
2. **Holdings**: for each ETF, the script tries its issuer's own website first.
   Issuers post holdings every business day, so this is the freshest public source:

   | Source        | ETFs                                 | Format              | As-of date      |
   |---------------|--------------------------------------|---------------------|-----------------|
   | `ishares`     | iShares (matched by ticker on the iShares fund list) | CSV     | In file         |
   | `ssga`        | names containing "SPDR" or starting with "State Street" | Excel workbook | In file |
   | `invesco`     | names starting with "Invesco"        | JSON API, keyed by the CUSIP on each fund's page | In response |
   | `first_trust` | names starting with "First Trust" or "FT Vest" | HTML page | On page        |
   | `global_x`    | names starting with "Global X"       | CSV, named by date  | In file name    |
   | `vanguard`    | names starting with "Vanguard"       | JSON API, paged     | In response (**month end**) |
   | `proshares`   | names starting with "ProShares"      | HTML fund page      | On page         |
   | `innovator`   | names starting with "Innovator"      | One CSV for all Innovator funds | In file |
   | `pacer`       | names starting with "Pacer"          | CSV                 | In file         |
   | `schwab`      | names starting with "Schwab"         | HTML pages, 100 holdings each, loaded in the browser (the site refuses scripts) | On page |
   | `vaneck`      | names starting with "VanEck"         | Excel workbook      | In file         |
   | `ark`         | names starting with "ARK"            | CSV, linked from the fund's list of documents | In file |
   | `direxion`    | names starting with "Direxion"       | CSV                 | In file, if given |
   | `dimensional` | names starting with "Dimensional" or "DFA", or ending in "Portfolio ETF Class Shares" | CSV, named by date | In file |
   | `avantis`     | Avantis (matched by ticker on the Avantis fund list) | Data embedded in the holdings page | On page |
   | `american_century` | American Century (matched by ticker on its fund list) | Data embedded in the holdings page | On page |
   | `fidelity`    | names starting with "Fidelity"       | HTML page, weights only | On page (**month end**) |
   | `goldman_sachs` | Goldman Sachs (the 48 ETFs listed in the script) | GraphQL query (POST), as its fund pages send it | In response |
   | `zacks`       | everything else, and the fallback when an issuer source fails | Web page, loaded in headless Chrome | **Not published** |

   Together the issuer sources cover about 2,350 of the 5,800 ETFs, including
   most of the largest.

   When a source has no file for a fund, errors out, or returns no rows, the
   next source is tried. The summary file records which source was used and why
   earlier ones failed.

## Usage

Set up once, in the project folder. A virtual environment keeps the packages
inside the project, which Homebrew's Python on macOS requires. Zacks and Schwab
pages are loaded in your installed Google Chrome, running hidden in the
background:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

In each new terminal, run `source .venv/bin/activate` first (the prompt then
starts with `(.venv)`). Then:

```bash
python etf_holdings_scraper.py                         # all ETFs
python etf_holdings_scraper.py --tickers SPY IVV QQQ   # a few ETFs
python etf_holdings_scraper.py --limit 100             # first 100 tickers, alphabetically
python etf_holdings_scraper.py --sources ishares ssga --no-fallback   # issuer data only
```

A full run makes about one request per ETF, at most `--rate` requests per
second to each website (default 2). Some sources need a few more: Vanguard
asks for stocks and bonds separately, Schwab shows 100 holdings per page, ARK
reads the fund page and its list of documents to find the file, and ProShares and VanEck may try a second
address. Innovator needs a single download for all its funds. VanEck's
robots.txt asks for 25 seconds between requests, which the script honors, so
its 90 funds take about 40 minutes in the background of a full run.
Zacks serves most ETFs, and its pages take a few seconds each in a browser, so
expect a full run to take roughly an hour.
Downloads are cached in `.cache/` for six hours, so an interrupted run restarts
quickly; `--refresh` ignores the cache. Ctrl-C stops the run and still writes
the summary for the ETFs finished so far.

| Option | Default | |
|---|---|---|
| `--tickers` | all ETFs | Space or comma separated |
| `--limit N` | | Only the first N ETFs |
| `--sources` | all issuer sources | Any of `ishares ssga invesco first_trust global_x vanguard proshares innovator pacer schwab vaneck ark direxion dimensional avantis american_century fidelity goldman_sachs` |
| `--no-fallback` | | Don't use Zacks. The browser is then only started for Schwab; if it can't start, Schwab ETFs are skipped |
| `--browser` | `chrome` | `chrome` (installed Google Chrome), `chromium` (after `playwright install chromium`), or a path to a browser |
| `--browser-workers` | 4 | Browsers loading Zacks pages in parallel |
| `--show-browser` | | Show the browser windows, for troubleshooting |
| `--output-dir` | `output` | |
| `--cache-dir` / `--no-cache` / `--refresh` | `.cache` | |
| `--workers` | 8 | Parallel downloads |
| `--rate` | 2 | Requests per second to any one website |
| `--user-agent` | a desktop Chrome string | |

## Output

`output/holdings.csv`: one row per position, all ETFs combined.

| Column | |
|---|---|
| `etf_ticker`, `etf_name` | The ETF |
| `source`, `as_of_date` | Where the row came from and the holdings date (blank for Zacks) |
| `holding_name`, `ticker`, `cusip`, `isin`, `sedol`, `other_id` | The position. Each source publishes a different subset of identifiers |
| `asset_class`, `sector`, `country`, `currency` | When the source provides them; blank sectors are filled in from iShares (see below) |
| `sector_source` | Where `sector` came from: the source's name, `ishares_lookup` if it was filled in, or blank |
| `shares` | Shares, contracts, or par value for bonds |
| `price`, `market_value` | When the source provides them |
| `notional_value` | For derivatives, when the source provides it: the exposure an option, future or swap controls, as opposed to its market value |
| `weight_pct` | Percent of net assets as the source publishes it: `8.34` means 8.34% |
| `value_weight_pct` | Percent of net assets computed from `market_value`, unrounded (see below). Blank when the source has no market values |
| `maturity_date`, `coupon_rate` | For bonds, when provided |

`output/etf_summary.csv`: one row per ETF with `status` (`ok` or `no_data`),
`source`, `as_of_date`, `holdings_count`, `weight_sum_pct` (a sanity check: it
should be close to 100), `source_url`, and `detail` (why sources failed).

## Things to know

- **Zacks data is undated and can lag.** Zacks doesn't say what date its
  holdings are from. In spot checks on 2026-10-11, its SPY share counts were
  close to but not equal to State Street's 2026-10-08 file, and its AGG
  holdings still listed mortgage TBAs settling in September, so that snapshot
  was at least about four weeks old. That is still far fresher than SEC N-PORT
  filings (quarter-end, published ~60 days later), but for funds where daily
  data matters, add an issuer source (see below).
- **Zacks blocks plain scripts** with a bot check ("Pardon Our Interruption"),
  which is why its pages are loaded in a real browser: the check clears itself
  there. If a page is still on the check after 30 seconds, it is skipped; after
  five such pages in a row, Zacks is skipped for the rest of the run and those
  ETFs say so in the summary's `detail` column. Successful downloads stay
  cached for six hours, so running again later only fetches what's missing.
- **Published weights are rounded**, which matters for funds with many small
  positions. iShares rounds to 0.01%: 10,837 of AGG's 13,418 bonds show 0.00%,
  though together they are 21.6% of the fund, so its weights add up to 81.88%.
  `value_weight_pct` fixes this by dividing each market value by the fund's total.
  It is filled where the source publishes market values: iShares, First
  Trust, Innovator (every one of its 179 funds), Dimensional, Avantis,
  American Century and Goldman Sachs (all checked), and Vanguard, ProShares,
  Pacer, VanEck, ARK, Direxion and Global X when their files include them. State
  Street and Invesco publish no market values but already give weights to six
  decimals. Schwab rounds its market values to a few digits ("$5.2B"), so they
  are left blank. Fidelity's pages give weights only. Zacks rounds to 0.01%
  (AGG adds up to 192% there) and publishes no market values, so its rows
  can't be corrected. Dimensional publishes weights as fractions (0.0567); they
  are converted to percent like the others.
- **How `value_weight_pct` treats shorts and derivatives.** Short positions and
  written options have negative market values, so they get negative weights and
  each fund still adds up to 100% of net assets. An option's weight is the value
  of the contract, not the exposure it controls: IVVW's written calls are worth
  −2.55% of the fund, but their notional is about −96% (`notional_value`).
  Futures and swaps are worth little after daily settlement. First Trust lists
  them at their notional amount in its market value column; rows whose market
  value is far larger than their published weight allows are treated as
  notional amounts and moved to `notional_value`. American Century does the
  opposite: it weights futures, swaps and currency forwards by their exposure
  (SDSI's Treasury future: 8.87%) while their market value is the day's gain
  or loss, so their `value_weight_pct` is small (0.003%) and the published
  weight is the exposure. If the remaining market values don't add up to the
  fund's net assets (implied by its published weights, within 1%),
  `value_weight_pct` is left blank for that fund and the summary's `detail`
  column says why. Holdings with neither a market value nor a weight, such as
  Avantis' worthless warrants, are left out. Checked against AGG, IVV, IVVW,
  FPE and FTLS: every computed weight is within 0.005 of the published one, the
  most rounding can hide.
- **Vanguard publishes month-end holdings**, a few weeks after the month ends:
  on 2026-10-15 its newest were as of 2026-08-31. That is dated and complete
  (with CUSIP, ISIN, SEDOL and market values), but can be older than Zacks. To
  use Zacks for Vanguard funds instead, leave `vanguard` out of `--sources`.
  Bonds come with their par value in `shares`; pools of mortgages give a range
  of maturities, written as `2027-11-01/2056-10-01`.
- **Fidelity's holdings are month-end, with weights only.** They come from
  the research pages on Fidelity's site: on 2026-10-15, FTEC's were as of
  2026-09-30 and the actively managed FBCG's as of 2026-08-31, since Fidelity's
  active equity ETFs disclose their full holdings a month later (daily they
  publish only a stand-in "tracking basket"). No share counts or market
  values. To use Zacks for Fidelity funds instead, leave `fidelity` out of
  `--sources`.
- **Goldman Sachs funds are looked up by ID.** Its fund pages load "All
  Holdings" with a GraphQL query that takes each fund's product number and
  share class ID; the Fund Finder that lists them is filled in by the browser,
  so the script carries a table of all 48 Goldman Sachs ETFs as of 2026-10-21
  (`GoldmanSachsSource.FUNDS`). A newer fund falls back to Zacks until it is
  added: its fund page address has both IDs, as in
  `am.gs.com/en-us/individual/funds/detail/PV102394/381430503/...`. Checked
  live on 2026-10-21: GSLC, GBND and GPIX returned 429, 733 and 489 holdings
  as of 2026-10-19, with weights adding up to 99.95–100.02%.
- **Leveraged and inverse funds list their exposure.** ProShares gives swaps
  and futures an exposure weight and an exposure value (`notional_value`) but
  no market value; their gains and losses are in the "Net Other Assets" row.
  Weights of a 3x fund such as TQQQ add up to about 300%, and
  `value_weight_pct` is left blank for those funds.
- **Direxion's test sample is representative.** Direxion's site blocked the
  tool used to capture the samples, so the sample uses the columns of an older
  copy of SOXL's file. The live file parses: on 2026-10-21, SOXL returned 44
  holdings as of 2026-10-20. Its weights include swap exposure, so a 3x fund's
  add up to about 300% or more.
- **Sectors are filled in from iShares.** Zacks publishes no sectors and State
  Street leaves them as "-", which left SPY, VOO, VTI and most other common ETFs
  without any. After a run, blank sectors are filled in from how iShares
  classifies the same security (GICS sectors), matched by ISIN, CUSIP or US
  ticker; `sector_source` says `ishares_lookup` for those. In a check against
  real data this covered 96–99.9% of the weight of common US equity ETFs such
  as SPY, VOO, VTI, QQQM and SCHD; XLF was lowest at 87%. Foreign stocks and
  bonds held through Zacks mostly stay blank, since Zacks gives them no ISIN or
  CUSIP. Sectors First Trust and Invesco publish themselves use their own
  names (e.g. Invesco's "Technology" rather than GICS "Information
  Technology"), so normalize those if you compare across sources.
- **Fund-of-funds list their direct holdings.** IVVW's file also has a
  look-through list of the stocks inside the IVV shares it holds; the script
  reads the direct holdings only.
- **Some ETFs have no holdings to list.** ETNs are bank debt, and single-asset
  trusts such as GLD or IBIT just hold the metal or the coin. They usually end
  up as `no_data`.
- **Websites change.** Each source is a few lines of parsing that can break
  when an issuer redesigns its site. Failures fall through to Zacks and show up
  in the summary's `detail` column. Run `pytest` to check the parsers against
  the saved sample files in `tests/fixtures/`.
- Check each website's terms of use before relying on this for anything
  commercial.

## Adding an issuer

Subclass `Source` with `handles(etf)` (does this source cover the ETF?) and
`fetch(etf)` (return a `Holdings`), then add it to `ISSUER_SOURCES`. For CSV,
Excel or HTML tables, `rows_to_holdings()` finds the header row and maps common
column names (Name, Ticker, CUSIP, Weight (%), Shares Held, Market Value...)
to the output columns, so most sources only need a URL and a reader. An issuer
with one CSV per fund at an address made from the ticker needs only a
`CsvSource` subclass with a URL and a name pattern, like `PacerSource`.
Pacer's and Innovator's files share one layout (Date, Account, StockTicker,
CUSIP, SecurityName, Shares, Price, MarketValue, Weightings, NetAssets...),
from their fund administrator, U.S. Bank; other funds it administers may
publish the same columns, which are already mapped. Sites that load holdings
with JavaScript usually fetch them from an API the script can call directly:
find the request in a browser's developer tools (Network tab, Fetch/XHR, then
reload the fund page) and copy it, as was done for Goldman Sachs.
`Source.download()` takes a JSON `payload` for APIs that need a POST.
Remaining issuers by fund count: Corgi, Leverage Shares, GraniteShares,
Defiance, Tradr, JPMorgan, PGIM, Franklin and WisdomTree. Of those checked so
far, Defiance's site blocked the tool used to capture samples, WisdomTree's
holdings files are addressed by document IDs rather than tickers, and
JPMorgan's fund pages load their data with JavaScript.

## Tests

```bash
source .venv/bin/activate
pip install pytest
pytest
```

`tests/test_browser_pool.py` also starts a real headless browser (Chrome, or
the one named in `ETF_TEST_BROWSER`) against a local stand-in for Zacks,
including a bot check that clears itself; it is skipped if no browser starts.
The other tests run offline against sample responses in `tests/fixtures/`, trimmed
from real downloads (a few counts edited to match the trimmed rows, as the
tests note), except the Global X and Direxion holdings files: those sites
blocked the tool used to capture the others, so their samples are
representative. The VanEck workbook is rebuilt in the test from rows of a real
download, and the Avantis and American Century samples keep the real holdings
data embedded in those pages but not the rest of the page around it.
