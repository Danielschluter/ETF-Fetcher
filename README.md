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
   | `invesco`     | names starting with "Invesco"        | JSON API            | In response     |
   | `first_trust` | names starting with "First Trust" or "FT Vest" | HTML page | On page        |
   | `global_x`    | names starting with "Global X"       | CSV, named by date  | In file name    |
   | `zacks`       | everything else, and the fallback when an issuer source fails | Web page, loaded in headless Chrome | **Not published** |

   When a source has no file for a fund, errors out, or returns no rows, the
   next source is tried. The summary file records which source was used and why
   earlier ones failed.

## Usage

Set up once, in the project folder. A virtual environment keeps the packages
inside the project, which Homebrew's Python on macOS requires. Zacks pages are
loaded in your installed Google Chrome, running hidden in the background:

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

A full run makes about one request per ETF (two or three for a few Global X
funds), at most `--rate` requests per second to each website (default 2).
Zacks serves most ETFs, and its pages take a few seconds each in a browser, so
expect a full run to take roughly an hour.
Downloads are cached in `.cache/` for six hours, so an interrupted run restarts
quickly; `--refresh` ignores the cache. Ctrl-C stops the run and still writes
the summary for the ETFs finished so far.

| Option | Default | |
|---|---|---|
| `--tickers` | all ETFs | Space or comma separated |
| `--limit N` | | Only the first N ETFs |
| `--sources` | all issuer sources | Any of `ishares ssga invesco first_trust global_x` |
| `--no-fallback` | | Don't use Zacks (no browser needed) |
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
| `asset_class`, `sector`, `country`, `currency` | When the source provides them |
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
  It is filled where the source publishes market values: iShares and First
  Trust (both checked), and Global X if its files include them. State Street
  and Invesco publish no market values but already give weights to six
  decimals. Zacks rounds to 0.01% (AGG adds up to 192% there) and publishes no
  market values, so its rows can't be corrected.
- **How `value_weight_pct` treats shorts and derivatives.** Short positions and
  written options have negative market values, so they get negative weights and
  each fund still adds up to 100% of net assets. An option's weight is the value
  of the contract, not the exposure it controls: IVVW's written calls are worth
  −2.55% of the fund, but their notional is about −96% (`notional_value`).
  Futures and swaps are worth little after daily settlement. First Trust lists
  them at their notional amount in its market value column; rows whose market
  value is far from what their published weight allows are treated as notional
  amounts and moved to `notional_value`. If the remaining market values don't
  add up to the fund's net assets (implied by its published weights, within 1%),
  `value_weight_pct` is left blank for that fund and the summary's `detail`
  column says why. Checked against AGG, IVV, IVVW, FPE and FTLS: every computed
  weight is within 0.005 of the published one, the most rounding can hide.
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
to the output columns, so most sources only need a URL and a reader. Good
candidates by fund count: Vanguard, ProShares, Direxion, Schwab, JPMorgan,
Fidelity, VanEck, WisdomTree and Innovator.

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
from real downloads, except the Global X holdings file: that site blocked the
tool used to capture the others, so its sample is representative.
