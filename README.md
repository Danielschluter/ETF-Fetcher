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
   | `zacks`       | everything else, and the fallback when an issuer source fails | HTML page | **Not published** |

   When a source has no file for a fund, errors out, or returns no rows, the
   next source is tried. The summary file records which source was used and why
   earlier ones failed.

## Usage

Set up once, in the project folder. A virtual environment keeps the packages
inside the project, which Homebrew's Python on macOS requires:

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
Zacks serves most ETFs, so expect a full run to take around 30–45 minutes.
Downloads are cached in `.cache/` for six hours, so an interrupted run restarts
quickly; `--refresh` ignores the cache. Ctrl-C stops the run and still writes
the summary for the ETFs finished so far.

| Option | Default | |
|---|---|---|
| `--tickers` | all ETFs | Space or comma separated |
| `--limit N` | | Only the first N ETFs |
| `--sources` | all issuer sources | Any of `ishares ssga invesco first_trust global_x` |
| `--no-fallback` | | Don't use Zacks |
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
| `weight_pct` | Percent of net assets: `8.34` means 8.34% |
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
- **Zacks sometimes answers with a bot check** ("Pardon Our Interruption")
  instead of the page. The script then pauses all Zacks requests for 30, 60,
  then 120 seconds, retrying after each pause. If Zacks keeps blocking, it is
  skipped for the rest of the run and those ETFs say so in the summary's
  `detail` column. Successful downloads stay cached for six hours, so running
  again later only fetches what's missing.
- **Zacks weights are rounded to 0.01%**, so funds with thousands of tiny
  positions can add up to well over 100% (AGG: 192%). Short positions such as
  written options also show positive weights there.
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

The tests run offline against sample responses in `tests/fixtures/`, trimmed
from real downloads, except the Global X holdings file: that site blocked the
tool used to capture the others, so its sample is representative.
