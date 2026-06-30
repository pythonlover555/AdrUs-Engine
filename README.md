# AdrUs-Engine — Random US Address Scraper

Collects random US addresses from
[bestrandoms.com](https://www.bestrandoms.com/random-address-in-us),
de-duplicates them, and accumulates them into `data/us_address.xlsx`.

The site's **Generate** button is just a form reload, reproducible with a plain
HTTP GET (`?quantity=10` → 10 fresh addresses per call). So the scraper hits
that endpoint directly with `requests` — no browser needed — parses the address
blocks, keeps only ones not already stored, and appends them to the workbook.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python scrape_addresses.py                 # run forever (Ctrl+C to stop)
python scrape_addresses.py --delay 3       # 3s between requests (politer)
python scrape_addresses.py --quantity 10   # addresses per call (default 10)
```

It runs endlessly, saving new addresses as it goes, until you press **Ctrl+C**.
Output goes to `data/us_address.xlsx` (override with `--output PATH`).

## How it works

| File | Role |
| --- | --- |
| [`scrape_addresses.py`](scrape_addresses.py) | CLI entry point / argument parsing |
| [`src/address_scraper/scraper.py`](src/address_scraper/scraper.py) | `AddressFetcher` (HTTP GET) → `parse_addresses` (BeautifulSoup) → `ExcelStore` (dedupe + append) |

Each row has: Street, City, State/Province/Area, Phone Number, Zip Code,
Country Calling Code, Country.

**De-duplication:** a row is treated as a duplicate only if **every field
(including phone number) matches**. The seen-set is rebuilt from the existing
workbook on startup, so de-dup persists across runs — re-running only ever adds
new addresses.

## Notes

- Be polite: the site rate-limits (HTTP 429) under rapid requests. The scraper
  catches errors and retries with a short backoff, but a larger `--delay`
  avoids them in the first place.
- Only scrape data you're allowed to.
