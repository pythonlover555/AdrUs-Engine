"""Entry point: scrape random US addresses into data/us_address.xlsx.

Repeatedly calls the bestrandoms.com "Generate" endpoint over HTTP, saving only
new (non-duplicate) addresses. Runs forever until you stop it with Ctrl+C.

Examples:
    python scrape_addresses.py                 # run forever (Ctrl+C to stop)
    python scrape_addresses.py --delay 2       # 2s between requests
"""

from __future__ import annotations

import argparse
from pathlib import Path

from src.address_scraper.scraper import DEFAULT_OUTPUT, scrape_forever


def main() -> None:
    parser = argparse.ArgumentParser(description="Random US address scraper")
    parser.add_argument(
        "--output", type=Path, default=DEFAULT_OUTPUT,
        help=f"Excel output path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--delay", type=float, default=1.0,
        help="Seconds to wait between requests (default: 1.0)",
    )
    parser.add_argument(
        "--quantity", type=int, default=10,
        help="Addresses requested per call (default: 10)",
    )
    args = parser.parse_args()

    scrape_forever(
        output=args.output,
        delay=args.delay,
        quantity=args.quantity,
    )


if __name__ == "__main__":
    main()
