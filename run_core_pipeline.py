#!/usr/bin/env python3
"""Download CORE/ICORE rankings and run the local workbook pipeline.

The CORE portal export button is a GET form submit:

    /conf-ranks/?search=&by=all&source=ICORE2026&sort=atitle&page=1&do=Export

This script uses that export endpoint to refresh CORE.csv, then runs:

    create_core_area_sheets.py
    scrape_core_area_deadlines.py
"""

from __future__ import annotations

import argparse
import csv
import os
import subprocess
import sys
from datetime import date
from pathlib import Path
from urllib import error, parse, request


PORTAL_EXPORT_URL = "https://portal.core.edu.au/conf-ranks/"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0 Safari/537.36"
)


def default_area() -> str:
    return os.environ.get("CORE_AREA", "4612")


def default_source() -> str:
    return os.environ.get("CORE_SOURCE", "ICORE2026")


def default_years() -> list[str]:
    current_year = date.today().year
    return [str(current_year + 1), str(current_year)]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download CORE.csv from the CORE portal and run the full area/deadline pipeline."
    )
    parser.add_argument(
        "area",
        nargs="?",
        default=default_area(),
        help="area code to enrich with deadlines, default: CORE_AREA or 4612",
    )
    parser.add_argument(
        "--source",
        default=default_source(),
        help="CORE/ICORE source to export, default: CORE_SOURCE or ICORE2026",
    )
    parser.add_argument(
        "--portal-url",
        default=PORTAL_EXPORT_URL,
        help=f"CORE portal export form URL, default: {PORTAL_EXPORT_URL}",
    )
    parser.add_argument(
        "--core-csv",
        default="CORE.csv",
        help="downloaded CORE CSV path, default: CORE.csv",
    )
    parser.add_argument(
        "--area-workbook",
        default="CORE_by_area.xlsx",
        help="area workbook path, default: CORE_by_area.xlsx",
    )
    parser.add_argument(
        "--deadline-workbook",
        default="deadline.xlsx",
        help="deadline workbook path, default: deadline.xlsx",
    )
    parser.add_argument(
        "--search",
        default="",
        help="portal search text, default: empty",
    )
    parser.add_argument(
        "--by",
        default="all",
        choices=["all", "title", "acronym", "rank", "for"],
        help="portal search field, default: all",
    )
    parser.add_argument(
        "--sort",
        default="atitle",
        help="portal sort key, default: atitle",
    )
    parser.add_argument(
        "--page",
        default="1",
        help="portal page parameter used for the export form, default: 1",
    )
    parser.add_argument(
        "--area-start-column",
        type=int,
        default=7,
        help="1-based CSV column where area codes start, default: 7",
    )
    parser.add_argument(
        "--include-ranks",
        nargs="+",
        default=["A*", "A", "B"],
        help="CORE ranks to scrape for deadlines, default: A* A B. Use ALL for every row.",
    )
    parser.add_argument(
        "--years",
        nargs="+",
        default=None,
        help="deadline years to try in order, default: next year then current year",
    )
    parser.add_argument(
        "--scrape-timeout",
        type=int,
        default=15,
        help="deadline scraper HTTP timeout in seconds, default: 15",
    )
    parser.add_argument(
        "--scrape-delay",
        type=float,
        default=0.75,
        help="deadline scraper delay between requests, default: 0.75",
    )
    parser.add_argument(
        "--max-bytes",
        type=int,
        default=2_000_000,
        help="maximum bytes read per scraped page, default: 2000000",
    )
    parser.add_argument(
        "--queries-per-year",
        type=int,
        default=4,
        help="search queries per conference/year for deadline scraping, default: 4",
    )
    parser.add_argument(
        "--max-search-results",
        type=int,
        default=6,
        help="search results kept from each query, default: 6",
    )
    parser.add_argument(
        "--pages-per-year",
        type=int,
        default=6,
        help="top search-result pages inspected per conference/year, default: 6",
    )
    parser.add_argument(
        "--child-links",
        type=int,
        default=3,
        help="same-site CFP/deadline links followed from each result page, default: 3",
    )
    parser.add_argument(
        "--scrape-cache",
        default=".core_deadline_scrape_cache.json",
        help='deadline scraper HTTP cache path, default: .core_deadline_scrape_cache.json. Use "" to disable.',
    )
    parser.add_argument(
        "--deadline-limit",
        type=int,
        default=0,
        help="limit deadline rows processed, default: all",
    )
    parser.add_argument(
        "--deadline-start-row",
        type=int,
        default=2,
        help="first row for deadline scraping, default: 2",
    )
    parser.add_argument(
        "--deadline-overwrite",
        action="store_true",
        help="replace existing URL/deadline/page-format values in the deadline workbook",
    )
    parser.add_argument(
        "--no-sort",
        action="store_true",
        help="keep original area sheet order instead of sorting by upcoming deadline",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="download timeout in seconds for the CORE export, default: 30",
    )
    parser.add_argument(
        "--skip-download",
        action="store_true",
        help="use existing CORE.csv instead of downloading from the portal",
    )
    parser.add_argument(
        "--skip-area-sheets",
        action="store_true",
        help="use existing area workbook instead of regenerating it",
    )
    parser.add_argument(
        "--skip-deadlines",
        action="store_true",
        help="stop after downloading CORE.csv and creating the area workbook",
    )
    parser.add_argument(
        "--dry-run-deadlines",
        action="store_true",
        help="run the deadline scraper without writing the deadline workbook",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="print detailed downloader and deadline-scraper logs",
    )
    return parser.parse_args()


def build_export_url(args: argparse.Namespace) -> str:
    query = {
        "search": args.search,
        "by": args.by,
        "source": args.source,
        "sort": args.sort,
        "page": args.page,
        "do": "Export",
    }
    return args.portal_url + "?" + parse.urlencode(query)


def decode_export(raw: bytes, content_type: str) -> str:
    if raw.startswith(b"%PDF"):
        raise ValueError(
            "The portal returned a PDF, not CSV. The current pipeline expects the CSV export."
        )

    charset = "utf-8-sig"
    if "charset=" in content_type:
        charset = content_type.split("charset=", 1)[1].split(";", 1)[0].strip()

    try:
        return raw.decode(charset)
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def validate_core_csv(text: str) -> list[list[str]]:
    if "<html" in text[:500].casefold():
        raise ValueError("The portal returned HTML instead of CSV. The export request may have failed.")

    rows = list(csv.reader(text.splitlines()))
    if not rows:
        raise ValueError("The exported CSV is empty.")

    widths = {len(row) for row in rows}
    if widths != {9}:
        preview = ", ".join(str(width) for width in sorted(widths))
        raise ValueError(f"Expected 9 CSV columns, but found column widths: {preview}")

    return rows


def download_core_csv(args: argparse.Namespace) -> int:
    export_url = build_export_url(args)
    output_path = Path(args.core_csv)

    if args.verbose:
        print(f"Downloading CORE export: {export_url}")

    req = request.Request(
        export_url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/csv,text/plain,application/octet-stream,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )

    try:
        with request.urlopen(req, timeout=args.timeout) as response:
            raw = response.read()
            content_type = response.headers.get("Content-Type", "")
    except (error.HTTPError, error.URLError, TimeoutError, OSError) as exc:
        raise RuntimeError(f"Could not download CORE export: {exc}") from exc

    text = decode_export(raw, content_type)
    rows = validate_core_csv(text)
    output_path.write_text(text, encoding="utf-8", newline="")
    print(f"Wrote {len(rows)} CORE row(s) to {output_path}")
    return len(rows)


def run_command(command: list[str]) -> None:
    print("+ " + " ".join(command))
    subprocess.run(command, check=True)


def create_area_workbook(args: argparse.Namespace) -> None:
    run_command(
        [
            sys.executable,
            "create_core_area_sheets.py",
            args.core_csv,
            "-o",
            args.area_workbook,
            "--area-start-column",
            str(args.area_start_column),
        ]
    )


def scrape_deadlines(args: argparse.Namespace) -> None:
    years = [str(year) for year in (args.years or default_years())]
    command = [
        sys.executable,
        "scrape_core_area_deadlines.py",
        args.area,
        "-i",
        args.area_workbook,
        "-o",
        args.deadline_workbook,
        "--include-ranks",
        *args.include_ranks,
        "--years",
        *years,
        "--start-row",
        str(args.deadline_start_row),
        "--timeout",
        str(args.scrape_timeout),
        "--delay",
        str(args.scrape_delay),
        "--max-bytes",
        str(args.max_bytes),
        "--queries-per-year",
        str(args.queries_per_year),
        "--max-search-results",
        str(args.max_search_results),
        "--pages-per-year",
        str(args.pages_per_year),
        "--child-links",
        str(args.child_links),
        "--cache",
        args.scrape_cache,
    ]

    if args.deadline_limit:
        command.extend(["--limit", str(args.deadline_limit)])
    if args.deadline_overwrite:
        command.append("--overwrite")
    if args.no_sort:
        command.append("--no-sort")
    if args.dry_run_deadlines:
        command.append("--dry-run")
    if args.verbose:
        command.append("--verbose")

    run_command(command)


def main() -> int:
    args = parse_args()

    if not args.skip_download:
        download_core_csv(args)
    else:
        print(f"Skipping download; using existing {args.core_csv}")

    if not args.skip_area_sheets:
        create_area_workbook(args)
    else:
        print(f"Skipping area workbook generation; using existing {args.area_workbook}")

    if not args.skip_deadlines:
        scrape_deadlines(args)
    else:
        print("Skipping deadline scraping.")

    print("Pipeline complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
