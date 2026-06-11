# CORE Area Sheets and Deadline Scraper

This folder contains the scripts and static site for a CORE conference deadline tracker:

- `run_core_pipeline.py`: downloads the CORE/ICORE export from the portal, creates area sheets, then scrapes deadlines for one area.
- `create_core_area_sheets.py`: reads `CORE.csv` and creates one Excel sheet per CORE area code.
- `scrape_core_area_deadlines.py`: reads one area sheet, searches the web for upcoming conference pages, adds URL/deadline/page-format columns, and sorts by upcoming deadline.
- `build_site_data.py`: converts `deadline.xlsx` into `site/data.js` for the static website.
- `build_site_config.py`: writes runtime UI settings such as countdown thresholds into `site/config.js`.

When `--years` is not supplied, the deadline scraper tries next year first, then the current year based on the day you run it. It skips CORE `C` rows by default. The output sheet is reordered by upcoming deadlines, with conferences that have the nearest future deadline at the top.

## Requirements

Use Python 3 with `openpyxl` installed:

```bash
python3 -m pip install -r requirements.txt
```

The scraper uses only Python standard-library networking/HTML parsing plus `openpyxl`.

## 1. Run the Full Pipeline

To refresh `CORE.csv` from the CORE portal, create `CORE_by_area.xlsx`, and scrape deadlines for area `4612`:

```bash
./run_core_pipeline.py 4612
```

This uses the CORE portal export URL:

```text
https://portal.core.edu.au/conf-ranks/?search=&by=all&source=ICORE2026&sort=atitle&page=1&do=Export
```

Default outputs:

```text
CORE.csv
CORE_by_area.xlsx
deadline.xlsx
```

### Full Pipeline Options

```bash
./run_core_pipeline.py [area] [options]
```

| Option | Default | Description |
| --- | --- | --- |
| `area` | `CORE_AREA` or `4612` | Area code to enrich with deadlines. |
| `--source` | `CORE_SOURCE` or `ICORE2026` | CORE/ICORE source to export. |
| `--portal-url` | `https://portal.core.edu.au/conf-ranks/` | CORE portal export form URL. |
| `--core-csv` | `CORE.csv` | Downloaded CORE CSV path. |
| `--area-workbook` | `CORE_by_area.xlsx` | Area workbook path. |
| `--deadline-workbook` | `deadline.xlsx` | Deadline workbook path. |
| `--search` | empty | Portal search text. |
| `--by` | `all` | Portal search field: `all`, `title`, `acronym`, `rank`, or `for`. |
| `--sort` | `atitle` | Portal sort key. |
| `--page` | `1` | Portal page parameter used with the export form. |
| `--area-start-column` | `7` | 1-based CSV column where area codes start. |
| `--include-ranks` | `A* A B` | CORE ranks to scrape for deadlines. Use `ALL` for every row. |
| `--years` | next year, current year | Deadline years to try, in order. |
| `--scrape-timeout` | `15` | Deadline scraper HTTP timeout in seconds. |
| `--scrape-delay` | `0.75` | Delay between deadline scraper HTTP requests. |
| `--max-bytes` | `2000000` | Maximum bytes read per scraped page. |
| `--queries-per-year` | `4` | Search queries per conference/year for deadline scraping. |
| `--max-search-results` | `6` | Search results kept from each query. |
| `--pages-per-year` | `6` | Top search-result pages inspected per conference/year. |
| `--child-links` | `3` | Same-site CFP/deadline links followed from each result page. |
| `--scrape-cache` | `.core_deadline_scrape_cache.json` | Deadline scraper HTTP cache path. Use `""` to disable. |
| `--deadline-limit` | `0` | Limit deadline rows processed. `0` means all matching rows. |
| `--deadline-start-row` | `2` | First row for deadline scraping. |
| `--deadline-overwrite` | off | Replace existing scraped values. |
| `--no-sort` | off | Keep original order instead of sorting by upcoming deadline. |
| `--timeout` | `30` | Download timeout in seconds for the CORE export. |
| `--skip-download` | off | Use existing `CORE.csv` instead of downloading. |
| `--skip-area-sheets` | off | Use existing area workbook instead of regenerating it. |
| `--skip-deadlines` | off | Stop after downloading and creating area sheets. |
| `--dry-run-deadlines` | off | Run the deadline scraper without writing the deadline workbook. |
| `--verbose` | off | Print detailed downloader and scraper logs. |

Examples:

```bash
./run_core_pipeline.py 4612
./run_core_pipeline.py 4602 --deadline-workbook deadline_4602.xlsx
./run_core_pipeline.py 4612 --deadline-limit 10 --verbose
./run_core_pipeline.py 4612 --skip-download --skip-area-sheets
CORE_AREA=4602 CORE_SOURCE=ICORE2026 ./run_core_pipeline.py
```

## 2. Build and View the Website

After `deadline.xlsx` is generated, build the website data:

```bash
./build_site_data.py deadline.xlsx
./build_site_config.py --urgent-days 10 --soon-days 30
```

Then open:

```text
site/index.html
```

The website is static and reads `site/data.js`, so it can be opened directly from disk or hosted with GitHub Pages. It includes area selection, search, rank filtering, deadline filtering, sorting, countdowns, URL links, and page-limit/format details.
It also reads `site/config.js` for UI thresholds such as when a deadline becomes urgent.

To build site data from a custom workbook:

```bash
./build_site_data.py deadline_4612.xlsx -o site/data.js
```

To change countdown highlighting without editing the UI:

```bash
./build_site_config.py --urgent-days 7 --soon-days 21
```

### GitHub Pages Hosting

The repository includes a GitHub Actions workflow at:

```text
.github/workflows/update-site.yml
```

In GitHub, enable Pages with:

```text
Settings -> Pages -> Build and deployment -> Source: GitHub Actions
```

The workflow deploys the static files in `site/` to GitHub Pages.

### Quarterly Auto-Refresh

The same workflow runs every three months using this cron schedule:

```text
15 3 1 */3 *
```

On each scheduled run it:

1. Downloads the configured CORE/ICORE export into `CORE.csv`.
2. Rebuilds `CORE_by_area.xlsx`.
3. Scrapes Researchr-style conference pages for the configured area.
4. Rebuilds `deadline.xlsx`.
5. Regenerates `site/data.js` and `site/config.js`.
6. Commits changed generated data back to `main`.
7. Deploys the refreshed `site/` folder to GitHub Pages.

You can also run it manually from the GitHub Actions tab with `workflow_dispatch`. Manual inputs override repository variables for that run.

### Workflow Configuration

Set these as GitHub Actions manual inputs or repository variables when you want the scheduled job to use different values:

| Variable / input | Default | Description |
| --- | --- | --- |
| `CORE_AREA` / `area` | `4612` | Area sheet to publish. |
| `CORE_SOURCE` / `source` | `ICORE2026` | CORE/ICORE export source. |
| `CORE_INCLUDE_RANKS` / `include_ranks` | `A*,A,B` | Comma-separated ranks to scrape. Use `ALL` for every row. |
| `CORE_SCRAPE_YEARS` / `years` | empty | Space-separated years. Empty means next year then current year. |
| `CORE_QUERIES_PER_YEAR` / `queries_per_year` | `0` | Search queries per conference/year in CI. |
| `CORE_PAGES_PER_YEAR` / `pages_per_year` | `4` | Search-result pages inspected per conference/year in CI. |
| `CORE_CHILD_LINKS` / `child_links` | `8` | Same-site CFP/deadline links followed in CI. |
| `CORE_URGENT_DAYS` / `urgent_days` | `10` | Number of days below which countdowns are highlighted as urgent. |
| `CORE_SOON_DAYS` / `soon_days` | `30` | Number of days below which countdowns are marked as soon. |

## 3. Create Area Sheets Only

Run this first if `CORE_by_area.xlsx` does not exist or if `CORE.csv` changed:

```bash
./create_core_area_sheets.py
```

Default input:

```text
CORE.csv
```

Default output:

```text
CORE_by_area.xlsx
```

Each conference is copied into every area sheet it belongs to. Rows are sorted by CORE ranking, with `A*` first, then `A`, `B`, `C`, and lower/non-ranked entries after that.

### Area Sheet Options

```bash
./create_core_area_sheets.py [input] [options]
```

| Option | Default | Description |
| --- | --- | --- |
| `input` | `CORE.csv` | Input CORE CSV file. |
| `-o`, `--output` | `CORE_by_area.xlsx` | Output Excel workbook. |
| `--area-start-column` | `7` | 1-based CSV column where area codes start. For the current CSV, columns 7-9 are area columns. |

Examples:

```bash
./create_core_area_sheets.py CORE.csv -o CORE_by_area.xlsx
./create_core_area_sheets.py CORE.csv --area-start-column 7
```

## 4. Scrape Deadlines for an Area Only

Run the scraper with the area code you want:

```bash
./scrape_core_area_deadlines.py 4612
```

Default input:

```text
CORE_by_area.xlsx
```

Default output:

```text
deadline.xlsx
```

The output workbook contains only the processed area sheet, for example a single `4612` sheet. The script appends:

```text
URL
Abstract Deadline
Submission Deadline
Next Deadline
Countdown
Page Limit / Format
```

By default, the scraper processes only CORE `A*`, `A`, and `B` rows. It does not scrape CORE `C` rows unless you ask it to.

The `Next Deadline` column is the earlier upcoming date found from `Abstract Deadline` and `Submission Deadline`. Rows with upcoming deadlines are sorted first, nearest deadline first. Rows with only past deadlines come after those, and rows without parseable deadlines are placed at the bottom.

The `Countdown` column is computed from `Next Deadline` against the current local date and time. Since scraped deadlines usually have no exact hour, the countdown treats each deadline as the end of that date.

The `Page Limit / Format` column captures author-instruction hints when available, such as `19 pages`, `8-12 pages`, `LNCS format`, `ACM format`, or `IEEE format`.

For software-engineering conferences, the scraper tries Researchr first using URLs like `https://conf.researchr.org/home/fse-2027` and `https://conf.researchr.org/dates/fse-2027`, then follows same-site important-date/submission/author links before falling back to general web search.

### Deadline Scraper Options

```bash
./scrape_core_area_deadlines.py [area] [options]
```

| Option | Default | Description |
| --- | --- | --- |
| `area` | `CORE_AREA` or `4612` | Area code / worksheet name to enrich. |
| `-i`, `--input` | `CORE_by_area.xlsx` | Input workbook containing area sheets. |
| `-o`, `--output` | `deadline.xlsx` | Output workbook. |
| `--sheet` | empty | Worksheet to enrich. Overrides the positional `area`. |
| `--include-ranks` | `A* A B` | CORE ranks to scrape. Use `ALL` to scrape every row. |
| `--years` | next year, current year | Conference years to try, in order. Earlier values are preferred. |
| `--limit` | `0` | Maximum number of rows to process. `0` means all matching rows. |
| `--start-row` | `2` | First worksheet row to process. Row 1 is the header. |
| `--only-acronym` | empty | Process only one acronym, such as `FSE`, for debugging or rerunning one conference. |
| `--overwrite` | off | Replace existing URL/deadline values. Without this, existing filled values are kept. |
| `--dry-run` | off | Run without writing the output workbook. |
| `--no-sort` | off | Keep the original sheet order instead of sorting by upcoming deadline. |
| `--timeout` | `15` | HTTP timeout in seconds. |
| `--delay` | `0.75` | Delay between HTTP requests, in seconds. |
| `--max-bytes` | `2000000` | Maximum bytes read per fetched page. |
| `--queries-per-year` | `4` | Number of search queries to try per conference/year. |
| `--max-search-results` | `6` | Search results kept from each query. |
| `--pages-per-year` | `6` | Top search-result pages inspected per conference/year. |
| `--child-links` | `3` | Same-site CFP/deadline links followed from each result page. |
| `--cache` | `.core_deadline_scrape_cache.json` | HTTP cache file. Use `--cache ""` to disable caching. |
| `--verbose` | off | Print search and fetch details. |

## Common Commands

Scrape area `4612` into `deadline.xlsx`:

```bash
./scrape_core_area_deadlines.py 4612
```

Scrape another area:

```bash
./scrape_core_area_deadlines.py 4602
```

Write to a custom workbook:

```bash
./scrape_core_area_deadlines.py 4612 -o deadline_4612.xlsx
```

Scrape only the first 10 matching rows for a quick test:

```bash
./scrape_core_area_deadlines.py 4612 --limit 10
```

Preview without writing a workbook:

```bash
./scrape_core_area_deadlines.py 4612 --dry-run --limit 5
```

Include CORE `C` conferences too:

```bash
./scrape_core_area_deadlines.py 4612 --include-ranks "A*" A B C
```

Scrape every rank, including national/regional/unranked entries:

```bash
./scrape_core_area_deadlines.py 4612 --include-ranks ALL
```

Prefer only 2027 pages:

```bash
./scrape_core_area_deadlines.py 4612 --years 2027
```

Prefer 2028, then 2027, then 2026:

```bash
./scrape_core_area_deadlines.py 4612 --years 2028 2027 2026
```

Resume from a later row:

```bash
./scrape_core_area_deadlines.py 4612 --start-row 25
```

Overwrite existing deadline columns:

```bash
./scrape_core_area_deadlines.py 4612 --overwrite
```

Rerun just one conference acronym:

```bash
./scrape_core_area_deadlines.py 4612 --only-acronym FSE --overwrite --verbose
```

Disable the HTTP cache:

```bash
./scrape_core_area_deadlines.py 4612 --cache ""
```

Print detailed search/fetch logs:

```bash
./scrape_core_area_deadlines.py 4612 --verbose
```

Keep the original CORE ranking order instead of sorting by deadline:

```bash
./scrape_core_area_deadlines.py 4612 --no-sort
```

## Notes

- The scraper relies on live search results and conference websites, so not every row will produce a URL or deadline.
- Deadline formats vary a lot across websites. The script looks for abstract and paper/submission deadline language near the configured years.
- Before searching, parenthesized notes in CORE conference names are removed, so entries like `was ESEC/FSE` do not pollute the search query.
- Page limits and format instructions also vary; the script looks for page-count and template/format language on the same pages it fetches for deadlines.
- After scraping, the output is sorted by `Next Deadline` unless `--no-sort` is used.
- `deadline.xlsx` is regenerated from the input workbook each run. Use a custom `-o` name if you want to keep several versions.
- The cache file speeds up reruns and avoids repeatedly fetching the same pages.
