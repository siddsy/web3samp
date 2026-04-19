# Dietitians directory scraper

This repo includes `scrape_dietitians.py`, which automates:

1. Opening the Dietitians of Canada directory page.
2. Clicking **Find** with no filters (to get all available results).
3. Walking every results page via **Next**.
4. Exporting unique rows to **Excel (`.xlsx`)** or CSV.

By default, it writes `dietitians.xlsx` and checks for **602 entries**.

## Setup

```bash
python -m pip install playwright
python -m playwright install chromium
```

## Run (Excel output)

```bash
python scrape_dietitians.py --output dietitians.xlsx --expected-count 602
```

## Optional

```bash
# CSV instead of Excel
python scrape_dietitians.py --output dietitians.csv --expected-count 602

# Show browser UI for debugging selectors
python scrape_dietitians.py --headed --output dietitians.xlsx --expected-count 602
```

If the count does not match 602, the script exits with a warning status so you can review selector changes.
