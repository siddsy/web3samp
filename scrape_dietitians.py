#!/usr/bin/env python3
"""Export Dietitians of Canada directory results to CSV or XLSX.

By default this script expects 602 entries and writes `dietitians.xlsx`.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import time
import zipfile
from typing import List
from xml.sax.saxutils import escape

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

DIRECTORY_URL = (
    "https://members.dietitians.ca/Web/Web/Membership/Directory/"
    "Find_a_Dietitian_Directory.aspx?hkey=81abbd4d-e041-4223-8723-7af2d390f04c"
)


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _click_find_button(page) -> None:
    candidates = [
        page.get_by_role("button", name=re.compile(r"^\s*Find\s*$", re.I)),
        page.locator("input[type='submit'][value*='Find' i]"),
        page.locator("button:has-text('Find')"),
        page.locator("a:has-text('Find')"),
    ]

    for locator in candidates:
        try:
            if locator.first.is_visible(timeout=2000):
                locator.first.click()
                return
        except Exception:
            continue

    raise RuntimeError("Could not locate the Find button on the directory page.")


def _extract_largest_table(page) -> tuple[List[str], List[List[str]]]:
    script = """
() => {
  const tables = Array.from(document.querySelectorAll('table'));
  let best = null;

  for (const table of tables) {
    const rows = Array.from(table.querySelectorAll('tr'));
    if (rows.length < 2) continue;

    const parsed = rows.map((tr) => {
      const cells = Array.from(tr.querySelectorAll('th,td'));
      return cells.map((c) => c.innerText.replace(/\s+/g, ' ').trim());
    }).filter((r) => r.some(Boolean));

    if (!parsed.length) continue;

    const score = parsed.length * (parsed[0]?.length || 1);
    if (!best || score > best.score) {
      best = { score, rows: parsed };
    }
  }

  if (!best) return null;

  const rows = best.rows;
  const first = rows[0] || [];
  const hasHeader = first.some((x) => /name|city|province|postal|email|phone|language|concern/i.test(x));
  const headers = hasHeader ? first : first.map((_, i) => `column_${i + 1}`);
  const data = hasHeader ? rows.slice(1) : rows;

  return { headers, data };
}
"""
    result = page.evaluate(script)
    if not result:
        return [], []

    headers = [_normalize(x) for x in result["headers"]]
    rows = [[_normalize(c) for c in row] for row in result["data"]]
    rows = [row for row in rows if any(row)]
    return headers, rows


def _click_next(page) -> bool:
    candidates = [
        page.get_by_role("link", name=re.compile(r"^\s*Next\s*$", re.I)),
        page.get_by_role("button", name=re.compile(r"^\s*Next\s*$", re.I)),
        page.locator("a:has-text('Next')"),
        page.locator("button:has-text('Next')"),
        page.locator("a[aria-label*='Next' i]"),
    ]

    for locator in candidates:
        try:
            element = locator.first
            if not element.is_visible(timeout=1500):
                continue
            classes = (element.get_attribute("class") or "").lower()
            aria_disabled = (element.get_attribute("aria-disabled") or "").lower()
            if "disabled" in classes or aria_disabled == "true":
                continue
            element.click()
            return True
        except Exception:
            continue

    return False


def scrape(headless: bool = True, max_pages: int = 1000) -> tuple[List[str], List[List[str]]]:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        page.goto(DIRECTORY_URL, wait_until="domcontentloaded", timeout=120000)
        _click_find_button(page)

        seen = set()
        all_rows: List[List[str]] = []
        headers: List[str] = []

        for page_no in range(1, max_pages + 1):
            try:
                page.wait_for_timeout(1200)
            except PlaywrightTimeoutError:
                pass

            current_headers, rows = _extract_largest_table(page)
            if current_headers and not headers:
                headers = current_headers

            new_rows = 0
            for row in rows:
                key = "|".join(row)
                if key and key not in seen:
                    seen.add(key)
                    all_rows.append(row)
                    new_rows += 1

            print(f"Page {page_no}: found {len(rows)} rows, added {new_rows} new rows")

            if not _click_next(page):
                break

            time.sleep(1)

        browser.close()

    if not headers and all_rows:
        width = max(len(r) for r in all_rows)
        headers = [f"column_{i + 1}" for i in range(width)]

    if not headers and not all_rows:
        raise RuntimeError(
            "No tabular results were detected after pressing Find. "
            "Open the site in headed mode (--headed) and adjust selectors if needed."
        )

    return headers, all_rows


def _xlsx_col_name(index: int) -> str:
    name = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def _write_xlsx(path: str, headers: List[str], rows: List[List[str]]) -> None:
    all_rows = [headers] + rows
    width = len(headers)
    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]

    for r_idx, row in enumerate(all_rows, start=1):
        sheet_lines.append(f'<row r="{r_idx}">')
        padded = row + [""] * (width - len(row))
        for c_idx, value in enumerate(padded[:width], start=1):
            ref = f"{_xlsx_col_name(c_idx)}{r_idx}"
            txt = escape(value)
            sheet_lines.append(f'<c r="{ref}" t="inlineStr"><is><t>{txt}</t></is></c>')
        sheet_lines.append("</row>")

    sheet_lines.extend(["</sheetData>", "</worksheet>"])
    sheet_xml = "".join(sheet_lines)

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>"""

    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>"""

    workbook = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"dietitians\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>"""

    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>"""

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def save_output(path: str, headers: List[str], rows: List[List[str]]) -> None:
    if path.lower().endswith(".xlsx"):
        _write_xlsx(path, headers, rows)
        return

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in rows:
            padded = row + [""] * (len(headers) - len(row))
            writer.writerow(padded[: len(headers)])


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape all Dietitians directory results")
    parser.add_argument("--output", default="dietitians.xlsx", help="Output file (.xlsx or .csv)")
    parser.add_argument("--expected-count", type=int, default=602, help="Expected number of entries")
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Run browser in headed mode (helps debugging selectors)",
    )
    parser.add_argument("--max-pages", type=int, default=1000)
    args = parser.parse_args()

    try:
        headers, rows = scrape(headless=not args.headed, max_pages=args.max_pages)
        save_output(args.output, headers, rows)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Saved {len(rows)} unique results to: {args.output}")
    if args.expected_count and len(rows) != args.expected_count:
        print(
            f"WARNING: expected {args.expected_count} entries but got {len(rows)}.",
            file=sys.stderr,
        )
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
