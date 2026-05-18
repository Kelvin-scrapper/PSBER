"""
PSBER_MFRP Scraper — Poland State Budget Execution Report
Source: https://www.gov.pl/web/finance/estimated-execution-of-the-state-budget
Fetches PDF reports (one per monthly period) and saves to downloads/.
Usage: fetch_data(downloads_dir, latest_only=True)
       Returns list of (period_hint, local_pdf_path) tuples.
"""

import os
import re
import requests
from bs4 import BeautifulSoup

PAGE_URL = "https://www.gov.pl/web/finance/estimated-execution-of-the-state-budget"
BASE_URL = "https://www.gov.pl"

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

_MONTH_MAP = {
    "january": "01", "february": "02", "march": "03", "april": "04",
    "may": "05", "june": "06", "july": "07", "august": "08",
    "september": "09", "october": "10", "november": "11", "december": "12",
}


def _parse_period_hint(link_text: str) -> str | None:
    """Extract period like '2026-04' from link text 'Estimated execution January - April 2026...'"""
    months = re.findall(
        r"(January|February|March|April|May|June|July|August|September|October|November|December)",
        link_text, re.IGNORECASE
    )
    year_m = re.search(r"(\d{4})", link_text)
    if months and year_m:
        return f"{year_m.group(1)}-{_MONTH_MAP[months[-1].lower()]}"
    return None


def _discover_links() -> list[tuple[str, str]]:
    """
    Returns list of (period_hint, attachment_url) sorted newest-first.
    period_hint may be None if link text could not be parsed.
    """
    r = requests.get(PAGE_URL, headers=_HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    results = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not href.startswith("/attachment/"):
            continue
        text = a.get_text(" ", strip=True)
        if "estimated execution" not in text.lower():
            continue
        period = _parse_period_hint(text)
        full_url = BASE_URL + href
        results.append((period, full_url))

    return results


def fetch_data(downloads_dir: str = "downloads", latest_only: bool = True) -> list[tuple[str, str]]:
    """
    Download PDF report(s) and save to downloads_dir.
    Returns list of (period_hint, local_path).
    """
    os.makedirs(downloads_dir, exist_ok=True)

    links = _discover_links()
    if not links:
        raise RuntimeError("[scraper] No attachment links found on page")

    if latest_only:
        links = links[:1]

    downloaded = []
    for period, url in links:
        safe_period = period.replace("-", "_") if period else "unknown"
        filename = f"psber_{safe_period}.pdf"
        local_path = os.path.join(downloads_dir, filename)

        if os.path.exists(local_path):
            print(f"[scraper] Already downloaded: {filename}")
            downloaded.append((period, local_path))
            continue

        print(f"[scraper] Downloading {period or '?'}: {url}")
        r = requests.get(url, headers=_HEADERS, timeout=60)
        r.raise_for_status()
        with open(local_path, "wb") as f:
            f.write(r.content)
        print(f"[scraper] Saved: {filename} ({len(r.content):,} bytes)")
        downloaded.append((period, local_path))

    return downloaded
