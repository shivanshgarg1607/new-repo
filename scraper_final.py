"""
scraper_final.py  v11.0
------------------------
Architecture (mirrors the original working scripts exactly):

  review_finder.py  — used DuckDuckGo INSIDE the real Playwright browser.
                      A real browser avoids the bot blocks that kill plain
                      HTTP requests to DDG/Bing/Google from cloud IPs.
  review_scraper.py — scraped pages using URLs already stored in Excel.
                      No searching needed when the URL is already known.

v11 combines both into one script with the same core logic:

  1. Read "Edit URL" from the Hospitals sheet.
     If a URL already exists → go straight to scraping (no DDG needed).
     If no URL → use DDG browser search (same as original review_finder.py).

  2. DDG browser search tries hexahealth.com → practo.com → justdial.com.
     Fallback: site's own search page.
     Newly found URLs are saved back to Excel so future runs are instant.

  3. Scrape review pages with the confirmed selectors:
       HexaHealth : div.reviewCard  /  p.review  /  span.text-capitalize
       Practo     : div[data-qa-id='review_card'] and several fallbacks
       JustDial   : div.reviewBox / div.rating-text / multiple fallbacks

  4. Phase 1 — non-success hospitals (need URL + reviews).
     Phase 2 — gap fill on already-successful hospitals (URL known, just scrape more).

Pre-flight (run once):
    pip install playwright openpyxl beautifulsoup4 requests lxml
    python -m playwright install chromium
"""

import re
import sys
import json
import time
import random
import hashlib
import logging
import traceback
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

try:
    import requests
    from bs4 import BeautifulSoup
    from openpyxl import load_workbook
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError as _e:
    sys.exit(
        f"\n[FATAL] Missing dependency: {_e}\n"
        "Run:\n"
        "  pip install playwright openpyxl beautifulsoup4 requests lxml\n"
        "  python -m playwright install chromium\n"
    )

# ─────────────────────────────────────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────────────────────────────────────

_ROOT           = Path(__file__).resolve().parent
EXCEL_FILE      = _ROOT / "output" / "HospitalAutomation.xlsx"
LOG_DIR         = _ROOT / "logs"
CACHE_DIR       = _ROOT / "cache"
LOG_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)
CHECKPOINT_FILE = CACHE_DIR / "scraper_final_checkpoint.json"

# ─────────────────────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────────────────────

TARGET_NEW_REVIEWS    = 1500
PHASE1_MAX_PER_SOURCE = 10
PHASE2_PER_HOSPITAL   = 4
SCROLL_ROUNDS         = 4
SEARCH_DELAY_MIN      = 2.0   # delay between DDG searches (be polite)
SEARCH_DELAY_MAX      = 4.0
HEADLESS              = True
BROWSER_TIMEOUT       = 30_000
PAGE_WAIT_MS          = 2_500
SCROLL_WAIT_MS        = 1_500

NON_SUCCESS_STATUSES = {
    "No Match", "No Reviews", "Failed", "Scraper Error",
    "No Review", "Not Found", "Error", "Pending", "Low Reviews", None, "",
}

SEARCH_ORDER = ["hexahealth.com", "practo.com", "justdial.com"]

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────

_log_file = LOG_DIR / f"scraper_final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("scraper_final")

# ─────────────────────────────────────────────────────────────────────────────
# Checkpoint
# ─────────────────────────────────────────────────────────────────────────────

def save_checkpoint(phase, index, name, total_new):
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump({"phase": phase, "index": index, "name": name,
                   "total_new": total_new, "saved_at": datetime.now().isoformat()}, f, indent=2)

def load_checkpoint():
    if CHECKPOINT_FILE.exists():
        with open(CHECKPOINT_FILE, encoding="utf-8") as f:
            return json.load(f)
    return None

def clear_checkpoint():
    if CHECKPOINT_FILE.exists():
        CHECKPOINT_FILE.unlink()

# ─────────────────────────────────────────────────────────────────────────────
# Deduplication
# ─────────────────────────────────────────────────────────────────────────────

def _hash(hospital, reviewer, review_text):
    raw = f"{hospital.strip().lower()}|{reviewer.strip().lower()}|{review_text.strip().lower()}"
    return hashlib.sha256(raw.encode()).hexdigest()

# ─────────────────────────────────────────────────────────────────────────────
# Excel helpers
# ─────────────────────────────────────────────────────────────────────────────

def load_existing_hashes():
    wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    ws = wb["Reviews"]
    hashes = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            h = _hash(str(row[0] or ""), str(row[1] or ""), str(row[3] or ""))
            hashes.add(h)
        except Exception:
            pass
    wb.close()
    log.info(f"Loaded {len(hashes):,} existing review hashes.")
    return hashes


def load_hospitals():
    """
    Load hospitals from the 'Hospitals' sheet.
    Reads all columns so we can find 'Edit URL' by header name.
    Returns list of dicts with: name, status, edit_url, row_number
    """
    if not EXCEL_FILE.exists():
        sys.exit(f"\n[FATAL] Excel file not found: {EXCEL_FILE}")
    wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    if "Hospitals" not in wb.sheetnames:
        sys.exit(f"\n[FATAL] Sheet 'Hospitals' not found. Sheets: {wb.sheetnames}")
    ws = wb["Hospitals"]

    # Read header row to find column indices
    headers = {}
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    for idx, h in enumerate(header_row):
        if h:
            headers[str(h).strip().lower()] = idx

    log.info(f"Hospitals sheet columns: {list(headers.keys())}")

    # Identify key columns — try multiple possible names
    name_col   = _find_col(headers, ["hospital name", "name", "hospital"])
    status_col = _find_col(headers, ["status", "scrape status", "review status"])
    url_col    = _find_col(headers, ["edit url", "editurl", "url", "review url",
                                      "hexahealth url", "source url", "link"])
    count_col  = _find_col(headers, ["review count", "reviews", "count", "total reviews"])

    log.info(f"  name_col={name_col}, status_col={status_col}, "
             f"url_col={url_col}, count_col={count_col}")

    out = []
    for rn, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        name = _cell(row, name_col)
        if not name:
            continue
        out.append({
            "name":       name,
            "status":     _cell(row, status_col),
            "edit_url":   _cell(row, url_col),
            "count_col":  count_col,
            "row_number": rn,
        })

    wb.close()
    log.info(f"Loaded {len(out):,} hospitals from Excel.")
    return out


def _find_col(headers: dict, candidates: list) -> int | None:
    for c in candidates:
        if c in headers:
            return headers[c]
    return None


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return str(row[idx] or "").strip()


def append_reviews(new_rows):
    if not new_rows:
        return
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Reviews"]
    for r in new_rows:
        ws.append([
            r.get("hospital_name", ""),
            r.get("reviewer",      "Unknown"),
            r.get("rating",        0),
            r.get("review",        ""),
            r.get("date",          ""),
            r.get("source",        ""),
            "",
        ])
    wb.save(EXCEL_FILE)
    wb.close()


def update_hospital_status(hospital_name: str, reviews_added: int, status: str):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Hospitals"]
        header_row = [c.value for c in ws[1]]
        status_idx = _header_idx(header_row, ["status", "scrape status", "review status"])
        count_idx  = _header_idx(header_row, ["review count", "reviews", "count", "total reviews"])
        for row in ws.iter_rows(min_row=2):
            name_cell = row[1] if len(row) > 1 else None
            if name_cell and str(name_cell.value or "").strip() == hospital_name:
                if status_idx is not None:
                    row[status_idx].value = status
                if count_idx is not None:
                    row[count_idx].value = reviews_added
                break
        wb.save(EXCEL_FILE)
        wb.close()
    except Exception as e:
        log.warning(f"Could not update status for '{hospital_name}': {e}")


def save_url_to_excel(hospital_name: str, url: str):
    """Save a newly-discovered URL back to the Hospitals sheet."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Hospitals"]
        header_row = [c.value for c in ws[1]]
        url_idx = _header_idx(header_row, ["edit url", "editurl", "url", "review url",
                                            "hexahealth url", "source url", "link"])
        if url_idx is None:
            wb.close()
            return
        for row in ws.iter_rows(min_row=2):
            name_cell = row[1] if len(row) > 1 else None
            if name_cell and str(name_cell.value or "").strip() == hospital_name:
                row[url_idx].value = url
                break
        wb.save(EXCEL_FILE)
        wb.close()
        log.info(f"  Saved URL to Excel for '{hospital_name}'")
    except Exception as e:
        log.warning(f"Could not save URL for '{hospital_name}': {e}")


def _header_idx(header_row, candidates):
    lower = [str(h or "").strip().lower() for h in header_row]
    for c in candidates:
        if c in lower:
            return lower.index(c)
    return None

# ─────────────────────────────────────────────────────────────────────────────
# URL discovery — DDG inside the real Playwright browser
# (same approach as the original review_finder.py that worked)
# ─────────────────────────────────────────────────────────────────────────────

def _duckduckgo_search(page, hospital_name: str, site: str) -> str | None:
    """
    Search DuckDuckGo using the real Playwright browser.
    This is what the original review_finder.py used — it avoids the
    bot-detection blocks that kill plain HTTP requests from cloud IPs.
    """
    query   = f"{hospital_name} reviews site:{site}"
    ddg_url = f"https://duckduckgo.com/?q={quote_plus(query)}&ia=web"
    log.info(f"  DDG: {ddg_url}")
    try:
        page.goto(ddg_url, wait_until="domcontentloaded", timeout=BROWSER_TIMEOUT)
        page.wait_for_timeout(2_500)
        for anchor in page.locator("a[href]").all():
            href = (anchor.get_attribute("href") or "").strip()
            if (href.startswith("http")
                    and site in href
                    and "duckduckgo.com" not in href):
                log.info(f"  DDG hit: {href}")
                return href
    except Exception as e:
        log.warning(f"  DDG error ({site}): {e}")
    return None


def _site_direct_search(page, hospital_name: str, site: str) -> str | None:
    """
    Fallback: use the site's own search page when DDG returns nothing.
    """
    try:
        if site == "hexahealth.com":
            page.goto(
                f"https://www.hexahealth.com/search?query={quote_plus(hospital_name)}",
                wait_until="domcontentloaded", timeout=BROWSER_TIMEOUT,
            )
            page.wait_for_timeout(2_500)
            # Correct pattern: /hospital/ (singular) — confirmed from real URL
            for sel in ["a[href*='/hospital/']", "a[href*='/hospitals/']"]:
                link = page.locator(sel).first
                if link.count():
                    href = link.get_attribute("href") or ""
                    if not href.startswith("http"):
                        href = "https://www.hexahealth.com" + href
                    return href

        elif site == "practo.com":
            page.goto(
                f"https://www.practo.com/search/hospitals?query={quote_plus(hospital_name)}",
                wait_until="domcontentloaded", timeout=BROWSER_TIMEOUT,
            )
            page.wait_for_timeout(2_500)
            # Practo uses /clinic/ and /hospital/ — check both
            for sel in ["a[href*='/clinic/']", "a[href*='/hospital/']"]:
                link = page.locator(sel).first
                if link.count():
                    href = link.get_attribute("href") or ""
                    if not href.startswith("http"):
                        href = "https://www.practo.com" + href
                    return href

        elif site == "justdial.com":
            page.goto(
                f"https://www.justdial.com/search?q={quote_plus(hospital_name + ' hospital')}",
                wait_until="domcontentloaded", timeout=BROWSER_TIMEOUT,
            )
            page.wait_for_timeout(2_500)
            link = page.locator("a[href*='justdial.com']").first
            if link.count():
                href = link.get_attribute("href") or ""
                if not href.startswith("http"):
                    href = "https://www.justdial.com" + href
                return href

    except Exception as e:
        log.warning(f"  Direct search error ({site}): {e}")
    return None


def _fix_url(url: str, site: str) -> str:
    """
    Clean the URL and ensure it points to the reviews page.

    HexaHealth review URLs must be:  /{city}/hospital/{slug}/reviews
    DDG sometimes returns deep sub-pages like:
        /{city}/hospital/{slug}/doctors-list/endocrinologist
    Strip everything after the slug, then append /reviews.
    """
    if site == "hexahealth.com":
        match = re.match(
            r"(https?://[^/]+/[^/]+/(?:hospital|doctor)/[^/?#]+)", url
        )
        if match:
            url = match.group(1).rstrip("/") + "/reviews"
        elif "/reviews" not in url.lower():
            url = url.rstrip("/") + "/reviews"

    elif site in ("practo.com", "justdial.com"):
        # Append /reviews if not already present and not already on a reviews URL
        clean = url.split("?")[0].rstrip("/")
        if not clean.endswith("/reviews"):
            url = clean + "/reviews"

    return url


def _has_reviews(page, url: str) -> bool:
    """
    Navigate to URL and do a lenient keyword check.
    Returns True if the page looks like it has review content.
    (Same approach as original review_finder._has_reviews)
    """
    keywords = ["review", "rating", "stars", "rated", "experience",
                "patient", "recommended", "helpful"]
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=BROWSER_TIMEOUT)
        page.wait_for_timeout(2_000)
        body = page.locator("body").inner_text().lower()
        for kw in keywords:
            if kw in body:
                log.info(f"  Keyword '{kw}' found — page has reviews.")
                return True
        log.info(f"  No review keywords on: {url}")
        return False
    except Exception as e:
        log.warning(f"  Could not verify {url}: {e}")
        return False


def _detect_site(url: str) -> str:
    for site in ["hexahealth.com", "practo.com", "justdial.com"]:
        if site in url:
            return site
    return "unknown"


def find_review_url(page, hospital_name: str) -> tuple:
    """
    Find a verified review page URL for a hospital.
    Tries DDG browser search → site direct search for each site in order.
    Returns (site, url) or (None, None).
    """
    for site in SEARCH_ORDER:
        log.info(f"  Searching {site} for: {hospital_name}")
        try:
            url = _duckduckgo_search(page, hospital_name, site)
            if not url:
                log.info("  DDG empty — trying direct site search…")
                url = _site_direct_search(page, hospital_name, site)
            if not url:
                log.info(f"  No URL found on {site}.")
                time.sleep(random.uniform(SEARCH_DELAY_MIN, SEARCH_DELAY_MAX))
                continue

            url = _fix_url(url, site)
            log.info(f"  Candidate URL: {url}")

            if _has_reviews(page, url):
                log.info(f"  ✓ Reviews confirmed on {site}")
                return site, url
            else:
                log.info(f"  No reviews confirmed on {site}, trying next…")

        except Exception as e:
            log.error(f"  Error on {site}: {e}")

        time.sleep(random.uniform(SEARCH_DELAY_MIN, SEARCH_DELAY_MAX))

    log.info(f"  No review page found for: {hospital_name}")
    return None, None

# ─────────────────────────────────────────────────────────────────────────────
# Browser helpers
# ─────────────────────────────────────────────────────────────────────────────

def _safe_navigate(page, url):
    for attempt in range(1, 3):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=BROWSER_TIMEOUT)
            return True
        except PWTimeout:
            log.warning(f"  Timeout (attempt {attempt}): {url}")
        except Exception as e:
            msg = str(e)
            if "interrupted by another navigation" in msg:
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=8_000)
                except Exception:
                    pass
                time.sleep(1)
                continue
            log.warning(f"  Nav error (attempt {attempt}): {msg[:120]}")
        if attempt < 2:
            time.sleep(2)
    return False


def _scroll_and_load(page):
    load_more_sels = [
        "button:has-text('Load More')",
        "button:has-text('Show More')",
        "button:has-text('See All Reviews')",
        "button:has-text('View More')",
        "[class*='loadMore']",
        "[class*='load-more']",
        "[class*='showMore']",
    ]
    for _ in range(SCROLL_ROUNDS):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(SCROLL_WAIT_MS)
            for sel in load_more_sels:
                try:
                    btn = page.locator(sel).first
                    if btn.count() and btn.is_visible():
                        btn.scroll_into_view_if_needed()
                        btn.click()
                        page.wait_for_timeout(SCROLL_WAIT_MS)
                        break
                except Exception:
                    pass
        except Exception:
            pass

# ─────────────────────────────────────────────────────────────────────────────
# Parsers — confirmed selectors from live HTML
# ─────────────────────────────────────────────────────────────────────────────

def _parse_hexahealth(html: str, max_n: int) -> list:
    """
    Confirmed live selectors (user-verified):
      div.reviewCard  — card
      p.review        — text
      span.text-capitalize — reviewer
      img[src*='staryellow'] — filled star
    """
    soup  = BeautifulSoup(html, "html.parser")
    cards = soup.select("div.reviewCard")
    out, seen = [], set()
    for card in cards:
        if len(out) >= max_n:
            break
        p    = card.select_one("p.review")
        text = p.get_text(strip=True) if p else ""
        if not text or len(text) < 10 or text in seen:
            continue
        seen.add(text)
        name_tag = card.select_one("span.text-capitalize")
        reviewer = name_tag.get_text(strip=True) if name_tag else "Unknown"
        rating   = len(card.select("img[src*='staryellow']"))
        date = ""
        spans = card.select("span")
        if len(spans) >= 2:
            date = spans[-1].get_text(strip=True)
        out.append({"reviewer": reviewer, "rating": rating, "review": text, "date": date})
    return out


def _parse_practo(html: str, max_n: int) -> list:
    soup  = BeautifulSoup(html, "html.parser")
    cards = (
        soup.select("div[data-qa-id='review_card']")
        or soup.select("div[class*='doctor-review']")
        or soup.select("div[class*='review-card']")
        or soup.select("div[class*='ReviewCard']")
        or soup.select("div[class*='review_card']")
        or soup.select("div[class*='reviewCard']")
    )
    out, seen = [], set()
    for card in cards:
        if len(out) >= max_n:
            break
        name_tag = (
            card.find(attrs={"data-qa-id": "reviewer_name"})
            or card.find("p", class_=re.compile(r"reviewer|author|username", re.I))
            or card.find("span", class_=re.compile(r"reviewer|author|username", re.I))
            or card.find("strong")
        )
        reviewer = name_tag.get_text(strip=True) if name_tag else "Unknown"
        stars  = card.select("i.c-icon--star-filled, i[class*='star'][class*='fill'], "
                              "span[class*='star-fill']")
        rating = len(stars)
        text_tag = (
            card.find(attrs={"data-qa-id": "review_text"})
            or card.find("p", class_=re.compile(r"review.?text|comment|body", re.I))
            or card.find("div", class_=re.compile(r"review.?text|comment|body", re.I))
            or card.find("p")
        )
        text = text_tag.get_text(strip=True) if text_tag else ""
        if not text or len(text) < 10 or text in seen:
            continue
        seen.add(text)
        date_tag = card.find(class_=re.compile(r"date|time|when|ago", re.I))
        date     = date_tag.get_text(strip=True) if date_tag else ""
        out.append({"reviewer": reviewer, "rating": rating, "review": text, "date": date})
    return out


def _parse_justdial(html: str, max_n: int) -> list:
    soup  = BeautifulSoup(html, "html.parser")
    cards = (
        soup.select("div.reviewBox")
        or soup.select("div[class*='review-box']")
        or soup.select("div[class*='ReviewBox']")
        or soup.select("li[class*='review']")
        or soup.select("div[class*='user-review']")
    )
    out, seen = [], set()
    for card in cards:
        if len(out) >= max_n:
            break
        name_tag = (
            card.find(class_=re.compile(r"reviewer|username|user.?name|ratinguser", re.I))
            or card.find("strong")
            or card.find("b")
        )
        reviewer = name_tag.get_text(strip=True) if name_tag else "Unknown"
        rating_tag = card.find(class_=re.compile(r"rating|stars|score", re.I))
        rating = 0
        if rating_tag:
            m = re.search(r"(\d+\.?\d*)", rating_tag.get_text())
            if m:
                rating = float(m.group(1))
        text_tag = (
            card.find(class_=re.compile(r"review.?text|comment|description|rating.?text", re.I))
            or card.find("p")
        )
        text = text_tag.get_text(strip=True) if text_tag else ""
        if not text or len(text) < 10 or text in seen:
            continue
        seen.add(text)
        date_tag = card.find(class_=re.compile(r"date|time|when|ago", re.I))
        date     = date_tag.get_text(strip=True) if date_tag else ""
        out.append({"reviewer": reviewer, "rating": rating, "review": text, "date": date})
    return out


_PARSERS = {
    "hexahealth.com": _parse_hexahealth,
    "practo.com":     _parse_practo,
    "justdial.com":   _parse_justdial,
}


def _scrape_known_url(page, site: str, url: str, max_n: int) -> list:
    """Navigate directly to a known URL and parse reviews. No searching needed."""
    if not _safe_navigate(page, url):
        return []
    page.wait_for_timeout(PAGE_WAIT_MS)
    _scroll_and_load(page)
    try:
        html = page.content()
    except Exception:
        return []
    parser = _PARSERS.get(site, _PARSERS["hexahealth.com"])
    results = parser(html, max_n)
    log.info(f"  Parsed {len(results)} reviews from {url}")
    return results

# ─────────────────────────────────────────────────────────────────────────────
# Process one hospital
# ─────────────────────────────────────────────────────────────────────────────

def process_hospital(page, hospital: dict, existing_hashes: set,
                     max_reviews: int) -> list:
    name      = hospital["name"]
    edit_url  = hospital.get("edit_url", "").strip()
    new_rows  = []

    if edit_url:
        # ── Fast path: URL already stored in Excel ──────────────────────────
        site = _detect_site(edit_url)
        url  = _fix_url(edit_url, site)
        log.info(f"  Using stored URL ({site}): {url}")
        raw = _scrape_known_url(page, site, url, max_reviews)
        print(f"    [{site}] scraped={len(raw)}", end="")

    else:
        # ── Slow path: find URL via DDG browser search ───────────────────────
        log.info(f"  No stored URL — searching via DDG browser…")
        site, url = find_review_url(page, name)
        if not url:
            log.info(f"  No URL found for '{name}'")
            print("  — No URL found on any site.")
            return []
        # Save it to Excel so future runs skip the search
        save_url_to_excel(name, url)
        # Page is already at the URL from _has_reviews check — just parse
        _scroll_and_load(page)
        try:
            html = page.content()
        except Exception:
            html = ""
        parser = _PARSERS.get(site, _PARSERS["hexahealth.com"])
        raw = parser(html, max_reviews)
        print(f"    [{site}] scraped={len(raw)}", end="")

    added = 0
    for r in raw:
        h = _hash(name, r.get("reviewer", ""), r.get("review", ""))
        if h in existing_hashes or not r.get("review", "").strip():
            continue
        existing_hashes.add(h)
        new_rows.append({
            "hospital_name": name,
            "reviewer":      r.get("reviewer", "Unknown"),
            "rating":        r.get("rating",   0),
            "review":        r.get("review",   ""),
            "date":          r.get("date",     ""),
            "source":        site if edit_url else site,
        })
        added += 1

    print(f", new={added}")
    return new_rows

# ─────────────────────────────────────────────────────────────────────────────
# Phase 1 — Non-success hospitals
# ─────────────────────────────────────────────────────────────────────────────

def run_phase1(page, hospitals, hashes, checkpoint):
    candidates = [h for h in hospitals if h["status"] in NON_SUCCESS_STATUSES]
    total, total_new, start = len(candidates), 0, 0

    if checkpoint and checkpoint.get("phase") == "phase1":
        start     = checkpoint.get("index", 0)
        total_new = checkpoint.get("total_new", 0)
        print(f"\n[RESUME] Phase 1 from #{start} | Collected: {total_new:,}")

    print(f"\n{'='*70}")
    print(f"PHASE 1 — {total:,} non-success hospitals  |  Target: {TARGET_NEW_REVIEWS:,}")
    print(f"{'='*70}")

    for i, hosp in enumerate(candidates[start:], start=start):
        if total_new >= TARGET_NEW_REVIEWS:
            print(f"\n[TARGET REACHED] {total_new:,} reviews — stopping Phase 1.")
            break

        name = hosp["name"]
        print(f"\n[PROGRESS] Phase 1: {i+1}/{total} — {name}")
        log.info(f"Phase 1 [{i+1}/{total}]: {name}")

        try:
            rows = process_hospital(page, hosp, hashes, PHASE1_MAX_PER_SOURCE)
        except Exception as e:
            log.error(f"  Unhandled error for '{name}': {e}")
            log.debug(traceback.format_exc())
            rows = []

        if rows:
            append_reviews(rows)
            total_new += len(rows)
            update_hospital_status(name, len(rows), "Success")
            print(f"  ✓ +{len(rows)} reviews | Total: {total_new:,} / {TARGET_NEW_REVIEWS:,}")
        else:
            update_hospital_status(name, 0, "No Reviews")
            print("  — No reviews found.")

        save_checkpoint("phase1", i + 1, name, total_new)

    print(f"\nPhase 1 done. New reviews this run: {total_new:,}")
    return total_new

# ─────────────────────────────────────────────────────────────────────────────
# Phase 2 — Gap fill on already-successful hospitals (URLs stored, just scrape)
# ─────────────────────────────────────────────────────────────────────────────

def run_phase2(page, hospitals, hashes, checkpoint, so_far):
    candidates = [h for h in hospitals if h["status"] not in NON_SUCCESS_STATUSES]
    total, added, start = len(candidates), 0, 0

    if checkpoint and checkpoint.get("phase") == "phase2":
        start = checkpoint.get("index", 0)
        added = max(0, checkpoint.get("total_new", 0) - so_far)
        print(f"\n[RESUME] Phase 2 from #{start}")

    still_need = TARGET_NEW_REVIEWS - so_far
    print(f"\n{'='*70}")
    print(f"PHASE 2 — {total:,} success hospitals | Still need: {still_need:,}")
    print(f"{'='*70}")

    for i, hosp in enumerate(candidates[start:], start=start):
        if so_far + added >= TARGET_NEW_REVIEWS:
            print(f"\n[TARGET REACHED] {so_far + added:,} total.")
            break

        name = hosp["name"]
        print(f"\n[PROGRESS] Phase 2: {i+1}/{total} — {name}")
        log.info(f"Phase 2 [{i+1}/{total}]: {name}")

        try:
            rows = process_hospital(page, hosp, hashes, PHASE2_PER_HOSPITAL)
        except Exception as e:
            log.error(f"  Unhandled error for '{name}': {e}")
            rows = []

        if rows:
            append_reviews(rows)
            added += len(rows)
            print(f"  ✓ +{len(rows)} reviews | Total: {so_far + added:,} / {TARGET_NEW_REVIEWS:,}")
        else:
            print("  — No new reviews.")

        save_checkpoint("phase2", i + 1, name, so_far + added)

    print(f"\nPhase 2 done. Additional reviews: {added:,}")
    return added

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  HOSPITAL REVIEW AUTOMATION — scraper_final.py v11.0")
    print(f"  Excel : {EXCEL_FILE}")
    print(f"  Target: {TARGET_NEW_REVIEWS:,} new unique reviews")
    print(f"  Start : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    checkpoint = load_checkpoint()
    if checkpoint:
        print(f"\n[CHECKPOINT] phase='{checkpoint['phase']}' | "
              f"index={checkpoint['index']} | collected={checkpoint['total_new']:,}")
    else:
        print("\n[FRESH START] No checkpoint found.")

    hospitals = load_hospitals()
    hashes    = load_existing_hashes()
    total_new = 0

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=HEADLESS,
            args=["--disable-gpu", "--disable-dev-shm-usage",
                  "--no-sandbox", "--disable-extensions"],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/138.0.0.0 Safari/537.36"
            ),
            locale="en-US",
            viewport={"width": 1280, "height": 800},
        )
        # Block images/fonts/media to speed up page loads
        context.route(
            "**/*.{png,jpg,jpeg,gif,webp,svg,ico,woff,woff2,ttf,eot,mp4,mp3}",
            lambda route: route.abort(),
        )
        page = context.new_page()

        # Run phases
        skip_phase1 = checkpoint and checkpoint.get("phase") == "phase2"
        if skip_phase1:
            total_new = checkpoint.get("total_new", 0)
            print(f"\n[SKIP] Phase 1 already done ({total_new:,} reviews) — Phase 2…")
        else:
            total_new = run_phase1(page, hospitals, hashes, checkpoint)

        if total_new < TARGET_NEW_REVIEWS:
            total_new += run_phase2(page, hospitals, hashes,
                                     None if skip_phase1 else checkpoint,
                                     total_new)
        else:
            print(f"\n[TARGET MET] {total_new:,} reviews — Phase 2 not needed.")

        browser.close()

    clear_checkpoint()
    print("\n" + "=" * 70)
    print("  SCRAPING COMPLETE")
    print(f"  Total NEW unique reviews : {total_new:,}")
    print(f"  Saved to                 : {EXCEL_FILE}")
    print(f"  Log                      : {_log_file}")
    print(f"  Finished                 : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)


if __name__ == "__main__":
    main()
