"""
review_finder.py
----------------
Finds a review page URL for a hospital using a shared Playwright page.

Algorithm:
  Try HexaHealth → Practo → Justdial in order.
  For each site:
    1. Search DuckDuckGo with the real browser (avoids HTTP-202 bot blocks).
    2. If DDG returns nothing, fall back to the site's own search.
    3. Verify the page actually has reviews (lenient check).
    4. Return (site, url) on first confirmed hit.
  Return (None, None) if all three fail.

The caller owns the browser; this module never opens its own.
"""

import time
import random
from urllib.parse import quote_plus

from config import SEARCH_DELAY_MIN, SEARCH_DELAY_MAX
from logger import get_logger

log = get_logger("REVIEW_FINDER")

SEARCH_ORDER = [
    "hexahealth.com",
    "practo.com",
    "justdial.com",
]

# Keywords that reliably appear on review pages across all three sites.
# We check page text instead of fragile CSS selectors.
REVIEW_KEYWORDS = [
    "review", "rating", "stars", "rated", "experience",
    "patient", "recommended", "helpful",
]


def _delay():
    time.sleep(random.uniform(SEARCH_DELAY_MIN, SEARCH_DELAY_MAX))


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _duckduckgo_search(page, hospital_name: str, site: str) -> str | None:
    """
    Search DuckDuckGo in the real Playwright browser and return the
    first result URL that belongs to `site`.
    Real browser = no HTTP-202 CAPTCHA that killed the old requests approach.
    """
    query = quote_plus(f"{hospital_name} reviews site:{site}")
    ddg_url = f"https://duckduckgo.com/?q={query}&ia=web"

    log.info(f"  DDG: {ddg_url}")

    try:
        page.goto(ddg_url, wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(2_500)

        for anchor in page.locator("a[href]").all():
            href = (anchor.get_attribute("href") or "").strip()
            if (
                href.startswith("http")           # must be an absolute URL
                and site in href                   # must belong to the target site
                and "duckduckgo.com" not in href   # must not be a DDG internal link
            ):
                log.info(f"  DDG result: {href}")
                return href

    except Exception as exc:
        log.warning(f"  DDG error ({site}): {exc}")

    return None


def _site_direct_search(page, hospital_name: str, site: str) -> str | None:
    """
    Fallback: use the site's own internal search when DDG returns nothing.
    """
    try:
        if site == "hexahealth.com":
            page.goto(
                f"https://www.hexahealth.com/search?query={quote_plus(hospital_name)}",
                wait_until="domcontentloaded", timeout=30_000,
            )
            page.wait_for_timeout(2_000)
            link = page.locator("a[href*='/hospital']").first
            if link.count():
                href = link.get_attribute("href") or ""
                if not href.startswith("http"):
                    href = "https://www.hexahealth.com" + href
                return href

        elif site == "practo.com":
            page.goto(
                f"https://www.practo.com/search/hospitals?query={quote_plus(hospital_name)}",
                wait_until="domcontentloaded", timeout=30_000,
            )
            page.wait_for_timeout(2_000)
            link = page.locator("a[href*='/hospital']").first
            if link.count():
                href = link.get_attribute("href") or ""
                if not href.startswith("http"):
                    href = "https://www.practo.com" + href
                return href

        elif site == "justdial.com":
            page.goto(
                f"https://www.justdial.com/search?q={quote_plus(hospital_name + ' hospital')}",
                wait_until="domcontentloaded", timeout=30_000,
            )
            page.wait_for_timeout(2_000)
            link = page.locator("a[href*='justdial.com']").first
            if link.count():
                href = link.get_attribute("href") or ""
                if not href.startswith("http"):
                    href = "https://www.justdial.com" + href
                return href

    except Exception as exc:
        log.warning(f"  Direct search error ({site}): {exc}")

    return None


def _fix_url(url: str, site: str) -> str:
    """
    Apply site-specific URL corrections.

    HexaHealth review URLs must look like:
        /{city}/hospital/{slug}/reviews
        /{city}/doctor/{slug}/reviews

    DDG sometimes returns sub-pages like:
        /{city}/hospital/{slug}/doctors-list/endocrinologist

    We strip everything after the entity slug (the first path segment
    immediately after /hospital/ or /doctor/) so we always get a clean
    base URL, then append /reviews.  This avoids HexaHealth "Oops" pages
    caused by /reviews appended to a deeply nested path.
    """
    import re
    if site == "hexahealth.com":
        # Capture up to and including /{city}/{hospital|doctor}/{slug}
        match = re.match(
            r"(https?://[^/]+/[^/]+/(?:hospital|doctor)/[^/?#]+)",
            url,
        )
        if match:
            base = match.group(1).rstrip("/")
            url = base + "/reviews"
        elif "/reviews" not in url.lower():
            url = url.rstrip("/") + "/reviews"
    return url


def _has_reviews(page, url: str, site: str) -> bool:
    """
    Navigate to the URL and do a LENIENT check for review content.

    We check page text for common review keywords instead of relying
    on exact CSS selectors (which break whenever a site updates its HTML).
    Returns True if ANY review keyword is found in the visible text.
    """
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(2_000)

        # Check visible page text — lenient and site-agnostic
        body_text = page.locator("body").inner_text().lower()
        for keyword in REVIEW_KEYWORDS:
            if keyword in body_text:
                log.info(f"  Review keyword '{keyword}' found on page.")
                return True

        log.info(f"  No review keywords found on: {url}")
        return False

    except Exception as exc:
        log.warning(f"  Could not load {url}: {exc}")
        return False


# ─────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────

def find_review_url(page, hospital_name: str) -> tuple:
    """
    Find a verified review page for `hospital_name`.

    Args:
        page:          Shared Playwright Page (owned by the caller).
        hospital_name: Hospital name string.

    Returns:
        (site, url)   e.g. ("practo.com", "https://...")
        (None, None)  if no review page confirmed on any site.
    """
    for site in SEARCH_ORDER:
        log.info(f"Trying {site} for: {hospital_name}")

        try:
            url = _duckduckgo_search(page, hospital_name, site)

            if not url:
                log.info("  DDG empty — trying direct site search…")
                url = _site_direct_search(page, hospital_name, site)

            if not url:
                log.info(f"  No URL found on {site}.")
                _delay()
                continue

            url = _fix_url(url, site)
            log.info(f"  Candidate URL: {url}")

            if _has_reviews(page, url, site):
                log.info(f"  ✓ Reviews confirmed on {site}")
                return site, url
            else:
                log.info(f"  No reviews confirmed on {site}, trying next…")

        except Exception as exc:
            log.error(f"  Error on {site}: {exc}")

        _delay()

    log.info(f"No review page found for: {hospital_name}")
    return None, None
