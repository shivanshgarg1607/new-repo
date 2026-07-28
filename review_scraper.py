"""
review_scraper.py
-----------------
Scrapes reviews from a HexaHealth review page.

Accepts an optional `page` (Playwright Page).  When provided the HTML is
taken directly from the already-rendered browser page — no HTTP request is
made.  This is essential because HexaHealth is a React SPA: the /reviews
path returns 404 when fetched with plain requests, but renders fine in a
real browser.
"""

import requests
from bs4 import BeautifulSoup


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/138.0 Safari/537.36"
    )
}


def scrape_reviews(url: str, max_reviews: int = 5, page=None):
    """
    Scrape reviews from a HexaHealth review page.

    Args:
        url:         Full URL of the /reviews page.
        max_reviews: Maximum number of reviews to return.
        page:        Optional Playwright Page object.  When supplied the
                     already-rendered HTML is used instead of a fresh HTTP
                     request (required for SPAs like HexaHealth).

    Returns:
        list[dict]  — each dict has: reviewer, rating, review, date
    """
    if page is not None:
        # The Playwright page is already at this URL (navigated by
        # review_finder._has_reviews).  Grab the live-rendered HTML.
        html = page.content()
    else:
        # Fallback: plain HTTP request (works only for server-rendered pages).
        response = requests.get(url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        html = response.text

    soup = BeautifulSoup(html, "html.parser")

    review_cards = soup.select("div.reviewCard")

    reviews = []

    for card in review_cards:

        review_text = ""
        review = card.select_one("p.review")
        if review:
            review_text = review.get_text(strip=True)

        reviewer = "Unknown"
        name = card.select_one("span.text-capitalize")
        if name:
            reviewer = name.get_text(strip=True)

        stars = len(card.select("img[src*='staryellow']"))

        date = ""
        spans = card.select("span")
        if len(spans) >= 2:
            date = spans[-1].get_text(strip=True)

        reviews.append(
            {
                "reviewer": reviewer,
                "rating":   stars,
                "review":   review_text,
                "date":     date,
            }
        )

        if len(reviews) >= max_reviews:
            break

    return reviews


if __name__ == "__main__":
    url = input("URL: ")
    reviews = scrape_reviews(url)
    print()
    print("Found", len(reviews), "reviews")
    for r in reviews:
        print(r)
