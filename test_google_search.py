from playwright.sync_api import sync_playwright

HOSPITAL_NAME = "Prakash Hospital - Best Hospital in Noida"

with sync_playwright() as p:

    browser = p.chromium.launch(headless=False)

    page = browser.new_page()

    print("Opening DuckDuckGo...")

    page.goto("https://duckduckgo.com/")

    page.get_by_role("textbox").fill(
        HOSPITAL_NAME + " Google Maps"
    )

    page.keyboard.press("Enter")

    page.wait_for_load_state("networkidle")

    print("Searching complete.")

    links = page.locator("a")

    google_link = None

    for i in range(links.count()):

        href = links.nth(i).get_attribute("href")

        if href and "google.com/maps" in href:

            google_link = href

            break

    if not google_link:

        print("No Google Maps link found.")

        input("Press ENTER to close...")

        browser.close()

        raise SystemExit()

    print("\nGoogle Maps URL found:\n")
    print(google_link)

    page.goto(google_link)

    page.wait_for_load_state("domcontentloaded")

    input("\nBrowser paused. Check that the correct hospital opened.\nPress ENTER to close...")

    browser.close()