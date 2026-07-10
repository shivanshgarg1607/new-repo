from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto("https://karunahealthlifepartner.com/index.php/account/admin")

    input("Login manually, navigate to the Hospital List, then press ENTER here...")

    rows = page.locator("table tbody tr")

    print(f"Found {rows.count()} rows\n")

    first_row = rows.first

    print(first_row.inner_html())

    input("\nPress ENTER to close...")
    browser.close()