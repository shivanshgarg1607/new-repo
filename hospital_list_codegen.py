import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://karunahealthlifepartner.com/index.php/account/admin")
    page.get_by_role("textbox", name="Enter User Name").click()
    page.get_by_role("textbox", name="Enter User Name").fill("karuna__admin")
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill("karuna__admin")
    page.get_by_role("button", name="SIGN IN").click()
    page.get_by_role("link", name="  Services Categories ").click()
    page.get_by_role("link", name=" Hospital").click()
    page.get_by_role("link", name="2", exact=True).click()
    page.get_by_role("link", name="3").click()
    page.get_by_role("link", name="4").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright) 