import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://karunahealthlifepartner.com/index.php/account/admin")
    page.get_by_role("textbox", name="Enter User Name").click()
    page.get_by_role("textbox", name="Enter User Name").fill("karuna__admin")
    page.get_by_role("textbox", name="Enter User Name").press("Enter")
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill("karuna__admin")
    page.get_by_role("button", name="SIGN IN").click()
    page.get_by_role("link", name=" Reviews & Ratings 1 pending").click()
    page.get_by_role("button", name=" Add Review").click()
    page.get_by_role("textbox", name="e.g. Rahul S., Priya G.,").click()
    page.get_by_role("textbox", name="e.g. Rahul S., Priya G.,").fill("rahul")
    page.get_by_role("textbox", name="Type to search listings (e.g").click()
    page.get_by_role("textbox", name="Type to search listings (e.g").fill("vinayak")
    page.get_by_text("Vinayak Hospital — Hospital", exact=True).click()
    page.locator("#amrStarRow > i:nth-child(5)").click()
    page.get_by_role("textbox", name="Write the review content…").click()
    page.get_by_role("textbox", name="Write the review content…").fill("the doctors are amazing ")
    page.get_by_role("radio", name="Pending (need to approve").check()
    page.get_by_role("button", name=" Save Review").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
