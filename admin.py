"""
admin.py
---------
Admin website controller for the Hospital Review Automation Bot.

Responsibilities
----------------
- Browser lifecycle
- Login
- Navigation
- Shared Playwright helpers

Business logic (extracting hospitals, collecting reviews, uploading reviews)
belongs in separate modules.
"""

from pathlib import Path
from typing import Optional

from playwright.sync_api import (
    sync_playwright,
    BrowserContext,
    Page,
    TimeoutError as PlaywrightTimeoutError,
)

from config import (
    BASE_URL,
    USERNAME,
    PASSWORD,
    HEADLESS,
    DEFAULT_TIMEOUT,
    SCREENSHOT_DIR,
)

from logger import get_logger

log = get_logger("ADMIN")


class AdminBot:

    def __init__(self):

        self.playwright = None
        self.browser = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

        self.profile_path = Path("cache/browser_profile")

    # ==========================================================
    # Browser
    # ==========================================================

    def start(self):

        log.info("Starting browser...")

        self.playwright = sync_playwright().start()

        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(self.profile_path),
            headless=HEADLESS,
        )

        self.page = self.context.new_page()

        self.page.set_default_timeout(DEFAULT_TIMEOUT)

        log.info("Browser started.")

    def stop(self):

        log.info("Closing browser...")

        try:

            if self.context:
                self.context.close()

        finally:

            if self.playwright:
                self.playwright.stop()

        log.info("Browser closed.")

    # ==========================================================
    # Helpers
    # ==========================================================

    def screenshot(self, name: str):

        path = SCREENSHOT_DIR / f"{name}.png"

        self.page.screenshot(path=str(path))

        log.info(f"Screenshot saved -> {path}")

    def wait(self, milliseconds: int = 1000):

        self.page.wait_for_timeout(milliseconds)

    # ==========================================================
    # Login
    # ==========================================================

    def login(self):

        log.info("Opening login page...")

        self.page.goto(BASE_URL)

        self.page.get_by_role(
            "textbox",
            name="Enter User Name"
        ).fill(USERNAME)

        self.page.get_by_role(
            "textbox",
            name="Password"
        ).fill(PASSWORD)

        self.page.get_by_role(
            "button",
            name="SIGN IN"
        ).click()

        self.wait(2000)

        log.info("Login successful.")

    # ==========================================================
    # Navigation
    # ==========================================================

    def goto_hospital_section(self):

        log.info("Opening Hospital section...")

        self.page.get_by_role(
            "link",
            name="  Services Categories "
        ).click()

        self.wait(1000)

        self.page.get_by_role(
            "link",
            name=" Hospital"
        ).click()

        self.wait(2000)

        log.info("Hospital list opened.")

    def goto_reviews_section(self):

        log.info("Opening Reviews section...")

        self.page.get_by_role(
            "link",
            name=" Reviews & Ratings"
        ).click()

        self.wait(2000)

        log.info("Review page opened.")

    # ==========================================================
    # Runtime Pause
    # ==========================================================

    def pause(self, reason: str):

        print("\n" + "=" * 70)
        print("AUTOMATION PAUSED")
        print("=" * 70)

        print(reason)

        while True:

            command = input(
                "\nType resume / retry / skip / quit : "
            ).strip().lower()

            if command in [
                "resume",
                "retry",
                "skip",
                "quit",
            ]:
                return command

            print("Invalid command.")

    # ==========================================================
    # Safe Action
    # ==========================================================

    def safe_click(self, locator):

        try:

            locator.click()

            return True

        except Exception as e:

            log.exception(e)

            self.screenshot("click_error")

            command = self.pause(str(e))

            return command