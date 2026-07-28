from playwright.sync_api import sync_playwright
from config import BASE_URL, USERNAME, PASSWORD


class AdminBot:

    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        

    def launch_browser(self, headless=False):
        self.playwright = sync_playwright().start()

        self.browser = self.playwright.chromium.launch(
            headless=headless
        )

        self.context = self.browser.new_context()
        self.page = self.context.new_page()

    def login(self):
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

        self.page.wait_for_load_state("networkidle")

    def open_hospital_page(self, offset=0):

        url = (
            "https://karunahealthlifepartner.com/"
            f"Servicemodules/categoryitemSetting/1/{offset}"
        )

        print(f"\nOpening offset {offset}...")

        for attempt in range(3):
            try:
                self.page.goto(
                    url,
                    wait_until="domcontentloaded",
                    timeout=60000
                )
                return
            except Exception as e:
                print("Waiting 15 seconds before retry...")
                self.page.wait_for_timeout(15000)
                print(f"Retry {attempt+1}/3")
                print(e)

        print("Navigation failed.")
        while True:
            command = input(
                "\nType resume after fixing the problem (or quit): "
            ).strip().lower()

            if command == "resume":
                break

            if command == "quit":
                raise SystemExit("Stopped by user.")

        try:
            self.page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=60000
            )
            return
        except Exception:
            pass

        self.page.wait_for_load_state("networkidle")

    def get_hospital_rows(self):
        return self.page.locator("table tbody tr")

    def get_hospital_data(self, row):

        cells = row.locator("td")

        image = cells.nth(1).locator("img")
        image_src = image.get_attribute("src") or ""

        hospital_name = (
            cells.nth(2)
            .text_content()
            .replace("updated", "")
            .strip()
        )

        edit_url = (
            row.locator("a", has_text="Edit")
            .get_attribute("href")
        )

        return {
            "name": hospital_name,
            "image_src": image_src,
            "edit_url": edit_url,
            "row": row
        }

    
    def go_to_next_page(self):

        print("\nChecking for next page...")

        # Current active page number
        active = self.page.locator("li.page-item.active a.page-link")

        if active.count() == 0:

            print("Could not determine current page.")

            return False

        current_page = int(active.inner_text().strip())

        print(f"Current page: {current_page}")

        # Get ALL numbered page links
        links = self.page.locator("li.page-item a.page-link")

        total = links.count()

        best_link = None
        best_number = None

        for i in range(total):

            text = links.nth(i).inner_text().strip()

            if text.isdigit():

                number = int(text)

                if number > current_page:

                    if best_number is None or number < best_number:

                        best_number = number
                        best_link = links.nth(i)

        if best_link:

            print(f"Opening page {best_number}...")

            best_link.click()

            self.page.wait_for_load_state("networkidle")

            return True

        print("No next page.")

        return False
    
    # def go_to_next_page(self):

    #     print("\nChecking for next page...")

    #     next_button = self.page.locator(
    #         "a.page-link",
    #         has_text=">"
    #     )

    #     if next_button.count() == 0:

    #         print("No next page.")

    #         return False

    #     if not next_button.first.is_visible():

    #         print("Next button not visible.")

    #         return False

    #     print("Opening next page...")

    #     next_button.first.click()

    #     self.page.wait_for_load_state("networkidle")

    #     return True

    def open_edit_page(self, hospital):

        print(f"\nOpening edit page for: {hospital['name']}")

        self.page.goto(hospital["edit_url"])

        self.page.wait_for_load_state("networkidle")

        print("✓ Edit page opened successfully.")

    

    
    
    
    
    
    
    
    
    
    def close(self):

        if self.browser:
            self.browser.close()

        if self.playwright:
            self.playwright.stop()
    def get_hospital_data(self, row):

        website_id = (
            row.locator("input.row-check")
            .get_attribute("value")
        )

        hospital_name = (
            row.locator("td").nth(2)
            .inner_text()
            .replace("updated", "")
            .strip()
        )

        phone_email = (
            row.locator("td").nth(5)
            .inner_text()
            .split("\n")
        )

        phone = phone_email[0].strip() if len(phone_email) > 0 else ""

        email = phone_email[1].strip() if len(phone_email) > 1 else ""

        status = (
            row.locator("td").nth(6)
            .inner_text()
            .strip()
        )

        edit_url = row.locator(
            "a",
            has_text="Edit"
        ).get_attribute("href")

        return {
            "id": website_id,
            "name": hospital_name,
            "phone": phone,
            "email": email,
            "status": status,
            "edit_url": edit_url,
        }


    def get_all_hospitals(self):

        rows = self.get_hospital_rows()

        hospitals = []

        for i in range(rows.count()):

            hospitals.append(
                self.get_hospital_data(
                    rows.nth(i)
                )
            )

        return hospitals            