from admin import AdminBot
from missing_logger import log_missing_hospital


def main():

    bot = AdminBot()

    skipped_hospitals = set()

    try:

        bot.launch_browser()

        bot.login()

        start_page = int(
            input("\nEnter starting page (1-143): ")
        )


        bot.open_hospital_page(start_page)

        while True:

            hospital = bot.find_first_placeholder_hospital(
                skipped_hospitals
            )

            # ----------------------------------------
            # No placeholders on current page
            # ----------------------------------------

            if hospital is None:

                print("\nNo placeholder hospitals on this page.")

                skipped_hospitals.clear()

                if bot.go_to_next_page():

                    continue

                print("\n==============================")
                print("ALL PAGES COMPLETED")
                print("==============================")

                break

            print("\n==============================")
            print("Hospital Found")
            print("==============================")

            print(f"Name : {hospital['name']}")

            bot.open_edit_page(hospital)

            bot.open_choose_image_modal()

            success = bot.upload_image(hospital)

            if success:

                print("\nImage uploaded.")

                bot.submit_hospital()

            else:

                print("\nNo suitable image found.")

                log_missing_hospital(
                    hospital["name"],
                    "No suitable image found"
                )

                skipped_hospitals.add(
                    hospital["name"]
                )

                bot.close_choose_image_modal()

                bot.submit_hospital()

        print("\nAutomation Complete.")

    finally:

        input("\nPress ENTER to close...")

        bot.close()


if __name__ == "__main__":
    main()