"""
extract_hospitals.py
--------------------
Phase 1:
Extract all hospitals from the admin panel and save them
to HospitalAutomation.xlsx.
"""

from admin import AdminBot
from excel_writer import ExcelWriter
from logger import get_logger
from checkpoint import (
    load_checkpoint,
    save_checkpoint,
    clear_checkpoint,
)

log = get_logger("EXTRACT")


def main():

    bot = AdminBot()
    writer = ExcelWriter()

    try:

        bot.launch_browser(headless=False)

        bot.login()

        checkpoint = load_checkpoint()

        if checkpoint:

            offset = checkpoint["hospital_index"]

            print(f"\nResuming from offset {offset}")

        else:

            offset = 0

        while True:

            print("\n" + "=" * 60)
            print(f"Opening Offset : {offset}")
            print("=" * 60)

            bot.open_hospital_page(offset)

            hospitals = bot.get_all_hospitals()

            if len(hospitals) == 0:

                print("\nNo hospitals found.")
                print("Extraction Complete.")

                break

            print(f"Found {len(hospitals)} hospitals")

            for hospital in hospitals:

                saved = writer.add_hospital(hospital)

                if saved:
                    print(f"✓ {hospital['name']}")
                else:
                    print(f"Skipped (already exists): {hospital['name']}")

            save_checkpoint(
                phase="extract",
                hospital_index=offset + 10,
                hospital_name=hospitals[-1]["name"],
            )
            offset += 10
            
    finally:
        clear_checkpoint() 
        bot.close()


if __name__ == "__main__":
    main()