"""
excel_writer.py
----------------
Handles all Excel writing for the project.
"""

from pathlib import Path
from openpyxl import Workbook, load_workbook

from config import HOSPITAL_LIST_FILE
from logger import get_logger

log = get_logger("EXCEL")


class HospitalExcelWriter:

    HEADERS = [
        "Website ID",
        "Hospital Name",
        "Phone",
        "Email",
        "Status"
    ]

    def __init__(self):

        self.file = HOSPITAL_LIST_FILE

        if not self.file.exists():

            wb = Workbook()

            ws = wb.active

            ws.title = "Hospitals"

            ws.append(self.HEADERS)

            wb.save(self.file)

            log.info(f"Created {self.file}")

    def append_hospital(
        self,
        website_id,
        hospital_name,
        phone,
        email,
        status
    ):

        wb = load_workbook(self.file)

        ws = wb.active

        ws.append([
            website_id,
            hospital_name,
            phone,
            email,
            status
        ])

        wb.save(self.file)

        log.info(f"Added {hospital_name}")