from excel_writer import HospitalExcelWriter

writer = HospitalExcelWriter()

writer.append_hospital(
    1,
    "Apollo Hospital",
    "9876543210",
    "apollo@test.com",
    "ACTIVE"
)

writer.append_hospital(
    2,
    "Fortis Hospital",
    "9999999999",
    "fortis@test.com",
    "ACTIVE"
)

print("Done")