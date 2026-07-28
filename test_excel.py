from excel_writer import ExcelWriter

writer = ExcelWriter()

writer.add_hospital({
    "id": "1897",
    "name": "Prakash Hospital",
    "phone": "8826000033",
    "email": "test@gmail.com",
    "status": "Active",
    "edit_url": "https://example.com/edit/1897"
})

print("Done")