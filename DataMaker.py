import docx

docx_file_path = 'test\\test.docx'
doc = docx.Document(docx_file_path)


# =========================== INSTITUTION DETAILS =========================== #

# Variables
inside_institution_details = False
name_of_institution = ""
place = ""
phone_number = ""
email_id = ""

for paragraph in doc.paragraphs:
    text = paragraph.text
    # Check if paragraph contains the "Institution Details"
    if text.startswith("Institution Details"):
        inside_institution_details = True
    elif inside_institution_details:
        # Split the paragraph by the colon
        parts = text.split(':')
        if len(parts) == 2:
            key = parts[0].strip()
            value = parts[1].strip()

            # Assign values to variables
            if key == "Name of the Institution":
                name_of_institution = value
            elif key == "Place":
                place = value
            elif key == "Phone number":
                phone_number = value
            elif key == "Email Id":
                email_id = value

# Extracted data
print(name_of_institution)
print(place)
print(phone_number)
print(email_id)


# ============================= STUDENT DETAILS ============================= #
# Iterate through the tables in the document
for table in doc.tables:
    for row in table.rows:
        first_column = row.cells[0].text
        if first_column != "" and first_column != "STUDENT NAME":
            name = row.cells[0].text
            standard = row.cells[1].text
            ifsc = row.cells[2].text
            acc_no = row.cells[3].text
            holder = row.cells[4].text
            branch = row.cells[5].text

            # Extracted data
            print(f"{name},{standard},{ifsc},{acc_no},{holder},{branch}")
