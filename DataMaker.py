import docx     # Docx parsing
import os       # Directory path support
import glob     # Finding files with extensions


def main():
    # ========================= DIRECTORY SCANNING ========================= #
    directory_path = 'test'
    docx_files = glob.glob(os.path.join(directory_path, '*.docx'))

    for docx_file in docx_files:
        print(docx_file)
        file_name = os.path.basename(docx_file)
        print(file_name)

    docx_file_path = 'test\\test.docx'

    institution = institutionDetails(docx_file_path)  # It iz le dictionary
    inst_name = institution["name"]
    inst_place = institution["place"]
    inst_number = institution["number"]
    inst_email = institution["email"]
    print(f"{inst_name}, {inst_place}, {inst_number}, {inst_email}")

    student_data = studentDetails(docx_file_path)
    for key, value in student_data.items():
        name = value[0]
        standard = value[1]
        ifsc = value[2]
        acc_no = value[3]
        holder = value[4]
        branch = value[5]
        print(f"{name}, {standard}, {ifsc}, {acc_no}, {holder}, {branch}")


# =========================== INSTITUTION DETAILS =========================== #

def institutionDetails(docx_file):
    """
    Docstring for institutionDetails
    --------------------------------
    Parameters: Document.docx file
    Returns Dictionary of following details:
    name, place, number, email
    """

    doc = docx.Document(docx_file)
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
    data = {
        "name": name_of_institution,
        "place": place,
        "number": phone_number,
        "email": email_id
    }
    return data


# ============================= STUDENT DETAILS ============================= #

def studentDetails(docx_file):
    """
    Docstring for studentDetails
    ----------------------------
    Parameter: Document.docx file
    returns a dictionary of tuples with the following details:
    (name, standard, ifsc, acc_no, holder, branch)
    """

    doc = docx.Document(docx_file)
    data = {}
    i = 0
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
                data[i] = name, standard, ifsc, acc_no, holder, branch
                i = i + 1
    return data


if __name__ == "__main__":
    main()
