import docx     # Docx parsing
import os       # Directory path support
import glob     # Finding files with extensions


def main():
    docx_file_dir = "test"
    docx_file_list = getDocxFileList(docx_file_dir)

    for docx_file in docx_file_list:

        file_name = os.path.basename(docx_file)
        print(f"==== {file_name} ====")

        institution = getInstitutionDetails(docx_file)  # It iz le dictionary
        inst_name = institution["name"]
        inst_place = institution["place"]
        inst_number = institution["number"]
        inst_email = institution["email"]
        print(f"{inst_name}, {inst_place}, {inst_number}, {inst_email}")

        student_data = getStudentDetails(docx_file)  # It iz le tuple
        for key, value in student_data.items():
            name = value[0]
            standard = value[1]
            ifsc = value[2]
            acc_no = value[3]
            holder = value[4]
            branch = value[5]
            print(f"{name}, {standard}, {ifsc}, {acc_no}, {holder}, {branch}")


# ================================ FUNCTIONS ================================ #


def getDocxFileList(dir):
    """
    Parameters: Directory path
    Returns: A list of files with docx extension.

    [file1.docx, file2.docx, file3.docx]
    """

    docx_files = glob.glob(os.path.join(dir, '*.docx'))
    docx_list = []
    for docx_file in docx_files:
        docx_list = docx_list + [docx_file]
    return docx_list


def getInstitutionDetails(docx_file):
    """
    Parameters: Document.docx file
    Returns: Dictionary of Institution Details

    data = {
        "name": name_of_institution,
        "place": place,
        "number": phone_number,
        "email": email_id
    }
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


def getStudentDetails(docx_file):
    """
    Parameter: Document.docx file
    Returns: A dictionary of tuples with Student details

    data = {
        0: (name, standard, ifsc, acc_no, holder, branch),
        1: (name, standard, ifsc, acc_no, holder, branch),
        2: (name, standard, ifsc, acc_no, holder, branch)
    }
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
