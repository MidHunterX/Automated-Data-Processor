import pdfplumber   # PDF parsing
import docx         # Docx parsing
import os           # Directory path support
import glob         # Finding files with extensions


def main():
    input_dir = "test"

    pdf_file_list = getPdfFileList(input_dir)
    if pdf_file_list:
        for pdf_file in pdf_file_list:

            file_name = os.path.basename(pdf_file)
            print(f"==== {file_name} ====")

            institution = getInstitutionDetailsPdf(pdf_file)
            inst_name = institution["name"]
            inst_place = institution["place"]
            inst_number = institution["number"]
            inst_email = institution["email"]
            print(f"{inst_name},{inst_place},{inst_number},{inst_email}")

            student_details_pdf = getStudentDetailsPdf(pdf_file)
            for entry in student_details_pdf:
                name, standard, ifsc, acc_no, holder, branch = entry
                print(f"{name},{standard},{ifsc},{acc_no},{holder},{branch}")

    docx_file_list = getDocxFileList(input_dir)
    if docx_file_list:
        for docx_file in docx_file_list:

            file_name = os.path.basename(docx_file)
            print(f"==== {file_name} ====")

            institution = getInstitutionDetailsDocx(docx_file)
            inst_name = institution["name"]
            inst_place = institution["place"]
            inst_number = institution["number"]
            inst_email = institution["email"]
            print(f"{inst_name},{inst_place},{inst_number},{inst_email}")

            student_data = getStudentDetailsDocx(docx_file)
            for key, value in student_data.items():
                name = value[0]
                standard = value[1]
                ifsc = value[2]
                acc_no = value[3]
                holder = value[4]
                branch = value[5]
                print(f"{name},{standard},{ifsc},{acc_no},{holder},{branch}")


# ================================ FUNCTIONS ================================ #


def getStudentDetailsPdf(pdf_file):
    """
    Parameter: PDF File
    Returns: A list of lists containing student table data
    [['NAME', 'CLASS', 'IFSC', 'ACC NO', 'ACC HOLDER', 'BRANCH'],
    ['Sophia', '1', '7afc5b', '9988 7766554433', 'Isabel', 'Rwanda'],
    ['Ina', '2', '80afb1', '6 54138546513135', 'Elmer', 'Anguilla'],
    ['Ophelia', '3', 'bf 841e', '387465354116', 'Marian', 'Bahrain']]
    """
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Generate CSV list from PDF table
            table = page.extract_table()

            # Replace \n substring with space
            if table:
                cleaned_table = []
                for row in table:
                    cleaned_row = []
                    for cell in row:
                        if isinstance(cell, str):
                            cleaned_row.append(cell.replace('\n', ' '))
                        else:
                            cleaned_row.append(cell)
                    cleaned_table.append(cleaned_row)
                table = cleaned_table
    # Remove Header in List
    table.pop(0)
    return table


def getInstitutionDetailsPdf(pdf_file):
    """
    Parameters: Document.pdf file
    Returns: Dictionary of Institution Details

    data = {
        "name": name_of_institution,
        "place": place,
        "number": phone_number,
        "email": email_id
    }
    """

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Extract PDF as text
            text = page.extract_text()
            if "Institution Details" in text:
                start = text.index("Name of the Institution")
                end = text.index("Student Details")
                institution_details = text[start:end]

    # Splitting text at '\n' into a list
    lines = institution_details.split('\n')

    for line in lines:
        parts = line.split(':')
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


def getPdfFileList(dir):
    """
    Parameters: Directory path
    Returns: A list of files with pdf extension.

    [file1.pdf, file2.pdf, file3.pdf]
    """

    pdf_files = glob.glob(os.path.join(dir, '*.pdf'))
    pdf_list = []
    for docx_file in pdf_files:
        pdf_list = pdf_list + [docx_file]
    return pdf_list


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


def getInstitutionDetailsDocx(docx_file):
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


def getStudentDetailsDocx(docx_file):
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
