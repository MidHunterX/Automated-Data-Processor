import pdfplumber   # PDF parsing
import docx         # Docx parsing
import csv          # CSV file manipulation
import os           # Directory path support
import glob         # Finding files with extensions


def main():
    input_dir = "test"
    csv_file = "output.csv"

    # ------------------------------------------------- [ PDF FILE PROCESSING ]

    pdf_file_list = getPdfFileList(input_dir)
    if pdf_file_list:
        for pdf_file in pdf_file_list:
            file_name = os.path.basename(pdf_file)
            print(f"\n==== {file_name} ====")
            institution = getInstitutionDetailsPdf(pdf_file)
            student_data = getStudentDetailsPdf(pdf_file)
            printInstitution(institution)
            printStudentData(student_data)
            writeToCSV(csv_file, institution, student_data)

    # ------------------------------------------------ [ DOCX FILE PROCESSING ]

    docx_file_list = getDocxFileList(input_dir)
    if docx_file_list:
        for docx_file in docx_file_list:
            file_name = os.path.basename(docx_file)
            print(f"\n==== {file_name} ====")
            institution = getInstitutionDetailsDocx(docx_file)
            student_data = getStudentDetailsDocx(docx_file)
            printInstitution(institution)
            printStudentData(student_data)
            writeToCSV(csv_file, institution, student_data)


# =============================== PROCEDURES ================================ #


def writeToCSV(csv_file, institution, student_data):
    """
    Parameter: (csv_file, institution, student_data)
    Returns: CSV File in working directory
    """
    with open(csv_file, mode="a", newline="") as file:
        writer = csv.writer(file)
        # Institution details
        for value in institution.values():
            writer.writerow([value])
        # Student details
        for row in student_data.values():
            writer.writerow(row)


def printInstitution(institution):
    inst_name = institution["name"]
    inst_place = institution["place"]
    inst_number = institution["number"]
    inst_email = institution["email"]
    print(f"{inst_name},{inst_place},{inst_number},{inst_email}")


def printStudentData(student_data):
    for key, value in student_data.items():
        name = value[0]
        standard = value[1]
        ifsc = value[2]
        acc_no = value[3]
        holder = value[4]
        branch = value[5]
        print(f"{name},{standard},{ifsc},{acc_no},{holder},{branch}")


# ================================ FUNCTIONS ================================ #


def countPdfPages(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        page_count = len(pdf.pages)
    return page_count


def getStudentDetailsPdf(pdf_file):
    """
    Parameter: PDF File
    Returns: A dictionary of tuples with Student details

    data = {
        0: (name, standard, ifsc, acc_no, holder, branch),
        1: (name, standard, ifsc, acc_no, holder, branch),
        2: (name, standard, ifsc, acc_no, holder, branch)
    }
    """
    with pdfplumber.open(pdf_file) as pdf:
        data = {}
        i = 0
        for page in pdf.pages:
            # Generate CSV list from PDF table
            table = page.extract_table()
            if table:
                for row in table:
                    # Replace \n substring with space
                    cleaned_row = []
                    for cell in row:
                        if isinstance(cell, str):
                            cleaned_row.append(cell.replace('\n', ' '))
                        else:
                            cleaned_row.append(cell)

                    name = cleaned_row[0]
                    standard = cleaned_row[1]
                    ifsc = cleaned_row[2]
                    acc_no = cleaned_row[3]
                    holder = cleaned_row[4]
                    branch = cleaned_row[5]

                    # Extracted data
                    if name:  # For avoiding empty rows
                        data[i] = name, standard, ifsc, acc_no, holder, branch
                        i = i + 1
    return data


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
