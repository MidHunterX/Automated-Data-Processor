# Libraries Used
Instead of Re-Inventing the wheel, let's use a premade module for simplicity.
Let's follow the path of the people who went before.

## DOCX Module
```
pip install python-docx
```
The python-docx library is one of the most popular and widely used libraries for working with DOCX (Microsoft Word) files in Python. It provides a comprehensive set of features for creating, modifying, and extracting information from DOCX files. For many use cases, it is indeed an excellent choice.
The python-docx library is well-documented and has a strong user community, making it a reliable choice for most tasks involving DOCX files.

## PDF Plumber
To extract information from PDFs with a specific structure, you can use Python libraries such as PyPDF2, pdfplumber, or Camelot. PyPDF2 primarily extracts raw text and doesn't provide as much layout information as pdfplumber. Therefore PDF Plumber is the best choice here.
```
pip install pdfplumber
```
### PDF Plumber usage example
```py
import pdfplumber

# Replace 'your_pdf.pdf' with the path to your PDF file.
pdf_file = 'your_pdf.pdf'

with pdfplumber.open(pdf_file) as pdf:
    for page in pdf.pages:
        text = page.extract_text()

        # Check for the section headers and extract the content.
        if "Institution Details" in text:
            start = text.index("Institution Details")
            end = text.index("Student Details")
            institution_details = text[start:end]

        if "Student Details" in text:
            table = page.extract_table()

# Print or process the extracted information.
print("Institution Details:")
print(institution_details)

print("Student Details:")
print(student_details)
```

## Camelot
PDF Plumber might be better than PyPDF2 but, for table extractions, it gives unusable data.
Therefore Camelot is superior than PDF Plumber.
Camelot gives sophisticated controls for data extraction specifically.
Even though Camelot gives advanced control, we are still going to use pdf plumber as it provides basic extraction.

## RazorPay IFSC Dataset
The RazorPay IFSC Dataset is a comprehensive and up-to-date collection of Indian Financial System Code (IFSC) information, provided by RazorPay, a leading payment gateway and financial services company. This dataset contains detailed information about IFSC codes, which are unique identifiers for individual bank branches in India. It includes data such as bank names, branch names, addresses, and other relevant details.

# Requirements
- [x] DOCX Parsing
- [x] PDF Parsing
- [x] DOCX Data Extraction
- [x] PDF Data Extraction
- [x] Extracted Data to CSV conversion
- [x] Organizing processed files
- [x] Separate unsupported files
- [x] Check for changes in document structure
- [x] Full process logging

# Further Problems
- No need for file renaming when reprocessing
- Steps are needeed as Investigating faulty files are an iterative process.
- Writing to database is done after iterative checking is done.
- Functions for each filetypes reduces flexibility
- District Recognition checking from top to bottom misses Kollam if Unknown
```
BANK,IFSC,BRANCH,CENTRE,DISTRICT,STATE,ADDRESS,CONTACT,IMPS,RTGS,CITY,ISO3166,NEFT,MICR,UPI,SWIFT
Bank,IOBA0001851,THANKASERY,KOLLAM,KOLLAM,KERALA,THIRUVANANTHAPURAM,2464429,true,true,KOLLAM,IN-KL
```
- Data String Case Validation: Name (Title Case), IFSC (Upper), Holder (Title), Branch (Upper)
- Multiple word branches like: [Town, District]

# Further Requirements
- [ ] Name (Title Case), IFSC (Upper), Holder (Title), Branch (Upper)
- [x] Overhaul District recognition algorithm
    - [x] Get record from RBI Dataset
    - [x] Convert CSV record to List
    - [x] Get most frequently used (mfu) data
    - [x] Do comparison searching to mfu data with District Dataset
    - [x] Return First match
- [x] Abstract correctDocxFormat(docx_file) and correctPdfFormat(pdf_file)
- [x] Step by step processing
- [x] Step 1: Filename Renaming
    - [x] Step 1.1: Get all filenames in directory
    - [x] Step 1.2: For each filename change to incremental numbers
- [x] Step 2: Form checking and sorting
    - [x] Step 2.1: Get list of all files
    - [x] Step 2.2: For each file, Separate out unsupported files
    - [x] Step 2.3: For each file, Check the file structure
    - [x] Step 2.4: Separate out each well structured file
    - [x] Step 2.5: Separate out suspicious files for investigation
- [x] Step 3: Form to database writing

# Data Analysis

## Class Data Conversion
(best approach would be to check similarity or REGEXP)

```py
std_dataset = {
    1: [ "1", "one", "1a", "1b", "1c", "1d", "1e", "i", "1st", "first" ],
    2: [ "2", "two" "2a", "2b", "2c", "2d", "2e", "11", "ii", "2nd", "second" ],
    3: [ "3", "three", "3a", "3b", "3c", "3d", "3e", "111", "iii", "3rd", "third" ],
    4: [ "4", "four", "4a", "4b", "4c", "4d", "4e", "1v", "iv", "4th", "fourth" ],
    5: [ "5", "five", "5a", "5b", "5c", "5d", "5e", "v", "5th", "fifth", ],
    6: [ "6", "six", "6a", "6b", "6c", "6d", "6e", "v1", "vi", "6th", "sixth", ],
    7: [ "7", "seven", "7a", "7b", "7c", "7d", "7e", "v11", "vii", "7th", "seventh", ],
    8: [ "8", "eight", "8a", "8b", "8c", "8d", "8e", "v111", "viii", "8th", "eighth", ],
    9: [ "9", "nine", "9a", "9b", "9c", "9d", "9e", "1x", "ix", "9th", "nineth", ],
    10: [ "10", "ten", "10a", "10b", "10c", "10d", "10e", "x", "10th", "tenth", ],
    11: [ "11", "x1", "xi", "11th", "plus one", "plusone", "+1", "+1 science", "+1 commerce", "+1 humanities", ],
    12: [ "12", "x11", "xii", "12th", "plus two", "plustwo", "+2", "+2 science", "+2 commerce", "+2 humanities", ],
    13: [ "1 dc", "1dc", "i dc", "idc", "ist dc", "1stdc", "1st dc" ],
    14: [ "2 dc", "2dc", "ii dc", "iidc", "iind dc", "2nddc", "2nd dc" ],
    15: [ "3 dc", "3dc", "iii dc", "iiidc", "iiird dc", "3rddc", "3rd dc" ],
    16: [ "1 pg", "1pg", "i pg", "ipg", "ist pg", "1st pg", "1stpg" ],
    17: [ "2 pg", "2pg", "ii pg", "iipg", "iind pg", "2ndpg", "2nd pg" ],
}
```

## District Recognition
1. Read all IFSC Code
2. Find District of all IFSC Code
3. Make a list out of districts
4. Find the most occuring value in the list
5. Return the value

```py
def most_common(a_list):
    from collections import Counter
    count = Counter(a_list)
    mostCommon = count.most_common(1)
    return mostCommon[0][0]
```

## District Recognition Algorithm 2.0
1. Store all IFSC Code into a List
2. For each IFSC Code, get CSV Record from RBI Dataset
3. Convert CSV Record into a List
4. Get Most Recurring Value (MRV) of the Record List
5. Compare Each MRV with District Dataset
6. Return first matching district for each IFSC Code
7. Store all return value into a list
8. Get (MRV) of the Return Value List

# Database Design
Tables:

Schools
    SchoolID (Primary Key)
    SchoolName
    Place
    District
    Phone
    Email

```sql
CREATE TABLE "Schools" (
	"SchoolID"	INTEGER NOT NULL UNIQUE,
	"SchoolName"	TEXT NOT NULL,
	"District"	TEXT,
	"Place"	TEXT,
	"Phone"	TEXT,
	"Email"	TEXT,
	PRIMARY KEY("SchoolID" AUTOINCREMENT)
);
```

Students
    StudentID (Primary Key)
    SchoolID (Foreign Key referencing the Schools table)
    Name
    Class
    IFSC
    AccNo
    AccHolder
    Branch
    Verified

```sql
CREATE TABLE Students (
    StudentID   INTEGER NOT NULL UNIQUE,
    SchoolID    INTEGER REFERENCES Schools (SchoolID),
    StudentName TEXT    NOT NULL,
    Class       INTEGER NOT NULL,
    IFSC        TEXT    NOT NULL,
    AccNo       TEXT    NOT NULL UNIQUE,
    AccHolder   TEXT    NOT NULL,
    Branch      TEXT    NOT NULL,
    Verified    TEXT    NOT NULL DEFAULT (False),
    PRIMARY KEY (StudentID AUTOINCREMENT),
    FOREIGN KEY (SchoolID) REFERENCES Schools (SchoolID)
);
```

## ISO 3166-2 District Abbreviations
```
Local  ISO   District Name
--------------------------------
TVM    TV    Thiruvananthapuram
KLM    KL    Kollam
PTA    PT    Pathanamthitta
ALP    AL    Alappuzha
KTM    KT    Kottayam
IDK    ID    Idukki
EKM    ER    Ernakulam
TSR    TS    Thrissur
PKD    PL    Palakkad
MLP    MA    Malappuram
KKD    KZ    Kozhikode
WYD    WA    Wayanad
KNR    KN    Kannur
KSD    KS    Kasargod
```

## Account Number Validation
RBI dictates certain rules over Indian Bank Account Number structures (9 - 18).
https://www.rbi.org.in/scripts/PublicationReportDetails.aspx?ID=695#UAN

    - Most of the banks have unique account numbers.
    - Account number length varies from 9 digits to 18 digits.
    - Most of the banks (67 out of 78) have included branch code as part of the account number structure. Some banks have product code as part of the account number structure.
    - 40 out of 78 banks do not have check digit as part of the account number structure.
    - All banks have purely numeric account numbers, except one or two foreign banks.
    - Only in the case of 20 banks, account numbers are formed without any pattern with a unique running serial number.

Indian Bank Account Number Validation Regex:
```
^\d{9,18}$
```

A better way to validate would be to select the right bank and then have checks in place as per the bank which have been outlined and analyzed by the RBI here:
https://www.rbi.org.in/scripts/PublicationReportDetails.aspx?ID=695#A3

