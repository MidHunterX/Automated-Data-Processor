# Literature Survey

## StackOverflow Answer
To open a DOCX file programmatically, DOCX is a zip file contatining an XML of the document.
You can open it with 7Zip and look into xml data.
You can open the zip, read the file and parse data using ElementTree.
The advantage of this technique is that you don't need any extra python libraries installed.
```py
import zipfile
import xml.etree.ElementTree

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'

with zipfile.ZipFile('<path to docx file>') as docx:
    tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))

for table in tree.iter(TABLE):
    for row in table.iter(ROW):
        for cell in row.iter(CELL):
            print ''.join(node.text for node in cell.iter(TEXT))
```

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


# Data Analysis

## Class Data Conversion
(best approach would be to check similarity or REGEXP)

```py
std_dataset = {
    1: ["1", "i", "1st", "first", "one"],
    2: ["2", "ii", "2nd", "second", "two"],
    3: ["3", "iii", "3rd", "third", "three"],
    4: ["4", "iv", "4th", "fourth", "four"],
    5: ["5", "v", "5th", "fifth", "five"],
    6: ["6", "vi", "6th", "sixth", "six"],
    7: ["7", "vii", "7th", "seventh", "seven"],
    8: ["8", "viii", "8th", "eighth", "eight"],
    9: ["9", "ix", "9th", "nineth", "nine"],
    10: ["10", "x", "10th", "tenth", "ten"],
    11: ["11", "xi", "11th", "plus one", "+1", "plusone"],
    12: ["12", "xii", "12th", "plus two", "+2", "plustwo"],
    13: ["1dc", "1 dc", "ist dc", "i dc", "idc", "1stdc", "1st dc"],
    14: ["2dc", "2 dc", "iind dc", "ii dc", "iidc", "2nddc", "2nd dc"],
    15: ["3dc", "3 dc", "iiird dc", "iii dc", "iiidc", "3rddc", "3rd dc"],
    16: ["1pg", "1 pg", "ist pg", "i pg", "ipg", "1stpg", "1st pg"],
    17: ["2pg", "2 pg", "iind pg", "ii pg", "iipg", "2ndpg", "2nd pg"],
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

```sql
CREATE TABLE Students (
	StudentID	INTEGER NOT NULL UNIQUE,
	SchoolID INTEGER,
	StudentName	TEXT NOT NULL,
	Class	INTEGER NOT NULL,
	IFSC	TEXT NOT NULL,
	AccNo	TEXT NOT NULL,
	AccHolder	TEXT NOT NULL,
	Branch	TEXT NOT NULL,
	PRIMARY KEY(StudentID AUTOINCREMENT),
	FOREIGN KEY (SchoolID) REFERENCES Schools(SchoolID)
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

