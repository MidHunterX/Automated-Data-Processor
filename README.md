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

## Cleaning up \n from a List of List 
v1
```py
data = [
    ["cell 1\nwith\nline breaks", "cell 2", "cell 3"],
    ["cell 4", "cell 5\nwith\nline breaks", "cell 6"]
]

cleaned_data = [[cell.replace('\n', ' ') if isinstance(cell, str) else cell for cell in row] for row in data]
```
v2
```py
data = [
    ["cell 1\nwith\nline breaks", "cell 2", "cell 3"],
    ["cell 4", "cell 5\nwith\nline breaks", "cell 6"]
]

cleaned_data = []

for row in data:
    cleaned_row = []
    for cell in row:
        if isinstance(cell, str):
            cleaned_row.append(cell.replace('\n', ' '))
        else:
            cleaned_row.append(cell)
    cleaned_data.append(cleaned_row)
```

# Requirements
- [x] DOCX Parsing
- [x] PDF Parsing
- [x] DOCX Data Extraction
- [x] PDF Data Extraction
- [x] Extracted Data to CSV conversion
- [ ] Organizing processed files
- [ ] Separate unsupported files
- [ ] Check for changes in document structure
- [ ] Full process logging
