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
## DOCX Module
Instead of Re-Inventing the wheel, let's use a premade module for simplicity.
Let's follow the path of the people who went before.
```
pip install python-docx
```

# Requirements
- [ ] Support for PDF and DOCX
- [x] Collects required data
- [ ] CSV export
