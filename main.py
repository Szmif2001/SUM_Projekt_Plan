from docx import Document
import aspose.words as aw
import csv


Doc = aw.Document('_Plan\\Fizjoterapia-jmgr-1r-z.doc')
Doc.save('Fizjoterapia-jmgr-1r-z.docx')
Docx = Document('Fizjoterapia-jmgr-1r-z.docx')

rowNo = 0
collumnNo = 0
toCSVTemp = []
toCSV = []

for table in Docx.tables:
    for row in table.rows:
        temp = []
        for cell in row.cells:
            stringUTF = cell.text.encode('utf8')
            temp.append(stringUTF)
        toCSVTemp.append(temp)

dni = toCSVTemp[0]
zaj = toCSVTemp[1]

for i in range(0, 5, 1):
    strings = zaj[i].decode('utf8').split('\n')
    temp = [dni[i].decode('utf8'), strings]
    toCSV.append(temp)
