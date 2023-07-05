from docx import Document
import aspose.words as aw
import csv
import mechanize
from time import sleep
import requests
import os
import shutil

browser = mechanize.Browser()

browser.open('https://student.sum.edu.pl/dziekanat-wydzialu-nauk-o-zdrowiu-w-katowicach/#harmonogramy')

filetypes = ['.doc','.docx']
files = []

for l in browser.links():
    for t in filetypes:
        if t in str(l):
            files.append(l)

def download_link(l):
    filename = ''
    url = l.url
    if url.find('/'):
        filename = url.rsplit('/', 1)[1]

    r = requests.get(url, allow_redirects = True)
    open(filename, mode = 'wb').write(r.content)

def move_file(filename):
    destination = ''
    if filename.endswith('.docx') or filename.endswith('.doc'):
        if filename.startswith('CMUS'):
            destination = 'Wydzialy\\WNOZK\\Coaching medyczny'
        if filename.startswith('ER'):
            destination = 'Wydzialy\\WNOZK\\Elektroradiologia'
        if filename.startswith('Fizjoterapia'):
            destination = 'Wydzialy\\WNOZK\\Fizjoterapia'
        if filename.startswith('PLL') or filename.startswith('PLUS') or filename.startswith('PLUN'):
            destination = 'Wydzialy\\WNOZK\\Pielegniarstwo'
        if filename.startswith('POL') or filename.startswith('POUS') or filename.startswith('POUN'):
            destination = 'Wydzialy\\WNOZK\\Poloznictwo'
        shutil.move(filename, destination)



filesInDir = os.listdir()
for i in filesInDir:
    move_file(i)

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
