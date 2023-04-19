import codecs
import csv
from ofxparse import OfxParser
from os.path import exists
from tkinter import messagebox
from tkinter.filedialog import askopenfilenames


def writeFile(filename):
    csvFile = open(filename, 'w')
    writer = csv.writer(csvFile)
    writer.writerow(['date', 'payee', 'amount'])

    for transaction in statement.transactions:
        writer.writerow(
            [transaction.date, transaction.memo, transaction.amount])

    csvFile.close()

files = askopenfilenames(
    title='OfxParser - Choose a file', filetypes=[('OFX', '*.ofx')])

if (len(files) == 0):
    exit()

for file in files:
    with codecs.open(file, encoding='latin-1') as ofxFile:
        ofx = OfxParser.parse(ofxFile)

    account = ofx.account
    statement = account.statement

    csvFile = f'{file[:-4]}.csv'

    overwriteFile = True
    if (exists(csvFile)):
        overwriteFile = messagebox.askyesno(
            'OfxParser', f'The file "{csvFile.split("/")[-1]}" already exists. Do you want to overwrite it?')

    if overwriteFile:
       writeFile(csvFile)
    else:
        break

    ofxFile.close()

messagebox.showinfo('OfxParser', 'All files parsed')
