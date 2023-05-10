"""Módulo que converte arquivos OFX em CSV"""
import codecs
import csv
from os.path import exists
from tkinter import messagebox
from tkinter.filedialog import askopenfilenames

from ofxparse import OfxParser


def search_files():
    """Busca arquivos para serem convertidos"""
    files = askopenfilenames(
        title='OfxParser - Choose a file', filetypes=[('OFX', '*.ofx')])

    if len(files) == 0:
        exit()

    open_files(files)


def open_files(files):
    """Abre arquivos selecionado"""
    for file in files:
        with codecs.open(file, encoding='latin-1') as ofx_file:
            ofx_data = OfxParser.parse(ofx_file)

        account = ofx_data.account

        csv_file = f'{file[:-4]}.csv'

        overwrite_file = True
        if exists(csv_file):
            overwrite_file = messagebox.askyesno(
                'OfxToCsv',
                f'O arquivo "{csv_file.split("/")[-1]}" já existe. Deseja subtituí-lo?'
            )

        if overwrite_file:
            create_file(csv_file, account.statement)
        else:
            break

        ofx_file.close()


def create_file(filename, statement):
    """Cria um novo csv a partir do arquivo selecionado"""
    with open(filename, 'w', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['date', 'payee', 'amount'])

        for transaction in statement.transactions:
            payee, date = filter_payee(transaction.memo, transaction.date)
            writer.writerow(
                [date, payee, transaction.amount])


def filter_payee(payee_data, date):
    """Filtra os dados da transação"""
    split_payee_data = payee_data.split(' - ')

    match split_payee_data[0]:
        case 'Compra com Cartão' | 'Transferência recebida' | 'Depósito Online TAA':
            payee = split_payee_data[1][12:]
            correct_date = split_payee_data[1][:5]

            return f'{split_payee_data[0]} - {payee}', f'{correct_date}/{date.year}'
        case 'Pix':
            correct_date = split_payee_data[2][:5]
            if 'Enviado' in split_payee_data:
                payee = split_payee_data[2].split(' ')[2:]
            else:
                payee = split_payee_data[2].split(' ')[3:]

            return f'{split_payee_data[0]} - {" ".join(payee)}', f'{correct_date}/{date.year}'
        case _:
            return payee_data, f'{date:%d}/{date:%m}/{date.year}'


search_files()
messagebox.showinfo('OfxParser', 'All files parsed')
