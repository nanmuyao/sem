# encoding=utf-8

import csv
import openpyxl
import chardet

def handle_csv2xlsx():
    source = 'content_1.csv'

    f = open(source,'rb')
    data = f.read()
    encoding = chardet.detect(data).get('encoding')

    wb = openpyxl.Workbook()
    ws = wb.active
    with open(source, mode='r', encoding=encoding) as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            ws.append(row)

    wb.save('content_1.xlsx')

if __name__ == '__main__':
    handle_csv2xlsx()
