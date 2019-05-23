# encoding=utf-8

import glob
import csv
import openpyxl
import chardet

def handle_csv2xlsx(src_file_name, dest_file_name):
    source = src_file_name

    f = open(source,'rb')
    data = f.read()
    encoding = chardet.detect(data).get('encoding')

    wb = openpyxl.Workbook()
    ws = wb.active
    with open(source, mode='r', encoding=encoding) as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            ws.append(row)

    wb.save(dest_file_name)

if __name__ == '__main__':
    path = "*.csv"
    for fname in glob.glob(path):
        print(fname)
        src_file_name = fname
        dest_file_name = src_file_name.split('.')[0] 
        dest_file_name = '%s_%s.xlsx' % (dest_file_name, 'converted')
        handle_csv2xlsx(src_file_name, dest_file_name)

