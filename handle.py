# encoding=utf-8

from openpyxl import load_workbook
from openpyxl import Workbook

from stat_services import handle_sen, \
        anylizer, \
        get_anylized_word_map, \
        get_sen_list_by_word


WB = Workbook()


def handle():
    wb = load_workbook('content_1.XLSX')
    source_sheet = None
    for sheet in wb:
        #print (sheet.title)
        source_sheet = sheet.title

    sheet = wb.get_sheet_by_name(source_sheet)
    for row in sheet.rows:
        for cell in row:
            pass
            #print (cell.value)


    for column in sheet.columns:
        for cell in column:
            if cell.value:
                pass

    handle_filter_sheet(sheet)
    handle_export(sheet) 


def bulk_insert_col(sheet, row, col, value_list):
    row = row
    for value in value_list:
        sheet.cell(column=col, row=row, value="%s" % (value))
        row = row + 1


def filter_label_list(is_filter=True):
    exclude = get_exclude_list()
    label_list = []
    anylized_workd_map = get_anylized_word_map()
    for word_dict in anylized_workd_map:
        word, info = word_dict
        column_field_name = '-'.join([word, str(info.get('count'))])
        print(word, exclude)
        if is_filter:
            if word not in exclude: 
                label_list.append(column_field_name)
        else:
            label_list.append(column_field_name)
    return label_list


def filter_label_name():
    exclude = get_exclude_list()
    name_list = []
    anylized_workd_map = get_anylized_word_map()
    for word_map in anylized_workd_map:
        word, _= word_map
        if word not in exclude:
            name_list.append(word)
    return name_list
    

def init_data(sheet):
    sen_list = [] 
    for i in range(4, 10):
        if sheet.cell(i, 1).value is not None:
            sen = sheet.cell(i, 1).value.strip() 
            if sen not in sen_list:
                sen_list.append(sen) 
                handle_sen(sen)


def add_col(analyze_sheet):
    label_name_list = filter_label_name()
    row = 2
    col = 1
    for word in label_name_list:
        row_data_list = []
        row_data_list = get_sen_list_by_word(word)
        #print('word=%s, row_data_list=%s' % (word, row_data_list))
        if row_data_list:
            bulk_insert_col(analyze_sheet, row, col, row_data_list)
        col = col + 1


def get_exclude_list():
    exclude_list = []
    try:
        wb = load_workbook('result.XLSX')
        filter_sheet = wb.get_sheet_by_name('filter')
        for row in range(2, filter_sheet.max_row):
            value = filter_sheet.cell(row=row, column=2).value
            if value and value not in exclude_list:
                exclude_list.append(value)
    except Exception:
        pass
    return exclude_list


def handle_export(sheet):
    wb = load_workbook('result.XLSX')
    analyze_sheet = wb.create_sheet("analyze_word", index=2)

    label_list = filter_label_list()
    analyze_sheet.append(label_list)
    add_col(analyze_sheet)

    analyze_sheet.title = "analyze_word"
    wb.save('result.XLSX')


def handle_filter_sheet(sheet):
    init_data(sheet)
    try:
        wb = load_workbook('result.XLSX')
        filter_sheet = wb.get_sheet_by_name('filter')
    except Exception as e:
        wb = Workbook()
        filter_sheet = wb.create_sheet('filter', index=1) 

    filter_sheet.cell(row=1, column=1, value='word') 
    filter_sheet.cell(row=1, column=2, value='exclude') 
    word_list = filter_label_list(is_filter=False)
    bulk_insert_col(filter_sheet, 2, 1, word_list)
    wb.save('result.XLSX')


if __name__=='__main__':
    handle()
