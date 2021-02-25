import sys
import xlfileread

def after_book_open(workbook):
    print('after_book_open: callback')
    return True

def before_sheet_open(sheet):
    print('before_sheet_open: ' + str(sheet['name']))
    # if sheet['name'] != '汇总表':
    #     return False
    return True

row_list = []
is_first_row = True
is_empty_row = True
cell_count = 0

sql = ''
cell_values = []
def before_row_open(sheet):
    global is_empty_row, is_first_row, cell_count, sql, cell_values
    # print('before_row_open: callback')

    # reference: <https://www.cnblogs.com/BackingStar/p/10986775.html> write-on-copy
    # row_list = []
    row_list.clear()
    is_empty_row = True
    cell_count = 0
    sql = 'insert into t_projects values (NULL, '
    cell_values = []
    if is_first_row:
        is_first_row = False
        return False
    return True

def cell_open(sheet):
    global is_empty_row, cell_count, sql
    # print('cell_open: ' + str(sheet['cell'].value))
    cell = sheet['cell']
    cell_index = sheet['cell_index']

    cell_count += 1
    if cell_count >= 16:
        return False

    value = ''
    if str(cell.value).startswith('='):
        value = '-'
        is_empty_row = False
    elif str(cell.value) == 'None':
        value = ''
    else:
        value = cell.value
        is_empty_row = False
    row_list.append(value)

    if cell_index == 7:
        if type(value) == float:
            value = int(value)
        value = str(value)

    if type(value) == str:
        cell_values.append(repr(value))
    else:
        cell_values.append(str(value))
    return True

def after_row_open(sheet):
    global is_empty_row, sql
    # print('after_row_open: callback')
    if not is_empty_row:
        # print(str(row_list))
        sql += ', '.join(cell_values) + ');'
        # print(', '.join(cell_values))
        print(sql)
    return True

xlfileread.fileread(
    # './tmp-file/210111-210117.xls'
    # './tmp-file/210118-210124.xls'
    './tmp-file/__projects.xls'
    , 'Sheet1'
    , {
        'after_book_open': after_book_open
        , 'before_sheet_open': before_sheet_open
        , 'before_row_open': before_row_open
        , 'cell_open': cell_open
        , 'after_row_open': after_row_open
    }
)

