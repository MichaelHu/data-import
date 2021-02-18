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
def before_row_open(sheet):
    global is_empty_row
    # print('before_row_open: callback')

    # reference: <https://www.cnblogs.com/BackingStar/p/10986775.html> write-on-copy
    # row_list = []
    row_list.clear()
    is_empty_row = True
    return True

def cell_open(sheet):
    global is_empty_row
    # print('cell_open: ' + str(sheet['cell'].value))
    cell = sheet['cell'];
    if str(cell.value).startswith('='):
        row_list.append('-')
        is_empty_row = False
    elif str(cell.value) == 'None':
        row_list.append('')
    else:
        row_list.append(cell.value)
        is_empty_row = False
    return True

def after_row_open(sheet):
    global is_empty_row
    # print('after_row_open: callback')
    if not is_empty_row:
        print(str(row_list))
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

