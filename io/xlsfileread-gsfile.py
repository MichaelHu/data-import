import os
import sys
import xlfileread

sheet_count = 0
row_count = 0
project_names = []

def after_book_open(workbook):
    # print('after_book_open: callback')
    return True

def before_sheet_open(sheet):
    global row_count, sheet_count, project_names
    row_count = 0
    sheet_count += 1
    project_names.clear()
    # print('before_sheet_open: ' + str(sheet['name']))
    # if sheet['name'] != '汇总表':
    #     return False

    if sheet_count == 2:
        return True
    else:
        return False


row_list = []
is_first_row = True
is_empty_row = True
cell_count = 0

sql = ''
cell_values = []
def before_row_open(sheet):
    global row_count, is_empty_row, is_first_row, cell_count, sql, cell_values
    row_count += 1

    # if row_count > 10:
    #     return False

    row_index = sheet['row_index']
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
    row_index = sheet['row_index']
    cell = sheet['cell']
    cell_index = sheet['cell_index']

    cell_count += 1
    # if cell_count >= 16:
    #     return False

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

    # if type(value) == str:
    #     cell_values.append(repr(value))
    # else:
    #     cell_values.append(str(value))
    cell_values.append(str(value))

    if row_index == 2:
        project_names.append(value)

    if type(cell.value) == float and row_index > 3:
        project_name = project_names[cell_index-1]
        cell_2_value = cell_values[1]
        row_content = []
        if ( cell.value > 0 and cell_2_value != '' and cell_2_value != "''"
                and project_name != '' ): 
            row_content.append(str(date_range))
            row_content.append(str(cell_values[0]))
            row_content.append(str(cell_values[1]))
            row_content.append(str(cell_values[2]))
            row_content.append(str(value))
            row_content.append(str(project_name))
            print('\t'.join(row_content))

    return True

def after_row_open(sheet):
    global is_empty_row, sql
    # print('after_row_open: callback')
    if not is_empty_row:
        # print(str(row_list))
        sql += ', '.join(cell_values) + ');'
        # print(', '.join(cell_values))
        # print(sql)
    return True

date_range = ''
if len(sys.argv) > 2:
    input_file = sys.argv[1]
    cur_dir = os.path.abspath(os.getcwd())
    file_path = os.path.join(cur_dir, input_file)
    date_range = sys.argv[2]
    # print(file_path, date_range)
    xlfileread.fileread(
        # './tmp-file/210111-210117.xls'
        # './tmp-file/210118-210124.xls'
        # './tmp-file/__210104-210110.xls'
        file_path
        , 'Sheet1'
        , {
            'after_book_open': after_book_open
            , 'before_sheet_open': before_sheet_open
            , 'before_row_open': before_row_open
            , 'cell_open': cell_open
            , 'after_row_open': after_row_open
        }
    )

