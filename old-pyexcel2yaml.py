# -*- coding: utf-8 -*- 

import xlrd

import codecs, sys


/*
reload(sys)
sys.setdefaultencoding('utf-8')

sys.stdout = codecs.getwriter('utf_8')(sys.stdout)
sys.stdin = codecs.getreader('utf_8')(sys.stdin)

book = xlrd.open_workbook('./Excel2YAML_1.0.0.xlsm')
print book.nsheets
for name in book.sheet_names():
    print name

inventory_sheet = book.sheet_by_index(0)



#for row_index in range(inventory_sheet.nrows):
#    for col_index in range(inventory_sheet.ncols):
#        val = inventory_sheet.cell_value(rowx=row_index, colx=col_index)
#        print('cell[{},{}] = {}'.format(row_index, col_index, val))

target = []

for row_index in range(19, inventory_sheet.nrows):
    print inventory_sheet.row_values(row_index)
    
    inventory_row = inventory_sheet.row_values(row_index)
    index = int(inventory_row[1])
    target[index].append({
        host : inventory_row[2],
        user : inventory_row[3],
        password : inventory_row[4],
        ansible : {},
        serverspec :{}
    })

    print target[index]



    
# 1~4
#    for col_index in range(inventory_sheet.ncols):
#        val = inventory_sheet.cell_value(rowx=row_index)
#        print '{}'.format(val)


*/
