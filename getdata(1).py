import os
import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy
import re


# 读取文件夹下所有文件
def del_excel(filePath):
    f = open(filePath)

    # 生成Excel
    xls = xlwt.Workbook()
    sheet = xls.add_sheet('sheet1', cell_overwrite_ok=True)
    x = 0
    while True:
        line = f.readline()
        if not line:
            break
        for i in range(len(line.split(','))):
            item = re.sub('"', '', line.split(',')[i])
            sheet.write(x, i, item)
        x += 1

    f.close()
    xls.save(filePath.split('.')[0]+'.xlsx')


def delToExcelWithColumnName(columnFilePath, filePath):
    f = open(filePath)
    # 打开列名所在的工作簿
    workbook = xlrd.open_workbook(columnFilePath)
    # sheets = workbook.sheet_names()
    worksheet = workbook.sheet_by_index(0)
    rows_old = worksheet.nrows
    new_wordbook = copy(workbook)
    new_wordsheet = new_wordbook.get_sheet(0)

    #追加
    row = rows_old
    while True:
        line = f.readline()
        if not line:
            break
        for col in range(0,len(line.split(','))):
            item = re.sub('"', '', line.split(',')[col])
            new_wordsheet.write(row, col+1, item)
        row += 1

    f.close()
    new_wordbook.save(filePath.split('.')[0] + '.xlsx')



#
# oslist = os.listdir(filePath)
#
# for i in oslist:
#     f = open(filePath+"\"+i)
#     line = f.readline()
#     while line:
#         print(line)
#         print(type(line))
#         line = f.readline()
#     f.close()
# excel_path=filePath+'\\'+i
# data = pd.read_excel(io=excel_path)
# print(data)

# excel_path=filePath+'\\'+oslist[0]
# data = pd.read_excel(io=excel_path)
# print(data)

if __name__ == "__main__":
    filePath = r'C:\Users\Administrator\Desktop\八月\同业\PBC_CLTYCK_20210731_ADD_828.del'
    # oslist = os.listdir(filePath)
    # for i in oslist:
    #     excel_path = filePath + '\\' + i
    #     # data = pd.read_excel(io=excel_path)
    #     del_excel(excel_path)
    #     print(excel_path)
    #     print(excel_path.split('.')[0]+'.xlsx')
    columnFilePath = r'C:\Users\Administrator\Desktop\八月\col\column.xls'
    delToExcelWithColumnName(columnFilePath,filePath)