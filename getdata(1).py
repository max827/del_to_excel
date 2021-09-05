import os
import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy
import re
import tarfile
import gzip


# 读取文件夹下所有文件
def del_excel(filePath):
    f = open(filePath)

    # 生成Excel
    xls = xlwt.Workbook()
    sheet = xls.add_sheet('sheet1', cell_overwrite_ok=True)
    x = 1
    while True:
        line = f.readline()
        if not line:
            break
        for i in range(len(line.split(','))):
            item = re.sub('"', '', line.split(',')[i])
            sheet.write(x, i+1, formatCell(item))
        x += 1

    f.close()
    xls.save(filePath.split('.')[0] + '.xlsx')


def delToExcelWithColumnName(columnFilePath, filePath):
    #如果没有对应的表头，那就直接造一个新的Excel出来
    if(columnFilePath == ""):
        del_excel(filePath)
        return
    # 如果有对应的表头，就在表头后面追加内容
    f = open(filePath)
    # 打开表头所在的工作簿
    workbook = xlrd.open_workbook(columnFilePath)
    worksheet = workbook.sheet_by_index(0)
    rows_old = worksheet.nrows
    new_wordbook = copy(workbook)
    new_worksheet = new_wordbook.get_sheet(0)

    # 追加
    row = rows_old
    while True:
        line = f.readline()
        if not line:
            break
        for col in range(0, len(line.split(','))):
            item = re.sub('"', '', line.split(',')[col])
            new_worksheet.write(row, col + 1, formatCell(item))
        row += 1

    f.close()
    new_wordbook.save(filePath.split('.')[0] + '.xlsx')


def unTar(fileName):
    """
       解压tar.gz文件
       :param fname: 压缩文件名
       :param dirs: 解压后的存放路径
       :return: bool
    """
    try:
        t = tarfile.open(fileName)
        names = t.getnames()
        dirs = fileName + "_files"
        # 创建文件夹用来存放解压后的文件
        if os.path.isdir(dirs):
            pass
        else:
            os.mkdir(dirs)
        for name in names:
            t.extract(name, dirs)
        # t.extractall(path = dirs)
        t.close()
        return True
    except Exception as e:
        print(e)
        return False

def getPathDict(delFilesLongPath, colNameLongPath):
    delFilesPathList = os.listdir(delFilesLongPath)
    colNamePathList = os.listdir(colNameLongPath)
    dict = {}
    for delFilesPath in delFilesPathList:
        # 如果文件名包含del后缀，再对这个文件进行操作
        if "del" in delFilesPath:
            for colNamePath in colNamePathList:
                # 保证这个文件是表头
                if "xls" in colNamePath:
                    colName = colNamePath.split('.')[0]
                    colPath = colNameLongPath + '\\' + colNamePath
                    delPath = delFilesLongPath + '\\' + delFilesPath
                    if colName in delFilesPath:
                        dict[delPath] = colPath
                        break
                    else:
                        dict[delPath] = ""
    return dict


def formatCell(cell):
    '''
    将单元格的格式进行转换
    :param cell: 单元格的值
    :return: 转换好的单元格值
    '''
    newCell = cell

    return newCell

if __name__ == "__main__":
    delFilesLongPath = r'C:\Users\admin\Desktop\PBC\delFiles'
    colNameLongPath = r'C:\Users\admin\Desktop\PBC\ColumnName'

    dict = getPathDict(delFilesLongPath, colNameLongPath)
    for key, value in dict.items():
        # print(key+": "+value)
        delToExcelWithColumnName(value, key)
