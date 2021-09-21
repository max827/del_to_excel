import os
import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy
import re
import tarfile
import gzip


# 读取文件夹下所有文件
def del_excel(filePath, targetPath):
    """
    :param filePath: del文件路径
    :return: 无
    """
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
            sheet.write(x, i + 1, item, style=formatCell(item))
        x += 1

    f.close()

    # 保存到目标文件夹
    filename = filePath.split('\\')[-1]
    xls.save(targetPath + '\\' + filename.split('.')[0] + '.xlsx')


def delToExcelWithColumnName(columnFilePath, filePath, targetPath):
    """
    :param columnFilePath: 表头的文件路径
    :param filePath: del文件路径
    :return: 无
    """
    # 如果没有对应的表头，那就直接造一个新的Excel出来
    if columnFilePath == "":
        del_excel(filePath, targetPath)
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
            new_worksheet.write(row, col + 1, item, style=formatCell(item))
        row += 1

    f.close()
    # 保存到目标文件夹
    filename = filePath.split('\\')[-1]
    new_wordbook.save(targetPath + '\\' + filename.split('.')[0] + '.xlsx')


def getPathDict(delFilesLongPath, colNameLongPath):
    delFilesPathList = os.listdir(delFilesLongPath)
    colNamePathList = os.listdir(colNameLongPath)
    nameDict = {}
    for delFilesPath in delFilesPathList:
        # 如果文件名包含del后缀，再对这个文件进行操作
        if "del" in delFilesPath and ".gz" not in delFilesPath:
            for colNamePath in colNamePathList:
                # 保证这个文件是表头
                if "xls" in colNamePath:
                    colName = colNamePath.split('.')[0]
                    colPath = colNameLongPath + '\\' + colNamePath
                    delPath = delFilesLongPath + '\\' + delFilesPath
                    if colName in delFilesPath:
                        nameDict[delPath] = colPath
                        break
                    else:
                        nameDict[delPath] = ""
    return nameDict


def formatCell(cell):
    """
    将单元格的格式进行转换
    :param cell: 单元格的值
    :return: 单元格对应的格式
    """
    # 数字变成常规，其他变成文本
    style1 = xlwt.XFStyle()
    # 如果有小数点就当成数字
    if '.' in cell:
        style1.num_format_str = 'general'
    else:
        style1.num_format_str = '@'

    return style1


def unTar(fileName):
    """
       解压tar.gz文件
       :param fname: 压缩文件名
       :param dirs: 解压后的存放路径
       :return: bool
    """
    try:
        t = tarfile.open(fileName)
        t.extractall()
        # names = t.getnames()
        # dirs = fileName + "_files"
        # # 创建文件夹用来存放解压后的文件
        # if os.path.isdir(dirs):
        #     pass
        # else:
        #     os.mkdir(dirs)
        # for name in names:
        #     t.extract(name, dirs)
        # t.close()
        return True
    except Exception as e:
        print(e)
        return False


def uncompress(filesPath):
    fileslist = os.listdir(filesPath)
    for path in fileslist:
        # 只找压缩文件
        if 'tar' in path:
            filepath = filesPath + '\\' + path
            if os.path.isfile(filepath):
                unTar(filepath)

if __name__ == "__main__":
    # delFilesLongPath = r'C:\Users\admin\Desktop\PBC\delFiles'
    # colNameLongPath = r'C:\Users\admin\Desktop\PBC\ColumnName'
    # targetPath = r'C:\Users\admin\Desktop\456'
    #
    # nameDict = getPathDict(delFilesLongPath, colNameLongPath)
    # for key, value in nameDict.items():
    #     print(key+": "+value)
    #     delToExcelWithColumnName(value, key, targetPath)
    filespath = r'C:\Users\admin\Desktop\PBC\delFiles'
    uncompress(filespath)
