import os
import xlwt
import xlrd
from xlutils.copy import copy
import re


class DelToExcel:
    """
    将del文件变成Excel文件
    """

    def __init__(self, targetPath=r'C:\Users\admin\Desktop\456'):
        # 如果不存在就创建
        if not os.path.exists(targetPath):
            os.mkdir(targetPath)
        self.targetPath = targetPath

    def delToExcel(self, filePath, targetPath):
        """
        没有表头的情况下为del文件创建一个对应的Excel文件
        :param targetPath: 目标路径，是一个文件夹
        :param filePath: del文件的路径，
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
                sheet.write(x, i + 1, item, style=self.formatCell(item))
            x += 1

        f.close()

        # 保存到目标文件夹
        filename = filePath.split('\\')[-1]
        xls.save(targetPath + '\\' + filename.split('.')[0] + '.xlsx')

    def delToExcelWithColumnName(self, columnFilePath, filePath, targetPath):
        """
        有表头的情况下为del文件创建一个对应的Excel文件
        :param targetPath: 目标路径，是一个文件夹
        :param columnFilePath: 表头的文件路径
        :param filePath: del文件路径
        :return: 无
        """
        # 如果没有对应的表头，那就直接造一个新的Excel出来
        if columnFilePath == "":
            self.delToExcel(filePath, targetPath)
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
                new_worksheet.write(row, col + 1, item, style=self.formatCell(item))
            row += 1

        f.close()
        # 保存到目标文件夹
        filename = filePath.split('\\')[-1]
        new_wordbook.save(targetPath + '\\' + filename.split('.')[0] + '.xlsx')

    def getPathDict(self, delFilesLongPath, colNameLongPath):
        """
        为了获取表头和文件之间的对应，会返回一个字典，一个文件对应一个表头
        :param delFilesLongPath: 文件所在的路径，应该是一个文件夹
        :param colNameLongPath: 表头所在的路径，也应该是一个文件夹
        :return: 文件和表头一一对应的字典
        """
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

    def formatCell(self, cell):
        """
        将单元格的格式进行转换
        当前主要是将数字变成常规，非数字变成文本
        :param cell: 单元格的值
        :return: 单元格对应的格式
        """
        # 数字变成常规，其他变成文本
        style = xlwt.XFStyle()
        # 如果有小数点就当成数字
        if '.' in cell:
            style.num_format_str = 'general'
        else:
            style.num_format_str = '@'
        return style

    def delToExcels(self, delFilesLongPath, colNameLongPath):
        nameDict = self.getPathDict(delFilesLongPath, colNameLongPath)
        for key, value in nameDict.items():
            self.delToExcelWithColumnName(value, key, self.targetPath)
