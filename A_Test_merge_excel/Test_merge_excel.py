# -*- coding:utf8 -*-
# xlrd模块主要用于读取Excel
import xlrd as xl
import os
import re
 
# 过滤重复的人，并保存到txt文件里
 
def readExcel(fileName="", sheetName="Sheet1"):
    # 打开fileName表格
    xls_file = xl.open_workbook(fileName)  # 打开文件
    xls_sheet = xls_file.sheet_by_name(sheetName)  # 通过工作簿名称获
    cv0 = xls_sheet.col_values(0)  # 第一列所有的值
    cv1 = xls_sheet.col_values(1)  # 第二列所有的值
    cv2 = xls_sheet.col_values(2)  # 第三列所有的值
 
    # 打开2.xlsx
    xls_file2 = xl.open_workbook("2.xlsx")  # 打开文件
    xls_sheet2 = xls_file2.sheet_by_name(sheetName)  # 通过工作簿名称获
    sheet2cv0 = xls_sheet2.col_values(0)  # 第一列所有的值
    sheet2cv1 = xls_sheet2.col_values(1)  # 第二列所有的值
 
#    print(cv0)


    for index, nameItem in enumerate(cv0):
        tmpIndex = sheet2cv0.index(nameItem)
        # print("tmpIndex=", tmpIndex)
        # “a+” 以”读写”模式打开,追加
        #with open("合并.csv", "ab+", encoding="utf-8") as f:
        with open("合并.csv", "a+") as f:
            f.writelines(
                [nameItem,  ",", cv1[index], ",", cv2[index], ",", sheet2cv1[tmpIndex], "\n"])


if __name__ == '__main__':
    readExcel("1.xlsx")
