import datetime
import os
import re
import sys
import time
import pandas as pd
import xlsxwriter

curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

data_path = os.path.join(curPath, 'analyze_data')


def get_source_data():
    source_data = {}
    for name in list(filter(os.path.isdir,
                            map(lambda filename: os.path.join(data_path, filename),
                                os.listdir(data_path)))):
        path = os.path.join(data_path, name)
        year_name = os.path.basename(path)
        print(year_name)
        year_data = {}
        for file in list(filter(os.path.isfile,
                                map(lambda filename: os.path.join(path, filename),
                                    os.listdir(path)))):
            file_name = os.path.join(path, file)
            print(file_name)
            range_name = os.path.basename(os.path.splitext(file_name)[0])
            if range_name.startswith("~"):
                continue
            df = pd.read_excel(file_name, sheet_name='Sheet1')
            year_data[range_name] = df
            print(range_name)
        source_data[year_name] = year_data
    return source_data


def analyze():
    source_data = get_source_data()
    result_file_name = os.path.join(data_path, "result.xlsx")
    workbook = xlsxwriter.Workbook(result_file_name)
    for year_tag in source_data.keys():
        worksheet = workbook.add_worksheet(year_tag)
        write_sheet(workbook, worksheet, year_tag, source_data)
    workbook.close()
    print("输入结果文件：{0}".format(result_file_name))
    os.startfile(result_file_name)

def write_sheet(workbook=xlsxwriter.Workbook, worksheet=xlsxwriter.Workbook.worksheet_class, year_tag=str,
                source_data=dict):
    data = source_data[year_tag]
    merge_format = workbook.add_format({'align': 'center'})
    worksheet.merge_range("A1:A2", "资产段", merge_format)
    worksheet.merge_range("B1:F1", year_tag, merge_format)
    worksheet.write("B2", "客户数")
    worksheet.write("C2", "客户数占比")
    worksheet.write("D2", "年日均资产占比")
    worksheet.write("E2", "年交易量占比")
    worksheet.write("F2", "年净拥金占比")
    start_row = "3"
    count_row = len(data.keys())
    for index, group_name in enumerate(data.keys()):
        row_index = str(index + 3)
        percent_format = workbook.add_format({'num_format': '0%'})
        data_format = workbook.add_format({'num_format': '0.00_ '})
        worksheet.write("A" + row_index, group_name)
        worksheet.write("B" + row_index, data[group_name].客户号.count())
        formula = "==B{0}/SUM(B${1}:B${2})".format(row_index, start_row, count_row)
        worksheet.write_formula("C" + row_index, formula, percent_format)
        worksheet.write("D" + row_index, data[group_name].该年日均资产.sum(), data_format)
        worksheet.write("E" + row_index, data[group_name].该年交易量.sum(), data_format)
        worksheet.write("F" + row_index, data[group_name].该年净佣金.sum(), data_format)
if __name__ == '__main__':
    print("")
    analyze()
    # io = "D:\\罗伟伟\\20200222\\数\\202002数据\\客户综合查询（20年0-5万）.xlsx"
    # print('file name:', os.path.splitext(io)[0])
    # print('file ext name:', os.path.splitext(io)[1])
    # df = pd.read_excel(io, sheet_name='Sheet1')
