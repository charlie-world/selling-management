# -*- coding: utf-8 -*-
import xlrd
import xlwt
import os
import zipfile
from openpyxl import load_workbook


def parse_product_size(pname, status, s):
    ps = s.splitlines()
    d = []
    for p in ps:
        try:
            if p != '':
                original = p.replace(".", "").replace(" ", "")
                name = original.split("(")[0].strip().split("|")[0].strip().replace("_", "")
                num = original.split("|")[-1].replace(" ", "").replace("개", "").split("(")[0]

                d.append([pname, status, name, int(num)])
        except ValueError:
            print(ps)


    return d

def parse_xlsx(_file_name):
    wb = xlrd.open_workbook(_file_name)
    worksheet = wb.sheet_by_index(0)

    product_name = _file_name.split("_")[-1].split(".xlsx")[0]

    data = []
    result = []
    pp = {}
    for row in range(worksheet.nrows):
        if worksheet.cell_value(row, 2) != '' and worksheet.cell_value(row, 2) != '구분2':
            x = []
            x.extend(parse_product_size(product_name, worksheet.cell_value(row, 2), worksheet.cell_value(row, 4)))
            data.extend(x)

    for k in data:
        if pp.get(k[2]):
            if k[1] != '신규주문':
                pp[k[2]] = pp[k[2]]
            else:
                pp[k[2]] = pp[k[2]] + k[3]
        else:
            pp[k[2]] = k[3]

    for e,v in pp.items():
        result.append([product_name, e, v])

    return result


def save_data(file_name):
    zip_file = zipfile.ZipFile(f'./upload/{file_name}')
    zip_file.extractall('./data')

    zip_file.close()

    file_list = os.listdir("./data")

    data = []

    for file_name in file_list:
        try:
            print(file_name)
            data.extend(parse_xlsx(f"./data/{file_name}"))

        except xlrd.biffh.XLRDError as e:
            print(e)

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1')

    for i in range(len(data)):
        sheet.write(i, 0, data[i][0])
        sheet.write(i, 1, data[i][1])
        sheet.write(i, 2, data[i][2])

    workbook.save("./output/output.xlsx")


def init():
    upload_list = os.listdir("./upload")

    for file_name in upload_list:
        os.remove(f"./upload/{file_name}")

    file_list = os.listdir("./data")

    for file_name in file_list:
        os.remove(f"./data/{file_name}")
