# -*- coding: utf-8 -*-
import xlrd
import os
import zipfile


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

    return data

def init():
    upload_list = os.listdir("./upload")

    for file_name in upload_list:
        os.remove(f"./upload/{file_name}")

    file_list = os.listdir("./data")

    for file_name in file_list:
        os.remove(f"./data/{file_name}")


def make_html(data):

    base = """
        <table style="width:60%">
            <tr>
                <th>품명</th>
                <th>색상+사이즈</th>
                <th>수량</th>
            </tr>
    """

    for i in range(len(data)):
        name, color, num = data[i][0], data[i][1], data[i][2]
        base = base + f"""
            <tr>
                <td>{name}</td>
                <td>{color}</td>
                <td>{num}</td>
            </tr>
        """

    base = base + "</table>"
    return base
