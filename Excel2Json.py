#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2018/7/27 17:43
# @Function: 将标准的二维excel表格转化为json数据
# @Author  : Tricky

import xlrd
import json
import codecs

def read_excel(file):
    try:
        wb = xlrd.open_workbook(file)
        sheets_name = wb.sheet_names()
        # print('共有 %d 个表格'%worksheets)

        print('sheets name are:')
        for i,sheetname in enumerate(sheets_name):
            print i,sheetname
        # print([sheetname for sheetname in sheets_name])
        sheetid = int(raw_input('请输入表名对应的编号进行解析：'))
        # sn = str(sheetname).decode('utf-8') # 输入表名
        sheet = wb.sheet_by_index(sheetid)
        return sheet
    except Exception, e:
        print u'excel表格读取失败：%s' % e
        return None


def save(filename,data):
    with codecs.open(filename + ".json",'w',encoding='utf-8')as f:
        f.write(data)


def excel2json(file):
    filename = file.split('.')[0]
    sheet = read_excel(file)
    keys = sheet.row(0) # .decode('unicode_escape')
    nrows = sheet.nrows  # 行号
    ncols = sheet.ncols  # 列号

    result = {}
    result["rows"] = nrows  # 行数
    for nrow in range(1,nrows):
        line = {}
        for ncol in range(ncols):
            key_en = str(keys[ncol]).decode('unicode_escape')
            key = key_en.split("'")[1]
            value = sheet.row_values(nrow)[ncol]
            line[key] = value
            # if type(value) is float:
            #     line[key] = int(value)
            # else:
            #     line[key] = value
            result[nrow+1] = line

    json_data = json.dumps(result, indent=4, sort_keys=True).decode('unicode_escape')
    print(json_data)

    save(filename, json_data)

if __name__ == "__main__":
    excelname = raw_input('请输入excel文件名：')
    en = excelname.decode('utf-8')
    excel2json(en)