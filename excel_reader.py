#! /usr/bin/env python
# -*- coding: utf8 -*-

import os
import sys

try:
    import xlrd
except ImportError as e:
    print str(e)
    print "Please install xlrd"
    print "pip install xlrd\n"
    sys.exit(0)

ABC_TO_NUM = {
    "A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "M": 12,
    "N": 13, "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19, "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24,
    "Z": 25, "AA": 26, "AB": 27, "AC": 28, "AD": 29, "AE": 30, "AF": 31, "AG": 32, "AH": 33, "AI": 34, "AJ": 35,
    "AK": 36, "AL": 37, "AM": 38, "AN": 39, "AO": 40, "AP": 41, "AQ": 42, "AR": 43, "AS": 44, "AT": 45, "AU": 46,
    "AV": 47, "AW": 48, "AX": 49, "AY": 50, "AZ": 51, "BA": 52, "BB": 53, "BC": 54, "BD": 55, "BE": 56, "BF": 57,
    "BG": 58, "BH": 59, "BI": 60, "BJ": 61, "BK": 62, "BL": 63, "BM": 64, "BN": 65, "BO": 66, "BP": 67, "BQ": 68,
    "BR": 69, "BS": 70, "BT": 71, "BU": 72, "BV": 73, "BW": 74, "BX": 75, "BY": 76, "BZ": 77}


def read_excel(file_path, map, sheet=1, begin=1, end=100000):
    """

    :param file_path: 要解析的excel的路径
    :param map: dict类型 类似于 {A: ID, B: Name}
    :param sheet: 要解析的哪个Sheet
    :param begin: 从第几行开始读取数据
    :param end:
    :return:
    """
    if not os.path.exists(file_path):
        print 'File Not Found!'
        return None
    if not map:
        print 'please define the mapping data'
        return None
    if not isinstance(map, dict):
        print 'only dict data can be accepted'
        return None
    if not begin < end:
        print 'end row should greater than begin row'
        return None

    print "read file : " + file_path
    workbook = xlrd.open_workbook(file_path)
    sheet_page = workbook.sheet_by_index(sheet - 1)

    end = sheet_page.nrows if end > sheet_page.nrows else end

    data_list = []

    for row in xrange(begin, end):
        # print sheet_page.row_values(row)
        test_data = dict()
        try:
            for key in map:
                data_key = map[key]
                data_value = sheet_page.cell_value(rowx=row, colx=ABC_TO_NUM[key])
                if data_value:
                    test_data[data_key] = data_value
        except Exception, e:
            print str(e)
        else:
            data_list.append(test_data)
    return data_list


def read_excel_with_caption(file_path, sheet=1, end=65535):
    """

    :param file_path: 要解析的excel的路径
    :param map: dict类型 类似于 {A: ID, B: Name}
    :param sheet: 要解析的哪个Sheet
    :param begin: 从第几行开始读取数据
    :param end:
    :return:
    """
    print u"Read File: %s" % file_path
    if not file_path or not os.path.exists(file_path):
        print 'File Not Found!'
        return None

    workbook = xlrd.open_workbook(file_path)
    sheet_page = workbook.sheet_by_index(sheet - 1)

    end = sheet_page.nrows if end > sheet_page.nrows else end

    caption = list()
    for col in range(len(ABC_TO_NUM.keys())):
        try:
            caption.append(sheet_page.cell_value(0, col))
        except BaseException, _:
            break

    print caption
    test_data_list = list()
    for row in xrange(1, end):
        # print sheet_page.row_values(row)
        test_data = dict()
        try:
            for key in caption:
                data_value = sheet_page.cell_value(rowx=row, colx=caption.index(key))
                _value = data_value if data_value else ""
                test_data[key] = _value
            test_data_list.append(test_data)
        except Exception, e:
            print str(e)

    print str(len(test_data_list)) + " records found from \n%s \n" % file_path
    return test_data_list


if __name__ == "__main__":
    pass
