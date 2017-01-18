#! /usr/bin/env python
# -*- coding: utf8 -*-
import os

try:
    import xlwt
except ImportError as e:
    print str(e)
    print "Please install xlwt"
    print "pip install xlwt\n"
    sys.exit(0)

import string

MAX_ROWS = 65535


def create_xls(rows=None, titles=None, path=None, keys=None, dict_list=None, name=None):
    """

    :param rows: 每一行的数据 [[1, 2, 3, 4], [1, 2, 3, 4]]
    :param titles: list 第一列数据 格式 ['1', '2', '3', '4']
    :param path: 文件存放路径 默认当前路径
    :param keys: list 如果数据是按照字典列表传递 ['a', 'b', 'c', 'd']
    :param dict_list: 一个字典列表 [{'a': 1, 'b': 2, 'c': 3, 'd': 4}, {'a': 1, 'b': 2, 'c': 3, 'd': 4}]
    :param name: 文件名字 默认 result.xls
    :return:
    """
    print "Create xls file"
    if not path:
        path = os.path.dirname(os.path.abspath(__file__))
    if not os.path.exists(path):
        os.makedirs(path)
    if not name:
        name = "result.xls"
    else:
        if not name.endswith(".xls"):
            name = u"%s%s" % (name, ".xls")

    all_rows = []
    if keys and dict_list:
        for dict_data in dict_list:
            row = []
            for key in keys:
                if key in dict_data:
                    row.append(dict_data[key])
                else:
                    row.append(u"空")
            all_rows.append(row)
    elif rows:
        all_rows = rows
    else:
        raise ValueError("Please input correct values")

    if len(all_rows) >= MAX_ROWS:
        raise ValueError("rows number limit is %s, actual %s" % (MAX_ROWS, len(all_rows)))

    w = xlwt.Workbook(encoding='utf8')
    sheet_name = "Sheet1"
    report = w.add_sheet(sheet_name)

    start = 0

    if titles:
        start = 1
        col = 0
        for title in titles:
            report.write(0, col, title)
            col += 1

    row = start
    for row_data in all_rows:
        # print "write line : %s" % str(row)
        col = 0
        for data in row_data:
            if isinstance(data, list):
                data = '\n'.join(str(item) for item in data)
            report.write(row, col, data)
            col += 1
        row += 1

    filename = os.path.join(path, name)
    w.save(filename)

    return filename


def create_xls_by_dict_list(dict_list, name, path):
    """

    :param dict_list:
    :param name:
    :param path:
    :return:
    """
    print "create_xls_by_dict_list"
    if not dict_list:
        print u"请传入合理的非空字典列表"
        return None

    titles = dict_list[0].keys()
    keys = dict_list[0].keys()

    return create_xls(titles=titles, keys=keys, dict_list=dict_list, name=name, path=path)


if __name__ == "__main__":
    pass
