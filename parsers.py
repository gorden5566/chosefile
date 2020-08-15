# -*- coding: utf-8 -*-
# /usr/bin/python3

import xlrd


class Parser:
    def __init__(self, logger):
        self.logger = logger

    # 读取xls文件
    def parse_excel(self, file_path, column_title):
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)

        # 查找title位置
        title_pos = self.find_title(sheet, column_title)

        if title_pos is None:
            self.logger.Log("[解析excel失败]\t" + column_title)
            return None

        title_col = title_pos["col"]
        title_row = title_pos["row"]

        name_arr = []
        for row_index in range(title_row + 1, sheet.nrows):
            name = sheet.cell_value(row_index, title_col)
            strip_name = name.replace(' ', '')
            if strip_name == '':
                continue
            name_arr.append(strip_name)
        return name_arr

    @staticmethod
    def find_title(sheet, column_title):
        for i in range(20):
            title_arr = sheet.row_values(i)
            # 不是这一行
            if column_title not in title_arr:
                continue

            # 找出是第几列
            column_index = title_arr.index(column_title)
            if column_index >= 0:
                return {"row": i, "col": column_index}

        return None
