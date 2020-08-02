# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import shutil
import xlwt


class Processor:
    def __init__(self, logger):
        self.logger = logger

    # 复制文件
    def copy_file(self, source, target, file_name):
        if not os.path.isdir(source):
            self.logger.Log("[源文件夹不存在]\t" + source)
            return False

        if not os.path.isdir(target):
            self.logger.Log("[目标文件夹不存在]\t" + target)
            return False

        source_name = os.path.join(source, file_name)
        target_name = os.path.join(target, file_name)
        if not os.path.exists(source_name):
            self.logger.Log("[文件不存在]\t" + source_name)
            return False

        try:
            shutil.copyfile(source_name, target_name)
            self.logger.Log("[复制成功]\t" + source_name)
            return True
        except Exception:
            self.logger.Log("[复制失败]\t" + source_name)

        return False

    # 导出模板文件
    def export_template(self, file_name, column_title):
        workbook = xlwt.Workbook()

        sheet = workbook.add_sheet("Sheet1")

        sheet.write(0, 0, column_title)
        sheet.write(1, 0, "1-1")

        workbook.save(file_name)
