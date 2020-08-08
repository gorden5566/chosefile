# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import shutil
import xlwt

from index import IndexTool


class Processor:
    def __init__(self, logger, setting):
        self.logger = logger
        self.setting = setting
        self.index_tool = IndexTool(self.setting.get_max_depth())

    # 构建索引
    def build_index(self):
        source_dir = self.setting.get_source_dir()
        if source_dir is None:
            self.logger.Log("源文件夹未设置，请先打开[config.ini]设置[sourceDir]")
            return False
        if not os.path.isdir(source_dir):
            self.logger.Log("源文件夹不存在，请先打开[config.ini]设置[sourceDir]")
            return False
        self.index_tool.build_index(source_dir).save()
        self.logger.Log("[构建索引]\t" + source_dir)
        return True

    # 前置检查
    def pre_check(self):
        # 检查索引是否存在，若不存在则构建
        has_build_db = self.index_tool.check_db()
        if not has_build_db:
            return self.build_index()
        return True

    # 复制文件
    def do_copy_file(self, source, target, file_name):
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

    def process(self, target_path, name_arr):
        # 复制文件
        total = 0
        success_num = 0
        failed_files = []
        for name in name_arr:
            source_name = name + self.setting.get_ext_name();
            source_path = self.get_source_path(source_name)
            total += 1

            if source_path is None:
                self.logger.Log("[索引结果空]\t" + source_name)
                failed_files.append(source_name)
                continue

            success = self.do_copy_file(source_path, target_path, source_name)
            if success:
                success_num += 1
            else:
                failed_files.append(source_name)

        message = "共处理" + str(total) + "个文件，处理成功" + str(success_num) + "个"
        self.logger.Log("[结果汇总]\t" + message)
        failed_file_name = self.logger.log_failed(failed_files)
        if failed_file_name is not None:
            self.logger.Log("[失败记录]\t" + failed_file_name)
        self.logger.Log("--------------------------------------------------------------------")

        return {"total": total, "success_num": success_num, "failed_files": failed_files}

    # 查询文件所在目录
    def get_source_path(self, source_name):
        # 默认位置为 self.setting.getsourcedir()
        # 因为文件可能在子文件夹中，所以还需考虑递归遍历所有子文件夹
        # 为加快查询速度，如下为从索引中查询对应结果
        index = self.index_tool.find(source_name)
        return self.index_tool.get_full_path(index)

    # 导出模板文件
    def export_template(self, file_name, column_title):
        workbook = xlwt.Workbook()

        sheet = workbook.add_sheet("Sheet1")

        sheet.write(0, 0, column_title)
        sheet.write(1, 0, "1-1")

        workbook.save(file_name)
