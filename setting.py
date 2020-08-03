# -*- coding: utf-8 -*-
# /usr/bin/python3

import configparser
import os


class Setting:

    def __init__(self):
        self.config = self.get_config("config.ini")

    def get_source_dir(self):
        if self.has_not_config():
            return os.getcwd()
        return self.config.get("sourceDir")

    def get_target_dir(self):
        if self.has_not_config():
            return os.getcwd()
        return self.config.get("targetDir")

    def get_ext_name(self):
        if self.has_not_config():
            return ".dwg"
        return self.config.get("extName")

    def get_column_title(self):
        if self.has_not_config():
            return "图号"
        return self.config.get("columnTitle")

    def get_excel_path(self):
        if self.has_not_config():
            return ""
        return self.config.get("excelPath")

    def get_max_depth(self):
        if self.has_not_config():
            return 5
        depth = self.config.get("maxDepth")
        try:
            return int(depth)
        except ValueError:
            return 5

    def has_not_config(self):
        if self.config is None:
            return True
        return False

    # 读取配置文件
    def get_config(self, config_name):
        current_path = os.getcwd()
        config_path = os.path.join(current_path, config_name)
        if not os.path.exists(config_path):
            self.init_config(config_path)

        conf = configparser.ConfigParser()
        conf.read(config_path, encoding="utf-8")

        section = "default"
        config = {'sourceDir': self.get_config_val(conf, section, "sourceDir", ""),
                  'targetDir': self.get_config_val(conf, section, "targetDir", ""),
                  'extName': self.get_config_val(conf, section, "extName", ".dwg"),
                  'columnTitle': self.get_config_val(conf, section, "columnTitle", "图号"),
                  'excelPath': self.get_config_val(conf, section, "excelPath", ""),
                  'maxDepth': self.get_config_val(conf, section, "maxDepth", 5)
                  }

        return config

    def init_config(self, config_path):
        content_arr = [
            "[default]\n",
            "; 模板文件扩展名\n",
            "extName = .dwg\n\n",
            "; 模板文件路径\n",
            "sourceDir = \n\n",
            "; 默认的目标文件夹\n",
            "targetDir = \n\n",
            "; excel文件路径\n",
            "excelPath = \n\n",
            "; 标题，只处理该标题下的单元格\n",
            "columnTitle = 图号\n\n",
            "; 索引最大深度\n",
            "maxDepth = 5\n"
        ]
        content = "".join(content_arr)
        with open('config.ini', 'w') as configfile:
            configfile.write(content)

    def get_config_val(self, conf, section, key, default):
        try:
            val = conf.get(section, key)
            if val is None or val == "":
                return default
            return val
        except BaseException:
            return default
