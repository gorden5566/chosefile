# -*- coding: utf-8 -*-
# /usr/bin/python3

import configparser
import os


class Setting:

    def __init__(self):
        self.config_name = "config.ini"
        self.section = "default"
        self.config = self.build_config()

    def get_source_dir(self):
        if self.has_not_config():
            return os.getcwd()
        return self.config.get(self.section, "sourceDir")

    def get_target_dir(self):
        if self.has_not_config():
            return os.getcwd()
        return self.config.get(self.section, "targetDir")

    def get_ext_name(self):
        if self.has_not_config():
            return ".dwg"
        return self.config.get(self.section, "extName")

    def get_column_title(self):
        if self.has_not_config():
            return "图号"
        return self.config.get(self.section, "columnTitle")

    def get_excel_path(self):
        if self.has_not_config():
            return ""
        return self.config.get(self.section, "excelPath")

    def get_max_depth(self):
        if self.has_not_config():
            return 5
        depth = self.config.get(self.section, "maxDepth")
        try:
            return int(depth)
        except ValueError:
            return 5

    def has_not_config(self):
        if self.config is None:
            return True
        return False

    # 读取配置文件
    def build_config(self):
        current_path = os.getcwd()
        config_path = os.path.join(current_path, self.config_name)
        if not os.path.exists(config_path):
            self.init_config(config_path)

        conf = configparser.ConfigParser()
        conf.read(config_path, encoding="utf-8")
        return conf

    # 设置配置
    def set_config(self, key, value):
        self.config.set(self.section, key, value)

    # 设置图库文件夹
    def set_source_dir(self, dir):
        self.config.set(self.section, "sourceDir", dir)

    # 设置清单文件
    def set_excel_file(self, dir):
        self.config.set(self.section, "excelPath", dir)

    # 设置目标文件夹
    def set_target_dir(self, dir):
        self.config.set(self.section, "targetDir", dir)

    # 保存配置
    def save_config(self):
        current_path = os.getcwd()
        config_path = os.path.join(current_path, self.config_name)
        with open(config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

        return True

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
        with open(config_path, 'w', encoding='utf-8') as configfile:
            configfile.write(content)
            configfile.close()

    def get_config_val(self, conf, section, key, default):
        try:
            val = conf.get(section, key)
            if val is None or val == "":
                return default
            return val
        except BaseException:
            return default
