# -*- coding: utf-8 -*-
# /usr/bin/python3

import configparser
import os


class Setting:

    def __init__(self):
        self.config = self.getconfig("config.ini")

    def getsourcedir(self):
        return self.config.get("sourceDir")

    def gettargetdir(self):
        return self.config.get("targetDir")

    def getextname(self):
        return self.config.get("extName")

    def getcolumntitle(self):
        return self.config.get("columnTitle")

    def getexcelpath(self):
        return self.config.get("excelPath")

    # 读取配置文件
    def getconfig(self, configname):
        if not os.path.exists(configname):
            self.Log("[配置文件不存在]\t" + configname)
            return

        conf = configparser.ConfigParser()
        conf.read(configname)

        section = "default"
        config = {'sourceDir': self.getconfigval(conf, section, "sourceDir", "."),
                  'targetDir': self.getconfigval(conf, section, "targetDir", "."),
                  'extName': self.getconfigval(conf, section, "extName", ".dwg"),
                  'columnTitle': self.getconfigval(conf, section, "columnTitle", "图号"),
                  'excelPath': self.getconfigval(conf, section, "excelPath", "")
                  }

        return config

    def getconfigval(self, conf, section, key, default):
        try:
            val = conf.get(section, key)
            if val is None or val == "":
                return default
            return val
        except BaseException:
            return default
