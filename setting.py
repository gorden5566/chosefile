# -*- coding: utf-8 -*-
# /usr/bin/python3

import configparser
import os


class Setting:

    def __init__(self, logger):
        self.logger = logger
        self.config = self.getconfig("config.ini")

    def getsourcedir(self):
        if self.hasnotconfig():
            return os.getcwd()
        return self.config.get("sourceDir")

    def gettargetdir(self):
        if self.hasnotconfig():
            return os.getcwd()
        return self.config.get("targetDir")

    def getextname(self):
        if self.hasnotconfig():
            return ".dwg"
        return self.config.get("extName")

    def getcolumntitle(self):
        if self.hasnotconfig():
            return "图号"
        return self.config.get("columnTitle")

    def getexcelpath(self):
        if self.hasnotconfig():
            return ""
        return self.config.get("excelPath")

    def getmaxdepth(self):
        if self.hasnotconfig():
            return 5
        depth = self.config.get("maxDepth")
        try:
            return int(depth)
        except ValueError:
            return 5

    def hasnotconfig(self):
        if self.config is None:
            return True
        return False

    # 读取配置文件
    def getconfig(self, configname):
        currentpath = os.getcwd()
        configpath = os.path.join(currentpath, configname)
        if not os.path.exists(configpath):
            self.logger.Log("[配置文件不存在]\t" + configpath)
            return None

        conf = configparser.ConfigParser()
        conf.read(configpath, encoding="utf-8")

        section = "default"
        config = {'sourceDir': self.getconfigval(conf, section, "sourceDir", "."),
                  'targetDir': self.getconfigval(conf, section, "targetDir", "."),
                  'extName': self.getconfigval(conf, section, "extName", ".dwg"),
                  'columnTitle': self.getconfigval(conf, section, "columnTitle", "图号"),
                  'excelPath': self.getconfigval(conf, section, "excelPath", ""),
                  'maxDepth': self.getconfigval(conf, section, "maxDepth", 5)
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
