# -*- coding: utf-8 -*-
# /usr/bin/python3
import os
import configparser
import xlrd as xd
import shutil as st
import tkinter as tk
from tkinter import filedialog


class Tools:
    # 读取xls文件
    def parseXls(self, filePath, columnIndex, columnTitle):
        workbook = xd.open_workbook(filePath)
        worksheet = workbook.sheet_by_index(0)

        ## 计算实际序号
        columnId = ord(columnIndex.upper()) - 65

        nameArr = []
        for rowIndex in range(0, worksheet.nrows):
            name = worksheet.cell_value(rowIndex, columnId)
            stripName = name.replace(' ', '')
            if stripName == '' or stripName == columnTitle:
                continue
            nameArr.append(name)
        return nameArr

    # 复制文件
    def copyfile(self, sourceDir, targetDir, fileName):
        if not os.path.isdir(sourceDir):
            print("源文件夹不存在")

        if not os.path.isdir(targetDir):
            print("目标文件夹不存在")

        sourceName = os.path.join(sourceDir, fileName)
        targetName = os.path.join(targetDir, fileName)
        if not os.path.exists(sourceName):
            print("源文件不存在：" + sourceName)

        st.copyfile(sourceName, targetName)
        print('成功复制：' + targetName)

    # 读取配置文件
    def getConfig(self, configName):
        if not os.path.exists(configName):
            print("配置文件[" + configName + "]不存在")
            return

        conf = configparser.ConfigParser()
        conf.read(configName)

        section = "default"
        option = conf.options('default')
        properties = {}
        properties['excelPath'] = conf.get(section, "excelPath")
        properties['sourceDir'] = conf.get(section, "sourceDir")
        properties['targetDir'] = conf.get(section, "targetDir")
        properties['extName'] = conf.get(section, "extName")
        properties['columnIndex'] = conf.get(section, "columnIndex")
        properties['columnTitle'] = conf.get(section, "columnTitle")

        return properties

    def process(self):
        # 读取配置文件
        configName = "config.ini"
        properties = self.getConfig(configName)
        excelPath = properties['excelPath']
        sourceDir = properties['sourceDir']
        targetDir = properties['targetDir']
        extName = properties['extName']
        columnIndex = properties['columnIndex']
        columnTitle = properties['columnTitle']

        # 未配置excel文件路径，弹出选择文件对话框
        if excelPath == '':
            root = tk.Tk()
            root.withdraw()
            excelPath = filedialog.askopenfilename(filetypes=[('Excel', 'xls')], title='请选择要处理的Excel文件')
            # 未选择文件
            if excelPath == '':
                return

        # 为配置目标文件夹，默认创建与文件名相同的文件夹
        if targetDir == '':
            filePath = os.path.dirname(excelPath)
            fileBaseName = os.path.basename(excelPath)
            fileName = os.path.splitext(fileBaseName)[0]
            targetDir = os.path.join(filePath, fileName)

        if not os.path.exists(targetDir):
            os.makedirs(targetDir)

        # 解析excel，提取模板名
        nameArr = self.parseXls(excelPath, columnIndex, columnTitle)

        # 复制文件
        for name in nameArr:
            self.copyfile(sourceDir, targetDir, name + extName)

        input("按任意按键继续！");


# main方法
if __name__ == "__main__":
    tools = Tools()
    tools.process();
