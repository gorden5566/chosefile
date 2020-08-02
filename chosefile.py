# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import shutil

import wx
import xlrd
import xlwt

from index import IndexTool
from logger import Logger
from setting import Setting


class ChoseFile(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title='ChoseFile', size=(640, 505),
                         style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN)

        # panel
        self.MakePanel()

        # logger
        self.logger = Logger(self.ConsoleContent)

        # config
        self.setting = Setting(self.logger)

        # 文件索引
        self.indextool = IndexTool(self.setting.getmaxdepth())

        # create a menu bar
        self.MakeMenuBar()

        # status bar
        self.MakeStatusBar()

        self.initDefault()

    def initDefault(self):
        excelpath = self.setting.getexcelpath()
        if excelpath is None:
            return
        self.FileName.SetValue(excelpath)

    def MakePanel(self):
        # 选择文件按钮
        self.OnSelectBtn = wx.Button(self, label='选择清单', pos=(10, 10), size=(80, 25))
        self.OnSelectBtn.Bind(wx.EVT_BUTTON, self.OnSelect)

        # 已选择的文件
        self.FileName = wx.TextCtrl(self, pos=(105, 10), size=(400, 25), style=wx.TE_READONLY)

        # 处理文件
        self.ProcessBtn = wx.Button(self, label='批量复制', pos=(10, 40), size=(80, 25))
        self.ProcessBtn.Bind(wx.EVT_BUTTON, self.OnProcess)

        # 清空控制台日志
        self.ClearBtn = wx.Button(self, label='清空日志', pos=(105, 40), size=(80, 25))
        self.ClearBtn.Bind(wx.EVT_BUTTON, self.OnClearConsoleContent)

        # 控制台
        self.ConsoleContent = wx.TextCtrl(self, pos=(10, 70), size=(605, 345), style=wx.TE_MULTILINE | wx.TE_READONLY)

    # 打开文件
    def OnSelect(self, event):
        wildcard = 'Microsoft Excel 97/2000/XP/2003 Workbook(*.xls)|*.xls|Microsoft Excel 2007/2010 Workbook(*.xlsx)|*.xlsx'
        dialog = wx.FileDialog(None, "请选择要处理的Excel文件", os.getcwd(), '', wildcard)

        if dialog.ShowModal() == wx.ID_OK:
            self.FileName.SetValue(dialog.GetPath())
            dialog.Destroy

    def OnProcess(self, event):
        # 检查索引是否存在，若不存在则构建
        hasbuilddb = self.indextool.checkdb()
        if not hasbuilddb:
            self.buildIndex()

        fileName = self.FileName.GetValue()
        if fileName is None or fileName == '':
            wx.MessageBox("请先选择清单文件", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        if not os.path.exists(fileName):
            wx.MessageBox("文件不存在: " + fileName, "处理结果", wx.OK | wx.ICON_WARNING)
            return

        targetPath = None
        dlg = wx.DirDialog(self, message="请选择要保存的路径", defaultPath=self.setting.gettargetdir(),
                           style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            targetPath = dlg.GetPath()
        dlg.Destroy()

        if targetPath is None or targetPath == '':
            return

        self.logger.Log("[目标文件夹]\t" + targetPath)

        nameArr = self.ParseXls(fileName, self.setting.getcolumntitle())
        if nameArr is None:
            wx.MessageBox("解析结果为空", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        # 复制文件
        total = 0
        successNum = 0
        for name in nameArr:
            sourceName = name + self.setting.getextname();
            sourcePath = self.getsourcepath(sourceName)
            total += 1

            if sourcePath is None:
                self.logger.Log("[索引结果空]\t" + sourceName)
                continue

            success = self.Copyfile(sourcePath, targetPath, sourceName)
            if success:
                successNum += 1

        message = "共处理" + str(total) + "个文件，处理成功" + str(successNum) + "个"
        self.logger.Log("[结果汇总]\t" + message)
        self.logger.Log("--------------------------------------------------------------------")

        wx.MessageBox(message, "处理结果", wx.OK | wx.ICON_INFORMATION)
        return

    # 查询文件所在目录
    def getsourcepath(self, sourceName):
        # 默认位置为 self.setting.getsourcedir()
        # 因为文件可能在子文件夹中，所以还需考虑递归遍历所有子文件夹
        # 为加快查询速度，如下为从索引中查询对应结果
        index = self.indextool.find(sourceName)
        return self.indextool.getfullpath(index)

    def OnClearConsoleContent(self, event):
        self.ConsoleContent.SetValue("")
        # wx.MessageBox("已清空", "处理结果", wx.OK | wx.ICON_INFORMATION)

    # 退出菜单
    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.logger.close()
        self.Close(True)

    # 生成菜单
    def MakeMenuBar(self):
        # Make the menu bar and add the two menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menuBar = wx.MenuBar()

        menuBar.Append(self.MakeFileMenu(), "&文件")
        menuBar.Append(self.MakeSettingMenu(), "&设置")
        menuBar.Append(self.MakeHelpMenu(), "&帮助")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

    # 文件菜单
    def MakeFileMenu(self):
        fileMenu = wx.Menu()

        # The "\t..." syntax defines an accelerator key that also triggers
        # the same event
        openItem = fileMenu.Append(-1, "&选择清单\tCtrl-O", "选择Excel清单文件")
        self.Bind(wx.EVT_MENU, self.OnSelect, openItem)

        exportItem = fileMenu.Append(-1, "&导出模板\tCtrl-E", "导出模板文件")
        self.Bind(wx.EVT_MENU, self.OnExportTemplate, exportItem)

        # 分隔符
        fileMenu.AppendSeparator()

        exitItem = fileMenu.Append(-1, "&退出\tCtrl-Q", "退出")
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)

        return fileMenu

    # 模板菜单
    def MakeSettingMenu(self):
        settingMenu = wx.Menu()

        buildIndexItem = settingMenu.Append(-1, "&重建索引\tCtrl-B", "修改源文件路径[sourceDir]后，需要重新构建文件索引")
        self.Bind(wx.EVT_MENU, self.OnBuildIndexTemplate, buildIndexItem)

        return settingMenu

    # 重建索引设置
    def OnBuildIndexTemplate(self, event):
        result = self.buildIndex()
        if result:
            sourcedir = self.setting.getsourcedir()
            message = "重建索引成功[" + sourcedir + "]"
            wx.MessageBox(message, "提示", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox("重建索引失败", "提示", wx.OK | wx.ICON_WARNING)

    # 帮助菜单
    def MakeHelpMenu(self):
        helpMenu = wx.Menu()

        usageItem = helpMenu.Append(-1, "&说明\tCtrl-H", "查看使用说明")
        self.Bind(wx.EVT_MENU, self.OnUsage, usageItem)

        aboutItem = helpMenu.Append(-1, "&关于", "关于")
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

        return helpMenu

    # 导出模板文件
    def OnExportTemplate(self, event):
        fd = wx.FileDialog(self, message='导出模板文件', defaultDir='', defaultFile='图号清单模板',
                           wildcard='Microsoft Excel 97/2000/XP/2003 Workbook(*.xls)|*.xls|Microsoft Excel 2007/2010 Workbook(*.xlsx)|*.xlsx',
                           style=wx.FD_SAVE)
        if fd.ShowModal() == wx.ID_OK:
            try:
                file_name = fd.GetFilename()
                dir_name = fd.GetDirectory()
                self.SaveTemplate(os.path.join(dir_name, file_name))
                save_msg = wx.MessageDialog(self, '保存成功！', '提示')
            except FileNotFoundError:
                save_msg = wx.MessageDialog(self, '保存失败，无效的保存路径', '提示')

            save_msg.ShowModal()
            save_msg.Destroy()

    # 保存模板文件
    def SaveTemplate(self, fileName):
        workbook = xlwt.Workbook()

        sheet = workbook.add_sheet("Sheet1")

        xlwt.easyxf()
        sheet.write(0, 0, self.setting.getcolumntitle())
        sheet.write(1, 0, "1-1")

        workbook.save(fileName)

    # 使用说明
    def OnUsage(self, event):
        message = "1.打开[config.ini]，配置[sourceDir]等参数\n" \
                  + "2.点击[选择]按钮，选择要处理的清单文件\n" \
                  + "3.点击[处理]按钮，选择要保存的路径，确认后执行\n" \
                  + "4.执行完后，查看日志，确认是否执行成功\n" \
                  + "注意：若修改了[sourceDir]参数，请点击[设置]->[重建索引]菜单重新构建文件索引\n"
        wx.MessageBox(message, "使用说明", wx.OK | wx.ICON_INFORMATION)

    # 关于菜单
    def OnAbout(self, event):
        """Display an About Dialog"""
        wx.MessageBox("批量复制文件", self.GetVersion(), wx.OK | wx.ICON_INFORMATION)

    # 状态栏
    def MakeStatusBar(self):
        self.CreateStatusBar()
        self.SetStatusText("欢迎使用 " + self.GetVersion())

    # 版本号
    def GetVersion(self):
        return "ChoseFile V0.0.1"

    # 复制文件
    def Copyfile(self, source, target, fileName):
        if not os.path.isdir(source):
            self.logger.Log("[源文件夹不存在]\t" + source)
            return False

        if not os.path.isdir(target):
            self.logger.Log("[目标文件夹不存在]\t" + target)
            return False

        sourceName = os.path.join(source, fileName)
        targetName = os.path.join(target, fileName)
        if not os.path.exists(sourceName):
            self.logger.Log("[文件不存在]\t" + sourceName)
            return False

        try:
            shutil.copyfile(sourceName, targetName)
            self.logger.Log("[复制成功]\t" + sourceName)
            return True
        except Exception:
            self.logger.Log("[复制失败]\t" + sourceName)

        return False

    # 读取xls文件
    def ParseXls(self, filePath, columnTitle):
        workbook = xlrd.open_workbook(filePath)
        sheet = workbook.sheet_by_index(0)

        # 查找title位置
        titlePos = self.findTitle(sheet, columnTitle)

        if titlePos is None:
            self.logger.Log("[解析excel失败]\t" + columnTitle)
            return None

        titleCol = titlePos["col"]
        titleRow = titlePos["row"]

        nameArr = []
        for rowIndex in range(titleRow + 1, sheet.nrows):
            name = sheet.cell_value(rowIndex, titleCol)
            stripName = name.replace(' ', '')
            if stripName == '':
                continue
            nameArr.append(stripName)
        return nameArr

    def findTitle(self, sheet, columnTitle):
        for i in range(20):
            titleArr = sheet.row_values(i)
            columnIndex = titleArr.index(columnTitle)
            if columnIndex >= 0:
                return {"row": i, "col": columnIndex}

        return None

    def buildIndex(self):
        sourcedir = self.setting.getsourcedir()
        if sourcedir is None:
            wx.MessageBox("源文件夹未设置，请先打开[config.ini]设置[sourceDir]", "提示", wx.OK | wx.ICON_WARNING)
            return False
        if not os.path.isdir(sourcedir):
            wx.MessageBox("源文件夹不存在，请先打开[config.ini]设置[sourceDir]", "提示", wx.OK | wx.ICON_WARNING)
            return False
        self.indextool.buildindex(sourcedir).save()
        return True


if __name__ == '__main__':
    app = wx.App()

    frame = ChoseFile()
    frame.Show()

    app.MainLoop()
