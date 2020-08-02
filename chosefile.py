# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import wx

from index import IndexTool
from logger import Logger
from setting import Setting
from parser import Parser
from processor import Processor


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

        # parser
        self.parser = Parser(self.logger)

        # processor
        self.processor = Processor(self.logger, self.setting)

        # 文件索引
        self.indextool = IndexTool(self.setting.get_max_depth())

        # create a menu bar
        self.MakeMenuBar()

        # status bar
        self.MakeStatusBar()

        self.init_default()

    def init_default(self):
        excel_path = self.setting.get_excel_path()
        if excel_path is None:
            return
        self.FileName.SetValue(excel_path)

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
        hasbuilddb = self.indextool.check_db()
        if not hasbuilddb:
            result = self.buildIndex()
            if not result:
                return

        fileName = self.FileName.GetValue()
        if fileName is None or fileName == '':
            wx.MessageBox("请先选择清单文件", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        if not os.path.exists(fileName):
            wx.MessageBox("文件不存在: " + fileName, "处理结果", wx.OK | wx.ICON_WARNING)
            return

        targetPath = None
        dlg = wx.DirDialog(self, message="请选择要保存的路径", defaultPath=self.setting.get_target_dir(),
                           style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            targetPath = dlg.GetPath()
        dlg.Destroy()

        if targetPath is None or targetPath == '':
            return

        self.logger.Log("[目标文件夹]\t" + targetPath)

        nameArr = self.parser.parse_excel(fileName, self.setting.get_column_title())
        if nameArr is None:
            wx.MessageBox("解析结果为空", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        # 复制文件
        total = 0
        successNum = 0
        for name in nameArr:
            sourceName = name + self.setting.get_ext_name();
            sourcePath = self.getsourcepath(sourceName)
            total += 1

            if sourcePath is None:
                self.logger.Log("[索引结果空]\t" + sourceName)
                continue

            success = self.processor.copy_file(sourcePath, targetPath, sourceName)
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
        return self.indextool.get_full_path(index)

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
            sourcedir = self.setting.get_source_dir()
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
                self.processor.export_template(os.path.join(dir_name, file_name))
                save_msg = wx.MessageDialog(self, '保存成功！', '提示')
            except FileNotFoundError:
                save_msg = wx.MessageDialog(self, '保存失败，无效的保存路径', '提示')

            save_msg.ShowModal()
            save_msg.Destroy()

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

    def buildIndex(self):
        sourcedir = self.setting.get_source_dir()
        if sourcedir is None:
            wx.MessageBox("源文件夹未设置，请先打开[config.ini]设置[sourceDir]", "提示", wx.OK | wx.ICON_WARNING)
            return False
        if not os.path.isdir(sourcedir):
            wx.MessageBox("源文件夹不存在，请先打开[config.ini]设置[sourceDir]", "提示", wx.OK | wx.ICON_WARNING)
            return False
        self.indextool.build_index(sourcedir).save()
        return True


if __name__ == '__main__':
    app = wx.App()

    frame = ChoseFile()
    frame.Show()

    app.MainLoop()
