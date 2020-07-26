import os

import wx
import xlwt
import shutil


class ChoseFile(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title='选图工具', size=(640, 480))

        self.MakePanel()

        # create a menu bar
        self.MakeMenuBar()

        # status bar
        self.MakeStatusBar()

    def MakePanel(self):
        # 选择文件按钮
        self.OnOpenBtn = wx.Button(self, label='打开', pos=(10, 10), size=(80, 25))
        self.OnOpenBtn.Bind(wx.EVT_BUTTON, self.OnOpen)

        # 已选择的文件
        self.FileName = wx.TextCtrl(self, pos=(105, 10), size=(400, 25), style=wx.TE_READONLY)

        # 读取文件
        self.SelBtn = wx.Button(self, label='解析', pos=(10, 40), size=(80, 25))
        self.SelBtn.Bind(wx.EVT_BUTTON, self.OnReadFile)

        # 处理文件
        self.ProcessBtn = wx.Button(self, label='处理', pos=(105, 40), size=(80, 25))
        self.ProcessBtn.Bind(wx.EVT_BUTTON, self.OnProcess)

        # 清空控制台日志
        self.ClearBtn = wx.Button(self, label='清空日志', pos=(200, 40), size=(80, 25))
        self.ClearBtn.Bind(wx.EVT_BUTTON, self.OnClearConsoleContent)

        # 控制台
        self.ConsoleContent = wx.TextCtrl(self, pos=(10, 70), size=(620, 355), style=wx.TE_MULTILINE | wx.TE_READONLY)

    # 打开文件
    def OnOpen(self, event):
        wildcard = 'Microsoft Excel 97/2000/XP/2003 Workbook(*.xls)|*.xls|Microsoft Excel 2007/2010 Workbook(*.xlsx)|*.xlsx'
        dialog = wx.FileDialog(None, "请选择要处理的Excel文件", os.getcwd(), '', wildcard)

        if dialog.ShowModal() == wx.ID_OK:
            self.FileName.SetValue(dialog.GetPath())
            dialog.Destroy

    # 解析文件
    def OnReadFile(self, event):
        fileName = self.FileName.GetValue()
        if not os.path.exists(fileName):
            wx.MessageBox("请先选择文件", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        file = open(self.FileName.GetValue())
        self.ConsoleContent.SetValue(file.read())
        file.close()

    def OnProcess(self, event):
        wx.MessageBox("共处理24个文件", "处理结果", wx.OK | wx.ICON_INFORMATION)
        return

    def OnClearConsoleContent(self, event):
        self.ConsoleContent.SetValue("")
        wx.MessageBox("已清空", "处理结果", wx.OK | wx.ICON_INFORMATION)

    # 退出菜单
    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)

    # 生成菜单
    def MakeMenuBar(self):
        # Make the menu bar and add the two menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menuBar = wx.MenuBar()
        menuBar.Append(self.MakeFileMenu(), "&文件")
        menuBar.Append(self.MakeTemplateMenu(), "&模板")
        menuBar.Append(self.MakeHelpMenu(), "&帮助")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

    # 文件菜单
    def MakeFileMenu(self):
        fileMenu = wx.Menu()

        # The "\t..." syntax defines an accelerator key that also triggers
        # the same event
        openItem = fileMenu.Append(-1, "&打开\tCtrl-O", "打开要处理的Excel文件")

        # 分隔符
        fileMenu.AppendSeparator()

        exitItem = fileMenu.Append(-1, "&退出\tCtrl-Q", "退出")

        # Associate a handler function with the EVT_MENU event for each of the menu items.
        self.Bind(wx.EVT_MENU, self.OnOpen, openItem)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)

        return fileMenu

    # 模板菜单
    def MakeTemplateMenu(self):
        templateMenu = wx.Menu()

        settingItem = templateMenu.Append(-1, "&设置\tCtrl-,", "模板参数设置")
        self.Bind(wx.EVT_MENU, self.OnSettingTemplate, settingItem)

        exportItem = templateMenu.Append(-1, "&导出\tCtrl-E", "导出模板文件")
        self.Bind(wx.EVT_MENU, self.OnExportTemplate, exportItem)

        return templateMenu

    # 模板设置
    def OnSettingTemplate(self, event):
        pass

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
        sheet.write(0, 0, "图号")
        sheet.write(1, 0, "1-1")

        workbook.save(fileName)

    # 使用说明
    def OnUsage(self, event):
        message = "1.点击[打开]按钮，选择要处理的清单文件\n" \
                  + "2.点击[执行]按钮，处理清单文件"
        wx.MessageBox(message, "使用说明", wx.OK | wx.ICON_INFORMATION)

    # 关于菜单
    def OnAbout(self, event):
        """Display an About Dialog"""
        wx.MessageBox("这是一个简单的选图工具", self.GetVersion(), wx.OK | wx.ICON_INFORMATION)

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
            print("源文件夹不存在")
            return

        if not os.path.isdir(target):
            print("目标文件夹不存在")
            return

        sourceName = os.path.join(source, fileName)
        targetName = os.path.join(target, fileName)
        if not os.path.exists(sourceName):
            print("源文件不存在：" + sourceName)
            return

        try:
            shutil.copyfile(sourceName, targetName)
            print('成功复制：' + targetName)
        except Exception:
            print("复制失败")


if __name__ == '__main__':
    app = wx.App()

    frame = ChoseFile()
    frame.Show()

    app.MainLoop()
