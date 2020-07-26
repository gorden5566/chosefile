import os

import wx


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
        file = open(self.FileName.GetValue())
        self.ConsoleContent.SetValue(file.read())
        file.close()

    def OnProcess(self, event):
        return

    def OnClearConsoleContent(self, event):
        self.ConsoleContent.SetValue("")

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
        menuBar.Append(self.MakeHelpMenu(), "&帮助")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

    # 文件菜单
    def MakeFileMenu(self):
        fileMenu = wx.Menu()

        # The "\t..." syntax defines an accelerator key that also triggers
        # the same event
        openItem = fileMenu.Append(-1, "&打开\tCtrl-O", "打开要处理的Excel文件")
        fileMenu.AppendSeparator()

        exitItem = fileMenu.Append(-1, "&退出\tCtrl-Q", "退出")

        # Associate a handler function with the EVT_MENU event for each of the menu items.
        self.Bind(wx.EVT_MENU, self.OnOpen, openItem)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)

        return fileMenu

    # 帮助菜单
    def MakeHelpMenu(self):
        helpMenu = wx.Menu()
        # aboutItem = helpMenu.Append(wx.ID_ABOUT)
        aboutItem = helpMenu.Append(-1, "&关于", "关于")
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

        usageItem = helpMenu.Append(-1, "&说明\tCtrl-H", "查看使用说明")
        self.Bind(wx.EVT_MENU, self.OnUsage, usageItem)

        return helpMenu

    # 使用说明
    def OnUsage(self, event):
        message = "1.点击[打开]按钮，选择要处理的清单文件\n"\
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


if __name__ == '__main__':
    app = wx.App()

    frame = ChoseFile()
    frame.Show()

    app.MainLoop()
