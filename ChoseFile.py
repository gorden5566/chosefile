import os
import wx


class ChoseFile(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title='选图工具', size=(640, 480))

        self.MakePanel()

        # create a menu bar
        self.MakeMenuBar()

        self.CreateStatusBar()
        self.SetStatusText("欢迎使用" + self.GetVersion())

    def MakePanel(self):
        # 选择文件按钮
        self.OnOpenBtn = wx.Button(self, label='打开', pos=(10, 10), size=(80, 25))
        self.OnOpenBtn.Bind(wx.EVT_BUTTON, self.OnOpen)

        # 已选择的文件
        self.FileName = wx.TextCtrl(self, pos=(105, 10), size=(400, 25), style=wx.TE_READONLY)

        # 读取文件
        self.SelBtn = wx.Button(self, label='执行', pos=(10, 40), size=(80, 25))
        self.SelBtn.Bind(wx.EVT_BUTTON, self.OnReadFile)

        # 控制台
        self.FileContent = wx.TextCtrl(self, pos=(10, 70), size=(620, 355), style=wx.TE_MULTILINE | wx.TE_READONLY)

    # 打开文件
    def OnOpen(self, event):
        wildcard = 'Microsoft Excel 97/2000/XP/2003 Workbook(*.xls)|*.xls|Microsoft Excel 2007/2010 Workbook(*.xlsx)|*.xlsx'
        dialog = wx.FileDialog(None, "请选择要处理的Excel文件", os.getcwd(), '', wildcard)

        if dialog.ShowModal() == wx.ID_OK:
            self.FileName.SetValue(dialog.GetPath())
            dialog.Destroy

    def OnReadFile(self, event):
        file = open(self.FileName.GetValue())
        self.FileContent.SetValue(file.read())
        file.close()

    # 退出菜单
    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)

    # 生成菜单
    def MakeMenuBar(self):
        # Make a file menu with Hello and Exit items
        fileMenu = wx.Menu()
        # The "\t..." syntax defines an accelerator key that also triggers
        # the same event
        openItem = fileMenu.Append(-1, "&打开\tCtrl-O", "打开要处理的Excel文件")
        fileMenu.AppendSeparator()

        # When using a stock ID we don't need to specify the menu item's
        # label
        # exitItem = fileMenu.Append(wx.ID_EXIT)
        exitItem = fileMenu.Append(-1, "&退出\tCtrl-Q", "退出")

        # Now a help menu for the about item
        helpMenu = wx.Menu()
        # aboutItem = helpMenu.Append(wx.ID_ABOUT)
        aboutItem = helpMenu.Append(-1, "&关于", "关于")

        # Make the menu bar and add the two menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&文件")
        menuBar.Append(helpMenu, "&帮助")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

        # Associate a handler function with the EVT_MENU event for each of the menu items.
        self.Bind(wx.EVT_MENU, self.OnOpen, openItem)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    # 关于菜单
    def OnAbout(self, event):
        """Display an About Dialog"""
        wx.MessageBox("这是一个简单的选图工具",
                      self.GetVersion(),
                      wx.OK | wx.ICON_INFORMATION)

    # 版本号
    def GetVersion(self):
        return "选图工具 V0.0.1"


if __name__ == '__main__':
    app = wx.App()

    SiteFrame = ChoseFile()
    SiteFrame.Show()

    app.MainLoop()
