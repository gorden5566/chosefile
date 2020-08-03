# -*- coding: utf-8 -*-
# /usr/bin/python3

import os

import wx

from index import IndexTool
from logger import Logger
from parser import Parser
from processor import Processor
from setting import Setting


class ChoseFile(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title='ChoseFile', size=(640, 505),
                         style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN)

        # ui
        self.init_ui()

        # logger
        self.logger = Logger(self.console_text)

        # config
        self.setting = Setting(self.logger)

        # parser
        self.parser = Parser(self.logger)

        # processor
        self.processor = Processor(self.logger)

        # index
        self.index_tool = IndexTool(self.setting.get_max_depth())

        self.init_default()

    def init_ui(self):
        # panel
        self.make_panel()

        # create a menu bar
        self.make_menu_bar()

        # status bar
        self.make_status_bar()

    def init_default(self):
        excel_path = self.setting.get_excel_path()
        if excel_path is not None:
            self.file_name_text.SetValue(excel_path)

        target_dir = self.setting.get_target_dir()
        if target_dir is not None:
            self.target_dir_text.SetValue(target_dir)

    def make_panel(self):
        # 选择文件按钮
        self.select_btn = wx.Button(self, label='选择清单', pos=(10, 10), size=(80, 25))
        self.select_btn.Bind(wx.EVT_BUTTON, self.on_select)

        # 已选择的文件
        self.file_name_text = wx.TextCtrl(self, pos=(105, 10), size=(400, 25), style=wx.TE_READONLY)

        # 目标地址
        self.target_btn = wx.Button(self, label='目标地址', pos=(10, 40), size=(80, 25))
        self.target_btn.Bind(wx.EVT_BUTTON, self.on_target)

        # 已选择的目标地址
        self.target_dir_text = wx.TextCtrl(self, pos=(105, 40), size=(400, 25), style=wx.TE_READONLY)

        # 处理文件
        self.process_btn = wx.Button(self, label='批量复制', pos=(10, 70), size=(80, 25))
        self.process_btn.Bind(wx.EVT_BUTTON, self.on_process)

        # 清空控制台日志
        self.clear_btn = wx.Button(self, label='清空日志', pos=(105, 70), size=(80, 25))
        self.clear_btn.Bind(wx.EVT_BUTTON, self.on_clear_console_content)

        # 控制台
        self.console_text = wx.TextCtrl(self, pos=(10, 100), size=(605, 345), style=wx.TE_MULTILINE | wx.TE_READONLY)

    # 打开文件
    def on_select(self, event):
        wildcard = 'Microsoft Excel 97/2000/XP/2003 Workbook(*.xls)|*.xls|Microsoft Excel 2007/2010 Workbook(*.xlsx)|*.xlsx'
        dialog = wx.FileDialog(None, "请选择要处理的Excel文件", os.getcwd(), '', wildcard)

        if dialog.ShowModal() == wx.ID_OK:
            self.file_name_text.SetValue(dialog.GetPath())
            dialog.Destroy

    # 打开文件
    def on_target(self, event):
        dialog = wx.DirDialog(self, message="请选择要保存的路径", defaultPath=self.setting.get_target_dir(),
                              style=wx.DD_DEFAULT_STYLE)
        if dialog.ShowModal() == wx.ID_OK:
            self.target_dir.SetValue(dialog.GetPath())
        dialog.Destroy()

    def on_process(self, event):
        # 检查索引是否存在，若不存在则构建
        has_build_db = self.index_tool.check_db()
        if not has_build_db:
            result = self.build_index()
            if not result:
                return

        file_name = self.file_name_text.GetValue()
        if file_name is None or file_name == '':
            wx.MessageBox("请先选择清单文件", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        if not os.path.exists(file_name):
            wx.MessageBox("文件不存在: " + file_name, "处理结果", wx.OK | wx.ICON_WARNING)
            return

        target_path = self.target_dir_text.GetValue()
        if target_path is None or target_path == '':
            wx.MessageBox("请先选择目标地址", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        self.logger.Log("[目标文件夹]\t" + target_path)

        name_arr = self.parser.parse_excel(file_name, self.setting.get_column_title())
        if name_arr is None:
            wx.MessageBox("解析结果为空", "处理结果", wx.OK | wx.ICON_WARNING)
            return

        self.do_copy_file(target_path, name_arr)

        return

    def do_copy_file(self, target_path, name_arr):
        # 复制文件
        total = 0
        success_num = 0
        failed_files = []
        for name in name_arr:
            source_name = name + self.setting.get_ext_name();
            source_path = self.get_source_path(source_name)
            total += 1

            if source_path is None:
                self.logger.Log("[索引结果空]\t" + source_name)
                failed_files.append(source_name)
                continue

            success = self.processor.copy_file(source_path, target_path, source_name)
            if success:
                success_num += 1
            else:
                failed_files.append(source_name)

        message = "共处理" + str(total) + "个文件，处理成功" + str(success_num) + "个"
        self.logger.Log("[结果汇总]\t" + message)
        failed_file_name = self.logger.log_failed(failed_files)
        if failed_file_name is not None:
            self.logger.Log("[失败记录]\t" + failed_file_name)
        self.logger.Log("--------------------------------------------------------------------")

        wx.MessageBox(message, "处理结果", wx.OK | wx.ICON_INFORMATION)

    # 查询文件所在目录
    def get_source_path(self, source_name):
        # 默认位置为 self.setting.getsourcedir()
        # 因为文件可能在子文件夹中，所以还需考虑递归遍历所有子文件夹
        # 为加快查询速度，如下为从索引中查询对应结果
        index = self.index_tool.find(source_name)
        return self.index_tool.get_full_path(index)

    def on_clear_console_content(self, event):
        self.console_text.SetValue("")
        # wx.MessageBox("已清空", "处理结果", wx.OK | wx.ICON_INFORMATION)

    # 退出菜单
    def on_exit(self, event):
        """Close the frame, terminating the application."""
        self.logger.close()
        self.Close(True)

    # 生成菜单
    def make_menu_bar(self):
        # Make the menu bar and add the two menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menu_bar = wx.MenuBar()

        menu_bar.Append(self.make_file_menu(), "&文件")
        menu_bar.Append(self.make_setting_menu(), "&设置")
        menu_bar.Append(self.make_help_menu(), "&帮助")

        # Give the menu bar to the frame
        self.SetMenuBar(menu_bar)

    # 文件菜单
    def make_file_menu(self):
        file_menu = wx.Menu()

        # The "\t..." syntax defines an accelerator key that also triggers
        # the same event
        open_item = file_menu.Append(-1, "&选择清单\tCtrl-O", "选择Excel清单文件")
        self.Bind(wx.EVT_MENU, self.on_select, open_item)

        export_item = file_menu.Append(-1, "&导出模板\tCtrl-E", "导出模板文件")
        self.Bind(wx.EVT_MENU, self.on_export_template, export_item)

        # 分隔符
        file_menu.AppendSeparator()

        exit_item = file_menu.Append(-1, "&退出\tCtrl-Q", "退出")
        self.Bind(wx.EVT_MENU, self.on_exit, exit_item)

        return file_menu

    # 模板菜单
    def make_setting_menu(self):
        setting_menu = wx.Menu()

        build_index_item = setting_menu.Append(-1, "&重建索引\tCtrl-B", "修改源文件路径[sourceDir]后，需要重新构建文件索引")
        self.Bind(wx.EVT_MENU, self.on_build_index, build_index_item)

        return setting_menu

    # 重建索引设置
    def on_build_index(self, event):
        result = self.build_index()
        if result:
            source_dir = self.setting.get_source_dir()
            message = "重建索引成功[" + source_dir + "]"
            wx.MessageBox(message, "提示", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox("重建索引失败", "提示", wx.OK | wx.ICON_WARNING)

    # 帮助菜单
    def make_help_menu(self):
        help_menu = wx.Menu()

        usage_item = help_menu.Append(-1, "&说明\tCtrl-H", "查看使用说明")
        self.Bind(wx.EVT_MENU, self.on_usage, usage_item)

        about_item = help_menu.Append(-1, "&关于", "关于")
        self.Bind(wx.EVT_MENU, self.on_about, about_item)

        return help_menu

    # 导出模板文件
    def on_export_template(self, event):
        fd = wx.FileDialog(self, message='导出模板文件', defaultDir='', defaultFile='图号清单模板',
                           wildcard='Microsoft Excel 97/2000/XP/2003 Workbook(*.xls)|*.xls|Microsoft Excel 2007/2010 Workbook(*.xlsx)|*.xlsx',
                           style=wx.FD_SAVE)
        if fd.ShowModal() == wx.ID_OK:
            try:
                file_name = fd.GetFilename()
                dir_name = fd.GetDirectory()
                column_title = self.setting.get_column_title()
                self.processor.export_template(os.path.join(dir_name, file_name), column_title)
                save_msg = wx.MessageDialog(self, '保存成功！', '提示')
            except FileNotFoundError:
                save_msg = wx.MessageDialog(self, '保存失败，无效的保存路径', '提示')

            save_msg.ShowModal()
            save_msg.Destroy()

    # 使用说明
    def on_usage(self, event):
        message = "1.打开[config.ini]，配置[sourceDir]等参数\n" \
                  + "2.点击[选择]按钮，选择要处理的清单文件\n" \
                  + "3.点击[处理]按钮，选择要保存的路径，确认后执行\n" \
                  + "4.执行完后，查看日志，确认是否执行成功\n" \
                  + "注意：若修改了[sourceDir]参数，请点击[设置]->[重建索引]菜单重新构建文件索引\n"
        wx.MessageBox(message, "使用说明", wx.OK | wx.ICON_INFORMATION)

    # 关于菜单
    def on_about(self, event):
        """Display an About Dialog"""
        wx.MessageBox("批量复制文件", self.get_version(), wx.OK | wx.ICON_INFORMATION)

    # 状态栏
    def make_status_bar(self):
        self.CreateStatusBar()
        self.SetStatusText("欢迎使用 " + self.get_version())

    # 版本号
    def get_version(self):
        return "ChoseFile V0.0.1"

    def build_index(self):
        source_dir = self.setting.get_source_dir()
        if source_dir is None:
            wx.MessageBox("源文件夹未设置，请先打开[config.ini]设置[sourceDir]", "提示", wx.OK | wx.ICON_WARNING)
            return False
        if not os.path.isdir(source_dir):
            wx.MessageBox("源文件夹不存在，请先打开[config.ini]设置[sourceDir]", "提示", wx.OK | wx.ICON_WARNING)
            return False
        self.index_tool.build_index(source_dir).save()
        return True


if __name__ == '__main__':
    app = wx.App()

    frame = ChoseFile()
    frame.Show()

    app.MainLoop()
