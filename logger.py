# -*- coding: utf-8 -*-
# /usr/bin/python3

import time


class Logger:
    def __init__(self, consolecontent):
        self.console_content = consolecontent
        self.file_logger = self.create_file_logger()

    @staticmethod
    def create_file_logger():
        today = time.strftime("%Y%m%d", time.localtime())
        f = open(today + ".log", "a+")
        return f

    def close(self):
        self.file_logger.close()

    def Log(self, message):
        now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        format_message = now + "\t" + message + "\n"

        self.console_content.AppendText(format_message)

        self.file_logger.write(format_message)
