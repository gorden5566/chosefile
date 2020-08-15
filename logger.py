# -*- coding: utf-8 -*-
# /usr/bin/python3

import time
import os


class Logger:
    def __init__(self, console_content):
        self.console_content = console_content
        self.file_logger = None

    @staticmethod
    def create_file_logger():
        today = time.strftime("%Y%m%d", time.localtime())
        f = open(today + ".log", "a+")
        return f

    def close(self):
        if self.file_logger is not None:
            self.file_logger.close()

    def Log(self, message):
        if self.file_logger is None:
            self.file_logger = self.create_file_logger()

        now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        format_message = now + "\t" + message + "\n"

        self.console_content.AppendText(format_message)

        self.file_logger.write(format_message)

    def log_failed(self, file_path, failed_files):
        if not failed_files:
            return None

        now = time.strftime("%Y%m%d%H%M%S", time.localtime())
        failed_name = os.path.join(file_path, "failed" + now + ".txt")
        f = open(failed_name, "a+")
        for file in failed_files:
            f.write(file + "\n")

        f.close()

        return failed_name

