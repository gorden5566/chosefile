# -*- coding: utf-8 -*-
# /usr/bin/python3

import time


class Logger:
    def __init__(self, consolecontent):
        self.consolecontent = consolecontent
        self.createfilelogger()

    def createfilelogger(self):
        today = time.strftime("%Y%m%d", time.localtime())
        f = open(today + ".log", "a+")
        self.filelogger = f

    def close(self):
        self.filelogger.close()

    def Log(self, message):
        now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        formatmessage = now + "\t" + message + "\n"

        self.consolecontent.AppendText(formatmessage)

        self.filelogger.write(formatmessage)
