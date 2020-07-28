# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import pickle


class IndexTool:
    def __init__(self):
        self.index = None
        self.dbname = "index.db"
        pass

    # 检查索引文件是否存在
    def checkdb(self):
        if os.path.exists(self.dbname):
            return True
        return False

    # 查找文件
    def find(self, name):
        if self.index is None:
            self.index = self.load()
        return self.dofind(self.index, name)

    # 从索引中查找文件
    def dofind(self, index, name):
        # 文件
        if not index.getisdir():
            if index.equals(name):
                return index
            else:
                return None

        # 文件夹
        next = index.getnext()
        if next is None:
            return None

        for i in next:
            result = self.dofind(i, name)
            if result is not None:
                return result

        return None

    # 构建索引
    def buildindex(self, path, name, isdir):
        index = Index(path, name, isdir)
        if not isdir:
            return index

        pathwithname = os.path.join(path, name)
        files = os.listdir(pathwithname)
        for file in files:
            filePath = os.path.join(pathwithname, file)

            if os.path.isdir(filePath):
                if file[0] != '.':
                    index.addnext(self.buildindex(pathwithname, file, True))
            elif os.path.isfile(filePath):
                if file[0] != '.':
                    index.addnext(self.buildindex(pathwithname, file, False))

        return index

    # 保存索引文件
    def save(self, index):
        f = open(self.dbname, 'wb')
        pickle.dump(index, f)
        f.close()

    # 加载索引文件
    def load(self):
        f = open(self.dbname, 'rb')
        index = pickle.load(f)
        f.close()
        return index


class Index:
    def __init__(self, path, name, isdir):
        # 所在目录
        self.path = path
        # 文件夹/文件名
        self.name = name
        # 是否为目录
        self.isdir = isdir
        # 下一级列表
        self.next = None

    def getpath(self):
        return self.path

    def getname(self):
        return self.name

    def getisdir(self):
        return self.isdir

    def getnext(self):
        return self.next

    def addnext(self, next):
        if self.next is None:
            self.next = []
        self.next.append(next)

    def equals(self, name):
        if name == self.name:
            return True
        return False
