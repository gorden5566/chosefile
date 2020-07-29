# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import pickle


class IndexTool:
    def __init__(self, maxdepth):
        self.maxdepth = maxdepth
        self.indexfile = None
        self.dbname = "index.db"
        pass

    # 基准路径
    def getbasepath(self):
        return self.indexfile.getbasepath()

    # 组装路径
    def getfullpath(self, index):
        if index is None:
            return None
        return os.path.normpath(os.path.join(self.getbasepath(), index.getpath()))

    # 检查索引文件是否存在
    def checkdb(self):
        if os.path.exists(self.dbname):
            return True
        return False

    # 查找文件
    def find(self, name):
        if self.indexfile is None:
            self.load()
        index = self.indexfile.getindex()
        return self.dofind(index, name)

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
    def buildindex(self, path):
        self.indexfile = IndexFile(path)
        index = self.dobuildindex(None, ".", "", True, 0)
        self.indexfile.setindex(index)
        return self

    def dobuildindex(self, parent, path, name, isdir, depth):
        # 最大深度
        if depth > self.maxdepth:
            return None
        index = Index(parent, path, name, isdir)
        if not isdir:
            return index

        parentpath = ""
        if not parent is None:
            parentpath = parent.getpath()
        pathwithname = os.path.join(self.indexfile.getbasepath(), parentpath, path, name)

        files = os.listdir(pathwithname)
        for file in files:
            if file[0] == '.':
                continue

            filePath = os.path.join(pathwithname, file)
            if os.path.isdir(filePath):
                index.addnext(self.dobuildindex(index, name, file, True, depth + 1))
            elif os.path.isfile(filePath):
                index.addnext(self.dobuildindex(index, name, file, False, depth + 1))

        return index

    # 保存索引文件
    def save(self):
        f = open(self.dbname, 'wb')
        pickle.dump(self.indexfile, f)
        f.close()
        return self

    # 加载索引文件
    def load(self):
        f = open(self.dbname, 'rb')
        self.indexfile = pickle.load(f)
        f.close()
        return self


class IndexFile:
    def __init__(self, basepath):
        self.basepath = basepath
        self.index = None

    def getbasepath(self):
        return self.basepath

    def getindex(self):
        return self.index

    def setindex(self, index):
        self.index = index


class Index:
    def __init__(self, parent, path, name, isdir):
        self.parent = parent
        # 所在目录
        self.path = path
        # 文件夹/文件名
        self.name = name
        # 是否为目录
        self.isdir = isdir
        # 下一级列表
        self.next = None

    def getpath(self):
        if self.parent is None:
            return self.path
        return os.path.join(self.parent.getpath(), self.path)

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
