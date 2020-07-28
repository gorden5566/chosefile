# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import pickle


class IndexTool:
    def __init__(self):
        self.index = None
        pass

    def find(self, name):
        if self.index is None:
            self.index = self.load()
        return self.dofind(self.index, name)

    def dofind(self, index, name):
        # 文件
        if not index.isdir:
            if index.equals(name):
                return index
            else:
                return None

        # 文件夹
        next = index.next
        if next is None:
            return None

        for i in next:
            result = self.dofind(i, name)
            if result is not None:
                return result

        return None

    def buildindex(self, parentpath, path, name, isdir):
        index = Index(parentpath, path, name, isdir)
        if not isdir:
            return index

        files = os.listdir(parentpath)
        for file in files:
            filePath = os.path.join(parentpath, file)

            if os.path.isdir(filePath):
                if file[0] != '.':
                    index.addnext(self.buildindex(path, filePath, file, True))
            elif os.path.isfile(filePath):
                if file[0] != '.':
                    index.addnext(self.buildindex(path, filePath, file, False))

        return index

    def save(self, index):
        f = open('index.db', 'wb')
        pickle.dump(index, f)
        f.close()

    def load(self):
        f = open('index.db', 'rb')
        index = pickle.load(f)
        f.close()
        return index


class Index:
    def __init__(self, parentpath, path, name, isdir):
        self.parentpath = parentpath
        self.path = path
        self.name = name
        self.isdir = isdir
        self.next = None

    def getparentpath(self):
        return self.parentpath

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
