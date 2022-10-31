# -*- coding: utf-8 -*-
# /usr/bin/python3

import os
import pickle


class IndexTool:
    def __init__(self, max_depth, logger):
        self.max_depth = max_depth
        self.logger = logger
        self.index_file = None
        self.db_name = "index.db"
        pass

    # 基准路径
    def get_base_path(self):
        return self.index_file.get_base_path()

    # 组装路径
    def get_full_path(self, index):
        if index is None:
            return None
        return os.path.normpath(os.path.join(self.get_base_path(), index.get_path()))

    # 检查索引文件是否存在
    def check_db(self):
        if not os.path.exists(self.db_name):
            return False

        try:
            self.load()
            return True
        except Exception:
            return False

    # 查找文件
    def find(self, name):
        if self.index_file is None:
            self.load()
        index = self.index_file.get_index()
        return self.do_find(index, name)

    # 从索引中查找文件
    def do_find(self, index, name):
        if index is None:
            return None

        # 文件
        if not index.get_isdir():
            if index.equals(name):
                return index
            else:
                return None

        # 文件夹
        next_level = index.get_next()
        if next_level is None:
            return None

        for i in next_level:
            result = self.do_find(i, name)
            if result is not None:
                return result

        return None

    # 打印索引
    def print_index(self):
        if self.index_file is None:
            self.load()
        index = self.index_file.get_index()

        self.logger.Log("--------------------------------------------------------------------")
        depth = 0
        self.do_print_index(index, depth)
        self.logger.Log("--------------------------------------------------------------------")
        return

    def do_print_index(self, index, depth):
        if index is None:
            return

        pre_str = ""
        for num in range(1, depth):
            pre_str = pre_str + "\t"
        
        # 文件
        if not index.get_isdir():
            self.logger.raw_log(pre_str + "├─" + index.get_name() + "\n")
            return

        # 文件夹
        if index.get_name() == "":
            self.logger.raw_log(".\n")
        else:
            self.logger.raw_log(pre_str + "├─" + index.get_name() + "\n")
        next_level = index.get_next()
        if next_level is None:
            return None

        for i in next_level:
            self.do_print_index(i, depth + 1)

        return

    # 构建索引
    def build_index(self, path):
        self.index_file = IndexFile(path)
        index = self.do_build_index(None, ".", "", True, 0)
        self.index_file.set_index(index)
        return self

    def do_build_index(self, parent, path, name, isdir, depth):
        # 最大深度
        if depth > self.max_depth:
            return None
        index = Index(parent, path, name, isdir)
        if not isdir:
            return index

        parent_path = ""
        if parent is not None:
            parent_path = parent.get_path()
        path_with_name = os.path.join(self.index_file.get_base_path(), parent_path, path, name)

        files = os.listdir(path_with_name)
        files.sort()
        for file in files:
            if file[0] == '.':
                continue

            file_path = os.path.join(path_with_name, file)
            if os.path.isdir(file_path):
                index.add_next(self.do_build_index(index, name, file, True, depth + 1))
            elif os.path.isfile(file_path):
                index.add_next(self.do_build_index(index, name, file, False, depth + 1))

        return index

    # 保存索引文件
    def save(self):
        f = open(self.db_name, 'wb')
        pickle.dump(self.index_file, f)
        f.close()
        return self

    # 加载索引文件
    def load(self):
        f = open(self.db_name, 'rb')
        self.index_file = pickle.load(f)
        f.close()
        return self


class IndexFile:
    def __init__(self, base_path):
        self.base_path = base_path
        self.index = None

    def get_base_path(self):
        return self.base_path

    def get_index(self):
        return self.index

    def set_index(self, index):
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

    def get_path(self):
        if self.parent is None:
            return self.path
        return os.path.join(self.parent.get_path(), self.path)

    def get_name(self):
        return self.name

    def get_isdir(self):
        return self.isdir

    def get_next(self):
        return self.next

    def add_next(self, next):
        if self.next is None:
            self.next = []
        self.next.append(next)

    def equals(self, name):
        if name == self.name:
            return True
        return False
