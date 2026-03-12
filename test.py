#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/3/12 下午2:03
# @Site    : 
# @File    : test.py
# @Software: PyCharm

import subprocess
import json
import os

class MyOCR:
    def init(self, exe_path):
        # exe_path 是你电脑上 PaddleOCR-json.exe 的路径
        self.path = exe_path
        self.process = subprocess.Popen(
        [exe_path, "--use_gpu=0", "--ensure_ascii=false"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=True,
        encoding="utf-8"
        )
        # 读取掉引擎启动时的欢迎语
        self.process.stdout.readline()