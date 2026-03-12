#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/2/6 下午8:18
# @Site    :
# @File    : app.py
# @Software: PyCharm
# ! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2025/4/3 上午11:40
# @Site    :
# @File    : pywinautoWDT.py
# @Software: PyCharm
import os
import sys
import traceback
import psutil
from pywinauto import Application


class autoLoginTB():

    def check_process(self, exename):
        target_processes = [p for p in psutil.process_iter(['pid', 'name', 'create_time'])
                            if exename in p.info['name']]
        sorted_processes = sorted(target_processes, key=lambda p: p.create_time())
        main_process = sorted_processes if sorted_processes else None
        return [item.pid for item in main_process] if main_process else []

    def start_app(self):
        pid_list = self.check_process("Weixin.exe")
        for id in pid_list:
            print("pid:",id)
            try:
                app = Application(backend="uia").connect(process=id)
                win = app.window(class_name='Qt51514QWindowIcon')
                # 检测窗口是否可访问（不改变焦点）
                if win.exists() and win.is_visible():
                    return win
            except  Exception as e:
                print(f"进程 {id} 检测超时或失败，跳过...")
                traceback.print_exc()
                continue
        # os.system("taskkill /f /t /im  WeChatAppEx.exe")
        return None


if __name__ == '__main__':
    print(autoLoginTB().start_app())
