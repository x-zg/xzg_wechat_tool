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
import time
from pywinauto import Application


class autoLoginTB():

    def check_process(self, exename):
        target_processes = [p for p in psutil.process_iter(['pid', 'name', 'create_time'])
                            if exename in p.info['name']]
        sorted_processes = sorted(target_processes, key=lambda p: p.create_time())
        main_process = sorted_processes if sorted_processes else None
        return [item.pid for item in main_process] if main_process else []

    def start_app(self):
        """启动并连接微信窗口 - 简化版"""
        pid_list = self.check_process("Weixin.exe")
        if not pid_list:
            print("未找到微信进程")
            return None
        
        print(f"找到 {len(pid_list)} 个微信进程：{pid_list}")
        
        # 直接连接第一个进程
        pid = pid_list[0]
        try:
            app = Application(backend="uia").connect(process=pid)
            print(f"已连接到进程 {pid}")
            
            # 尝试窗口
            win = app.window(class_name_re='mmui::.*')
            print(f"  找到窗口")
            return win
            
        except Exception as e:
            print(f"连接失败：{e}")
            return None


if __name__ == '__main__':
    print(autoLoginTB().start_app())
