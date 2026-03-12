#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/3/12
# @File    : agent.py
# @Software: PyCharm

import sys
import base64
import time
import logging
import psutil
from typing import Any, Dict, Optional, Tuple
from io import BytesIO
from pathlib import Path

import pyautogui
pyautogui.FAILSAFE = False
import pygetwindow as gw
from PIL import Image, ImageGrab
import pyperclip
from pywinauto import Application
import win32gui
import win32ui
import win32con
import numpy as np

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("WeChat_Tool")


class WeChatManager:
    """微信客户端管理器"""
    
    WECHAT_EXE = "Weixin.exe"
    WAKE_UP_HOTKEY = ('ctrl', 'alt', 'w')

    def __init__(self):
        self._app = None
        self._window = None
        self._gw_window = None  # pygetwindow 窗口
    
    def check_process(self) -> list:
        """检查微信进程，返回 PID 列表"""
        try:
            processes = [p for p in psutil.process_iter(['pid', 'name'])
                        if self.WECHAT_EXE in p.info['name']]
            return [p.info['pid'] for p in processes]
        except Exception as e:
            logger.error(f"检查进程失败: {e}")
            return []
    
    def _find_gw_window(self) -> Optional[Any]:
        """使用 pygetwindow 查找微信窗口"""
        for keyword in ["微信", "WeChat", "Weixin"]:
            windows = gw.getWindowsWithTitle(keyword)
            for w in windows:
                if 500 < w.width < 2000:
                    return w
        return None
    
    def _wake_up_wechat(self) -> bool:
        """唤醒微信窗口（Ctrl+Alt+W）"""
        logger.debug("执行唤醒快捷键...")
        pyautogui.hotkey(*self.WAKE_UP_HOTKEY)
        time.sleep(0.5)
        
        # 检查窗口是否出现
        w = self._find_gw_window()
        return w is not None
    
    def _bring_window_to_front(self, w) -> bool:
        """将窗口置顶并激活"""
        try:
            # 方法1: pygetwindow 激活
            try:
                w.activate()
                time.sleep(0.3)
            except:
                pass
            
            # 方法2: 使用 win32gui 强制置顶
            try:
                hwnd = win32gui.FindWindow(None, w.title)
                if hwnd:
                    # 如果窗口最小化，先恢复
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    # 置顶窗口
                    win32gui.SetForegroundWindow(hwnd)
                    # 确保窗口在最前
                    win32gui.BringWindowToTop(hwnd)
                    time.sleep(0.3)
            except:
                pass
            
            return True
        except Exception as e:
            logger.debug(f"激活窗口失败: {e}")
            return False
    
    def get_main_window(self, force_refresh: bool = False, activate_first: bool = False) -> Optional[Any]:
        """获取微信主窗口（返回 pygetwindow 窗口对象）"""
        # 1. 检查微信是否运行
        pid_list = self.check_process()
        if not pid_list:
            logger.debug("微信未运行")
            return None
        
        # 2. 如果需要激活窗口，强制刷新
        if activate_first:
            force_refresh = True
        
        # 3. 使用缓存的窗口
        if self._gw_window and not force_refresh:
            return self._gw_window
        
        # 4. 查找窗口
        w = self._find_gw_window()
        
        # 5. 如果没找到或需要激活，尝试唤醒
        if not w or activate_first:
            if not self._wake_up_wechat():
                logger.warning("唤醒微信失败")
                return None
            w = self._find_gw_window()
        
        if not w:
            logger.warning("未找到微信窗口")
            return None
        
        # 6. 激活窗口（确保在前台）
        self._bring_window_to_front(w)
        
        logger.debug(f"找到窗口: {w.title}, 大小: {w.width}x{w.height}")
        
        # 7. 尝试连接 pywinauto（用于 OCR）
        for pid in pid_list:
            try:
                self._app = Application(backend="uia").connect(process=pid, timeout=2)
                self._window = self._app.window(title="微信")
                break
            except:
                continue
        
        self._gw_window = w
        return w
    
    def get_window_rect(self) -> Optional[Dict]:
        """获取窗口位置和大小（使用 pygetwindow）"""
        w = self._gw_window or self.get_main_window()
        if not w:
            return None
        
        try:
            return {
                "left": w.left,
                "top": w.top,
                "right": w.right,
                "bottom": w.bottom,
                "width": w.width,
                "height": w.height
            }
        except Exception as e:
            logger.error(f"获取窗口位置失败: {e}")
            return None
    
    def capture(self) -> Optional[Image.Image]:
        """截取微信窗口（先激活窗口，确保获取最新内容）"""
        # 先激活窗口
        w = self.get_main_window(activate_first=True)
        if not w:
            return None
        
        try:
            time.sleep(0.3)  # 等待窗口激活完成
            
            # 优先使用 ImageGrab 截取屏幕区域（保证最新）
            bbox = (w.left, w.top, w.right, w.bottom)
            return ImageGrab.grab(bbox=bbox)
            
        except Exception as e:
            logger.error(f"截图失败: {e}")
            return None
    
    def _capture_window_hwnd(self, hwnd) -> Optional[Image.Image]:
        """使用 win32 API 截取窗口（即使被遮挡）"""
        try:
            # 获取窗口尺寸
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            
            # 获取窗口 DC
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            # 创建位图
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            
            # 截取窗口内容
            result = saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
            
            # 转换为 PIL Image
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )
            
            # 清理资源
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            
            return img
        except Exception as e:
            logger.error(f"win32 截图失败: {e}")
            return None
    
    def get_status(self) -> Dict:
        """获取微信状态"""
        if not self.check_process():
            return {"status": "not_running", "message": "微信未运行"}
        
        w = self.get_main_window()
        if not w:
            return {"status": "not_running", "message": "未找到微信窗口"}
        
        rect = self.get_window_rect()
        
        return {
            "status": "running",
            "title": w.title,
            "position": {"x": rect["left"], "y": rect["top"]} if rect else None,
            "size": {"width": rect["width"], "height": rect["height"]} if rect else None
        }
    
    def click(self, x: int, y: int) -> bool:
        """点击指定坐标（先激活窗口）"""
        try:
            # 先激活窗口
            w = self.get_main_window(activate_first=True)
            if not w:
                return False
            time.sleep(0.1)
            
            pyautogui.click(int(x), int(y))
            return True
        except Exception as e:
            logger.error(f"点击失败: {e}")
            return False
    
    def input_text(self, content: str, x: int = None, y: int = None, 
                   send_enter: bool = False) -> bool:
        """输入文本（先激活窗口）"""
        try:
            # 先激活窗口
            w = self.get_main_window(activate_first=True)
            if not w:
                return False
            time.sleep(0.1)
            
            if x is not None and y is not None:
                pyautogui.click(int(x), int(y))
                time.sleep(0.2)
            
            pyperclip.copy(content)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            
            if send_enter:
                pyautogui.press('enter')
            
            return True
        except Exception as e:
            logger.error(f"输入失败: {e}")
            return False
    
    def scroll(self, direction: str = 'down', amount: int = 300, 
               x: int = None, y: int = None) -> bool:
        """滚动页面（先激活窗口）
        
        Args:
            direction: 滚动方向，"up" 或 "down"
            amount: 滚动量（像素）
            x: 滚动位置 X 坐标，默认窗口中心
            y: 滚动位置 Y 坐标，默认窗口中心
        """
        try:
            # 先激活窗口
            w = self.get_main_window(activate_first=True)
            if not w:
                return False
            time.sleep(0.1)
            
            # 如果未指定坐标，使用窗口中心
            if x is None or y is None:
                rect = self.get_window_rect()
                if rect:
                    x = rect["left"] + rect["width"] // 2
                    y = rect["top"] + rect["height"] // 2
                else:
                    x, y = w.left + w.width // 2, w.top + w.height // 2
            
            # 移动鼠标到指定位置
            pyautogui.moveTo(int(x), int(y))
            time.sleep(0.1)
            
            scroll_amount = -amount if direction == 'down' else amount
            pyautogui.scroll(scroll_amount, int(x), int(y))
            return True
        except Exception as e:
            logger.error(f"滚动失败: {e}")
            return False
    
    def send_message(self, message: str) -> Tuple[bool, Optional[str]]:
        """发送消息"""
        logger.info(f"发送消息: {message}")
        
        # 1. 确保窗口可见
        w = self.get_main_window(activate_first=True)
        if not w:
            return False, "未找到微信窗口"
        
        rect = self.get_window_rect()
        if not rect:
            return False, "无法获取窗口位置"
        
        logger.debug(f"窗口位置: ({rect['left']}, {rect['top']}), 大小: ({rect['width']}, {rect['height']})")
        
        # 2. 点击输入框
        input_x = rect["left"] + rect["width"] // 2
        input_y = rect["bottom"] - 60
        self.click(input_x, input_y)
        time.sleep(0.2)
        
        # 3. 输入消息
        pyperclip.copy(message)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        
        # 4. 发送
        pyautogui.press('enter')
        time.sleep(0.3)
        
        logger.info("消息发送完成")
        return True, None
    
    def get_ocr_result(self, word: str = None) -> Dict:
        """获取 OCR 结果（每次获取最新截图）
        
        Args:
            word: 搜索关键词（可选）
        
        Returns:
            Dict: OCR 结果
        """
        try:
            from OCR import ocr_endpoint
            
            # 获取 pygetwindow 窗口
            w = self.get_main_window(activate_first=True)
            if not w:
                return {"status": "error", "message": "无法获取微信窗口"}
            
            # 调用 OCR（始终获取最新截图）
            results = ocr_endpoint(w, word=word)
            return {"status": "success", "data": {"results": results, "count": len(results)}}
            
        except ImportError as e:
            return {"status": "error", "message": f"OCR 模块导入失败: {e}"}
        except Exception as e:
            return {"status": "error", "message": f"OCR 识别失败: {str(e)}"}
    
    def get_page_context(self) -> Dict:
        """获取页面上下文（OCR，每次获取最新截图）"""
        return self.get_ocr_result()
    
    def take_screenshot(self, save_path: str = None) -> Dict:
        """截图"""
        img = self.capture()
        if not img:
            return {"status": "error", "message": "无法获取微信窗口"}
        
        if save_path:
            img.save(save_path)
            return {"status": "success", "message": f"已保存: {save_path}"}
        else:
            buffered = BytesIO()
            img.save(buffered, format="PNG")
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            return {"status": "success", "data": {"image_base64": img_base64}}


# 全局管理器实例
_manager = WeChatManager()


# 兼容旧接口的函数
def get_wechat_main_window():
    """获取微信主窗口"""
    return _manager.get_main_window()

def get_wechat_status():
    """获取微信状态"""
    return _manager.get_status()

def send_message_to_current(message):
    """发送消息"""
    return _manager.send_message(message)

def screenshot(save_path=None):
    """截图"""
    return _manager.take_screenshot(save_path)

def get_ocr_result(word=None):
    """获取 OCR 结果（每次获取最新截图）"""
    return _manager.get_ocr_result(word=word)

def click_coordinate(x, y):
    """点击坐标"""
    if not x or not y:
        return {"status": "error", "message": "缺少 x 或 y 参数"}
    success = _manager.click(x, y)
    return {"status": "success" if success else "error", "message": f"点击 ({x}, {y})"}

def click_and_type(content, x=None, y=None, send_enter=False):
    """点击并输入"""
    if not content:
        return {"status": "error", "message": "缺少 content 参数"}
    success = _manager.input_text(content, x, y, send_enter)
    return {"status": "success" if success else "error", "message": "输入成功" if success else "输入失败"}

def scroll(direction='down', amount=300, x=None, y=None):
    """滚动"""
    success = _manager.scroll(direction, amount, x, y)
    return {"status": "success" if success else "error", "message": "滚动成功" if success else "滚动失败"}

def get_page_context():
    """获取页面上下文（每次获取最新截图）"""
    return _manager.get_page_context()


if __name__ == "__main__":
    import argparse
    import json
    
    parser = argparse.ArgumentParser(description="微信自动化工具")
    parser.add_argument("action", help="执行动作",
                        choices=["screenshot", "get_ocr_result", "click_coordinate",
                                 "click_and_type", "scroll", "get_page_context",
                                 "get_wechat_status", "send_message"])
    parser.add_argument("params", nargs="?", default="{}", help="JSON 格式参数")
    args = parser.parse_args()
    
    try:
        params = json.loads(args.params) if args.params else {}
    except json.JSONDecodeError:
        print(json.dumps({"status": "error", "message": "参数 JSON 格式错误"}))
        sys.exit(1)
    
    result = {"status": "error", "message": "未知操作"}
    
    if args.action == "screenshot":
        result = screenshot(params.get("save_path"))
    elif args.action == "get_ocr_result":
        result = get_ocr_result()
    elif args.action == "click_coordinate":
        result = click_coordinate(params.get("x"), params.get("y"))
    elif args.action == "click_and_type":
        result = click_and_type(
            params.get("content"),
            params.get("x"),
            params.get("y"),
            params.get("send_enter", False)
        )
    elif args.action == "scroll":
        result = scroll(
            params.get("direction", "down"), 
            params.get("amount", 300),
            params.get("x"),
            params.get("y")
        )
    elif args.action == "get_page_context":
        result = get_page_context()
    elif args.action == "get_wechat_status":
        result = get_wechat_status()
    elif args.action == "send_message":
        success, err = send_message_to_current(params.get("message"))
        if success:
            result = {"status": "success", "message": "消息发送成功"}
        else:
            result = {"status": "error", "message": err}
    
    print(json.dumps(result, ensure_ascii=False, indent=2))
