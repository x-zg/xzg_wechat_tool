#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/3/12
# @File    : agent.py
# @Software: PyCharm

import sys
import io
import os

# 强制设置 stdout 编码为 UTF-8（解决 Windows 控制台编码问题）
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

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
import win32api
import win32process
import numpy as np

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("WeChat_Tool")


class WeChatManager:
    """微信客户端管理器"""
    
    WECHAT_EXE = "Weixin.exe"
    WAKE_UP_HOTKEY = ('ctrl', 'alt', 'w')
    # 微信安装路径（常见路径，按优先级排序）
    WECHAT_PATHS = [
        os.path.expandvars(r"%LOCALAPPDATA%\Tencent\WeChat\WeChat.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Tencent\WeChat\WeChat.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Tencent\WeChat\WeChat.exe"),
        r"D:\Tencent\WeChat\WeChat.exe",
        r"C:\Program Files\Tencent\WeChat\WeChat.exe",
        r"C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
    ]

    def __init__(self):
        self._app = None
        self._window = None
        self._gw_window = None  # pygetwindow 窗口
    
    def _find_wechat_path(self) -> Optional[str]:
        """查找微信安装路径"""
        for path in self.WECHAT_PATHS:
            if os.path.exists(path):
                return path
        return None
    
    def start_wechat(self) -> Tuple[bool, str]:
        """启动微信
        
        Returns:
            Tuple[bool, str]: (是否成功, 消息)
        """
        # 先检查是否已运行
        if self.check_process():
            return True, "微信已在运行"
        
        # 查找微信路径
        wechat_path = self._find_wechat_path()
        if not wechat_path:
            return False, "未找到微信安装路径"
        
        try:
            # 启动微信
            import subprocess
            subprocess.Popen([wechat_path], shell=True)
            logger.info(f"正在启动微信: {wechat_path}")
            
            # 等待微信启动（最多等待 10 秒）
            for i in range(20):
                time.sleep(0.5)
                if self.check_process():
                    # 再等待窗口出现
                    time.sleep(1)
                    return True, "微信启动成功"
            
            return False, "微信启动超时"
        except Exception as e:
            logger.error(f"启动微信失败: {e}")
            return False, f"启动微信失败: {e}"
    
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
    
    def _is_window_visible(self, w) -> bool:
        """检查窗口是否可见且不在最小化状态"""
        try:
            # 获取窗口句柄
            hwnd = w._hWnd if hasattr(w, '_hWnd') else None
            if hwnd:
                # 使用 Win32 API 检查窗口状态
                if not win32gui.IsWindowVisible(hwnd):
                    return False
                # 检查是否最小化
                if win32gui.IsIconic(hwnd):
                    return False
                return True
            
            # 回退到属性检查
            if w.width < 100 or w.height < 100:
                return False
            if w.left < -w.width or w.top < -w.height:
                return False
            return True
        except Exception:
            return False
    
    def _wake_up_wechat(self) -> bool:
        """唤醒微信窗口（尝试多种方法）"""
        w = self._find_gw_window()
        if not w:
            logger.debug("未找到微信窗口")
            return False
        
        # 如果窗口已可见，直接返回成功
        if self._is_window_visible(w):
            logger.debug("窗口已可见，跳过唤醒")
            return True
        
        # 获取窗口句柄
        hwnd = w._hWnd if hasattr(w, '_hWnd') else None
        if hwnd:
            logger.debug(f"尝试通过 Win32 API 激活窗口: hwnd={hwnd}")
            # 如果窗口最小化，先恢复
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)
            
            # 尝试置顶
            try:
                # 使用 AttachThreadInput 绕过限制
                foreground_hwnd = win32gui.GetForegroundWindow()
                foreground_thread = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]
                current_thread = win32api.GetCurrentThreadId()
                target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
                
                if current_thread != target_thread:
                    win32gui.AttachThreadInput(current_thread, target_thread, True)
                    if foreground_thread != target_thread:
                        win32gui.AttachThreadInput(foreground_thread, target_thread, True)
                
                win32gui.SetForegroundWindow(hwnd)
                win32gui.BringWindowToTop(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                
                if current_thread != target_thread:
                    win32gui.AttachThreadInput(current_thread, target_thread, False)
                    if foreground_thread != target_thread:
                        win32gui.AttachThreadInput(foreground_thread, target_thread, False)
                
                time.sleep(0.2)
                
                if self._is_window_visible(w):
                    logger.debug("Win32 API 激活窗口成功")
                    return True
            except Exception as e:
                logger.debug(f"Win32 API 激活失败: {e}")
        
        # 回退：使用快捷键
        logger.debug("尝试使用快捷键唤醒...")
        pyautogui.hotkey(*self.WAKE_UP_HOTKEY)
        time.sleep(0.5)
        
        return self._is_window_visible(w)
    
    def _bring_window_to_front(self, w) -> bool:
        """将窗口置顶并激活"""
        try:
            # 获取窗口句柄（优先使用 pygetwindow 的 _hWnd 属性）
            hwnd = None
            if hasattr(w, '_hWnd'):
                hwnd = w._hWnd
            else:
                # 回退：尝试通过标题查找
                hwnd = win32gui.FindWindow(None, w.title)
            
            if not hwnd:
                logger.warning("无法获取窗口句柄")
                return False
            
            # 方法1: 如果窗口最小化，先恢复
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.1)
            
            # 方法2: 使用 AttachThreadInput 绕过 SetForegroundWindow 限制
            foreground_hwnd = win32gui.GetForegroundWindow()
            foreground_thread = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]
            current_thread = win32api.GetCurrentThreadId()
            target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
            
            # 附加线程输入
            if current_thread != target_thread:
                win32gui.AttachThreadInput(current_thread, target_thread, True)
                if foreground_thread != target_thread:
                    win32gui.AttachThreadInput(foreground_thread, target_thread, True)
            
            # 置顶窗口
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            
            # 解除线程附加
            if current_thread != target_thread:
                win32gui.AttachThreadInput(current_thread, target_thread, False)
                if foreground_thread != target_thread:
                    win32gui.AttachThreadInput(foreground_thread, target_thread, False)
            
            time.sleep(0.2)
            logger.debug(f"窗口已激活: hwnd={hwnd}")
            return True
            
        except Exception as e:
            logger.error(f"激活窗口失败: {e}")
            # 最后尝试 pygetwindow 方法
            try:
                w.activate()
                time.sleep(0.3)
                return True
            except Exception:
                return False
    
    def get_main_window(self, force_refresh: bool = False, activate_first: bool = False, 
                         auto_start: bool = True) -> Optional[Any]:
        """获取微信主窗口（返回 pygetwindow 窗口对象）
        
        Args:
            force_refresh: 是否强制刷新窗口
            activate_first: 是否先激活窗口
            auto_start: 微信未运行时是否自动启动
        """
        # 1. 检查微信是否运行，未运行时尝试启动
        pid_list = self.check_process()
        if not pid_list:
            if auto_start:
                logger.info("微信未运行，尝试自动启动...")
                success, msg = self.start_wechat()
                if not success:
                    logger.warning(f"自动启动微信失败: {msg}")
                    return None
                pid_list = self.check_process()
                if not pid_list:
                    return None
            else:
                logger.debug("微信未运行")
                return None
        
        # 2. 如果需要激活窗口，强制刷新
        if activate_first:
            force_refresh = True
        
        # 3. 验证缓存窗口是否仍然有效
        if self._gw_window and not force_refresh:
            try:
                # 验证窗口是否仍然存在且有效
                _ = self._gw_window.left  # 尝试访问属性
                if 500 < self._gw_window.width < 2000:
                    return self._gw_window
            except Exception:
                # 窗口已失效，清除缓存
                self._gw_window = None
                logger.debug("缓存窗口已失效，重新查找")
        
        # 4. 查找窗口
        w = self._find_gw_window()
        
        # 5. 如果窗口不可见或未找到，尝试唤醒
        if not w or not self._is_window_visible(w):
            if not self._wake_up_wechat():
                logger.warning("唤醒微信失败")
                return None
            w = self._find_gw_window()
        
        if not w:
            logger.warning("未找到微信窗口")
            return None
        
        # 6. 激活窗口（确保在前台）
        if activate_first:
            self._bring_window_to_front(w)
        
        logger.debug(f"找到窗口: {w.title}, 大小: {w.width}x{w.height}")
        
        # 7. 尝试连接 pywinauto（用于 OCR）
        for pid in pid_list:
            try:
                self._app = Application(backend="uia").connect(process=pid, timeout=2)
                self._window = self._app.window(title="微信")
                break
            except Exception:
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
        """截取微信窗口（使用 win32 API，支持自动启动）"""
        # 先激活窗口（强制刷新，自动启动微信）
        w = self.get_main_window(force_refresh=True, activate_first=True, auto_start=True)
        if not w:
            logger.error("无法获取微信窗口")
            return None
        
        # 获取窗口句柄
        hwnd = w._hWnd if hasattr(w, '_hWnd') else None
        if not hwnd:
            # 回退：通过标题查找
            try:
                hwnd = win32gui.FindWindow(None, w.title)
            except Exception:
                pass
        
        if not hwnd:
            logger.error("无法获取微信窗口句柄")
            return None
        
        # 验证窗口是否已激活到前台，最多尝试 3 次
        for attempt in range(3):
            foreground_hwnd = win32gui.GetForegroundWindow()
            if foreground_hwnd == hwnd:
                break  # 微信已在前台
            
            logger.debug(f"微信窗口未在前台 (尝试 {attempt + 1}/3)，重新激活...")
            self._bring_window_to_front(w)
            time.sleep(0.3)
        else:
            # 3 次尝试后仍未激活，但窗口句柄有效，仍可尝试用 Win32 API 截图
            logger.warning("微信窗口未能激活到前台，尝试直接截图")
        
        try:
            # 获取窗口位置
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            
            # 验证窗口尺寸有效
            width = right - left
            height = bottom - top
            if width < 100 or height < 100:
                logger.error(f"微信窗口尺寸异常: {width}x{height}")
                return None
            
            # 优先使用 win32 API 截图（即使窗口被遮挡也能正确截图）
            img = self._capture_window_hwnd(hwnd)
            if img:
                return img
            
            # 回退方案：使用 ImageGrab 截取屏幕区域
            logger.debug("win32 截图失败，回退到 ImageGrab")
            bbox = (left, top, right, bottom)
            return ImageGrab.grab(bbox=bbox)
            
        except Exception as e:
            logger.error(f"截图失败: {e}")
            return None
    
    def _capture_window_hwnd(self, hwnd) -> Optional[Image.Image]:
        """使用 win32 API 截取窗口（即使被遮挡）"""
        hwndDC = None
        mfcDC = None
        saveDC = None
        saveBitMap = None
        
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
            saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
            
            # 转换为 PIL Image
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )
            
            return img
        except Exception as e:
            logger.error(f"win32 截图失败: {e}")
            return None
        finally:
            # 确保清理资源（即使发生异常）
            try:
                if saveBitMap:
                    win32gui.DeleteObject(saveBitMap.GetHandle())
            except Exception:
                pass
            try:
                if saveDC:
                    saveDC.DeleteDC()
            except Exception:
                pass
            try:
                if mfcDC:
                    mfcDC.DeleteDC()
            except Exception:
                pass
            try:
                if hwndDC:
                    win32gui.ReleaseDC(hwnd, hwndDC)
            except Exception:
                pass
    
    def get_status(self) -> Dict:
        """获取微信状态（确保窗口激活，支持自动启动）"""
        # 直接调用 get_main_window，它会自动处理启动逻辑
        w = self.get_main_window(activate_first=True, auto_start=True)
        if not w:
            # 检查是否是微信未运行
            if not self.check_process():
                return {"status": "not_running", "message": "微信未运行且自动启动失败"}
            return {"status": "not_running", "message": "未找到微信窗口"}
        
        rect = self.get_window_rect()
        
        return {
            "status": "running",
            "title": w.title,
            "position": {"x": rect["left"], "y": rect["top"]} if rect else None,
            "size": {"width": rect["width"], "height": rect["height"]} if rect else None
        }
    
    def click(self, x: int, y: int) -> bool:
        """点击指定坐标（先激活窗口，支持自动启动）"""
        try:
            # 先激活窗口（自动启动微信）
            w = self.get_main_window(activate_first=True, auto_start=True)
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
        """输入文本（先激活窗口，支持自动启动）"""
        try:
            # 先激活窗口（自动启动微信）
            w = self.get_main_window(activate_first=True, auto_start=True)
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
        """滚动页面（先激活窗口，支持自动启动）
        
        Args:
            direction: 滚动方向，"up" 或 "down"
            amount: 滚动量（像素）
            x: 滚动位置 X 坐标，默认窗口中心
            y: 滚动位置 Y 坐标，默认窗口中心
        """
        try:
            # 先激活窗口（自动启动微信）
            w = self.get_main_window(activate_first=True, auto_start=True)
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
        """发送消息（支持自动启动微信）"""
        logger.info(f"发送消息: {message}")
        
        # 1. 确保窗口可见（自动启动微信）
        w = self.get_main_window(activate_first=True, auto_start=True)
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
            
            # 获取 pygetwindow 窗口（强制刷新，自动启动微信）
            w = self.get_main_window(force_refresh=True, activate_first=True, auto_start=True)
            if not w:
                return {"status": "error", "message": "无法获取微信窗口"}
            
            # 验证窗口有效性
            try:
                _ = w.left  # 测试窗口是否有效
            except Exception:
                # 窗口失效，重新获取
                logger.debug("OCR: 窗口失效，重新获取")
                self._gw_window = None
                w = self.get_main_window(force_refresh=True, activate_first=True, auto_start=True)
                if not w:
                    return {"status": "error", "message": "无法获取微信窗口"}
            
            # 调用 OCR（始终获取最新截图）
            results = ocr_endpoint(w, word=word)
            return {"status": "success", "data": {"results": results, "count": len(results)}}
            
        except ImportError as e:
            return {"status": "error", "message": f"OCR 模块导入失败: {e}"}
        except Exception as e:
            logger.error(f"OCR 识别失败: {e}")
            return {"status": "error", "message": f"OCR 识别失败: {str(e)}"}
    
    def get_page_context(self) -> Dict:
        """获取页面上下文（OCR，每次获取最新截图）"""
        return self.get_ocr_result()
    
    # ==================== 控件获取功能（已禁用，速度太慢）====================
    # def get_controls(self, control_type: str = None, name_filter: str = None,
    #                  max_depth: int = 10, include_invisible: bool = False) -> Dict:
    #     """获取微信窗口的所有控件信息"""
    #     pass
    # 
    # def _traverse_controls(self, element, controls: list, depth: int, max_depth: int,
    #                       control_type: str, name_filter: str, include_invisible: bool):
    #     """递归遍历控件树"""
    #     pass
    # 
    # def _get_control_info(self, element) -> Optional[Dict]:
    #     """获取单个控件的信息"""
    #     pass
    # ==========================================================================
    
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
    
    # 使用子命令方式，每个动作有独立参数
    subparsers = parser.add_subparsers(dest="action", help="执行动作")
    
    # screenshot
    p_screenshot = subparsers.add_parser("screenshot", help="截图")
    p_screenshot.add_argument("--save_path", type=str, default=None, help="保存路径")
    
    # get_wechat_status
    subparsers.add_parser("get_wechat_status", help="获取微信状态")
    
    # get_ocr_result
    subparsers.add_parser("get_ocr_result", help="OCR识别")
    
    # click_coordinate
    p_click = subparsers.add_parser("click_coordinate", help="点击坐标")
    p_click.add_argument("--x", type=int, required=True, help="X坐标")
    p_click.add_argument("--y", type=int, required=True, help="Y坐标")
    
    # click_and_type
    p_type = subparsers.add_parser("click_and_type", help="输入文字")
    p_type.add_argument("--content", type=str, required=True, help="文字内容")
    p_type.add_argument("--x", type=int, default=None, help="点击X坐标")
    p_type.add_argument("--y", type=int, default=None, help="点击Y坐标")
    p_type.add_argument("--send_enter", action="store_true", help="输入后按回车")
    
    # scroll
    p_scroll = subparsers.add_parser("scroll", help="滚动")
    p_scroll.add_argument("--direction", type=str, default="down", choices=["up", "down"], help="方向")
    p_scroll.add_argument("--amount", type=int, default=300, help="滚动量")
    p_scroll.add_argument("--x", type=int, default=None, help="X坐标")
    p_scroll.add_argument("--y", type=int, default=None, help="Y坐标")
    
    # get_page_context
    subparsers.add_parser("get_page_context", help="获取页面上下文")
    
    # send_message
    p_send = subparsers.add_parser("send_message", help="发送消息")
    p_send.add_argument("--message", type=str, required=True, help="消息内容")
    
    args = parser.parse_args()
    
    if args.action is None:
        parser.print_help()
        sys.exit(1)
    
    result = {"status": "error", "message": "未知操作"}
    
    if args.action == "screenshot":
        result = screenshot(args.save_path)
    elif args.action == "get_ocr_result":
        result = get_ocr_result()
    elif args.action == "click_coordinate":
        result = click_coordinate(args.x, args.y)
    elif args.action == "click_and_type":
        result = click_and_type(args.content, args.x, args.y, args.send_enter)
    elif args.action == "scroll":
        result = scroll(args.direction, args.amount, args.x, args.y)
    elif args.action == "get_page_context":
        result = get_page_context()
    elif args.action == "get_wechat_status":
        result = get_wechat_status()
    elif args.action == "send_message":
        success, err = send_message_to_current(args.message)
        if success:
            result = {"status": "success", "message": "消息发送成功"}
        else:
            result = {"status": "error", "message": err}
    
    print(json.dumps(result, ensure_ascii=False, indent=2))
