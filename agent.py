#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/3/12
# @File    : agent.py
# @Software: PyCharm

import sys
import io
import os

# ============== 强制统一编码为 UTF-8（解决 Windows 控制台编码问题）==============
if sys.platform == 'win32':
    # 强制 stdout/stderr 为 UTF-8
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    # 设置控制台代码页为 UTF-8
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass

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

# ============== 配置日志输出编码为 UTF-8 ==============
# 创建 StreamHandler，确保使用 UTF-8 编码
_log_handler = logging.StreamHandler(sys.stdout)
_log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# 配置根日志器
logging.basicConfig(
    level=logging.INFO,
    handlers=[_log_handler]
)
logger = logging.getLogger("WeChat_Tool")


class WeChatManager:
    """微信客户端管理器"""
    
    # 支持多种进程名（不同版本微信进程名可能不同）
    WECHAT_EXE_NAMES = ["WeChat.exe", "Weixin.exe"]
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
                        if any(exe.lower() in p.info['name'].lower() for exe in self.WECHAT_EXE_NAMES)]
            return [p.info['pid'] for p in processes]
        except Exception as e:
            logger.error(f"检查进程失败: {e}")
            return []
    
    def _find_gw_window(self) -> Optional[Any]:
        """使用 pygetwindow 查找微信窗口"""
        for keyword in ["微信", "WeChat", "Weixin"]:
            windows = gw.getWindowsWithTitle(keyword)
            for w in windows:
                # 放宽窗口尺寸限制：宽度 > 200 且高度 > 200
                if w.width > 200 and w.height > 200:
                    logger.debug(f"找到微信窗口: title={w.title}, size=({w.width}x{w.height})")
                    return w
                else:
                    logger.debug(f"跳过窗口(尺寸过小): title={w.title}, size=({w.width}x{w.height})")
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
        """唤醒微信窗口
        
        逻辑：
        1. 如果窗口不存在（微信在托盘区）→ 用快捷键唤醒
        2. 如果窗口存在但最小化 → 用 Win32 API 恢复
        3. 如果窗口存在且已显示 → 直接激活到前台（不用快捷键，避免关闭）
        """
        w = self._find_gw_window()
        
        # 情况1：找不到窗口，微信可能在托盘区，用快捷键唤醒
        if not w:
            logger.debug("未找到微信窗口，尝试快捷键唤醒...")
            pyautogui.hotkey(*self.WAKE_UP_HOTKEY)
            time.sleep(0.5)
            w = self._find_gw_window()
            return w is not None and self._is_window_visible(w)
        
        # 获取窗口句柄
        hwnd = w._hWnd if hasattr(w, '_hWnd') else None
        if not hwnd:
            logger.warning("无法获取窗口句柄")
            return False
        
        # 情况2：窗口最小化，需要恢复
        if win32gui.IsIconic(hwnd):
            logger.debug("窗口已最小化，恢复窗口...")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)
        
        # 情况3：窗口已显示，激活到前台（不用快捷键）
        logger.debug("激活窗口到前台...")
        self._bring_window_to_front(w)
        time.sleep(0.2)
        
        return self._is_window_visible(w)
    
    def _ensure_window_ready(self, auto_start: bool = True) -> Tuple[Optional[Any], Optional[str]]:
        """确保微信窗口就绪（统一的状态检查方法）
        
        所有操作方法应调用此方法来确保窗口状态正确。
        
        Args:
            auto_start: 是否自动启动微信（如果未运行）
        
        Returns:
            Tuple[Optional[window], Optional[error_msg]]:
                - 成功: (window_object, None)
                - 失败: (None, error_message)
        """
        logger.info("开始检查微信窗口状态...")
        
        # 步骤1: 检查微信进程是否存在
        pid_list = self.check_process()
        logger.info(f"微信进程检测结果: {pid_list}")
        
        if not pid_list:
            # 进程不存在
            if auto_start:
                logger.info("微信未运行，尝试自动启动...")
                success, msg = self.start_wechat()
                if not success:
                    return None, f"微信未运行且自动启动失败: {msg}"
                pid_list = self.check_process()
                if not pid_list:
                    return None, "微信启动超时"
            else:
                return None, "微信未运行"
        
        # 步骤2: 查找微信窗口（最多尝试 3 次）
        w = None
        for attempt in range(3):
            w = self._find_gw_window()
            if w:
                break
            logger.debug(f"未找到窗口 (尝试 {attempt + 1}/3)")
            time.sleep(0.3)
        
        # 步骤3: 判断窗口状态并处理
        if not w:
            # 窗口不存在（微信在托盘区），用快捷键唤醒（最多尝试 3 次）
            logger.info("窗口不存在，尝试快捷键唤醒...")
            for attempt in range(3):
                pyautogui.hotkey(*self.WAKE_UP_HOTKEY)
                time.sleep(0.8)
                w = self._find_gw_window()
                if w:
                    logger.info(f"快捷键唤醒成功 (尝试 {attempt + 1}/3)")
                    break
                logger.debug(f"快捷键唤醒未找到窗口 (尝试 {attempt + 1}/3)")
            
            if not w:
                return None, "无法唤醒微信窗口（可能在托盘区或快捷键无效）"
        
        # 步骤4: 获取窗口句柄
        hwnd = w._hWnd if hasattr(w, '_hWnd') else None
        if not hwnd:
            return None, "无法获取微信窗口句柄"
        
        logger.info(f"找到微信窗口: hwnd={hwnd}, title={w.title}")
        
        # 步骤5: 检查窗口是否最小化
        if win32gui.IsIconic(hwnd):
            logger.info("窗口已最小化，恢复窗口...")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)
        
        # 步骤6: 激活窗口到前台（最多尝试 3 次）
        for attempt in range(3):
            self._bring_window_to_front(w)
            time.sleep(0.3)
            
            # 步骤7: 验证窗口是否真的在前台
            foreground_hwnd = win32gui.GetForegroundWindow()
            if foreground_hwnd == hwnd:
                logger.info("窗口已成功激活到前台")
                break
            
            logger.debug(f"窗口未在前台 (尝试 {attempt + 1}/3)，hwnd={hwnd}, foreground={foreground_hwnd}")
        else:
            logger.warning("窗口激活可能失败，但继续操作")
        
        self._gw_window = w
        return w, None
    
    def _bring_window_to_front(self, w) -> bool:
        """将窗口置顶并激活（使用多种方法确保成功）"""
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
            
            logger.info(f"  正在激活窗口 hwnd={hwnd}...")
            
            # ===== 方法1: 先恢复窗口（如果最小化）=====
            if win32gui.IsIconic(hwnd):
                logger.info("  窗口已最小化，正在恢复...")
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
            
            # ===== 方法2: 使用 AttachThreadInput 绕过 SetForegroundWindow 限制 =====
            foreground_hwnd = win32gui.GetForegroundWindow()
            foreground_thread = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]
            current_thread = win32api.GetCurrentThreadId()
            target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
            
            logger.info(f"  线程信息: 当前={current_thread}, 前台={foreground_thread}, 目标={target_thread}")
            
            # 附加线程输入
            attach_threads = []
            if current_thread != target_thread:
                win32gui.AttachThreadInput(current_thread, target_thread, True)
                attach_threads.append((current_thread, target_thread))
            if foreground_thread != target_thread and foreground_thread != current_thread:
                win32gui.AttachThreadInput(foreground_thread, target_thread, True)
                attach_threads.append((foreground_thread, target_thread))
            
            # ===== 方法3: 设置窗口位置和状态 =====
            # 先显示窗口
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            # 设置为顶层窗口（不改变大小）
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
            # 取消顶层（避免一直置顶）
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
            
            # ===== 方法4: 激活窗口 =====
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetFocus(hwnd)
            
            # 解除线程附加
            for t1, t2 in attach_threads:
                try:
                    win32gui.AttachThreadInput(t1, t2, False)
                except Exception:
                    pass
            
            time.sleep(0.3)
            
            # ===== 方法5: 如果前面方法失败，使用模拟按键绕过限制 =====
            new_foreground = win32gui.GetForegroundWindow()
            if new_foreground != hwnd:
                logger.info("  常规方法激活失败，尝试模拟 Alt 键绕过限制...")
                # 模拟按下 Alt 键（这会允许 SetForegroundWindow 工作）
                win32api.keybd_event(0x12, 0, 0, 0)  # Alt down
                time.sleep(0.1)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.1)
                win32api.keybd_event(0x12, 0, 2, 0)  # Alt up
                time.sleep(0.2)
            
            # 验证激活结果
            new_foreground = win32gui.GetForegroundWindow()
            if new_foreground == hwnd:
                logger.info(f"  ✓ 窗口激活成功!")
                return True
            else:
                # ===== 方法6: 最后尝试 pygetwindow =====
                logger.warning(f"  Win32 激活失败，尝试 pygetwindow...")
                try:
                    w.activate()
                    time.sleep(0.3)
                    new_foreground = win32gui.GetForegroundWindow()
                    if new_foreground == hwnd:
                        logger.info(f"  ✓ pygetwindow 激活成功!")
                        return True
                except Exception as e:
                    logger.warning(f"  pygetwindow 激活也失败: {e}")
                
                logger.warning(f"  ⚠ 窗口激活可能失败，当前前台窗口 hwnd={new_foreground}")
                return True  # 仍然返回 True，让调用者决定是否重试
            
        except Exception as e:
            logger.error(f"激活窗口失败: {e}")
            # 最后尝试 pygetwindow 方法
            try:
                w.activate()
                time.sleep(0.3)
                logger.info("  使用 pygetwindow.activate() 成功")
                return True
            except Exception as e2:
                logger.error(f"pygetwindow 激活也失败: {e2}")
                return False
    
    def get_main_window(self, force_refresh: bool = False, activate_first: bool = False, 
                         auto_start: bool = True) -> Optional[Any]:
        """获取微信主窗口（返回 pygetwindow 窗口对象）
        
        Args:
            force_refresh: 是否强制刷新窗口
            activate_first: 是否先激活窗口
            auto_start: 微信未运行时是否自动启动
        """
        # 使用统一的窗口状态检查方法
        w, error = self._ensure_window_ready(auto_start=auto_start)
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
    
    # ==================== 聊天列表监控功能 ====================
    
    def get_chat_list(self, count: int = 5) -> Dict:
        """获取左侧聊天列表的前N个联系人
        
        通过 OCR 识别左侧聊天列表区域，提取联系人信息。
        
        Args:
            count: 获取的联系人数量（默认5个）
        
        Returns:
            Dict: {
                "status": "success/error",
                "data": {
                    "contacts": [
                        {
                            "index": 0,
                            "name": "联系人名称",
                            "last_message": "最后一条消息",
                            "position": {"x": 100, "y": 150},  # 点击位置
                            "rect": {"left": 0, "top": 120, "right": 250, "bottom": 170}
                        },
                        ...
                    ],
                    "total": 5
                }
            }
        """
        try:
            from OCR import ocr_endpoint
            
            # 确保窗口就绪
            w, error = self._ensure_window_ready()
            if error:
                return {"status": "error", "message": error}
            
            # 获取窗口位置
            rect = self.get_window_rect()
            if not rect:
                return {"status": "error", "message": "无法获取窗口位置"}
            
            # 左侧聊天列表区域（大约左侧 280 像素宽度）
            chat_list_width = 280
            chat_list_left = rect["left"]
            chat_list_top = rect["top"] + 60  # 跳过搜索框
            chat_list_right = rect["left"] + chat_list_width
            chat_list_bottom = rect["bottom"]
            
            # 截取聊天列表区域
            img = self.capture()
            if not img:
                return {"status": "error", "message": "截图失败"}
            
            # 裁剪左侧聊天列表区域
            chat_list_img = img.crop((chat_list_left, chat_list_top, chat_list_right, chat_list_bottom))
            
            # 保存临时截图用于 OCR
            temp_path = os.path.join(os.path.dirname(__file__), "_temp_chat_list.png")
            chat_list_img.save(temp_path)
            
            # OCR 识别
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
            result, _ = ocr(temp_path)
            
            # 删除临时文件
            try:
                os.remove(temp_path)
            except:
                pass
            
            if not result:
                return {"status": "error", "message": "OCR 未识别到内容"}
            
            # 解析 OCR 结果，提取联系人信息
            # result 格式: [[box, text, confidence], ...]
            contacts = []
            current_y = 0
            contact_height = 65  # 每个联系人项大约 65 像素高度
            
            # 按位置分组（同一行的文字属于同一个联系人）
            lines = []
            for item in result:
                box, text, conf = item
                # box 是四个点的坐标 [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
                y_pos = (box[0][1] + box[2][1]) / 2  # 取中心 Y 坐标
                x_pos = box[0][0]
                
                # 过滤无效文本（太短的、特殊字符等）
                if len(text.strip()) < 1:
                    continue
                
                lines.append({
                    "text": text,
                    "y": y_pos,
                    "x": x_pos,
                    "box": box
                })
            
            # 按 Y 坐标排序并分组
            lines.sort(key=lambda l: l["y"])
            
            # 分组：相邻的行（Y 差距小于 30）属于同一个联系人
            groups = []
            current_group = []
            last_y = -100
            
            for line in lines:
                if abs(line["y"] - last_y) < 30:
                    current_group.append(line)
                else:
                    if current_group:
                        groups.append(current_group)
                    current_group = [line]
                last_y = line["y"]
            
            if current_group:
                groups.append(current_group)
            
            # 解析每个分组为联系人信息
            for i, group in enumerate(groups[:count]):
                if not group:
                    continue
                
                # 联系人名称通常是该组中最左边且靠上的文字
                name_text = None
                message_text = None
                name_y = float('inf')
                
                for line in group:
                    text = line["text"].strip()
                    # 名称通常在左侧，消息在右侧或下方
                    if line["x"] < 80 and line["y"] < name_y:
                        # 可能是名称
                        if not name_text or line["y"] < name_y:
                            name_text = text
                            name_y = line["y"]
                    elif len(text) > 2:
                        message_text = text
                
                # 计算点击位置（转换回屏幕坐标）
                if group:
                    avg_y = sum(l["y"] for l in group) / len(group)
                    click_x = chat_list_left + 140  # 列表中间位置
                    click_y = chat_list_top + avg_y + 30  # 加上偏移，确保点击到项目
                    
                    contacts.append({
                        "index": i,
                        "name": name_text or f"未知联系人{i+1}",
                        "last_message": message_text or "",
                        "position": {"x": int(click_x), "y": int(click_y)},
                        "rect": {
                            "left": chat_list_left,
                            "top": int(chat_list_top + avg_y - 20),
                            "right": chat_list_right,
                            "bottom": int(chat_list_top + avg_y + 50)
                        }
                    })
            
            logger.info(f"识别到 {len(contacts)} 个联系人")
            return {
                "status": "success",
                "data": {
                    "contacts": contacts,
                    "total": len(contacts)
                }
            }
            
        except ImportError as e:
            return {"status": "error", "message": f"OCR 模块导入失败: {e}"}
        except Exception as e:
            logger.error(f"获取聊天列表失败: {e}")
            return {"status": "error", "message": f"获取聊天列表失败: {str(e)}"}
    
    def click_contact(self, contact_name: str = None, position: dict = None) -> Dict:
        """点击联系人进入聊天
        
        Args:
            contact_name: 联系人名称（需要先通过 get_chat_list 获取）
            position: 点击位置 {"x": 100, "y": 200}
        
        Returns:
            Dict: {"status": "success/error", "message": "..."}
        """
        try:
            # 确保窗口就绪
            w, error = self._ensure_window_ready()
            if error:
                return {"status": "error", "message": error}
            
            # 如果提供了位置，直接点击
            if position and "x" in position and "y" in position:
                pyautogui.click(position["x"], position["y"])
                time.sleep(0.5)
                return {"status": "success", "message": f"已点击位置 ({position['x']}, {position['y']})"}
            
            # 如果提供了联系人名称，先获取聊天列表查找位置
            if contact_name:
                result = self.get_chat_list()
                if result["status"] != "success":
                    return result
                
                for contact in result["data"]["contacts"]:
                    if contact_name in contact["name"]:
                        pyautogui.click(contact["position"]["x"], contact["position"]["y"])
                        time.sleep(0.5)
                        return {"status": "success", "message": f"已点击联系人: {contact['name']}"}
                
                return {"status": "error", "message": f"未找到联系人: {contact_name}"}
            
            return {"status": "error", "message": "请提供 contact_name 或 position 参数"}
            
        except Exception as e:
            logger.error(f"点击联系人失败: {e}")
            return {"status": "error", "message": f"点击联系人失败: {str(e)}"}
    
    def monitor_chat_changes(self, previous_state: dict = None, interval: float = 2.0) -> Dict:
        """监控聊天列表变化
        
        Args:
            previous_state: 上一次的状态（用于对比变化）
            interval: 监控间隔（秒）
        
        Returns:
            Dict: {
                "status": "success",
                "data": {
                    "current_state": {...},  # 当前状态
                    "changes": [  # 变化列表
                        {
                            "type": "new_message/position_change",
                            "contact": "联系人名称",
                            "details": "详细描述"
                        }
                    ],
                    "unread_contacts": [  # 有未读消息的联系人
                        {"name": "...", "position": {...}}
                    ]
                }
            }
        """
        try:
            # 获取当前聊天列表状态
            current_result = self.get_chat_list()
            if current_result["status"] != "success":
                return current_result
            
            current_contacts = current_result["data"]["contacts"]
            changes = []
            unread_contacts = []
            
            # 如果有之前的状态，进行对比
            if previous_state and "contacts" in previous_state:
                prev_contacts = {c["name"]: c for c in previous_state["contacts"]}
                
                for i, contact in enumerate(current_contacts):
                    name = contact["name"]
                    prev_contact = prev_contacts.get(name)
                    
                    # 位置变化（新消息置顶）
                    if prev_contact:
                        prev_index = prev_contact.get("index", -1)
                        if i != prev_index:
                            changes.append({
                                "type": "position_change",
                                "contact": name,
                                "details": f"位置从 {prev_index} 变为 {i}（可能有新消息）"
                            })
                        
                        # 消息内容变化
                        if contact["last_message"] != prev_contact.get("last_message", ""):
                            changes.append({
                                "type": "new_message",
                                "contact": name,
                                "details": f"新消息: {contact['last_message']}"
                            })
                            unread_contacts.append(contact)
                    else:
                        # 新出现的联系人（置顶的新消息）
                        changes.append({
                            "type": "new_contact",
                            "contact": name,
                            "details": f"新联系人出现在列表: {name}"
                        })
                        unread_contacts.append(contact)
            
            return {
                "status": "success",
                "data": {
                    "current_state": {"contacts": current_contacts},
                    "changes": changes,
                    "unread_contacts": unread_contacts,
                    "has_changes": len(changes) > 0
                }
            }
            
        except Exception as e:
            logger.error(f"监控聊天变化失败: {e}")
            return {"status": "error", "message": f"监控聊天变化失败: {str(e)}"}
    
    def auto_reply_to_contact(self, contact_name: str, message: str) -> Dict:
        """自动回复指定联系人
        
        Args:
            contact_name: 联系人名称
            message: 回复内容
        
        Returns:
            Dict: {"status": "success/error", "message": "..."}
        """
        try:
            # 确保窗口就绪
            w, error = self._ensure_window_ready()
            if error:
                return {"status": "error", "message": error}
            
            # 1. 点击联系人进入聊天
            click_result = self.click_contact(contact_name=contact_name)
            if click_result["status"] != "success":
                return click_result
            
            time.sleep(0.5)
            
            # 2. 发送消息
            success, err = self.send_message(message)
            if success:
                return {"status": "success", "message": f"已回复 {contact_name}: {message}"}
            else:
                return {"status": "error", "message": f"发送失败: {err}"}
                
        except Exception as e:
            logger.error(f"自动回复失败: {e}")
            return {"status": "error", "message": f"自动回复失败: {str(e)}"}
    
    def start_chat_monitor(self, reply_handler: callable = None, interval: float = 3.0, 
                           max_loops: int = 100, auto_reply_message: str = None) -> Dict:
        """启动聊天监控（阻塞式）
        
        Args:
            reply_handler: 自定义回复处理函数 (contact_name, last_message) -> reply_message
            interval: 检查间隔（秒）
            max_loops: 最大循环次数
            auto_reply_message: 自动回复的消息内容（如果不使用 reply_handler）
        
        Returns:
            Dict: 监控结果统计
        """
        logger.info("启动聊天监控...")
        
        # 获取初始状态
        initial_result = self.get_chat_list()
        if initial_result["status"] != "success":
            return initial_result
        
        previous_state = {"contacts": initial_result["data"]["contacts"]}
        stats = {"loops": 0, "replies_sent": 0, "changes_detected": 0}
        
        try:
            for loop in range(max_loops):
                stats["loops"] = loop + 1
                time.sleep(interval)
                
                # 检测变化
                monitor_result = self.monitor_chat_changes(previous_state)
                if monitor_result["status"] != "success":
                    logger.warning(f"监控检查失败: {monitor_result['message']}")
                    continue
                
                # 更新状态
                previous_state = monitor_result["data"]["current_state"]
                
                # 处理变化
                if monitor_result["data"]["has_changes"]:
                    stats["changes_detected"] += 1
                    
                    for change in monitor_result["data"]["changes"]:
                        logger.info(f"检测到变化: {change}")
                    
                    # 自动回复
                    if reply_handler or auto_reply_message:
                        for contact in monitor_result["data"]["unread_contacts"]:
                            contact_name = contact["name"]
                            last_message = contact["last_message"]
                            
                            # 确定回复内容
                            if reply_handler:
                                reply = reply_handler(contact_name, last_message)
                            else:
                                reply = auto_reply_message
                            
                            if reply:
                                reply_result = self.auto_reply_to_contact(contact_name, reply)
                                if reply_result["status"] == "success":
                                    stats["replies_sent"] += 1
                                    logger.info(f"已自动回复 {contact_name}: {reply}")
                                else:
                                    logger.warning(f"回复失败: {reply_result['message']}")
                                
                                time.sleep(1)  # 回复间隔
            
        except KeyboardInterrupt:
            logger.info("监控被用户中断")
        
        return {
            "status": "success",
            "data": {
                "stats": stats,
                "message": f"监控结束，共检测 {stats['changes_detected']} 次变化，发送 {stats['replies_sent']} 条回复"
            }
        }


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

def get_chat_list(count=5):
    """获取聊天列表"""
    return _manager.get_chat_list(count=count)

def click_contact(contact_name=None, position=None):
    """点击联系人"""
    return _manager.click_contact(contact_name=contact_name, position=position)

def monitor_chat_changes(previous_state=None):
    """监控聊天变化"""
    return _manager.monitor_chat_changes(previous_state=previous_state)

def auto_reply(contact_name, message):
    """自动回复联系人"""
    return _manager.auto_reply_to_contact(contact_name=contact_name, message=message)

def start_monitor(interval=3.0, max_loops=100, auto_reply_message=None):
    """启动聊天监控"""
    return _manager.start_chat_monitor(
        interval=interval, 
        max_loops=max_loops, 
        auto_reply_message=auto_reply_message
    )


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
    
    # get_chat_list
    p_chat_list = subparsers.add_parser("get_chat_list", help="获取聊天列表")
    p_chat_list.add_argument("--count", type=int, default=5, help="获取数量")
    
    # click_contact
    p_click_contact = subparsers.add_parser("click_contact", help="点击联系人")
    p_click_contact.add_argument("--name", type=str, default=None, help="联系人名称")
    p_click_contact.add_argument("--x", type=int, default=None, help="点击X坐标")
    p_click_contact.add_argument("--y", type=int, default=None, help="点击Y坐标")
    
    # auto_reply
    p_auto_reply = subparsers.add_parser("auto_reply", help="自动回复联系人")
    p_auto_reply.add_argument("--name", type=str, required=True, help="联系人名称")
    p_auto_reply.add_argument("--message", type=str, required=True, help="回复内容")
    
    # start_monitor
    p_monitor = subparsers.add_parser("start_monitor", help="启动聊天监控")
    p_monitor.add_argument("--interval", type=float, default=3.0, help="检查间隔(秒)")
    p_monitor.add_argument("--max_loops", type=int, default=100, help="最大循环次数")
    p_monitor.add_argument("--auto_reply", type=str, default=None, help="自动回复内容")
    
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
    elif args.action == "get_chat_list":
        result = get_chat_list(args.count)
    elif args.action == "click_contact":
        position = {"x": args.x, "y": args.y} if args.x and args.y else None
        result = click_contact(contact_name=args.name, position=position)
    elif args.action == "auto_reply":
        result = auto_reply(contact_name=args.name, message=args.message)
    elif args.action == "start_monitor":
        result = start_monitor(
            interval=args.interval, 
            max_loops=args.max_loops, 
            auto_reply_message=args.auto_reply
        )
    
    print(json.dumps(result, ensure_ascii=False, indent=2))
