#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/3/12
# @File    : agent.py
# @Software: PyCharm

import sys
import io
import os

# ============== 统一编码处理（解决 Windows 控制台编码问题）==============
if sys.platform == 'win32':
    # 1. 先包装 stdout/stderr 为 UTF-8（关键！解决 print 输出乱码）
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    # 2. 再设置控制台代码页为 UTF-8
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleOutputCP(65001)
    kernel32.SetConsoleCP(65001)

# 输出编码标记（供其他模块使用）
_OUTPUT_ENCODING = 'utf-8'

import base64
import time
import logging
import json
import subprocess
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

# 获取 user32 API (ctypes 已在文件开头导入)
user32 = ctypes.windll.user32

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

    WAKE_UP_HOTKEY = ('ctrl', 'alt', 'w')
    STATE_FILE = "wechat_chat_records.json"  # 聊天记录持久化文件

    # 跳过的联系人类型
    SKIP_CONTACTS = [
        "文件传输助手", "File Transfer", "文件传输",
        "公众号", "服务号", "微信支付",
        "腾讯", "系统", "官方",
        "折叠聊天", "折叠的聊天"  # 折叠聊天功能
    ]

    def __init__(self):
        self._app = None
        self._window = None
        self._gw_window = None  # pygetwindow 窗口
        # 聊天记录状态 {contacts_order: [...], contacts: {name: {last_opponent_message, ...}}}
        self._chat_records = {}
        self._load_records()  # 启动时加载记录

    def _get_state_file_path(self) -> str:
        """获取状态文件路径"""
        # 状态文件保存在agent.py所在目录
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, self.STATE_FILE)

    def _load_records(self):
        """从文件加载聊天记录"""
        try:
            state_file = self._get_state_file_path()
            if os.path.exists(state_file):
                with open(state_file, 'r', encoding='utf-8') as f:
                    self._chat_records = json.load(f)
        except Exception as e:
            logger.debug(f"加载聊天记录失败: {e}")
            self._chat_records = {}

    def _save_records(self):
        """保存聊天记录到文件"""
        try:
            state_file = self._get_state_file_path()
            with open(state_file, 'w', encoding='utf-8') as f:
                json.dump(self._chat_records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.debug(f"保存聊天记录失败: {e}")

    def _find_gw_window(self) -> Optional[Any]:
        """使用 pygetwindow 查找微信窗口"""
        # 微信窗口标题通常是：微信、WeChat、Weixin，或者以这些开头的标题
        # 排除 IDE/编辑器窗口
        script_name = os.path.basename(__file__).lower()

        # IDE 和编辑器关键词（用于排除）
        IDE_KEYWORDS = [".py", "pycharm", "idea", "vscode", "visual studio",
                        "sublime", "atom", "notepad", "editor", "ide"]

        for keyword in ["微信", "WeChat", "Weixin"]:
            windows = gw.getWindowsWithTitle(keyword)
            for w in windows:
                title_lower = w.title.lower()

                # 排除脚本名或 agent.py
                if script_name in title_lower or "agent.py" in title_lower:
                    logger.debug(f"跳过脚本相关窗口: title={w.title}")
                    continue

                # 排除 IDE/编辑器窗口
                is_ide = False
                for ide_keyword in IDE_KEYWORDS:
                    if ide_keyword in title_lower:
                        # 特殊处理：微信标题可能包含这些词（罕见），但尺寸应该匹配
                        is_ide = True
                        break

                if is_ide:
                    logger.debug(f"跳过 IDE 窗口: title={w.title}")
                    continue

                # 微信窗口特征：尺寸较大，且标题通常是 "微信" 或 "WeChat"
                # 允许标题以微信关键词开头
                title_starts_with_keyword = title_lower.startswith(keyword.lower())

                # 如果标题以关键词开头，直接接受
                if title_starts_with_keyword and w.width > 200 and w.height > 200:
                    logger.debug(f"找到微信窗口: title={w.title}, size=({w.width}x{w.height})")
                    return w

                # 否则，检查窗口类名（微信主窗口类名：WeChatMainWndForPC）
                try:
                    hwnd = w._hWnd if hasattr(w, '_hWnd') else None
                    if hwnd:
                        class_name = win32gui.GetClassName(hwnd)
                        if "WeChat" in class_name or "wechat" in class_name.lower():
                            if w.width > 200 and w.height > 200:
                                logger.debug(f"找到微信窗口(类名匹配): title={w.title}, class={class_name}")
                                return w
                except Exception:
                    pass

                # 最后：如果标题完全是关键词（精确匹配）
                if title_lower.strip() in ["微信", "wechat", "weixin"]:
                    if w.width > 200 and w.height > 200:
                        logger.debug(f"找到微信窗口(精确匹配): title={w.title}")
                        return w

                logger.debug(f"跳过窗口(不匹配): title={w.title}, size=({w.width}x{w.height})")
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

    def _ensure_window_ready(self) -> Tuple[Optional[Any], Optional[str]]:
        """确保微信窗口就绪（统一的状态检查方法）

        所有操作方法应调用此方法来确保窗口状态正确。
        微信需要用户手动打开，此方法只负责检测和置顶窗口。

        Returns:
            Tuple[Optional[window], Optional[error_msg]]:
                - 成功: (window_object, None)
                - 失败: (None, error_message)
        """
        # 步骤1: 查找微信窗口（最多尝试 3 次）
        w = None
        for attempt in range(3):
            w = self._find_gw_window()
            if w:
                break
            time.sleep(0.1)

        # 步骤2: 判断窗口状态并处理
        if not w:
            # 窗口不存在（微信在托盘区），用快捷键唤醒（最多尝试 3 次）
            for attempt in range(3):
                pyautogui.hotkey(*self.WAKE_UP_HOTKEY)
                time.sleep(0.3)
                w = self._find_gw_window()
                if w:
                    break

            if not w:
                return None, "微信未运行，请手动打开微信"

        # 步骤3: 获取窗口句柄
        hwnd = w._hWnd if hasattr(w, '_hWnd') else None
        if not hwnd:
            return None, "无法获取微信窗口句柄"

        # 步骤4: 检查窗口是否最小化
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.1)

        # 步骤5: 激活窗口到前台（最多尝试 3 次）
        for attempt in range(3):
            self._bring_window_to_front(w)
            time.sleep(0.1)

            # 验证窗口是否真的在前台
            foreground_hwnd = win32gui.GetForegroundWindow()
            if foreground_hwnd == hwnd:
                break

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
                return False

            # ===== 方法1: 先恢复窗口（如果最小化）=====
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.05)

            # ===== 方法2: 使用 AttachThreadInput 绕过 SetForegroundWindow 限制 =====
            foreground_hwnd = win32gui.GetForegroundWindow()
            foreground_thread = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]
            current_thread = win32api.GetCurrentThreadId()
            target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]

            # 附加线程输入（使用 ctypes 调用 AttachThreadInput）
            attach_threads = []
            try:
                if current_thread != target_thread:
                    user32.AttachThreadInput(current_thread, target_thread, True)
                    attach_threads.append((current_thread, target_thread))
                if foreground_thread != target_thread and foreground_thread != current_thread:
                    user32.AttachThreadInput(foreground_thread, target_thread, True)
                    attach_threads.append((foreground_thread, target_thread))
            except Exception:
                pass

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
                    user32.AttachThreadInput(t1, t2, False)
                except Exception:
                    pass

            time.sleep(0.1)  # 缩短等待

            # ===== 方法5: 如果前面方法失败，使用模拟按键绕过限制 =====
            new_foreground = win32gui.GetForegroundWindow()
            if new_foreground != hwnd:
                # 模拟按下 Alt 键（这会允许 SetForegroundWindow 工作）
                win32api.keybd_event(0x12, 0, 0, 0)  # Alt down
                time.sleep(0.05)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                win32api.keybd_event(0x12, 0, 2, 0)  # Alt up
                time.sleep(0.1)

            # 验证激活结果
            new_foreground = win32gui.GetForegroundWindow()
            if new_foreground == hwnd:
                return True
            else:
                # ===== 方法6: 最后尝试 pygetwindow =====
                try:
                    w.activate()
                    time.sleep(0.1)
                    new_foreground = win32gui.GetForegroundWindow()
                    if new_foreground == hwnd:
                        return True
                except Exception:
                    pass

                return True  # 仍然返回 True，让调用者决定是否重试

        except Exception:
            # 最后尝试 pygetwindow 方法
            try:
                w.activate()
                time.sleep(0.3)
                return True
            except Exception:
                return False

    def get_main_window(self, force_refresh: bool = False, activate_first: bool = False) -> Optional[Any]:
        """获取微信主窗口（返回 pygetwindow 窗口对象）

        Args:
            force_refresh: 是否强制刷新窗口
            activate_first: 是否先激活窗口
        """
        # 使用统一的窗口状态检查方法
        w, error = self._ensure_window_ready()
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

    def capture(self, show_flash: bool = False) -> Optional[Image.Image]:
        """截取微信窗口（使用 win32 API）

        Args:
            show_flash: 是否显示截图闪烁提示（用系统图片查看器打开截图0.5秒）
        """
        should_flash = show_flash

        # 先激活窗口
        w = self.get_main_window(force_refresh=True, activate_first=True)
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
            time.sleep(0.1)
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

            # 如果启用了闪烁提示，显示截图区域
            if should_flash:
                self._flash_capture_area(left, top, right, bottom, FLASH_DURATION)

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

            logger.debug(f"开始截图: hwnd={hwnd}, 尺寸={width}x{height}")

            # 获取窗口 DC
            hwndDC = win32gui.GetWindowDC(hwnd)
            if not hwndDC:
                logger.error("无法获取窗口 DC")
                return None

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

            logger.debug(f"截图成功: {bmpinfo['bmWidth']}x{bmpinfo['bmHeight']}")
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
        """获取微信状态"""
        # 直接调用 get_main_window
        w = self.get_main_window(activate_first=True)
        if not w:
            return {"status": "not_running", "message": "微信未运行，请手动打开微信"}

        rect = self.get_window_rect()

        return {
            "status": "running",
            "title": w.title,
            "position": {"x": rect["left"], "y": rect["top"]} if rect else None,
            "size": {"width": rect["width"], "height": rect["height"]} if rect else None
        }

    def click(self, x: int, y: int) -> bool:
        """点击指定坐标"""
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
        """输入文本"""
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
        """滚动页面

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
        # 1. 确保窗口可见
        w = self.get_main_window(activate_first=True)
        if not w:
            return False, "未找到微信窗口"

        rect = self.get_window_rect()
        if not rect:
            return False, "无法获取窗口位置"

        logger.debug(f"窗口位置: ({rect['left']}, {rect['top']}), 大小: ({rect['width']}, {rect['height']})")

        # 2. 点击输入框（使用相对比例）
        input_x = rect["left"] + rect["width"] // 2
        input_y = rect["bottom"] - int(rect["height"] * 0.35)  # 输入框在底部约8%位置
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

        return True, None

    def send_image(self, image_path: str) -> Tuple[bool, Optional[str]]:
        """发送图片

        Args:
            image_path: 图片文件路径

        Returns:
            Tuple[bool, Optional[str]]: (成功?, 错误信息)
        """
        import win32clipboard

        # 1. 检查图片文件是否存在
        if not os.path.exists(image_path):
            return False, f"图片文件不存在: {image_path}"

        # 2. 确保窗口可见
        w = self.get_main_window(activate_first=True)
        if not w:
            return False, "未找到微信窗口"

        rect = self.get_window_rect()
        if not rect:
            return False, "无法获取窗口位置"

        logger.debug(f"窗口位置: ({rect['left']}, {rect['top']}), 大小: ({rect['width']}, {rect['height']})")

        # 3. 将图片复制到剪贴板
        try:
            # 打开图片
            img = Image.open(image_path)

            # 将图片复制到剪贴板（Windows 方式）
            output = BytesIO()
            img.convert('RGB').save(output, 'BMP')
            data = output.getvalue()[14:]  # BMP 文件头是 14 字节，需要去掉
            output.close()

            # 设置剪贴板数据
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()

            logger.info("图片已复制到剪贴板")

        except Exception as e:
            logger.error(f"复制图片到剪贴板失败: {e}")
            return False, f"复制图片到剪贴板失败: {str(e)}"

        time.sleep(0.3)

        # 4. 点击输入框（使用相对比例）
        input_x = rect["left"] + rect["width"] // 2
        input_y = rect["bottom"] - int(rect["height"] * 0.08)  # 输入框在底部约8%位置
        self.click(input_x, input_y)
        time.sleep(0.3)

        # 5. 粘贴图片 (Ctrl+V)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.8)  # 等待图片预览加载

        # 6. 发送 (Enter)
        pyautogui.press('enter')
        time.sleep(0.5)

        logger.info("图片发送完成")
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
            w = self.get_main_window(force_refresh=True, activate_first=True)
            if not w:
                return {"status": "error", "message": "无法获取微信窗口"}

            # 验证窗口有效性
            try:
                _ = w.left  # 测试窗口是否有效
            except Exception:
                # 窗口失效，重新获取
                logger.debug("OCR: 窗口失效，重新获取")
                self._gw_window = None
                w = self.get_main_window(force_refresh=True, activate_first=True)
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
        """获取微信窗口OCR结果

        直接返回OCR识别的原始结果，由AI来判断联系人和聊天内容。
        不再在代码中解析，避免窗口大小变化导致的坐标判断问题。

        Args:
            count: 保留参数（兼容旧接口）

        Returns:
            Dict: {
                "status": "success/error",
                "data": {
                    "ocr_results": [
                        {
                            "text": "识别的文本",
                            "confidence": 0.95,
                            "position": {"x": 100, "y": 200},  # 左上角坐标
                            "box": [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]  # 四个角坐标
                        },
                        ...
                    ],
                    "window_rect": {"left": 0, "top": 0, "width": 800, "height": 600},
                    "total": 10
                }
            }
        """
        try:
            # 确保窗口就绪
            w, error = self._ensure_window_ready()
            if error:
                return {"status": "error", "message": error}

            # 获取窗口位置
            rect = self.get_window_rect()
            if not rect:
                return {"status": "error", "message": "无法获取窗口位置"}

            # 截取整个微信窗口
            img = self.capture()
            if not img:
                logger.error("截图失败")
                return {"status": "error", "message": "截图失败"}

            # 保存窗口位置用于后续计算点击坐标
            self._window_rect = rect

            # OCR 识别整个窗口（直接使用 numpy 数组）
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
            result, _ = ocr(np.array(img))

            if not result:
                logger.error("OCR 未识别到内容")
                return {"status": "error", "message": "OCR 未识别到内容"}

            # 构建OCR结果列表
            ocr_results = []
            for item in result:
                box, text, conf = item
                # box 是四个点的坐标 [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
                ocr_results.append({
                    "text": text,
                    "confidence": float(conf),
                    "position": {"x": int(box[0][0]), "y": int(box[0][1])},  # 左上角坐标
                    "box": [[int(p[0]), int(p[1])] for p in box]  # 四个角坐标
                })

            # 打印OCR结果
            logger.info("===== OCR 识别结果 =====")
            for i, item in enumerate(ocr_results):
                logger.info(f"  [{i}] 文本: '{item['text']}' | 置信度: {item['confidence']:.2f} | 位置: ({item['position']['x']}, {item['position']['y']})")
            logger.info(f"===== 共识别到 {len(ocr_results)} 条文本 =====")

            return {
                "status": "success",
                "data": {
                    "ocr_results": ocr_results,
                    "window_rect": rect,
                    "total": len(ocr_results)
                }
            }

        except ImportError as e:
            return {"status": "error", "message": f"OCR 模块导入失败: {e}"}
        except Exception as e:
            logger.error(f"OCR识别失败: {e}")
            return {"status": "error", "message": f"OCR识别失败: {str(e)}"}

    def _verify_chat_window_open(self) -> bool:
        """验证聊天窗口是否已打开（通过OCR检查是否有发送按钮）

        Returns:
            bool: True=聊天窗口已打开，False=聊天窗口未打开
        """
        try:
            # 获取窗口位置
            rect = self.get_window_rect()
            if not rect:
                return False

            # 截图
            img = self.capture()
            if not img:
                return False

            # 裁剪右侧底部区域（发送按钮通常在右下角）
            # 使用相对比例，适应不同窗口大小
            # 微信聊天窗口的发送按钮位置：右侧底部
            send_button_region = img.crop((
                rect["width"] // 2,  # 左边界：窗口中间
                int(rect["height"] * 0.75),  # 上边界：窗口高度75%处
                rect["width"],  # 右边界：窗口右边
                int(rect["height"] * 0.95)  # 下边界：窗口高度95%处
            ))

            # OCR识别
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
            result, _ = ocr(np.array(send_button_region))

            if not result:
                return False

            # 检查是否有"发送"按钮
            # 微信发送按钮通常显示"发送(S)"或"发送"
            for item in result:
                text = item[1]  # item格式: [box, text, confidence]
                if "发送" in text:
                    return True

            return False

        except Exception as e:
            logger.debug(f"验证聊天窗口失败: {e}")
            return False

    def click_contact(self, position: dict) -> Dict:
        """点击指定坐标进入聊天

        Args:
            position: 点击位置 {"x": 100, "y": 200}（屏幕绝对坐标）

        Returns:
            Dict: {"status": "success/error", "message": "..."}
        """
        try:
            # 确保窗口就绪
            w, error = self._ensure_window_ready()
            if error:
                return {"status": "error", "message": error}

            # 获取点击位置
            if not position or "x" not in position or "y" not in position:
                return {"status": "error", "message": "请提供 position 参数，格式: {\"x\": 100, \"y\": 200}"}

            click_x = position["x"]
            click_y = position["y"]

            # 点击并验证窗口状态
            max_attempts = 2

            for attempt in range(max_attempts):
                pyautogui.click(click_x, click_y)
                time.sleep(0.5)

                if self._verify_chat_window_open():
                    return {
                        "status": "success",
                        "message": f"已点击位置 ({click_x}, {click_y}) 并进入聊天窗口"
                    }
                else:
                    if attempt < max_attempts - 1:
                        time.sleep(0.3)
                        continue

            return {"status": "error", "message": "无法打开聊天窗口，请检查微信状态"}

        except Exception as e:
            logger.error(f"点击失败: {e}")
            return {"status": "error", "message": f"点击失败: {str(e)}"}

    def auto_reply_to_contact(self, position: dict, message: str) -> Dict:
        """自动回复指定位置的联系人

        Args:
            position: 点击位置 {"x": 100, "y": 200}（屏幕绝对坐标）
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
            click_result = self.click_contact(position=position)
            if click_result["status"] != "success":
                return click_result

            time.sleep(0.3)

            # 2. 发送消息
            success, err = self.send_message(message)
            if success:
                return {"status": "success", "message": f"已发送消息: {message}"}
            else:
                return {"status": "error", "message": f"发送失败: {err}"}

        except Exception as e:
            logger.error(f"自动回复失败: {e}")
            return {"status": "error", "message": f"自动回复失败: {str(e)}"}

    def _should_skip_contact(self, contact_name: str) -> bool:
        """判断是否应该跳过此联系人"""
        if not contact_name:
            return True
        contact_lower = contact_name.lower()
        for skip in self.SKIP_CONTACTS:
            if skip.lower() in contact_lower:
                return True
        return False

    # ==================== 布局常量（适应不同窗口大小）====================
    # 微信左侧联系人列表：名称约 230px，时间约 330px
    # 这里使用固定宽度 + 动态调整
    LEFT_LIST_WIDTH = 350          # 左侧联系人列表宽度（包含时间）
    LEFT_LIST_MIN_RATIO = 0.35     # 最小比例（小窗口时）
    LEFT_LIST_MAX_RATIO = 0.40     # 最大比例（大窗口时）

    # 消息区域参数
    CHAT_AREA_MARGIN = 60          # 聊天区域边距
    OPPONENT_MSG_RATIO_MAX = 0.60  # 对方消息 X 坐标最大比例
    CHAT_BOTTOM_MARGIN = 80        # 聊天区域底部边距

    def _get_left_boundary(self, window_width: int) -> int:
        """计算左侧联系人列表边界（适应不同窗口大小）

        微信左侧联系人列表宽度固定约 230px，
        小窗口时使用比例，大窗口时使用固定宽度。
        """
        # 计算固定宽度边界
        fixed_boundary = self.LEFT_LIST_WIDTH

        # 计算比例边界
        ratio_boundary = int(window_width * self.LEFT_LIST_MIN_RATIO)

        # 计算最大比例边界
        max_ratio_boundary = int(window_width * self.LEFT_LIST_MAX_RATIO)

        # 选择策略：
        # - 小窗口（<700px）：使用比例，避免左侧列表占用太多空间
        # - 中等窗口（700-900px）：使用固定宽度
        # - 大窗口（>900px）：使用固定宽度，但不超过最大比例
        if window_width < 700:
            return ratio_boundary
        elif window_width < 900:
            return fixed_boundary
        else:
            return min(fixed_boundary, max_ratio_boundary)

    def parse_contacts_from_ocr(self, ocr_results: list, window_width: int) -> list:
        """从 OCR 结果中解析联系人列表

        Args:
            ocr_results: OCR 识别结果列表
            window_width: 窗口宽度

        Returns:
            list: 联系人列表 [{"name": "张三", "x": 125, "y": 200, "preview": "你好"}, ...]
        """
        import re

        # 计算左侧边界
        left_boundary = self._get_left_boundary(window_width)

        # 过滤左侧区域的文本
        left_texts = []
        for item in ocr_results:
            pos = item.get('position', {})
            x = pos.get('x', 0)
            y = pos.get('y', 0)
            text = item.get('text', '')

            # 只取左侧区域，排除头像区域（X > 50）
            if 50 < x < left_boundary:
                # 排除搜索框
                if '搜索' in text:
                    continue
                left_texts.append({'x': x, 'y': y, 'text': text})

        # 按 Y 坐标排序
        left_texts.sort(key=lambda t: t['y'])

        # 分组：每个联系人条目高度约 64 像素
        # 名称和预览消息的 Y 差通常 < 30px，相邻联系人的 Y 差 > 50px
        # 使用 40px 作为阈值，确保同一联系人的内容在同一组
        groups = []
        current_group = []
        prev_y = -100

        for item in left_texts:
            if item['y'] - prev_y > 40:  # 新的一组
                if current_group:
                    groups.append(current_group)
                current_group = [item]
            else:  # 同一组
                current_group.append(item)
            prev_y = item['y']

        if current_group:
            groups.append(current_group)




        # 从每组中解析联系人信息
        contacts = []
        for group in groups:
            group.sort(key=lambda t: t['y'])  # 按 Y 坐标排序

            # 过滤条件1：组内只有单个文本，可能是右侧消息（正常联系人应该有名称+时间+预览）
            if len(group) < 2:
                continue

            # 名称是 Y 坐标最小的（最上面那行）
            # 微信联系人条目结构：名称在上，消息预览在下
            name_item = group[0]

            # 过滤条件1.5：名称不能是时间格式（右侧消息的时间可能被错误识别为名称）
            if self._is_time_format(name_item['text']):
                continue

            # 检查整组是否应该跳过
            should_skip_group = False
            for item in group:
                if self._should_skip_contact(item['text']):
                    should_skip_group = True
                    break

            if should_skip_group:
                continue  # 跳过整组（包括公众号的消息预览）

            # 过滤条件2：检查组内是否有时间格式（正常联系人条目应该有时间）
            time_item = self._find_time_item(group)
            if not time_item:
                continue  # 没有时间，可能是右侧聊天消息，跳过

            # 过滤条件3：名称和时间应该在同一行（Y 差值小）
            # 微信联系人列表中，名称和时间 Y 坐标几乎相同（Y 差 < 10）
            name_y = name_item['y']
            time_y = time_item['y']
            if abs(time_y - name_y) > 10:
                continue  # 名称和时间不在同一行，可能是右侧内容

            # 长度限制：联系人名称通常 1-20 个字符
            if 1 <= len(name_item['text']) <= 20:
                # 解析消息预览
                preview = self._extract_message_preview(group, name_item)

                contacts.append({
                    'name': name_item['text'],
                    'x': name_item['x'],
                    'y': name_item['y'],
                    'preview': preview,
                    'time': time_item['text']  # 添加时间字段
                })

        return contacts

    def _is_time_format(self, text: str) -> bool:
        """检查文本是否是时间格式

        Args:
            text: 待检查的文本

        Returns:
            bool: 是否是时间格式
        """
        import re

        time_patterns = [
            r'^\d{1,2}:\d{2}$',           # 17:00, 9:30
            r'^昨天', r'^前天',            # 昨天, 前天, 昨天14:39
            r'^今天',                      # 今天
            r'^刚刚',                      # 刚刚
            r'^\d+分钟前$',                # 5分钟前
            r'^星期[一二三四五六日]$',       # 星期一
            r'^周[一二三四五六日]$',         # 周一
            r'^\d{1,2}月\d{1,2}日$',       # 3月15日
        ]

        for pattern in time_patterns:
            if re.match(pattern, text):
                return True

        return False

    def _find_time_item(self, group: list) -> dict:
        """找到组内的时间项

        Args:
            group: 同一组的文本列表

        Returns:
            dict: 时间项，如果没有则返回 None
        """
        import re

        time_patterns = [
            r'^\d{1,2}:\d{2}$',           # 17:00, 9:30
            r'^昨天', r'^前天',            # 昨天, 前天, 昨天14:39
            r'^今天',                      # 今天
            r'^刚刚', r'^刚刚',            # 刚刚
            r'^\d+分钟前$',                # 5分钟前
            r'^星期[一二三四五六日]$',       # 星期一
            r'^周[一二三四五六日]$',         # 周一
            r'^\d{1,2}月\d{1,2}日$',       # 3月15日
        ]

        # 按 X 坐标排序，找 X 最大的（时间通常在右侧）
        group_by_x = sorted(group, key=lambda t: t['x'], reverse=True)

        for item in group_by_x:
            text = item['text']
            # 1. 先检查标准时间格式
            for pattern in time_patterns:
                if re.match(pattern, text):
                    return item
            # 2. 检查是否可能是 OCR 误识别的时间（X 坐标大，文本短）
            # 时间通常在右侧（X > 280），且文本较短（< 15 字符）
            if item['x'] > 280 and len(text) <= 15:
                # 检查是否和名称在同一行（Y 差值 < 10）
                name_y = group[0]['y']  # 名称是 Y 最小的
                if abs(item['y'] - name_y) <= 10:
                    return item

        return None

    def _extract_message_preview(self, group: list, name_item: dict) -> str:
        """从分组中提取消息预览

        微信联系人条目结构：
        - 联系人名称：X坐标最小
        - 消息预览：在名称下方，X坐标和名称相近
        - 时间：X坐标最大（在右侧），如 "17:00"、"昨天"、"星期四"

        Args:
            group: 同一组的文本列表
            name_item: 联系人名称项

        Returns:
            str: 消息预览内容，如果没有则返回空字符串
        """
        import re

        # 时间格式的正则表达式
        time_patterns = [
            r'^\d{1,2}:\d{2}$',           # 17:00, 9:30
            r'^昨天$', r'^前天$',          # 昨天, 前天
            r'^星期[一二三四五六日]$',       # 星期一
            r'^周[一二三四五六日]$',         # 周一
            r'^\d{1,2}月\d{1,2}日$',       # 3月15日
        ]

        def is_time_format(text: str) -> bool:
            """判断是否是时间格式"""
            for pattern in time_patterns:
                if re.match(pattern, text):
                    return True
            return False

        # 找出所有非名称的文本
        other_items = [item for item in group if item['text'] != name_item['text']]

        if not other_items:
            return ""  # 没有其他文本

        # 按 X 坐标排序，找 X 最大的（通常是时间）
        other_items_by_x = sorted(other_items, key=lambda t: t['x'], reverse=True)

        # 检查 X 最大的项是否是时间
        max_x_item = other_items_by_x[0]

        if is_time_format(max_x_item['text']):
            # 有时间，检查剩余项是否是消息预览
            remaining = [item for item in other_items if item['text'] != max_x_item['text']]
            if remaining:
                # 剩余的按 Y 排序，取 Y 最小的（最靠近名称的）
                remaining.sort(key=lambda t: t['y'])
                return remaining[0]['text']
            else:
                return ""  # 只有时间，没有消息预览
        else:
            # X 最大的不是时间格式
            # 检查是否是消息预览（X 坐标和名称相近）
            name_x = name_item['x']
            for item in other_items_by_x:
                # X 坐标差距小于 50 认为是消息预览
                if item['x'] - name_x < 50 and not is_time_format(item['text']):
                    return item['text']

            # 如果都不是，返回 Y 最小的（排除名称后）
            other_items.sort(key=lambda t: t['y'])
            return other_items[0]['text'] if other_items else ""

    def parse_opponent_message_from_ocr(self, ocr_results: list, window_width: int, window_height: int) -> str:
        """从 OCR 结果中解析对方的最后一条消息

        Args:
            ocr_results: OCR 识别结果列表
            window_width: 窗口宽度
            window_height: 窗口高度

        Returns:
            str: 对方最后一条消息
        """
        import re

        # 计算区域边界
        left_boundary = self._get_left_boundary(window_width)
        chat_area_left = left_boundary

        # 对方消息的 X 坐标范围
        opponent_x_min = chat_area_left + self.CHAT_AREA_MARGIN
        opponent_x_max = int(window_width * self.OPPONENT_MSG_RATIO_MAX)

        # 聊天区域底部（排除输入框）
        chat_area_bottom = window_height - self.CHAT_BOTTOM_MARGIN

        # 过滤对方消息
        opponent_msgs = []
        for item in ocr_results:
            pos = item.get('position', {})
            x = pos.get('x', 0)
            y = pos.get('y', 0)
            text = item.get('text', '')
            confidence = item.get('confidence', 0)

            # 在对方消息区域内
            if opponent_x_min < x < opponent_x_max:
                # 排除底部输入区域
                if y < chat_area_bottom:
                    # 检查是否是有效消息
                    if self._is_valid_message(text, confidence):
                        opponent_msgs.append({'x': x, 'y': y, 'text': text, 'confidence': confidence})

        # 按 Y 坐标排序，取最后一条
        if opponent_msgs:
            opponent_msgs.sort(key=lambda t: t['y'])
            return opponent_msgs[-1]['text']

        return ""

    def _is_valid_message(self, text: str, confidence: float) -> bool:
        """判断是否是有效的消息文本"""
        import re

        # 1. 置信度太低可能是乱码
        if confidence < 0.85:
            return False

        # 2. 长度太短可能是误识别
        if len(text) < 2:
            return False

        # 3. 包含特殊字符组合（OCR 乱码特征）
        garbage_patterns = [
            r'^[@→※口\-\+\*\#\!\?\.\,]+$',  # 纯特殊字符
            r'^.*[→※].*$',  # 包含箭头等特殊符号
            r'^\d+[\"\'].*',  # "4\" (" 这种格式
            r'^[\d\.\,]+$',  # 纯数字
        ]
        for pattern in garbage_patterns:
            if re.match(pattern, text):
                return False

        # 4. 时间格式
        if re.match(r'^昨天\d', text) or re.match(r'^前天\d', text):
            return False
        if re.match(r'^\d{1,2}:\d{2}$', text):
            return False

        # 5. 系统关键词
        if text in ['发送（S)', '发送(S)', '发送']:
            return False

        return True

    def check_new_messages(self) -> Dict:
        """检查微信联系人列表，返回解析后的联系人信息

        返回联系人列表（带预览消息）和已保存的聊天记录，供AI对比判断是否有新消息。

        Returns:
            Dict: {
                "status": "success",
                "data": {
                    "contacts": [
                        {
                            "name": "联系人名称",
                            "x": 125, "y": 200,
                            "preview": "预览消息",
                            "time": "16:09"
                        }
                    ],
                    "saved_records": {
                        "contacts": {name: {last_opponent_message, last_opponent_time}},
                        "contacts_order": [...]
                    },
                    "is_first_init": True/False,
                    "window_rect": {...}
                }
            }
        """
        try:
            # 获取OCR结果
            ocr_result = self.get_chat_list()
            if ocr_result["status"] != "success":
                return ocr_result

            ocr_results = ocr_result["data"]["ocr_results"]
            window_rect = ocr_result["data"]["window_rect"]
            window_width = window_rect["width"]

            # 解析联系人列表（带预览消息）
            contacts = self.parse_contacts_from_ocr(ocr_results, window_width)

            # 为每个联系人添加屏幕绝对坐标（用于点击）
            for contact in contacts:
                contact["screen_x"] = window_rect["left"] + contact["x"]
                contact["screen_y"] = window_rect["top"] + contact["y"]

            # 检查是否是首次初始化
            is_first_init = len(self._chat_records) == 0

            return {
                "status": "success",
                "data": {
                    "contacts": contacts,
                    "saved_records": self._chat_records,
                    "is_first_init": is_first_init,
                    "window_rect": window_rect
                }
            }

        except Exception as e:
            logger.error(f"检查新消息失败: {e}")
            return {"status": "error", "message": f"检查新消息失败: {str(e)}"}

    def verify_chat_window(self, expected_name: str) -> Dict:
        """验证当前聊天窗口是否是指定联系人

        通过OCR识别对话框顶部的联系人名称，验证是否与预期一致。

        Args:
            expected_name: 预期的联系人名称

        Returns:
            Dict: {
                "status": "success/error",
                "data": {
                    "actual_name": "识别到的名称",
                    "expected_name": "预期名称",
                    "matched": True/False
                }
            }
        """
        try:
            # 获取窗口位置
            rect = self.get_window_rect()
            if not rect:
                return {"status": "error", "message": "无法获取窗口位置"}

            # 截图并OCR
            img = self.capture()
            if not img:
                return {"status": "error", "message": "截图失败"}

            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
            result, _ = ocr(np.array(img))

            if not result:
                return {"status": "error", "message": "OCR未识别到内容"}

            # 构建OCR结果
            ocr_results = []
            for item in result:
                box, text, conf = item
                ocr_results.append({
                    "text": text,
                    "confidence": float(conf),
                    "position": {"x": int(box[0][0]), "y": int(box[0][1])}
                })

            # 聊天窗口顶部名称位置：Y < 100, X 在右侧聊天区域
            left_boundary = self._get_left_boundary(rect["width"])
            top_area_texts = []
            for item in ocr_results:
                x = item["position"]["x"]
                y = item["position"]["y"]
                text = item["text"]
                # 顶部区域（Y < 100），右侧聊天区域（X > 左侧边界）
                if y < 100 and x > left_boundary:
                    # 排除一些非名称文本
                    if text not in ["搜索", "发送", "更多", "...", "×"]:
                        top_area_texts.append({"x": x, "y": y, "text": text})

            # 找到顶部最左边的文本作为名称（通常名称在顶部最左边）
            if top_area_texts:
                top_area_texts.sort(key=lambda t: t["x"])
                actual_name = top_area_texts[0]["text"]
            else:
                actual_name = ""

            # 匹配检查（支持部分匹配，如 "CF-李壮" 匹配 "CF-李壮@百博电商"）
            matched = expected_name in actual_name or actual_name in expected_name

            return {
                "status": "success",
                "data": {
                    "actual_name": actual_name,
                    "expected_name": expected_name,
                    "matched": matched
                }
            }

        except Exception as e:
            logger.error(f"验证聊天窗口失败: {e}")
            return {"status": "error", "message": f"验证聊天窗口失败: {str(e)}"}

    def get_chat_records(self) -> Dict:
        """获取当前所有联系人的聊天记录"""
        return {
            "status": "success",
            "data": {
                "chat_records": self._chat_records,
                "contacts_order": self._chat_records.get("contacts_order", []),
                "total": len(self._chat_records.get("contacts_order", []))
            }
        }

    def save_chat_record(self, contact_name: str, last_opponent_message: str,
                         last_opponent_time: str = None, contact_order: int = None) -> Dict:
        """保存单个联系人的聊天记录

        Args:
            contact_name: 联系人名称
            last_opponent_message: 对方发的最后一条消息
            last_opponent_time: 消息时间
            contact_order: 联系人在列表中的顺序（0-4）

        Returns:
            Dict: {"status": "success/error", "message": "..."}
        """
        if not contact_name:
            return {"status": "error", "message": "请提供联系人名称"}

        # 初始化 contacts 字典
        if "contacts" not in self._chat_records:
            self._chat_records["contacts"] = {}

        # 保存联系人聊天记录
        self._chat_records["contacts"][contact_name] = {
            "last_opponent_message": last_opponent_message,
            "last_opponent_time": last_opponent_time or ""
        }

        # 更新联系人顺序
        if contact_order is not None:
            if "contacts_order" not in self._chat_records:
                self._chat_records["contacts_order"] = []
            # 确保列表长度足够
            while len(self._chat_records["contacts_order"]) <= contact_order:
                self._chat_records["contacts_order"].append(None)
            self._chat_records["contacts_order"][contact_order] = contact_name

        self._save_records()
        return {"status": "success", "message": f"已保存联系人 '{contact_name}' 的聊天记录"}

    def get_chat_record(self, contact_name: str) -> Dict:
        """获取单个联系人的聊天记录

        Args:
            contact_name: 联系人名称

        Returns:
            Dict: {"status": "success", "data": {"contact_name": "...", "record": {...}, "exists": true/false}}
        """
        if not contact_name:
            return {"status": "error", "message": "请提供联系人名称"}

        contacts = self._chat_records.get("contacts", {})
        record = contacts.get(contact_name)

        return {
            "status": "success",
            "data": {
                "contact_name": contact_name,
                "record": record,
                "exists": contact_name in contacts
            }
        }

    def update_preview_after_reply(self, contact_name: str, reply_message: str = None) -> Dict:
        """回复完成后更新联系人的预览消息

        重要：在回复消息后调用此方法，记录已回复的内容。
        下次检测时会跳过包含此内容的预览，防止误判。

        Args:
            contact_name: 联系人名称
            reply_message: 回复的消息内容（可选，用于标记已回复）

        Returns:
            Dict: {"status": "success/error", "data": {"preview": "新的预览消息", "time": "时间"}}
        """
        try:
            # 1. 等待微信界面更新（回复后预览需要一点时间更新）
            time.sleep(0.5)

            # 2. 重新获取OCR结果
            ocr_result = self.get_chat_list()
            if ocr_result["status"] != "success":
                return {"status": "error", "message": f"获取OCR失败: {ocr_result.get('message', '')}"}

            ocr_results = ocr_result["data"]["ocr_results"]
            window_rect = ocr_result["data"]["window_rect"]
            window_width = window_rect["width"]

            # 3. 解析联系人列表
            contacts = self.parse_contacts_from_ocr(ocr_results, window_width)

            # 4. 查找目标联系人
            target_contact = None
            for contact in contacts:
                # 支持部分匹配
                if contact_name in contact['name'] or contact['name'] in contact_name:
                    target_contact = contact
                    break

            if not target_contact:
                # 如果找不到，可能联系人已不在当前显示范围
                # 使用回复消息作为标记
                if reply_message:
                    if "contacts" not in self._chat_records:
                        self._chat_records["contacts"] = {}
                    self._chat_records["contacts"][contact_name] = {
                        "last_opponent_message": "",  # 清空，表示已处理
                        "last_opponent_time": "",
                        "last_reply_message": reply_message,  # 记录自己的回复
                        "last_update_time": time.strftime("%Y-%m-%d %H:%M:%S")
                    }
                    self._save_records()
                return {
                    "status": "warning",
                    "message": f"未在当前列表找到联系人 '{contact_name}'，已记录回复内容",
                    "data": {"contact_name": contact_name, "reply_message": reply_message}
                }

            # 5. 获取新的预览消息和时间
            new_preview = target_contact.get('preview', '')
            new_time = target_contact.get('time', '')

            # 更新记录
            if "contacts" not in self._chat_records:
                self._chat_records["contacts"] = {}

            # 保存当前状态，标记这是自己回复后的预览
            # 使用前8字符进行判断（简单可靠）
            preview_prefix = new_preview[:8] if new_preview else ""
            reply_prefix = reply_message[:8] if reply_message else ""

            self._chat_records["contacts"][contact_name] = {
                "last_opponent_message": "",  # 清空对方消息，表示已处理
                "last_opponent_time": new_time,
                "last_preview": preview_prefix,  # 记录当前预览前8字符
                "last_reply_preview": reply_prefix,  # 记录回复内容前8字符
                "last_update_time": time.strftime("%Y-%m-%d %H:%M:%S")
            }

            self._save_records()

            logger.info(f"已更新联系人 '{contact_name}' 的状态，预览: {new_preview}")

            return {
                "status": "success",
                "message": f"已更新联系人 '{contact_name}' 的状态",
                "data": {
                    "contact_name": contact_name,
                    "preview": new_preview,
                    "time": new_time,
                    "reply_message": reply_message
                }
            }

        except Exception as e:
            logger.error(f"更新预览失败: {e}")
            return {"status": "error", "message": f"更新预览失败: {str(e)}"}

    def update_contacts_order(self, contacts_order: list) -> Dict:
        """更新联系人顺序列表

        Args:
            contacts_order: 联系人名称列表，按顺序排列

        Returns:
            Dict: {"status": "success/error", "message": "..."}
        """
        self._chat_records["contacts_order"] = contacts_order
        self._save_records()
        return {"status": "success", "message": f"已更新联系人顺序，共 {len(contacts_order)} 人"}

    def reset_chat_records(self) -> Dict:
        """重置所有聊天记录（删除本地文件）"""
        self._chat_records = {}
        # 删除本地文件
        try:
            state_file = self._get_state_file_path()
            if os.path.exists(state_file):
                os.remove(state_file)
        except Exception as e:
            logger.debug(f"删除聊天记录文件失败: {e}")
        return {"status": "success", "message": "已重置所有聊天记录"}

    # 兼容旧接口
    def get_contact_states(self) -> Dict:
        """获取当前所有联系人的状态（兼容旧接口）"""
        return self.get_chat_records()

    def save_contact_state(self, contact_name: str, state: dict) -> Dict:
        """保存单个联系人的状态（兼容旧接口）"""
        return self.save_chat_record(
            contact_name=contact_name,
            last_opponent_message=state.get("last_message", ""),
            last_opponent_time=state.get("last_time", "")
        )

    def get_contact_state(self, contact_name: str) -> Dict:
        """获取单个联系人的状态（兼容旧接口）"""
        return self.get_chat_record(contact_name=contact_name)

    def reset_contact_states(self) -> Dict:
        """重置所有联系人状态（兼容旧接口）"""
        return self.reset_chat_records()


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

def send_image_to_current(image_path):
    """发送图片"""
    success, err = _manager.send_image(image_path)
    if success:
        return {"status": "success", "message": f"图片发送成功: {image_path}"}
    else:
        return {"status": "error", "message": err}

def screenshot(save_path=None):
    """截图"""
    return _manager.take_screenshot(save_path)

def get_ocr_result(word=None):
    """获取 OCR 结果"""
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
    """获取页面上下文"""
    return _manager.get_page_context()

def get_chat_list(count=5):
    """获取聊天列表"""
    return _manager.get_chat_list(count=count)

def click_contact(position):
    """点击指定坐标"""
    return _manager.click_contact(position=position)

def auto_reply(position, message):
    """自动回复指定位置的联系人"""
    return _manager.auto_reply_to_contact(position=position, message=message)

def check_new_messages():
    """检查是否有新消息"""
    return _manager.check_new_messages()

def verify_chat_window(expected_name):
    """验证当前聊天窗口是否是指定联系人"""
    return _manager.verify_chat_window(expected_name=expected_name)

def get_chat_records():
    """获取所有联系人聊天记录"""
    return _manager.get_chat_records()

def save_chat_record(contact_name, last_opponent_message, last_opponent_time=None, contact_order=None):
    """保存单个联系人聊天记录"""
    return _manager.save_chat_record(
        contact_name=contact_name,
        last_opponent_message=last_opponent_message,
        last_opponent_time=last_opponent_time,
        contact_order=contact_order
    )

def get_chat_record(contact_name):
    """获取单个联系人聊天记录"""
    return _manager.get_chat_record(contact_name=contact_name)

def update_preview_after_reply(contact_name, reply_message=None):
    """回复完成后更新联系人的预览消息（防止误判）"""
    return _manager.update_preview_after_reply(contact_name=contact_name, reply_message=reply_message)

def update_contacts_order(contacts_order):
    """更新联系人顺序"""
    return _manager.update_contacts_order(contacts_order=contacts_order)

def reset_chat_records():
    """重置所有聊天记录"""
    return _manager.reset_chat_records()

# 兼容旧接口
def get_contact_states():
    """获取联系人状态"""
    return _manager.get_contact_states()

def save_contact_state(contact_name, state):
    """保存单个联系人状态"""
    return _manager.save_contact_state(contact_name=contact_name, state=state)

def get_contact_state(contact_name):
    """获取单个联系人状态"""
    return _manager.get_contact_state(contact_name=contact_name)

def reset_contact_states():
    """重置联系人状态"""
    return _manager.reset_contact_states()


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(description="微信自动化工具")
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

    # send_image
    p_send_image = subparsers.add_parser("send_image", help="发送图片")
    p_send_image.add_argument("--path", type=str, required=True, help="图片路径")

    # get_chat_list
    p_chat_list = subparsers.add_parser("get_chat_list", help="获取聊天列表")
    p_chat_list.add_argument("--count", type=int, default=5, help="获取数量")

    # click_contact
    p_click_contact = subparsers.add_parser("click_contact", help="点击指定坐标进入聊天")
    p_click_contact.add_argument("--x", type=int, required=True, help="点击X坐标")
    p_click_contact.add_argument("--y", type=int, required=True, help="点击Y坐标")

    # auto_reply
    p_auto_reply = subparsers.add_parser("auto_reply", help="自动回复指定位置的联系人")
    p_auto_reply.add_argument("--x", type=int, required=True, help="点击X坐标")
    p_auto_reply.add_argument("--y", type=int, required=True, help="点击Y坐标")
    p_auto_reply.add_argument("--message", type=str, required=True, help="回复内容")

    # check_new_messages
    subparsers.add_parser("check_new_messages", help="检查是否有新消息")

    # get_chat_records
    subparsers.add_parser("get_chat_records", help="获取所有联系人聊天记录")

    # save_chat_record
    p_save_record = subparsers.add_parser("save_chat_record", help="保存单个联系人聊天记录")
    p_save_record.add_argument("--name", type=str, required=True, help="联系人名称")
    p_save_record.add_argument("--message", type=str, required=True, help="对方最后一条消息")
    p_save_record.add_argument("--time", type=str, default="", help="消息时间")
    p_save_record.add_argument("--order", type=int, default=None, help="联系人顺序(0-4)")

    # get_chat_record
    p_get_record = subparsers.add_parser("get_chat_record", help="获取单个联系人聊天记录")
    p_get_record.add_argument("--name", type=str, required=True, help="联系人名称")

    # update_preview_after_reply
    p_update_preview = subparsers.add_parser("update_preview_after_reply", help="回复完成后更新预览消息")
    p_update_preview.add_argument("--name", type=str, required=True, help="联系人名称")
    p_update_preview.add_argument("--reply_message", type=str, default=None, help="回复的消息内容")

    # update_contacts_order
    p_update_order = subparsers.add_parser("update_contacts_order", help="更新联系人顺序")
    p_update_order.add_argument("--order", type=str, required=True, help="联系人顺序JSON数组")

    # reset_chat_records
    subparsers.add_parser("reset_chat_records", help="重置所有聊天记录")

    # verify_chat_window
    p_verify = subparsers.add_parser("verify_chat_window", help="验证当前聊天窗口是否是指定联系人")
    p_verify.add_argument("--expected_name", type=str, required=True, help="预期的联系人名称")

    # 兼容旧接口
    # get_contact_states
    subparsers.add_parser("get_contact_states", help="获取联系人状态")

    # save_contact_state
    p_save_state = subparsers.add_parser("save_contact_state", help="保存单个联系人状态")
    p_save_state.add_argument("--name", type=str, required=True, help="联系人名称")
    p_save_state.add_argument("--state", type=str, required=True, help="状态JSON字符串")

    # get_contact_state
    p_get_state = subparsers.add_parser("get_contact_state", help="获取单个联系人状态")
    p_get_state.add_argument("--name", type=str, required=True, help="联系人名称")

    # reset_contact_states
    subparsers.add_parser("reset_contact_states", help="重置联系人状态")

    args = parser.parse_args()

    if args.action is None:
        parser.print_help()
        sys.exit(1)

    result = {"status": "error", "message": "未知操作"}

    if args.action == "screenshot":
        img = _manager.capture()
        if not img:
            result = {"status": "error", "message": "无法获取微信窗口"}
        elif args.save_path:
            img.save(args.save_path)
            result = {"status": "success", "message": f"已保存: {args.save_path}"}
        else:
            buffered = BytesIO()
            img.save(buffered, format="PNG")
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            result = {"status": "success", "data": {"image_base64": img_base64}}
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
    elif args.action == "send_image":
        result = send_image_to_current(args.path)
    elif args.action == "get_chat_list":
        result = get_chat_list(args.count)
    elif args.action == "click_contact":
        position = {"x": args.x, "y": args.y}
        result = click_contact(position=position)
    elif args.action == "auto_reply":
        position = {"x": args.x, "y": args.y}
        result = auto_reply(position=position, message=args.message)
    elif args.action == "check_new_messages":
        result = check_new_messages()
    elif args.action == "get_chat_records":
        result = get_chat_records()
    elif args.action == "save_chat_record":
        result = save_chat_record(
            contact_name=args.name,
            last_opponent_message=args.message,
            last_opponent_time=args.time,
            contact_order=args.order
        )
    elif args.action == "get_chat_record":
        result = get_chat_record(contact_name=args.name)
    elif args.action == "update_preview_after_reply":
        result = update_preview_after_reply(contact_name=args.name, reply_message=args.reply_message)
    elif args.action == "update_contacts_order":
        try:
            order_list = json.loads(args.order)
            result = update_contacts_order(contacts_order=order_list)
        except json.JSONDecodeError:
            result = {"status": "error", "message": "order 参数必须是有效的 JSON 数组"}
    elif args.action == "reset_chat_records":
        result = reset_chat_records()
    elif args.action == "verify_chat_window":
        result = verify_chat_window(expected_name=args.expected_name)
    elif args.action == "get_contact_states":
        result = get_contact_states()
    elif args.action == "save_contact_state":
        try:
            state_dict = json.loads(args.state)
            result = save_contact_state(contact_name=args.name, state=state_dict)
        except json.JSONDecodeError:
            result = {"status": "error", "message": "state 参数必须是有效的 JSON 字符串"}
    elif args.action == "get_contact_state":
        result = get_contact_state(contact_name=args.name)
    elif args.action == "reset_contact_states":
        result = reset_contact_states()

    print(json.dumps(result, ensure_ascii=False, indent=2))

