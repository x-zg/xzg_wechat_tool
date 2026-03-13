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
import ctypes

# 获取 user32 API
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
    STATE_FILE = "wechat_contact_states.json"  # 状态持久化文件
    
    def __init__(self):
        self._app = None
        self._window = None
        self._gw_window = None  # pygetwindow 窗口
        self._contact_states = {}  # 记录每个联系人的状态 {name: {"last_message": "...", "replied": True}}
        self._load_states()  # 启动时加载状态
    
    def _get_state_file_path(self) -> str:
        """获取状态文件路径"""
        # 状态文件保存在agent.py所在目录
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, self.STATE_FILE)
    
    def _load_states(self):
        """从文件加载联系人状态"""
        try:
            state_file = self._get_state_file_path()
            if os.path.exists(state_file):
                with open(state_file, 'r', encoding='utf-8') as f:
                    self._contact_states = json.load(f)
        except Exception as e:
            logger.debug(f"加载联系人状态失败: {e}")
            self._contact_states = {}
    
    def _save_states(self):
        """保存联系人状态到文件"""
        try:
            state_file = self._get_state_file_path()
            with open(state_file, 'w', encoding='utf-8') as f:
                json.dump(self._contact_states, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.debug(f"保存联系人状态失败: {e}")

    def _find_gw_window(self) -> Optional[Any]:
        """使用 pygetwindow 查找微信窗口"""
        # 微信窗口标题通常是：微信、WeChat、Weixin，或者以这些开头的标题
        # 排除包含 "agent.py" 或当前脚本路径的窗口
        script_name = os.path.basename(__file__).lower()

        for keyword in ["微信", "WeChat", "Weixin"]:
            windows = gw.getWindowsWithTitle(keyword)
            for w in windows:
                title_lower = w.title.lower()
                # 排除 IDE/编辑器窗口（包含脚本名或路径）
                if script_name in title_lower or "agent.py" in title_lower:
                    logger.debug(f"跳过 IDE 窗口: title={w.title}")
                    continue
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

        # 4. 点击输入框
        input_x = rect["left"] + rect["width"] // 2
        input_y = rect["bottom"] - 60
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

            # 截取聊天列表区域
            img = self.capture()
            if not img:
                logger.error("截图失败")
                return {"status": "error", "message": "截图失败"}

            # 左侧聊天列表区域（相对于图像的坐标）
            # 图像坐标从 (0, 0) 开始
            chat_list_width = 280
            crop_left = 0  # 图像左边缘
            crop_top = 60  # 跳过搜索框
            crop_right = chat_list_width
            crop_bottom = img.height  # 图像底部

            # 裁剪左侧聊天列表区域
            chat_list_img = img.crop((crop_left, crop_top, crop_right, crop_bottom))

            # 保存窗口位置用于后续计算点击坐标
            self._window_rect = rect

            # OCR 识别（直接使用 numpy 数组）
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
            result, _ = ocr(np.array(chat_list_img))

            if not result:
                logger.error("OCR 未识别到内容")
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

                # 按 Y 坐标排序（上面的先处理）
                group.sort(key=lambda l: l["y"])

                # 微信聊天列表布局：每组第一个是名称，后面是消息
                # 例如：['小许', '天气'] -> 名称='小许', 消息='天气'
                name_text = None
                message_text = None
                is_from_me = False

                # 第一个（Y 坐标最小）是名称
                name_text = group[0]["text"].strip()

                # 其余的是消息，取最后一个作为最新消息
                if len(group) > 1:
                    message_text = group[-1]["text"].strip()

                # ===== 改进的 is_from_me 判断 =====
                # 微信聊天列表可能显示的格式：
                # 1. "我: 消息内容" / "我：消息内容"（中文/英文冒号）
                # 2. "[我] 消息内容"（部分版本）
                # 3. "我 消息内容"（无冒号，少见）
                # 注意：如果消息内容本身包含"我"字（如"我们"），不要误判
                if message_text:
                    # 方法1：检测标准前缀
                    prefixes = ["我:", "我：", "[我]", "【我】"]
                    for prefix in prefixes:
                        if message_text.startswith(prefix):
                            is_from_me = True
                            message_text = message_text[len(prefix):].strip()
                            logger.debug(f"检测到 '{prefix}' 前缀，is_from_me=True")
                            break

                    # 方法2：如果名称和消息中的发送者名称一致，可能是群消息中我发的
                    # 例如：群名"工作群"，消息显示"我: 好的" -> is_from_me=True
                    # 但这种情况上面已经处理了

                    # 方法3：检测消息开头是否包含联系人名称（群消息格式）
                    # 例如："许志国：@老三" 可能表示群消息，发送者是"许志国"
                    # 这种情况不标记为 is_from_me，因为需要和当前用户名对比
                    # 由于我们不知道当前登录的微信用户名，暂时无法准确判断

                # 计算点击位置（转换回屏幕坐标）
                if group and hasattr(self, '_window_rect'):
                    avg_y = sum(l["y"] for l in group) / len(group)
                    # 转换回屏幕坐标：窗口左边 + 裁剪偏移 + 相对坐标
                    click_x = self._window_rect["left"] + 140  # 列表中间位置
                    click_y = self._window_rect["top"] + 60 + avg_y + 30  # 加上裁剪偏移

                    contacts.append({
                        "index": i,
                        "name": name_text or f"未知联系人{i+1}",
                        "last_message": message_text or "",
                        "is_from_me": is_from_me,  # 新增：标记消息来源
                        "position": {"x": int(click_x), "y": int(click_y)},
                        "rect": {
                            "left": self._window_rect["left"],
                            "top": int(self._window_rect["top"] + 60 + avg_y - 20),
                            "right": self._window_rect["left"] + 280,
                            "bottom": int(self._window_rect["top"] + 60 + avg_y + 50)
                        }
                    })

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
            # 微信聊天窗口的发送按钮位置：右侧底部
            send_button_region = img.crop((
                rect["width"] // 2,  # 左边界：窗口中间
                rect["height"] - 150,  # 上边界：底部150像素以上
                rect["width"],  # 右边界：窗口右边
                rect["height"] - 30  # 下边界：底部30像素以上
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

    def click_contact(self, contact_name: str = None, position: dict = None) -> Dict:
        """点击联系人进入聊天

        Args:
            contact_name: 联系人名称（需要先通过 get_chat_list 获取）
            position: 点击位置 {"x": 100, "y": 200}

        Returns:
            Dict: {"status": "success/error", "message": "...", "matched_name": "匹配到的联系人名称"}
        """
        try:
            # 确保窗口就绪
            w, error = self._ensure_window_ready()
            if error:
                return {"status": "error", "message": error}

            # 获取点击位置
            click_x, click_y = None, None
            contact_display_name = None
            
            if position and "x" in position and "y" in position:
                click_x = position["x"]
                click_y = position["y"]
            elif contact_name:
                # 获取聊天列表查找位置
                result = self.get_chat_list()
                if result["status"] != "success":
                    return result

                # ===== 改进的匹配逻辑：精确匹配优先，模糊匹配备选 =====
                matched_contact = None
                
                # 优先：精确匹配
                for contact in result["data"]["contacts"]:
                    if contact["name"] == contact_name:
                        matched_contact = contact
                        break
                
                # 备选：模糊匹配（但要求名称包含关系是双向的或目标较短）
                if not matched_contact:
                    for contact in result["data"]["contacts"]:
                        # 模糊匹配：联系人名称包含目标，或目标包含联系人名称
                        if contact_name in contact["name"] or contact["name"] in contact_name:
                            # 额外检查：避免误匹配（如"小许"匹配到"小许2"）
                            # 只有当目标名称较短时才允许包含匹配
                            if len(contact_name) <= len(contact["name"]):
                                matched_contact = contact
                                logger.debug(f"模糊匹配: '{contact_name}' -> '{contact['name']}'")
                                break
                
                if matched_contact:
                    click_x = matched_contact["position"]["x"]
                    click_y = matched_contact["position"]["y"]
                    contact_display_name = matched_contact["name"]
                else:
                    return {"status": "error", "message": f"未找到联系人: {contact_name}"}
            else:
                return {"status": "error", "message": "请提供 contact_name 或 position 参数"}

            # ===== 点击联系人并验证窗口状态（关键！）=====
            # 微信联系人列表有开关切换行为：
            # - 如果不在聊天窗口，点击会打开
            # - 如果已在聊天窗口，点击会关闭
            # 因此需要验证并处理
            
            max_attempts = 2  # 最多尝试2次点击
            
            for attempt in range(max_attempts):
                # 点击联系人
                pyautogui.click(click_x, click_y)
                time.sleep(0.5)
                
                # 验证聊天窗口是否打开
                if self._verify_chat_window_open():
                    # 窗口已打开，成功
                    return {
                        "status": "success", 
                        "message": f"已进入与 {contact_display_name} 的聊天窗口",
                        "matched_name": contact_display_name  # 返回匹配到的名称
                    }
                else:
                    # 窗口未打开，可能是因为之前已经在这个窗口，点击导致关闭了
                    # 如果不是最后一次尝试，继续点击重新打开
                    if attempt < max_attempts - 1:
                        time.sleep(0.3)  # 等待一下再重试
                        continue
            
            # 所有尝试都失败
            return {"status": "error", "message": "无法打开聊天窗口，请检查微信状态"}

        except Exception as e:
            logger.error(f"点击联系人失败: {e}")
            return {"status": "error", "message": f"点击联系人失败: {str(e)}"}

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

            # 获取实际匹配到的联系人名称
            matched_name = click_result.get("matched_name", contact_name)

            time.sleep(0.3)

            # 2. 发送消息
            success, err = self.send_message(message)
            if success:
                # ===== 关键：发送成功后，立即更新所有"我发的消息"的联系人状态 =====
                # 这样可以防止发错人后，错误目标被检测为"有新消息"
                time.sleep(0.5)  # 等待微信更新聊天列表
                
                # 重新获取聊天列表
                list_result = self.get_chat_list()
                if list_result["status"] == "success":
                    for contact in list_result["data"]["contacts"]:
                        name = contact["name"]
                        # 如果最后一条消息是我发的，标记为已回复
                        if contact.get("is_from_me", False):
                            if name in self._contact_states:
                                self._contact_states[name]["replied"] = True
                                self._contact_states[name]["last_message"] = contact.get("last_message", "")
                                self._contact_states[name]["is_from_me"] = True
                                logger.debug(f"已标记 {name} 为已回复（我发的消息）")
                            else:
                                # 如果联系人不在状态中，添加并标记为已回复
                                self._contact_states[name] = {
                                    "last_message": contact.get("last_message", ""),
                                    "is_from_me": True,
                                    "replied": True
                                }
                                logger.debug(f"添加并标记 {name} 为已回复")
                    self._save_states()  # 保存状态
                
                return {"status": "success", "message": f"已回复 {matched_name}: {message}"}
            else:
                return {"status": "error", "message": f"发送失败: {err}"}

        except Exception as e:
            logger.error(f"自动回复失败: {e}")
            return {"status": "error", "message": f"自动回复失败: {str(e)}"}

    def check_new_messages(self) -> Dict:
        """检查是否有新消息（单次检测，不自动回复）

        Returns:
            Dict: {
                "status": "success",
                "data": {
                    "has_new_messages": True/False,
                    "new_messages": [...],
                    "all_contacts": [...]
                }
            }
        """
        try:
            # 获取当前聊天列表状态
            current_result = self.get_chat_list()
            if current_result["status"] != "success":
                return current_result

            current_contacts = current_result["data"]["contacts"]
            new_messages = []

            # 检查每个联系人是否有新消息
            has_changes = False

            # 初始化所有联系人的状态（首次运行）
            for contact in current_contacts:
                name = contact["name"]
                if name not in self._contact_states:
                    self._contact_states[name] = {
                        "last_message": contact.get("last_message", ""),
                        "is_from_me": contact.get("is_from_me", False),
                        "replied": True  # 首次见到，标记为已回复
                    }
                    has_changes = True  # 初始化后需要保存状态
            for contact in current_contacts:
                name = contact["name"]
                current_message = contact.get("last_message", "")
                current_is_from_me = contact.get("is_from_me", False)
                saved_state = self._contact_states.get(name, {})
                saved_message = saved_state.get("last_message", "")
                saved_is_from_me = saved_state.get("is_from_me", False)
                already_replied = saved_state.get("replied", False)

                # ===== 关键改进：检查是否手动回复 =====
                # 如果当前最后一条消息来自我，说明已经回复了（手动或自动）
                # 无论之前的状态如何，都应该标记为已回复
                if current_is_from_me and not already_replied:
                    # 检测到手动回复
                    self._contact_states[name]["replied"] = True
                    has_changes = True
                    # 不加入 new_messages，因为已经回复了
                    # 更新消息内容
                    self._contact_states[name]["last_message"] = current_message
                    self._contact_states[name]["is_from_me"] = current_is_from_me
                    continue

                # 检测消息是否变化
                if current_message != saved_message or current_is_from_me != saved_is_from_me:
                    # 更新保存的消息
                    self._contact_states[name]["last_message"] = current_message
                    self._contact_states[name]["is_from_me"] = current_is_from_me
                    has_changes = True

                    # 只有对方发的新消息才需要回复
                    if not current_is_from_me:
                        self._contact_states[name]["replied"] = False
                        new_messages.append({
                            "contact": name,
                            "message": current_message,
                            "position": contact.get("position"),
                            "is_from_me": False
                        })
                    else:
                        # 当前消息来自我，标记为已回复
                        self._contact_states[name]["replied"] = True

                elif not already_replied and not current_is_from_me:
                    # 消息没有变化，但之前标记为未回复，且当前消息不是来自我
                    # 这种情况可能是：对方发了消息，但还没有回复
                    new_messages.append({
                        "contact": name,
                        "message": current_message,
                        "position": contact.get("position"),
                        "is_from_me": False
                    })

            # 如果有状态变化，保存到文件
            if has_changes:
                self._save_states()

            return {
                "status": "success",
                "data": {
                    "has_new_messages": len(new_messages) > 0,
                    "new_messages": new_messages,
                    "all_contacts": current_contacts,
                    "total_new": len(new_messages)
                }
            }

        except Exception as e:
            logger.error(f"检查新消息失败: {e}")
            return {"status": "error", "message": f"检查新消息失败: {str(e)}"}

    def get_contact_states(self) -> Dict:
        """获取当前所有联系人的状态"""
        return {
            "status": "success",
            "data": {
                "contact_states": self._contact_states,
                "total": len(self._contact_states)
            }
        }

    def reset_contact_states(self) -> Dict:
        """重置所有联系人状态"""
        self._contact_states = {}
        self._save_states()  # 保存空状态
        return {"status": "success", "message": "已重置所有联系人状态"}


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

def click_contact(contact_name=None, position=None):
    """点击联系人"""
    return _manager.click_contact(contact_name=contact_name, position=position)

def auto_reply(contact_name, message):
    """自动回复联系人"""
    return _manager.auto_reply_to_contact(contact_name=contact_name, message=message)

def check_new_messages():
    """检查是否有新消息"""
    return _manager.check_new_messages()

def get_contact_states():
    """获取联系人状态"""
    return _manager.get_contact_states()

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
    p_click_contact = subparsers.add_parser("click_contact", help="点击联系人")
    p_click_contact.add_argument("--name", type=str, default=None, help="联系人名称")
    p_click_contact.add_argument("--x", type=int, default=None, help="点击X坐标")
    p_click_contact.add_argument("--y", type=int, default=None, help="点击Y坐标")

    # auto_reply
    p_auto_reply = subparsers.add_parser("auto_reply", help="自动回复联系人")
    p_auto_reply.add_argument("--name", type=str, required=True, help="联系人名称")
    p_auto_reply.add_argument("--message", type=str, required=True, help="回复内容")

    # check_new_messages
    subparsers.add_parser("check_new_messages", help="检查是否有新消息")

    # get_contact_states
    subparsers.add_parser("get_contact_states", help="获取联系人状态")

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
        position = {"x": args.x, "y": args.y} if args.x and args.y else None
        result = click_contact(contact_name=args.name, position=position)
    elif args.action == "auto_reply":
        result = auto_reply(contact_name=args.name, message=args.message)
    elif args.action == "check_new_messages":
        result = check_new_messages()
    elif args.action == "get_contact_states":
        result = get_contact_states()
    elif args.action == "reset_contact_states":
        result = reset_contact_states()

    print(json.dumps(result, ensure_ascii=False, indent=2))

