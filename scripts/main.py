#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
微信自动化技能 - OpenClaw Skill
1. 获取页面截图
2. 获取 OCR 结果（带坐标，字符级别）
3. 点击坐标操作
4. 根据坐标点击后输入内容
5. 上下滚动
6. 获取页面上下文（OCR + UI控件，辅助模型判断）
"""

import sys
import base64
import time
import logging
import os
from typing import Any, Dict, List, Optional
from datetime import datetime
from io import BytesIO

import pyautogui
import pyperclip

from app import autoLoginTB

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("OpenClaw_WeChat_Agent")


class WeChatSkill:
    """精简版微信自动化技能"""
    
    SUPPORTED_ACTIONS = [
        "screenshot",           # 1. 获取页面截图
        "get_ocr_result",       # 2. 获取 OCR 结果
        "click_coordinate",     # 3. 点击坐标操作
        "click_and_type",       # 4. 点击后输入内容
        "scroll",               # 5. 上下滚动
        "get_page_context"      # 6. 获取页面上下文（OCR + UI控件）
    ]

    def __init__(self):
        self.app_controller = autoLoginTB()
        self.window = None
        
        # 配置参数
        self.config = {
            'scroll_amount': 300,       # 滚动量
        }

    def ensure_connection(self) -> bool:
        """确保微信窗口连接"""
        if self.window and self.window.exists():
            return True
        self.window = self.app_controller.start_app()
        return self.window is not None

    def _ensure_window_focused(self) -> bool:
        """
        确保窗口已聚焦到前台
        
        步骤：
        1. 确保窗口连接
        2. 使用 set_focus() 设置焦点
        3. 使用 pyautogui 移动鼠标到窗口（辅助置顶）
        4. 短暂等待窗口激活
        
        返回:
            bool: 是否成功聚焦
        """
        if not self.ensure_connection():
            return False
        
        try:
            # 1. 使用 pywinauto 设置焦点
            try:
                self.window.set_focus()
                logger.info("已聚焦微信窗口")
            except Exception as e:
                logger.debug(f"set_focus 尝试：{e}")
            
            # 2. 获取窗口位置
            rect = self.window.rectangle()
            center_x = rect.left + int(rect.width() * 0.4)
            center_y = rect.top + int(rect.height() * 0.3)

            # 3. 使用 pyautogui 移动鼠标到窗口区域（帮助窗口置顶）
            pyautogui.moveTo(center_x, center_y, duration=0.2)
            time.sleep(0.3)
            
            # 4. 点击一下窗口（确保在前台）
            # 使用左键轻微点击
            pyautogui.click(center_x, center_y, button='left', clicks=1, interval=0.1)
            time.sleep(0.3)
            
            logger.info("窗口已置顶并聚焦")
            return True
            
        except Exception as e:
            logger.warning(f"聚焦窗口失败：{e}")
            # 即使失败也返回 True，因为窗口可能已经在前台
            return True

    def _wake_up_page(self) -> bool:
        """
        唤醒微信页面 - 简化版
        
        策略：
        1. 尝试直接连接窗口，如果成功则点击窗口确保在前台
        2. 如果连接失败，使用 Ctrl+Alt+W 唤醒
        3. 再次尝试连接
        
        返回:
            bool: 是否成功唤醒
        """
        try:

            # 1. 尝试直接连接窗口
            if self.ensure_connection():
                # 窗口已连接，点击确保在前台
                logger.info("微信窗口已连接，确保在前台...")
                
                rect = self.window.rectangle()
                center_x = rect.left + int(rect.width() * 0.3)
                center_y = rect.top + int(rect.height() * 0.3)
                
                pyautogui.moveTo(center_x, center_y, duration=0.3)
                time.sleep(0.2)
                pyautogui.click()
                time.sleep(0.3)
                
                logger.info("微信窗口已置顶")
                return True
            
            # 2. 连接失败，使用 Ctrl+Alt+W 唤醒
            logger.info("使用 Ctrl+Alt+W 唤醒微信...")
            pyautogui.hotkey('ctrl', 'alt', 'w')
            time.sleep(2.0)
            
            # 3. 再次尝试连接
            if self.ensure_connection():
                logger.info("微信已唤醒")
                return True
            
            logger.warning("微信唤醒失败")
            return False
                
        except Exception as e:
            logger.error(f"唤醒页面失败：{e}")
            return False

    def _dispatch_action(self, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        路由分发动作
        
        统一处理窗口唤醒逻辑：
        1. 检测微信是否打开
        2. 未打开则使用 Ctrl+Alt+W 唤醒
        3. 已打开则确保窗口在最上层
        """
        method_name = f"action_{action}"
        if hasattr(self, method_name):
            # 需要窗口交互的动作，先唤醒并聚焦窗口
            actions_need_window = [
                'screenshot', 'get_ocr_result', 'click_coordinate', 
                'click_and_type', 'scroll', 'get_page_context'
            ]
            
            if action in actions_need_window:
                # 检测并唤醒窗口（未打开则 Ctrl+Alt+W，已打开则置顶）
                if not self._wake_up_page():
                    return {"status": "error", "message": "无法唤醒微信窗口，请确保微信已启动并登录"}
                # 确保窗口聚焦在最上层
                if not self._ensure_window_focused():
                    return {"status": "error", "message": "无法聚焦微信窗口"}
            
            return getattr(self, method_name)(params)
        else:
            return {"status": "error", "message": f"不支持的动作：{action}"}

    def execute_action(self, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """执行动作（OpenClaw 调用入口）"""
        if action not in self.SUPPORTED_ACTIONS:
            return {"status": "error", "message": f"不支持的动作：{action}"}
        
        logger.info(f"执行动作：{action}, 参数：{params}")
        return self._dispatch_action(action, params)

    # ==================== 核心功能实现 ====================

    def action_screenshot(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        1. 获取链接程序页面截图
        
        参数:
            save_path: 可选，保存路径，不传则返回 base64
            
        返回:
            {
                "status": "success",
                "data": {
                    "image_base64": "base64 编码的截图",
                    "save_path": "保存路径（如果指定）",
                    "width": 1920,
                    "height": 1080
                }
            }
        """
        try:
            if not self.window:
                return {"status": "error", "message": "微信窗口未连接"}
            
            # 获取窗口截图
            screenshot = self.window.capture_as_image()
            
            # 获取尺寸
            rect = self.window.rectangle()
            width = rect.width()
            height = rect.height()
            
            # 转换为 base64
            buffer = BytesIO()
            screenshot.save(buffer, format='PNG')
            img_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
            result = {
                "status": "success",
                "data": {
                    "image_base64": img_base64,
                    "width": width,
                    "height": height,
                    "timestamp": time.time()
                }
            }
            
            # 如果指定了保存路径，保存图片
            save_path = params.get('save_path')
            if save_path:
                screenshot.save(save_path)
                result["data"]["save_path"] = save_path
                logger.info(f"截图已保存到：{save_path}")
            
            return result
            
        except Exception as e:
            logger.exception("截图失败")
            return {"status": "error", "message": f"截图失败：{str(e)}"}

    def action_get_ocr_result(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        2. 获取当前页面 OCR 结果，带坐标，字符级别
        
        参数:
            keyword: 可选，搜索特定文字
            include_all: 是否返回所有结果（默认 True）
            
        返回:
            {
                "status": "success",
                "data": {
                    "results": [
                        {
                            "text": "识别的文字",
                            "confidence": 0.95,
                            "bbox": [x1, y1, x2, y2],  # 边界框
                            "center": [x, y],           # 中心点坐标
                            "position": "左上 (x1,y1), 右下 (x2,y2)"
                        }
                    ],
                    "count": 10,
                    "search_keyword": "搜索的关键字（如果有）"
                }
            }
        """
        try:
            if not self.window:
                return {"status": "error", "message": "微信窗口未连接"}
            
            # 导入 OCR 引擎
            try:
                from OCR import ocr_endpoint
            except ImportError:
                return {"status": "error", "message": "OCR 引擎未安装"}
            except SystemError:
                return {"status": "error", "message": "OCR 引擎依赖问题：NumPy/ONNXRuntime 版本不兼容。建议运行：pip install numpy<2.0"}
            
            keyword = params.get('keyword')
            include_all = params.get('include_all', True)
            
            # 执行 OCR 识别
            logger.info("开始 OCR 识别...")
            try:
                results = ocr_endpoint(self.window, word=keyword if keyword else None)
            except SystemError as e:
                logger.error(f"OCR 系统错误：{e}")
                return {
                    "status": "error",
                    "message": "OCR 引擎依赖问题：NumPy/ONNXRuntime 版本不兼容",
                    "suggestion": "请运行：pip install numpy<2.0"
                }
            
            if not results:
                return {
                    "status": "success",
                    "data": {
                        "results": [],
                        "count": 0,
                        "message": "未识别到文字"
                    }
                }
            
            # 格式化结果
            formatted_results = []
            for item in results:
                # OCR 返回 box 字段，转换为 bbox
                box = item.get('box', [])
                bbox = [
                    item.get('x_min', 0),
                    item.get('y_min', 0),
                    item.get('x_max', 0),
                    item.get('y_max', 0)
                ] if not box else box
                
                formatted_item = {
                    "text": item.get('text', ''),
                    "confidence": item.get('scores', 0.0),
                    "bbox": bbox,
                    "center": item.get('center', [0, 0]),
                    "position": f"左上 ({bbox[0]},{bbox[1]}), 右下 ({bbox[2]},{bbox[3]})"
                }
                formatted_results.append(formatted_item)
            
            # 如果有关键字，过滤结果
            if keyword:
                filtered_results = [
                    r for r in formatted_results 
                    if keyword.lower() in r['text'].lower()
                ]
                logger.info(f"找到 {len(filtered_results)} 个包含 '{keyword}' 的结果")
                formatted_results = filtered_results
            
            return {
                "status": "success",
                "data": {
                    "results": formatted_results if include_all else formatted_results[:10],
                    "count": len(formatted_results),
                    "total_ocr_count": len(results),
                    "search_keyword": keyword
                }
            }
            
        except SystemError as e:
            error_msg = str(e)
            logger.error(f"OCR 系统错误：{error_msg}")
            return {
                "status": "error",
                "message": "OCR 引擎依赖问题：NumPy/ONNXRuntime 版本不兼容",
                "suggestion": "请运行：pip install numpy<2.0"
            }
        except Exception as e:
            logger.exception("OCR 识别失败")
            return {"status": "error", "message": f"OCR 识别失败：{str(e)}"}

    def action_click_coordinate(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        3. 点击坐标操作
        
        参数:
            x: X 坐标（必需）
            y: Y 坐标（必需）
            clicks: 点击次数（默认 1）
            button: 鼠标按钮（'left', 'right', 'middle'，默认 'left'）
            interval: 多次点击间隔（默认 0.1）
            
        返回:
            {
                "status": "success",
                "message": "点击成功",
                "data": {
                    "position": [x, y],
                    "clicks": 1,
                    "button": "left"
                }
            }
        """
        try:
            x = params.get('x')
            y = params.get('y')
            
            if x is None or y is None:
                return {"status": "error", "message": "缺少必需参数：x, y"}
            
            clicks = params.get('clicks', 1)
            button = params.get('button', 'left')
            interval = params.get('interval', 0.1)
            
            # 执行点击
            logger.info(f"点击坐标：({x}, {y}), 次数：{clicks}, 按钮：{button}")
            pyautogui.click(x=x, y=y, clicks=clicks, button=button, interval=interval)
            
            return {
                "status": "success",
                "message": f"成功点击坐标 ({x}, {y})",
                "data": {
                    "position": [x, y],
                    "clicks": clicks,
                    "button": button
                }
            }
            
        except Exception as e:
            logger.exception("点击失败")
            return {"status": "error", "message": f"点击失败：{str(e)}"}

    def _get_input_box_position(self, offset_x: float = 0.4, offset_y: float = 0.95) -> tuple:
        """
        获取输入框的位置（通用方法）
        
        策略：
        1. 尝试 OCR 识别关键词（如"发送"），输入框在其附近
        2. 使用窗口相对位置作为备选
        
        参数:
            offset_x: X 方向相对位置（0.0-1.0），默认 0.4（窗口宽度 40% 处）
            offset_y: Y 方向相对位置（0.0-1.0），默认 0.95（窗口底部 95% 处）
        
        返回:
            tuple: (x, y) 输入框中心坐标，失败返回 (None, None)
        """
        try:
            time.sleep(0.2)  # 短暂等待
            
            # 策略 1：尝试 OCR 识别关键词
            try:
                from OCR import ocr_endpoint
                # 可配置要识别的关键词
                keywords = ["发送", "Send", "回复", "Reply"]
                
                for keyword in keywords:
                    results = ocr_endpoint(self.window, word=keyword)
                    if results and len(results) > 0:
                        # 找到关键词，输入框通常在其附近
                        target = results[0]
                        bbox = target.get('bbox', [])
                        if len(bbox) >= 4:
                            x1, y1, x2, y2 = bbox
                            # 输入框在关键词左边或下方（根据具体 UI 调整）
                            input_x = x1 - 200  # 可配置偏移量
                            input_y = (y1 + y2) // 2
                            logger.info(f"通过'{keyword}'定位输入框：({input_x}, {input_y})")
                            return input_x, input_y
            except Exception as e:
                logger.debug(f"OCR 策略失败：{e}")
            
            # 策略 2：使用窗口相对位置
            if self.window:
                rect = self.window.rectangle()
                window_width = rect.width()
                window_height = rect.height()
                
                # 使用相对位置（可配置）
                input_x = rect.left + int(window_width * offset_x)
                input_y = rect.top + int(window_height * offset_y)
                
                logger.info(f"使用相对位置定位输入框：({input_x}, {input_y})")
                return input_x, input_y
            
            return None, None
            
        except Exception as e:
            logger.error(f"获取输入框位置失败：{e}")
            return None, None

    def action_click_and_type(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        4. 根据坐标点击后输入内容
        
        参数:
            x: X 坐标（可选，如果提供则先点击）
            y: Y 坐标（可选，如果提供则先点击）
            content: 要输入的内容（必需）
            clear_before: 输入前是否清空（默认 True）
            send_enter: 输入后是否按回车（默认 False）
            interval: 字符输入间隔（默认 0.05 秒）
            auto_locate_input: 是否自动定位输入框（默认 False，为 True 时忽略 x,y 参数）
            
        返回:
            {
                "status": "success",
                "message": "输入成功",
                "data": {
                    "content_length": 10,
                    "clicked_position": [x, y],
                    "content_preview": "输入内容的前 20 个字符"
                }
            }
        """
        try:
            x = params.get('x')
            y = params.get('y')
            content = params.get('content')
            auto_locate_input = params.get('auto_locate_input', False)
            
            if not content:
                return {"status": "error", "message": "缺少必需参数：content"}
            
            # 如果需要自动定位输入框
            if auto_locate_input:
                logger.info("自动定位输入框...")
                input_x, input_y = self._get_input_box_position()
                if input_x is not None and input_y is not None:
                    x, y = input_x, input_y
                    logger.info(f"输入框定位成功：({x}, {y})")
                else:
                    return {"status": "error", "message": "无法定位输入框"}
            
            # 如果需要，先点击坐标
            if x is not None and y is not None:
                logger.info(f"先点击坐标：({x}, {y})")
                pyautogui.click(x=x, y=y, clicks=1)
                time.sleep(0.3)  # 等待点击生效
            
            # 清空输入框（如果需要）
            clear_before = params.get('clear_before', True)
            if clear_before:
                logger.info("清空输入框")
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.press('delete')
                time.sleep(0.1)
            
            # 输入内容
            logger.info(f"输入内容：{content[:50]}{'...' if len(content) > 50 else ''}")

            # 使用剪贴板输入（支持中文）
            pyperclip.copy(content)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            
            # 按回车（如果需要）
            send_enter = params.get('send_enter', False)
            if send_enter:
                logger.info("按回车键发送")
                pyautogui.press('enter')
            
            return {
                "status": "success",
                "message": "输入成功",
                "data": {
                    "content_length": len(content),
                    "clicked_position": [x, y] if x and y else None,
                    "content_preview": content[:20] + ('...' if len(content) > 20 else ''),
                    "send_enter": send_enter
                }
            }
            
        except Exception as e:
            logger.exception("输入失败")
            return {"status": "error", "message": f"输入失败：{str(e)}"}

    def action_scroll(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        5. 上下滚动
        
        参数:
            direction: 方向（'up' 或 'down'，默认 'down'）
            amount: 滚动量（默认 300）
            clicks: 滚动次数（默认 1）
            
        返回:
            {
                "status": "success",
                "message": "滚动成功",
                "data": {
                    "direction": "down",
                    "amount": 300,
                    "clicks": 1
                }
            }
        """
        try:
            direction = params.get('direction', 'down')
            amount = params.get('amount', self.config['scroll_amount'])
            clicks = params.get('clicks', 1)
            
            # 计算实际滚动量
            scroll_amount = amount * clicks
            if direction == 'up':
                scroll_amount = -scroll_amount
            
            logger.info(f"滚动：方向={direction}, 量={scroll_amount}")
            
            # 执行滚动
            pyautogui.scroll(scroll_amount)
            
            return {
                "status": "success",
                "message": f"已{direction}滚动 {scroll_amount} 单位",
                "data": {
                    "direction": direction,
                    "amount": amount,
                    "clicks": clicks,
                    "total_scroll": abs(scroll_amount)
                }
            }
            
        except Exception as e:
            logger.exception("滚动失败")
            return {"status": "error", "message": f"滚动失败：{str(e)}"}

    def action_get_page_context(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        6. 获取页面上下文（OCR + UI控件）
        
        将页面文字和UI控件信息一起返回，辅助模型判断操作位置
        
        参数:
            include_ocr: 是否包含OCR结果（默认 True）
            include_controls: 是否包含控件信息（默认 True）
            format_for_model: 是否格式化为适合模型的文本（默认 False）
            max_items: 每类最大返回数量（默认 50）
            
        返回:
            {
                "status": "success",
                "data": {
                    "window_info": {"title": "...", "width": 800, "height": 600},
                    "ocr_results": [...],           # OCR识别结果
                    "controls": [...],               # 全部控件
                    "interactive_controls": [...],   # 可交互控件
                    "formatted_text": "..."          # 格式化的文本（如果启用）
                }
            }
        """
        try:
            if not self.window:
                return {"status": "error", "message": "微信窗口未连接"}
            
            # 导入 UI 检测模块
            try:
                from ui_inspector import get_page_context, format_context_for_model
            except ImportError:
                return {"status": "error", "message": "UI检测模块未安装"}
            
            # 获取参数
            include_ocr = params.get('include_ocr', True)
            include_controls = params.get('include_controls', True)
            format_for_model = params.get('format_for_model', False)
            max_items = params.get('max_items', 50)
            
            # 获取页面上下文
            logger.info("获取页面上下文...")
            context = get_page_context(
                self.window, 
                include_ocr=include_ocr, 
                include_controls=include_controls
            )
            
            # 统计信息
            ocr_count = len(context.get('ocr_results', []))
            ctrl_count = len(context.get('controls', []))
            interactive_count = len(context.get('interactive_controls', []))
            
            logger.info(f"页面上下文: OCR={ocr_count}, 控件={ctrl_count}, 可交互={interactive_count}")
            
            result = {
                "status": "success",
                "message": f"获取成功: {ocr_count}个文字区域, {ctrl_count}个控件, {interactive_count}个可交互元素",
                "data": context
            }
            
            # 格式化为适合模型的文本
            if format_for_model:
                result["data"]["formatted_text"] = format_context_for_model(context, max_items)
            
            return result
            
        except Exception as e:
            logger.exception("获取页面上下文失败")
            return {"status": "error", "message": f"获取页面上下文失败：{str(e)}"}


# OpenClaw 调用入口
def execute_action(action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    """OpenClaw 调用入口函数"""
    skill = WeChatSkill()
    return skill.execute_action(action, params)


if __name__ == "__main__":
    # 测试代码
    print("微信自动化精简版 - 核心功能")
    print(f"支持的动作：{WeChatSkill.SUPPORTED_ACTIONS}")
