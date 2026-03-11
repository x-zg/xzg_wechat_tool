#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""微信自动化技能 - OpenClaw Skill"""

import sys
import base64
import time
import logging
from typing import Any, Dict
from io import BytesIO

import pyperclip

try:
    from openclaw.sdk import BaseSkill, Context
    OPENCLAW_AVAILABLE = True
except ImportError:
    OPENCLAW_AVAILABLE = False
    BaseSkill = object
    Context = None

from app import autoLoginTB

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class WeChatSkill(BaseSkill):
    """微信自动化技能类"""
    
    SUPPORTED_ACTIONS = [
        "screenshot",
        "get_ocr_result",
        "click_coordinate",
        "click_and_type",
        "scroll",
        "get_page_context"
    ]
    
    def __init__(self):
        if OPENCLAW_AVAILABLE:
            super().__init__()
        self.app_controller = autoLoginTB()
        self.window = None
        
    def execute(self, context: Context) -> Dict[str, Any]:
        """OpenClaw 标准执行方法"""
        if not OPENCLAW_AVAILABLE:
            return {"status": "error", "message": "OpenClaw SDK 未安装"}
        
        try:
            action = context.get_param('action')
            if not action:
                return {"status": "error", "message": "缺少 action 参数"}
            
            logger.info(f"执行操作：{action}")
            
            # 调用对应方法
            if action == 'screenshot':
                return self._screenshot(context)
            elif action == 'get_ocr_result':
                return self._ocr(context)
            elif action == 'click_coordinate':
                return self._click(context)
            elif action == 'click_and_type':
                return self._type(context)
            elif action == 'scroll':
                return self._scroll(context)
            elif action == 'get_page_context':
                return self._page_context(context)
            else:
                return {"status": "error", "message": f"不支持的操作：{action}"}
                
        except Exception as e:
            logger.error(f"执行异常：{e}")
            return {"status": "error", "message": str(e)}
    
    def _screenshot(self, context):
        import pyautogui
        save_path = context.get_param('save_path')
        screenshot = pyautogui.screenshot()
        if save_path:
            screenshot.save(save_path)
            return {"status": "success", "message": f"已保存：{save_path}"}
        else:
            buffered = BytesIO()
            screenshot.save(buffered, format="PNG")
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            return {"status": "success", "data": {"image_base64": img_base64}}
    
    def _ocr(self, context):
        from OCR import OCRHelper
        import pyautogui
        ocr_helper = OCRHelper()
        screenshot = pyautogui.screenshot()
        results = ocr_helper.recognize(screenshot)
        return {"status": "success", "data": {"results": results, "count": len(results)}}
    
    def _click(self, context):
        import pyautogui
        x = context.get_param('x')
        y = context.get_param('y')
        if not x or not y:
            return {"status": "error", "message": "缺少 x 或 y 参数"}
        pyautogui.click(int(x), int(y))
        return {"status": "success", "message": f"点击 ({x}, {y})"}
    
    def _type(self, context):
        content = context.get_param('content')
        if not content:
            return {"status": "error", "message": "缺少 content 参数"}
        pyperclip.copy(content)
        import pyautogui
        pyautogui.hotkey('ctrl', 'v')
        if context.get_param('send_enter'):
            pyautogui.press('enter')
        return {"status": "success", "message": "输入成功"}
    
    def _scroll(self, context):
        import pyautogui
        direction = context.get_param('direction', 'down')
        amount = context.get_param('amount', 300)
        if direction == 'down':
            pyautogui.scroll(-amount)
        else:
            pyautogui.scroll(amount)
        return {"status": "success", "message": "滚动成功"}
    
    def _page_context(self, context):
        from OCR import OCRHelper
        import pyautogui
        ocr_helper = OCRHelper()
        screenshot = pyautogui.screenshot()
        ocr_results = ocr_helper.recognize(screenshot)
        return {"status": "success", "data": {"ocr_results": ocr_results}}


def execute_action(action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    """独立模式执行函数"""
    skill = WeChatSkill()
    
    class MockContext:
        def __init__(self, action, params):
            self._action = action
            self._params = params or {}
        def get_param(self, key, default=None):
            if key == 'action':
                return self._action
            return self._params.get(key, default)
    
    context = MockContext(action, params)
    return skill.execute(context)


if __name__ == "__main__":
    print(f"OpenClaw SDK: {'已安装' if OPENCLAW_AVAILABLE else '未安装'}")
    print(f"支持的操作：{', '.join(SUPPORTED_ACTIONS)}")
