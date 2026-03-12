#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信自动化工具测试文件
测试所有功能模块
"""
import sys
import time
import json
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

from agent import WeChatManager


def print_result(name, result):
    """打印测试结果"""
    print(f"\n{'='*50}")
    print(f"【{name}】")
    print(f"{'='*50}")
    if isinstance(result, dict):
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(result)


class TestWeChatManager:
    """微信管理器测试类"""
    
    def __init__(self):
        self.manager = WeChatManager()
    
    def test_check_process(self):
        """测试1: 检查微信进程"""
        print("\n[测试1] 检查微信进程")
        pid_list = self.manager.check_process()
        result = {
            "status": "success" if pid_list else "warning",
            "pid_list": pid_list,
            "message": f"找到 {len(pid_list)} 个微信进程" if pid_list else "微信未运行"
        }
        print_result("检查微信进程", result)
        return result
    
    def test_get_main_window(self):
        """测试2: 获取微信主窗口"""
        print("\n[测试2] 获取微信主窗口")
        win = self.manager.get_main_window(activate_first=True)
        if win:
            result = {
                "status": "success",
                "title": getattr(win, 'title', '未知'),
                "size": f"{getattr(win, 'width', '?')}x{getattr(win, 'height', '?')}",
                "message": "成功获取窗口"
            }
        else:
            result = {"status": "error", "message": "获取窗口失败"}
        print_result("获取微信主窗口", result)
        return result
    
    def test_get_window_rect(self):
        """测试3: 获取窗口位置"""
        print("\n[测试3] 获取窗口位置")
        rect = self.manager.get_window_rect()
        if rect:
            result = {
                "status": "success",
                "position": f"({rect['left']}, {rect['top']})",
                "size": f"{rect['width']}x{rect['height']}",
                "rect": rect
            }
        else:
            result = {"status": "error", "message": "获取窗口位置失败"}
        print_result("获取窗口位置", result)
        return result
    
    def test_get_status(self):
        """测试4: 获取微信状态"""
        print("\n[测试4] 获取微信状态")
        status = self.manager.get_status()
        print_result("获取微信状态", status)
        return status
    
    def test_capture(self):
        """测试5: 截图测试"""
        print("\n[测试5] 截图测试")
        img = self.manager.capture()
        if img:
            save_path = str(Path(__file__).parent / "test_screenshot.png")
            img.save(save_path)
            result = {
                "status": "success",
                "size": f"{img.width}x{img.height}",
                "save_path": save_path,
                "message": "截图成功"
            }
        else:
            result = {"status": "error", "message": "截图失败"}
        print_result("截图测试", result)
        return result
    
    def test_take_screenshot(self):
        """测试6: 截图保存测试"""
        print("\n[测试6] 截图保存测试")
        save_path = str(Path(__file__).parent / "test_screenshot2.png")
        result = self.manager.take_screenshot(save_path)
        print_result("截图保存测试", result)
        return result
    
    def test_ocr_result(self):
        """测试7: OCR识别测试"""
        print("\n[测试7] OCR识别测试")
        result = self.manager.get_ocr_result()
        print_result("OCR识别测试", result)
        
        # 显示部分OCR结果
        if result.get("status") == "success" and result.get("data"):
            items = result["data"].get("results", [])
            print(f"\n识别到 {len(items)} 个文本区域:")
            for i, item in enumerate(items[:10]):
                print(f"  [{i+1}] {item.get('text', '')}")
        return result
    
    def test_ocr_with_keyword(self):
        """测试8: OCR关键词搜索测试"""
        print("\n[测试8] OCR关键词搜索测试")
        # 搜索常见词
        result = self.manager.get_ocr_result(word="微信")
        print_result("OCR关键词搜索测试", result)
        return result
    
    def test_get_page_context(self):
        """测试9: 获取页面上下文"""
        print("\n[测试9] 获取页面上下文")
        print("正在获取最新截图...")
        result = self.manager.get_page_context()
        
        # 显示截图时间
        if result.get("status") == "success":
            from datetime import datetime
            print(f"截图时间: {datetime.now().strftime('%H:%M:%S')}")
        
        print_result("获取页面上下文", result)
        return result
    
    def test_click(self):
        """测试10: 点击测试"""
        print("\n[测试10] 点击测试")
        rect = self.manager.get_window_rect()
        if rect:
            # 点击窗口中心
            x = rect["left"] + rect["width"] // 2
            y = rect["top"] + rect["height"] // 2
            success = self.manager.click(x, y)
            result = {
                "status": "success" if success else "error",
                "click_position": f"({x}, {y})",
                "message": "点击成功" if success else "点击失败"
            }
        else:
            result = {"status": "error", "message": "无法获取窗口位置"}
        print_result("点击测试", result)
        return result
    
    def test_scroll(self):
        """测试11: 滚动测试"""
        print("\n[测试11] 滚动测试")
        success = self.manager.scroll(direction='down', amount=200)
        result = {
            "status": "success" if success else "error",
            "message": "向下滚动成功" if success else "滚动失败"
        }
        print_result("滚动测试", result)
        return result
    
    def test_input_text(self):
        """测试12: 输入文本测试"""
        print("\n[测试12] 输入文本测试")
        print("警告: 此测试会输入文本到微信，请确保当前焦点在安全位置")
        response = input("是否继续？(y/n): ")
        if response.lower() != 'y':
            result = {"status": "skipped", "message": "用户跳过测试"}
            print_result("输入文本测试", result)
            return result
        
        success = self.manager.input_text("【测试输入】", send_enter=False)
        result = {
            "status": "success" if success else "error",
            "message": "输入成功" if success else "输入失败"
        }
        print_result("输入文本测试", result)
        return result
    
    def test_send_message(self):
        """测试13: 发送消息测试"""
        print("\n[测试13] 发送消息测试")
        print("警告: 此测试会发送消息到当前聊天窗口")
        response = input("是否继续？(y/n): ")
        if response.lower() != 'y':
            result = {"status": "skipped", "message": "用户跳过测试"}
            print_result("发送消息测试", result)
            return result
        
        success, error = self.manager.send_message("【自动化测试消息】")
        result = {
            "status": "success" if success else "error",
            "message": "消息发送成功" if success else f"发送失败: {error}"
        }
        print_result("发送消息测试", result)
        return result


def show_menu():
    """显示交互菜单"""
    print("\n" + "="*60)
    print("        微信自动化工具 - 测试菜单")
    print("="*60)
    
    tests = [
        ("1", "检查进程", "process"),
        ("2", "获取窗口", "window"),
        ("3", "窗口位置", "rect"),
        ("4", "微信状态", "status"),
        ("5", "截图测试", "capture"),
        ("6", "截图保存", "screenshot"),
        ("7", "OCR识别", "ocr"),
        ("8", "OCR关键词", "ocr_keyword"),
        ("9", "页面上下文", "context"),
        ("10", "点击测试", "click"),
        ("11", "滚动测试", "scroll"),
        ("12", "输入测试", "input"),
        ("13", "发送消息", "send"),
    ]
    
    for num, desc, _ in tests:
        print(f"  [{num:>2}] {desc}")
    
    print()
    print("  [ a] 运行所有基础测试")
    print("  [ d] 运行危险操作测试")
    print("  [ q] 退出")
    print("="*60)
    
    return tests


def interactive_mode():
    """交互模式"""
    tester = TestWeChatManager()
    
    test_map = {
        "process": tester.test_check_process,
        "window": tester.test_get_main_window,
        "rect": tester.test_get_window_rect,
        "status": tester.test_get_status,
        "capture": tester.test_capture,
        "screenshot": tester.test_take_screenshot,
        "ocr": tester.test_ocr_result,
        "ocr_keyword": tester.test_ocr_with_keyword,
        "context": tester.test_get_page_context,
        "click": tester.test_click,
        "scroll": tester.test_scroll,
        "input": tester.test_input_text,
        "send": tester.test_send_message,
    }
    
    while True:
        tests = show_menu()
        choice = input("\n请选择测试项: ").strip().lower()
        
        if choice == 'q':
            print("\n退出测试")
            break
        
        if choice == 'a':
            print("\n>>> 运行所有基础测试 <<<")
            for key in ["process", "window", "rect", "status", "capture", "screenshot", "ocr", "ocr_keyword", "context"]:
                try:
                    test_map[key]()
                except Exception as e:
                    print(f"\n测试异常: {e}")
            input("\n按回车继续...")
            continue
        
        if choice == 'd':
            print("\n>>> 运行危险操作测试 <<<")
            for key in ["click", "scroll", "input", "send"]:
                try:
                    test_map[key]()
                except Exception as e:
                    print(f"\n测试异常: {e}")
            input("\n按回车继续...")
            continue
        
        # 查找对应测试
        found = False
        for num, desc, key in tests:
            if choice == num or choice == key:
                if key in test_map:
                    try:
                        test_map[key]()
                    except Exception as e:
                        print(f"\n测试异常: {e}")
                    found = True
                    break
        
        if not found and choice not in ['', 'a', 'd', 'q']:
            print(f"\n无效选择: {choice}")
        
        input("\n按回车继续...")


if __name__ == "__main__":
    interactive_mode()
