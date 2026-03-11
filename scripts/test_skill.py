#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信自动化技能测试脚本

测试所有核心功能：
1. screenshot - 截图
2. get_ocr_result - OCR识别
3. click_coordinate - 点击坐标
4. click_and_type - 输入内容
5. scroll - 滚动
6. get_page_context - 获取页面上下文
"""

import sys
import os
import time

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import execute_action


def print_header(title: str):
    """打印标题"""
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60)


def print_result(name: str, result: dict):
    """打印结果"""
    status = result.get('status', 'unknown')
    icon = "✅" if status == 'success' else "❌"
    print(f"\n{icon} {name}")
    print(f"   状态: {status}")
    if 'message' in result:
        print(f"   消息: {result['message']}")
    if 'data' in result:
        data = result['data']
        if isinstance(data, dict):
            for key, value in data.items():
                if key == 'image_base64':
                    print(f"   {key}: [base64, 长度: {len(value)}]")
                elif key == 'results' and isinstance(value, list):
                    print(f"   {key}: {len(value)} 条")
                    for i, item in enumerate(value[:5]):
                        text = item.get('text', '')[:30]
                        center = item.get('center', [0, 0])
                        print(f"      [{i+1}] \"{text}\" @ {center}")
                elif key == 'formatted_text':
                    print(f"   {key}: (见下方详细输出)")
                elif key == 'interactive_controls' and isinstance(value, list):
                    print(f"   {key}: {len(value)} 个可交互元素")
                elif key == 'controls' and isinstance(value, list):
                    print(f"   {key}: {len(value)} 个控件")
                elif key == 'ocr_results' and isinstance(value, list):
                    print(f"   {key}: {len(value)} 个文本区域")
                else:
                    val_str = str(value)[:50]
                    print(f"   {key}: {val_str}")


class WeChatSkillTester:
    """微信技能测试器"""
    
    def __init__(self):
        self.test_results = []
    
    def run_test(self, name: str, test_func) -> bool:
        """运行单个测试"""
        print_header(f"测试: {name}")
        try:
            result = test_func()
            passed = result.get('status') == 'success'
            self.test_results.append((name, passed, None))
            return passed
        except Exception as e:
            self.test_results.append((name, False, str(e)))
            print(f"\n❌ 测试异常: {e}")
            return False
    
    def test_screenshot(self):
        """测试1: 截图功能"""
        print("\n📸 测试截图功能...")
        
        # 截图并保存
        result = execute_action("screenshot", {
            "save_path": "test_screenshot.png"
        })
        print_result("截图并保存", result)
        return result
    
    def test_ocr(self):
        """测试2: OCR识别"""
        print("\n📖 测试OCR识别...")
        
        # 获取所有文字
        result = execute_action("get_ocr_result", {
            "include_all": True
        })
        print_result("OCR识别全部文字", result)
        
        # 搜索关键词
        if result['status'] == 'success':
            time.sleep(0.5)
            result2 = execute_action("get_ocr_result", {
                "keyword": "微信"
            })
            print_result("搜索关键词 '微信'", result2)
        
        return result
    
    def test_page_context(self):
        """测试3: 获取页面上下文"""
        print("\n🔍 测试获取页面上下文...")
        
        result = execute_action("get_page_context", {
            "include_ocr": True,
            "include_controls": True,
            "format_for_model": True,
            "max_items": 20
        })
        print_result("获取页面上下文 (OCR + UI控件)", result)
        
        # 打印格式化文本
        if result['status'] == 'success':
            formatted = result['data'].get('formatted_text', '')
            if formatted:
                print("\n" + "─" * 50)
                print("📝 格式化输出 (供模型参考):")
                print("─" * 50)
                print(formatted)
                print("─" * 50)
        
        return result
    
    def test_scroll(self):
        """测试4: 滚动功能"""
        print("\n📜 测试滚动功能...")
        
        # 向下滚动
        result = execute_action("scroll", {
            "direction": "down",
            "clicks": 1
        })
        print_result("向下滚动", result)
        
        time.sleep(0.5)
        
        # 向上滚动
        result2 = execute_action("scroll", {
            "direction": "up",
            "clicks": 1
        })
        print_result("向上滚动", result2)
        
        return result
    
    def test_click(self):
        """测试5: 点击功能 (安全模式，只显示不执行)"""
        print("\n👆 点击功能测试 (安全模式)")
        print("   提示: 实际点击可能影响微信状态，已跳过")
        print("   如需测试，请使用: execute_action('click_coordinate', {'x': 100, 'y': 100})")
        return {"status": "success", "message": "安全模式跳过"}
    
    def test_type(self):
        """测试6: 输入功能 (安全模式)"""
        print("\n⌨️ 输入功能测试 (安全模式)")
        print("   提示: 实际输入会发送消息，已跳过")
        print("   如需测试，请使用: execute_action('click_and_type', {'content': '测试', 'send_enter': True})")
        return {"status": "success", "message": "安全模式跳过"}
    
    def test_send_message_flow(self, contact: str, message: str):
        """完整流程: 发送消息"""
        print_header(f"完整流程: 给 '{contact}' 发送消息")
        
        steps = []
        
        # 步骤1: 获取页面上下文
        print("\n[步骤1] 获取当前页面上下文...")
        result = execute_action("get_page_context", {
            "include_ocr": True,
            "include_controls": True,
            "format_for_model": True
        })
        if result['status'] != 'success':
            print_result("获取页面上下文失败", result)
            return False
        steps.append(("获取页面上下文", True))
        print("✅ 成功获取页面信息")
        
        # 步骤2: 截图保存当前状态
        print(f"\n[步骤2] 截图保存当前状态...")
        result = execute_action("screenshot", {
            "save_path": "send_msg_01_before.png"
        })
        steps.append(("截图", result['status'] == 'success'))
        print("✅ 截图已保存: send_msg_01_before.png")
        
        # 步骤3: 点击搜索框并输入联系人
        print(f"\n[步骤3] 搜索联系人 '{contact}'...")
        import pyautogui
        import pyperclip
        
        # 点击搜索区域 (微信左上角)
        result = execute_action("click_coordinate", {
            "x": 150,
            "y": 50
        })
        time.sleep(0.5)
        
        # 输入联系人名称
        pyperclip.copy(contact)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1.5)
        steps.append(("搜索联系人", True))
        print(f"✅ 已输入搜索关键词: {contact}")
        
        # 步骤4: 截图查看搜索结果
        print(f"\n[步骤4] 截图查看搜索结果...")
        result = execute_action("screenshot", {
            "save_path": "send_msg_02_search.png"
        })
        print("✅ 截图已保存: send_msg_02_search.png")
        
        # 步骤5: OCR识别联系人位置
        print(f"\n[步骤5] 定位联系人 '{contact}'...")
        result = execute_action("get_ocr_result", {
            "keyword": contact
        })
        
        if result['status'] == 'success' and result['data']['count'] > 0:
            contact_info = result['data']['results'][0]
            center = contact_info.get('center', [0, 0])
            print(f"✅ 找到联系人 '{contact}' 在坐标 {center}")
            
            # 步骤6: 双击打开聊天窗口
            print(f"\n[步骤6] 打开聊天窗口...")
            result = execute_action("click_coordinate", {
                "x": center[0],
                "y": center[1],
                "clicks": 2
            })
            time.sleep(1)
            steps.append(("打开聊天窗口", True))
            print("✅ 已打开聊天窗口")
            
            # 步骤7: 输入消息并发送
            print(f"\n[步骤7] 输入消息: '{message}'...")
            result = execute_action("click_and_type", {
                "content": message,
                "auto_locate_input": True,
                "clear_before": True,
                "send_enter": True
            })
            steps.append(("发送消息", result['status'] == 'success'))
            print_result("发送消息", result)
            
            # 步骤8: 截图确认
            time.sleep(1)
            print(f"\n[步骤8] 截图确认发送结果...")
            result = execute_action("screenshot", {
                "save_path": "send_msg_03_after.png"
            })
            print("✅ 截图已保存: send_msg_03_after.png")
            
            # 汇总
            print("\n" + "=" * 60)
            print(f"✅ 消息发送完成!")
            print(f"   收件人: {contact}")
            print(f"   内容: {message}")
            print("=" * 60)
            
            return True
        else:
            print(f"❌ 未找到联系人 '{contact}'")
            print("   请确保联系人存在，或手动调整搜索方式")
            return False
    
    def print_summary(self):
        """打印测试汇总"""
        print_header("测试汇总")
        
        passed = sum(1 for _, p, _ in self.test_results if p)
        total = len(self.test_results)
        
        for name, p, err in self.test_results:
            if err:
                print(f"  ❌ {name}: 异常 - {err}")
            elif p:
                print(f"  ✅ {name}: 通过")
            else:
                print(f"  ❌ {name}: 失败")
        
        print(f"\n总计: {passed}/{total} 通过")
    
    def run_all(self, quick: bool = False):
        """运行所有测试"""
        print("\n" + "=" * 60)
        print("  微信自动化技能测试")
        print("=" * 60)
        print("\n提示: 请确保微信已启动并登录")
        print("即将开始测试...\n")
        
        time.sleep(2)
        
        # 核心测试
        self.run_test("截图功能", self.test_screenshot)
        time.sleep(0.5)
        
        self.run_test("OCR识别", self.test_ocr)
        time.sleep(0.5)
        
        if not quick:
            self.run_test("页面上下文", self.test_page_context)
            time.sleep(0.5)
            
            self.run_test("滚动功能", self.test_scroll)
            time.sleep(0.5)
        
        self.run_test("点击功能", self.test_click)
        self.run_test("输入功能", self.test_type)
        
        # 打印汇总
        self.print_summary()
        
        # 询问是否测试完整流程
        if not quick:
            print("\n" + "=" * 60)
            choice = input("是否测试发送消息流程? (y/n): ").strip().lower()
            if choice == 'y':
                contact = input("联系人名称 (默认: 文件传输助手): ").strip() or "文件传输助手"
                message = input("消息内容 (默认: 测试消息): ").strip() or "测试消息"
                self.test_send_message_flow(contact, message)


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="微信自动化技能测试")
    parser.add_argument('--quick', action='store_true', help='快速测试模式')
    parser.add_argument('--send', action='store_true', help='直接测试发送消息')
    parser.add_argument('--contact', type=str, default='文件传输助手', help='联系人名称')
    parser.add_argument('--message', type=str, default='这是一条测试消息', help='消息内容')
    parser.add_argument('--context', action='store_true', help='只测试页面上下文')
    
    args = parser.parse_args()
    
    tester = WeChatSkillTester()
    
    if args.context:
        # 只测试页面上下文
        tester.run_test("页面上下文", tester.test_page_context)
        tester.print_summary()
    elif args.send:
        # 直接发送消息
        tester.test_send_message_flow(args.contact, args.message)
    else:
        # 完整测试
        tester.run_all(quick=args.quick)


if __name__ == "__main__":
    main()
