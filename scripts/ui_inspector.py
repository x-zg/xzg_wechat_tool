#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UI控件检测模块
获取窗口的UI控件树，结合OCR结果辅助模型判断操作位置
"""

import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


def get_control_info(control) -> Dict[str, Any]:
    """
    获取单个控件的信息
    
    Args:
        control: pywinauto 控件对象
        
    Returns:
        控件信息字典
    """
    try:
        info = {
            'control_type': control.element_info.control_type,
            'class_name': control.element_info.class_name or '',
            'name': control.element_info.name or '',
            'automation_id': control.element_info.automation_id or '',
            'is_enabled': control.element_info.enabled,
            'is_visible': control.element_info.visible,
            'can_focus': control.element_info.can_focus,
        }
        
        # 获取位置信息
        try:
            rect = control.rectangle()
            info['bbox'] = [rect.left, rect.top, rect.right, rect.bottom]
            info['center'] = [(rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2]
            info['width'] = rect.width()
            info['height'] = rect.height()
        except Exception:
            info['bbox'] = []
            info['center'] = []
            info['width'] = 0
            info['height'] = 0
        
        # 获取控件文本
        try:
            info['text'] = control.window_text() or ''
        except Exception:
            info['text'] = ''
        
        # 尝试获取值（针对输入框等）
        try:
            if hasattr(control, 'get_value'):
                info['value'] = control.get_value() or ''
        except Exception:
            pass
        
        return info
        
    except Exception as e:
        logger.debug(f"获取控件信息失败: {e}")
        return None


def get_all_controls(window, max_depth: int = 5, max_controls: int = 200) -> List[Dict[str, Any]]:
    """
    获取窗口中所有可见控件
    
    Args:
        window: pywinauto 窗口对象
        max_depth: 最大遍历深度
        max_controls: 最大控件数量限制
        
    Returns:
        控件信息列表
    """
    controls = []
    
    def traverse(control, depth: int):
        if depth > max_depth or len(controls) >= max_controls:
            return
            
        # 获取当前控件信息
        info = get_control_info(control)
        if info and info.get('is_visible', True):
            # 过滤无效控件
            if info.get('width', 0) > 0 and info.get('height', 0) > 0:
                info['depth'] = depth
                controls.append(info)
        
        # 递归遍历子控件
        try:
            children = control.children()
            for child in children:
                traverse(child, depth + 1)
        except Exception:
            pass
    
    try:
        traverse(window, 0)
    except Exception as e:
        logger.error(f"遍历控件失败: {e}")
    
    return controls


def get_interactive_controls(window, max_controls: int = 100) -> List[Dict[str, Any]]:
    """
    获取可交互控件（按钮、输入框、列表项等）
    
    Args:
        window: pywinauto 窗口对象
        max_controls: 最大控件数量限制
        
    Returns:
        可交互控件列表
    """
    # 可交互控件类型
    interactive_types = {
        'Button', 'CheckBox', 'RadioButton', 'ComboBox', 
        'Edit', 'Document', 'ListItem', 'MenuItem', 'TabItem',
        'Hyperlink', 'Text', 'Pane', 'List', 'Tree', 'TreeItem',
        'ToolBar', 'SplitButton', 'Spinner'
    }
    
    # 可交互类名关键词
    interactive_keywords = {
        'button', 'edit', 'list', 'menu', 'combo', 'check',
        'radio', 'tab', 'hyperlink', 'click', 'input'
    }
    
    all_controls = get_all_controls(window, max_depth=6, max_controls=max_controls * 2)
    
    interactive = []
    for ctrl in all_controls:
        ctrl_type = ctrl.get('control_type', '')
        class_name = ctrl.get('class_name', '').lower()
        name = ctrl.get('name', '').lower()
        
        # 判断是否可交互
        is_interactive = (
            ctrl_type in interactive_types or
            any(kw in class_name for kw in interactive_keywords) or
            ctrl.get('can_focus', False) or
            ctrl.get('text', '') or
            name
        )
        
        if is_interactive:
            # 添加交互提示
            ctrl['interactable'] = True
            if 'Button' in ctrl_type or 'button' in class_name:
                ctrl['action_hint'] = '可点击'
            elif 'Edit' in ctrl_type or 'edit' in class_name:
                ctrl['action_hint'] = '可输入'
            elif 'List' in ctrl_type or 'list' in class_name:
                ctrl['action_hint'] = '列表项，可点击选择'
            elif ctrl.get('text'):
                ctrl['action_hint'] = '有文字，可能可点击'
            else:
                ctrl['action_hint'] = '可交互'
            
            interactive.append(ctrl)
            
            if len(interactive) >= max_controls:
                break
    
    return interactive


def get_page_context(window, include_ocr: bool = True, include_controls: bool = True) -> Dict[str, Any]:
    """
    获取页面完整上下文（OCR + UI控件）
    
    Args:
        window: pywinauto 窗口对象
        include_ocr: 是否包含OCR结果
        include_controls: 是否包含控件信息
        
    Returns:
        页面上下文字典
    """
    context = {
        'window_info': {},
        'ocr_results': [],
        'controls': [],
        'interactive_controls': []
    }
    
    # 获取窗口信息
    try:
        rect = window.rectangle()
        context['window_info'] = {
            'title': window.window_text() or '',
            'bbox': [rect.left, rect.top, rect.right, rect.bottom],
            'width': rect.width(),
            'height': rect.height()
        }
    except Exception as e:
        logger.error(f"获取窗口信息失败: {e}")
    
    # 获取OCR结果
    if include_ocr:
        try:
            from OCR import ocr_endpoint
            ocr_results = ocr_endpoint(window, word=None)
            context['ocr_results'] = ocr_results
            logger.info(f"OCR识别到 {len(ocr_results)} 个文本区域")
        except Exception as e:
            logger.error(f"OCR识别失败: {e}")
    
    # 获取控件信息
    if include_controls:
        try:
            all_controls = get_all_controls(window, max_depth=5, max_controls=150)
            interactive_controls = get_interactive_controls(window, max_controls=80)
            
            context['controls'] = all_controls
            context['interactive_controls'] = interactive_controls
            logger.info(f"获取到 {len(all_controls)} 个控件，{len(interactive_controls)} 个可交互")
        except Exception as e:
            logger.error(f"获取控件失败: {e}")
    
    return context


def format_context_for_model(context: Dict[str, Any], max_items: int = 50) -> str:
    """
    将页面上下文格式化为适合模型理解的文本
    
    Args:
        context: 页面上下文
        max_items: 每类最大显示数量
        
    Returns:
        格式化的文本描述
    """
    lines = []
    
    # 窗口信息
    win_info = context.get('window_info', {})
    if win_info:
        lines.append(f"【窗口】{win_info.get('title', '未知')}")
        lines.append(f"  尺寸: {win_info.get('width', 0)} x {win_info.get('height', 0)}")
    
    # OCR结果（关键文字）
    ocr_results = context.get('ocr_results', [])
    if ocr_results:
        lines.append(f"\n【页面文字】(共{len(ocr_results)}处)")
        for i, item in enumerate(ocr_results[:max_items]):
            text = item.get('text', '')
            center = item.get('center', [0, 0])
            if text.strip():
                lines.append(f"  [{i+1}] \"{text}\" @ ({center[0]}, {center[1]})")
    
    # 可交互控件
    interactive = context.get('interactive_controls', [])
    if interactive:
        lines.append(f"\n【可交互元素】(共{len(interactive)}个)")
        
        # 按类型分组
        buttons = [c for c in interactive if 'Button' in c.get('control_type', '') or 'button' in c.get('class_name', '').lower()]
        edits = [c for c in interactive if 'Edit' in c.get('control_type', '') or 'edit' in c.get('class_name', '').lower()]
        lists = [c for c in interactive if 'List' in c.get('control_type', '') or 'list' in c.get('class_name', '').lower()]
        others = [c for c in interactive if c not in buttons and c not in edits and c not in lists]
        
        if buttons:
            lines.append(f"  按钮({len(buttons)}个):")
            for btn in buttons[:10]:
                name = btn.get('name', '') or btn.get('text', '')
                center = btn.get('center', [0, 0])
                lines.append(f"    - \"{name}\" @ ({center[0]}, {center[1]})")
        
        if edits:
            lines.append(f"  输入框({len(edits)}个):")
            for edit in edits[:10]:
                name = edit.get('name', '') or edit.get('text', '')
                value = edit.get('value', '')
                center = edit.get('center', [0, 0])
                info = f"    - \"{name}\""
                if value:
                    info += f" [当前值: \"{value}\"]"
                info += f" @ ({center[0]}, {center[1]})"
                lines.append(info)
        
        if lists:
            lines.append(f"  列表/列表项({len(lists)}个):")
            for lst in lists[:10]:
                name = lst.get('name', '') or lst.get('text', '')
                center = lst.get('center', [0, 0])
                lines.append(f"    - \"{name}\" @ ({center[0]}, {center[1]})")
        
        if others:
            lines.append(f"  其他交互元素({len(others)}个):")
            for item in others[:10]:
                name = item.get('name', '') or item.get('text', '')
                ctrl_type = item.get('control_type', '')
                center = item.get('center', [0, 0])
                hint = item.get('action_hint', '')
                lines.append(f"    - [{ctrl_type}] \"{name}\" {hint} @ ({center[0]}, {center[1]})")
    
    return '\n'.join(lines)


# 测试
if __name__ == "__main__":
    from app import autoLoginTB
    
    print("正在连接微信窗口...")
    controller = autoLoginTB()
    window = controller.start_app()
    
    if window:
        print("获取页面上下文...\n")
        context = get_page_context(window, include_ocr=True, include_controls=True)
        
        # 打印格式化结果
        formatted = format_context_for_model(context)
        print(formatted)
        
        print(f"\n统计:")
        print(f"  OCR文本: {len(context.get('ocr_results', []))} 个")
        print(f"  全部控件: {len(context.get('controls', []))} 个")
        print(f"  可交互控件: {len(context.get('interactive_controls', []))} 个")
    else:
        print("未找到微信窗口")
