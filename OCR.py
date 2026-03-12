#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
优化版 OCR 模块 - 不依赖 pywinauto
- 使用 PIL 直接截图
- 简化图像预处理
- 快速模式
"""
import cv2
import numpy as np
import logging
import time
from pathlib import Path
from PIL import ImageGrab
from rapidocr import RapidOCR

# 配置日志（减少输出）
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def init_ocr_engine():
    """初始化 OCR 引擎（快速模式）"""
    try:
        # 使用默认参数初始化
        engine = RapidOCR()
        
        # 预热
        test_img = np.ones((100, 100, 3), dtype=np.uint8) * 255
        engine(test_img, use_det=True, use_cls=False, use_rec=True)
        
        return engine
    except Exception as e:
        logger.error(f"OCR引擎初始化失败：{e}")
        raise


# 全局引擎（单例，懒加载）
_engine = None

def get_engine():
    """获取 OCR 引擎"""
    global _engine
    if _engine is None:
        _engine = init_ocr_engine()
    return _engine


def ocr_endpoint(win, word=None, fast_mode=True):
    """
    优化版 OCR 识别（不依赖 pywinauto）
    
    Args:
        win: 窗口对象（pygetwindow 或 pywinauto）
        word: 搜索关键词（可选）
        fast_mode: 快速模式
    
    Returns:
        list: OCR 结果列表
    """
    global _engine
    
    try:
        start_time = time.time()
        
        # 强制重新初始化 OCR 引擎（清除可能的缓存）
        _engine = None
        
        # 1. 获取窗口位置和截图（始终获取最新截图）
        win_left, win_top, win_right, win_bottom, screenshot = _get_window_screenshot(win)
        
        if screenshot is None:
            logger.warning("截图失败")
            return []
        
        # 保存调试截图（可选，用于验证是否获取最新截图）
        debug_path = Path(__file__).parent / "debug_ocr_screenshot.png"
        cv2.imwrite(str(debug_path), screenshot)
        print(f"[OCR] 调试截图已保存: {debug_path}")
        
        # 2. 图像预处理
        if fast_mode:
            img = _preprocess_fast(screenshot)
        else:
            img = _preprocess_full(screenshot)
        
        # 3. OCR 识别（重新获取引擎）
        engine = get_engine()
        ocr_result = engine(img, use_det=True, use_cls=False, use_rec=True)
        
        # 解析 RapidOCR 结果
        if ocr_result is None:
            return []
        
        # RapidOCR 返回 RapidOCROutput 对象
        # ocr_result.boxes: 检测框
        # ocr_result.txts: 文本列表
        # ocr_result.scores: 置信度列表
        if hasattr(ocr_result, 'boxes') and ocr_result.boxes is not None:
            boxes = ocr_result.boxes
            txts = ocr_result.txts if ocr_result.txts else []
            scores = ocr_result.scores if ocr_result.scores else []
        elif isinstance(ocr_result, tuple) and len(ocr_result) > 0:
            # 旧版本格式
            raw_results = ocr_result[0]
            if raw_results is None:
                return []
            boxes = [item[0] for item in raw_results if len(item) >= 3]
            txts = [str(item[1]).strip() for item in raw_results if len(item) >= 3]
            scores = [item[2] for item in raw_results if len(item) >= 3]
        else:
            return []
        
        if len(boxes) == 0:
            return []
        
        processed_items = []
        for i in range(len(boxes)):
            try:
                box = boxes[i]
                text = txts[i] if i < len(txts) else ""
                score = scores[i] if i < len(scores) else 0.0
                
                if not text:
                    continue
                
                box_np = np.array(box, dtype=np.float32).reshape(-1, 2)
                box_np[:, 0] += win_left
                box_np[:, 1] += win_top
                screen_box = box_np.astype(np.int32).tolist()
                
                x_coords = [p[0] for p in screen_box]
                y_coords = [p[1] for p in screen_box]
                x_min, x_max = min(x_coords), max(x_coords)
                y_min, y_max = min(y_coords), max(y_coords)
                center = [(x_min + x_max) // 2, (y_min + y_max) // 2]
                
                processed_items.append({
                    'text': text,
                    'scores': float(score),
                    'box': screen_box,
                    'center': center,
                    'total_width': x_max - x_min,
                    'total_height': y_max - y_min,
                    'x_min': x_min,
                    'x_max': x_max,
                    'y_min': y_min,
                    'y_max': y_max
                })
            except Exception:
                continue
        
        elapsed = time.time() - start_time
        logger.debug(f"OCR 完成，耗时: {elapsed:.2f}s，识别到 {len(processed_items)} 个文本区域")
        
        # 5. 关键词匹配
        if word is None:
            return processed_items
        
        return _match_keyword(processed_items, word)
        
    except Exception as e:
        logger.error(f"OCR 识别失败: {e}")
        return []


def _get_window_screenshot(win):
    """获取窗口截图（即使被遮挡也能正确截取，始终获取最新截图）"""
    import time as time_module
    capture_time = time_module.strftime('%H:%M:%S')
    
    try:
        # 尝试 pygetwindow 窗口
        if hasattr(win, 'left'):
            left, top = win.left, win.top
            right, bottom = win.right, win.bottom
            
            # 优先使用 ImageGrab 截取屏幕（保证最新）
            try:
                bbox = (left, top, right, bottom)
                img = ImageGrab.grab(bbox=bbox)
                screenshot = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
                print(f"[OCR] 截图成功 - 时间: {capture_time}, 区域: ({left},{top})-({right},{bottom})")
                return (left, top, right, bottom, screenshot)
            except Exception as e:
                logger.debug(f"ImageGrab 截图失败: {e}")
            
            # 回退：使用 win32gui 截取窗口
            screenshot = _capture_window_win32(win.title)
            if screenshot is not None:
                print(f"[OCR] win32截图成功 - 时间: {capture_time}")
                return (left, top, right, bottom, screenshot)
        
        # 尝试 pywinauto 窗口（使用 XZGUtil）
        try:
            from XZGUtil.ImagePositioning import window_rectangle
            screenshot = window_rectangle(win)
            
            # 获取窗口位置
            if hasattr(win, 'wrapper_object'):
                elem = win.wrapper_object()
                rect = elem.rectangle()
            else:
                return (0, 0, 0, 0, screenshot)
            
            left, top = rect.left, rect.top
            right, bottom = rect.right, rect.bottom
            
            return (left, top, right, bottom, screenshot)
        except Exception as e:
            logger.debug(f"pywinauto 截图失败: {e}")
            return (0, 0, 0, 0, None)
            
    except Exception as e:
        logger.debug(f"截图失败: {e}")
        return (0, 0, 0, 0, None)


def _capture_window_win32(window_title: str = None):
    """使用 win32 API 截取窗口（即使被遮挡）"""
    try:
        import win32gui
        import win32ui
        import win32con
        
        # 查找窗口句柄
        hwnd = None
        if window_title:
            hwnd = win32gui.FindWindow(None, window_title)
        if not hwnd:
            hwnd = win32gui.FindWindow("WeChatMainWndForPC", None)
        
        if not hwnd:
            return None
        
        # 获取窗口尺寸
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        
        if width <= 0 or height <= 0:
            return None
        
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
        
        # 转换为 numpy 数组
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        img = np.frombuffer(bmpstr, dtype=np.uint8)
        img = img.reshape((bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        
        # 清理资源
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        
        return img
    except Exception as e:
        logger.debug(f"win32 截图失败: {e}")
        return None


def _preprocess_fast(img):
    """快速预处理"""
    if img.ndim == 2:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
    elif img.ndim == 3 and img.shape[2] == 4:
        return cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    elif img.ndim == 3 and img.shape[2] == 3:
        # BGR 转 RGB
        return cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    return img


def _preprocess_full(img):
    """完整预处理"""
    img_rgb = _preprocess_fast(img)
    gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    _, binary = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return cv2.cvtColor(binary, cv2.COLOR_GRAY2RGB)


def _match_keyword(items, word):
    """匹配关键词"""
    target_word = str(word).strip()
    if not target_word:
        return []
    
    matched = []
    
    for item in items:
        text = item.get('text', '')
        if not text or target_word not in text:
            continue
        
        # 精准匹配
        if text == target_word:
            matched.append({
                'box': item['box'],
                'center': item['center'],
                'text': target_word
            })
            continue
        
        # 包含匹配：使用均分方式
        start_idx = text.find(target_word)
        x_min, x_max = item['x_min'], item['x_max']
        y_min, y_max = item['y_min'], item['y_max']
        
        char_width = (x_max - x_min) / len(text)
        keyword_left = int(x_min + start_idx * char_width)
        keyword_right = int(x_min + (start_idx + len(target_word)) * char_width)
        
        matched.append({
            'box': [
                [keyword_left, y_min],
                [keyword_right, y_min],
                [keyword_right, y_max],
                [keyword_left, y_max]
            ],
            'center': [(keyword_left + keyword_right) // 2, (y_min + y_max) // 2],
            'text': target_word,
            'full_text': text
        })
    
    return matched
