#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
优化版 OCR 模块
- 多尺度检测减少文字遗漏
- 增强图像预处理
- 结果去重合并
- OCR结果日志记录
"""
import sys
import io
import os

# ============== 强制统一编码为 UTF-8（解决 Windows 控制台编码问题）==============
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass

import cv2
import numpy as np
import logging
import time
import json
from pathlib import Path
from PIL import ImageGrab
from rapidocr import RapidOCR

# 配置日志输出编码为 UTF-8
_log_handler = logging.StreamHandler(sys.stdout)
_log_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
logging.basicConfig(level=logging.WARNING, handlers=[_log_handler])
logger = logging.getLogger(__name__)

# 配置参数
OCR_CONFIG = {
    'min_confidence': 0.2,          # 降低置信度阈值，减少遗漏
    'scales': [2.0, 3.0],           # 多尺度检测
    'use_multi_scale': True,        # 是否启用多尺度
    'iou_threshold': 0.3,           # 去重IOU阈值
}

_engine = None


def get_engine():
    """获取 OCR 引擎（单例）"""
    global _engine
    if _engine is None:
        _engine = RapidOCR()
        # 预热
        test_img = np.ones((100, 100, 3), dtype=np.uint8) * 255
        _engine(test_img, use_det=True, use_cls=False, use_rec=True)
    return _engine


def ocr_endpoint(win, word=None, fast_mode=False):
    """
    OCR 识别
    
    Args:
        win: 窗口对象
        word: 搜索关键词（可选）
        fast_mode: 快速模式（跳过多尺度检测）
    
    Returns:
        list: OCR 结果
    """
    try:
        start_time = time.time()
        
        # 1. 获取截图
        win_left, win_top, win_right, win_bottom, screenshot = _get_window_screenshot(win)
        
        if screenshot is None:
            logger.warning("截图失败")
            return []
        
        all_results = []
        
        if fast_mode or not OCR_CONFIG['use_multi_scale']:
            # 单尺度检测
            img, scale = _preprocess_enhanced(screenshot, scale_factor=2.0)
            results = _run_ocr(img, scale, win_left, win_top)
            all_results.extend(results)
        else:
            # 多尺度检测，合并结果
            for scale_factor in OCR_CONFIG['scales']:
                img, scale = _preprocess_enhanced(screenshot, scale_factor=scale_factor)
                results = _run_ocr(img, scale, win_left, win_top)
                all_results.extend(results)
            
            # 去重合并
            all_results = _deduplicate_results(all_results)
        
        # 关键词匹配
        if word is None:
            return all_results
        
        return _match_keyword(all_results, word)
        
    except Exception as e:
        logger.error(f"OCR 识别失败: {e}")
        return []


def _run_ocr(img, scale, win_left, win_top):
    """执行单次OCR并返回处理后的结果"""
    engine = get_engine()
    ocr_result = engine(img, use_det=True, use_cls=False, use_rec=True)
    
    if ocr_result is None:
        return []
    
    # 解析结果
    if hasattr(ocr_result, 'boxes') and ocr_result.boxes is not None:
        boxes = ocr_result.boxes
        txts = ocr_result.txts if ocr_result.txts else []
        scores = ocr_result.scores if ocr_result.scores else []
    elif isinstance(ocr_result, tuple) and len(ocr_result) > 0:
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
    
    # 处理结果（坐标还原）
    processed_items = []
    min_conf = OCR_CONFIG['min_confidence']
    
    for i in range(len(boxes)):
        try:
            box = boxes[i]
            text = txts[i] if i < len(txts) else ""
            score = scores[i] if i < len(scores) else 0.0
            
            if not text or score < min_conf:
                continue
            
            box_np = np.array(box, dtype=np.float32).reshape(-1, 2)
            box_np = box_np / scale
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
    
    return processed_items


def _deduplicate_results(items):
    """基于位置和文本的去重"""
    if not items:
        return items
    
    # 按置信度排序，保留高置信度结果
    items = sorted(items, key=lambda x: x['scores'], reverse=True)
    
    unique = []
    for item in items:
        is_duplicate = False
        for existing in unique:
            # 检查位置重叠
            iou = _calculate_iou(item, existing)
            if iou > OCR_CONFIG['iou_threshold']:
                # 位置重叠，检查文本相似度
                if item['text'] == existing['text']:
                    is_duplicate = True
                    break
                # 文本不同但位置重叠，保留更长的文本
                if len(item['text']) > len(existing['text']):
                    unique.remove(existing)
                    break
        
        if not is_duplicate:
            unique.append(item)
    
    return unique


def _calculate_iou(item1, item2):
    """计算两个区域的IOU"""
    x1 = max(item1['x_min'], item2['x_min'])
    y1 = max(item1['y_min'], item2['y_min'])
    x2 = min(item1['x_max'], item2['x_max'])
    y2 = min(item1['y_max'], item2['y_max'])
    
    if x2 <= x1 or y2 <= y1:
        return 0.0
    
    intersection = (x2 - x1) * (y2 - y1)
    area1 = (item1['x_max'] - item1['x_min']) * (item1['y_max'] - item1['y_min'])
    area2 = (item2['x_max'] - item2['x_min']) * (item2['y_max'] - item2['y_min'])
    union = area1 + area2 - intersection
    
    return intersection / union if union > 0 else 0.0


def _get_window_screenshot(win):
    """获取窗口截图"""
    try:
        if hasattr(win, 'left'):
            try:
                left, top = win.left, win.top
                right, bottom = win.right, win.bottom
            except Exception as e:
                logger.error(f"窗口属性访问失败: {e}")
                return (0, 0, 0, 0, None)
            
            bbox = (left, top, right, bottom)
            img = ImageGrab.grab(bbox=bbox)
            screenshot = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            
            return (left, top, right, bottom, screenshot)
        
        return (0, 0, 0, 0, None)
            
    except Exception as e:
        logger.error(f"截图失败: {e}")
        return (0, 0, 0, 0, None)


def _preprocess_fast(img):
    """快速预处理"""
    if img.ndim == 2:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
    elif img.ndim == 3 and img.shape[2] == 4:
        return cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    elif img.ndim == 3 and img.shape[2] == 3:
        return cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    return img


def _preprocess_enhanced(img, scale_factor=2.0):
    """增强预处理（提高识别率）
    
    Args:
        img: 输入图像
        scale_factor: 缩放比例
    
    Returns:
        tuple: (处理后的图像, 缩放比例)
    """
    # 1. 转换格式
    if img.ndim == 2:
        img_rgb = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
    elif img.ndim == 3 and img.shape[2] == 4:
        img_rgb = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    elif img.ndim == 3 and img.shape[2] == 3:
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    else:
        img_rgb = img
    
    # 2. 放大图像（针对小字，提高识别率）
    h, w = img_rgb.shape[:2]
    scale = scale_factor
    if scale > 1:
        img_rgb = cv2.resize(img_rgb, None, fx=scale, fy=scale, 
                            interpolation=cv2.INTER_CUBIC)
    
    # 3. 转灰度
    gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
    
    # 4. 去噪（双边滤波保留边缘）
    denoised = cv2.bilateralFilter(gray, 9, 75, 75)
    
    # 5. 对比度增强（CLAHE）
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(denoised)
    
    # 6. 锐化
    kernel_sharpen = np.array([
        [0, -1, 0],
        [-1, 5, -1],
        [0, -1, 0]
    ])
    sharpened = cv2.filter2D(enhanced, -1, kernel_sharpen)
    
    # 7. 转回 RGB
    return cv2.cvtColor(sharpened, cv2.COLOR_GRAY2RGB), scale


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
