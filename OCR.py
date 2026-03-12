#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
优化版 OCR 模块
- 增强图像预处理
- 优化 OCR 参数
- 提高识别准确率
"""
import cv2
import numpy as np
import logging
import time
from pathlib import Path
from PIL import ImageGrab
from rapidocr import RapidOCR

logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


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
        fast_mode: 快速模式（跳过预处理）
    
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
        
        # 2. 图像预处理
        if fast_mode:
            img = _preprocess_fast(screenshot)
            scale = 1.0
        else:
            img, scale = _preprocess_enhanced(screenshot)
        
        # 3. OCR 识别
        engine = get_engine()
        ocr_result = engine(img, use_det=True, use_cls=False, use_rec=True)
        
        # 4. 解析结果
        if ocr_result is None:
            return []
        
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
        
        # 5. 处理结果（坐标还原到原始尺寸）
        processed_items = []
        for i in range(len(boxes)):
            try:
                box = boxes[i]
                text = txts[i] if i < len(txts) else ""
                score = scores[i] if i < len(scores) else 0.0
                
                if not text or score < 0.3:
                    continue
                
                box_np = np.array(box, dtype=np.float32).reshape(-1, 2)
                # 坐标还原（除以缩放比例）
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
        
        elapsed = time.time() - start_time
        logger.info(f"OCR 完成: {elapsed:.2f}s, 识别到 {len(processed_items)} 个文本")
        
        # 6. 关键词匹配
        if word is None:
            return processed_items
        
        return _match_keyword(processed_items, word)
        
    except Exception as e:
        logger.error(f"OCR 识别失败: {e}")
        return []


def _get_window_screenshot(win):
    """获取窗口截图"""
    try:
        if hasattr(win, 'left'):
            left, top = win.left, win.top
            right, bottom = win.right, win.bottom
            
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


def _preprocess_enhanced(img):
    """增强预处理（提高识别率）
    
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
    scale = 2.0 if min(h, w) < 1000 else 1.5
    if scale > 1:
        img_rgb = cv2.resize(img_rgb, None, fx=scale, fy=scale, 
                            interpolation=cv2.INTER_CUBIC)
    
    # 3. 转灰度
    gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
    
    # 4. 去噪（双边滤波保留边缘）
    denoised = cv2.bilateralFilter(gray, 9, 75, 75)
    
    # 5. 对比度增强（CLAHE）- 调整参数提高效果
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(4, 4))
    enhanced = clahe.apply(denoised)
    
    # 6. 锐化 - 使用更温和的核
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
