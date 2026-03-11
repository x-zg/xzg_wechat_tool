#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Author   : 许老三
# @Time    : 2026/3/2 上午10:09
# @Site    :
# @File    : OCR.py
# @Software: PyCharm
import cv2
import numpy as np
import logging
import time
from XZGUtil.ImagePositioning import window_rectangle

# 配置日志（仅保留必要日志）
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 延迟初始化OCR引擎
engine = None
fine_grained_engine = None
_ocr_initialized = False


def init_ocr_engine():
    """初始化并预热OCR引擎，确保首次调用可用（延迟初始化）"""
    global _ocr_initialized, engine, fine_grained_engine
    
    if _ocr_initialized:
        return engine
    
    try:
        from rapidocr import RapidOCR
        
        logger.info("开始初始化OCR引擎...")
        engine = RapidOCR()
        fine_grained_engine = RapidOCR()

        # 创建测试图像进行预热
        test_img = np.ones((100, 100, 3), dtype=np.uint8) * 255
        cv2.putText(test_img, "test", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)

        # 执行预热调用
        engine(test_img, use_det=True, use_cls=False, use_rec=True)
        fine_grained_engine(test_img, use_det=True, use_cls=False, use_rec=True)
        
        _ocr_initialized = True
        logger.info("OCR引擎初始化并预热完成")
        return engine
    except Exception as e:
        logger.error(f"OCR引擎初始化失败：{e}", exc_info=True)
        return None

# 全局缓存（可选，用于优化多次调用）
_SCREENSHOT_CACHE = {}
_CACHE_EXPIRE_TIME = 2  # 缓存过期时间（秒）


def get_fine_grained_position(original_img, roi_box, target_word, win_left, win_top):
    """
    对目标区域进行细粒度OCR，精准定位关键词（字符级）
    :param original_img: 原始截图（RGB格式）
    :param roi_box: 原始识别区域的坐标 [x_min, y_min, x_max, y_max]
    :param target_word: 目标关键词
    :param win_left: 窗口左偏移
    :param win_top: 窗口上偏移
    :return: 关键词的精准坐标信息
    """
    try:
        x_min, y_min, x_max, y_max = roi_box
        # 裁剪目标区域（增加边界检查）
        if x_min < 0 or y_min < 0 or x_max > original_img.shape[1] or y_max > original_img.shape[0]:
            logger.warning(f"ROI区域超出图像范围：{roi_box}，图像尺寸：{original_img.shape}")
            return None

        roi_img = original_img[y_min:y_max, x_min:x_max]
        if roi_img.size == 0:
            logger.warning("裁剪后的ROI图像为空")
            return None

        # 增强预处理（针对文字粘连）
        gray_roi = cv2.cvtColor(roi_img, cv2.COLOR_RGB2GRAY)
        # 增加二值化的对比度，强化文字边缘
        _, thresh_roi = cv2.threshold(gray_roi, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        # 形态学操作：纵向膨胀，横向腐蚀，分离粘连文字
        kernel1 = np.ones((2, 1), np.uint8)  # 纵向膨胀核
        kernel2 = np.ones((1, 2), np.uint8)  # 横向腐蚀核
        thresh_roi = cv2.dilate(thresh_roi, kernel1, iterations=1)
        thresh_roi = cv2.erode(thresh_roi, kernel2, iterations=1)
        roi_rgb = cv2.cvtColor(thresh_roi, cv2.COLOR_GRAY2RGB)

        # 确保引擎已初始化
        if fine_grained_engine is None:
            init_ocr_engine()
            if fine_grained_engine is None:
                logger.error("细粒度OCR引擎不可用")
                return None
        
        # 细粒度OCR识别（启用检测，关闭分类）
        fine_ocr_result = fine_grained_engine(roi_rgb, use_det=True, use_cls=False, use_rec=True)
        fine_results = fine_ocr_result[0] if isinstance(fine_ocr_result, tuple) and len(fine_ocr_result) > 0 else []

        if not fine_results:
            logger.info("细粒度OCR未识别到文本")
            return None

        # 整理细粒度识别结果（按x坐标排序）
        char_items = []
        for item in fine_results:
            if len(item) < 3:
                continue
            box, text, score = item[0], str(item[1]).strip(), item[2]
            if not text or score < 0.5:  # 过滤低置信度结果
                continue

            # 计算字符区域的中心x坐标（用于排序）
            box_np = np.array(box, dtype=np.float32).reshape(-1, 2)
            char_x_min = min(box_np[:, 0])
            char_x_max = max(box_np[:, 0])
            char_center_x = (char_x_min + char_x_max) / 2
            char_y_min = min(box_np[:, 1])
            char_y_max = max(box_np[:, 1])

            char_items.append({
                'text': text,
                'center_x': char_center_x,
                'x_min': char_x_min,
                'x_max': char_x_max,
                'y_min': char_y_min,
                'y_max': char_y_max,
                'score': score
            })

        # 按x坐标排序（保证文字顺序正确）
        char_items.sort(key=lambda x: x['center_x'])

        # 拼接字符并查找目标关键词
        full_text = ''.join([item['text'] for item in char_items])
        start_idx = full_text.find(target_word)
        if start_idx == -1:
            logger.info(f"在细粒度识别结果中未找到关键词：{target_word}（识别到：{full_text}）")
            return None

        # 提取关键词对应的字符坐标
        keyword_chars = char_items[start_idx:start_idx + len(target_word)]
        if len(keyword_chars) != len(target_word):
            logger.warning(f"关键词长度不匹配：期望{len(target_word)}，实际{len(keyword_chars)}")
            return None

        # 计算关键词的整体边界（仅覆盖目标关键词）
        keyword_x_min = min([c['x_min'] for c in keyword_chars]) + x_min + win_left
        keyword_x_max = max([c['x_max'] for c in keyword_chars]) + x_min + win_left
        keyword_y_min = min([c['y_min'] for c in keyword_chars]) + y_min + win_top
        keyword_y_max = max([c['y_max'] for c in keyword_chars]) + y_min + win_top

        # 构造边界框和中心点
        keyword_box = [
            [int(keyword_x_min), int(keyword_y_min)],
            [int(keyword_x_max), int(keyword_y_min)],
            [int(keyword_x_max), int(keyword_y_max)],
            [int(keyword_x_min), int(keyword_y_max)]
        ]
        keyword_center = [
            (keyword_x_min + keyword_x_max) // 2,
            (keyword_y_min + keyword_y_max) // 2
        ]

        return {
            'box': keyword_box,
            'center': keyword_center,
            'text': target_word
        }

    except Exception as e:
        logger.error(f"细粒度定位失败：{e}", exc_info=True)
        return None


def ocr_endpoint(win, word=None, fuzzy=False, use_cache=True):
    """
    最终修复版：精准定位关键词的像素坐标（字符级）
    - 不传word：返回全部OCR结果
    - 传word：返回关键词的精准字符级坐标（而非所在字符串的坐标）
    - use_cache：是否使用截图缓存（默认开启）
    """
    try:
        # 1. 截取窗口截图并获取窗口坐标（增加缓存优化）
        win_id = id(win)
        current_time = time.time()

        # 检查缓存
        if use_cache and win_id in _SCREENSHOT_CACHE:
            cache_data = _SCREENSHOT_CACHE[win_id]
            if current_time - cache_data['timestamp'] < _CACHE_EXPIRE_TIME:
                screenshot = cache_data['screenshot']
                win_left = cache_data['win_left']
                win_top = cache_data['win_top']
                logger.info("使用缓存的截图数据")
            else:
                # 缓存过期，删除并重新截图
                del _SCREENSHOT_CACHE[win_id]
                screenshot = window_rectangle(win)
        else:
            # 首次获取截图
            screenshot = window_rectangle(win)

        # 获取窗口坐标
        if screenshot is not None:
            window_rect = win.rectangle()
            win_left = int(getattr(window_rect, 'left', 0))
            win_top = int(getattr(window_rect, 'top', 0))

            # 更新缓存
            if use_cache:
                _SCREENSHOT_CACHE[win_id] = {
                    'screenshot': screenshot,
                    'win_left': win_left,
                    'win_top': win_top,
                    'timestamp': current_time
                }

        if screenshot is None:
            logger.warning("窗口截图为空！")
            return []
        if not isinstance(screenshot, np.ndarray):
            raise ValueError(f"screenshot必须是numpy数组！当前类型：{type(screenshot)}")

        # 2. 预处理（优化版本，增强文字分割）
        if screenshot.ndim == 2:
            img_rgb = cv2.cvtColor(screenshot, cv2.COLOR_GRAY2RGB)
        elif screenshot.ndim == 3:
            img_rgb = cv2.cvtColor(screenshot, cv2.COLOR_RGBA2RGB) if screenshot.shape[2] == 4 else screenshot
        else:
            logger.warning(f"截图维度异常（{screenshot.ndim}维），直接使用原始图")
            img_rgb = screenshot

        # 优化预处理流程：增强文字与背景的对比度
        gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
        # 双边滤波：保留边缘的同时去噪（比高斯模糊更适合文字识别）
        blur = cv2.bilateralFilter(gray, 9, 75, 75)
        # 自适应阈值：调整参数增强分割效果
        processed_gray = cv2.adaptiveThreshold(
            blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV, blockSize=15, C=7  # 调整blockSize和C值
        )
        # 形态学操作：先膨胀后腐蚀，强化文字轮廓
        kernel = np.ones((2, 2), np.uint8)
        processed_gray = cv2.morphologyEx(processed_gray, cv2.MORPH_DILATE, kernel)
        processed_gray = cv2.morphologyEx(processed_gray, cv2.MORPH_ERODE, kernel)
        processed_img = cv2.cvtColor(processed_gray, cv2.COLOR_GRAY2RGB)

        # 3. OCR 识别（增加超时和重试机制）
        max_retries = 2
        retry_count = 0
        ocr_result_obj = None

        # 确保引擎已初始化
        try:
            if engine is None:
                init_ocr_engine()
                if engine is None:
                    logger.error("OCR 引擎不可用")
                    return []
        except Exception as e:
            logger.error(f"OCR 引擎初始化失败：{e}")
            logger.error("请检查 NumPy 版本，建议运行：pip install numpy<2.0")
            return []
        
        while retry_count < max_retries:
            try:
                ocr_result_obj = engine(processed_img, use_det=True, use_cls=False, use_rec=True)
                if ocr_result_obj is not None:
                    break
            except Exception as e:
                retry_count += 1
                logger.warning(f"OCR 识别第{retry_count}次失败：{e}，将重试")
                time.sleep(0.1)

        if ocr_result_obj is None:
            logger.error("OCR 识别多次失败")
            return []

        # 解析OCR结果
        raw_ocr_results = []
        if isinstance(ocr_result_obj, tuple):
            raw_ocr_results = ocr_result_obj[0] if len(ocr_result_obj) > 0 else []
        else:
            if hasattr(ocr_result_obj, 'results'):
                raw_ocr_results = ocr_result_obj.results
            elif hasattr(ocr_result_obj, 'boxes'):
                for b, t, s in zip(ocr_result_obj.boxes, ocr_result_obj.txts, ocr_result_obj.scores):
                    raw_ocr_results.append([b, t, s])

        if not raw_ocr_results:
            logger.info("未识别到任何文本内容")
            return []

        # 4. 处理结果
        processed_items = []
        for item in raw_ocr_results:
            if len(item) < 3:
                continue
            box, text, score = item[0], item[1], item[2]

            text_str = str(text).strip() if text is not None else ""
            score_float = float(score) if score is not None else 0.0

            # 坐标转换（增加异常处理）
            try:
                box_np = np.array(box, dtype=np.float32).reshape(-1, 2)
                box_np[:, 0] += win_left
                box_np[:, 1] += win_top
                screen_box = box_np.astype(np.int32).tolist()
            except Exception as e:
                logger.warning(f"处理box失败：{box}，错误：{e}")
                continue

            # 计算边界
            if len(screen_box) >= 2:
                x_coords = [p[0] for p in screen_box]
                y_coords = [p[1] for p in screen_box]
                x_min, x_max = min(x_coords), max(x_coords)
                y_min, y_max = min(y_coords), max(y_coords)
                screen_center = [(x_min + x_max) // 2, (y_min + y_max) // 2]
                total_width = x_max - x_min
                total_height = y_max - y_min
            else:
                screen_center = [0, 0]
                total_width, total_height = 0, 0
                x_min, x_max, y_min, y_max = 0, 0, 0, 0

            processed_items.append({
                'text': text_str,
                'scores': score_float,
                'box': screen_box,
                'center': screen_center,
                'total_width': total_width,
                'total_height': total_height,
                'x_min': x_min,
                'x_max': x_max,
                'y_min': y_min,
                'y_max': y_max
            })

        # 5. 关键词匹配（核心修复：优先字符级精准定位）
        if word is None:
            return processed_items
        else:
            target_word = str(word).strip()
            if not target_word:
                logger.warning("关键词为空")
                return []

            matched_results = []

            for item in processed_items:
                text_str = item['text']
                if not text_str or item['total_width'] <= 0:
                    continue

                # 核心逻辑：只要关键词在文本中，就执行字符级精准定位
                if target_word in text_str:
                    # 准备ROI区域（转换为相对截图的坐标）
                    roi_box = [
                        item['x_min'] - win_left,
                        item['y_min'] - win_top,
                        item['x_max'] - win_left,
                        item['y_max'] - win_top
                    ]
                    # 细粒度定位关键词（拆分字符，精准获取关键词坐标）
                    fine_result = get_fine_grained_position(
                        img_rgb, roi_box, target_word, win_left, win_top
                    )
                    if fine_result:
                        # 细粒度定位成功 → 添加精准坐标
                        matched_results.append(fine_result)
                    else:
                        # 降级方案：均分方式（仅兜底）
                        start_idx = text_str.find(target_word)
                        word_len = len(target_word)
                        end_idx = start_idx + word_len

                        x_min, x_max = item['x_min'], item['x_max']
                        y_min, y_max = item['y_min'], item['y_max']
                        char_width = item['total_width'] / len(text_str)

                        keyword_left = int(x_min + start_idx * char_width)
                        keyword_right = int(x_min + end_idx * char_width)
                        keyword_left = max(x_min, min(keyword_left, x_max - 1))
                        keyword_right = max(keyword_left + 1, min(keyword_right, x_max))

                        keyword_box = [
                            [keyword_left, y_min],
                            [keyword_right, y_min],
                            [keyword_right, y_max],
                            [keyword_left, y_max]
                        ]
                        keyword_center = [(keyword_left + keyword_right) // 2, (y_min + y_max) // 2]

                        matched_results.append({
                            'box': keyword_box,
                            'center': keyword_center,
                            'text': target_word,
                            'full_text': text_str,
                            'note': '使用降级方案（均分方式）'
                        })
                # 保留精准匹配逻辑（仅当文本完全等于关键词时触发）
                elif text_str == target_word:
                    matched_results.append({
                        'box': item['box'],
                        'center': item['center'],
                        'text': target_word
                    })

            logger.info(f"匹配到'{target_word}'共{len(matched_results)}个结果")
            if len(matched_results) == 0:
                logger.info(f"未匹配到关键词'{target_word}'，已识别的文本：{[item['text'] for item in processed_items]}")
            return matched_results

    except Exception as e:
        logger.error(f"OCR识别过程出错：{e}", exc_info=True)
        return []


# 测试函数（验证精准定位效果）
def test_keyword_position(win, test_word):
    """测试关键词的字符级精准定位"""
    logger.info(f"\n=== 测试关键词：{test_word} ===")
    # 直接调用关键词定位（无需先执行全量OCR）
    result = ocr_endpoint(win, word=test_word, fuzzy=True)

    if result:
        logger.info(f"精准定位结果：")
        for idx, res in enumerate(result):
            logger.info(f"  结果{idx + 1}：")
            logger.info(f"    关键词：{res['text']}")
            logger.info(f"    边界框：{res['box']}")
            logger.info(f"    中心点：{res['center']}")
            if 'note' in res:
                logger.info(f"    备注：{res['note']}")
    else:
        logger.info("未定位到关键词")
    return result


# 示例调用（根据你的实际窗口对象调整）
if __name__ == "__main__":
    # 请替换为你的实际窗口对象
    # win = 你的窗口对象（如通过pywin32/qt等获取的窗口句柄）
    # test_keyword_position(win, "文章")
    pass
