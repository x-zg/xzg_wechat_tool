#! /usr/bin/env python3
# -*- coding: utf-8 -*-
import time
import json

print("=" * 60)
print("测试优化后的 OCR")
print("=" * 60)

from agent import WeChatManager

manager = WeChatManager()

print("\n[测试 1] OCR 识别（增强模式）...")
start = time.time()
result = manager.get_ocr_result()
elapsed = time.time() - start

if result.get("status") == "success":
    data = result.get("data", {})
    count = data.get("count", 0)
    results = data.get("results", [])
    
    print(f"[OK] 识别成功！")
    print(f"  耗时: {elapsed:.2f} 秒")
    print(f"  识别到: {count} 个文本区域")
    
    print("\n  所有识别结果:")
    for i, item in enumerate(results):
        text = item.get('text', '')
        score = item.get('scores', 0)
        print(f"    [{i+1}] {text} (置信度: {score:.2f})")
else:
    print(f"[ERROR] {result.get('message')}")

print("\n" + "=" * 60)
