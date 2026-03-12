#! /usr/bin/env python3
# -*- coding: utf-8 -*-
import time

print("=" * 60)
print("测试 OCR 识别效果")
print("=" * 60)

from agent import WeChatManager

manager = WeChatManager()

print("\n[测试] OCR 识别...")
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
    for i, item in enumerate(results[:20]):  # 只显示前20个
        text = item.get('text', '')
        score = item.get('scores', 0)
        center = item.get('center', [])
        print(f"    [{i+1}] {text} (置信度: {score:.2f}, 中心: {center})")
    
    if count > 20:
        print(f"    ... 还有 {count - 20} 个结果")
else:
    print(f"[ERROR] {result.get('message')}")

print("\n" + "=" * 60)
