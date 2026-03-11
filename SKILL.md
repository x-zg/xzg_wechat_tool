---
name: wechat-tool
version: 1.0.0
description: 微信客户端自动化控制工具，支持截图、OCR、点击、输入、滚动和消息发送
---

# WeChat Tool

Windows 微信客户端自动化控制工具。

## 功能

- 📸 截图：截取微信窗口或屏幕
- �� OCR：识别图片中的文字
- 👆 点击：点击指定坐标
- ⌨️ 输入：在指定位置输入文本
- 📜 滚动：上下滚动页面
- 📖 上下文：获取页面 OCR 结果
- ✉️ 消息：给指定联系人发送消息

## 安装

需要先安装 Python 依赖：

```bash
pip install pyautogui pygetwindow pillow pyperclip pytesseract opencv-python
```

## 使用方法

### 1. 获取微信状态

```python
from agent import get_wechat_status

status = get_wechat_status()
print(status)
```

### 2. 截取微信窗口

```python
from agent import capture_window

img = capture_window()
img.save("wechat_window.png")
```

### 3. 发送消息到当前聊天窗口

```python
from agent import send_message_to_current

# 给当前已打开的聊天窗口发送消息
success, err = send_message_to_current("你好！")

if success:
    print("消息发送成功")
else:
    print(f"发送失败: {err}")
```

### 4. 截图 + OCR

```python
from agent import capture_window, get_ocr_result

img = capture_window()
results = get_ocr_result(img)
```

### 5. 点击坐标

```python
from agent import click_coordinate

click_coordinate(100, 200)
```

### 6. 输入文本

```python
from agent import click_and_type

click_and_type("这是输入的文本", send_enter=True)
```

### 7. 滚动页面

```python
from agent import scroll

# 向下滚动
scroll(direction="down", amount=300)

# 向上滚动
scroll(direction="up", amount=300)
```

## 注意事项

1. 微信窗口需要保持打开状态
2. 发送消息时会自动激活微信窗口
3. 中文输入需要确保系统中文输入法正常工作
4. **自动识别联系人**：不传 contact 参数时，会自动从当前微信窗口标题获取联系人名称

## 文件结构

```
wechat-tool/
 agent.py           # 主程序
 app.py             # 微信自动化核心
 OCR.py             # OCR 帮助类
 ui_inspector.py    # UI 检查工具
 SKILL.md           # 技能描述
 _meta.json         # 元数据
 requirements.txt   # Python 依赖
```
