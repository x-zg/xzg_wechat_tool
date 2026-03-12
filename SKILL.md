---
name: wechat_tool
description: 微信客户端自动化控制工具，支持打开微信、截图、OCR识别、点击、输入、滚动、发送消息等操作
version: 1.0.0
author: xzg
permissions: 系统操作权限（控制微信窗口、鼠标键盘操作）
---

# 微信自动化技能

## 1. Description（技能详细说明）

这是一个微信客户端自动化控制工具，可以自动唤醒微信窗口、截图、识别界面内容、模拟点击和输入操作、发送消息。适用于自动化回复消息、批量操作、界面监控等场景。

## 2. When to use（触发场景）

当用户说以下内容时，使用此技能：

- "打开微信"、"唤醒微信"、"显示微信"
- "帮我截个微信的图"、"微信截图"
- "微信界面上有什么内容"、"识别微信界面文字"
- "点击微信上的XXX"、"在微信里输入XXX"
- "给XXX发消息"、"发送消息给XXX"
- "滚动微信聊天记录"
- "帮我操作微信"相关任务
- "微信状态"、"微信是否在运行"

## 3. How to use（调用逻辑）

### 3.1 打开/唤醒微信 / 截图

```bash
python {baseDir}/agent.py screenshot '{"save_path": "wechat.png"}'
```

### 3.2 获取微信状态

```bash
python {baseDir}/agent.py get_wechat_status '{}'
```

返回微信是否运行、当前聊天联系人、窗口位置等信息。

### 3.3 识别微信界面内容（OCR）

```bash
python {baseDir}/agent.py get_ocr_result '{}'
```

返回所有识别到的文字及其坐标位置。

### 3.4 点击指定位置

```bash
python {baseDir}/agent.py click_coordinate '{"x": 200, "y": 150}'
```

### 3.5 输入文字并发送

```bash
python {baseDir}/agent.py click_and_type '{"content": "消息内容", "send_enter": true}'
```

### 3.6 发送消息到当前聊天

```bash
python {baseDir}/agent.py send_message '{"message": "你好"}'
```

### 3.7 滚动页面

```bash
python {baseDir}/agent.py scroll '{"direction": "down", "amount": 300}'
```

### 3.8 获取完整页面信息

```bash
python {baseDir}/agent.py get_page_context '{}'
```

## 4. 典型操作流程示例

### 示例1：打开微信

**用户说**："打开微信"

**执行**：
```bash
python {baseDir}/agent.py screenshot '{}'
```

### 示例2：发送消息给某人

**用户说**："帮我给文件传输助手发一条消息：你好"

**执行步骤**：

1. 检查微信状态
```bash
python {baseDir}/agent.py get_wechat_status '{}'
```

2. 如果微信未运行，提示用户先启动微信

3. 发送消息（假设已在正确聊天窗口）
```bash
python {baseDir}/agent.py send_message '{"message": "你好"}'
```

### 示例3：识别界面内容

**用户说**："微信界面上有哪些文字？"

**执行**：
```bash
python {baseDir}/agent.py get_ocr_result '{}'
```

## 5. Edge cases（边缘场景处理）

- **微信未启动**：提示用户"微信未运行，请先启动微信并登录"
- **未找到窗口**：提示用户"未找到微信窗口，请确保微信已打开"
- **OCR 识别为空**：可能是微信窗口被遮挡，提示用户确保微信窗口在前台
- **发送消息失败**：提示用户确认是否已打开目标聊天窗口

## 6. 权限说明

此技能需要以下权限：
- 控制微信窗口（窗口置顶、激活）
- 鼠标点击操作
- 键盘输入操作
- 截图功能

## 7. 依赖安装

```bash
pip install pyautogui pygetwindow Pillow pyperclip rapidocr-onnxruntime numpy
```

## 8. 返回值格式

所有命令返回 JSON 格式：

**成功时**：
```json
{
  "status": "success",
  "message": "操作描述",
  "data": { ... }
}
```

**失败时**：
```json
{
  "status": "error",
  "message": "错误原因"
}
```

## 9. 支持的动作列表

| 动作 | 说明 | 必需参数 |
|------|------|----------|
| `screenshot` | 截图 | 无 |
| `get_wechat_status` | 获取微信状态 | 无 |
| `get_ocr_result` | OCR识别 | 无 |
| `click_coordinate` | 点击坐标 | x, y |
| `click_and_type` | 点击并输入 | content |
| `send_message` | 发送消息 | message |
| `scroll` | 滚动 | direction, amount |
| `get_page_context` | 获取页面上下文 | 无 |
