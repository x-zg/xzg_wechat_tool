---
name: wechat_tool
description: 微信客户端自动化控制工具，支持打开微信、截图、OCR识别、点击、输入、滚动、发送消息等操作
version: 1.1.0
author: xzg
permissions: 系统操作权限（控制微信窗口、鼠标键盘操作）
---

# 微信自动化技能

## 1. Description（技能详细说明）

这是一个微信客户端自动化控制工具，可以自动唤醒微信窗口、截图、识别界面内容、模拟点击和输入操作、发送消息。

**核心特性：**
- 自动唤醒微信窗口（Ctrl+Alt+W 快捷键）
- 所有操作前自动激活窗口，确保窗口在前台
- 每次操作都获取最新截图，不使用缓存
- 支持多窗口标题匹配（"微信"、"WeChat"、"Weixin"）

## 2. When to use（触发场景）

当用户说以下内容时，使用此技能：

- "打开微信"、"唤醒微信"、"显示微信"
- "帮我截个微信的图"、"微信截图"
- "微信界面上有什么内容"、"识别微信界面文字"
- "点击微信上的XXX"、"在微信里输入XXX"
- "发送消息"、"在当前聊天发送XXX"
- "滚动微信聊天记录"
- "帮我操作微信"相关任务
- "微信状态"、"微信是否在运行"

## 3. How to use（调用逻辑）

**调用格式：**
```bash
python {baseDir}/agent.py <action> '<json_params>'
```

**重要说明：**
- `<action>` 是动作名称，必须是以下支持的动作之一
- `<json_params>` 是 JSON 格式的参数，必须用单引号包裹
- 如果参数为空，使用 `'{}'`
- 所有命令返回 JSON 格式结果

---

### 3.1 screenshot - 截图

**功能：** 截取微信窗口截图

**调用：**
```bash
python {baseDir}/agent.py screenshot '{}'
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| save_path | string | 否 | 无 | 保存路径（绝对路径或相对路径）。不传则返回 base64 编码 |

**示例：**

保存到文件：
```bash
python {baseDir}/agent.py screenshot '{"save_path": "D:/screenshots/wechat.png"}'
```

返回 base64：
```bash
python {baseDir}/agent.py screenshot '{}'
```

**返回值：**

成功（保存文件）：
```json
{
  "status": "success",
  "message": "已保存: D:/screenshots/wechat.png"
}
```

成功（返回 base64）：
```json
{
  "status": "success",
  "data": {
    "image_base64": "iVBORw0KGgoAAAANSUhEUgAA..."
  }
}
```

失败：
```json
{
  "status": "error",
  "message": "无法获取微信窗口"
}
```

---

### 3.2 get_wechat_status - 获取微信状态

**功能：** 检查微信是否运行及窗口状态

**调用：**
```bash
python {baseDir}/agent.py get_wechat_status '{}'
```

**参数说明：** 无参数

**返回值：**

微信运行中：
```json
{
  "status": "running",
  "title": "微信",
  "position": {"x": 100, "y": 100},
  "size": {"width": 900, "height": 700}
}
```

微信未运行：
```json
{
  "status": "not_running",
  "message": "微信未运行"
}
```

未找到窗口：
```json
{
  "status": "not_running",
  "message": "未找到微信窗口"
}
```

---

### 3.3 get_ocr_result - OCR 识别

**功能：** 识别微信界面文字并返回坐标信息

**调用：**
```bash
python {baseDir}/agent.py get_ocr_result '{}'
```

**参数说明：** 无参数（已移除 word、image、use_cache 参数）

**返回值：**

成功：
```json
{
  "status": "success",
  "data": {
    "results": [
      {
        "text": "文件传输助手",
        "scores": 0.95,
        "box": [[150, 200], [280, 200], [280, 230], [150, 230]],
        "center": [215, 215],
        "total_width": 130,
        "total_height": 30,
        "x_min": 150,
        "x_max": 280,
        "y_min": 200,
        "y_max": 230
      }
    ],
    "count": 1
  }
}
```

失败：
```json
{
  "status": "error",
  "message": "无法获取微信窗口"
}
```

**字段说明：**

| 字段 | 类型 | 说明 |
|------|------|------|
| text | string | 识别的文字内容 |
| scores | float | 识别置信度（0-1） |
| box | array | 文字区域四个角坐标 [[x1,y1], [x2,y2], [x3,y3], [x4,y4]] |
| center | array | 文字中心坐标 [x, y]，可用于点击 |
| total_width | int | 文字区域宽度 |
| total_height | int | 文字区域高度 |
| x_min, x_max, y_min, y_max | int | 文字区域边界坐标 |

---

### 3.4 click_coordinate - 点击坐标

**功能：** 点击屏幕指定坐标位置

**调用：**
```bash
python {baseDir}/agent.py click_coordinate '{"x": 200, "y": 150}'
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| x | int | **是** | - | 屏幕 X 坐标（绝对坐标） |
| y | int | **是** | - | 屏幕 Y 坐标（绝对坐标） |

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "点击 (200, 150)"
}
```

失败（缺少参数）：
```json
{
  "status": "error",
  "message": "缺少 x 或 y 参数"
}
```

失败（窗口问题）：
```json
{
  "status": "error",
  "message": "点击 (200, 150)"
}
```

---

### 3.5 click_and_type - 输入文字

**功能：** 在指定位置输入文字（可选先点击再输入）

**调用：**
```bash
python {baseDir}/agent.py click_and_type '{"content": "你好世界"}'
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| content | string | **是** | - | 要输入的文字内容 |
| x | int | 否 | null | 点击位置 X 坐标，不传则不点击 |
| y | int | 否 | null | 点击位置 Y 坐标，不传则不点击 |
| send_enter | bool | 否 | false | 输入后是否按回车键 |

**示例：**

只输入文字：
```bash
python {baseDir}/agent.py click_and_type '{"content": "你好"}'
```

点击后输入并回车：
```bash
python {baseDir}/agent.py click_and_type '{"content": "消息内容", "x": 500, "y": 600, "send_enter": true}'
```

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "输入成功"
}
```

失败（缺少 content）：
```json
{
  "status": "error",
  "message": "缺少 content 参数"
}
```

失败（窗口问题）：
```json
{
  "status": "error",
  "message": "输入失败"
}
```

---

### 3.6 send_message - 发送消息

**功能：** 在当前聊天窗口发送消息（自动定位输入框并发送）

**调用：**
```bash
python {baseDir}/agent.py send_message '{"message": "你好"}'
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| message | string | **是** | - | 要发送的消息内容 |

**注意：** 
- 此命令会自动点击输入框（窗口底部中央位置）
- 使用剪贴板粘贴方式输入，然后按回车发送
- **已移除 contact 参数**，需先手动打开目标聊天窗口

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "消息发送成功"
}
```

失败：
```json
{
  "status": "error",
  "message": "未找到微信窗口"
}
```

---

### 3.7 scroll - 滚动页面

**功能：** 滚动微信聊天记录或列表

**调用：**
```bash
python {baseDir}/agent.py scroll '{"direction": "down", "amount": 300}'
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| direction | string | 否 | "down" | 滚动方向："up" 向上滚动，"down" 向下滚动 |
| amount | int | 否 | 300 | 滚动量（像素） |
| x | int | 否 | 窗口中心 X | 滚动位置 X 坐标（屏幕绝对坐标） |
| y | int | 否 | 窗口中心 Y | 滚动位置 Y 坐标（屏幕绝对坐标） |

**示例：**

默认滚动（窗口中心，向下）：
```bash
python {baseDir}/agent.py scroll '{}'
```

向上滚动 500 像素：
```bash
python {baseDir}/agent.py scroll '{"direction": "up", "amount": 500}'
```

指定位置滚动（聊天列表区域）：
```bash
python {baseDir}/agent.py scroll '{"x": 200, "y": 400, "amount": 300}'
```

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "滚动成功"
}
```

失败：
```json
{
  "status": "error",
  "message": "滚动失败"
}
```

---

### 3.8 get_page_context - 获取页面上下文

**功能：** 获取当前微信页面的 OCR 识别结果（等同于 get_ocr_result）

**调用：**
```bash
python {baseDir}/agent.py get_page_context '{}'
```

**参数说明：** 无参数（已移除 use_cache 参数）

**返回值：** 同 `get_ocr_result`

---

## 4. 完整操作流程示例

### 示例1：打开并截图微信

**用户说**："打开微信并截图"

**执行步骤：**
```bash
# 截图命令会自动唤醒并激活微信窗口
python {baseDir}/agent.py screenshot '{"save_path": "wechat.png"}'
```

### 示例2：发送消息到当前聊天

**用户说**："在微信里发送消息：你好，我是机器人"

**执行步骤：**

1. 先确认微信状态
```bash
python {baseDir}/agent.py get_wechat_status '{}'
```

2. 如果返回 `"status": "not_running"`，提示用户启动微信

3. 发送消息
```bash
python {baseDir}/agent.py send_message '{"message": "你好，我是机器人"}'
```

### 示例3：点击指定文字

**用户说**："点击微信上的文件传输助手"

**执行步骤：**

1. 获取 OCR 结果
```bash
python {baseDir}/agent.py get_ocr_result '{}'
```

2. 从返回结果中找到 "文件传输助手" 的 center 坐标

3. 点击该位置
```bash
python {baseDir}/agent.py click_coordinate '{"x": 215, "y": 215}'
```

### 示例4：复杂操作流程

**用户说**："打开微信，点击文件传输助手，发送消息：测试完成"

**执行步骤：**

1. 唤醒微信并获取界面信息
```bash
python {baseDir}/agent.py get_ocr_result '{}'
```

2. 从结果中找到 "文件传输助手" 的 center 坐标，点击
```bash
python {baseDir}/agent.py click_coordinate '{"x": 215, "y": 215}'
```

3. 等待界面切换后，发送消息
```bash
python {baseDir}/agent.py send_message '{"message": "测试完成"}'
```

---

## 5. Edge cases（边缘场景处理）

| 场景 | 处理方式 | 用户提示 |
|------|----------|----------|
| 微信未启动 | 返回 `status: "not_running"` | "微信未运行，请先启动微信并登录" |
| 窗口被最小化 | 自动恢复窗口并置顶 | 无需提示 |
| 窗口被遮挡 | 使用 win32gui 强制置顶 | 无需提示 |
| 找不到窗口 | 自动尝试 Ctrl+Alt+W 唤醒 | "未找到微信窗口，已尝试唤醒" |
| OCR 结果为空 | 返回空数组 `results: []` | "未能识别到文字，请确保微信窗口可见" |
| 点击/输入失败 | 返回 `status: "error"` | "操作失败，请重试" |
| 参数格式错误 | 返回错误信息 | "参数 JSON 格式错误" |

---

## 6. 错误码说明

| status | 说明 |
|--------|------|
| success | 操作成功 |
| error | 操作失败，查看 message 了解原因 |
| running | 微信正在运行（仅 get_wechat_status） |
| not_running | 微信未运行或未找到窗口（仅 get_wechat_status） |

---

## 7. 权限说明

此技能需要以下系统权限：

- **窗口控制**：置顶、激活、恢复最小化窗口
- **鼠标操作**：点击、滚动
- **键盘操作**：输入文字、快捷键（Ctrl+V, Ctrl+Alt+W, Enter）
- **截图功能**：截取屏幕区域
- **剪贴板**：复制文字用于粘贴输入

---

## 8. 依赖安装

```bash
pip install pyautogui pygetwindow Pillow pyperclip rapidocr-onnxruntime numpy pywinauto psutil pywin32
```

**依赖说明：**

| 包名 | 用途 |
|------|------|
| pyautogui | 鼠标点击、键盘操作、滚动 |
| pygetwindow | 窗口查找和激活 |
| Pillow | 截图和图像处理 |
| pyperclip | 剪贴板操作 |
| rapidocr-onnxruntime | OCR 文字识别 |
| numpy | 图像数据处理 |
| pywinauto | 窗口连接（备用） |
| psutil | 进程检测 |
| pywin32 | Win32 API（窗口置顶） |

---

## 9. 支持的动作列表

| 动作 | 说明 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| screenshot | 截图 | 无 | save_path |
| get_wechat_status | 获取微信状态 | 无 | 无 |
| get_ocr_result | OCR 识别 | 无 | 无 |
| click_coordinate | 点击坐标 | x, y | 无 |
| click_and_type | 输入文字 | content | x, y, send_enter |
| send_message | 发送消息 | message | 无 |
| scroll | 滚动页面 | 无 | direction, amount, x, y |
| get_page_context | 获取页面上下文 | 无 | 无 |

---

## 10. 技术细节

### 10.1 窗口检测流程

1. 检查 `Weixin.exe` 进程是否存在
2. 使用 `pygetwindow` 搜索标题包含 "微信"、"WeChat"、"Weixin" 的窗口
3. 过滤窗口宽度在 500-2000 之间的有效窗口
4. 如果未找到，执行唤醒快捷键 `Ctrl+Alt+W`

### 10.2 窗口激活机制

```
优先级：
1. pygetwindow.activate() 
2. win32gui.SetForegroundWindow(hwnd)
3. win32gui.BringWindowToTop(hwnd)
```

### 10.3 截图机制

- **优先**：`PIL.ImageGrab.grab()` 截取屏幕区域（保证最新内容）
- **备用**：`win32gui` BitBlt 截图（可截取被遮挡窗口）

### 10.4 OCR 坐标转换

OCR 识别的是窗口内相对坐标，系统会自动转换为屏幕绝对坐标：
```
screen_x = window_left + relative_x
screen_y = window_top + relative_y
```

---

## 11. 注意事项

1. **坐标系统**：所有坐标均为屏幕绝对坐标，左上角为 (0, 0)
2. **JSON 格式**：参数必须是合法的 JSON 字符串，用单引号包裹
3. **消息发送**：`send_message` 会自动定位输入框，无需手动点击
4. **窗口前置**：所有操作都会自动将微信窗口置顶
5. **无缓存**：每次 OCR 和截图都获取最新画面
