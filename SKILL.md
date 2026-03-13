---
name: wechat_tool
description: 微信客户端自动化控制工具，支持打开微信、截图、OCR识别、点击、输入、滚动、发送消息等操作
version: 1.2.0
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
- **"看看谁给我发消息了"、"查看聊天列表"** → 使用 `get_chat_list`
- **"监控微信消息"、"自动回复"** → 使用 `start_monitor` 或 `auto_reply`
- **"回复XXX"** → 使用 `auto_reply`

## 3. How to use（调用逻辑）

**调用格式：**
```bash
python {baseDir}/agent.py <action> [--参数名 参数值]...
```

**重要说明：**
- `<action>` 是动作名称，必须是以下支持的动作之一
- 参数使用 `--参数名 参数值` 格式传递，无需 JSON
- 必需参数必须提供，可选参数可不传
- 所有命令返回 JSON 格式结果

---

### 3.1 screenshot - 截图

**功能：** 截取微信窗口截图

**调用：**
```bash
python {baseDir}/agent.py screenshot
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --save_path | string | 否 | 无 | 保存路径。不传则返回 base64 编码 |

**示例：**

保存到文件：
```bash
python {baseDir}/agent.py screenshot --save_path D:/screenshots/wechat.png
```

返回 base64：
```bash
python {baseDir}/agent.py screenshot
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
python {baseDir}/agent.py get_wechat_status
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
python {baseDir}/agent.py get_ocr_result
```

**参数说明：** 无参数

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

**坐标系统说明：**
- 所有坐标均为**屏幕绝对坐标**（左上角为原点 `0,0`）
- OCR 已自动将窗口相对坐标转换为屏幕绝对坐标
- `center` 坐标可直接传给 `click_coordinate` 使用

**置信度解读：**
- `scores >= 0.9`：非常可信，识别准确
- `0.7 <= scores < 0.9`：较可信，大部分情况正确
- `0.5 <= scores < 0.7`：一般可信，可能有小错误
- `scores < 0.5`：可信度较低，建议人工确认

**box 坐标顺序：**
```
(x1,y1) -------- (x2,y2)
   |                |
   |    文字区域    |
   |                |
(x4,y4) -------- (x3,y3)
```

---

### 3.4 click_coordinate - 点击坐标

**功能：** 点击屏幕指定坐标位置

**调用：**
```bash
python {baseDir}/agent.py click_coordinate --x 200 --y 150
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --x | int | **是** | - | 屏幕 X 坐标（绝对坐标） |
| --y | int | **是** | - | 屏幕 Y 坐标（绝对坐标） |

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "点击 (200, 150)"
}
```

失败：
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
python {baseDir}/agent.py click_and_type --content "你好世界"
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --content | string | **是** | - | 要输入的文字内容 |
| --x | int | 否 | null | 点击位置 X 坐标，不传则不点击 |
| --y | int | 否 | null | 点击位置 Y 坐标，不传则不点击 |
| --send_enter | flag | 否 | false | 输入后是否按回车键（无需值，有则生效） |

**示例：**

只输入文字：
```bash
python {baseDir}/agent.py click_and_type --content "你好"
```

点击后输入：
```bash
python {baseDir}/agent.py click_and_type --content "消息内容" --x 500 --y 600
```

输入后按回车：
```bash
python {baseDir}/agent.py click_and_type --content "消息内容" --x 500 --y 600 --send_enter
```

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "输入成功"
}
```

失败：
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
python {baseDir}/agent.py send_message --message "你好"
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --message | string | **是** | - | 要发送的消息内容 |

**注意：** 
- 此命令会自动点击输入框（窗口底部中央位置）
- 使用剪贴板粘贴方式输入，然后按回车发送
- 需先打开目标聊天窗口

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
python {baseDir}/agent.py scroll --direction down --amount 300
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --direction | string | 否 | "down" | 滚动方向："up" 或 "down" |
| --amount | int | 否 | 300 | 滚动量（像素） |
| --x | int | 否 | 窗口中心 X | 滚动位置 X 坐标 |
| --y | int | 否 | 窗口中心 Y | 滚动位置 Y 坐标 |

**示例：**

默认滚动（窗口中心，向下）：
```bash
python {baseDir}/agent.py scroll
```

向上滚动 500 像素：
```bash
python {baseDir}/agent.py scroll --direction up --amount 500
```

指定位置滚动：
```bash
python {baseDir}/agent.py scroll --x 200 --y 400 --amount 300
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
python {baseDir}/agent.py get_page_context
```

**参数说明：** 无参数

**返回值：** 同 `get_ocr_result`

---

### 3.9 get_chat_list - 获取聊天列表

**功能：** 获取左侧聊天列表的前N个联系人信息

**调用：**
```bash
python {baseDir}/agent.py get_chat_list --count 5
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --count | int | 否 | 5 | 获取的联系人数量 |

**返回值：**

成功：
```json
{
  "status": "success",
  "data": {
    "contacts": [
      {
        "index": 0,
        "name": "张三",
        "last_message": "好的，明天见",
        "position": {"x": 140, "y": 150},
        "rect": {"left": 0, "top": 120, "right": 280, "bottom": 180}
      },
      {
        "index": 1,
        "name": "李四",
        "last_message": "收到",
        "position": {"x": 140, "y": 220},
        "rect": {"left": 0, "top": 185, "right": 280, "bottom": 250}
      }
    ],
    "total": 2
  }
}
```

**用途：**
- 获取当前聊天列表中联系人的名称和最后一条消息
- 返回的 `position` 可用于 `click_contact` 点击进入聊天

---

### 3.10 click_contact - 点击联系人

**功能：** 点击联系人进入聊天窗口

**调用：**
```bash
# 方式1：通过联系人名称
python {baseDir}/agent.py click_contact --name "张三"

# 方式2：通过坐标位置
python {baseDir}/agent.py click_contact --x 140 --y 150
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --name | string | 否 | 无 | 联系人名称（模糊匹配） |
| --x | int | 否 | 无 | 点击X坐标 |
| --y | int | 否 | 无 | 点击Y坐标 |

**注意：** `--name` 和 `--x/--y` 二选一，优先使用坐标

---

### 3.11 auto_reply - 自动回复联系人

**功能：** 点击联系人并自动发送回复消息

**调用：**
```bash
python {baseDir}/agent.py auto_reply --name "张三" --message "好的，收到了"
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --name | string | 是 | 无 | 联系人名称 |
| --message | string | 是 | 无 | 回复内容 |

**返回值：**

成功：
```json
{
  "status": "success",
  "message": "已回复 张三: 好的，收到了"
}
```

---

### 3.12 start_monitor - 启动聊天监控

**功能：** 启动持续监控，检测聊天列表变化并自动回复（阻塞式运行）

**调用：**
```bash
# 只监控，不自动回复
python {baseDir}/agent.py start_monitor --interval 3 --max_loops 100

# 监控并自动回复固定内容
python {baseDir}/agent.py start_monitor --interval 3 --max_loops 100 --auto_reply "好的，稍后回复你"
```

**参数说明：**

| 参数名 | 类型 | 必需 | 默认值 | 说明 |
|--------|------|------|--------|------|
| --interval | float | 否 | 3.0 | 检查间隔（秒） |
| --max_loops | int | 否 | 100 | 最大循环次数 |
| --auto_reply | string | 否 | 无 | 自动回复内容（不传则只监控不回复） |

**工作原理：**
1. 首次获取聊天列表作为基准状态
2. 每隔 `interval` 秒检查一次
3. 检测到变化（新消息、位置变化）时：
   - 记录变化日志
   - 如果设置了 `--auto_reply`，自动回复有新消息的联系人

**返回值：**
```json
{
  "status": "success",
  "data": {
    "stats": {
      "loops": 50,
      "replies_sent": 3,
      "changes_detected": 5
    },
    "message": "监控结束，共检测 5 次变化，发送 3 条回复"
  }
}
```

---

## 4. 完整操作流程示例

> **⚠️ 重要提示：以下所有示例仅供说明使用方法，禁止自动执行！**
> 
> - 这些示例展示的是**当用户明确请求时**应如何操作
> - 只有在用户**实际发出请求**后才执行相应命令
> - 加载技能时**绝对不要**自动执行任何示例中的操作

### 示例1：打开并截图微信

**用户说**："打开微信并截图"

**执行步骤：**
```bash
python {baseDir}/agent.py screenshot --save_path wechat.png
```

### 示例2：发送消息到当前聊天

> 📝 **此为示例说明，非执行指令**

**用户说**："在微信里发送消息：你好，我是机器人"

**执行步骤：**

1. 确认微信状态
```bash
python {baseDir}/agent.py get_wechat_status
```

2. 如果返回 `"status": "not_running"`，提示用户启动微信

3. 发送消息
```bash
python {baseDir}/agent.py send_message --message "你好，我是机器人"
```

### 示例3：点击指定文字

> 📝 **此为示例说明，非执行指令**

**用户说**："点击微信上的文件传输助手"

**执行步骤：**

1. 获取 OCR 结果
```bash
python {baseDir}/agent.py get_ocr_result
```

2. 从返回结果中找到 "文件传输助手" 的 center 坐标

3. 点击该位置
```bash
python {baseDir}/agent.py click_coordinate --x 215 --y 215
```

### 示例4：搜索联系人并发送消息

> 📝 **此为示例说明，非执行指令。加载技能时禁止自动执行！**

**用户说**："给张三发消息：你好，明天开会"

**执行步骤：**

1. 获取当前界面 OCR 结果，找到搜索框
```bash
python {baseDir}/agent.py get_ocr_result
```

2. 从返回结果中找到 "搜索" 的 center 坐标，点击搜索框
```bash
# 假设搜索框 center 为 [2065, 362]
python {baseDir}/agent.py click_coordinate --x 2065 --y 362
```

3. 在搜索框输入联系人名字
```bash
python {baseDir}/agent.py click_and_type --content "张三"
```

4. 等待搜索结果加载，再次 OCR 识别搜索结果
```bash
python {baseDir}/agent.py get_ocr_result
```

5. 从结果中找到目标联系人 "张三" 的 center 坐标，点击进入聊天
```bash
# 假设联系人 center 为 [2200, 450]
python {baseDir}/agent.py click_coordinate --x 2200 --y 450
```

6. 发送消息
```bash
python {baseDir}/agent.py send_message --message "你好，明天开会"
```

**完整流程图：**
```
用户请求: "给张三发消息：你好，明天开会"
                    │
                    ▼
步骤1: get_ocr_result ──→ 定位搜索框
                    │
                    ▼
步骤2: click_coordinate ──→ 点击搜索框
                    │
                    ▼
步骤3: click_and_type ──→ 输入 "张三"
                    │
                    ▼
步骤4: get_ocr_result ──→ 找到联系人
                    │
                    ▼
步骤5: click_coordinate ──→ 点击联系人
                    │
                    ▼
步骤6: send_message ──→ 发送消息
                    │
                    ▼
                  完成 ✓
```

**关键点：**
- `get_ocr_result` 返回的 `center` 坐标可直接用于 `click_coordinate`
- 搜索后需要重新 OCR 获取搜索结果
- `send_message` 会自动定位输入框

### 示例5：查看谁和我说话了，说了什么

> 📝 **此为示例说明，非执行指令**

**用户说**："看看谁给我发消息了"

**执行步骤：**

1. 获取当前微信界面 OCR 结果
```bash
python {baseDir}/agent.py get_ocr_result
```

2. 解析返回结果，识别聊天列表中的信息：联系人名字、消息预览、时间戳、未读标记

**返回数据示例：**
```json
{
  "status": "success",
  "data": {
    "results": [
      {"text": "张三", "center": [2200, 400], "scores": 0.98},
      {"text": "14:19", "center": [2346, 400], "scores": 0.95},
      {"text": "你好，明天开会", "center": [2250, 430], "scores": 0.92},
      {"text": "[3条]李东阳", "center": [2200, 500], "scores": 0.96}
    ],
    "count": 4
  }
}
```

3. 如果需要查看某个聊天的详细内容
```bash
python {baseDir}/agent.py click_coordinate --x 2200 --y 400
```

4. 再次 OCR 获取聊天详情
```bash
python {baseDir}/agent.py get_ocr_result
```

5. 如果需要查看更多历史消息
```bash
python {baseDir}/agent.py scroll --direction up --amount 300
```

**关键点：**
- 聊天列表中通常包含：联系人名、时间、消息预览
- 未读消息常有 "[N条]" 标记
- 进入聊天后需要重新 OCR 获取详情
- 向上滚动可查看历史记录

### 示例6：滚动查看历史聊天记录

> 📝 **此为示例说明，非执行指令**

**用户说**："查看我和张三的历史聊天记录"

**执行步骤：**

1. 先通过搜索找到联系人并进入聊天（参考示例4的前5步）

2. 获取当前聊天界面 OCR 结果
```bash
python {baseDir}/agent.py get_ocr_result
```

3. 向上滚动查看更早的消息
```bash
python {baseDir}/agent.py scroll --direction up --amount 500
```

4. 再次 OCR 获取滚动后的消息
```bash
python {baseDir}/agent.py get_ocr_result
```

5. 重复步骤3-4直到找到目标内容

**完整流程图：**
```
用户请求: "查看我和张三的历史聊天记录"
                    │
                    ▼
步骤1: 搜索并进入张三聊天（参考示例4）
                    │
                    ▼
步骤2: get_ocr_result ──→ 识别当前消息
                    │
                    ▼
        ┌───────────┴───────────┐
        │ 找到目标内容？          │
        └───────────┬───────────┘
              否    │
                    ▼
步骤3: scroll --direction up ──→ 向上滚动
                    │
                    ▼
步骤4: get_ocr_result ──→ 识别新消息
                    │
                    └──────────┐
                               │
                    ┌──────────┘
                    ▼
                  完成 ✓
```

**关键点：**
- 向上滚动加载历史消息
- 每次滚动后需要重新 OCR
- 滚动量 `--amount` 可根据需要调整（默认300像素）
- 可以指定滚动位置 `--x --y` 在特定区域滚动

### 聊天记录展示格式

**微信聊天界面布局：**
```
┌─────────────────────────────────────┐
│  小许                      ⋮  │  ← 标题栏
├─────────────────────────────────────┤
│                                     │
│     [对方消息]           10:30      │  ← 左侧气泡
│                                     │
│           10:32      [自己的消息]   │  ← 右侧气泡
│                                     │
│     [对方消息]                       │
│                                     │
├─────────────────────────────────────┤
│  [输入框]                    发送   │
└─────────────────────────────────────┘
```

**识别和展示规则：**

1. **区分消息与界面元素**
   - 界面元素：`文件传输助手`、`朋友圈`、`消息`、`搜索`、`通讯录` 等 → 这些是侧边栏，不是聊天内容
   - 时间戳：`10:30`、`14:05`、`昨天`、`刚刚` → 消息时间
   - 消息内容：气泡中的文字，通常在聊天区域中央

2. **区分发送方**
   - 左侧气泡 → 对方发的消息
   - 右侧气泡 → 自己发的消息
   - 可通过 `x_min` 坐标判断：窗口左半部分是对方，右半部分是自己

3. **展示格式**
```
【聊天记录 - 小许】
─────────────────
14:05
对方：Python编程：去重
16:04
我：好的，我看一下
15:30
对方：还有个问题
─────────────────
```

**坐标判断逻辑：**
```python
def classify_message(item, window_center_x):
    """根据坐标判断消息发送方"""
    x_min = item['x_min']
    text = item['text']

    # 判断是否为消息内容（排除界面元素）
    if text in ['文件传输助手', '朋友圈', '消息', '搜索', '通讯录', '发现', '我']:
        return None  # 不是消息内容

    # 判断发送方
    if x_min < window_center_x:
        return {'sender': '对方', 'text': text}
    else:
        return {'sender': '我', 'text': text}
```

**正确回复示例：**

❌ 错误回复（只列出关键词）：
```
识别到的文本：小许、文件传输助手、朋友圈、消息、14:05、16:04...
```

✅ 正确回复（整理为对话形式）：
```
【小许的聊天记录】

14:05
小许：Python编程：去重

15:30
小许：还有个问题想问你

16:04
我：好的，我看一下

需要查看更多历史记录吗？我可以向上滚动查看更早的消息。
```

### 示例7：获取聊天列表并监控消息

> 📝 **此为示例说明，非执行指令**

**用户说**："看看谁给我发消息了，帮我监控一下"

**执行步骤：**

1. 获取当前聊天列表
```bash
python {baseDir}/agent.py get_chat_list --count 5
```

**返回示例：**
```json
{
  "status": "success",
  "data": {
    "contacts": [
      {"index": 0, "name": "张三", "last_message": "明天开会", "position": {"x": 140, "y": 150}},
      {"index": 1, "name": "李四", "last_message": "收到", "position": {"x": 140, "y": 220}}
    ],
    "total": 2
  }
}
```

2. 如果需要自动回复特定联系人
```bash
python {baseDir}/agent.py auto_reply --name "张三" --message "好的，收到了"
```

3. 如果需要持续监控并自动回复
```bash
python {baseDir}/agent.py start_monitor --interval 3 --max_loops 50 --auto_reply "好的，稍后回复"
```

**完整流程图：**
```
用户请求: "帮我监控微信消息并自动回复"
                    │
                    ▼
步骤1: get_chat_list ──→ 获取初始聊天列表
                    │
                    ▼
步骤2: start_monitor ──→ 开始监控循环
                    │
                    ├──→ 检测到新消息
                    │         │
                    │         ▼
                    │    auto_reply ──→ 自动回复
                    │         │
                    │         ▼
                    │    继续监控...
                    │
                    ▼
                  完成 ✓
```

**关键点：**
- `get_chat_list` 返回联系人名称、最后消息和点击位置
- `start_monitor` 是阻塞式运行，会持续监控直到达到 max_loops 或手动中断
- 可以只监控不回复（不传 `--auto_reply` 参数）

### 联系人列表识别

**微信主界面布局：**
```
┌──────────────────────────────────────────────────┐
│  ←  搜索                    │
├──────────────────────────────────────────────────┤
│  [联系人1名字]     时间1   │
│   消息预览...              │
├──────────────────────────────────────────────────┤
│  [联系人2名字]     时间2   │
│   消息预览...              │
├──────────────────────────────────────────────────┤
│  ...                       │
└──────────────────────────────────────────────────┘
        ↑                    ↑
   x: 750-900           x: 940-960
   (联系人名字)           (时间戳)
```

**联系人识别规则：**

1. **坐标范围判断**
   - 联系人名字：`x_min` 在 750-900 范围内
   - 时间戳：`x_min` 在 940-960 范围内
   - 联系人按 `y` 坐标从上到下排列

2. **排除非联系人项**
   | 排除类型 | 示例 | 特征 |
   |----------|------|------|
   | 搜索框 | "Q搜索"、"搜索" | 文字包含"搜索" |
   | 时间戳 | "14:05"、"昨天09:57"、"星期五" | 纯时间格式，x≈950 |
   | 界面按钮 | "口"、"X"、"8"、"中" | 单字符或图标误识别 |
   | 图标误识别 | "CE"、"offen"、"offcn" | 无意义英文字母组合 |
   | 底部菜单 | "三招考信息"、"三中公服务" | x > 1000，通常在底部 |

3. **联系人关联信息**
   - 每个联系人通常有对应的时间戳（y 坐标相近，差值 < 30）
   - 可能有消息预览（y 坐标略大，通常差值 15-25）
   - 可能有未读标记如 `[N条]`

**识别代码示例：**
```python
def identify_contacts(results):
    """从 OCR 结果中识别联系人列表"""
    contacts = []
    
    # 时间戳特征
    time_patterns = [r'\d{1,2}:\d{2}', r'昨天', r'前天', r'星期', r'刚刚']
    
    for item in results:
        text = item['text']
        x_min = item.get('x_min', item['center'][0])
        
        # 排除：不在联系人区域
        if x_min < 700 or x_min > 950:
            continue
        
        # 排除：时间戳
        if any(re.match(p, text) for p in time_patterns):
            continue
        
        # 排除：界面元素
        if text in ['搜索', 'Q搜索', '文件传输助手']:
            continue
            
        # 排除：单字符或无意义内容
        if len(text) <= 1 or text.lower() in ['ce', 'x', 'offen', 'offcn']:
            continue
        
        # 可能是联系人
        contacts.append({
            'name': text,
            'center': item['center'],
            'y': item['center'][1]
        })
    
    # 按 y 坐标排序
    return sorted(contacts, key=lambda c: c['y'])
```

**识别示例：**

OCR 原始结果：
```json
{"text": "小许", "center": [809, 136]}
{"text": "16:04", "center": [953, 135]}
{"text": "你好", "center": [808, 154]}
{"text": "丰巢", "center": [809, 200]}
{"text": "15:30", "center": [953, 200]}
```

识别后的联系人列表：
```
1. 小许 (y=136) - 时间: 16:04 - 预览: 你好
2. 丰巢 (y=200) - 时间: 15:30
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
| screenshot | 截图 | 无 | --save_path |
| get_wechat_status | 获取微信状态 | 无 | 无 |
| get_ocr_result | OCR 识别 | 无 | 无 |
| click_coordinate | 点击坐标 | --x, --y | 无 |
| click_and_type | 输入文字 | --content | --x, --y, --send_enter |
| send_message | 发送消息 | --message | 无 |
| scroll | 滚动页面 | 无 | --direction, --amount, --x, --y |
| get_page_context | 获取页面上下文 | 无 | 无 |
| **get_chat_list** | 获取聊天列表 | 无 | --count |
| **click_contact** | 点击联系人 | 无 | --name, --x, --y |
| **auto_reply** | 自动回复联系人 | --name, --message | 无 |
| **start_monitor** | 启动聊天监控 | 无 | --interval, --max_loops, --auto_reply |

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

## 11. 数据处理原则（重要）

**⚠️ OCR 结果处理必须遵循以下原则：**

### 11.1 禁止修改原始数据

- **绝对禁止**修改 OCR 返回的 `text` 字段内容
- **绝对禁止**根据上下文"猜测"或"纠正"识别结果
- **绝对禁止**编造或添加 OCR 结果中不存在的文字
- **绝对禁止**用你认为"更合理"的内容替换原始结果

### 11.2 正确示例

OCR 返回：
```json
{"text": "小许", "center": [809, 136], "scores": 1.0}
```

✅ 正确处理：直接使用 `"小许"`
❌ 错误处理：改为 `"小明"`（即使你觉得应该是"小明"）
❌ 错误处理：改为 `"小许/小明"`（添加猜测）
❌ 错误处理：输出 "联系人：小明"（替换原始值）

### 11.3 处理不清晰结果

如果 OCR 结果不清晰或有歧义：

1. **如实报告**：展示原始识别结果
2. **标注置信度**：低置信度结果应提示用户确认
3. **不要猜测**：不要根据上下文擅自修改

示例回复：
```
OCR 识别到："[N条]Python技术迷"（置信度 0.91）
注意：部分文字可能识别不准确，建议人工确认。
```

### 11.4 完整数据链路

```
OCR.py (ocr_endpoint)
    → 返回原始识别结果
    → 写入 ocr_log.txt（日志）
    → agent.py (get_ocr_result) 返回给调用方
    → LLM 接收 JSON 数据
    → LLM 必须原样使用，不得修改
```

**任何环节都不应修改 OCR 识别的原始文字内容。**

### 11.5 验证机制

调用 `get_ocr_result` 后，LLM 必须确保：

1. **返回给用户的内容** 与 **OCR JSON 中的 text 字段** 完全一致
2. 如需展示联系人名字，必须从 OCR 结果中提取，不得凭空生成
3. 如果用户质疑结果，可以提示查看 `ocr_log.txt` 日志文件进行对比

**常见幻觉场景及纠正：**

| 错误行为 | 正确做法 |
|----------|----------|
| OCR 返回 "小许"，输出 "小明" | 输出 "小许" |
| OCR 返回 "丰巢"，输出 "蜂巢" | 输出 "丰巢" |
| OCR 返回 "许志国"，输出 "许志华" | 输出 "许志国" |
| OCR 返回部分内容，自己补全 | 只输出 OCR 实际识别到的内容 |

---

## 12. 注意事项

1. **坐标系统**：所有坐标均为屏幕绝对坐标，左上角为 (0, 0)
2. **参数格式**：使用 `--参数名 参数值` 格式，无需 JSON
3. **消息发送**：`send_message` 会自动定位输入框，无需手动点击
4. **窗口前置**：所有操作都会自动将微信窗口置顶
5. **无缓存**：每次 OCR 和截图都获取最新画面
6. **数据忠实**：必须原样使用 OCR 返回的文字，禁止修改
