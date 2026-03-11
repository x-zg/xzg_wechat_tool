---
name: wechat_tool
description: "微信客户端自动化控制工具：支持截图、OCR识别、点击、输入、滚动、获取页面上下文等功能。"
metadata:
  {
    "openclaw":
      {
        "requires": { "bins": ["python"], "env": [] },
        "os": ["win32"]
      }
  }
---

# 微信自动化技能

微信客户端自动化控制工具，提供 6 个核心动作。

## 支持的动作

### 1. screenshot - 获取页面截图

截取当前微信窗口的截图。

**参数**：
- `save_path` (可选): 保存路径，默认返回 base64

**示例**：
```python
execute_action("screenshot", {"save_path": "page.png"})
```

---

### 2. get_ocr_result - OCR 文字识别

识别微信窗口中的文字，返回文字和坐标。

**参数**：
- `keyword` (可选): 搜索关键词
- `include_all` (可选): 是否返回所有结果，默认 true

**返回**：
- `results`: 识别结果列表，包含 `text`, `center`, `bbox`
- `count`: 识别数量

**示例**：
```python
# 搜索特定文字
execute_action("get_ocr_result", {"keyword": "联系人名称"})

# 获取所有文字
execute_action("get_ocr_result", {})
```

---

### 3. click_coordinate - 点击坐标

点击指定坐标位置。

**参数**：
- `x` (必填): X 坐标
- `y` (必填): Y 坐标
- `clicks` (可选): 点击次数，默认 1
- `button` (可选): 鼠标按钮，默认 "left"

**示例**：
```python
# 单击
execute_action("click_coordinate", {"x": 200, "y": 150})

# 双击
execute_action("click_coordinate", {"x": 200, "y": 150, "clicks": 2})
```

---

### 4. click_and_type - 点击后输入

点击指定坐标并输入文字，或自动定位输入框并输入。

**参数**：
- `content` (必填): 要输入的内容
- `x`, `y` (可选): 坐标（与 auto_locate_input 二选一）
- `auto_locate_input` (可选): 是否自动定位输入框，默认 false
- `clear_before` (可选): 输入前是否清空，默认 true
- `send_enter` (可选): 输入后是否按回车，默认 false
- `offset_x`, `offset_y` (可选): 自动定位时的相对位置

**示例**：
```python
# 自动定位输入框并输入
execute_action("click_and_type", {
    "content": "你好",
    "auto_locate_input": True,
    "send_enter": True
})

# 指定坐标输入
execute_action("click_and_type", {
    "content": "测试消息",
    "x": 500,
    "y": 800,
    "send_enter": True
})
```

---

### 5. scroll - 上下滚动

在指定位置滚动页面。

**参数**：
- `direction` (可选): 方向 "up" 或 "down"，默认 "down"
- `clicks` (可选): 滚动次数，默认 1
- `amount` (可选): 滚动量，默认 300

**示例**：
```python
execute_action("scroll", {"direction": "down", "clicks": 3})
```

---

### 6. get_page_context - 获取页面上下文

获取页面上下文（OCR + UI控件），辅助模型判断操作位置。

**参数**：
- `include_ocr` (可选): 是否包含OCR结果，默认 true
- `include_controls` (可选): 是否包含控件信息，默认 true
- `format_for_model` (可选): 是否格式化为适合模型的文本，默认 false
- `max_items` (可选): 每类最大返回数量，默认 50

**示例**：
```python
execute_action("get_page_context", {
    "include_ocr": True,
    "include_controls": True,
    "format_for_model": True
})
```

---

## 使用流程示例

```python
from scripts.main import execute_action

# 1. 截图查看当前页面
execute_action("screenshot", {"save_path": "page.png"})

# 2. OCR 识别页面内容
result = execute_action("get_ocr_result", {"include_all": True})

# 3. 根据识别结果点击
if result['data']['count'] > 0:
    x, y = result['data']['results'][0]['center']
    execute_action("click_coordinate", {"x": x, "y": y})

# 4. 输入内容
execute_action("click_and_type", {
    "content": "自定义内容",
    "auto_locate_input": True,
    "send_enter": True
})
```

## 注意事项

1. **窗口唤醒**: 技能会自动处理窗口唤醒，未打开则使用 `Ctrl+Alt+W`
2. **输入框定位**: 支持OCR关键词定位和窗口相对位置两种策略
3. **仅支持 Windows**: 本技能仅适用于 Windows 平台
4. **微信状态**: 确保微信已登录且未最小化

## 依赖

- pywinauto>=0.6.8
- pyautogui>=0.9.50
- rapidocr>=1.0.0
- Pillow>=9.0.0
- numpy<2.0
