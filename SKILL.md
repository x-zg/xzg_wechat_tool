# 微信自动化技能 - OpenClaw Skill 使用指南

## 技能概述

本技能是一个**通用的微信自动化工具**，为 OpenClaw 平台提供 5 个核心动作：
1. 截图
2. OCR 文字识别
3. 点击坐标
4. 点击并输入
5. 滚动

**注意**：本技能是通用工具，不包含任何特定业务逻辑。所有操作都需要通过参数指定。

## 文件结构

```
xzg_wechat_tool/
├── SKILL_README.md          # 本文档
├── requirements.txt          # 依赖列表
└── scripts/
    ├── manifest.yaml         # 配置文件
    ├── app.py              # 窗口连接模块
    ├── main.py             # 技能主程序
    └── OCR.py              # OCR 识别模块
```

## 5 个核心动作详解

### 1. screenshot - 获取页面截图

**功能**：截取当前微信窗口的截图

**参数**：
```python
{
    "save_path": "screenshot.png"  # 可选，保存路径，默认 "wechat_screenshot.png"
}
```

**返回值**：
```python
{
    "status": "success",
    "message": "截图成功",
    "data": {
        "save_path": "保存路径",
        "width": 截图宽度，
        "height": 截图高度
    }
}
```

**使用示例**：
```python
# 截图并保存
result = skill.execute_action("screenshot", {
    "save_path": "current_page.png"
})
```

---

### 2. get_ocr_result - OCR 文字识别

**功能**：识别微信窗口中的文字，返回文字和坐标

**参数**：
```python
{
    "keyword": "搜索关键词",    # 可选，精确匹配某个词
    "include_all": True        # 可选，是否返回所有识别结果，默认 False
}
```

**返回值**：
```python
{
    "status": "success",
    "message": "识别成功",
    "data": {
        "count": 识别数量，
        "results": [
            {
                "text": "识别的文字",
                "center": [x, y],  # 中心坐标
                "bbox": [x1, y1, x2, y2]  # 边界框
            }
        ]
    }
}
```

**使用示例**：
```python
# 搜索特定文字
result = skill.execute_action("get_ocr_result", {
    "keyword": "联系人名称",
    "include_all": True
})

# 获取所有文字
result = skill.execute_action("get_ocr_result", {})
```

---

### 3. click_coordinate - 点击坐标

**功能**：点击指定坐标

**参数**：
```python
{
    "x": 500,       # X 坐标（必需）
    "y": 300,       # Y 坐标（必需）
    "clicks": 1,    # 可选，点击次数，默认 1
    "button": "left"  # 可选，鼠标按钮，默认 "left"
}
```

**返回值**：
```python
{
    "status": "success",
    "message": "点击成功",
    "data": {
        "position": [x, y]
    }
}
```

**使用示例**：
```python
# 单击
skill.execute_action("click_coordinate", {
    "x": 200,
    "y": 150
})

# 双击
skill.execute_action("click_coordinate", {
    "x": 200,
    "y": 150,
    "clicks": 2
})
```

---

### 4. click_and_type - 点击后输入内容

**功能**：点击指定坐标并输入文字，或自动定位输入框并输入

**参数**：
```python
{
    "content": "要输入的文字",      # 必需
    "x": 500,                      # 可选，X 坐标（与 auto_locate_input 二选一）
    "y": 300,                      # 可选，Y 坐标
    "clear_before": True,          # 可选，输入前是否清空，默认 True
    "send_enter": False,           # 可选，输入后是否按回车，默认 False
    "auto_locate_input": False,    # 可选，是否自动定位输入框，默认 False
    "interval": 0.05               # 可选，字符输入间隔，默认 0.05 秒
}
```

**返回值**：
```python
{
    "status": "success",
    "message": "输入成功",
    "data": {
        "content_length": 10,
        "clicked_position": [x, y],
        "content_preview": "输入内容的前 20 个字符"
    }
}
```

**使用示例**：
```python
# 方式 1：自动定位输入框并输入
skill.execute_action("click_and_type", {
    "content": "你好",
    "auto_locate_input": True,
    "send_enter": True
})

# 方式 2：指定坐标输入
skill.execute_action("click_and_type", {
    "content": "测试消息",
    "x": 500,
    "y": 800,
    "send_enter": True
})
```

**自动定位输入框说明**：
- 当 `auto_locate_input=True` 时，技能会尝试自动定位输入框
- 定位策略：
  1. OCR 识别关键词（如"发送"、"Send"等）
  2. 使用窗口相对位置（可配置 offset_x, offset_y）
- 可通过参数自定义位置：
  ```python
  skill.execute_action("click_and_type", {
      "content": "消息",
      "auto_locate_input": True,
      "offset_x": 0.4,  # X 方向相对位置
      "offset_y": 0.95  # Y 方向相对位置
  })
  ```

---

### 5. scroll - 上下滚动

**功能**：在指定位置滚动页面

**参数**：
```python
{
    "x": 500,            # 可选，X 坐标（默认屏幕中心）
    "y": 300,            # 可选，Y 坐标
    "clicks": 3,         # 可选，滚动次数，默认 1
    "direction": "down"  # 可选，滚动方向 "up" 或 "down"，默认 "down"
}
```

**返回值**：
```python
{
    "status": "success",
    "message": "滚动成功",
    "data": {
        "scroll_position": [x, y],
        "scroll_count": 3
    }
}
```

**使用示例**：
```python
# 向下滚动 3 次
skill.execute_action("scroll", {
    "clicks": 3,
    "direction": "down"
})

# 向上滚动
skill.execute_action("scroll", {
    "direction": "up",
    "clicks": 1
})
```

---

## 通用使用流程

### 场景：自动化操作微信

```python
from main import WeChatSkill

# 1. 创建技能实例
skill = WeChatSkill()

# 2. 截图查看当前页面
skill.execute_action("screenshot", {
    "save_path": "page_view.png"
})

# 3. OCR 识别页面内容
result = skill.execute_action("get_ocr_result", {
    "include_all": True
})

# 4. 根据识别结果点击
if result['data']['count'] > 0:
    x, y = result['data']['results'][0]['center']
    skill.execute_action("click_coordinate", {
        "x": x,
        "y": y,
        "clicks": 1
    })
    
    # 5. 输入内容
    skill.execute_action("click_and_type", {
        "content": "自定义内容",
        "auto_locate_input": True,
        "send_enter": True
    })
```

---

## 重要说明

### 1. 窗口唤醒机制

技能会自动处理窗口唤醒：
- 如果窗口未打开，使用 `Ctrl+Alt+W` 唤醒
- 如果窗口被遮挡，自动置顶到前台

### 2. 输入框定位

**自动定位策略**：
- 策略 1：OCR 识别关键词（"发送"、"Send"、"回复"、"Reply"）
- 策略 2：使用窗口相对位置（默认：宽度 40%，高度 95%）

**自定义位置**：
```python
skill.execute_action("click_and_type", {
    "content": "消息",
    "auto_locate_input": True,
    "offset_x": 0.5,  # 窗口中间
    "offset_y": 0.9   # 窗口底部 90% 处
})
```

### 3. 图片发送

图片发送需要使用剪贴板（外部工具）：
```python
import pyperclip
import pyautogui

# 复制图片到剪贴板
pyperclip.copy("图片路径")

# 定位输入框并粘贴
skill.execute_action("click_and_type", {
    "content": "",  # 空内容，只定位
    "auto_locate_input": True
})
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
```

---

## 错误处理

所有动作都可能返回错误，格式如下：

```python
{
    "status": "error",
    "message": "错误原因描述"
}
```

常见错误：
- `微信窗口未连接` - 需要先连接窗口
- `缺少必需参数` - 检查参数是否完整
- `无法定位输入框` - 尝试手动指定坐标或使用 offset_x/offset_y

---

## 依赖库

```
pywinauto>=0.6.8
pyautogui>=0.9.50
pyperclip>=1.8.0
Pillow>=9.0.0
numpy<2.0
rapidocr-onnxruntime
psutil>=5.9.0
```

---

## 注意事项

1. **通用性**：本技能是通用工具，不包含特定业务逻辑
2. **参数配置**：所有位置、偏移量等都可通过参数调整
3. **窗口状态**：确保微信窗口未最小化
4. **剪贴板**：发送图片会覆盖剪贴板内容
5. **网络延迟**：OCR 识别需要一定时间
6. **微信版本**：不同版本的微信界面可能略有差异，需要调整 offset 参数

---

## 配置参数说明

### _get_input_box_position 参数

| 参数 | 类型 | 范围 | 默认值 | 说明 |
|------|------|------|--------|------|
| offset_x | float | 0.0-1.0 | 0.4 | X 方向相对位置 |
| offset_y | float | 0.0-1.0 | 0.95 | Y 方向相对位置 |

**调整建议**：
- 如果输入框在窗口底部：`offset_y=0.95`
- 如果输入框在窗口中间：`offset_y=0.5`
- 如果输入框靠右：`offset_x=0.6`
- 如果输入框靠左：`offset_x=0.3`

---

**技能版本**：v1.0  
**更新时间**：2026-03-11  
**适用平台**：OpenClaw
