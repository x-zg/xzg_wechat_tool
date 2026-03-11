---
name: wechat_tool
description: "微信客户端自动化控制工具：支持截图、OCR识别、点击、输入、滚动、获取页面上下文等功能。"
user-invocable: true
metadata:
  {
    "openclaw":
      {
        "requires": { "bins": ["python"], "env": [] },
        "os": ["win32"],
        "emoji": "💬"
      }
  }
---

# 微信自动化技能

微信客户端自动化控制工具，通过 Python 脚本控制微信窗口。

## 使用方式

调用 Python 脚本执行操作：

```bash
python {baseDir}/scripts/main.py <action> '<json_params>'
```

**示例**：
```bash
# 截图
python {baseDir}/scripts/main.py screenshot '{"save_path": "page.png"}'

# OCR 识别
python {baseDir}/scripts/main.py get_ocr_result '{"keyword": "微信"}'

# 点击坐标
python {baseDir}/scripts/main.py click_coordinate '{"x": 200, "y": 150}'

# 输入内容
python {baseDir}/scripts/main.py click_and_type '{"content": "你好", "send_enter": true}'
```

---

## 支持的动作

### 1. screenshot - 获取页面截图

**参数**：`save_path` (可选) - 保存路径

**执行**：
```bash
python {baseDir}/scripts/main.py screenshot '{"save_path": "page.png"}'
```

---

### 2. get_ocr_result - OCR 文字识别

**参数**：
- `keyword` (可选): 搜索关键词
- `include_all` (可选): 是否返回所有结果，默认 true

**执行**：
```bash
python {baseDir}/scripts/main.py get_ocr_result '{"keyword": "联系人", "include_all": true}'
```

---

### 3. click_coordinate - 点击坐标

**参数**：
- `x` (必填): X 坐标
- `y` (必填): Y 坐标
- `clicks` (可选): 点击次数，默认 1

**执行**：
```bash
python {baseDir}/scripts/main.py click_coordinate '{"x": 200, "y": 150, "clicks": 2}'
```

---

### 4. click_and_type - 点击后输入

**参数**：
- `content` (必填): 要输入的内容
- `auto_locate_input` (可选): 是否自动定位输入框
- `send_enter` (可选): 输入后是否按回车

**执行**：
```bash
python {baseDir}/scripts/main.py click_and_type '{"content": "你好", "auto_locate_input": true, "send_enter": true}'
```

---

### 5. scroll - 上下滚动

**参数**：
- `direction` (可选): 方向 "up" 或 "down"
- `clicks` (可选): 滚动次数

**执行**：
```bash
python {baseDir}/scripts/main.py scroll '{"direction": "down", "clicks": 3}'
```

---

### 6. get_page_context - 获取页面上下文

**参数**：
- `include_ocr` (可选): 是否包含OCR结果
- `include_controls` (可选): 是否包含控件信息
- `format_for_model` (可选): 是否格式化输出

**执行**：
```bash
python {baseDir}/scripts/main.py get_page_context '{"include_ocr": true, "format_for_model": true}'
```

---

## 常用操作流程

### 打开微信并发送消息

1. 先截图查看当前状态
2. OCR 识别页面内容
3. 点击目标联系人
4. 输入消息并发送

```bash
# 1. 截图
python {baseDir}/scripts/main.py screenshot '{"save_path": "wechat.png"}'

# 2. OCR 识别联系人
python {baseDir}/scripts/main.py get_ocr_result '{"keyword": "文件传输助手"}'

# 3. 点击联系人（根据 OCR 返回的坐标）
python {baseDir}/scripts/main.py click_coordinate '{"x": 300, "y": 200, "clicks": 2}'

# 4. 输入并发送消息
python {baseDir}/scripts/main.py click_and_type '{"content": "测试消息", "auto_locate_input": true, "send_enter": true}'
```

---

## 注意事项

1. **窗口唤醒**: 技能会自动使用 `Ctrl+Alt+W` 唤醒微信
2. **仅支持 Windows**: 本技能仅适用于 Windows 平台
3. **微信状态**: 确保微信已登录

## 依赖安装

```bash
pip install -r {baseDir}/requirements.txt
```
