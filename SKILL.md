---
name: wechat_tool
description: "微信客户端自动化控制工具。当用户需要操作微信时使用，如：打开微信、截图、发消息、查找联系人等。"
user-invocable: true
metadata:
  {
    "openclaw":
      {
        "requires": { "bins": ["python"] },
        "emoji": "💬"
      }
  }
---

# 微信自动化技能

控制 Windows 微信客户端的自动化工具。

## 快速指南

### 打开/唤醒微信

当用户说"打开微信"、"唤醒微信"、"显示微信"时，执行截图操作即可自动唤醒：

```bash
python {baseDir}/scripts/main.py screenshot "{}"
```

> 技能会自动检测微信是否打开，未打开则使用 Ctrl+Alt+W 唤醒。

---

### 查看微信界面

```bash
python {baseDir}/scripts/main.py screenshot '{"save_path": "wechat_current.png"}'
```

---

### 识别界面内容（找文字/按钮位置）

```bash
python {baseDir}/scripts/main.py get_ocr_result '{"include_all": true}'
```

返回所有文字及其坐标位置。

---

### 搜索特定内容

```bash
python {baseDir}/scripts/main.py get_ocr_result '{"keyword": "联系人名称"}'
```

---

### 点击指定位置

```bash
python {baseDir}/scripts/main.py click_coordinate '{"x": 200, "y": 150}'
```

双击：
```bash
python {baseDir}/scripts/main.py click_coordinate '{"x": 200, "y": 150, "clicks": 2}'
```

---

### 输入文字并发送

```bash
python {baseDir}/scripts/main.py click_and_type '{"content": "消息内容", "auto_locate_input": true, "send_enter": true}'
```

---

### 滚动页面

```bash
python {baseDir}/scripts/main.py scroll '{"direction": "down", "clicks": 3}'
```

---

### 获取完整页面信息（OCR + UI控件）

```bash
python {baseDir}/scripts/main.py get_page_context '{"format_for_model": true}'
```

---

## 常见场景示例

### 场景：发送消息给某人

```bash
# 1. 唤醒并截图
python {baseDir}/scripts/main.py screenshot '{"save_path": "step1.png"}'

# 2. 查找联系人位置
python {baseDir}/scripts/main.py get_ocr_result '{"keyword": "文件传输助手"}'

# 3. 根据返回坐标双击联系人（假设返回 center: [300, 200]）
python {baseDir}/scripts/main.py click_coordinate '{"x": 300, "y": 200, "clicks": 2}'

# 4. 输入消息并发送
python {baseDir}/scripts/main.py click_and_type '{"content": "你好", "auto_locate_input": true, "send_enter": true}'
```

### 场景：查看聊天记录

```bash
# 1. 截图当前聊天
python {baseDir}/scripts/main.py screenshot '{}'

# 2. OCR 识别内容
python {baseDir}/scripts/main.py get_ocr_result '{"include_all": true}'

# 3. 向上滚动查看更多
python {baseDir}/scripts/main.py scroll '{"direction": "up", "clicks": 5}'
```

---

## 重要说明

1. **自动唤醒**：任何操作都会自动唤醒微信窗口
2. **微信必须已登录**：首次使用需手动登录微信
3. **Windows 专用**：仅支持 Windows 系统
4. **快捷键**：微信唤醒使用 Ctrl+Alt+W

---

## 返回值格式

所有命令返回 JSON：

```json
{
  "status": "success",
  "message": "操作描述",
  "data": { ... }
}
```

失败时：
```json
{
  "status": "error",
  "message": "错误原因"
}
```
