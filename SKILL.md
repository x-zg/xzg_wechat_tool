---
name: wechat-tool
version: 1.0.0
description: 微信客户端自动化控制工具
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

### 查看微信界面
```bash
python {baseDir}/scripts/main.py screenshot '{"save_path": "wechat_current.png"}'
```

### 识别界面内容
```bash
python {baseDir}/scripts/main.py get_ocr_result '{"include_all": true}'
```
返回所有文字及其坐标位置。

### 点击指定位置
```bash
python {baseDir}/scripts/main.py click_coordinate '{"x": 200, "y": 150}'
```

### 输入文字并发送
```bash
python {baseDir}/scripts/main.py click_and_type '{"content": "消息内容", "auto_locate_input": true, "send_enter": true}'
```

### 滚动页面
```bash
python {baseDir}/scripts/main.py scroll '{"direction": "down", "clicks": 3}'
```

## 重要说明

1. **自动唤醒**：任何操作都会自动唤醒微信窗口
2. **微信必须已登录**：首次使用需手动登录微信
3. **Windows 专用**：仅支持 Windows 系统
