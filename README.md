# Oracle Quick Open (OQO) Tool 🚀

> **Oracle Support工程师的必备工具** - 快速打开Oracle SR/KM文档/Bug/GRP/OLUEK/URL链接

## ✨ 功能特性

- **智能识别**: SR (3-xxxxxxxxx), CMOS (4-xxxxxxxxx), 文档 (xxxxx.x), Bug, JIRA, People等
- **一键处理**: 选中文本即可自动生成并打开相应链接
- **跨平台**: 支持Linux和macOS
- **智能搜索**: 自动Google搜索Oracle相关文档

## 🚀 快速开始

### 安装
```bash
# 克隆仓库
git clone https://github.com/yourusername/oqo.git
cd oqo

# 设置执行权限
chmod +x oqo.py

# 移动到bin目录（可选）
sudo cp oqo.py /usr/local/bin/oqo
# 或
cp oqo.py ~/bin/oqo.py
```

### 快捷键设置

#### Linux/macOS
```bash
# 设置执行权限
chmod +x ~/bin/oqo.py

# 在系统设置中添加快捷键
# Settings → Keyboard → Customize Shortcuts → Custom Shortcuts
# 命令: /usr/bin/python3 ~/bin/oqo.py
# 快捷键: Super(Windows/Meta) + J
```

#### 使用方法
1. 选中包含Oracle信息的文本
2. 按 `Super + J` 快捷键
3. 自动识别并打开相应链接

## 🛠️ 使用选项

```bash
# 正常模式（自动打开浏览器）
python3 oqo.py

# 只提取URL，不打开浏览器
python3 oqo.py --no-open

# 调试模式
python3 oqo.py --debug
```

## �� 支持格式

| 类型 | 格式 | 示例 |
|------|------|------|
| SR | `3-xxxxxxxxx` | `3-1234567890` |
| CMOS | `4-xxxxxxxxx` | `4-1234567890` |
| 文档 | `xxxxx.x` | `123456.1` |
| Bug | `xxxxxxx` | `1234567` |
| JIRA | `OLUEK-xxxx` | `OLUEK-1234` |
| People | `@username` | `@john` |

## 🐛 常见问题

- **权限问题**: `chmod +x oqo.py`
- **剪贴板**: 确保安装了 `xsel` 或 `xclip` (Linux)
- **浏览器**: 需要安装Google Chrome

## 许可证

MIT License

---

⭐ **Star支持一下！** ⭐
