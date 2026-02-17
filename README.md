# 🚀 Micro Agent 1.0: 极简而强大的 AI 智能体框架

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Rich Interface](https://img.shields.io/badge/UI-Rich%20CLI-orange.svg)](https://github.com/Textualize/rich)

**Micro Agent** 是一款基于 Python 开发的轻量级 AI 智能体框架。它不仅能通过 OpenAI 兼容接口连接各种大语言模型，还内置了一套强大的工具集，能够自动化生成专业文档、进行网络搜索、执行终端命令以及管理文件系统。

---

## ✨ 核心特性

### 🎨 绝美终端交互
- **彩虹 ASCII Logo**: 启动时展现华丽的视觉效果。
- **富文本输出**: 基于 `rich` 库，提供精美的表格、面板、进度条和 Markdown 渲染。
- **实时反馈**: 智能体思考与工具执行过程透明可见。

### 🛠️ 强大的工具箱 (Built-in Tools)
- **专业文档生成**: 
  - **Word**: 自动生成带封面、目录、图表和精美样式的专业报告。
  - **Excel**: 支持自动筛选、冻结窗格、条件格式及数据可视化。
  - **PowerPoint**: 内置多种主题（Ocean, Forest, Tech 等），自动排版精美幻灯片。
- **网络与信息**: 集成搜索工具与实时天气查询。
- **系统级操作**: 终端命令执行、文件系统深度管理。

### 🧠 智能上下文管理
- **动态适配**: 根据不同模型的 Context Limit 自动调整历史记录长度。
- **自动压缩**: 智能触发上下文摘要，确保长对话不中断。
- **多模型支持**: 轻松配置并切换多个 OpenAI 兼容的 API 模型。

---

## 🚀 快速开始

### 1. 克隆仓库
```bash
git clone https://github.com/your-username/micro-agent.git
cd micro-agent
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
```
*(注：主要依赖包括 `openai`, `rich`, `requests` 等)*

### 3. 启动程序
```bash
python main.py
```
首次启动将引导您配置 API Base URL 和 API Key。

---

## 📂 项目结构

```text
.
├── main.py                # 程序入口，负责 AI 逻辑与交互循环
├── config.py              # 动态配置管理与上下文计算
├── tools/                 # 工具集目录
│   ├── base.py            # 工具基类与工作空间管理
│   ├── manager.py         # 工具执行调度器
│   ├── definitions.py     # 工具函数定义 (JSON Schema)
│   └── ...                # 各类具体工具实现 (Word, Excel, PPT, etc.)
├── work_space/            # 智能体的工作目录，生成的文档将存放于此
└── .gitignore             # 排除敏感配置与缓存
```

---

## 🛠️ 进阶配置

程序运行后会生成以下配置文件（已在 `.gitignore` 中排除，确保安全）：
- `model_config.json`: 存储模型 API 信息。
- `tools_config.json`: 存储第三方工具（如搜索 API）的密钥。
- `workspace_config.json`: 配置工作空间路径。

---

## 🤝 贡献与支持

欢迎提交 Issue 或 Pull Request 来完善这个项目！

---

## 📄 开源协议

本项目采用 [MIT License](LICENSE) 开源协议。

---

<p align="center">
  <i>Built with ❤️ for the AI Community</i>
</p>
