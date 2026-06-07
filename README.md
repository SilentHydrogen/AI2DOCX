# AI2DOCX — AI 内容一键转 Word

将 AI 对话内容（DeepSeek / ChatGPT / Kimi / 豆包等）一键导出为排版精美的 Word 文档。支持 **LaTeX 数学公式** → **Word 原生可编辑公式**、**标题层级自动识别**、**表格自动转换**、**中英混排字体**。

**纯本地处理，数据不上传**，**无使用次数限制**。

## 效果预览

| 输入（AI 对话复制） | 输出（Word 文档） |
|:---|:---|
| Markdown 标题 → | Word 多级标题（仿宋_GB2312 + Times New Roman 混排） |
| `$$...$$` / `\(...\)` 数学公式 → | Word 原生 OMML 可编辑公式 |
| Markdown 表格 / HTML 表格 → | Word 表格（蓝色主题边框 + 表头） |
| 代码块 → | Consolas 等宽字体 + 灰色背景 |
| 纯中文无 `#` 标题 → | 自动检测并转为对应级别标题 |

## 环境配置

### 所需软件

| 软件 | 版本要求 | 用途 | 安装方式 |
|:---|:---:|:---|:---|
| **Python** | ≥ 3.8 | 运行脚本 | [python.org](https://python.org) 或微软商店 |
| **Pandoc** | ≥ 2.18 | 核心转换引擎（Markdown+LaTeX → Word OMML 公式） | [pandoc.org](https://pandoc.org/installing.html) |
| **python-docx** | 最新 | Word 文档后处理 | `pip install python-docx` |
| **pyperclip** | 最新 | 剪贴板读取 | `pip install pyperclip` |

> **💡 提示**：Pandoc 是**必装**的。如果不装，脚本会降级为纯 Python 模式，**LaTeX 公式将显示为源码**而非可编辑公式。

### 一键安装

```bash
# 安装 Python 依赖包（二选一）
pip install python-docx pyperclip

# 或者用 pip3
pip3 install python-docx pyperclip
```

Pandoc 需要单独下载安装：

- **Windows**：从 [pandoc.org/releases](https://github.com/jgm/pandoc/releases) 下载 `.msi` 安装包，双击安装
- **macOS**：`brew install pandoc`
- **Linux**：`sudo apt install pandoc`（Ubuntu/Debian）或 `sudo dnf install pandoc`（Fedora）

安装完成后验证：

```bash
python --version      # 应显示 Python 3.8+
pandoc --version      # 应显示 Pandoc 2.18+
```

### 验证安装

```bash
python ai2docx.py -h
```

如果一切正常，会显示帮助信息，并检测到 Pandoc 可用。

### 方式一：从剪贴板（最常用）

1. 在 AI 对话中点击「复制」按钮
2. 运行：

```bash
python ai2docx.py -o 我的文档.docx
```

> Windows 用户也可以双击 **`转换.bat`**（自动生成带日期的文件名）

### 方式二：从 Markdown 文件

```bash
python ai2docx.py 输入文件.md 输出文档.docx
```

### 方式三：简易模式

```bash
python ai2docx.py 输入文件.md
# 自动输出 output.docx
```

### 调试模式

```bash
python ai2docx.py --debug
```

输出诊断信息 + 保存预处理后的 Markdown 到 `.debug.md`，便于排查问题。

## 工作原理

```
剪贴板 / Markdown 文件
        │
        ▼
┌─────────────────────────────┐
│  预处理（_preprocess_content）│
│  ├─ 中文纯文本标题 → # 标题  │
│  ├─ 表格检测（三层兜底）      │
│  │  ├─ 管道符表格（|...|）   │
│  │  ├─ 制表符表格（TSV）     │
│  │  └─ HTML 剪贴板表格       │
│  └─ 公式转换（\(…\) → $…$） │
└─────────────────────────────┘
        │
        ▼
┌─────────────────────────────┐
│  Pandoc 转换引擎             │
│  -f markdown+tex_math_dollars│
│  -t docx                    │
│  → Word 文档（含 OMML 公式） │
└─────────────────────────────┘
        │
        ▼
┌─────────────────────────────┐
│  python-docx 后处理          │
│  ├─ 中英混排字体             │
│  ├─ 表格蓝色主题样式         │
│  ├─ "至少"行距（公式自适应）  │
│  └─ 页边距（国标论文格式）    │
└─────────────────────────────┘
        │
        ▼
    ✅ 输出 .docx
```

### 表格检测三层兜底

很多 AI 平台复制表格时结构会丢失，脚本采用三层检测：

| 层级 | 检测对象 | 示例 |
|:---|:---|:---|
| **① 管道符** | `\| a \| b \|` 格式 | 标准 Markdown 表格，自动补分隔行 |
| **② TSV** | 制表符分隔 | 从网页复制的 HTML 表格纯文本版 |
| **③ HTML** | Windows HTML 剪贴板 | 纯文本完全丢失时，直接从 `<table>` 提取 |

### 标题自动识别

**三种途径识别标题层级**（按优先级）：

1. **Markdown 标记**：`# ## ### ####` — Pandoc 原生支持
2. **HTML 剪贴板**：从 AI 网页复制时，`<h1>`~`<h6>` 标签自动转为对应级别
3. **纯文本智能检测**（无 `#` 标记时）：

| 格式 | 标题级别 | 示例 |
|:---|:---:|:---|
| 一、二、三、... | **H2** | `一、研究背景` |
| 第一章 / 第二节 /... | **H2** | `第一章 绪论` |
| （一）(一)... | **H2** | `（一）实验方法` |
| 1. 概述 / 1、概述 | **H2** | `1. 架构设计` |
| 1.1 背景 / 2.1 方法 | **H3** | `1.1 研究现状` |
| 1.1.1 细节 | **H4** | `1.1.1 参数设置` |
| 依据 N：/ 方法 N：/... | **H3** | `方法 1：梯度下降` |
| 方案 A：/ 方案 B：/... | **H4** | `方案 A：随机采样` |

> **智能防误判**：连续 3 条以上的编号（如操作步骤）会被识别为有序列表，不会错误转为标题。

### 文档样式

| 元素 | 中文字体 | 英文/数字字体 | 字号 | 行距 |
|:---|:---|:---|:---:|:---:|
| 一级标题 | 仿宋_GB2312 | Times New Roman | **18pt** | 段后 8pt |
| 二级标题 | 仿宋_GB2312 | Times New Roman | **16pt** | 段后 8pt |
| 三级标题 | 仿宋_GB2312 | Times New Roman | **14pt** | 段后 8pt |
| 正文 | 仿宋_GB2312 | Times New Roman | **11pt** | 至少 24pt |
| 代码 | Consolas | Consolas | 9.5pt | 固定 16pt |
| 表格表头 | 仿宋_GB2312 | Times New Roman | 10pt | — |
| 公式 | Cambria Math | Cambria Math | 与正文一致 | 自适应 |

## 项目结构

```
AI2DOCX/
├── ai2docx.py          # 主脚本（~1500 行）
├── 转换.bat             # Windows 一键转换（双击运行）
├── README.md            # 本文件
├── 测试示例.md           # 含公式 + 代码块的测试文件
└── .gitignore
```

## 许可证

MIT
