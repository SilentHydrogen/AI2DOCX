# AI2DOCX — AI 内容一键转 Word

将 AI 对话内容（DeepSeek / ChatGPT / Kimi / 豆包等）一键导出为排版精美的 Word 文档。支持 **LaTeX 数学公式** → **Word 原生可编辑公式**、**标题层级自动识别**、**表格自动转换**、**中英混排字体**。

对标 "向前文档解析" (docx.jackweb.cc) 的本地命令行版本，**纯本地处理，数据不上传**，**无使用次数限制**。

## 效果预览

| 输入（AI 对话复制） | 输出（Word 文档） |
|:---|:---|
| Markdown 标题 → | Word 多级标题（仿宋_GB2312 + Times New Roman 混排） |
| `$$...$$` / `\(...\)` 数学公式 → | Word 原生 OMML 可编辑公式 |
| Markdown 表格 / HTML 表格 → | Word 表格（蓝色主题边框 + 表头） |
| 代码块 → | Consolas 等宽字体 + 灰色背景 |
| 纯中文无 `#` 标题 → | 自动检测并转为对应级别标题 |

## 快速开始

### 前置要求

- **Python 3.8+**（推荐 3.13）
- **Pandoc** — Markdown+LaTeX 转 Word 引擎（[下载安装](https://pandoc.org)）
- **python-docx** — Word 文档后处理
- **pyperclip** — 剪贴板读取

```bash
pip install python-docx pyperclip
```

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

对纯中文文本（无 `#` 标记），自动识别标题层级：

| 格式 | 标题级别 |
|:---|:---:|
| 一、二、三、... | H2 |
| 第一章 / 第二节 /... | H2 |
| 依据 N：/ 方法 N：/... | H3 |
| 方案 A：/ 方案 B：/... | H4 |

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
├── ai2docx.py          # 主脚本（~1200 行）
├── 转换.bat             # Windows 一键转换（双击运行）
├── README.md            # 本文件
├── 测试示例.md           # 含公式 + 代码块的测试文件
└── .gitignore
```

## 依赖

已内置的系统环境：
- **Python 3.13**（[官网](https://python.org)）
- **Pandoc**（[下载](https://pandoc.org)）
- **python-docx**（`pip install python-docx`）
- **pyperclip**（`pip install pyperclip`）

## 与 web 版对比

| 功能 | 向前文档解析 (docx.jackweb.cc) | AI2DOCX (本地脚本) |
|------|:---:|:---:|
| AI 内容 → Word | ✅ | ✅ |
| LaTeX 公式 → Word 公式 | ✅ | ✅ (Pandoc 引擎) |
| 标题层级保留 | ✅ | ✅ (含纯中文文本检测) |
| 纯本地处理 | ✅ 浏览器端 | ✅ 命令行 |
| 无需联网 | ✅ | ✅ |
| 无使用次数限制 | ❌ (有免费额度) | ✅ 完全免费 |
| 批量转换 | ❌ | ✅ 可脚本化 |
| Mermaid 图表 | ✅ | ✅ (保留源码) |
| 表格自动修复 | ❌ | ✅ (三层兜底) |
| 中英混排字体 | ❌ | ✅ |
| PDF 导出 | ✅ | ❌ (仅 .docx) |

## 许可证

MIT
