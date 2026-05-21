---
name: "md2docx"
description: "将 Markdown 文件转换为 Word 文档，支持 Mermaid 图表渲染、代码高亮、表格、中文排版优化"
---

# md2docx — Markdown 转 Word

将 `.md` 文件转换为格式精良的 `.docx` 文档。

## 用法

```bash
python md2docx.py <输入文件.md> <输出文件.docx>
```

## 功能

- **中文排版**：全文微软雅黑，正文 10.5pt，表格 9pt
- **动态标题字号**：自动检测标题层级，末级 12.5pt，每升一级 +2pt
- **Mermaid 图表**：自动渲染为 PNG 嵌入文档，动态检测系统 Chrome/Edge，无需手动配置
- **代码块**：灰色背景 + 四边框 + Consolas 等宽字体
- **行内格式**：`**粗体**`、`*斜体*`、`` `代码` ``、`[链接](url)`
- **表格**：表头灰底，支持 HTML 转义
- **引用块 / 分割线 / 图片 / HTML 块**

## 依赖

### Python

```bash
pip install python-docx pygments pillow
```

| 包 | 必须 | 用途 |
|---|---|---|
| `python-docx` | 是 | 生成 Word 文档 |
| `pygments` | 否 | 代码块语法高亮（缺失时回退纯文本） |
| `pillow` | 否 | Mermaid 图片自动缩放（缺失时用默认宽度） |

### Node.js（Mermaid 支持）

如需渲染 Mermaid 图表：

```bash
cd <skill目录>
npm install
```

安装后 `node_modules/.bin/mmdc.cmd` 会被自动调用。

**浏览器配置**：脚本会自动检测系统已安装的 Chrome / Edge / Chromium，无需手动下载 chrome-headless-shell。若系统无任何浏览器，Mermaid 图表将以文字占位符代替。

若不需要 Mermaid 支持，可跳过 `npm install`。

## 个性化配置

在 `md2docx.py` 顶部修改常量：

```python
FONT_NAME         = '微软雅黑'  # 全文字体
BODY_FONT_SIZE    = 10.5        # 正文字号
TABLE_FONT_SIZE   = 9           # 表格字号
MIN_HEADING_SIZE  = 12.5        # 最末级标题字号
HEADING_SIZE_STEP = 2           # 每升一级递增字号
```

## 文件结构

```
md2docx/
├── md2docx.py        # 主转换脚本
├── code_block.py     # 代码块样式渲染
├── SKILL.md          # Skill 描述（本文件）
├── README.md         # 详细文档
└── package.json      # Node.js 依赖（Mermaid CLI）
```
