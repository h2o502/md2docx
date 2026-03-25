# md2docx — Markdown 转 Word 工具

将 Markdown 文件（`.md`）转换为格式精良的 Word 文档（`.docx`）。支持标题、表格、代码块语法高亮、Mermaid 图表渲染、引用块、粗体/斜体等全套 Markdown 语法，并对中文文档场景做了专项优化。

---

## 特性

- **全文微软雅黑**，正文 10.5pt，表格 9pt
- **动态标题字号**：自动检测文档最大标题层级，末级 12.5pt，每升一级 +2pt
- **表格**：表头灰底，整洁边框，支持 HTML 转义内容
- **代码块**：灰色背景 + 四边框 + Consolas 等宽字体，视觉还原度高
- **Mermaid 图表**：调用本地 `mmdc` 渲染为 PNG 并嵌入文档，支持自动等比缩放
- **行内格式**：`**粗体**`、`*斜体*`、`***粗斜体***`、`` `行内代码` ``、`[链接](url)` 均正确渲染
- **引用块**：`>` 引用使用 Word Quote 样式，支持多行
- **水平分割线**：`---` 转换为段落下边框
- **本地图片**：`![alt](path)` 自动插入图片
- **HTML 块**：支持内嵌 `<blockquote>`、`<p style="text-align:center">`、`<strong>` 等常见 HTML 标签

---

## 安装依赖

### Python 依赖

```bash
pip install python-docx pygments pillow
```

| 包            | 用途                                           |
| ------------- | ---------------------------------------------- |
| `python-docx` | 生成 Word 文档（必须）                         |
| `pygments`    | 代码块语法高亮（可选，缺失时回退为纯文本样式） |
| `pillow`      | 读取 Mermaid 渲染图片尺寸以自动缩放（可选）    |

### Node.js 依赖（Mermaid 图表支持）

```bash
cd ~/.workbuddy/skills/md2docx
npm install
```

安装完成后，`node_modules/.bin/mmdc.cmd` 将被自动调用来渲染 Mermaid 代码块。若不需要 Mermaid 支持，可跳过此步骤，图表将以文字占位符代替。

---

## 使用方法

```bash
python md2docx.py <输入文件.md> <输出文件.docx>
```

### 示例

```bash
# 基础用法
python md2docx.py document.md output.docx

# 指定完整路径
python md2docx.py D:/项目/report.md D:/项目/report.docx

# 输出到桌面
python md2docx.py README.md C:/Users/你的用户名/Desktop/README.docx
```

---

## 支持的 Markdown 语法

| 语法                     | 效果                             |
| ------------------------ | -------------------------------- |
| `# 标题` / `## 二级` …   | Word 标题样式，动态字号          |
| `**粗体**`               | 粗体                             |
| `*斜体*`                 | 斜体                             |
| `***粗斜体***`           | 粗斜体                           |
| `` `行内代码` ``         | 灰底等宽字体                     |
| ` ```python ` … ` ``` `  | 代码块，Consolas 字体 + 灰色背景 |
| ` ```mermaid ` … ` ``` ` | 渲染为 PNG 图片嵌入文档          |
| `\| 表格 \|`             | Word 表格，表头灰底              |
| `> 引用`                 | Word Quote 样式                  |
| `---`                    | 水平分割线                       |
| `- 列表项`               | Word 项目符号列表                |
| `![alt](path)`           | 嵌入本地图片                     |
| `[文字](url)`            | 显示为下划线文本                 |

---

## 文件结构

```
md2docx/
├── md2docx.py        # 主转换脚本
├── code_block.py     # 代码块样式渲染
├── SKILL.md          # WorkBuddy Skill 描述
├── README.md         # 本文档
├── package.json      # Node.js 依赖（Mermaid CLI）
└── package-lock.json # 依赖版本锁定
```

---

## 个性化配置

在 `md2docx.py` 顶部修改以下常量，可快速调整全局样式：

```python
FONT_NAME        = '微软雅黑'   # 全文字体
BODY_FONT_SIZE   = 10.5         # 正文字号（pt）
TABLE_FONT_SIZE  = 9            # 表格字号（pt）
MIN_HEADING_SIZE = 12.5         # 最末级标题字号（pt）
HEADING_SIZE_STEP = 2           # 每升一级标题递增字号（pt）
```

---

## 注意事项

- 输入文件编码须为 **UTF-8**
- Mermaid 图表渲染需要 Node.js 环境（建议 v18+）
- 图片路径支持相对路径（相对于 `.md` 文件所在目录）和绝对路径
- 若输出目录不存在，脚本会自动创建

---

## 作者

冯帆 / WorkBuddy 工作虾
