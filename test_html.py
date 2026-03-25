#!/usr/bin/env python3
"""用 Pygments 生成语法高亮 HTML，再用 playwright 截图转 PNG"""
import os, tempfile
from pygments.formatters import HtmlFormatter
from pygments.lexers import get_lexer_by_name, TextLexer
from pygments import highlight
from docx import Document
from docx.shared import Pt, Inches

# 生成语法高亮 HTML（带行号）
def generate_code_html(code_text, language=None):
    if language and language not in ('mermaid', ''):
        try:
            lexer = get_lexer_by_name(language)
        except Exception:
            lexer = TextLexer()
    else:
        lexer = TextLexer()

    formatter = HtmlFormatter(
        nowrap=False,
        cssclass='highlight',
        style='monokai',
        linenos=True,
        full=True,
    )
    return highlight(code_text, lexer, formatter)

code = """def rag_pipeline(query):
    docs = vector_db.search(query, top_k=10)
    context = format_docs(docs)
    response = llm.generate(context)
    return response"""

html = generate_code_html(code, 'python')
tmp_html = os.path.join(tempfile.gettempdir(), 'test_code.html')
with open(tmp_html, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'HTML generated: {tmp_html}')
print('HTML preview:', html[:300])
