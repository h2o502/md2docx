#!/usr/bin/env python3
"""测试: Pygments SVG → Pillow PNG → Word"""
import os, tempfile
from io import BytesIO
from pygments.formatters import SvgFormatter
from pygments.lexers import get_lexer_by_name
from pygments import highlight
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches

code = """def rag_pipeline(query):
    docs = vector_db.search(query, top_k=10)
    context = format_docs(docs)
    response = llm.generate(context)
    return response"""

lexer = get_lexer_by_name("python")
formatter = SvgFormatter(linenos=True, fontface='Courier New', fontsize='12px')
svg_data = highlight(code, lexer, formatter)
print(f'SVG length: {len(svg_data)}')

tmp_svg = os.path.join(tempfile.gettempdir(), 'test_code.svg')
with open(tmp_svg, 'w', encoding='utf-8') as f:
    f.write(svg_data)

# 用 Pillow 打开 SVG（需要 rsvg 或内置的）
try:
    img = Image.open(tmp_svg)
    print(f'Image size: {img.size}, mode: {img.mode}')
    tmp_png = os.path.join(tempfile.gettempdir(), 'test_code.png')
    img.save(tmp_png, 'PNG')
    print(f'PNG saved: {tmp_png}')
except Exception as e:
    print(f'Pillow SVG failed: {e}')

    # 备选：downgrade Pillow 方案
    print('Trying downgrade...')
