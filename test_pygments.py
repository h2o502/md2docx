from pygments.formatters import SvgFormatter
from pygments.lexers import PythonLexer
from pygments import highlight
import os

code = 'def hello():\n    print("hello")'
svg = highlight(code, PythonLexer(), SvgFormatter(linenos=True, fontface='Courier New'))
print('SVG generated, length:', len(svg))
print(svg[:500])

out_path = os.path.join(os.path.dirname(__file__), 'test_code.svg')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(svg)
print('Saved to:', out_path)
