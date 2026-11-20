"""<!--INLINE:path--> 주석을 실제 CSS/JS 내용으로 바꿔 단일 HTML 파일을 만든다."""
import os, re, sys
src, out = sys.argv[1], sys.argv[2]
base = os.path.dirname(os.path.abspath(src))
def rep(m):
    p = os.path.join(base, m.group(1)); c = open(p, encoding='utf-8').read()
    return f'<style>\n{c}\n</style>' if p.endswith('.css') else f'<script>\n{c}\n</script>'
html = re.sub(r'<!--INLINE:(.+?)-->', rep, open(src, encoding='utf-8').read())
open(out, 'w', encoding='utf-8').write(html); print(out, len(html))
