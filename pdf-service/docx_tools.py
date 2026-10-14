"""Word 範本工具：段落內文字替換，保留圖片（印章）、分頁符等非文字節點。
做法：把段落內「文字 run」的文字合併到第一個文字 run，其餘文字 run 清空（不刪 run，也不動含圖片的 run）。
"""
import copy
from docx.oxml.ns import qn

NON_TEXT = ('w:drawing', 'w:pict', 'w:sym', 'w:fldChar', 'w:instrText')  # 含這些的 run 原位保留；tab/br 由 python-docx 以 \t \n 重建

def iter_paragraphs(doc):
    def walk(container):
        for p in container.paragraphs: yield p
        for t in getattr(container, 'tables', []):
            seen = []
            for row in t.rows:
                for c in row.cells:
                    if any(c._tc is x for x in seen): continue
                    seen.append(c._tc); yield from walk(c)
    yield from walk(doc)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header, s.even_page_footer):
            try: yield from walk(part)
            except Exception: pass

def _has_special(r):
    return any(r._r.findall('.//' + qn(tag)) for tag in NON_TEXT)

def set_text(p, text):
    """把段落的文字整段換成 text；含圖片／分頁符／tab 的 run 原位保留。"""
    runs = p.runs
    text_runs = [r for r in runs if not _has_special(r)]
    if not text_runs:
        r = p.add_run(text); return
    keep = next((r for r in text_runs if r.text.strip()), text_runs[0])
    rpr = keep._r.find(qn('w:rPr'))
    first = text_runs[0]
    for r in text_runs:
        for tag in ('w:t', 'w:tab', 'w:br', 'w:cr'):
            for t in r._r.findall(qn(tag)): r._r.remove(t)
    if rpr is not None and first._r.find(qn('w:rPr')) is None: first._r.insert(0, copy.deepcopy(rpr))
    first.text = text  # 只在第一個文字 run 內寫入（此 run 無特殊節點）

def replace_in_paragraph(p, old, new, count=1):
    t = p.text
    if old not in t: return False
    set_text(p, t.replace(old, new, count)); return True
