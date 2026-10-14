"""Word 範本工具：在段落內以整段文字取代（保留第一個文字片段的格式），支援 {{欄位}}、{{chk:欄位=值}}、{{簽名}}"""
import copy, re
from docx.oxml.ns import qn

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
        for part in (s.header, s.footer):
            yield from walk(part)

def set_text(p, text):
    """把段落文字整段換成 text，格式沿用第一個有文字的 run"""
    runs = p.runs
    if not runs:
        p.add_run(text); return
    keep = next((r for r in runs if r.text.strip()), runs[0])
    rpr = keep._r.find(qn('w:rPr'))
    first = runs[0]
    for r in runs[1:]:
        if r._r.findall('.//' + qn('w:drawing')) or r._r.findall('.//' + qn('w:pict')): continue  # 保留圖片（印章）
        r._r.getparent().remove(r._r)
    old = first._r.find(qn('w:rPr'))
    if old is not None: first._r.remove(old)
    if rpr is not None: first._r.insert(0, copy.deepcopy(rpr))
    first.text = text

def replace_in_paragraph(p, old, new, count=1):
    t = p.text
    if old not in t: return False
    set_text(p, t.replace(old, new, count)); return True
