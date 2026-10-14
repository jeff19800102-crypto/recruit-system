"""Word 範本填入資料 → 轉 PDF → 合併（Cloud Run 與本機共用）"""
import copy, io, os, re, subprocess, tempfile, shutil, hashlib
import docx
from docx.shared import Cm
from docx.oxml.ns import qn
from docx_tools import iter_paragraphs, set_text

TAG = re.compile(r'\{\{([^}]+)\}\}')
DATE_KEYS = ['到職日', '最後工作日', '生日', '簽署日']

def expand_dates(v):
    out = dict(v)
    for k in DATE_KEYS:
        val = v.get(k)
        if isinstance(val, str) and re.match(r'^\d{4}-\d{2}-\d{2}$', val):
            y, m, d = [int(x) for x in val.split('-')]
            out[k + '_民國年'] = str(y - 1911); out[k + '_月'] = str(m); out[k + '_日'] = str(d)
            out[k + '_民國'] = f'{y - 1911}年{m}月{d}日'
    return out

def render_text(text, v):
    def rep(m):
        key = m.group(1).strip()
        if key.startswith('chk:'):
            k, want = key[4:].split('=', 1)
            have = v.get(k)
            on = (want in have) if isinstance(have, list) else (str(have) == want)
            return '■' if on else '□'
        if key == '簽名': return m.group(0)
        return str(v.get(key, ''))
    return TAG.sub(rep, text)

def put_signature(p, sig_png, height_cm=1.2):
    before, after = p.text.split('{{簽名}}', 1)
    set_text(p, before)
    rpr = p.runs[0]._r.find(qn('w:rPr')) if p.runs else None
    p.add_run().add_picture(io.BytesIO(sig_png), height=Cm(height_cm))
    if after:
        r = p.add_run(after)
        if rpr is not None: r._r.insert(0, copy.deepcopy(rpr))

def fill_docx(template_bytes, values, sig_png):
    d = docx.Document(io.BytesIO(template_bytes))
    for p in list(iter_paragraphs(d)):
        t = p.text
        if '{{' not in t: continue
        new = render_text(t, values)
        if new != t: set_text(p, new)
        if '{{簽名}}' in p.text: put_signature(p, sig_png)
    out = io.BytesIO(); d.save(out); return out.getvalue()

def _convert(tmp, i, name, data, v, sig_png):
    src = os.path.join(tmp, f'{i:02d}.docx')
    open(src, 'wb').write(fill_docx(data, v, sig_png))
    r = subprocess.run(['soffice', '--headless', '--norestore', f'-env:UserInstallation=file://{tmp}/lo', '--convert-to', 'pdf', '--outdir', tmp, src],
                       capture_output=True, text=True, timeout=120)
    out = os.path.join(tmp, f'{i:02d}.pdf')
    if r.returncode != 0 or not os.path.exists(out):
        raise RuntimeError(f'轉檔失敗 {name}: {r.stderr[-300:]}')
    return out

def build_pdf(templates, values, sig_png):
    """templates: [(檔名, docx bytes)]；回傳 (合併後 PDF bytes, sha256)"""
    v = expand_dates(values)
    tmp = tempfile.mkdtemp()
    try:
        pdfs = [_convert(tmp, i, n, d, v, sig_png) for i, (n, d) in enumerate(templates)]
        merged = os.path.join(tmp, 'merged.pdf')
        if len(pdfs) == 1: shutil.copy(pdfs[0], merged)
        else: subprocess.run(['pdfunite', *pdfs, merged], check=True)
        b = open(merged, 'rb').read()
        return b, hashlib.sha256(b).hexdigest()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def build_previews(templates, values, sig_png, dpi=72, quality=55):
    """每份文件各自轉成頁面 JPEG；回傳 [(name, [jpeg bytes...])]"""
    v = expand_dates(values)
    tmp = tempfile.mkdtemp()
    try:
        out = []
        for i, (n, d) in enumerate(templates):
            pdf = _convert(tmp, i, n, d, v, sig_png)
            prefix = os.path.join(tmp, f'p{i:02d}')
            subprocess.run(['pdftoppm', '-r', str(dpi), '-jpeg', '-jpegopt', f'quality={quality}', pdf, prefix], check=True, timeout=120)
            pages = sorted(f for f in os.listdir(tmp) if f.startswith(f'p{i:02d}') and f.endswith('.jpg'))
            out.append((n, [open(os.path.join(tmp, f), 'rb').read() for f in pages]))
        return out
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
