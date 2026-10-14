"""線上簽約伺服器（Cloud Run，台灣 asia-east1）
POST /sign ：員工送出簽署 → 驗證 → 保存個資（僅伺服器可讀）→ 產生 PDF 存到 Cloud Storage → 更新簽約狀態
GET  /pdf  ：下載已簽署 PDF（同仁需登入；員工用自己的簽約連結代碼）
"""
import base64, datetime, os, re
from flask import Flask, request, jsonify, Response
import firebase_admin
from firebase_admin import firestore, storage, auth as fauth
from filler import build_pdf

BUCKET = os.environ.get('BUCKET', 'recruit-system-904ba.firebasestorage.app')
firebase_admin.initialize_app(options={'storageBucket': BUCKET})
db = firestore.client()
bucket = storage.bucket()
ORIGINS = {'https://jeff19800102-crypto.github.io'}
COLS = {'test': ('test_contracts', 'test_ct_files', 'test_contract_private'),
        'prod': ('contracts', 'ct_files', 'contract_private')}
app = Flask(__name__)

@app.after_request
def cors(resp):
    o = request.headers.get('Origin', '')
    if o in ORIGINS:
        resp.headers['Access-Control-Allow-Origin'] = o
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
        resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return resp

def now_tw():
    return (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime('%Y-%m-%dT%H:%M')

def err(msg, code):
    return jsonify({'ok': False, 'error': msg}), code

def tw_id_ok(s):
    if not re.match(r'^[A-Z][12]\d{8}$', s or ''): return False
    n = 'ABCDEFGHJKLMNPQRSTUVXYWZIO'.index(s[0]) + 10
    t = n // 10 + (n % 10) * 9 + sum(int(s[i]) * w for i, w in zip(range(1, 9), [8, 7, 6, 5, 4, 3, 2, 1])) + int(s[9])
    return t % 10 == 0

@app.route('/sign', methods=['POST', 'OPTIONS'])
def sign():
    if request.method == 'OPTIONS': return ('', 204)
    b = request.get_json(force=True, silent=True) or {}
    mode = 'test' if b.get('test') else 'prod'
    ccol, fcol, pcol = COLS[mode]
    token = str(b.get('token', ''))
    if not re.match(r'^[A-Za-z0-9]{12,40}$', token): return err('連結無效', 400)
    ref = db.collection(ccol).document(token)
    snap = ref.get()
    if not snap.exists: return err('找不到簽約資料', 404)
    c = snap.to_dict()
    if c.get('status') not in ('sent', 'opened'): return err('這份簽約已完成或已作廢', 409)
    if c.get('expiresAt', '') < now_tw(): return err('連結已過期', 410)
    if str(b.get('last3', '')) != c['worker']['phone'][-3:]: return err('身分驗證失敗', 403)
    d = b.get('data') or {}
    if not tw_id_ok(str(d.get('idNo', '')).upper()): return err('身分證字號格式不正確', 400)
    sig = base64.b64decode(str(b.get('signaturePng', '')).split(',')[-1] or b'')
    if len(sig) < 200: return err('沒有簽名', 400)
    opts = b.get('opts') or {}
    snapdocs = (c.get('pkgSnap') or {}).get('docs') or []
    if not snapdocs: return err('這份簽約沒有設定文件範本', 400)
    docs = [x for x in snapdocs if not x.get('optional') or str(opts.get(x['optional']['field'])) in x['optional']['in']]
    templates = []
    for x in docs:
        f = db.collection(fcol).document(x['fileId']).get()
        if not f.exists: return err(f"範本檔案遺失：{x['name']}", 500)
        templates.append((x['name'], base64.b64decode(f.to_dict()['b64'])))
    today = now_tw()[:10]
    values = {**c.get('job', {}),
              '員工姓名': d.get('name', ''), '手機': d.get('phone', ''), '市話': d.get('tel', ''), '身分證字號': str(d.get('idNo', '')).upper(),
              '性別': d.get('sex', ''), '生日': d.get('birthday', ''), '通訊地址': d.get('address', ''), 'Email': d.get('email', ''),
              '銀行': d.get('bank', ''), '分行': d.get('branch', ''), '帳號': d.get('account', ''),
              '帳戶說明': f"{d.get('bank', '')} {d.get('branch', '')} 帳號 {d.get('account', '')}",
              '帳戶類型': '郵局' if d.get('bankCode') == '700' else '其它',
              '健保': opts.get('健保', ''), '勞退自提': str(opts.get('勞退', '0')), '離職原因': opts.get('離職原因', []), '簽署日': today}
    pdf, sha = build_pdf(templates, values, sig)
    path = f"{mode}/contracts/{c['contractNo']}.pdf"
    bucket.blob(path).upload_from_string(pdf, content_type='application/pdf')
    db.collection(pcol).document(token).set({**d, 'idNo': values['身分證字號'], 'savedAt': now_tw(), 'contractNo': c['contractNo']})
    ref.update({'status': 'signed', 'signedAt': now_tw(), 'opts': opts, 'agreed': b.get('agreed', []),
                'device': str(b.get('device', ''))[:180], 'ip': request.headers.get('X-Forwarded-For', request.remote_addr),
                'pdfPath': path, 'pdfSha256': sha, 'docsSigned': [x['name'] for x in docs]})
    return jsonify({'ok': True, 'sha256': sha})

@app.route('/pdf', methods=['GET', 'OPTIONS'])
def pdf():
    if request.method == 'OPTIONS': return ('', 204)
    mode = 'test' if request.args.get('test') else 'prod'
    ccol = COLS[mode][0]
    token = request.args.get('t', '')
    if not re.match(r'^[A-Za-z0-9]{12,40}$', token): return err('連結無效', 400)
    c = db.collection(ccol).document(token).get()
    if not c.exists or c.to_dict().get('status') != 'signed': return err('尚未簽署', 404)
    authz = request.headers.get('Authorization', '')
    if authz.startswith('Bearer '):
        try: fauth.verify_id_token(authz[7:])
        except Exception: return err('請重新登入', 401)
    data = bucket.blob(c.to_dict()['pdfPath']).download_as_bytes()
    return Response(data, mimetype='application/pdf', headers={'Content-Disposition': f"inline; filename={c.to_dict()['contractNo']}.pdf"})

@app.get('/')
def health():
    return 'ok'
