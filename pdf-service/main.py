"""線上簽約伺服器（Cloud Run，台灣 asia-east1）
POST /sign ：員工送出簽署 → 驗證 → 鎖定狀態 → 保存個資（僅伺服器可讀）→ 產生 PDF 存到 Cloud Storage → 更新簽約狀態
GET  /pdf  ：下載已簽署 PDF（同仁：Bearer 登入權杖；員工：token + 手機末 3 碼，簽署後 7 天內）
"""
import base64, datetime, io, os, re, logging
from flask import Flask, request, jsonify, Response
import firebase_admin
from firebase_admin import firestore, storage, auth as fauth
from PIL import Image
from filler import build_pdf

BUCKET = os.environ.get('BUCKET', 'recruit-system-904ba.firebasestorage.app')
firebase_admin.initialize_app(options={'storageBucket': BUCKET})
db = firestore.client()
bucket = storage.bucket()
ORIGINS = {'https://jeff19800102-crypto.github.io'}
COLS = {'test': ('test_contracts', 'test_ct_files', 'test_contract_private'),
        'prod': ('contracts', 'ct_files', 'contract_private')}
TZ = datetime.timezone(datetime.timedelta(hours=8))
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024
log = logging.getLogger('sign'); logging.basicConfig(level=logging.INFO)

@app.after_request
def cors(resp):
    o = request.headers.get('Origin', '')
    if o in ORIGINS:
        resp.headers['Access-Control-Allow-Origin'] = o
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
        resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    return resp

def now_tw():
    return datetime.datetime.now(TZ).strftime('%Y-%m-%dT%H:%M')

def err(msg, code):
    return jsonify({'ok': False, 'error': msg}), code

def tw_id_ok(s):
    if not re.match(r'^[A-Z][12]\d{8}$', s or ''): return False
    n = 'ABCDEFGHJKLMNPQRSTUVXYWZIO'.index(s[0]) + 10
    t = n // 10 + (n % 10) * 9 + sum(int(s[i]) * w for i, w in zip(range(1, 9), [8, 7, 6, 5, 4, 3, 2, 1])) + int(s[9])
    return t % 10 == 0

def clean(v, n):
    """字串：去控制字元、限長"""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', str(v or ''))[:n].strip()

def valid_token(t):
    return bool(re.match(r'^[A-Za-z0-9]{12,40}$', t or ''))

def parse_data(d):
    """員工填寫資料白名單與格式檢查；回傳 (資料, 錯誤訊息)"""
    out = {
        'name': clean(d.get('name'), 30), 'phone': re.sub(r'\D', '', clean(d.get('phone'), 20)),
        'sex': clean(d.get('sex'), 2), 'birthday': clean(d.get('birthday'), 10), 'idNo': clean(d.get('idNo'), 10).upper(),
        'tel': clean(d.get('tel'), 20), 'address': clean(d.get('address'), 120), 'email': clean(d.get('email'), 80),
        'bank': clean(d.get('bank'), 40), 'branch': clean(d.get('branch'), 30), 'account': re.sub(r'\D', '', clean(d.get('account'), 20)),
    }
    if not out['name']: return None, '姓名不可空白'
    if not re.match(r'^09\d{8}$', out['phone']): return None, '手機格式不正確'
    if out['sex'] not in ('男', '女'): return None, '請選擇性別'
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', out['birthday']): return None, '出生日期格式不正確'
    try:
        bd = datetime.date.fromisoformat(out['birthday'])
        if (datetime.date.today() - bd).days < 15 * 365: return None, '出生日期不正確（未滿 15 歲）'
    except ValueError: return None, '出生日期不正確'
    if not tw_id_ok(out['idNo']): return None, '身分證字號格式不正確'
    if not out['address']: return None, '請填寫通訊地址'
    if out['email'] and not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', out['email']): return None, 'Email 格式不正確'
    if not out['bank']: return None, '請選擇銀行'
    if not (10 <= len(out['account']) <= 16): return None, '帳號長度不正確'
    out['bankCode'] = out['bank'][:3] if re.match(r'^\d{3} ', out['bank']) else ''
    if out['bankCode'] == '700' and len(out['account']) != 14: return None, '郵局帳號應為 14 碼'
    return out, None

def parse_sig(s):
    try:
        raw = base64.b64decode(str(s or '').split(',')[-1], validate=True)
    except Exception: return None
    if not (200 <= len(raw) <= 500 * 1024) or not raw.startswith(b'\x89PNG'): return None
    try:
        im = Image.open(io.BytesIO(raw)); im.verify()
    except Exception: return None
    return raw

@app.route('/sign', methods=['POST', 'OPTIONS'])
def sign():
    if request.method == 'OPTIONS': return ('', 204)
    b = request.get_json(force=True, silent=True) or {}
    mode = 'test' if b.get('test') else 'prod'
    ccol, fcol, pcol = COLS[mode]
    token = str(b.get('token', ''))
    if not valid_token(token): return err('連結無效', 400)
    ref = db.collection(ccol).document(token)
    snap = ref.get()
    if not snap.exists: return err('找不到簽約資料', 404)
    c = snap.to_dict()
    phone = ((c.get('worker') or {}).get('phone') or '')
    if str(b.get('last3', '')) != phone[-3:] or len(phone) < 10:
        tries = int(c.get('tries') or 0) + 1
        ref.update({'tries': tries, **({'status': 'void', 'voidAt': now_tw(), 'voidNote': '末 3 碼錯誤超過 5 次'} if tries >= 5 and c.get('status') in ('sent', 'opened') else {})})
        return err('身分驗證失敗', 403)
    d, msg = parse_data(b.get('data') or {})
    if msg: return err(msg, 400)
    sig = parse_sig(b.get('signaturePng'))
    if not sig: return err('簽名圖無效，請重新簽名', 400)
    opts_in = b.get('opts') or {}
    opts = {'健保': clean(opts_in.get('健保'), 10), '勞退': clean(opts_in.get('勞退'), 2),
            '離職原因': [clean(x, 40) for x in (opts_in.get('離職原因') or []) if isinstance(x, str)][:20]}
    if opts['健保'] and opts['健保'] not in ('公司投保', '依附父母', '其他單位'): return err('健保選項不正確', 400)
    if opts['勞退'] and opts['勞退'] not in [str(i) for i in range(7)]: return err('勞退自提選項不正確', 400)
    snapdocs = (c.get('pkgSnap') or {}).get('docs') or []
    if not snapdocs: return err('這份簽約沒有設定文件範本', 400)

    # 以交易鎖定狀態：只有 sent/opened 且未過期才能進入 signing
    @firestore.transactional
    def lock(tx, ref):
        cur = ref.get(transaction=tx).to_dict() or {}
        if cur.get('status') not in ('sent', 'opened'): return cur.get('status')
        exp = cur.get('expiresTs')
        expired = (exp.timestamp() < datetime.datetime.now(datetime.timezone.utc).timestamp()) if hasattr(exp, 'timestamp') else ((cur.get('expiresAt') or '') < now_tw())
        if expired: return 'expired'
        tx.update(ref, {'status': 'signing', 'signingAt': now_tw()})
        return 'ok'
    st = lock(db.transaction(), ref)
    if st == 'signed': return err('這份簽約已完成', 409)
    if st == 'signing': return err('正在處理中，請稍候', 409)
    if st == 'expired': return err('連結已過期', 410)
    if st != 'ok': return err('這份簽約已作廢', 409)

    try:
        docs = [x for x in snapdocs if not x.get('optional') or str(opts.get(x['optional']['field'])) in x['optional']['in']]
        templates = []
        for x in docs:
            f = db.collection(fcol).document(x['fileId']).get()
            if not f.exists: raise RuntimeError(f"範本檔案遺失：{x['name']}")
            templates.append((x['name'], base64.b64decode(f.to_dict()['b64'])))
        today = now_tw()[:10]
        job = {k: clean(v, 120) for k, v in (c.get('job') or {}).items() if isinstance(v, (str, int, float))}
        values = {**job,
                  '員工姓名': d['name'], '手機': d['phone'], '市話': d['tel'], '身分證字號': d['idNo'], '性別': d['sex'], '生日': d['birthday'],
                  '通訊地址': d['address'], 'Email': d['email'], '銀行': d['bank'], '分行': d['branch'], '帳號': d['account'],
                  '帳戶說明': f"{d['bank']} {d['branch']} 帳號 {d['account']}", '帳戶類型': '郵局' if d['bankCode'] == '700' else '其它',
                  '健保': opts['健保'], '勞退自提': opts['勞退'] or '0', '離職原因': opts['離職原因'], '簽署日': today}
        pdf, sha = build_pdf(templates, values, sig)
        path = f"{mode}/contracts/{token}.pdf"
        bucket.blob(path).upload_from_string(pdf, content_type='application/pdf')
        db.collection(pcol).document(token).set({**d, 'savedAt': now_tw(), 'contractNo': c.get('contractNo', '')})
        ref.update({'status': 'signed', 'signedAt': now_tw(), 'signedTs': firestore.SERVER_TIMESTAMP, 'opts': opts, 'agreed': [clean(x, 60) for x in (b.get('agreed') or []) if isinstance(x, str)][:10],
                    'device': clean(b.get('device'), 180), 'ip': clean(request.headers.get('X-Forwarded-For', request.remote_addr), 60),
                    'pdfPath': path, 'pdfSha256': sha, 'docsSigned': [x['name'] for x in docs], 'nameEntered': d['name'] != (c.get('worker') or {}).get('name')})
        return jsonify({'ok': True, 'sha256': sha})
    except Exception as e:
        log.exception('sign failed %s', token)
        ref.update({'status': 'opened', 'signError': clean(str(e), 200), 'signErrorAt': now_tw()})
        return err('文件產生失敗，請稍後再試或聯絡招募人員', 500)

@app.route('/pdf', methods=['GET', 'OPTIONS'])
def pdf():
    if request.method == 'OPTIONS': return ('', 204)
    mode = 'test' if request.args.get('test') else 'prod'
    ccol = COLS[mode][0]
    token = request.args.get('t', '')
    if not valid_token(token): return err('連結無效', 400)
    snap = db.collection(ccol).document(token).get()
    if not snap.exists or snap.to_dict().get('status') != 'signed': return err('尚未簽署', 404)
    c = snap.to_dict()
    authz = request.headers.get('Authorization', '')
    ok = False
    if authz.startswith('Bearer '):
        try: fauth.verify_id_token(authz[7:]); ok = True
        except Exception: return err('請重新登入', 401)
    else:
        k = request.args.get('k', '')
        phone = (c.get('worker') or {}).get('phone') or ''
        signed = c.get('signedAt') or ''
        fresh = signed and (datetime.datetime.now(TZ) - datetime.datetime.strptime(signed, '%Y-%m-%dT%H:%M').replace(tzinfo=TZ)).days <= 7
        ok = bool(k) and k == phone[-3:] and fresh
    if not ok: return err('無權限下載，請登入後台或聯絡招募人員', 403)
    data = bucket.blob(c['pdfPath']).download_as_bytes()
    fn = re.sub(r'[^A-Za-z0-9\-]', '', c.get('contractNo', 'contract')) or 'contract'
    return Response(data, mimetype='application/pdf', headers={'Content-Disposition': f'attachment; filename={fn}.pdf'})

@app.get('/')
def health():
    return 'ok'
