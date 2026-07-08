"""
نظام جرد وإدارة الأصول التقنية — سيرفر محلي على شبكة الجمعية
=================================================================
كيف يشتغل:
- يشغّله شخص واحد (مسؤول تقنية المعلومات) على جهاز واحد ثابت متصل بالشبكة.
- أي موظف على نفس الشبكة يفتح الرابط من متصفحه العادي (بدون تثبيت أي شيء).
- البيانات تُخزّن في ملفات JSON داخل مجلد data/ على نفس الجهاز — لا إنترنت، لا خدمات خارجية.

قبل التشغيل:
1) عدّل كلمتي المرور بالأسفل (ACCOUNTS) لكلمات مرورك الفعلية.
2) شغّل: pip install flask
3) شغّل: python app.py
4) اعرف عنوان IP لهذا الجهاز على الشبكة (راجع ملف التعليمات المرفق)
5) شارك الرابط: http://<عنوان-IP-لهذا-الجهاز>:5000 مع الموظفين على نفس الشبكة
"""

from flask import Flask, jsonify, request, session, send_from_directory
from werkzeug.security import generate_password_hash, check_password_hash
import json
import os
import uuid
import secrets

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
STATIC_DIR = os.path.join(BASE_DIR, 'static')
SECRET_KEY_FILE = os.path.join(BASE_DIR, 'secret.key')

os.makedirs(DATA_DIR, exist_ok=True)

app = Flask(__name__, static_folder=STATIC_DIR)

# ------------------------------------------------------------------
# مفتاح تشفير الجلسات — يُنشأ تلقائيًا أول مرة ويبقى ثابتًا بعدها
# ------------------------------------------------------------------
if os.path.exists(SECRET_KEY_FILE):
    with open(SECRET_KEY_FILE, 'r') as f:
        app.secret_key = f.read().strip()
else:
    key = secrets.token_hex(32)
    with open(SECRET_KEY_FILE, 'w') as f:
        f.write(key)
    app.secret_key = key

# ------------------------------------------------------------------
# ⚠️ كلمتا المرور تُقرآن من accounts.json (ملف محلي على هذا الجهاز فقط،
# غير مرفوع لأي مستودع). عدّل كلمتي المرور من ذاك الملف مباشرة.
# أول تشغيل على جهاز جديد ينشئ الملف تلقائيًا بكلمتي مرور مؤقتة (CHANGE_ME)
# لازم تغييرها فورًا.
# ------------------------------------------------------------------
ACCOUNTS_FILE = os.path.join(BASE_DIR, 'accounts.json')
DEFAULT_ACCOUNTS = {
    'administrator': {'password': 'CHANGE_ME_ADMIN_PASSWORD', 'role': 'admin'},
    'admin': {'password': 'CHANGE_ME_USER_PASSWORD', 'role': 'user'},
}

if os.path.exists(ACCOUNTS_FILE):
    with open(ACCOUNTS_FILE, 'r', encoding='utf-8') as f:
        raw_accounts = json.load(f)
else:
    raw_accounts = DEFAULT_ACCOUNTS
    with open(ACCOUNTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(raw_accounts, f, ensure_ascii=False, indent=2)

ACCOUNTS = {
    username: {
        'password_hash': generate_password_hash(info['password']),
        'role': info['role'],
    }
    for username, info in raw_accounts.items()
}

TABLES = ['employees', 'devices', 'switches', 'routers', 'servers', 'security', 'doors', 'secure']
ADMIN_ONLY_TABLES = {'secure'}


def data_path(table):
    return os.path.join(DATA_DIR, f'{table}.json')


def load_table(table):
    path = data_path(table)
    if not os.path.exists(path):
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def save_table(table, records):
    path = data_path(table)
    tmp_path = path + '.tmp'
    with open(tmp_path, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)  # كتابة آمنة تمنع تلف الملف عند انقطاع الكهرباء أثناء الحفظ


def current_role():
    return session.get('role')


def is_logged_in():
    return 'username' in session


# ------------------------------------------------------------------
# تسجيل الدخول / الخروج / التحقق من الجلسة
# ------------------------------------------------------------------
@app.route('/api/login', methods=['POST'])
def login():
    body = request.get_json(silent=True) or {}
    username = (body.get('username') or '').strip()
    password = body.get('password') or ''

    account = ACCOUNTS.get(username)
    if not account or not check_password_hash(account['password_hash'], password):
        return jsonify({'error': 'بيانات الدخول غير صحيحة'}), 401

    session['username'] = username
    session['role'] = account['role']
    return jsonify({'username': username, 'role': account['role']})


@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'success': True})


@app.route('/api/session', methods=['GET'])
def get_session():
    if not is_logged_in():
        return jsonify({'loggedIn': False})
    return jsonify({'loggedIn': True, 'username': session['username'], 'role': session['role']})


# ------------------------------------------------------------------
# بيانات الجداول (موظفون / أجهزة / سويتشات / راوترات / القسم الآمن)
# ------------------------------------------------------------------
def require_login():
    if not is_logged_in():
        return jsonify({'error': 'الرجاء تسجيل الدخول أولاً'}), 401
    return None


def require_admin():
    if not is_logged_in():
        return jsonify({'error': 'الرجاء تسجيل الدخول أولاً'}), 401
    if current_role() != 'admin':
        return jsonify({'error': 'هذا القسم مخصص لمسؤول النظام فقط'}), 403
    return None


@app.route('/api/<table>', methods=['GET'])
def get_table(table):
    if table not in TABLES:
        return jsonify({'error': 'جدول غير معروف'}), 404
    err = require_admin() if table in ADMIN_ONLY_TABLES else require_login()
    if err:
        return err
    return jsonify(load_table(table))


@app.route('/api/<table>', methods=['POST'])
def upsert_record(table):
    if table not in TABLES:
        return jsonify({'error': 'جدول غير معروف'}), 404
    err = require_admin() if table in ADMIN_ONLY_TABLES else require_login()
    if err:
        return err

    record = request.get_json(silent=True) or {}
    records = load_table(table)

    if table == 'employees' and record.get('device'):
        owner = next((r for r in records if r.get('device') == record['device'] and r.get('id') != record.get('id')), None)
        if owner:
            return jsonify({'error': f"الجهاز {record['device']} بعهدة {owner.get('name')} بالفعل"}), 409

    if record.get('id'):
        idx = next((i for i, r in enumerate(records) if r.get('id') == record['id']), None)
        if idx is not None:
            records[idx] = record
        else:
            records.append(record)
    else:
        record['id'] = uuid.uuid4().hex
        records.append(record)

    save_table(table, records)
    return jsonify(record)


@app.route('/api/<table>/<record_id>', methods=['DELETE'])
def delete_record(table, record_id):
    if table not in TABLES:
        return jsonify({'error': 'جدول غير معروف'}), 404
    err = require_admin() if table in ADMIN_ONLY_TABLES else require_login()
    if err:
        return err

    records = load_table(table)
    records = [r for r in records if r.get('id') != record_id]
    save_table(table, records)
    return jsonify({'success': True})


# ------------------------------------------------------------------
# تقديم واجهة الموقع
# ------------------------------------------------------------------
@app.route('/')
def index():
    return send_from_directory(STATIC_DIR, 'index.html')


if __name__ == '__main__':
    print('=' * 60)
    print('نظام جرد الأصول التقنية يعمل الآن.')
    print('افتح على نفس الجهاز: http://localhost:5000')
    print('شارك مع باقي الموظفين على نفس الشبكة عبر عنوان IP هذا الجهاز، مثال:')
    print('http://192.168.1.10:5000')
    print('=' * 60)
    app.run(host='0.0.0.0', port=5000, debug=False)
