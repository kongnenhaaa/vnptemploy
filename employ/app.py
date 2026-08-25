"""
VNPT Employee Web App  ·  app.py  (v2 - OneBSS auth flow)
Auth flow khớp với app.js gốc:
  Step 1: POST /quantri/user/xacthuc_tapdoan  → secretCode
  Step 2: POST /quantri/oauth/token (secretCode + OTP) → access_token
"""
import os, re, json, time, base64, hashlib, threading, subprocess, shutil, secrets
import sys
from functools import wraps
from flask import (Flask, render_template, request, session,
                   redirect, url_for, jsonify, flash, Response)
import requests
import urllib3
from urllib.parse import urlparse, parse_qsl
from flask_session import Session
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Windows can start the source build with a legacy cp1252 console.  Logging a
# Vietnamese status message must never abort session restoration or delete a
# still-valid persistent login.
for _stream in (sys.stdout, sys.stderr):
    try:
        if _stream is not None and hasattr(_stream, 'reconfigure'):
            _stream.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError):
        pass

# ── eKYC / ĐKTTTB service ──────────────────────────────────────────────────
try:
    import ekyc_service as _ekyc
    EKYC_AVAILABLE = True
except ImportError as _ekyc_err:
    EKYC_AVAILABLE = False
    print(f"[WARN] ekyc_service not available: {_ekyc_err}")

app = Flask(__name__)
app.secret_key = 'vnpt-employ-static-key-2026-x9k2m'
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

# ─── Server-side session (tùy filesystem, không bị giới hạn 4KB cookie) ───
app.config['SESSION_TYPE'] = 'filesystem'
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    app_base_dir = sys._MEIPASS
    data_base_dir = sys._MEIPASS
    app.template_folder = os.path.join(app_base_dir, 'templates')
    app.static_folder = os.path.join(app_base_dir, 'static')
else:
    # Running in a normal Python environment
    app_base_dir = os.path.dirname(os.path.abspath(__file__))
    data_base_dir = os.path.join(app_base_dir, '..')

# Dữ liệu đăng nhập phải nằm ở thư mục người dùng để bản portable có thể chạy
# từ Downloads, USB hoặc thư mục chỉ đọc mà không cần cài đặt/quyền admin.
_credential_root = os.path.join(
    os.environ.get('VNPT_EMPLOY_DATA_ROOT') or
    os.path.join(os.environ.get('LOCALAPPDATA') or os.path.dirname(sys.executable), 'VNPTEmploy')
)
if getattr(sys, 'frozen', False) or os.environ.get('VNPT_EMPLOY_DATA_ROOT'):
    session_dir = os.path.join(_credential_root, 'flask_sessions')
else:
    session_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'flask_sessions')

os.makedirs(session_dir, exist_ok=True)
app.config['SESSION_FILE_DIR'] = session_dir
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True
Session(app)

# ─── Saved Employee accounts ────────────────────────────────────────────────
# Passwords are encrypted with Windows DPAPI and never returned to the browser.
# LOCALAPPDATA keeps the account list stable even when the PyInstaller build is
# replaced or the desktop app starts its webview on a different random port.
SAVED_ACCOUNTS_FILE = os.path.join(_credential_root, 'employee_accounts.dat')
_saved_accounts_lock = threading.RLock()

# ─── SIM Kit batch Excel files ─────────────────────────────────────────────
# Keep these in Documents so the files are stable, user-visible, and can be
# opened by Excel.  Microsoft Store Python virtualizes writes to LOCALAPPDATA;
# Excel is a separate process and cannot see those virtualized files.
_legacy_sim_batch_dir = os.path.join(_credential_root, 'SIM_Kit_Batch')
_documents_root = os.path.join(
    os.environ.get('VNPT_EMPLOY_DOCUMENTS_ROOT') or
    os.path.join(os.environ.get('USERPROFILE') or os.path.expanduser('~'), 'Documents')
)
SIM_BATCH_DIR = os.path.join(_documents_root, 'VNPTEmploy', 'SIM_Kit_Batch')
SIM_BATCH_INPUT_FILE = os.path.join(SIM_BATCH_DIR, 'SIM_Kit_Input.xlsx')
SIM_BATCH_OUTPUT_FILE = os.path.join(SIM_BATCH_DIR, 'SIM_Kit_Output.xlsx')
_sim_batch_file_lock = threading.Lock()

# Batch thao tác IC/OC dùng workbook riêng để không lẫn Serial SIM của luồng
# khởi tạo SIM Kit. File được giữ trong Documents giống batch SIM Kit.
ICOC_BATCH_DIR = os.path.join(_documents_root, 'VNPTEmploy', 'IC_OC_Batch')
ICOC_BATCH_INPUT_FILE = os.path.join(ICOC_BATCH_DIR, 'IC_OC_Input.xlsx')
ICOC_BATCH_OUTPUT_FILE = os.path.join(ICOC_BATCH_DIR, 'IC_OC_Output.xlsx')
_icoc_batch_file_lock = threading.Lock()

def _ensure_sim_batch_files():
    """Create the stable input template and an empty legacy output template."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter

    os.makedirs(SIM_BATCH_DIR, exist_ok=True)
    # Preserve templates/results created by older versions in LOCALAPPDATA.
    for file_name in ('SIM_Kit_Input.xlsx', 'SIM_Kit_Output.xlsx'):
        legacy_file = os.path.join(_legacy_sim_batch_dir, file_name)
        current_file = os.path.join(SIM_BATCH_DIR, file_name)
        if not os.path.exists(current_file) and os.path.isfile(legacy_file):
            try:
                shutil.copy2(legacy_file, current_file)
            except OSError:
                # A missing/inaccessible legacy file must not prevent creation
                # of a fresh workbook at the externally visible location.
                pass
    if not os.path.exists(SIM_BATCH_INPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'SIM_Input'
        sheet.append(['SĐT', 'Serial SIM'])
        sheet.freeze_panes = 'A2'
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 20
        guide = workbook.create_sheet('Huong_dan')
        guide.append(['Hướng dẫn'])
        guide.append(['Mỗi dòng trong sheet SIM_Input gồm SĐT thuê bao và Serial SIM trắng.'])
        guide.append(['SĐT nhận dạng đầu 84, đầu 0 hoặc 9 chữ số không có tiền tố.'])
        guide.append(['Ví dụ tương đương: 84849531207 / 0849531207 / 849531207'])
        guide.append(['Ví dụ một dòng: 0849531207 | 1184229391'])
        workbook.save(SIM_BATCH_INPUT_FILE)
    else:
        try:
            workbook = load_workbook(SIM_BATCH_INPUT_FILE)
            guide = workbook['Huong_dan'] if 'Huong_dan' in workbook.sheetnames else workbook.create_sheet('Huong_dan')
            instructions = [
                'Hướng dẫn',
                'Mỗi dòng trong sheet SIM_Input gồm SĐT thuê bao và Serial SIM trắng.',
                'SĐT nhận dạng đầu 84, đầu 0 hoặc 9 chữ số không có tiền tố.',
                'Ví dụ tương đương: 84849531207 / 0849531207 / 849531207',
                'Ví dụ một dòng: 0849531207 | 1184229391',
            ]
            changed = False
            for row_index, value in enumerate(instructions, start=1):
                if guide.cell(row=row_index, column=1).value != value:
                    guide.cell(row=row_index, column=1, value=value)
                    changed = True
            if changed:
                workbook.save(SIM_BATCH_INPUT_FILE)
            workbook.close()
        except PermissionError:
            pass
    if not os.path.exists(SIM_BATCH_OUTPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'SIM_Output'
        sheet.append(['SĐT', 'Serial SIM', 'Kết quả', 'Thời gian', 'User chạy'])
        sheet.freeze_panes = 'A2'
        for column, width in {'A':20, 'B':20, 'C':80, 'D':22, 'E':24}.items():
            sheet.column_dimensions[column].width = width
        workbook.save(SIM_BATCH_OUTPUT_FILE)
    else:
        # Output cũ từng có cột địa chỉ riêng; kết quả đã chứa địa chỉ nên bỏ
        # cột trùng lặp mà vẫn giữ nguyên toàn bộ các dòng lịch sử.
        try:
            workbook = load_workbook(SIM_BATCH_OUTPUT_FILE)
            sheet = workbook['SIM_Output'] if 'SIM_Output' in workbook.sheetnames else workbook.active
            address_column = next((
                index for index, cell in enumerate(sheet[1], start=1)
                if str(cell.value or '').strip().casefold() == 'địa chỉ đầy đủ'.casefold()
            ), None)
            if address_column:
                sheet.delete_cols(address_column, 1)
                workbook.save(SIM_BATCH_OUTPUT_FILE)
            workbook.close()
        except PermissionError:
            pass


def _ensure_icoc_batch_files():
    """Create the stable input/output workbooks for batch IC/OC changes."""
    from openpyxl import Workbook, load_workbook

    os.makedirs(ICOC_BATCH_DIR, exist_ok=True)
    if not os.path.exists(ICOC_BATCH_INPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'IC_OC_Input'
        sheet.append(['SĐT'])
        sheet.freeze_panes = 'A2'
        sheet.column_dimensions['A'].width = 20
        guide = workbook.create_sheet('Huong_dan')
        guide.append(['Hướng dẫn'])
        guide.append(['Mỗi dòng trong sheet IC_OC_Input gồm một SĐT thuê bao.'])
        guide.append(['SĐT nhận dạng đầu 84, đầu 0 hoặc 9 chữ số không có tiền tố.'])
        guide.append(['Ví dụ tương đương: 84846216326 / 0846216326 / 846216326'])
        workbook.save(ICOC_BATCH_INPUT_FILE)
    else:
        try:
            workbook = load_workbook(ICOC_BATCH_INPUT_FILE)
            guide = workbook['Huong_dan'] if 'Huong_dan' in workbook.sheetnames else workbook.create_sheet('Huong_dan')
            instructions = [
                'Hướng dẫn',
                'Mỗi dòng trong sheet IC_OC_Input gồm một SĐT thuê bao.',
                'SĐT nhận dạng đầu 84, đầu 0 hoặc 9 chữ số không có tiền tố.',
                'Ví dụ tương đương: 84846216326 / 0846216326 / 846216326',
            ]
            changed = False
            for row_index, value in enumerate(instructions, start=1):
                if guide.cell(row=row_index, column=1).value != value:
                    guide.cell(row=row_index, column=1, value=value)
                    changed = True
            if changed:
                workbook.save(ICOC_BATCH_INPUT_FILE)
            workbook.close()
        except PermissionError:
            pass
    if not os.path.exists(ICOC_BATCH_OUTPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'IC_OC_Output'
        sheet.append(['SĐT', 'Kết quả', 'Thời gian', 'User chạy'])
        sheet.freeze_panes = 'A2'
        for column, width in {'A': 20, 'B': 80, 'C': 22, 'D': 24}.items():
            sheet.column_dimensions[column].width = width
        workbook.save(ICOC_BATCH_OUTPUT_FILE)

def _open_local_file(path):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f'Khong tim thay file: {path}')
    if os.name == 'nt':
        os.startfile(path, 'open')
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def _new_batch_output_path(directory, prefix):
    """Return a unique workbook path for one completed batch run.

    A run must never append to a previous output workbook: operators often keep
    the old files open in Excel and need an immutable audit file per assignment.
    The random suffix also prevents two workers finishing in the same millisecond
    from choosing the same filename.
    """
    os.makedirs(directory, exist_ok=True)
    stamp = time.strftime('%Y%m%d_%H%M%S')
    return os.path.join(directory, f'{prefix}_{stamp}_{secrets.token_hex(3)}.xlsx')


def _owned_batch_file_path(requested, directory, fallback):
    """Resolve a browser-provided output path without allowing path traversal."""
    candidate = os.path.abspath(str(requested or '').strip()) if requested else os.path.abspath(fallback)
    root = os.path.abspath(directory)
    try:
        if os.path.commonpath([candidate, root]) != root or not candidate.lower().endswith('.xlsx'):
            return os.path.abspath(fallback)
    except ValueError:
        return os.path.abspath(fallback)
    return candidate


def _authenticated_batch_user_map():
    user_name = str(session.get('username') or '--')
    users = {user_name.casefold(): user_name}
    for account in (session.get('multi_accounts') or {}).values():
        if isinstance(account, dict) and account.get('username'):
            value = str(account['username'])
            users[value.casefold()] = value
    return user_name, users

def _dpapi_crypt(data, protect=True):
    if os.name != 'nt':
        raise RuntimeError('Saved accounts require Windows DPAPI')
    import ctypes
    from ctypes import wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [('cbData', wintypes.DWORD),
                    ('pbData', ctypes.POINTER(ctypes.c_byte))]

    source_buffer = ctypes.create_string_buffer(data)
    source = DATA_BLOB(len(data), ctypes.cast(source_buffer, ctypes.POINTER(ctypes.c_byte)))
    output = DATA_BLOB()
    if protect:
        ok = ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(source), 'VNPT Employ', None, None, None, 0,
            ctypes.byref(output))
    else:
        ok = ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(source), None, None, None, None, 0,
            ctypes.byref(output))
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output.pbData, output.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(output.pbData)

def _read_saved_accounts():
    with _saved_accounts_lock:
        try:
            with open(SAVED_ACCOUNTS_FILE, 'rb') as f:
                raw = base64.b64decode(f.read())
            accounts = json.loads(_dpapi_crypt(raw, protect=False).decode('utf-8'))
            if not isinstance(accounts, list):
                return []
            return [a for a in accounts if isinstance(a, dict) and
                    a.get('id') and a.get('username') and a.get('password')]
        except (FileNotFoundError, ValueError, OSError, RuntimeError, json.JSONDecodeError):
            return []

def _write_saved_accounts(accounts):
    with _saved_accounts_lock:
        os.makedirs(_credential_root, exist_ok=True)
        payload = json.dumps(accounts[:30], ensure_ascii=False).encode('utf-8')
        encrypted = base64.b64encode(_dpapi_crypt(payload, protect=True))
        temp_file = SAVED_ACCOUNTS_FILE + '.tmp'
        with open(temp_file, 'wb') as f:
            f.write(encrypted)
        os.replace(temp_file, SAVED_ACCOUNTS_FILE)

def _ensure_application_storage():
    """Create every user-writable file/folder required by the portable app."""
    os.makedirs(_credential_root, exist_ok=True)
    os.makedirs(session_dir, exist_ok=True)
    if not os.path.exists(SAVED_ACCOUNTS_FILE):
        _write_saved_accounts([])
    _ensure_sim_batch_files()
    _ensure_icoc_batch_files()

def _account_id(username):
    return hashlib.sha256(username.strip().casefold().encode('utf-8')).hexdigest()[:24]

def save_employee_account(username, password):
    username = str(username or '').strip()
    password = str(password or '')
    if not username or not password:
        return False
    account_id = _account_id(username)
    with _saved_accounts_lock:
        accounts = [a for a in _read_saved_accounts() if a.get('id') != account_id]
        accounts.insert(0, {
            'id': account_id,
            'username': username,
            'password': password,
            'saved_at': int(time.time())
        })
        _write_saved_accounts(accounts)
    return True

def saved_account_summaries():
    return [{'id': a['id'], 'username': a['username']}
            for a in _read_saved_accounts()]

def get_saved_employee_account(account_id):
    return next((a for a in _read_saved_accounts()
                 if a.get('id') == account_id), None)

# ─── CONFIG (Dynamic) ───────────────────────
BASE_URL      = 'https://api-onebss.vnpt.vn'
APP_CFG = {
    'CLIENT_ID': 'clientapp',
    'CLIENT_SECRET': 'password',
    'MENU_ID': 810241,
    'SELECTED_MENU': '810241',
    'APP_VERSION': '1.5.41.007'
}

# IDG Token-id / Token-key (từ upload_mobile.js)
IDG_TOKEN_ID  = '04c0a953-7fb8-5461-e063-62199f0aeda6'
IDG_TOKEN_KEY = 'MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAKjy7FK9SegSCW0cuUIbEDUsbRZOCoxijNPLMfvgX+8/XA7HebHXMN4/PO5c5mwK31Yk31RKuMXYLLp6X6oZPDKcAwEAAQ=='

# ─────────────────────────────────────────────────────────────
#  Parser
# ─────────────────────────────────────────────────────────────
DATA_FILE = os.path.join(data_base_dir, 'request_body_full.txt')

def parse_doc(path):
    with open(path, encoding='utf-8', errors='ignore') as f:
        raw = f.read().lstrip('\ufeff')
    sections = []
    parts = re.split(r'={60,}', raw)
    i = 3
    while i < len(parts) - 1:
        heading = parts[i].strip()
        body    = parts[i+1]
        m = re.search(r'^(\d+)\.\s+(.+)$', heading, re.M)
        if not m:
            i += 2; continue
        section = {'id': m.group(1), 'name': m.group(2).strip(),
                   'full': f"{m.group(1)}. {m.group(2).strip()}", 'apis': []}
        subs = re.split(r'-{60,}', body)
        j = 1
        while j < len(subs) - 1:
            tm = re.search(r'(\d+\.\d+[\.\d]*)\s{2,}(.+)', subs[j].strip())
            if tm:
                api = _parse_api(tm.group(1)+'  '+tm.group(2).strip(), subs[j+1])
                if api: section['apis'].append(api)
            j += 2
        if section['apis']: sections.append(section)
        i += 2
    return sections

def _parse_api(title, text):
    api = {'title': title, 'method': 'GET', 'endpoints': [],
           'content_type': 'application/json', 'body': '', 'body_type': 'json',
           'params': {}, 'headers': {}, 'note': '', 'source': '',
           'response_fields': [], 'primary_endpoint': ''}
    lines = text.split('\n')
    i = 0; in_body = False; in_resp = False; body_lines = []; resp_lines = []
    while i < len(lines):
        s = lines[i].strip()
        if re.match(r'^Method\s*:\s*', s, re.I):
            api['method'] = re.sub(r'^Method\s*:\s*', '', s, flags=re.I).strip()
            in_body = in_resp = False
        elif re.match(r'^Endpoints?\s*:', s, re.I):
            val = re.sub(r'^Endpoints?\s*:\s*', '', s, flags=re.I).strip()
            if val:
                for ep in re.split(r'\s+hoac\s+|\s{3,}', val):
                    if ep.strip(): api['endpoints'].append(ep.strip())
            k = i+1
            while k < len(lines) and lines[k].strip().startswith('/'):
                api['endpoints'].append(lines[k].strip()); k+=1
            i = k-1; in_body = in_resp = False
        elif re.match(r'^Content-Type\s*:', s, re.I):
            ct = re.sub(r'^Content-Type\s*:\s*', '', s, flags=re.I).strip()
            api['content_type'] = ct
            if 'form' in ct: api['body_type'] = 'form'
        elif re.match(r'^Body(\s*\(JSON\))?\s*:|^Params\s*:', s, re.I):
            in_body = True; in_resp = False; body_lines = []
            api['body_type'] = 'query' if re.match(r'^Params', s, re.I) else 'json'
        elif re.match(r'^Note\s*:', s, re.I):
            api['note'] = re.sub(r'^Note\s*:\s*', '', s, flags=re.I).strip()
            in_body = in_resp = False
        elif re.match(r'^Source\s*:', s, re.I):
            api['source'] = re.sub(r'^Source\s*:\s*', '', s, flags=re.I).strip()
            in_body = in_resp = False
        elif re.match(r'^Response(\s+fields?)?\s*:|^Context\s+data', s, re.I):
            in_resp = True; in_body = False
            val = re.sub(r'^[^:]+:\s*', '', s).strip()
            if val: resp_lines.append(val)
        elif in_body:
            if s == '' and body_lines and any(l.strip().startswith('}') for l in body_lines):
                in_body = False
            elif s: body_lines.append(lines[i])
        elif in_resp:
            if s == '' or re.match(r'^Source\s*:', s, re.I):
                in_resp = False
                if re.match(r'^Source\s*:', s, re.I):
                    api['source'] = re.sub(r'^Source\s*:\s*', '', s, flags=re.I).strip()
            else: resp_lines.append(s)
        i += 1
    raw_body = '\n'.join(body_lines).strip()
    api['body'] = raw_body
    if raw_body:
        try: api['params'] = json.loads(raw_body)
        except:
            for m2 in re.finditer(r'"([\w_]+)"\s*:\s*"([^"]*)"', raw_body):
                api['params'][m2.group(1)] = m2.group(2)
            for m2 in re.finditer(r'^([^\s=]+)\s*=\s*(.*)$', raw_body, re.M):
                api['params'][m2.group(1)] = m2.group(2).strip()
    rf = ' '.join(resp_lines)
    api['response_fields'] = [f.strip() for f in re.split(r'[,\n]', rf) if f.strip()]
    api['primary_endpoint'] = api['endpoints'][0] if api['endpoints'] else ''
    return api if (api['endpoints'] or api['body'] or api['note']) else None

SECTIONS = parse_doc(DATA_FILE)
SECTION_MAP = {s['id']: s for s in SECTIONS}

# ─────────────────────────────────────────────────────────────
# Persistent session helper (tự động khôi phục trong 4 tiếng)
# ─────────────────────────────────────────────────────────────
PERSISTENT_SESSION_FILE = os.path.join(_credential_root, 'session_state.dat')
NO_DEFAULT_ACCOUNTS_MARKER = os.path.join(
    _credential_root, 'no_default_accounts_v1.marker')
_persistent_session_lock = threading.RLock()
_multi_account_lock = threading.RLock()
# OneBSS invalidates the previous secretCode when a second OTP is requested.
# Keep one in-flight request per username and reject rapid duplicates so a
# double-click or two WebView events cannot send two competing OTPs.
_otp_request_locks = {}
_otp_request_locks_guard = threading.RLock()
OTP_REQUEST_COOLDOWN_SECONDS = 60
SESSION_MAX_INACTIVE_SECONDS = 4 * 3600  # 4 tiếng (14,400 giây)


def _otp_lock_for(username):
    key = str(username or '').strip().casefold()
    with _otp_request_locks_guard:
        return _otp_request_locks.setdefault(key, threading.Lock())

def _save_persistent_session():
    """Lưu trạng thái đăng nhập vào đĩa kèm thời điểm thao tác/tắt tool cuối cùng."""
    if 'access_token' not in session:
        return
    try:
        os.makedirs(_credential_root, exist_ok=True)
        data = {
            'access_token': session.get('access_token'),
            'refresh_token': session.get('refresh_token'),
            'expires_in': session.get('expires_in'),
            'token_time': session.get('token_time'),
            'device_id': session.get('device_id'),
            'app_secret': session.get('app_secret'),
            'username': session.get('username'),
            'menus': session.get('menus', []),
            'active_menu_id': session.get('active_menu_id'),
            'multi_accounts': session.get('multi_accounts', {}),
            'last_active_time': int(time.time())
        }
        payload = json.dumps(data, ensure_ascii=False).encode('utf-8')
        if os.name == 'nt':
            encrypted = base64.b64encode(_dpapi_crypt(payload, protect=True))
        else:
            encrypted = base64.b64encode(payload)
        with _persistent_session_lock:
            temp_file = f'{PERSISTENT_SESSION_FILE}.{os.getpid()}.tmp'
            with open(temp_file, 'wb') as f:
                f.write(encrypted)
            os.replace(temp_file, PERSISTENT_SESSION_FILE)
    except Exception as e:
        print(f"[WARN] Không thể lưu persistent session: {e}")

def _clear_persistent_session():
    try:
        if os.path.exists(PERSISTENT_SESSION_FILE):
            os.remove(PERSISTENT_SESSION_FILE)
    except Exception:
        pass

def _remove_legacy_default_accounts_once():
    """Clear credentials/session left by builds that shipped saved defaults.

    The marker makes this a one-time migration. Accounts added after the user
    signs in and verifies OTP continue to persist normally on later launches.
    """
    if os.path.exists(NO_DEFAULT_ACCOUNTS_MARKER):
        return False
    os.makedirs(_credential_root, exist_ok=True)
    _write_saved_accounts([])
    _clear_persistent_session()
    try:
        cache = getattr(app.session_interface, 'cache', None)
        if cache is not None:
            cache.clear()
    except Exception:
        pass
    temp_marker = NO_DEFAULT_ACCOUNTS_MARKER + '.tmp'
    with open(temp_marker, 'w', encoding='utf-8') as marker:
        marker.write('Legacy/default accounts removed.\n')
    os.replace(temp_marker, NO_DEFAULT_ACCOUNTS_MARKER)
    return True

def _try_restore_persistent_session():
    """Tự động khôi phục phiên đăng nhập nếu tắt tool chưa quá 4 tiếng."""
    if 'access_token' in session:
        return True

    if not os.path.exists(PERSISTENT_SESSION_FILE):
        return False

    try:
        with _persistent_session_lock:
            with open(PERSISTENT_SESSION_FILE, 'rb') as f:
                raw = base64.b64decode(f.read())
        if os.name == 'nt':
            decrypted = _dpapi_crypt(raw, protect=False).decode('utf-8')
        else:
            decrypted = raw.decode('utf-8')
        data = json.loads(decrypted)

        last_active = data.get('last_active_time', 0)
        elapsed = time.time() - last_active

        # Nếu thời gian kể từ lúc tắt/dùng tool cuối cùng đã >= 4 tiếng -> Yêu cầu đăng nhập lại
        if elapsed >= SESSION_MAX_INACTIVE_SECONDS:
            print(f"[SESSION] Phiên đăng nhập đã quá 4 tiếng ({elapsed/3600:.2f}h). Yêu cầu đăng nhập lại.")
            _clear_persistent_session()
            return False

        # Còn trong vòng 4 tiếng -> Khôi phục session và tự động đăng nhập
        session['access_token']   = data.get('access_token')
        session['refresh_token']  = data.get('refresh_token')
        session['expires_in']     = data.get('expires_in')
        session['token_time']     = data.get('token_time')
        session['device_id']      = data.get('device_id')
        session['app_secret']     = data.get('app_secret')
        session['username']       = data.get('username')
        session['menus']          = data.get('menus', [])
        session['active_menu_id'] = data.get('active_menu_id')
        restored_accounts = data.get('multi_accounts') or {}
        session['multi_accounts'] = {
            str(account_id): account for account_id, account in restored_accounts.items()
            if isinstance(account, dict) and account.get('username') and account.get('access_token')
        } if isinstance(restored_accounts, dict) else {}

        # App versions can become invalid while a saved session is still active.
        # Refresh only the version metadata while preserving the original device id.
        session['app_secret'] = build_app_secret()

        print(f"[SESSION] Tự động khôi phục đăng nhập (lần cuối dùng tool cách đây {elapsed/60:.1f} phút).")
        _save_persistent_session()
        return True
    except Exception as e:
        print(f"[WARN] Lỗi khi khôi phục persistent session: {e}")
        _clear_persistent_session()
        return False

# ─────────────────────────────────────────────────────────────
#  Auth helpers
# ─────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'access_token' not in session:
            if not _try_restore_persistent_session():
                return redirect(url_for('login'))
        _save_persistent_session()
        return f(*args, **kwargs)
    return decorated

def _decode_app_secret(value):
    """Decode an existing base64 JSON app-secret without trusting its version."""
    if not isinstance(value, str) or not value:
        return {}
    try:
        padded = value + ('=' * (-len(value) % 4))
        decoded = json.loads(base64.b64decode(padded).decode('utf-8'))
        return decoded if isinstance(decoded, dict) else {}
    except (ValueError, TypeError, UnicodeDecodeError, json.JSONDecodeError):
        return {}


def _build_app_secret_value(existing_value='', device_id=''):
    """Build app-secret without mutating Flask's current login session."""
    existing = _decode_app_secret(existing_value)
    device_id = device_id or existing.get('device_id') or '0f8c2d3fb0c51653'
    obj = dict(existing)
    defaults = {
        "device_id": device_id,
        "device_ip": "Unknown",
        "device_name": "Web-Browser",
        "mac_address": "Unknown",
        "mobile_id": "web-generated-id",
        "app_id": "1",
        "os_version": "Chrome/Web",
    }
    for key, value in defaults.items():
        obj.setdefault(key, value)
    obj['device_id'] = device_id
    obj['app_version'] = APP_CFG['APP_VERSION']
    return base64.b64encode(
        json.dumps(obj, separators=(',', ':')).encode('utf-8')
    ).decode()


def build_app_secret():
    """Tạo app-secret giống app mobile (JSON base64) với phiên bản hiện tại."""
    existing = _decode_app_secret(session.get('app_secret'))
    device_id = (session.get('device_id') or existing.get('device_id') or
                 '0f8c2d3fb0c51653')
    session['device_id'] = device_id

    # Preserve device metadata from the authenticated session, but always replace
    # app_version so an old cached session cannot keep sending an invalid version.
    obj = dict(existing)
    defaults = {
        "device_id":   device_id,
        "device_ip":   "Unknown",
        "device_name": "Web-Browser",
        "mac_address": "Unknown",
        "mobile_id":   "web-generated-id",
        "app_id":      "1",
        "os_version":  "Chrome/Web"
    }
    for key, value in defaults.items():
        obj.setdefault(key, value)
    obj['device_id'] = device_id
    obj['app_version'] = APP_CFG['APP_VERSION']
    return base64.b64encode(
        json.dumps(obj, separators=(',', ':')).encode('utf-8')
    ).decode()


def current_app_secret():
    """Return and cache an app-secret synchronized with APP_CFG."""
    value = build_app_secret()
    session['app_secret'] = value
    return value

def _primary_account_context():
    username = str(session.get('username') or '').strip()
    return {
        'id': _account_id(username) if username else '',
        'username': username,
        'access_token': session.get('access_token', ''),
        'refresh_token': session.get('refresh_token', ''),
        'expires_in': session.get('expires_in', 3600),
        'token_time': session.get('token_time', 0),
        'device_id': session.get('device_id', ''),
        'app_secret': session.get('app_secret', ''),
        'primary': True,
    }


def _get_account_context(account_id=None):
    """Resolve an authenticated account without returning secrets to the UI."""
    primary = _primary_account_context()
    requested = str(account_id or '').strip()
    if not requested or requested == primary.get('id'):
        return primary
    accounts = session.get('multi_accounts') or {}
    context = accounts.get(requested) if isinstance(accounts, dict) else None
    if not isinstance(context, dict) or not context.get('access_token'):
        raise ValueError('Tài khoản chạy không tồn tại hoặc chưa xác thực OTP')
    resolved = dict(context)
    resolved['id'] = requested
    resolved['primary'] = False
    return resolved


def _account_seconds_remaining(context):
    try:
        elapsed = time.time() - float(context.get('token_time') or 0)
        return max(0, int(float(context.get('expires_in') or 3600) - elapsed))
    except (TypeError, ValueError):
        return 0


def get_headers(menu_id=None, account_id=None):
    """Headers chuẩn cho tất cả API calls - dùng active_menu_id từ session"""
    active_mid = menu_id or session.get('active_menu_id', APP_CFG['SELECTED_MENU'])
    context = _get_account_context(account_id)
    account_app_secret = _build_app_secret_value(
        context.get('app_secret', ''), context.get('device_id', ''))
    return {
        'Content-Type':   'application/json',
        'Accept':         'application/json',
        'authorization':  f"Bearer {context.get('access_token', '')}",
        'app-secret':     account_app_secret,
        'selectedmenuid': active_mid,
        'SelectedMenuId': active_mid,
    }

def api_call(method, path, **kwargs):
    url = path if path.startswith('http') else \
        BASE_URL.rstrip('/') + '/' + path.lstrip('/')
    kw = dict(headers=get_headers(), verify=False, timeout=15)
    kw.update(kwargs)
    try:
        return requests.request(method.upper(), url, **kw)
    except Exception:
        return None


# OneBSS validates SelectedMenuId independently for each business module. A
# single menu selected in the sidebar cannot be reused for calls made by a tab
# that contains more than one module (Cashless/DCRS is the common example).
ENDPOINT_MENU_ROUTES = (
    ('/ccbs/chonSo/', '699161'),
    ('/app-banhang/donhang_simkit/', '699161'),
    ('/app-com/danhmuc/get_danhmuc', '699161'),
    ('/app-thuno/VnptPay/', '699161'),
    ('/app-banhang/giohang/', '11074'),
    ('/app-banhang/banhang_dcrs/', '11177'),
    ('/tichhop/smcs/banHang/', '11177'),
    ('/app-banhang/cashless/', '11176'),
    ('/app-ccdv/vietqr/', '11176'),
    ('/app-banhang/thanhtoan_hopdong/', '11176'),
    ('/app-banhang/cmgs/', '11173'),
    ('/app-banhang/hopdong/', '11199'),
    ('/app-banhang/hoadondientu/', '11076'),
    ('/app-banhang/phieuyeucau', '11143'),
    ('/app-cskh/', '11141'),
    ('/app-banhang/b2a/', '11141'),
    ('/app-banhang/b2c/', '11167'),
    ('/app-banhang/baocao_banhang/', '700214'),
    # Full SIM-change flow captured from the Employee mobile application.
    ('/ccbs/oneBss/', '11175'),
    ('/app-banhang/Ekyc/insert_log_ekyc_doisim_v2', '11175'),
    ('/app-banhang/thuebaodidong/app_tb_doisim_v2', '11175'),
    # Mobile captures use menu 11213 for IC/OC status, permission and change.
    ('/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen', '11213'),
    ('/app-banhang/thuebaodidong/khoamo_ic_oc', '11213'),
    ('/app-banhang/thietbi_thuebao/', '11175'),
    ('/app-banhang/ccbs/tracuu_anh_thuebao', '810241'),
    ('/app-banhang/ccbs/', '11213'),
    ('/app-banhang/danhba/', '11213'),
    ('/app-banhang/thongtincuoc/', '11213'),
    ('/app-banhang/thongtindanhba/', '11213'),
    ('/app-banhang/thuebaodidong/', '11213'),
)

# Các endpoint chỉ đọc theo source mobile. Một số cụm OneBSS hiện vẫn dùng
# GET, trong khi một số gateway khác trả 405 cho GET và chỉ nhận POST. Proxy
# được phép thử phương thức còn lại riêng cho danh sách này; tuyệt đối không
# áp dụng cơ chế đó cho API tạo/cập nhật/hủy để tránh phát sinh giao dịch kép.
READ_ONLY_ENDPOINTS = frozenset({
    '/app-banhang/cashless/ds_donhang_cashless',
    '/app-banhang/cashless/chitiet_donhang_cashless',
    '/app-banhang/banhang_dcrs/ds_donhang_dcrs',
    '/app-banhang/cashless/ds_kho',
    '/app-banhang/cmgs/ds_chuyenmang_giuso',
    '/app-banhang/cmgs/goicuocdexuat_tratruocmnp',
    '/app-banhang/cmgs/goicuocdexuat_trasaucmnp',
    '/app-banhang/hopdong/chitiet_hd_thuebao',
    '/app-cskh/dvitcare/lay_tt_nhanvien_yc_khaosat',
    '/app-cskh/quantrithuebao/lay_ds_pbh_smp',
    '/app-banhang/b2a/listdsphieukhaosatb2a',
    '/app-banhang/b2c/ds_tuvan_dichvu_mytv_fiber',
    '/app-banhang/phieuyeucau/listdsphieuyeucau',
    '/app-banhang/baocao_banhang/baocao_sanluong_donhang',
})


def endpoint_menu_id(endpoint_path, fallback=None, body=None):
    # This permission endpoint is shared by multiple mobile modules.  Its
    # SelectedMenuId follows ma_quyen, not the panel that happened to run the
    # previous request: DOISIM=11175, CATMODICHVU (IC/OC)=11213.
    if endpoint_path == '/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen':
        permission_code = str(
            (body or {}).get('ma_quyen', '') if isinstance(body, dict) else ''
        ).strip().upper()
        if permission_code == 'DOISIM':
            return '11175'
        if permission_code == 'CATMODICHVU':
            return '11213'
    for prefix, menu_id in ENDPOINT_MENU_ROUTES:
        if endpoint_path.startswith(prefix):
            return menu_id
    return str(fallback or session.get('active_menu_id') or
               APP_CFG['SELECTED_MENU'])


def normalize_onebss_body(endpoint_path, body, endpoint='', account_username=None):
    """Translate legacy dashboard aliases to DTOs accepted by OneBSS."""
    if not isinstance(body, dict):
        return body
    normalized = dict(body)
    # Preserve query parameters if a legacy GET has to be retried as POST.
    for key, value in parse_qsl(urlparse(endpoint).query, keep_blank_values=True):
        normalized.setdefault(key, value)

    phone_paths = (
        '/app-banhang/ccbs/tracuu_thongtin_thuebao',
        '/app-banhang/thuebaodidong/tracuu_tb_didong',
        '/app-banhang/thuebaodidong/tracuu_taikhoan_tien',
        '/app-banhang/thuebaodidong/cuocnong_didong',
        '/app-banhang/thuebaodidong/dichvu_sudung',
        '/app-banhang/thuebaodidong/lichsu_thuebao',
        '/app-banhang/thuebaodidong/lichsu_goicuoc',
        '/app-banhang/thuebaodidong/tracuu_taikhoan_nhom',
        '/app-banhang/thongtincuoc/thongtin_cuoc_thuebao',
    )
    if endpoint_path in phone_paths:
        so_tb = (normalized.get('p_so_tb') or normalized.get('p_isdn') or
                 normalized.get('isdn') or normalized.get('p_msisdn') or
                 normalized.get('msisdn') or normalized.get('so_tb'))
        if so_tb:
            normalized['p_so_tb'] = re.sub(r'\D', '', str(so_tb))
        for legacy_key in ('p_isdn', 'isdn', 'p_msisdn', 'msisdn', 'so_tb'):
            normalized.pop(legacy_key, None)

    if endpoint_path == '/app-banhang/danhba/chitiet_khachhang':
        customer_id = (normalized.get('p_khachhang_id') or
                       normalized.get('p_doituong_id') or normalized.get('id'))
        if customer_id not in (None, ''):
            normalized['p_khachhang_id'] = customer_id
        normalized.pop('p_doituong_id', None)
        normalized.pop('id', None)

    if endpoint_path == '/app-banhang/danhba/chitiet_thuebao':
        subscriber_id = (normalized.get('p_thuebao_id') or
                         normalized.get('p_id_thuebao') or
                         normalized.get('p_isdn') or normalized.get('isdn') or
                         normalized.get('id'))
        if subscriber_id not in (None, ''):
            normalized['p_thuebao_id'] = subscriber_id
        for legacy_key in ('p_id_thuebao', 'p_isdn', 'isdn', 'id'):
            normalized.pop(legacy_key, None)

    if endpoint_path == '/app-banhang/hopdong/chitiet_hd_thuebao':
        contract_id = (normalized.get('p_hdtb_id') or normalized.get('id') or
                       normalized.get('hopdong_id'))
        if contract_id not in (None, ''):
            normalized['p_hdtb_id'] = contract_id
        normalized.pop('id', None)
        normalized.pop('hopdong_id', None)

    if endpoint_path == '/app-banhang/cashless/chitiet_donhang_cashless':
        sale_trans_id = (normalized.get('sale_trans_id') or
                         normalized.get('saleId') or normalized.get('id'))
        if sale_trans_id not in (None, ''):
            normalized['sale_trans_id'] = sale_trans_id
        normalized.pop('saleId', None)
        normalized.pop('id', None)

    if endpoint_path == '/app-banhang/thongtincuoc/thongtin_cuoc_thuebao':
        normalized.setdefault('p_kycuoc_yc', time.strftime('%m/%Y'))

    if endpoint_path in (
            '/app-banhang/cashless/access_smcs',
            '/app-banhang/cashless/ds_kho'):
        normalized.setdefault('p_account', account_username or session.get('username', ''))

    return normalized

# ─────────────────────────────────────────────────────────────
#  Routes – Auth (2-step: xacthuc_tapdoan → OTP → token)
# ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    if 'access_token' in session or _try_restore_persistent_session():
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/settings/update', methods=['POST'])
def update_settings():
    data = request.get_json(silent=True) or {}
    APP_CFG['CLIENT_ID'] = data.get('client_id', APP_CFG['CLIENT_ID'])
    APP_CFG['CLIENT_SECRET'] = data.get('client_secret', APP_CFG['CLIENT_SECRET'])
    APP_CFG['SELECTED_MENU'] = data.get('selected_menu', APP_CFG['SELECTED_MENU'])
    APP_CFG['MENU_ID'] = int(APP_CFG['SELECTED_MENU']) if str(APP_CFG['SELECTED_MENU']).isdigit() else APP_CFG['MENU_ID']
    requested_version = str(data.get('app_version') or APP_CFG['APP_VERSION']).strip()
    APP_CFG['APP_VERSION'] = requested_version
    
    # Đồng bộ session active_menu_id nếu người dùng lưu cài đặt
    session['active_menu_id'] = APP_CFG['SELECTED_MENU']
    current_app_secret()
    _save_persistent_session()
    
    return jsonify({'ok': True, 'msg': 'Đã cập nhật cài đặt!'})

def _render_login():
    return render_template(
        'login.html',
        saved_accounts=saved_account_summaries(),
        app_version=APP_CFG['APP_VERSION'])

def _begin_employee_login(username, password):
    """Run OneBSS login step 1 and retain the password only until OTP succeeds."""
    import random
    username_key = str(username or '').strip().casefold()
    now = time.time()
    # A repeated POST from the same login page must reuse the still-valid
    # secretCode.  Requesting a new one would invalidate the OTP already sent.
    if (str(session.get('username') or '').strip().casefold() == username_key and
            session.get('secret_code') and
            now - float(session.get('otp_issued_at') or 0) < OTP_REQUEST_COOLDOWN_SECONDS):
        session['password'] = password
        return True, 'OTP đã được gửi trước đó; hãy dùng mã mới nhất trong tin nhắn.'

    otp_lock = _otp_lock_for(username)
    if not otp_lock.acquire(blocking=False):
        return False, 'Đang gửi OTP cho tài khoản này; không gửi lại để tránh mã bị vô hiệu.'
    session.pop('secret_code', None)
    session.pop('otp_issued_at', None)
    device_id = ''.join(random.choices('0123456789abcdef', k=16))
    session['device_id'] = device_id
    try:
        resp = requests.post(
            f'{BASE_URL}/quantri/user/xacthuc_tapdoan',
            json={'username': username, 'password': password,
                  'os_type': '1', 'device_id': device_id},
            headers={'Content-Type': 'application/json', 'Accept': 'application/json'},
            verify=False, timeout=15
        )
        data = resp.json()
    except Exception as exc:
        return False, f'Lỗi kết nối: {exc}'
    finally:
        otp_lock.release()

    if resp.status_code == 200 and data.get('error_code') == 'BSS-00000000':
        session['secret_code'] = (data.get('data') or {}).get('secretCode', '')
        session['username'] = username
        session['password'] = password
        session['otp_issued_at'] = time.time()
        return True, 'Đăng nhập bước 1 thành công. Vui lòng nhập OTP.'
    return False, data.get('message', 'Tài khoản hoặc mật khẩu không đúng.')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        if 'access_token' in session or _try_restore_persistent_session():
            return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        if not username or not password:
            flash('Vui lòng nhập tên đăng nhập và mật khẩu.', 'error')
            return _render_login()
        ok, message = _begin_employee_login(username, password)
        flash(message, 'success' if ok else 'error')
        if ok:
            return redirect(url_for('otp_page'))
    return _render_login()

@app.post('/login/saved')
def login_saved():
    account = get_saved_employee_account(request.form.get('account_id', ''))
    if not account:
        flash('Không tìm thấy tài khoản đã lưu.', 'error')
        return redirect(url_for('login'))
    ok, message = _begin_employee_login(account['username'], account['password'])
    flash(message, 'success' if ok else 'error')
    if ok:
        return redirect(url_for('otp_page'))
    return redirect(url_for('login'))

@app.route('/otp', methods=['GET', 'POST'])
def otp_page():
    if 'username' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        otp = request.form.get('otp', '').strip()
        secret_code = session.get('secret_code', '')

        if not otp:
            flash('Vui lòng nhập mã OTP.', 'error')
            return render_template('otp.html')

        # ── Bước 2: oauth/token với secretCode + OTP (app.js dòng 321-379) ──
        try:
            resp = requests.post(
                f'{BASE_URL}/quantri/oauth/token',
                json={'grant_type':   'password',
                      'client_id':    APP_CFG['CLIENT_ID'],
                      'client_secret': APP_CFG['CLIENT_SECRET'],
                      'secretCode':   secret_code,
                      'otp':          otp},
                headers={'Content-Type': 'application/json', 'Accept': 'application/json'},
                verify=False, timeout=25
            )
            data = resp.json()
        except Exception as e:
            flash(f'Lỗi kết nối khi xác thực OTP: {e}', 'error')
            return render_template('otp.html')

        if data.get('access_token'):
            # Persist only after OTP succeeds; never save a failed credential.
            pending_username = session.get('username', '')
            pending_password = session.get('password', '')
            session['access_token']  = data['access_token']
            session['refresh_token'] = data.get('refresh_token', '')
            session['expires_in']    = data.get('expires_in', 3600)
            session['token_time']    = time.time()
            # Keep any server-provided device metadata, then force the current
            # supported app version into the app-secret.
            if data.get('app_secret'):
                session['app_secret'] = data['app_secret']
            session['app_secret'] = current_app_secret()
            try:
                save_employee_account(pending_username, pending_password)
            except Exception as save_error:
                print(f"[WARN] Could not save Employee account: {save_error}")
            # Xoá thông tin nhạy cảm khỏi session
            session.pop('password', None)
            session.pop('secret_code', None)
            session.pop('otp_issued_at', None)

            # ── Fetch danh sách chức năng (menu) ngay sau login ──
            try:
                _h = {
                    'Content-Type':  'application/json',
                    'Accept':        'application/json',
                    'authorization': f"Bearer {data['access_token']}",
                    'App-secret':    build_app_secret(),
                    'SelectedMenuId': APP_CFG['SELECTED_MENU'],
                    'selectedmenuid': APP_CFG['SELECTED_MENU'],
                }
                _rx = requests.get(
                    f"{BASE_URL}/quantri/user/khoitao_ungdung?p_idmodule=21",
                    headers=_h, verify=False, timeout=15
                )
                if _rx.status_code == 200:
                    _d = _rx.json()
                    _menus = (_d.get('data') or {}).get('ds_chucnang') or []
                    if not _menus and isinstance(_d.get('data'), list):
                        _menus = _d['data']
                    session['menus'] = _menus
                    session['active_menu_id'] = str(APP_CFG['MENU_ID'])
                    print(f"[OK] Loaded {len(_menus)} menu items into session")
                else:
                    session['menus'] = []
                    print(f"[WARN] khoitao_ungdung returned {_rx.status_code}")
            except Exception as _ex:
                session['menus'] = []
                print(f"[WARN] Could not fetch menus: {_ex}")

            flash('Đăng nhập thành công!', 'success')
            _save_persistent_session()
            return redirect(url_for('dashboard'))
        else:
            err = data.get('message', 'Mã OTP không hợp lệ hoặc đã hết hạn.')
            flash(err, 'error')

    return render_template('otp.html')

@app.route('/logout')
def logout():
    session.clear()
    _clear_persistent_session()
    return redirect(url_for('login'))


def _public_account_summary(context):
    return {
        'id': context.get('id', ''),
        'username': context.get('username', ''),
        'primary': bool(context.get('primary')),
        'expires_in': _account_seconds_remaining(context),
        'authenticated': bool(context.get('access_token')),
    }


@app.get('/api/accounts')
@login_required
def api_accounts_list():
    """List authenticated accounts and saved usernames without exposing secrets."""
    primary = _primary_account_context()
    accounts = [_public_account_summary(primary)]
    extras = session.get('multi_accounts') or {}
    if isinstance(extras, dict):
        for account_id, value in extras.items():
            if not isinstance(value, dict) or not value.get('username'):
                continue
            context = dict(value)
            context['id'] = str(account_id)
            context['primary'] = False
            accounts.append(_public_account_summary(context))
    logged_ids = {account['id'] for account in accounts}
    saved = [account for account in saved_account_summaries()
             if account.get('id') not in logged_ids]
    return jsonify({'ok': True, 'accounts': accounts, 'saved_accounts': saved})


@app.post('/api/accounts/begin')
@login_required
def api_accounts_begin():
    """Start an independent OneBSS login without replacing the primary session."""
    payload = request.get_json(silent=True) or {}
    saved_id = str(payload.get('saved_account_id') or '').strip()
    saved = get_saved_employee_account(saved_id) if saved_id else None
    username = str((saved or {}).get('username') or payload.get('username') or '').strip()
    password = str((saved or {}).get('password') or payload.get('password') or '')
    if not username or not password:
        return jsonify({'ok': False, 'error': 'Nhập đầy đủ tài khoản và mật khẩu'}), 400
    account_id = _account_id(username)
    if account_id == _primary_account_context().get('id'):
        return jsonify({'ok': False, 'error': 'Tài khoản này đang là phiên đăng nhập chính'}), 409
    extras = session.get('multi_accounts') or {}
    if len(extras) >= 10 and account_id not in extras:
        return jsonify({'ok': False, 'error': 'Chỉ cho phép tối đa 10 tài khoản trong một phiên'}), 400

    # Reuse the pending challenge during the cooldown.  OneBSS invalidates a
    # secretCode as soon as a newer OTP is requested for the same username.
    pending = session.get('pending_extra_accounts') or {}
    if isinstance(pending, dict):
        recent = next((value for value in pending.values()
                       if isinstance(value, dict)
                       and str(value.get('account_id') or '') == account_id
                       and time.time() - float(value.get('created_at') or 0) < OTP_REQUEST_COOLDOWN_SECONDS), None)
        if recent:
            pending_id = next(key for key, value in pending.items() if value is recent)
            return jsonify({'ok': True, 'pending_id': pending_id, 'username': username,
                            'message': 'OTP đã được gửi trước đó; hãy dùng mã mới nhất trong tin nhắn.'})

    otp_lock = _otp_lock_for(username)
    if not otp_lock.acquire(blocking=False):
        return jsonify({'ok': False,
                        'error': 'Đang gửi OTP cho tài khoản này; không gửi lại để tránh mã bị vô hiệu.'}), 409

    device_id = secrets.token_hex(8)
    try:
        response = requests.post(
            f'{BASE_URL}/quantri/user/xacthuc_tapdoan',
            json={'username': username, 'password': password,
                  'os_type': '1', 'device_id': device_id},
            headers={'Content-Type': 'application/json', 'Accept': 'application/json'},
            verify=False, timeout=15)
        data = response.json()
    except Exception as exc:
        return jsonify({'ok': False, 'error': f'Lỗi kết nối đăng nhập: {exc}'}), 502
    finally:
        otp_lock.release()
    if response.status_code != 200 or data.get('error_code') != 'BSS-00000000':
        return jsonify({'ok': False, 'error': data.get('message') or 'Tài khoản hoặc mật khẩu không đúng'}), 400
    secret_code = str((data.get('data') or {}).get('secretCode') or '')
    if not secret_code:
        return jsonify({'ok': False, 'error': 'OneBSS không trả secretCode để xác thực OTP'}), 502

    pending_id = secrets.token_urlsafe(18)
    pending = session.get('pending_extra_accounts') or {}
    pending = dict(pending) if isinstance(pending, dict) else {}
    pending[pending_id] = {
        'account_id': account_id,
        'username': username,
        'password': password,
        'device_id': device_id,
        'secret_code': secret_code,
        'created_at': time.time(),
    }
    pending = {key: value for key, value in pending.items()
               if time.time() - float(value.get('created_at') or 0) < 600}
    session['pending_extra_accounts'] = pending
    return jsonify({'ok': True, 'pending_id': pending_id, 'username': username,
                    'message': 'Đã gửi OTP. Nhập OTP để hoàn tất tài khoản phụ.'})


@app.post('/api/accounts/confirm')
@login_required
def api_accounts_confirm():
    payload = request.get_json(silent=True) or {}
    pending_id = str(payload.get('pending_id') or '').strip()
    otp = str(payload.get('otp') or '').strip()
    pending_map = session.get('pending_extra_accounts') or {}
    pending = pending_map.get(pending_id) if isinstance(pending_map, dict) else None
    if not pending or time.time() - float(pending.get('created_at') or 0) >= 600:
        return jsonify({'ok': False, 'error': 'Phiên thêm tài khoản đã hết hạn; hãy đăng nhập lại'}), 400
    if not otp:
        return jsonify({'ok': False, 'error': 'Nhập mã OTP'}), 400
    try:
        response = requests.post(
            f'{BASE_URL}/quantri/oauth/token',
            json={'grant_type': 'password', 'client_id': APP_CFG['CLIENT_ID'],
                  'client_secret': APP_CFG['CLIENT_SECRET'],
                  'secretCode': pending.get('secret_code'), 'otp': otp},
            headers={'Content-Type': 'application/json', 'Accept': 'application/json'},
            verify=False, timeout=25)
        data = response.json()
    except Exception as exc:
        return jsonify({'ok': False, 'error': f'Lỗi kết nối xác thực OTP: {exc}'}), 502
    if not data.get('access_token'):
        return jsonify({'ok': False, 'error': data.get('message') or 'OTP không hợp lệ hoặc đã hết hạn'}), 400

    account_id = str(pending.get('account_id'))
    context = {
        'id': account_id,
        'username': pending.get('username'),
        'access_token': data.get('access_token'),
        'refresh_token': data.get('refresh_token', ''),
        'expires_in': data.get('expires_in', 3600),
        'token_time': time.time(),
        'device_id': pending.get('device_id'),
        'app_secret': _build_app_secret_value(
            data.get('app_secret', ''), pending.get('device_id', '')),
    }
    with _multi_account_lock:
        extras = session.get('multi_accounts') or {}
        extras = dict(extras) if isinstance(extras, dict) else {}
        extras[account_id] = context
        session['multi_accounts'] = extras
        clean_pending = dict(pending_map)
        clean_pending.pop(pending_id, None)
        session['pending_extra_accounts'] = clean_pending
    try:
        save_employee_account(pending.get('username'), pending.get('password'))
    except Exception as exc:
        print(f'[WARN] Could not save extra Employee account: {exc}')
    _save_persistent_session()
    context['primary'] = False
    return jsonify({'ok': True, 'account': _public_account_summary(context)})


@app.post('/api/accounts/remove')
@login_required
def api_accounts_remove():
    account_id = str((request.get_json(silent=True) or {}).get('account_id') or '').strip()
    if not account_id or account_id == _primary_account_context().get('id'):
        return jsonify({'ok': False, 'error': 'Không thể gỡ phiên tài khoản chính'}), 400
    with _multi_account_lock:
        extras = session.get('multi_accounts') or {}
        extras = dict(extras) if isinstance(extras, dict) else {}
        removed = extras.pop(account_id, None)
        session['multi_accounts'] = extras
    _save_persistent_session()
    return jsonify({'ok': True, 'removed': bool(removed)})


@app.post('/api/accounts/refresh')
@login_required
def api_accounts_refresh():
    account_id = str((request.get_json(silent=True) or {}).get('account_id') or '').strip()
    try:
        context = _get_account_context(account_id)
    except ValueError as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 404
    if not context.get('refresh_token'):
        return jsonify({'ok': False, 'error': 'Tài khoản không có refresh token; hãy thêm lại và nhập OTP'}), 400
    try:
        response = requests.post(
            f'{BASE_URL}/quantri/oauth/token',
            json={'grant_type': 'refresh_token', 'refresh_token': context.get('refresh_token'),
                  'client_id': APP_CFG['CLIENT_ID'], 'client_secret': APP_CFG['CLIENT_SECRET']},
            headers={'Content-Type': 'application/json'}, verify=False, timeout=15)
        data = response.json()
    except Exception as exc:
        return jsonify({'ok': False, 'error': f'Lỗi làm mới token: {exc}'}), 502
    if not data.get('access_token'):
        return jsonify({'ok': False, 'error': data.get('message') or 'Không làm mới được token'}), 400
    context['access_token'] = data.get('access_token')
    context['refresh_token'] = data.get('refresh_token', context.get('refresh_token'))
    context['expires_in'] = data.get('expires_in', 3600)
    context['token_time'] = time.time()
    if data.get('app_secret'):
        context['app_secret'] = _build_app_secret_value(data.get('app_secret'), context.get('device_id'))
    if context.get('primary'):
        for key in ('access_token', 'refresh_token', 'expires_in', 'token_time', 'app_secret'):
            if key in context:
                session[key] = context[key]
    else:
        extras = dict(session.get('multi_accounts') or {})
        stored = dict(context)
        stored.pop('primary', None)
        extras[account_id] = stored
        session['multi_accounts'] = extras
    _save_persistent_session()
    return jsonify({'ok': True, 'account': _public_account_summary(context)})

# ── Refresh token ──
@app.route('/token/refresh', methods=['POST'])
@login_required
def refresh_token():
    try:
        resp = requests.post(
            f'{BASE_URL}/quantri/oauth/token',
            json={'grant_type':    'refresh_token',
                  'refresh_token': session.get('refresh_token', ''),
                  'client_id':     APP_CFG['CLIENT_ID'],
                  'client_secret': APP_CFG['CLIENT_SECRET']},
            headers={'Content-Type': 'application/json'},
            verify=False, timeout=10
        )
        data = resp.json()
        if data.get('access_token'):
            session['access_token']  = data['access_token']
            session['refresh_token'] = data.get('refresh_token', session['refresh_token'])
            session['token_time']    = time.time()
            return jsonify({'ok': True, 'expires_in': data.get('expires_in', 3600)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    return jsonify({'ok': False, 'error': 'Failed'})

# ─────────────────────────────────────────────────────────────
#  Dashboard
# ─────────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    elapsed    = time.time() - session.get('token_time', time.time())
    expires_in = max(0, int(session.get('expires_in', 3600) - elapsed))
    return render_template('dashboard.html',
        sections=SECTIONS,
        user=session.get('username', ''),
        base_url=BASE_URL,
        expires_in=expires_in,
        app_cfg=APP_CFG,
        CLIENT_ID=APP_CFG['CLIENT_ID'])

# ─────────────────────────────────────────────────────────────
#  Generic proxy (tất cả API call qua đây)
# ─────────────────────────────────────────────────────────────
@app.route('/proxy', methods=['POST'])
@login_required
def proxy():
    data      = request.json or {}
    endpoint  = data.get('endpoint', '')
    account_id = str(data.get('account_id') or '').strip()
    method    = data.get('method', 'GET').upper()
    body      = data.get('body', {})
    # Work on our own object: normalising/injecting proxy metadata must not
    # mutate Flask's parsed request JSON.
    if isinstance(body, dict):
        body = dict(body)
    body_type = data.get('body_type', 'json')
    extra_hdr = data.get('headers', {})
    if not isinstance(extra_hdr, dict):
        extra_hdr = {}
    timeout   = int(data.get('timeout', 15))

    if not endpoint:
        return jsonify({'error': 'endpoint required'}), 400

    url = endpoint if endpoint.startswith('http') else \
          BASE_URL.rstrip('/') + '/' + endpoint.lstrip('/')

    endpoint_path = endpoint.split('?', 1)[0].rstrip('/')
    try:
        account_context = _get_account_context(account_id)
    except ValueError as exc:
        return jsonify({'error': str(exc), 'status': 401}), 401
    body = normalize_onebss_body(
        endpoint_path, body, endpoint, account_context.get('username'))
    requested_mid = extra_hdr.get('SelectedMenuId') or extra_hdr.get('selectedmenuid')
    active_mid = endpoint_menu_id(endpoint_path, requested_mid, body)
    # app_ds_dauso is the only chonSo call captured with an empty DTO. The
    # mobile search_isdn DTO explicitly contains menu_id=699161, so preserve
    # and inject it like the other SIM-kit requests.
    strict_body_endpoints = {
        '/ccbs/chonSo/app_ds_dauso',
        # Mobile sends only ?so_msin=... for this final read-only status check.
        '/ccbs/chonSo/checkSimStatus',
    }
    inject_menu_id = endpoint_path not in strict_body_endpoints
    if inject_menu_id and isinstance(body, dict) and 'menu_id' not in body:
        body['menu_id'] = active_mid

    hdrs = get_headers(active_mid, account_id)
    protected_headers = {'authorization', 'app-secret'}
    hdrs.update({key: value for key, value in extra_hdr.items()
                 if str(key).casefold() not in protected_headers})
    # Route metadata is authoritative for known OneBSS modules.  Avoid an old
    # caller-provided header silently undoing the correction above.
    if any(endpoint_path.startswith(prefix) for prefix, _ in ENDPOINT_MENU_ROUTES):
        hdrs['SelectedMenuId'] = active_mid
        hdrs['selectedmenuid'] = active_mid
    kw = dict(headers=hdrs, verify=False, timeout=timeout)

    t0 = time.time()
    try:
        method_used = method
        retried_from_get = False
        retried_from_post = False
        if method == 'GET':
            resp = requests.get(url, params=body, **kw)
            # The current gateway exposes a number of read-only routes as POST
            # although the mobile catalog still labels them GET. Retry only a
            # server-declared 405 and only for OneBSS application routes.
            if (resp.status_code == 405 and
                    endpoint_path.startswith(('/app-', '/ccbs/'))):
                resp = requests.post(url, json=body, **kw)
                method_used = 'POST'
                retried_from_get = True
        elif body_type == 'form':
            hdrs['Content-Type'] = 'application/x-www-form-urlencoded'
            resp = requests.request(method, url, data=body, **kw)
        else:
            resp = requests.request(method, url, json=body, **kw)
        if (method == 'POST' and resp.status_code == 405 and
                endpoint_path in READ_ONLY_ENDPOINTS):
            resp = requests.get(url, params=body, **kw)
            method_used = 'GET'
            retried_from_post = True
        elapsed = round((time.time()-t0)*1000)
        try: rb = resp.json()
        except: rb = resp.text
        return jsonify({'status': resp.status_code, 'elapsed': elapsed,
                        'method_used': method_used,
                        'retried_from_get': retried_from_get,
                        'retried_from_post': retried_from_post,
                        'selected_menu_id': active_mid,
                        'headers': dict(resp.headers), 'body': rb})
    except requests.exceptions.ConnectionError as e:
        return jsonify({'error': f'Lỗi kết nối: {e}'}), 502
    except requests.exceptions.Timeout:
        return jsonify({'error': 'Request timeout'}), 504
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/proxy-image', methods=['GET'])
@login_required
def proxy_image():
    """Tải ảnh OneBSS bằng session headers rồi trả ảnh về cùng origin.

    Không để thẻ <img> gọi thẳng api-onebss.vnpt.vn vì request của thẻ ảnh
    không gửi được authorization/app-secret như luồng app.js.
    """
    raw_url = (request.args.get('url') or '').strip()
    if not raw_url:
        return jsonify({'error': 'url required'}), 400

    # Cho phép URL tuyệt đối của đúng API host hoặc path tương đối.
    if raw_url.startswith('//'):
        target_url = 'https:' + raw_url
    elif raw_url.startswith('/'):
        target_url = BASE_URL.rstrip('/') + raw_url
    elif raw_url.lower().startswith(('http://', 'https://')):
        target_url = raw_url
    else:
        target_url = BASE_URL.rstrip('/') + '/' + raw_url.lstrip('/')

    target = urlparse(target_url)
    base = urlparse(BASE_URL)
    if (target.scheme not in ('http', 'https') or
            target.hostname != base.hostname):
        return jsonify({'error': 'image host not allowed'}), 403

    headers = get_headers()
    headers.pop('Content-Type', None)
    headers['Accept'] = 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8'
    last_error = None
    for attempt in range(3):
        try:
            upstream = requests.get(target_url, headers=headers,
                                    verify=False, timeout=15)
            if 200 <= upstream.status_code < 300:
                content_type = (upstream.headers.get('Content-Type') or
                                'application/octet-stream').split(';', 1)[0].strip().lower()
                if not content_type.startswith('image/'):
                    return jsonify({'error': f'upstream did not return an image ({content_type})'}), 502
                return Response(upstream.content, status=200, mimetype=content_type,
                                headers={'Cache-Control': 'private, max-age=300'})
            last_error = f'upstream image status {upstream.status_code}'
            if upstream.status_code < 500:
                break
        except requests.exceptions.RequestException as exc:
            last_error = f'image connection failed: {exc}'
        if attempt < 2:
            time.sleep(0.4)

    return jsonify({'error': last_error or 'image request failed'}), 502

@app.route('/session/info')
@login_required
def session_info():
    elapsed  = time.time() - session.get('token_time', time.time())
    exp_left = max(0, int(session.get('expires_in', 3600) - elapsed))
    token    = session.get('access_token', '')
    token_claims = {}
    try:
        payload_part = token.split('.')[1]
        payload_part += '=' * (-len(payload_part) % 4)
        token_claims = json.loads(base64.urlsafe_b64decode(payload_part).decode('utf-8'))
    except (IndexError, ValueError, TypeError, UnicodeDecodeError,
            json.JSONDecodeError):
        token_claims = {}
    return jsonify({
        'user':           session.get('username', ''),
        # Employee mobile uses this claim as p_ma_hrm in the SIM-change DTO.
        'staff_code':     token_claims.get('user_vi') or
                          token_claims.get('ma_nhanvien_ccbs') or '',
        'base_url':       BASE_URL,
        'client_id':      APP_CFG['CLIENT_ID'],
        'expires_in':     exp_left,
        'token_prefix':   token[:30] + '...' if len(token) > 30 else token,
        'full_token':     token,
        'app_secret':     session.get('app_secret', '')[:40] + '...',
        'selected_menu':  session.get('active_menu_id', APP_CFG['SELECTED_MENU']),
        'menu_count':     len(session.get('menus', [])),
    })


@app.route('/api/menus')
@login_required
def api_menus():
    """Trả về danh sách menu đã load từ khoitao_ungdung, phân theo parent."""
    menus = session.get('menus', [])
    active = session.get('active_menu_id', APP_CFG['SELECTED_MENU'])

    # Auto-fetch nếu session chưa có menu (user login trước khi deploy code mới)
    if not menus and session.get('access_token'):
        try:
            _h = {
                'Content-Type':  'application/json',
                'Accept':        'application/json',
                'authorization': f"Bearer {session.get('access_token', '')}",
                'App-secret':    build_app_secret(),
                'SelectedMenuId': active,
                'selectedmenuid': active,
            }
            _rx = requests.get(
                f"{BASE_URL}/quantri/user/khoitao_ungdung?p_idmodule=21",
                headers=_h, verify=False, timeout=15
            )
            if _rx.status_code == 200:
                _d = _rx.json()
                menus = (_d.get('data') or {}).get('ds_chucnang') or []
                if not menus and isinstance(_d.get('data'), list):
                    menus = _d['data']
                session['menus'] = menus
                print(f"[OK] api_menus: auto-fetched {len(menus)} items")
        except Exception as _e:
            print(f"[WARN] api_menus auto-fetch: {_e}")

    # Build tree: root (p_id=None) + children
    roots = [m for m in menus if not m.get('p_id')]
    children_map = {}
    for m in menus:
        pid = m.get('p_id')
        if pid:
            children_map.setdefault(pid, []).append(m)

    tree = []
    for root in sorted(roots, key=lambda x: x.get('stt', 0)):
        tree.append({
            'id':       root.get('id'),
            'name':     root.get('name', root.get('displayName', '')),
            'icon':     root.get('icon', ''),
            'url':      root.get('url'),
            'level':    root.get('level', 1),
            'children': sorted(
                children_map.get(root.get('id'), []),
                key=lambda x: x.get('stt', 0)
            )
        })

    return jsonify({'tree': tree, 'flat': menus, 'active': active})


@app.route('/api/debug_menus')
@login_required
def api_debug_menus():
    """Debug: xem raw menus data trong session"""
    menus = session.get('menus', [])
    if not menus:
        return jsonify({'count': 0, 'sample': [], 'keys': [], 'tree_roots': 0})
    sample = menus[:5]
    keys = list(sample[0].keys()) if sample and isinstance(sample[0], dict) else []
    roots = [m for m in menus if not m.get('p_id')]
    return jsonify({
        'count': len(menus),
        'root_count': len(roots),
        'keys': keys,
        'sample': sample,
        'roots_sample': roots[:5]
    })


@app.route('/api/set_menu', methods=['POST'])
@login_required
def api_set_menu():
    """Đổi active menu ID cho các API call tiếp theo."""
    mid = str(request.json.get('menu_id', APP_CFG['SELECTED_MENU']))
    session['active_menu_id'] = mid
    return jsonify({'ok': True, 'active_menu_id': mid})


@app.route('/api/sim-batch/files')
@login_required
def api_sim_batch_files():
    """Create and return the stable SIM batch input/output workbooks."""
    try:
        with _sim_batch_file_lock:
            _ensure_sim_batch_files()
        return jsonify({
            'ok': True,
            'input_path': SIM_BATCH_INPUT_FILE,
            'output_path': SIM_BATCH_OUTPUT_FILE,
            'output_dir': SIM_BATCH_DIR,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/sim-batch/open', methods=['POST'])
@login_required
def api_sim_batch_open():
    """Open only one of the two application-owned workbooks."""
    payload = request.get_json(silent=True) or {}
    kind = str(payload.get('type', '')).strip().lower()
    path = {'input': SIM_BATCH_INPUT_FILE, 'output': SIM_BATCH_OUTPUT_FILE}.get(kind)
    if kind == 'output':
        path = _owned_batch_file_path(payload.get('path'), SIM_BATCH_DIR, SIM_BATCH_OUTPUT_FILE)
    if not path:
        return jsonify({'ok': False, 'error': 'type phải là input hoặc output'}), 400
    try:
        with _sim_batch_file_lock:
            _ensure_sim_batch_files()
        _open_local_file(path)
        return jsonify({'ok': True, 'path': path})
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/sim-batch/read-input')
@login_required
def api_sim_batch_read_input():
    """Read the application-owned input workbook for one-click batch runs."""
    try:
        from openpyxl import load_workbook
        with _sim_batch_file_lock:
            _ensure_sim_batch_files()
            workbook = load_workbook(SIM_BATCH_INPUT_FILE, data_only=True, read_only=True)
            sheet = workbook['SIM_Input'] if 'SIM_Input' in workbook.sheetnames else workbook.active
            rows = [[cell if cell is not None else '' for cell in row] for row in sheet.iter_rows(values_only=True)]
            workbook.close()
        return jsonify({
            'ok': True,
            'file_name': os.path.basename(SIM_BATCH_INPUT_FILE),
            'sheet_name': sheet.title,
            'rows': rows,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/sim-batch/append-output', methods=['POST'])
@login_required
def api_sim_batch_append_output():
    """Write one completed run to a new workbook; never merge old runs."""
    payload = request.get_json(silent=True) or {}
    table = payload.get('table')
    if not isinstance(table, list) or len(table) < 2 or not isinstance(table[0], list):
        return jsonify({'ok': False, 'error': 'Thiếu bảng kết quả gồm header và dữ liệu'}), 400
    if len(table) > 5001:
        return jsonify({'ok': False, 'error': 'Mỗi lần chỉ ghi tối đa 5000 dòng'}), 400
    raw_headers = table[0][:100]
    kept_columns = [
        index for index, value in enumerate(raw_headers)
        if str(value or '').strip().casefold() != 'địa chỉ đầy đủ'.casefold()
    ]
    incoming_headers = [
        str(raw_headers[index] or '').strip() or f'Cột {index + 1}'
        for index in kept_columns
    ]
    incoming_rows = [
        [row[index] if index < len(row) else '' for index in kept_columns]
        for row in table[1:] if isinstance(row, list)
    ]
    if not incoming_rows:
        return jsonify({'ok': False, 'error': 'Không có dòng kết quả để ghi'}), 400
    try:
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
        with _sim_batch_file_lock:
            _ensure_sim_batch_files()
            output_path = _new_batch_output_path(SIM_BATCH_DIR, 'SIM_Kit_Output')
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'SIM_Output'
            sheet.append(incoming_headers)
            user_name, authenticated_users = _authenticated_batch_user_map()
            user_index = next((index for index, header in enumerate(incoming_headers)
                               if header.casefold() == 'user chạy'.casefold()), None)
            for incoming in incoming_rows:
                mapped = (list(incoming) + [''] * len(incoming_headers))[:len(incoming_headers)]
                if user_index is not None:
                    requested_user = str(mapped[user_index] or '').strip().casefold()
                    mapped[user_index] = authenticated_users.get(requested_user, user_name)
                sheet.append(mapped)
            for index, header in enumerate(incoming_headers, start=1):
                width = 80 if header.casefold() in {'kết quả'.casefold(), 'địa chỉ'.casefold()} else (24 if index > 2 else 20)
                sheet.column_dimensions[get_column_letter(index)].width = width
            sheet.freeze_panes = 'A2'
            workbook.save(output_path)
            workbook.close()
        return jsonify({'ok': True, 'appended': len(incoming_rows),
                        'path': output_path, 'total_rows': len(incoming_rows),
                        'new_file': True})
    except PermissionError:
        return jsonify({
            'ok': False,
            'error': 'File output đang mở trong Excel. Hãy đóng file rồi bấm ghi lại kết quả.'
        }), 409
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/icoc-batch/files')
@login_required
def api_icoc_batch_files():
    """Create and return the stable IC/OC batch input/output workbooks."""
    try:
        with _icoc_batch_file_lock:
            _ensure_icoc_batch_files()
        return jsonify({
            'ok': True,
            'input_path': ICOC_BATCH_INPUT_FILE,
            'output_path': ICOC_BATCH_OUTPUT_FILE,
            'output_dir': ICOC_BATCH_DIR,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/icoc-batch/open', methods=['POST'])
@login_required
def api_icoc_batch_open():
    payload = request.get_json(silent=True) or {}
    kind = str(payload.get('type', '')).strip().lower()
    path = {'input': ICOC_BATCH_INPUT_FILE, 'output': ICOC_BATCH_OUTPUT_FILE}.get(kind)
    if kind == 'output':
        path = _owned_batch_file_path(payload.get('path'), ICOC_BATCH_DIR, ICOC_BATCH_OUTPUT_FILE)
    if not path:
        return jsonify({'ok': False, 'error': 'type phải là input hoặc output'}), 400
    try:
        with _icoc_batch_file_lock:
            _ensure_icoc_batch_files()
        _open_local_file(path)
        return jsonify({'ok': True, 'path': path})
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/icoc-batch/read-input')
@login_required
def api_icoc_batch_read_input():
    try:
        from openpyxl import load_workbook
        with _icoc_batch_file_lock:
            _ensure_icoc_batch_files()
            workbook = load_workbook(ICOC_BATCH_INPUT_FILE, data_only=True, read_only=True)
            sheet = workbook['IC_OC_Input'] if 'IC_OC_Input' in workbook.sheetnames else workbook.active
            rows = [[cell if cell is not None else '' for cell in row] for row in sheet.iter_rows(values_only=True)]
            sheet_name = sheet.title
            workbook.close()
        return jsonify({
            'ok': True,
            'file_name': os.path.basename(ICOC_BATCH_INPUT_FILE),
            'sheet_name': sheet_name,
            'rows': rows,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/icoc-batch/append-output', methods=['POST'])
@login_required
def api_icoc_batch_append_output():
    """Write one completed IC/OC run to a new workbook; never merge old runs."""
    payload = request.get_json(silent=True) or {}
    table = payload.get('table')
    if not isinstance(table, list) or len(table) < 2 or not isinstance(table[0], list):
        return jsonify({'ok': False, 'error': 'Thiếu bảng kết quả gồm header và dữ liệu'}), 400
    if len(table) > 5001:
        return jsonify({'ok': False, 'error': 'Mỗi lần chỉ ghi tối đa 5000 dòng'}), 400
    incoming_headers = [
        str(value or '').strip() or f'Cột {index + 1}'
        for index, value in enumerate(table[0][:100])
    ]
    incoming_rows = [
        [row[index] if index < len(row) else '' for index in range(len(incoming_headers))]
        for row in table[1:] if isinstance(row, list)
    ]
    if not incoming_rows:
        return jsonify({'ok': False, 'error': 'Không có dòng kết quả để ghi'}), 400

    try:
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
        with _icoc_batch_file_lock:
            _ensure_icoc_batch_files()
            output_path = _new_batch_output_path(ICOC_BATCH_DIR, 'IC_OC_Output')
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'IC_OC_Output'
            sheet.append(incoming_headers)
            user_name, authenticated_users = _authenticated_batch_user_map()
            user_index = next((index for index, header in enumerate(incoming_headers)
                               if header.casefold() == 'user chạy'.casefold()), None)
            for incoming in incoming_rows:
                mapped = (list(incoming) + [''] * len(incoming_headers))[:len(incoming_headers)]
                if user_index is not None:
                    requested_user = str(mapped[user_index] or '').strip().casefold()
                    mapped[user_index] = authenticated_users.get(requested_user, user_name)
                sheet.append(mapped)
            for index, header in enumerate(incoming_headers, start=1):
                width = 80 if header.casefold() == 'kết quả'.casefold() else (24 if index > 1 else 20)
                sheet.column_dimensions[get_column_letter(index)].width = width
            sheet.freeze_panes = 'A2'
            workbook.save(output_path)
            workbook.close()
        return jsonify({'ok': True, 'appended': len(incoming_rows),
                        'path': output_path, 'total_rows': len(incoming_rows),
                        'new_file': True})
    except PermissionError:
        return jsonify({
            'ok': False,
            'error': 'File output đang mở trong Excel. Hãy đóng file rồi bấm ghi lại kết quả.'
        }), 409
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/reload_menus')
@login_required
def api_reload_menus():
    """Reload danh sách menu từ server (không cần logout)."""
    try:
        h = {
            'Content-Type':  'application/json',
            'Accept':        'application/json',
            'authorization': f"Bearer {session.get('access_token', '')}",
            'App-secret':    build_app_secret(),
            'SelectedMenuId': session.get('active_menu_id', APP_CFG['SELECTED_MENU']),
            'selectedmenuid': session.get('active_menu_id', APP_CFG['SELECTED_MENU']),
        }
        rx = requests.get(
            f"{BASE_URL}/quantri/user/khoitao_ungdung?p_idmodule=21",
            headers=h, verify=False, timeout=15
        )
        if rx.status_code == 200:
            d = rx.json()
            menus = (d.get('data') or {}).get('ds_chucnang') or []
            if not menus and isinstance(d.get('data'), list):
                menus = d['data']
            session['menus'] = menus
            return jsonify({'ok': True, 'count': len(menus)})
        return jsonify({'ok': False, 'status': rx.status_code, 'body': rx.text[:300]})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/dump_menus')
@login_required
def dump_menus():
    """Route cũ: brute-force menu ID range."""
    try:
        import json as _json, base64
        app_secret_obj = {
            'device_id': session.get('device_id', ''),
            'device_ip': 'Unknown',
            'device_name': 'Web-Browser',
            'device_model': 'Web',
            'os_type': 'Android',
            'os_version': 'Chrome',
            'app_version': APP_CFG['APP_VERSION']
        }
        h_app_secret = base64.b64encode(_json.dumps(app_secret_obj, separators=(',', ':')).encode()).decode()
        h = {
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'app-secret': h_app_secret,
            'authorization': f"Bearer {session.get('access_token')}"
        }
        
        results = {}
        for mid in range(810239, 810244):
            h['selectedmenuid'] = str(mid)
            rx = requests.post(f"{BASE_URL}/quantri/user/thongtin_nv", headers=h, json={}, verify=False, timeout=5)
            try:
                results[mid] = {'status': rx.status_code, 'resp': rx.json()}
            except:
                results[mid] = {'status': rx.status_code, 'resp': rx.text[:100]}
        
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/dump_menus_v2')
@login_required
def dump_menus_v2():
    """
    Gọi đúng API khoitao_ungdung_v2 để lấy danh sách chức năng (menu).
    Phân tích từ blutter pp.txt:
      - Endpoint: /quantri/user/khoitao_ungdung_v2?p_idmodule=21
      - Header key đúng: "SelectedMenuId" (có viết hoa)
      - Body params: p_device_id, p_phanvung_id, p_refresh_token
      - SQLite table schema: chuc_nang(id, p_id, name, url, level, menu_type, ...)
    """
    try:
        import json as _json, base64
        device_id = session.get('device_id', '0f8c2d3fb0c51653')
        app_secret_obj = {
            'device_id':   device_id,
            'device_ip':   'Unknown',
            'device_name': 'Web-Browser',
            'mac_address': 'Unknown',
            'mobile_id':   'web-generated-id',
            'app_id':      '1',
            'app_version': APP_CFG['APP_VERSION'],
            'os_version':  'Chrome/Web'
        }
        h_app_secret = base64.b64encode(_json.dumps(app_secret_obj, separators=(',', ':')).encode()).decode()
        token = session.get('access_token', '')

        # Header dùng cả 2 biến thể (viết hoa và lowercase)
        h = {
            'Content-Type':  'application/json',
            'Accept':        'application/json',
            'authorization': f"Bearer {token}",
            'App-secret':    h_app_secret,
            'SelectedMenuId': APP_CFG['SELECTED_MENU'],
            'selectedmenuid': APP_CFG['SELECTED_MENU'],
        }

        # Gọi v2 (mới hơn)
        body_v2 = {
            'p_device_id':     device_id,
            'p_phanvung_id':   0,
            'p_refresh_token': session.get('refresh_token', ''),
        }
        rx2 = requests.post(
            f"{BASE_URL}/quantri/user/khoitao_ungdung_v2?p_idmodule=21",
            headers=h, json=body_v2, verify=False, timeout=15
        )

        # Gọi v1 fallback (GET)
        rx1 = requests.get(
            f"{BASE_URL}/quantri/user/khoitao_ungdung?p_idmodule=21",
            headers=h, verify=False, timeout=15
        )

        result = {
            'v2': {
                'status': rx2.status_code,
                'body':   rx2.json() if rx2.headers.get('content-type','').startswith('application/json') else rx2.text[:2000],
            },
            'v1': {
                'status': rx1.status_code,
                'body':   rx1.json() if rx1.headers.get('content-type','').startswith('application/json') else rx1.text[:2000],
            },
            'note': (
                'chuc_nang table schema (từ blutter): '
                'id(=menu_id), p_id(parent), name, url, level, icon, menu_type(1=folder), '
                'is_required, isDependantHRM, min_ver'
            )
        }

        # Parse và format menu tree nếu có data
        for ver_key, resp_obj in [('v2', rx2), ('v1', rx1)]:
            try:
                data = resp_obj.json()
                # Tìm danh sách chuc_nang trong response
                menus = None
                if isinstance(data, dict):
                    for k in ('data', 'chuc_nang', 'ds_chucnang', 'result', 'menus'):
                        if k in data and isinstance(data[k], list):
                            menus = data[k]
                            break
                    if menus is None and 'data' in data and isinstance(data['data'], dict):
                        for k in ('chuc_nang', 'ds_chucnang', 'menus', 'functions'):
                            if k in data['data'] and isinstance(data['data'][k], list):
                                menus = data['data'][k]
                                break
                elif isinstance(data, list):
                    menus = data

                if menus:
                    result[ver_key]['menu_count'] = len(menus)
                    result[ver_key]['menus_flat'] = [
                        {
                            'id':        m.get('id', m.get('menu_id', '?')),
                            'p_id':      m.get('p_id', m.get('parent_id', 0)),
                            'name':      m.get('name', m.get('ten', '?')),
                            'url':       m.get('url', ''),
                            'level':     m.get('level', 0),
                            'menu_type': m.get('menu_type', 0),
                        }
                        for m in menus
                    ]
            except Exception:
                pass

        return jsonify(result)
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()})

# =============================================================================
# ĐKTTTB (eKYC) ROUTES
# =============================================================================
@app.route('/dktttb')
@login_required
def dktttb_page():
    """Trang Đăng Ký Thông Tin Thuê Bao mới."""
    if EKYC_AVAILABLE:
        _ekyc.cleanup_old_tasks()
    return render_template('dktttb.html',
        user=session.get('username', ''),
        ekyc_available=EKYC_AVAILABLE)


@app.route('/api/dktttb/run', methods=['POST'])
@login_required
def dktttb_run():
    """Khởi chạy 1 luồng ĐKTTTB (background). Trả về task_id."""
    if not EKYC_AVAILABLE:
        return jsonify({'ok': False, 'error': 'ekyc_service chưa được cài đặt'}), 503

    phone    = (request.form.get('phone')    or '').strip()
    password = (request.form.get('password') or '').strip()
    otp      = (request.form.get('otp')      or '').strip()
    nfc_raw  = (request.form.get('nfc_data') or '').strip()

    if not phone:
        return jsonify({'ok': False, 'error': 'Thiếu số điện thoại'}), 400
    if not password and not otp:
        return jsonify({'ok': False, 'error': 'Cần nhập mật khẩu hoặc OTP'}), 400

    # Đọc file ảnh nếu có
    def _read_file(field: str) -> bytes | None:
        f = request.files.get(field)
        return f.read() if f and f.filename else None

    face_bytes       = _read_file('face_image')
    cccd_front_bytes = _read_file('cccd_front')
    cccd_back_bytes  = _read_file('cccd_back')
    digital_sig      = _read_file('digital_sig')

    nfc_data = None
    if nfc_raw:
        try:
            nfc_data = json.loads(nfc_raw)
        except Exception:
            nfc_data = {'raw': nfc_raw}

    onebss_token = session.get('access_token', '')  # token ONEBSS hiện tại trong session

    task_id = _ekyc.start_dktttb_background(
        phone=phone, password=password, otp=otp,
        face_bytes=face_bytes,
        cccd_front_bytes=cccd_front_bytes,
        cccd_back_bytes=cccd_back_bytes,
        nfc_data=nfc_data,
        digital_sig=digital_sig,
        onebss_token=onebss_token,
    )
    return jsonify({'ok': True, 'task_id': task_id})


@app.route('/api/dktttb/status/<task_id>')
@login_required
def dktttb_status(task_id: str):
    """Polling kết quả task (logs + status)."""
    if not EKYC_AVAILABLE:
        return jsonify({'error': 'ekyc_service unavailable'}), 503
    task = _ekyc.get_task(task_id)
    if not task:
        return jsonify({'error': 'task not found'}), 404
    return jsonify({
        'status':  task.get('status'),
        'logs':    task.get('logs', []),
        'result':  task.get('result'),
        'elapsed': task.get('elapsed', 0),
    })


@app.route('/api/dktttb/batch', methods=['POST'])
@login_required
def dktttb_batch():
    """
    Chạy ĐKTTTB cho nhiều số cùng lúc.
    Body JSON: {items: [{phone, password, otp?}, ...], onebss_token?}
    Trả về danh sách task_ids.
    """
    if not EKYC_AVAILABLE:
        return jsonify({'ok': False, 'error': 'ekyc_service chưa được cài đặt'}), 503

    data  = request.get_json(silent=True) or {}
    items = data.get('items', [])
    if not items or not isinstance(items, list):
        return jsonify({'ok': False, 'error': 'Cần truyền items: [{phone, password}...]'}), 400

    onebss_token = session.get('access_token', '')
    task_ids = []
    for item in items[:50]:   # giới hạn 50 số/batch
        phone    = str(item.get('phone', '')).strip()
        password = str(item.get('password', '')).strip()
        otp      = str(item.get('otp', '')).strip()
        if not phone or (not password and not otp):
            continue
        tid = _ekyc.start_dktttb_background(
            phone=phone, password=password, otp=otp,
            onebss_token=onebss_token,
        )
        task_ids.append({'phone': phone, 'task_id': tid})

    return jsonify({'ok': True, 'tasks': task_ids, 'count': len(task_ids)})


# =============================================================================
# MAIN
# =============================================================================
if __name__ == '__main__':
    total = sum(len(s['apis']) for s in SECTIONS)
    print(f"[OK] {total} APIs | Base: {BASE_URL} | client_id: {APP_CFG['CLIENT_ID']}")
    print(f"[OK] eKYC/ĐKTTTB module: {'available' if EKYC_AVAILABLE else 'NOT AVAILABLE'}")
    app.run(debug=True, host='0.0.0.0', port=5056)
