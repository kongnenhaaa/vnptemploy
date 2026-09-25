"""
VNPT Employee Web App  ·  app.py  (v2 - OneBSS auth flow)
Auth flow khớp với app.js gốc:
  Step 1: POST /quantri/user/xacthuc_tapdoan  → secretCode
  Step 2: POST /quantri/oauth/token (secretCode + OTP) → access_token
"""
import os, re, json, time, base64, hashlib, threading, subprocess, shutil, secrets, sqlite3
import sys
from contextlib import contextmanager, nullcontext
from functools import wraps
from flask import (Flask, render_template, request, session,
                   redirect, url_for, jsonify, flash, Response, send_file)
import requests
import urllib3
from urllib.parse import urlparse, parse_qsl, unquote
from flask_session import Session
from onebss_report_batch import OneBssReportBatchManager, OneBssReportError
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
SIM_BATCH_RUN_PREFIX = 'SIM_Kit_Phien'
SIM_BATCH_OUTPUT_HEADERS = (
    'SĐT', 'Serial SIM', 'Kết quả', 'Mã đơn hàng', 'Loại đăng ký',
    'Tên khách hàng', 'Địa chỉ khách hàng', 'Số tiền', 'Thời gian', 'User chạy'
)
SIM_BATCH_OUTPUT_META_SHEET = '_Luu_tu_dong'
_sim_batch_file_lock = threading.Lock()

# Tra cứu seri SIM dùng file 1 cột input và file 2 cột output cố định.
# Output là nhật ký cộng dồn: mỗi dòng tra xong được ghi ngay để
# thao tác Dừng không làm mất phần đã hoàn tất.
SERIAL_LOOKUP_BATCH_DIR = os.path.join(_documents_root, 'VNPTEmploy', 'Tra_Cuu_Seri')
SERIAL_LOOKUP_INPUT_FILE = os.path.join(SERIAL_LOOKUP_BATCH_DIR, 'Tra_Cuu_Seri_Input.xlsx')
SERIAL_LOOKUP_OUTPUT_FILE = os.path.join(SERIAL_LOOKUP_BATCH_DIR, 'Tra_Cuu_Seri_Output.xlsx')
SERIAL_LOOKUP_OUTPUT_HEADERS = ('MSIN / Seri SIM', 'SĐT')
SERIAL_LOOKUP_OUTPUT_META_SHEET = '_Luu_tu_dong'
_serial_lookup_file_lock = threading.Lock()

# Batch thao tác IC/OC dùng workbook riêng để không lẫn Serial SIM của luồng
# khởi tạo SIM Kit. File được giữ trong Documents giống batch SIM Kit.
ICOC_BATCH_DIR = os.path.join(_documents_root, 'VNPTEmploy', 'IC_OC_Batch')
ICOC_BATCH_INPUT_FILE = os.path.join(ICOC_BATCH_DIR, 'IC_OC_Input.xlsx')
ICOC_BATCH_OUTPUT_FILE = os.path.join(ICOC_BATCH_DIR, 'IC_OC_Output.xlsx')
ICOC_BATCH_RUN_PREFIX = 'IC_OC_Phien'
ICOC_BATCH_OUTPUT_HEADERS = ('SĐT', 'Kết quả', 'Thời gian', 'User chạy')
ICOC_BATCH_OUTPUT_META_SHEET = '_Luu_tu_dong'
_icoc_batch_file_lock = threading.Lock()

# Đăng ký gói cước (menu 11077): output phiên và nhật ký cộng dồn. Mỗi kết
# quả được ghi ngay sau khi OneBSS trả lời để dừng/lỗi giữa danh sách không làm
# mất các thuê bao đã xử lý.
PACKAGE_REGISTRATION_DIR = os.path.join(
    _documents_root, 'VNPTEmploy', 'Dang_Ky_Goi_Cuoc')
PACKAGE_REGISTRATION_OUTPUT_FILE = os.path.join(
    PACKAGE_REGISTRATION_DIR, 'Dang_Ky_Goi_Cuoc_Output.xlsx')
PACKAGE_REGISTRATION_RUN_PREFIX = 'Dang_Ky_Goi_Cuoc_Phien'
PACKAGE_REGISTRATION_HEADERS = (
    'STT', 'SĐT', 'Mã gói', 'Hệ thống', 'Chu kỳ (ngày)', 'Giá',
    'Trạng thái', 'Kết quả', 'Thời gian', 'User chạy',
)
PACKAGE_REGISTRATION_META_SHEET = '_Luu_tu_dong'
_package_registration_file_lock = threading.Lock()

# OneBSS report 378 shown in the Employ dashboard.  The login is maintained by
# employ_web and refreshed there; the worker deliberately re-reads this file
# before every day so a renewed session is picked up without restarting Employ.
ONEBSS_REPORT_SESSION_FILE = os.environ.get(
    'VNPT_EMPLOY_ONEBSS_SESSION_FILE',
    os.path.join(
        os.environ.get('USERPROFILE') or os.path.expanduser('~'),
        'Pictures', 'employ_web', 'onebss_session.json',
    ),
)
ONEBSS_REPORT_OUTPUT_DIR = os.path.join(
    _documents_root, 'VNPTEmploy', 'Bao_Cao_OneBSS')
_onebss_report_manager = OneBssReportBatchManager(
    ONEBSS_REPORT_SESSION_FILE,
    ONEBSS_REPORT_OUTPUT_DIR,
)


def _batch_header_key(value):
    return re.sub(r'\s+', ' ', str(value or '').strip()).casefold()


def _sim_batch_legacy_details(result):
    """Recover separated fields from the pipe-delimited result used by old builds."""
    text = str(result or '').strip()
    details = {'order': '', 'registration': '', 'customer': '', 'address': '', 'amount': ''}
    order_match = re.search(r'(?:đơn|mã đơn hàng)\s*[:#-]?\s*([0-9]+)', text, re.IGNORECASE)
    if order_match:
        details['order'] = order_match.group(1)
    segments = [segment.strip() for segment in text.split('|') if segment.strip()]
    order_index = next((index for index, value in enumerate(segments)
                        if re.match(r'^(?:đơn|mã đơn hàng)\b', value, re.IGNORECASE)), None)
    if order_index is not None and order_index + 1 < len(segments):
        tail = segments[order_index + 1:]
        if tail:
            details['customer'] = tail[0]
        if len(tail) > 1 and ('đăng ký' in tail[1].casefold() or 'nhân viên hỗ trợ' in tail[1].casefold()):
            details['registration'] = tail[1]
        if len(tail) > 2 and re.search(r'\d', tail[2]):
            details['amount'] = tail[2]
        if len(tail) > 3:
            details['address'] = ' | '.join(tail[3:])
    lowered = text.casefold()
    if not details['registration']:
        if 'nhân viên hỗ trợ' in lowered:
            details['registration'] = 'Nhân viên hỗ trợ'
        elif 'kh tự đăng ký' in lowered or 'khách hàng tự đăng ký' in lowered:
            details['registration'] = 'Khách hàng tự đăng ký'
    return details


def _sim_batch_convert_output_row(headers, row):
    source = {
        _batch_header_key(header): row[index] if index < len(row) else ''
        for index, header in enumerate(headers)
    }
    value = lambda name: source.get(_batch_header_key(name), '')
    result = value('Kết quả') or value('Nhóm kết quả')
    parsed = _sim_batch_legacy_details(result)
    registration = value('Loại đăng ký') or parsed['registration']
    registration_key = _batch_header_key(registration)
    if (registration_key.startswith('kh tự đăng ký') or
            registration_key.startswith('khách hàng tự đăng ký')):
        registration = 'Khách hàng tự đăng ký'
    elif 'nhân viên hỗ trợ' in registration_key:
        registration = 'Nhân viên hỗ trợ'
    converted = [
        value('SĐT'),
        value('Serial SIM'),
        result,
        value('Mã đơn hàng') or parsed['order'],
        registration,
        value('Tên khách hàng') or parsed['customer'],
        value('Địa chỉ khách hàng') or value('Địa chỉ') or value('Địa chỉ đầy đủ') or parsed['address'],
        value('Số tiền') or parsed['amount'],
        value('Thời gian'),
        value('User chạy'),
    ]
    # Preserve values from unknown legacy columns inside Kết quả instead of
    # silently dropping history during the one-time schema migration.
    known = {_batch_header_key(name) for name in SIM_BATCH_OUTPUT_HEADERS}
    known.update({_batch_header_key('Nhóm kết quả'), _batch_header_key('Địa chỉ'),
                  _batch_header_key('Địa chỉ đầy đủ')})
    extras = [
        f'{headers[index]}: {item}' for index, item in enumerate(row)
        if index < len(headers) and item not in (None, '') and _batch_header_key(headers[index]) not in known
    ]
    if extras:
        converted[2] = ' | '.join(filter(None, [str(converted[2] or '').strip(), 'Dữ liệu cũ: ' + '; '.join(extras)]))
    return converted


def _prepare_sim_batch_output_workbook(workbook):
    """Migrate the stable cumulative log without deleting any existing rows."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    changed = False
    sheet = workbook['SIM_Output'] if 'SIM_Output' in workbook.sheetnames else workbook.active
    sheet.title = 'SIM_Output'
    current_headers = [cell.value for cell in sheet[1]][:sheet.max_column] if sheet.max_row else []
    if tuple(str(value or '').strip() for value in current_headers) != SIM_BATCH_OUTPUT_HEADERS:
        old_headers = current_headers
        old_rows = list(sheet.iter_rows(min_row=2, values_only=True)) if sheet.max_row > 1 else []
        converted = [_sim_batch_convert_output_row(old_headers, row) for row in old_rows]
        if sheet.max_row:
            sheet.delete_rows(1, sheet.max_row)
        sheet.append(list(SIM_BATCH_OUTPUT_HEADERS))
        for row in converted:
            sheet.append(row)
        changed = True

    widths = [20, 20, 48, 18, 26, 28, 68, 18, 22, 24]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = f'A1:J{max(1, sheet.max_row)}'

    if SIM_BATCH_OUTPUT_META_SHEET not in workbook.sheetnames:
        metadata = workbook.create_sheet(SIM_BATCH_OUTPUT_META_SHEET)
        metadata.append(['entry_id', 'saved_at'])
        metadata.sheet_state = 'hidden'
        changed = True
    else:
        metadata = workbook[SIM_BATCH_OUTPUT_META_SHEET]
        metadata.sheet_state = 'hidden'
    return sheet, metadata, changed


def _migrate_legacy_sim_run_files():
    """Merge old SIM_Kit_Output_* workbooks into the all-time log once."""
    from openpyxl import load_workbook

    if not os.path.isfile(SIM_BATCH_OUTPUT_FILE):
        return 0
    legacy_pattern = re.compile(
        r'^SIM_Kit_Output_\d{8}_\d{6}_[0-9a-fA-F]{6}\.xlsx$')
    legacy_names = sorted(
        name for name in os.listdir(SIM_BATCH_DIR)
        if legacy_pattern.match(name)
    )
    if not legacy_names:
        return 0

    workbook = load_workbook(SIM_BATCH_OUTPUT_FILE)
    sheet, metadata, _ = _prepare_sim_batch_output_workbook(workbook)
    existing_ids = {
        str(cell.value or '').strip()
        for cell in metadata['A'][1:]
        if str(cell.value or '').strip()
    }
    existing_rows = {
        tuple(str(value or '') for value in row)
        for row in sheet.iter_rows(min_row=2, max_col=len(SIM_BATCH_OUTPUT_HEADERS), values_only=True)
    }
    appended = 0
    changed = False
    for name in legacy_names:
        file_marker = f'legacy-file:{name}'
        if file_marker in existing_ids:
            continue
        legacy_path = os.path.join(SIM_BATCH_DIR, name)
        try:
            legacy = load_workbook(legacy_path, data_only=True, read_only=True)
            legacy_sheet = (legacy['SIM_Output']
                            if 'SIM_Output' in legacy.sheetnames else legacy.active)
            rows = legacy_sheet.iter_rows(values_only=True)
            headers = [str(value or '').strip() for value in next(rows, ())]
            for row_number, row in enumerate(rows, start=2):
                entry_id = f'legacy:{name}:{row_number}'
                if entry_id in existing_ids or not any(value not in (None, '') for value in row):
                    continue
                converted = _sim_batch_convert_output_row(headers, row)
                fingerprint = tuple(str(value or '') for value in converted)
                if fingerprint in existing_rows:
                    continue
                sheet.append(converted)
                metadata.append([entry_id, time.strftime('%Y-%m-%d %H:%M:%S')])
                existing_ids.add(entry_id)
                existing_rows.add(fingerprint)
                appended += 1
                changed = True
            legacy.close()
            metadata.append([file_marker, time.strftime('%Y-%m-%d %H:%M:%S')])
            existing_ids.add(file_marker)
            changed = True
        except (OSError, ValueError, KeyError):
            continue
    if changed:
        sheet.auto_filter.ref = f'A1:J{max(1, sheet.max_row)}'
        _atomic_save_workbook(workbook, SIM_BATCH_OUTPUT_FILE)
    else:
        workbook.close()
    return appended


def _prepare_icoc_cumulative_workbook(workbook):
    """Normalize the all-time IC/OC log while preserving legacy rows."""
    from openpyxl.styles import Alignment, Font, PatternFill

    changed = False
    sheet = workbook['IC_OC_Output'] if 'IC_OC_Output' in workbook.sheetnames else workbook.active
    sheet.title = 'IC_OC_Output'
    current_headers = [cell.value for cell in sheet[1]][:sheet.max_column] if sheet.max_row else []
    normalized_headers = tuple(str(value or '').strip() for value in current_headers)
    if normalized_headers != ICOC_BATCH_OUTPUT_HEADERS:
        old_rows = list(sheet.iter_rows(min_row=2, values_only=True)) if sheet.max_row > 1 else []
        header_indexes = {
            _batch_header_key(header): index for index, header in enumerate(current_headers)
        }

        def old_value(row, *names):
            for name in names:
                index = header_indexes.get(_batch_header_key(name))
                if index is not None and index < len(row):
                    return row[index]
            return ''

        converted = [[
            old_value(row, 'SĐT', 'Số điện thoại', 'MSISDN', 'Số TB'),
            old_value(row, 'Kết quả'),
            old_value(row, 'Thời gian'),
            old_value(row, 'User chạy'),
        ] for row in old_rows]
        if sheet.max_row:
            sheet.delete_rows(1, sheet.max_row)
        sheet.append(list(ICOC_BATCH_OUTPUT_HEADERS))
        for row in converted:
            sheet.append(row)
        changed = True

    for cell in sheet[1][:4]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for column, width in {'A': 20, 'B': 80, 'C': 22, 'D': 24}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = f'A1:D{max(1, sheet.max_row)}'

    if ICOC_BATCH_OUTPUT_META_SHEET not in workbook.sheetnames:
        metadata = workbook.create_sheet(ICOC_BATCH_OUTPUT_META_SHEET)
        metadata.append(['entry_id', 'saved_at'])
        changed = True
    else:
        metadata = workbook[ICOC_BATCH_OUTPUT_META_SHEET]
    metadata.sheet_state = 'hidden'
    return sheet, metadata, changed


def _prepare_icoc_run_workbook(workbook, headers):
    """Prepare a per-run IC/OC workbook that keeps the imported columns."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    sheet = workbook['IC_OC_Output'] if 'IC_OC_Output' in workbook.sheetnames else workbook.active
    sheet.title = 'IC_OC_Output'
    if sheet.max_row == 0 or not any(cell.value not in (None, '') for cell in sheet[1]):
        sheet.append(list(headers))
    current_headers = tuple(str(cell.value or '').strip() for cell in sheet[1][:len(headers)])
    if current_headers != tuple(headers):
        raise ValueError('Header file output phiên IC/OC không khớp phiên đang chạy')
    for index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=index)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        sheet.column_dimensions[get_column_letter(index)].width = (
            80 if _batch_header_key(header) == _batch_header_key('Kết quả')
            else (20 if index == 1 else 24)
        )
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = f'A1:{get_column_letter(max(1, len(headers)))}{max(1, sheet.max_row)}'
    if ICOC_BATCH_OUTPUT_META_SHEET not in workbook.sheetnames:
        metadata = workbook.create_sheet(ICOC_BATCH_OUTPUT_META_SHEET)
        metadata.append(['entry_id', 'saved_at'])
    else:
        metadata = workbook[ICOC_BATCH_OUTPUT_META_SHEET]
    metadata.sheet_state = 'hidden'
    return sheet, metadata


def _migrate_legacy_icoc_run_files():
    """Merge old per-run IC_OC_Output_* workbooks into the all-time log once."""
    from openpyxl import load_workbook

    if not os.path.isfile(ICOC_BATCH_OUTPUT_FILE):
        return 0
    legacy_pattern = re.compile(
        r'^IC_OC_Output_\d{8}_\d{6}_[0-9a-fA-F]{6}\.xlsx$')
    legacy_names = sorted(
        name for name in os.listdir(ICOC_BATCH_DIR)
        if legacy_pattern.match(name)
    )
    if not legacy_names:
        return 0

    workbook = load_workbook(ICOC_BATCH_OUTPUT_FILE)
    sheet, metadata, _ = _prepare_icoc_cumulative_workbook(workbook)
    existing_ids = {
        str(cell.value or '').strip()
        for cell in metadata['A'][1:]
        if str(cell.value or '').strip()
    }
    existing_rows = {
        tuple(str(value or '') for value in row)
        for row in sheet.iter_rows(min_row=2, max_col=len(ICOC_BATCH_OUTPUT_HEADERS), values_only=True)
    }
    appended = 0
    changed = False
    for name in legacy_names:
        file_marker = f'legacy-file:{name}'
        if file_marker in existing_ids:
            continue
        legacy_path = os.path.join(ICOC_BATCH_DIR, name)
        try:
            legacy = load_workbook(legacy_path, data_only=True, read_only=True)
            legacy_sheet = (legacy['IC_OC_Output']
                            if 'IC_OC_Output' in legacy.sheetnames else legacy.active)
            rows = legacy_sheet.iter_rows(values_only=True)
            headers = [str(value or '').strip() for value in next(rows, ())]
            phone_index = _icoc_header_index(
                headers, 'SĐT', 'Số điện thoại', 'MSISDN', 'Số TB')
            result_index = _icoc_header_index(headers, 'Kết quả')
            time_index = _icoc_header_index(headers, 'Thời gian')
            user_index = _icoc_header_index(headers, 'User chạy')
            for row_number, row in enumerate(rows, start=2):
                entry_id = f'legacy:{name}:{row_number}'
                if entry_id in existing_ids or not any(value not in (None, '') for value in row):
                    continue
                converted = [
                    row[phone_index] if phone_index is not None and phone_index < len(row) else row[0],
                    row[result_index] if result_index is not None and result_index < len(row) else '',
                    row[time_index] if time_index is not None and time_index < len(row) else '',
                    row[user_index] if user_index is not None and user_index < len(row) else '--',
                ]
                fingerprint = tuple(str(value or '') for value in converted)
                if fingerprint in existing_rows:
                    continue
                sheet.append(converted)
                metadata.append([entry_id, time.strftime('%Y-%m-%d %H:%M:%S')])
                existing_ids.add(entry_id)
                existing_rows.add(fingerprint)
                appended += 1
                changed = True
            legacy.close()
            metadata.append([file_marker, time.strftime('%Y-%m-%d %H:%M:%S')])
            existing_ids.add(file_marker)
            changed = True
        except (OSError, ValueError, KeyError):
            # A damaged/open historical workbook must not block current runs.
            continue
    if changed:
        sheet.auto_filter.ref = f'A1:D{max(1, sheet.max_row)}'
        _atomic_save_workbook(workbook, ICOC_BATCH_OUTPUT_FILE)
    else:
        workbook.close()
    return appended

def _ensure_sim_batch_files():
    """Create the stable input template and cumulative SIM output log."""
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
        sheet.append(['SĐT', 'Serial SIM', 'CCCD'])
        sheet.freeze_panes = 'A2'
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 22
        for column in ('A', 'B', 'C'):
            sheet.column_dimensions[column].number_format = '@'
        guide = workbook.create_sheet('Huong_dan')
        guide.append(['Hướng dẫn'])
        guide.append(['Mỗi dòng trong sheet SIM_Input gồm SĐT thuê bao, Serial SIM trắng và CCCD của khách hàng.'])
        guide.append(['SĐT nhận dạng đầu 84, đầu 0 hoặc 9 chữ số không có tiền tố.'])
        guide.append(['CCCD/số giấy tờ là bắt buộc, gồm 8 đến 20 chữ số và nên lưu ở định dạng Text.'])
        guide.append(['Ví dụ tương đương: 84849531207 / 0849531207 / 849531207'])
        guide.append(['Ví dụ một dòng: 0849531207 | 1184229391 | 079123456789'])
        workbook.save(SIM_BATCH_INPUT_FILE)
    else:
        try:
            workbook = load_workbook(SIM_BATCH_INPUT_FILE)
            sheet = workbook['SIM_Input'] if 'SIM_Input' in workbook.sheetnames else workbook.active
            guide = workbook['Huong_dan'] if 'Huong_dan' in workbook.sheetnames else workbook.create_sheet('Huong_dan')
            instructions = [
                'Hướng dẫn',
                'Mỗi dòng trong sheet SIM_Input gồm SĐT thuê bao, Serial SIM trắng và CCCD của khách hàng.',
                'SĐT nhận dạng đầu 84, đầu 0 hoặc 9 chữ số không có tiền tố.',
                'CCCD/số giấy tờ là bắt buộc, gồm 8 đến 20 chữ số và nên lưu ở định dạng Text.',
                'Ví dụ tương đương: 84849531207 / 0849531207 / 849531207',
                'Ví dụ một dòng: 0849531207 | 1184229391 | 079123456789',
            ]
            changed = False
            input_headers = ('SĐT', 'Serial SIM', 'CCCD')
            for column_index, value in enumerate(input_headers, start=1):
                if sheet.cell(row=1, column=column_index).value != value:
                    sheet.cell(row=1, column=column_index, value=value)
                    changed = True
            for column, width in (('A', 20), ('B', 20), ('C', 22)):
                dimension = sheet.column_dimensions[column]
                if dimension.width != width:
                    dimension.width = width
                    changed = True
                if dimension.number_format != '@':
                    dimension.number_format = '@'
                    changed = True
            if sheet.freeze_panes != 'A2':
                sheet.freeze_panes = 'A2'
                changed = True
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
        sheet.append(list(SIM_BATCH_OUTPUT_HEADERS))
        _prepare_sim_batch_output_workbook(workbook)
        workbook.save(SIM_BATCH_OUTPUT_FILE)
        workbook.close()
    else:
        # One-time migration keeps every prior result while making column E the
        # requested registration type for all future appended rows.
        try:
            workbook = load_workbook(SIM_BATCH_OUTPUT_FILE)
            _, _, changed = _prepare_sim_batch_output_workbook(workbook)
            if changed:
                workbook.save(SIM_BATCH_OUTPUT_FILE)
            workbook.close()
        except PermissionError:
            pass

    try:
        _migrate_legacy_sim_run_files()
    except PermissionError:
        # The cumulative file may be open in Excel during application startup.
        pass


def _ensure_icoc_batch_files():
    """Create the stable input/output workbooks for batch IC/OC changes."""
    from openpyxl import Workbook, load_workbook
    from zipfile import BadZipFile, is_zipfile

    os.makedirs(ICOC_BATCH_DIR, exist_ok=True)

    def quarantine_invalid_workbook(path):
        """Keep the broken file for recovery, then allow a clean replacement."""
        stamp = time.strftime('%Y%m%d-%H%M%S')
        backup = f'{path}.corrupt-{stamp}'
        suffix = 1
        while os.path.exists(backup):
            backup = f'{path}.corrupt-{stamp}-{suffix}'
            suffix += 1
        try:
            shutil.move(path, backup)
            print(f'[IC_OC] Workbook lỗi đã được giữ lại tại: {backup}')
            return True
        except OSError as exc:
            print(f'[IC_OC] Không thể di chuyển workbook lỗi {path}: {exc}')
            return False

    def load_existing_workbook(path):
        """Return None after quarantining any empty/non-OOXML workbook."""
        try:
            valid_archive = (
                os.path.isfile(path)
                and os.path.getsize(path) > 0
                and is_zipfile(path)
            )
            if not valid_archive:
                if not quarantine_invalid_workbook(path):
                    raise PermissionError(path)
                return None
            return load_workbook(path)
        except PermissionError:
            raise
        except (BadZipFile, KeyError, OSError, ValueError) as exc:
            print(f'[IC_OC] Workbook không hợp lệ {path}: {exc}')
            if not quarantine_invalid_workbook(path):
                raise PermissionError(path) from exc
            return None

    def create_input_workbook():
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
        workbook.close()

    def create_output_workbook():
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'IC_OC_Output'
        sheet.append(list(ICOC_BATCH_OUTPUT_HEADERS))
        _prepare_icoc_cumulative_workbook(workbook)
        workbook.save(ICOC_BATCH_OUTPUT_FILE)
        workbook.close()

    if not os.path.exists(ICOC_BATCH_INPUT_FILE):
        create_input_workbook()
    else:
        try:
            workbook = load_existing_workbook(ICOC_BATCH_INPUT_FILE)
            if workbook is None:
                create_input_workbook()
            else:
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
        create_output_workbook()
    else:
        try:
            workbook = load_existing_workbook(ICOC_BATCH_OUTPUT_FILE)
            if workbook is None:
                create_output_workbook()
            else:
                _, _, changed = _prepare_icoc_cumulative_workbook(workbook)
                if changed:
                    workbook.save(ICOC_BATCH_OUTPUT_FILE)
                workbook.close()
        except PermissionError:
            pass
    try:
        _migrate_legacy_icoc_run_files()
    except PermissionError:
        # The cumulative file may be open in Excel during application startup.
        pass


def _prepare_serial_lookup_output_workbook(workbook):
    """Validate and format the cumulative two-column serial lookup output."""
    from openpyxl.styles import Alignment, Font, PatternFill

    sheet = (workbook['Tra_Cuu_Seri'] if 'Tra_Cuu_Seri' in workbook.sheetnames
             else workbook.active)
    sheet.title = 'Tra_Cuu_Seri'
    current_headers = tuple(
        str(sheet.cell(row=1, column=index).value or '').strip()
        for index in range(1, 3)
    )
    if not any(current_headers):
        for index, header in enumerate(SERIAL_LOOKUP_OUTPUT_HEADERS, start=1):
            sheet.cell(row=1, column=index, value=header)
    elif current_headers != SERIAL_LOOKUP_OUTPUT_HEADERS:
        raise ValueError(
            'Tra_Cuu_Seri_Output.xlsx không còn đúng 2 cột '
            'MSIN / Seri SIM và SĐT; hãy đóng/di chuyển file cũ để ứng dụng tạo lại.'
        )

    for cell in sheet[1][:2]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    sheet.column_dimensions['A'].width = 28
    sheet.column_dimensions['B'].width = 18
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = f'A1:B{max(1, sheet.max_row)}'
    for row in sheet.iter_rows(min_row=2, max_col=2):
        for cell in row:
            cell.number_format = '@'

    if SERIAL_LOOKUP_OUTPUT_META_SHEET not in workbook.sheetnames:
        metadata = workbook.create_sheet(SERIAL_LOOKUP_OUTPUT_META_SHEET)
        metadata.append(['entry_id', 'saved_at', 'output_row'])
    else:
        metadata = workbook[SERIAL_LOOKUP_OUTPUT_META_SHEET]
        metadata.cell(row=1, column=1, value='entry_id')
        metadata.cell(row=1, column=2, value='saved_at')
        metadata.cell(row=1, column=3, value='output_row')
    metadata.sheet_state = 'hidden'
    return sheet, metadata


def _ensure_serial_lookup_files():
    """Create the stable one-column input and two-column cumulative output."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    os.makedirs(SERIAL_LOOKUP_BATCH_DIR, exist_ok=True)
    if not os.path.exists(SERIAL_LOOKUP_INPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Input'
        sheet.append(['MSIN / Seri SIM'])
        sheet.freeze_panes = 'A2'
        sheet.column_dimensions['A'].width = 28
        cell = sheet['A1']
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        for row in sheet.iter_rows(min_row=2, max_col=1):
            row[0].number_format = '@'
        workbook.save(SERIAL_LOOKUP_INPUT_FILE)
        workbook.close()

    if not os.path.exists(SERIAL_LOOKUP_OUTPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Tra_Cuu_Seri'
        sheet.append(list(SERIAL_LOOKUP_OUTPUT_HEADERS))
        _prepare_serial_lookup_output_workbook(workbook)
        workbook.save(SERIAL_LOOKUP_OUTPUT_FILE)
        workbook.close()
    else:
        workbook = load_workbook(SERIAL_LOOKUP_OUTPUT_FILE)
        _prepare_serial_lookup_output_workbook(workbook)
        workbook.close()


def _prepare_package_registration_workbook(workbook):
    """Validate and format a package-registration audit workbook."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    sheet = (workbook['Dang_Ky_Goi_Cuoc']
             if 'Dang_Ky_Goi_Cuoc' in workbook.sheetnames else workbook.active)
    sheet.title = 'Dang_Ky_Goi_Cuoc'
    header_count = len(PACKAGE_REGISTRATION_HEADERS)
    current_headers = tuple(
        str(sheet.cell(row=1, column=index).value or '').strip()
        for index in range(1, header_count + 1)
    )
    if not any(current_headers):
        for index, header in enumerate(PACKAGE_REGISTRATION_HEADERS, start=1):
            sheet.cell(row=1, column=index, value=header)
    elif current_headers != PACKAGE_REGISTRATION_HEADERS:
        raise ValueError(
            'File output Đăng ký gói cước đã bị đổi cấu trúc; '
            'hãy đóng hoặc đổi tên file để ứng dụng tạo lại.'
        )

    for cell in sheet[1][:header_count]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    widths = (7, 18, 24, 14, 16, 16, 14, 62, 21, 24)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = (
        f'A1:{get_column_letter(header_count)}{max(1, sheet.max_row)}'
    )
    for row in sheet.iter_rows(min_row=2, max_col=header_count):
        row[1].number_format = '@'

    if PACKAGE_REGISTRATION_META_SHEET not in workbook.sheetnames:
        metadata = workbook.create_sheet(PACKAGE_REGISTRATION_META_SHEET)
        metadata.append(['entry_id', 'saved_at'])
    else:
        metadata = workbook[PACKAGE_REGISTRATION_META_SHEET]
        metadata.cell(row=1, column=1, value='entry_id')
        metadata.cell(row=1, column=2, value='saved_at')
    metadata.sheet_state = 'hidden'
    return sheet, metadata


def _ensure_package_registration_files():
    """Create the cumulative output used by menu 11077."""
    from openpyxl import Workbook, load_workbook

    os.makedirs(PACKAGE_REGISTRATION_DIR, exist_ok=True)
    if not os.path.exists(PACKAGE_REGISTRATION_OUTPUT_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Dang_Ky_Goi_Cuoc'
        sheet.append(list(PACKAGE_REGISTRATION_HEADERS))
        _prepare_package_registration_workbook(workbook)
        workbook.save(PACKAGE_REGISTRATION_OUTPUT_FILE)
        workbook.close()
        return
    workbook = load_workbook(PACKAGE_REGISTRATION_OUTPUT_FILE)
    _prepare_package_registration_workbook(workbook)
    workbook.close()


def _append_package_registration_workbook(
        path, rows, entry_ids, authenticated_users, fallback_user):
    """Append completed rows idempotently and replace the workbook atomically."""
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    sheet, metadata = _prepare_package_registration_workbook(workbook)
    existing_ids = {
        str(cell.value or '').strip()
        for cell in metadata['A'][1:]
        if str(cell.value or '').strip()
    }
    appended = 0
    skipped = 0
    expected = len(PACKAGE_REGISTRATION_HEADERS)
    for index, incoming in enumerate(rows):
        entry_id = str(
            entry_ids[index] if index < len(entry_ids) else ''
        ).strip()[:200] or secrets.token_hex(16)
        if entry_id in existing_ids:
            skipped += 1
            continue
        mapped = (list(incoming) + [''] * expected)[:expected]
        requested_user = str(mapped[9] or '').strip().casefold()
        mapped[9] = authenticated_users.get(requested_user, fallback_user)
        mapped[1] = str(mapped[1] or '').strip()
        sheet.append(mapped)
        sheet.cell(row=sheet.max_row, column=2).number_format = '@'
        metadata.append([entry_id, time.strftime('%Y-%m-%d %H:%M:%S')])
        existing_ids.add(entry_id)
        appended += 1
    _prepare_package_registration_workbook(workbook)
    total_rows = max(0, sheet.max_row - 1)
    _atomic_save_workbook(workbook, path)
    return appended, skipped, total_rows

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


def _owned_batch_run_path(requested, directory, prefix):
    """Validate a per-run workbook path created by this application."""
    if not requested:
        return ''
    candidate = os.path.abspath(str(requested).strip())
    root = os.path.abspath(directory)
    name = os.path.basename(candidate)
    try:
        if (os.path.commonpath([candidate, root]) != root or
                not candidate.lower().endswith('.xlsx') or
                not name.startswith(f'{prefix}_')):
            return ''
    except ValueError:
        return ''
    return candidate


def _atomic_save_workbook(workbook, path):
    """Replace one workbook atomically so an interrupted save keeps the old file."""
    temp_path = f'{path}.{secrets.token_hex(4)}.tmp.xlsx'
    try:
        workbook.save(temp_path)
        workbook.close()
        os.replace(temp_path, path)
    finally:
        try:
            workbook.close()
        except Exception:
            pass
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except OSError:
            pass


def _append_sim_output_workbook(path, incoming_headers, incoming_rows,
                                entry_ids, authenticated_users, fallback_user):
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    sheet, metadata, _ = _prepare_sim_batch_output_workbook(workbook)
    existing_ids = {
        str(cell.value or '').strip()
        for cell in metadata['A'][1:]
        if str(cell.value or '').strip()
    }
    appended = 0
    skipped = 0
    for index, incoming in enumerate(incoming_rows):
        supplied_id = entry_ids[index] if index < len(entry_ids) else ''
        entry_id = str(supplied_id or secrets.token_hex(16)).strip()[:200]
        if entry_id in existing_ids:
            skipped += 1
            continue
        mapped = _sim_batch_convert_output_row(incoming_headers, incoming)
        requested_user = str(mapped[9] or '').strip().casefold()
        mapped[9] = authenticated_users.get(requested_user, fallback_user)
        sheet.append(mapped)
        metadata.append([entry_id, time.strftime('%Y-%m-%d %H:%M:%S')])
        existing_ids.add(entry_id)
        appended += 1
    sheet.auto_filter.ref = f'A1:J{max(1, sheet.max_row)}'
    total_rows = max(0, sheet.max_row - 1)
    _atomic_save_workbook(workbook, path)
    return appended, skipped, total_rows


def _icoc_header_index(headers, *names):
    indexes = {_batch_header_key(header): index for index, header in enumerate(headers)}
    for name in names:
        index = indexes.get(_batch_header_key(name))
        if index is not None:
            return index
    return None


def _append_icoc_output_workbook(path, incoming_headers, incoming_rows,
                                 entry_ids, authenticated_users, fallback_user,
                                 cumulative=False):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    workbook = load_workbook(path)
    if cumulative:
        sheet, metadata, _ = _prepare_icoc_cumulative_workbook(workbook)
    else:
        sheet, metadata = _prepare_icoc_run_workbook(workbook, incoming_headers)
    existing_ids = {
        str(cell.value or '').strip()
        for cell in metadata['A'][1:]
        if str(cell.value or '').strip()
    }
    phone_index = _icoc_header_index(
        incoming_headers, 'SĐT', 'Số điện thoại', 'MSISDN', 'Số TB')
    result_index = _icoc_header_index(incoming_headers, 'Kết quả')
    time_index = _icoc_header_index(incoming_headers, 'Thời gian')
    user_index = _icoc_header_index(incoming_headers, 'User chạy')
    appended = 0
    skipped = 0
    for index, incoming in enumerate(incoming_rows):
        supplied_id = entry_ids[index] if index < len(entry_ids) else ''
        entry_id = str(supplied_id or secrets.token_hex(16)).strip()[:200]
        if entry_id in existing_ids:
            skipped += 1
            continue
        requested_user = (
            str(incoming[user_index] or '').strip().casefold()
            if user_index is not None and user_index < len(incoming) else '')
        safe_user = authenticated_users.get(requested_user, fallback_user)
        if cumulative:
            mapped = [
                incoming[phone_index] if phone_index is not None and phone_index < len(incoming) else incoming[0],
                incoming[result_index] if result_index is not None and result_index < len(incoming) else '',
                incoming[time_index] if time_index is not None and time_index < len(incoming) else '',
                safe_user,
            ]
        else:
            mapped = (list(incoming) + [''] * len(incoming_headers))[:len(incoming_headers)]
            if user_index is not None:
                mapped[user_index] = safe_user
        sheet.append(mapped)
        metadata.append([entry_id, time.strftime('%Y-%m-%d %H:%M:%S')])
        existing_ids.add(entry_id)
        appended += 1
    sheet.auto_filter.ref = (
        f'A1:D{max(1, sheet.max_row)}' if cumulative else
        f'A1:{get_column_letter(max(1, len(incoming_headers)))}{max(1, sheet.max_row)}'
    )
    total_rows = max(0, sheet.max_row - 1)
    _atomic_save_workbook(workbook, path)
    return appended, skipped, total_rows


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
    _ensure_serial_lookup_files()
    _ensure_package_registration_files()

def _account_id(username):
    return hashlib.sha256(username.strip().casefold().encode('utf-8')).hexdigest()[:24]

def _normalise_account_phone(value):
    """Return a plausible phone value without exposing arbitrary profile data."""
    if value is None or isinstance(value, (dict, list, tuple, set)):
        return ''
    text = str(value).strip()
    digits = re.sub(r'\D', '', text)
    if 9 <= len(digits) <= 15:
        return text
    return ''


_ACCOUNT_PHONE_KEYS = {
    'phone', 'phonenumber', 'mobile', 'mobilephone', 'telephone',
    'sdt', 'sodt', 'sodienthoai', 'dienthoai', 'dienthoainguoidung',
    'so_dt', 'so_dien_thoai', 'dien_thoai',
}


def _extract_account_phone(value):
    """Extract the employee phone from the varying OneBSS profile shapes."""
    if isinstance(value, dict):
        for key, item in value.items():
            normalised_key = re.sub(r'[^a-z0-9_]', '', str(key).casefold())
            if normalised_key in _ACCOUNT_PHONE_KEYS:
                phone = _normalise_account_phone(item)
                if phone:
                    return phone
        for item in value.values():
            phone = _extract_account_phone(item)
            if phone:
                return phone
    elif isinstance(value, (list, tuple)):
        for item in value:
            phone = _extract_account_phone(item)
            if phone:
                return phone
    return ''


def _extract_account_phone_from_token(token):
    """Use a phone claim when the access token already contains one."""
    payload = _decode_access_token_payload(token)
    return _extract_account_phone(payload)


def _decode_access_token_payload(token):
    try:
        payload_part = str(token or '').split('.')[1]
        payload_part += '=' * (-len(payload_part) % 4)
        return json.loads(base64.urlsafe_b64decode(payload_part).decode('utf-8'))
    except (IndexError, ValueError, TypeError, UnicodeDecodeError,
            json.JSONDecodeError):
        return {}


def _normalise_pasted_token(token):
    """Accept a raw token or the common ``Bearer <token>`` clipboard form."""
    value = str(token or '').strip()
    if value[:7].casefold() == 'bearer ':
        value = value[7:].strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        value = value[1:-1].strip()
    # Chat/Markdown and wrapped text can add escapes or line breaks which are
    # not part of a base64url JWT. Normalising them makes clipboard login match
    # the token that OneBSS originally issued.
    value = value.replace('\\_', '_').replace('\\-', '-')
    return ''.join(value.split())


_ACCOUNT_USERNAME_KEYS = {
    'username', 'preferredusername', 'uniquename', 'useraccount',
    'account', 'accountname', 'login', 'loginname', 'tendangnhap',
    'taikhoan', 'manv', 'manhanvien', 'employeecode', 'staffcode',
}


def _normalise_account_username(value):
    if value is None or isinstance(value, (dict, list, tuple, set)):
        return ''
    username = str(value).strip()
    if not username or len(username) > 160 or any(char.isspace() for char in username):
        return ''
    return username


def _extract_account_username(value):
    """Extract a OneBSS login name from a token/profile of varying shape."""
    if isinstance(value, dict):
        for key, item in value.items():
            normalised_key = re.sub(r'[^a-z0-9]', '', str(key).casefold())
            if normalised_key in _ACCOUNT_USERNAME_KEYS:
                username = _normalise_account_username(item)
                if username:
                    return username
        for item in value.values():
            username = _extract_account_username(item)
            if username:
                return username
    elif isinstance(value, (list, tuple)):
        for item in value:
            username = _extract_account_username(item)
            if username:
                return username
    return ''


def _extract_account_username_from_token(token):
    claims = _decode_access_token_payload(token)
    username = _extract_account_username(claims)
    if username:
        return username
    # Some identity providers only expose the login in ``sub``. Avoid using a
    # numeric/internal subject id as p_account for OneBSS business requests.
    subject = _normalise_account_username(claims.get('sub'))
    return subject if subject and not subject.isdigit() else ''


def _extract_token_device_id(token):
    claims = _decode_access_token_payload(token)
    for key in ('id_thietbi', 'device_id', 'deviceId'):
        value = claims.get(key)
        if value is not None and not isinstance(value, (dict, list, tuple, set)):
            value = str(value).strip()
            if value:
                return value
    return ''


def _is_plausible_signed_access_token(token):
    """Allow offline login only for a signed-looking, unexpired identity JWT.

    Signature verification remains the responsibility of OneBSS on every API
    request. This fallback is needed because thongtin_nv also validates device
    metadata, so it can reject a valid token pasted on another app instance.
    """
    parts = str(token or '').split('.')
    if len(parts) != 3 or not all(parts):
        return False
    try:
        encoded_header = parts[0] + '=' * (-len(parts[0]) % 4)
        header = json.loads(base64.urlsafe_b64decode(encoded_header).decode('utf-8'))
        algorithm = str(header.get('alg') or '').strip().casefold()
        claims = _decode_access_token_payload(token)
        expires_at = float(claims.get('exp'))
    except (ValueError, TypeError, UnicodeDecodeError, json.JSONDecodeError):
        return False
    return bool(
        algorithm and algorithm != 'none' and
        expires_at > time.time() and
        _extract_account_username_from_token(token)
    )


def _token_seconds_remaining(token):
    claims = _decode_access_token_payload(token)
    try:
        expires_at = float(claims.get('exp'))
    except (TypeError, ValueError):
        return 3600
    return max(0, int(expires_at - time.time()))


def _account_phone(context):
    return (_normalise_account_phone(context.get('phone')) or
            _extract_account_phone_from_token(context.get('access_token')))


def _fetch_account_phone(context):
    """Fetch the signed-in employee profile from OneBSS."""
    cached = _account_phone(context)
    if cached:
        return cached
    active_menu = session.get('active_menu_id', APP_CFG['SELECTED_MENU'])
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'authorization': f"Bearer {context.get('access_token', '')}",
        'app-secret': _build_app_secret_value(
            context.get('app_secret', ''), context.get('device_id', '')),
        'selectedmenuid': active_menu,
        'SelectedMenuId': active_menu,
    }
    try:
        response = requests.post(
            f"{BASE_URL}/quantri/user/thongtin_nv",
            headers=headers, json={}, verify=False, timeout=8)
        if 200 <= response.status_code < 300:
            return _extract_account_phone(response.json())
    except (requests.RequestException, ValueError, TypeError):
        pass
    return ''


def _request_account_profile(context):
    """Validate an account token and return the OneBSS employee profile."""
    active_menu = session.get('active_menu_id', APP_CFG['SELECTED_MENU'])
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'authorization': f"Bearer {context.get('access_token', '')}",
        'app-secret': _build_app_secret_value(
            context.get('app_secret', ''), context.get('device_id', '')),
        'selectedmenuid': active_menu,
        'SelectedMenuId': active_menu,
    }
    try:
        response = requests.post(
            f"{BASE_URL}/quantri/user/thongtin_nv",
            headers=headers, json={}, verify=False, timeout=8)
        try:
            body = response.json()
        except (ValueError, TypeError):
            body = None
        return response, body
    except requests.RequestException:
        return None, None


def _remember_account_phone(account_id, phone):
    """Cache a profile phone in the active session and encrypted account list."""
    phone = _normalise_account_phone(phone)
    account_id = str(account_id or '').strip()
    if not phone or not account_id:
        return
    primary_id = _account_id(str(session.get('username') or '').strip())
    if account_id == primary_id:
        session['account_phone'] = phone
    else:
        extras = session.get('multi_accounts') or {}
        if isinstance(extras, dict) and isinstance(extras.get(account_id), dict):
            extras = dict(extras)
            context = dict(extras[account_id])
            context['phone'] = phone
            extras[account_id] = context
            session['multi_accounts'] = extras
    with _saved_accounts_lock:
        saved = _read_saved_accounts()
        changed = False
        for account in saved:
            if str(account.get('id') or '') == account_id:
                if account.get('phone') != phone:
                    account['phone'] = phone
                    changed = True
                break
        if changed:
            _write_saved_accounts(saved)
    _save_persistent_session()


def save_employee_account(username, password, phone=''):
    username = str(username or '').strip()
    password = str(password or '')
    if not username or not password:
        return False
    account_id = _account_id(username)
    with _saved_accounts_lock:
        all_accounts = _read_saved_accounts()
        existing = next((account for account in all_accounts if account.get('id') == account_id), None)
        accounts = [a for a in all_accounts if a.get('id') != account_id]
        account_data = {
            'id': account_id,
            'username': username,
            'password': password,
            'saved_at': int(time.time())
        }
        if phone or (existing and existing.get('phone')):
            account_data['phone'] = _normalise_account_phone(phone) or existing.get('phone', '')
        accounts.insert(0, account_data)
        _write_saved_accounts(accounts)
    return True

def saved_account_summaries():
    return [{'id': a['id'], 'username': a['username'],
             **({'phone': a['phone']} if a.get('phone') else {})}
            for a in _read_saved_accounts()]

def get_saved_employee_account(account_id):
    return next((a for a in _read_saved_accounts()
                 if a.get('id') == account_id), None)


def delete_saved_employee_account(account_id):
    """Delete one encrypted saved credential without touching other accounts."""
    account_id = str(account_id or '').strip()
    if not account_id:
        return False
    with _saved_accounts_lock:
        accounts = _read_saved_accounts()
        remaining = [account for account in accounts
                     if str(account.get('id') or '') != account_id]
        if len(remaining) == len(accounts):
            return False
        _write_saved_accounts(remaining)
    return True

# ─── CONFIG (Dynamic) ───────────────────────
BASE_URL = 'https://api-onebss.vnpt.vn'
DEFAULT_APP_VERSION = '1.5.41.130'
APP_SETTINGS_FILE = os.path.join(_credential_root, 'app_settings.json')
_app_settings_lock = threading.RLock()


def _normalize_app_version(value):
    """Accept the four-part OneBSS app version shown on the download page."""
    version = str(value or '').strip()
    if not re.fullmatch(r'\d{1,4}(?:\.\d{1,4}){3}', version):
        raise ValueError('APP_VERSION không hợp lệ. Ví dụ: 1.5.41.130')
    return version


def _load_saved_app_version():
    try:
        with _app_settings_lock:
            with open(APP_SETTINGS_FILE, 'r', encoding='utf-8') as settings_file:
                settings = json.load(settings_file)
        saved = _normalize_app_version(settings.get('app_version'))
        # OneBSS currently requires at least 1.5.41.130. An older
        # persisted version produces an app-secret that the current gateway
        # rejects with 401, so upgrade old settings while preserving any newer
        # version the operator configured later.
        saved_parts = tuple(int(part) for part in saved.split('.'))
        minimum_parts = tuple(int(part) for part in DEFAULT_APP_VERSION.split('.'))
        return saved if saved_parts >= minimum_parts else DEFAULT_APP_VERSION
    except (OSError, TypeError, ValueError, json.JSONDecodeError):
        return DEFAULT_APP_VERSION


def _save_app_version(version):
    """Persist APP_VERSION outside the EXE so upgrades keep the user's value."""
    version = _normalize_app_version(version)
    os.makedirs(_credential_root, exist_ok=True)
    temp_file = f'{APP_SETTINGS_FILE}.{os.getpid()}.tmp'
    payload = {'app_version': version, 'updated_at': int(time.time())}
    with _app_settings_lock:
        try:
            with open(temp_file, 'w', encoding='utf-8') as settings_file:
                json.dump(payload, settings_file, ensure_ascii=False, indent=2)
            os.replace(temp_file, APP_SETTINGS_FILE)
        finally:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except OSError:
                pass
    return version


APP_CFG = {
    'CLIENT_ID': 'clientapp',
    'CLIENT_SECRET': 'password',
    'MENU_ID': 810241,
    'SELECTED_MENU': '810241',
    'APP_VERSION': _load_saved_app_version()
}

if EKYC_AVAILABLE and hasattr(_ekyc, 'set_app_version'):
    _ekyc.set_app_version(APP_CFG['APP_VERSION'])

# IDG Token-id / Token-key (từ upload_mobile.js)
IDG_TOKEN_ID  = '04c0a953-7fb8-5461-e063-62199f0aeda6'
IDG_TOKEN_KEY = 'MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAKjy7FK9SegSCW0cuUIbEDUsbRZOCoxijNPLMfvgX+8/XA7HebHXMN4/PO5c5mwK3lYk3lRKuMXYLLp6X6oZPDkCAwEAAQ=='

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
            'account_phone': session.get('account_phone', ''),
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
        session['account_phone']  = data.get('account_phone', '')
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
        'phone': session.get('account_phone', ''),
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
    ('/quantri/user/log_sudung_chucnang', '11077'),
    ('/ccbs/oneBss/app_tb_tc_thongtin', '11077'),
    ('/ccbs/goicuoc/', '11077'),
    ('/ccbs/pttb/get_sotb_by_msin', '699060'),
    ('/ccbs/tracuu/ts_tracuu_stb_serial', '699060'),
    # Customer self-registration (p_id_hinhthuc_dk_tttb=2) is exposed by the
    # newer kenhban-simkit module.  The mobile capture uses menu 810641 for
    # these shared chonSo calls, while legacy SIM-order screens still use
    # menu 699161 for the broader /ccbs/chonSo namespace.
    ('/ccbs/chonSo/app_ds_dauso', '810641'),
    ('/ccbs/chonSo/checkSimStatus', '810641'),
    ('/app-banhang/kenhban-simkit/', '810641'),
    ('/web-quantri/danhmuc-chung/lay_tt_ts_diadanh_moi', '810641'),
    ('/app-ccdv/vietqr/check_donhang', '810641'),
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
    # Device-change identity verification uses the eKYC utility menu captured
    # from Employee mobile.  Keep these before broader /app-banhang prefixes.
    ('/app-com/Config/token_ekyc', '810241'),
    ('/app-com/Config/app_config', '810241'),
    ('/quantri/user/get_ekyc_config', '810241'),
    ('/app-banhang/Ekyc/init_log_uuid', '810241'),
    ('/app-banhang/Ekyc/log_ekyc', '810241'),
    ('/app-banhang/quanlyfile/', '810241'),
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
        if permission_code == 'BANGOICUOCDIDONG':
            return '11077'
        if permission_code == 'KHOITAOTB':
            return '810641'
    # get_danhmuc is shared by many modules.  Preserve the captured SIM Kit
    # menu only for this known flow instead of forcing the legacy 699161 menu.
    if (endpoint_path == '/app-com/danhmuc/get_danhmuc' and
            isinstance(body, dict) and
            str(body.get('menu_id') or '').strip() == '810641'):
        return '810641'
    # VNPT Pay is also shared by legacy and customer-self-registration SIM
    # flows.  The completed mobile capture keeps menu 810641 for wallet auth,
    # balance and token refresh in this branch.
    if (endpoint_path.startswith('/app-thuno/VnptPay/') and
            isinstance(body, dict) and
            str(body.get('menu_id') or '').strip() == '810641'):
        return '810641'
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
@login_required
def update_settings():
    data = request.get_json(silent=True) or {}
    APP_CFG['CLIENT_ID'] = data.get('client_id', APP_CFG['CLIENT_ID'])
    APP_CFG['CLIENT_SECRET'] = data.get('client_secret', APP_CFG['CLIENT_SECRET'])
    APP_CFG['SELECTED_MENU'] = data.get('selected_menu', APP_CFG['SELECTED_MENU'])
    APP_CFG['MENU_ID'] = int(APP_CFG['SELECTED_MENU']) if str(APP_CFG['SELECTED_MENU']).isdigit() else APP_CFG['MENU_ID']
    try:
        requested_version = _save_app_version(
            data.get('app_version', APP_CFG['APP_VERSION']))
    except (OSError, ValueError) as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 400
    APP_CFG['APP_VERSION'] = requested_version
    if EKYC_AVAILABLE and hasattr(_ekyc, 'set_app_version'):
        _ekyc.set_app_version(requested_version)
    
    # Đồng bộ session active_menu_id nếu người dùng lưu cài đặt
    session['active_menu_id'] = APP_CFG['SELECTED_MENU']
    current_app_secret()
    _save_persistent_session()
    
    return jsonify({
        'ok': True,
        'msg': f'Đã lưu APP_VERSION {requested_version}',
        'app_version': requested_version,
    })

def _render_login():
    return render_template(
        'login.html',
        saved_accounts=saved_account_summaries(),
        app_version=APP_CFG['APP_VERSION'])


def _validated_pasted_token_context(token, expected_username='', device_id='', app_secret='', primary=False):
    """Validate a pasted token and build an account context without mutating session."""
    token = _normalise_pasted_token(token)
    expected_username = str(expected_username or '').strip()
    if not token:
        return None, 'Hãy dán access token.', 400

    remaining = _token_seconds_remaining(token)
    if remaining <= 0:
        return None, 'Token đã hết hạn.', 401

    # Prefer the device id embedded in a pasted JWT over stale OTP/session
    # metadata. OneBSS may bind thongtin_nv to the device that obtained token.
    device_id = str(
        device_id or _extract_token_device_id(token) or
        session.get('device_id') or secrets.token_hex(8))
    app_secret = _build_app_secret_value(
        app_secret or session.get('app_secret', ''), device_id)
    context = {
        'username': expected_username,
        'access_token': token,
        'refresh_token': '',
        'expires_in': remaining,
        'token_time': time.time(),
        'device_id': device_id,
        'app_secret': app_secret,
        'primary': bool(primary),
    }
    response, profile = _request_account_profile(context)
    profile_succeeded = bool(
        response is not None and profile is not None and
        _business_payload_succeeded(response.status_code, profile))
    if not profile_succeeded:
        # A newly-issued JWT can still be rejected by thongtin_nv when the
        # pasted session lacks the original app-secret. Keep login usable and
        # let OneBSS validate the signature on the first real business call.
        if not _is_plausible_signed_access_token(token):
            if response is None:
                return None, 'Không kết nối được OneBSS để kiểm tra token.', 502
            return None, 'Token không hợp lệ hoặc đã hết hạn.', 401
        profile = {}

    detected_username = (
        _extract_account_username(profile) or
        _extract_account_username_from_token(token)
    )
    if (expected_username and detected_username and
            expected_username.casefold() != detected_username.casefold()):
        return None, 'Token không khớp user được chọn.', 400
    username = detected_username or expected_username
    if not username:
        return None, 'Token hợp lệ nhưng OneBSS không trả tên user.', 400

    phone = _extract_account_phone(profile) or _extract_account_phone_from_token(token)
    context['id'] = _account_id(username)
    context['username'] = username
    if phone:
        context['phone'] = phone
    return context, '', 200


def _activate_primary_token_session(token, expected_username='', device_id='', app_secret=''):
    """Validate a pasted token and replace the current primary login session."""
    context, error, status = _validated_pasted_token_context(
        token,
        expected_username=expected_username,
        device_id=device_id,
        app_secret=app_secret,
        primary=True,
    )
    if not context:
        return None, error, status
    token = context['access_token']
    username = context['username']
    phone = context.get('phone', '')
    session.clear()
    session['username'] = username
    session['access_token'] = token
    session['refresh_token'] = ''
    session['expires_in'] = context.get('expires_in', 3600)
    session['token_time'] = context['token_time']
    session['device_id'] = context.get('device_id', '')
    session['app_secret'] = context.get('app_secret', '')
    session['menus'] = []
    session['active_menu_id'] = str(APP_CFG['MENU_ID'])
    session['multi_accounts'] = {}
    session['account_phone'] = phone
    account_id = _account_id(username)
    if phone:
        _remember_account_phone(account_id, phone)
    else:
        _save_persistent_session()
    return _primary_account_context(), '', 200

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


@app.post('/login/token')
def login_token():
    context, error, _status = _activate_primary_token_session(
        request.form.get('token', ''))
    if not context:
        flash(error, 'error')
        return redirect(url_for('login'))
    flash(f"Đăng nhập {context.get('username', '')} bằng token thành công!", 'success')
    return redirect(url_for('dashboard'))

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
    summary = {
        'id': context.get('id', ''),
        'username': context.get('username', ''),
        'primary': bool(context.get('primary')),
        'expires_in': _account_seconds_remaining(context),
        'authenticated': bool(context.get('access_token')),
    }
    phone = _account_phone(context)
    if phone:
        summary['phone'] = phone
    token = str(context.get('access_token') or '')
    if token:
        summary['access_token'] = token
    return summary


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


@app.post('/api/accounts/profile')
@login_required
def api_accounts_profile():
    """Return the employee phone for one active account."""
    account_id = str((request.get_json(silent=True) or {}).get('account_id') or '').strip()
    try:
        context = _get_account_context(account_id)
    except ValueError as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 404
    phone = _fetch_account_phone(context)
    if phone:
        _remember_account_phone(context.get('id') or account_id, phone)
    return jsonify({'ok': True, 'account_id': context.get('id') or account_id,
                    'phone': phone})


@app.post('/api/accounts/token-login')
@login_required
def api_accounts_token_login():
    """Add or update one independent account from a pasted access token."""
    payload = request.get_json(silent=True) or {}
    account_id = str(payload.get('account_id') or '').strip()
    token = _normalise_pasted_token(
        payload.get('token') or payload.get('access_token'))
    if not token:
        return jsonify({'ok': False, 'error': 'Hãy dán access token'}), 400

    primary = _primary_account_context()
    primary_id = primary.get('id')
    extras = session.get('multi_accounts') or {}
    existing = extras.get(account_id) if isinstance(extras, dict) else None
    saved = get_saved_employee_account(account_id) if account_id else None
    username = str(payload.get('username') or
                   (existing or {}).get('username') or
                   (saved or {}).get('username') or
                   (primary.get('username') if account_id == primary_id else '') or '').strip()

    base_context = dict(existing or {})
    if account_id == primary_id:
        base_context = primary
    context, error, status = _validated_pasted_token_context(
        token,
        expected_username=username,
        device_id=base_context.get('device_id') or session.get('device_id', ''),
        app_secret=base_context.get('app_secret') or session.get('app_secret', ''),
        primary=False,
    )
    if not context:
        return jsonify({'ok': False, 'error': error}), status
    detected_id = str(context.get('id') or _account_id(context.get('username', '')))
    if account_id and account_id != detected_id:
        return jsonify({'ok': False, 'error': 'Token không khớp user được chọn'}), 400

    if detected_id == primary_id:
        for key in ('username', 'access_token', 'refresh_token', 'expires_in',
                    'token_time', 'device_id', 'app_secret'):
            if key in context:
                session[key] = context[key]
        if context.get('phone'):
            session['account_phone'] = context['phone']
        context['primary'] = True
    else:
        with _multi_account_lock:
            extras = session.get('multi_accounts') or {}
            extras = dict(extras) if isinstance(extras, dict) else {}
            if len(extras) >= 10 and detected_id not in extras:
                return jsonify({
                    'ok': False,
                    'error': 'Chỉ cho phép tối đa 10 tài khoản trong một phiên',
                }), 400
            stored = dict(context)
            stored.pop('primary', None)
            extras[detected_id] = stored
            session['multi_accounts'] = extras
        context['primary'] = False
    if context.get('phone'):
        _remember_account_phone(detected_id, context['phone'])
    else:
        _save_persistent_session()
    return jsonify({'ok': True, 'account': _public_account_summary(context),
                    'message': f"Đã thêm {context.get('username', username)} bằng token"})


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


@app.post('/api/accounts/delete')
@login_required
def api_accounts_delete():
    """Delete a saved account and sign out its active token, if present."""
    account_id = str((request.get_json(silent=True) or {}).get('account_id') or '').strip()
    if not account_id:
        return jsonify({'ok': False, 'error': 'Thiếu tài khoản cần xóa'}), 400

    primary_id = _primary_account_context().get('id')
    deleted_saved = delete_saved_employee_account(account_id)
    if account_id == primary_id:
        session.clear()
        _clear_persistent_session()
        return jsonify({
            'ok': True,
            'deleted': True,
            'deleted_saved': deleted_saved,
            'logout_required': True,
        })

    with _multi_account_lock:
        extras = session.get('multi_accounts') or {}
        extras = dict(extras) if isinstance(extras, dict) else {}
        removed_session = extras.pop(account_id, None) is not None
        session['multi_accounts'] = extras
    _save_persistent_session()
    return jsonify({
        'ok': True,
        'deleted': bool(deleted_saved or removed_session),
        'deleted_saved': deleted_saved,
        'removed_session': removed_session,
        'logout_required': False,
    })


@app.post('/api/accounts/refresh')
@login_required
def api_accounts_refresh():
    account_id = str((request.get_json(silent=True) or {}).get('account_id') or '').strip()
    try:
        context = _get_account_context(account_id)
    except ValueError as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 404
    remaining = _account_seconds_remaining(context)
    if remaining > 120:
        return jsonify({
            'ok': True,
            'skipped': True,
            'message': f'Token vẫn còn {remaining} giây; chưa cần làm mới',
            'account': _public_account_summary(context),
        })
    if not context.get('refresh_token'):
        return jsonify({'ok': False, 'error': 'Tài khoản không có refresh token; hãy thêm lại và nhập OTP'}), 400
    try:
        account_key = str(context.get('id') or account_id or 'primary')
        with _business_request_slot(account_key, {
                'workflow': 'auth', 'mutation': False,
                'cooldown': BUSINESS_SIM_COOLDOWN_SECONDS}):
            response = requests.post(
                f'{BASE_URL}/quantri/oauth/token',
                json={'grant_type': 'refresh_token', 'refresh_token': context.get('refresh_token'),
                      'client_id': APP_CFG['CLIENT_ID'], 'client_secret': APP_CFG['CLIENT_SECRET']},
                headers={'Content-Type': 'application/json'}, verify=False, timeout=15)
            retry_after = _business_record_response(account_key, response)
        data = response.json()
    except BusinessGuardRejected as exc:
        return jsonify({'ok': False, 'error': str(exc),
                        'retry_after': exc.retry_after}), exc.status
    except Exception as exc:
        return jsonify({'ok': False, 'error': f'Lỗi làm mới token: {exc}'}), 502
    if response.status_code == 429:
        return jsonify({'ok': False,
                        'error': data.get('message') or 'OneBSS đang giới hạn tần suất',
                        'retry_after': retry_after}), 429
    if not data.get('access_token'):
        status = response.status_code if response.status_code >= 400 else 400
        return jsonify({'ok': False,
                        'error': data.get('message') or 'Không làm mới được token'}), status
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
        context = _primary_account_context()
        account_key = str(context.get('id') or 'primary')
        with _business_request_slot(account_key, {
                'workflow': 'auth', 'mutation': False,
                'cooldown': BUSINESS_SIM_COOLDOWN_SECONDS}):
            resp = requests.post(
                f'{BASE_URL}/quantri/oauth/token',
                json={'grant_type':    'refresh_token',
                      'refresh_token': session.get('refresh_token', ''),
                      'client_id':     APP_CFG['CLIENT_ID'],
                      'client_secret': APP_CFG['CLIENT_SECRET']},
                headers={'Content-Type': 'application/json'},
                verify=False, timeout=10
            )
            retry_after = _business_record_response(account_key, resp)
        data = resp.json()
        if resp.status_code == 429:
            return jsonify({'ok': False,
                            'error': data.get('message') or 'OneBSS đang giới hạn tần suất',
                            'retry_after': retry_after}), 429
        if data.get('access_token'):
            session['access_token']  = data['access_token']
            session['refresh_token'] = data.get('refresh_token', session['refresh_token'])
            session['token_time']    = time.time()
            return jsonify({'ok': True, 'expires_in': data.get('expires_in', 3600)})
    except BusinessGuardRejected as exc:
        return jsonify({'ok': False, 'error': str(exc),
                        'retry_after': exc.retry_after}), exc.status
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    status = resp.status_code if resp.status_code >= 400 else 400
    return jsonify({'ok': False,
                    'error': data.get('message') or 'Failed'}), status

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


@app.route('/api/onebss-report/status')
@login_required
def api_onebss_report_status():
    """Return progress without exposing the saved OneBSS access token."""
    return jsonify({'ok': True, **_onebss_report_manager.status()})


@app.route('/api/onebss-report/start', methods=['POST'])
@login_required
def api_onebss_report_start():
    """Start/resume the daily OneBSS export from April 2026 to today."""
    from datetime import date, datetime

    payload = request.get_json(silent=True) or {}
    start_text = str(payload.get('start_date') or '2026-04-01').strip()
    end_text = str(payload.get('end_date') or date.today().isoformat()).strip()
    account_id = str(payload.get('account_id') or '').strip()
    try:
        start_day = datetime.strptime(start_text, '%Y-%m-%d').date()
        end_day = datetime.strptime(end_text, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({
            'ok': False,
            'error': 'Ngày phải đúng định dạng YYYY-MM-DD.',
        }), 400
    if end_day > date.today():
        return jsonify({
            'ok': False,
            'error': 'Đến ngày không được lớn hơn ngày hiện tại.',
        }), 400
    try:
        access_token = ''
        account_username = ''
        if account_id:
            context = _get_account_context(account_id)
            access_token = str(context.get('access_token') or '')
            account_username = str(context.get('username') or '')
        state = _onebss_report_manager.start(
            start_day,
            end_day,
            access_token=access_token,
            account_username=account_username,
        )
        return jsonify({'ok': True, **state})
    except (OneBssReportError, ValueError) as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 409


@app.route('/api/onebss-report/stop', methods=['POST'])
@login_required
def api_onebss_report_stop():
    return jsonify({'ok': True, **_onebss_report_manager.stop()})


@app.route('/api/onebss-report/download')
@login_required
def api_onebss_report_download():
    state = _onebss_report_manager.status()
    output_path = os.path.abspath(str(state.get('output_path') or ''))
    try:
        owned = os.path.commonpath([
            output_path,
            os.path.abspath(ONEBSS_REPORT_OUTPUT_DIR),
        ]) == os.path.abspath(ONEBSS_REPORT_OUTPUT_DIR)
    except ValueError:
        owned = False
    if not owned or not os.path.isfile(output_path):
        return jsonify({'ok': False, 'error': 'Chưa có file Excel để tải.'}), 404
    return send_file(
        output_path,
        as_attachment=True,
        download_name=os.path.basename(output_path),
    )


# ─────────────────────────────────────────────────────────────
#  OneBSS business guard – lookup, SIM Kit and IC/OC
# ─────────────────────────────────────────────────────────────
# Browser-side flags are not authoritative: two WebViews or a direct /proxy
# caller can otherwise overlap requests for the same Employee account.  Keep
# the safety boundary here, immediately before outbound OneBSS traffic.
BUSINESS_SIM_COOLDOWN_SECONDS = max(
    0.5, float(os.environ.get('VNPT_EMPLOY_SIM_COOLDOWN_SECONDS', '1.5')))
BUSINESS_ICOC_COOLDOWN_SECONDS = max(
    1.0, float(os.environ.get('VNPT_EMPLOY_ICOC_COOLDOWN_SECONDS', '2.0')))
BUSINESS_LOOKUP_COOLDOWN_SECONDS = max(
    1.0, float(os.environ.get('VNPT_EMPLOY_LOOKUP_COOLDOWN_SECONDS', '1.5')))
BUSINESS_MUTATION_CACHE_SECONDS = 24 * 3600
BUSINESS_GUARD_DB = os.path.join(_credential_root, 'business_guard.sqlite3')

_business_guard_lock = threading.RLock()
_business_account_locks = {}
_business_account_completed_at = {}
_business_account_paused_until = {}
_business_db_lock = threading.RLock()

_SIM_MUTATION_ENDPOINTS = frozenset({
    '/app-banhang/donhang_simkit/chonso_kit_v2',
    '/app-banhang/donhang_simkit/huy_donhang',
    '/app-banhang/donhang_simkit/dangky_goicuoc',
    '/app-banhang/donhang_simkit/nhap_thongtin_khachhang_v3',
    '/app-banhang/donhang_simkit/xacnhan_thanhtoan',
    '/app-banhang/donhang_simkit/khoitao_thuebao',
    '/app-banhang/kenhban-simkit/chonso_kit_v2',
    '/app-banhang/kenhban-simkit/dangky_goicuoc',
    '/app-banhang/kenhban-simkit/nhap_thongtin_khachhang_v3',
    '/app-banhang/kenhban-simkit/khoitao_thuebao',
    '/app-banhang/kenhban-simkit/xacnhan_thanhtoan',
    '/app-banhang/kenhban-simkit/hoanthanh_donhang_tratruoc',
})
_ICOC_MUTATION_ENDPOINTS = frozenset({
    '/app-banhang/thuebaodidong/khoamo_ic_oc',
})
_PACKAGE_REGISTRATION_MUTATION_ENDPOINTS = frozenset({
    '/ccbs/goicuoc/dangky',
})
_LOOKUP_ENDPOINTS = frozenset({
    '/ccbs/pttb/get_sotb_by_msin',
    '/ccbs/tracuu/ts_tracuu_stb_serial',
    '/app-banhang/ccbs/tracuu_anh_thuebao',
    '/app-banhang/ccbs/tracuu_thongtin_thuebao',
    '/app-banhang/ccbs/verify_otp',
    '/app-banhang/thuebaodidong/lichsu_thuebao',
})
_LOOKUP_OTP_ENDPOINT = '/app-banhang/ccbs/send_otp'


class BusinessGuardRejected(RuntimeError):
    def __init__(self, message, status=429, retry_after=0):
        super().__init__(message)
        self.status = int(status)
        self.retry_after = max(0, int(retry_after or 0))


def _business_endpoint_policy(endpoint_path, body):
    """Return guard metadata for lookup, SIM Kit and IC/OC workflows."""
    path = str(endpoint_path or '').rstrip('/').casefold()
    payload = body if isinstance(body, dict) else {}
    menu_id = str(payload.get('menu_id') or '').strip()

    is_sim = (
        path.startswith('/ccbs/chonso/') or
        path.startswith('/app-banhang/donhang_simkit/') or
        path.startswith('/app-banhang/kenhban-simkit/') or
        path.startswith('/app-thuno/vnptpay/') or
        (path == '/app-com/danhmuc/get_danhmuc' and
         menu_id in ('699161', '810641')) or
        path == '/web-quantri/danhmuc-chung/lay_tt_ts_diadanh_moi' or
        (path == '/app-ccdv/vietqr/check_donhang' and
         menu_id == '810641') or
        (path == '/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen' and
         str(payload.get('ma_quyen') or '').strip().upper() == 'KHOITAOTB')
    )
    if is_sim:
        return {
            'workflow': 'sim',
            'mutation': path in _SIM_MUTATION_ENDPOINTS,
            'cooldown': BUSINESS_SIM_COOLDOWN_SECONDS,
        }

    is_icoc = (
        path == '/app-banhang/thuebaodidong/tracuu_tb_didong' or
        path == '/app-banhang/thuebaodidong/khoamo_ic_oc' or
        (path == '/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen' and
         str(payload.get('ma_quyen') or '').strip().upper() == 'CATMODICHVU')
    )
    if is_icoc:
        return {
            'workflow': 'icoc',
            'mutation': path in _ICOC_MUTATION_ENDPOINTS,
            'cooldown': BUSINESS_ICOC_COOLDOWN_SECONDS,
        }

    is_package_registration = (
        path == '/quantri/user/log_sudung_chucnang' or
        path == '/ccbs/onebss/app_tb_tc_thongtin' or
        path.startswith('/ccbs/goicuoc/') or
        (path == '/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen' and
         str(payload.get('ma_quyen') or '').strip().upper() ==
         'BANGOICUOCDIDONG')
    )
    if is_package_registration:
        return {
            'workflow': 'package_registration',
            'mutation': path in _PACKAGE_REGISTRATION_MUTATION_ENDPOINTS,
            'cooldown': BUSINESS_LOOKUP_COOLDOWN_SECONDS,
        }
    if path == _LOOKUP_OTP_ENDPOINT:
        return {
            'workflow': 'lookup',
            'mutation': False,
            'cooldown': 30.0,
        }
    if path in _LOOKUP_ENDPOINTS:
        return {
            'workflow': 'lookup',
            'mutation': False,
            'cooldown': BUSINESS_LOOKUP_COOLDOWN_SECONDS,
        }
    return None


def _business_account_lock(account_key):
    with _business_guard_lock:
        return _business_account_locks.setdefault(account_key, threading.Lock())


def _retry_after_seconds(response, default=30):
    value = str(response.headers.get('Retry-After') or '').strip()
    try:
        return max(1, min(300, int(float(value))))
    except (TypeError, ValueError):
        return max(1, min(300, int(default)))


@contextmanager
def _business_request_slot(account_key, policy):
    """Serialize requests per account; different accounts run independently."""
    account_lock = _business_account_lock(account_key)
    account_lock.acquire()
    slot_entered = False
    try:
        now = time.monotonic()
        with _business_guard_lock:
            paused_until = _business_account_paused_until.get(account_key, 0)
            completed_at = _business_account_completed_at.get(account_key, 0)
        if paused_until > now:
            retry_after = max(1, int(paused_until - now + 0.999))
            raise BusinessGuardRejected(
                'Tài khoản đang tạm dừng do OneBSS giới hạn tần suất',
                retry_after=retry_after)

        remaining = completed_at + float(policy['cooldown']) - now
        if remaining > 0:
            time.sleep(remaining)
        slot_entered = True
        yield
    finally:
        if slot_entered:
            with _business_guard_lock:
                _business_account_completed_at[account_key] = time.monotonic()
        account_lock.release()


def _business_record_response(account_key, response):
    if response.status_code == 429:
        retry_after = _retry_after_seconds(response)
        with _business_guard_lock:
            _business_account_paused_until[account_key] = max(
                _business_account_paused_until.get(account_key, 0),
                time.monotonic() + retry_after)
        return retry_after
    if response.status_code == 401:
        with _business_guard_lock:
            _business_account_paused_until[account_key] = max(
                _business_account_paused_until.get(account_key, 0),
                time.monotonic() + 60)
    return 0


def _ensure_business_guard_db():
    with _business_db_lock:
        os.makedirs(_credential_root, exist_ok=True)
        connection = sqlite3.connect(BUSINESS_GUARD_DB, timeout=10)
        try:
            connection.execute('PRAGMA journal_mode=WAL')
            connection.execute('''
                CREATE TABLE IF NOT EXISTS mutation_guard (
                    operation_key TEXT PRIMARY KEY,
                    account_id TEXT NOT NULL,
                    endpoint TEXT NOT NULL,
                    state TEXT NOT NULL,
                    response_json TEXT,
                    updated_at REAL NOT NULL
                )
            ''')
            connection.execute(
                'DELETE FROM mutation_guard WHERE updated_at < ?',
                (time.time() - 7 * 24 * 3600,))
            connection.commit()
        finally:
            connection.close()


def _business_operation_key(account_key, endpoint_path, body):
    path = str(endpoint_path or '').rstrip('/').casefold()
    payload = body if isinstance(body, dict) else {}
    identity = {'account': account_key, 'endpoint': path}
    if path.endswith('/chonso_kit_v2'):
        identity['phone'] = payload.get('p_so_dt')
    elif path.endswith('/huy_donhang'):
        identity['order'] = payload.get('p_id_donhang')
    elif path.endswith('/dangky_goicuoc'):
        identity['order'] = payload.get('p_id_donhang')
        identity['recharge'] = payload.get('p_id_hinhthuc_napthe')
    elif path.endswith('/nhap_thongtin_khachhang_v3'):
        identity['order'] = payload.get('p_id_donhang')
    elif path.endswith('/xacnhan_thanhtoan'):
        identity['order'] = payload.get('p_id_donhang')
    elif path.endswith('/hoanthanh_donhang_tratruoc'):
        identity['order'] = payload.get('p_id_donhang')
    elif path.endswith('/khoitao_thuebao'):
        identity['order'] = payload.get('p_id_donhang')
    elif path.endswith('/khoamo_ic_oc'):
        identity.update({
            'phone': payload.get('p_so_tb'),
            'ic': payload.get('p_goi_den'),
            'oc': payload.get('p_goi_di'),
        })
    else:
        identity['body'] = payload
    encoded = json.dumps(identity, sort_keys=True, separators=(',', ':'),
                         ensure_ascii=False).encode('utf-8')
    return hashlib.sha256(encoded).hexdigest()


def _business_claim_mutation(operation_key, account_key, endpoint_path):
    """Claim a mutation or return a recent completed response for deduping."""
    _ensure_business_guard_db()
    now = time.time()
    with _business_db_lock:
        connection = sqlite3.connect(BUSINESS_GUARD_DB, timeout=10)
        try:
            connection.execute('BEGIN IMMEDIATE')
            row = connection.execute(
                'SELECT state, response_json, updated_at FROM mutation_guard WHERE operation_key=?',
                (operation_key,)).fetchone()
            if row:
                state, response_json, updated_at = row
                age = max(0, now - float(updated_at or 0))
                is_icoc_change = str(endpoint_path).casefold().endswith(
                    '/khoamo_ic_oc')
                completed_ttl = (60 if is_icoc_change
                                 else BUSINESS_MUTATION_CACHE_SECONDS)
                if state == 'completed' and age < completed_ttl:
                    connection.commit()
                    try:
                        return 'completed', json.loads(response_json or '{}')
                    except (TypeError, ValueError, json.JSONDecodeError):
                        return 'blocked', None
                uncertain_ttl = (60 if is_icoc_change
                                 else BUSINESS_MUTATION_CACHE_SECONDS)
                if state in ('in_flight', 'uncertain') and age < uncertain_ttl:
                    connection.commit()
                    return state, None
                connection.execute(
                    'UPDATE mutation_guard SET state=?, response_json=NULL, updated_at=? WHERE operation_key=?',
                    ('in_flight', now, operation_key))
            else:
                connection.execute(
                    'INSERT INTO mutation_guard(operation_key, account_id, endpoint, state, response_json, updated_at) VALUES(?,?,?,?,?,?)',
                    (operation_key, account_key, endpoint_path, 'in_flight', None, now))
            connection.commit()
            return 'claimed', None
        finally:
            connection.close()


def _business_finish_mutation(operation_key, state, response_payload=None):
    if not operation_key:
        return
    _ensure_business_guard_db()
    encoded = (json.dumps(response_payload, ensure_ascii=False)
               if response_payload is not None else None)
    with _business_db_lock:
        connection = sqlite3.connect(BUSINESS_GUARD_DB, timeout=10)
        try:
            if state == 'failed':
                connection.execute(
                    'DELETE FROM mutation_guard WHERE operation_key=?',
                    (operation_key,))
            else:
                connection.execute(
                    'UPDATE mutation_guard SET state=?, response_json=?, updated_at=? WHERE operation_key=?',
                    (state, encoded, time.time(), operation_key))
            connection.commit()
        finally:
            connection.close()


def _business_payload_succeeded(status_code, payload):
    if not 200 <= int(status_code) < 300:
        return False
    if payload in (None, ''):
        return True
    if not isinstance(payload, dict):
        return not re.search(r'\b(error|fail|exception|no_data)\b',
                             str(payload), flags=re.IGNORECASE)
    if payload.get('success') is False or payload.get('ok') is False:
        return False
    error_code = str(payload.get('error_code') or
                     payload.get('errorCode') or '').strip().upper()
    if error_code and error_code not in {
            '0', '200', 'SUCCESS', 'BSS-00000000'}:
        return False
    error = payload.get('error')
    if (error not in (None, '', False, 0, '0', 200, '200') and
            str(error).strip().upper() not in {'SUCCESS', 'BSS-00000000'}):
        return False
    return True

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
    try:
        timeout = max(3, min(60, int(data.get('timeout', 15))))
    except (TypeError, ValueError):
        timeout = 15

    if not endpoint:
        return jsonify({'error': 'endpoint required'}), 400

    url = endpoint if endpoint.startswith('http') else \
          BASE_URL.rstrip('/') + '/' + endpoint.lstrip('/')

    target_url = urlparse(url)
    base_target = urlparse(BASE_URL)
    try:
        target_port = target_url.port
    except ValueError:
        target_port = -1
    if (target_url.scheme != 'https' or
            target_url.hostname != base_target.hostname or
            target_port not in (None, 443) or
            target_url.username is not None or target_url.password is not None):
        return jsonify({'error': 'proxy host not allowed', 'status': 403}), 403
    endpoint_path = str(target_url.path or '/')
    for _ in range(3):
        decoded_path = unquote(endpoint_path)
        if decoded_path == endpoint_path:
            break
        endpoint_path = decoded_path
    if any(part in ('.', '..') for part in endpoint_path.split('/')):
        return jsonify({'error': 'proxy path not allowed', 'status': 403}), 403
    endpoint_path = re.sub(r'/+', '/', endpoint_path).rstrip('/') or '/'
    try:
        account_context = _get_account_context(account_id)
    except ValueError as exc:
        return jsonify({'error': str(exc), 'status': 401}), 401
    body = normalize_onebss_body(
        endpoint_path, body, endpoint, account_context.get('username'))
    requested_mid = extra_hdr.get('SelectedMenuId') or extra_hdr.get('selectedmenuid')
    active_mid = endpoint_menu_id(endpoint_path, requested_mid, body)
    # A few read-only endpoints are captured without a JSON DTO. Keep their
    # query/body exact; SelectedMenuId is still routed independently above.
    strict_body_endpoints = {
        '/ccbs/chonSo/app_ds_dauso',
        # Mobile sends only ?so_msin=... for this final read-only status check.
        '/ccbs/chonSo/checkSimStatus',
        # Tra SIM ra số sends only ?msin=...; do not add menu_id to the query.
        '/ccbs/pttb/get_sotb_by_msin',
        # The Employee serial lookup DTO contains only so_sim.
        '/ccbs/tracuu/ts_tracuu_stb_serial',
        # Captured as a body-less GET; menu 810641 is carried in headers.
        '/web-quantri/danhmuc-chung/lay_tt_ts_diadanh_moi',
        # The captured payment DTO does not contain menu_id.  Its menu is
        # carried only by SelectedMenuId, like the Employee mobile request.
        '/app-banhang/kenhban-simkit/xacnhan_thanhtoan',
    }
    inject_menu_id = endpoint_path not in strict_body_endpoints
    if inject_menu_id and isinstance(body, dict) and 'menu_id' not in body:
        body['menu_id'] = active_mid

    policy = _business_endpoint_policy(endpoint_path, body)
    if policy and policy.get('mutation') and method != 'POST':
        return jsonify({'error': 'Business mutation requires POST',
                        'status': 405}), 405
    account_key = str(account_context.get('id') or account_id or
                      _account_id(account_context.get('username', '')) or 'primary')
    operation_key = ''
    if policy and policy.get('mutation'):
        operation_key = _business_operation_key(
            account_key, endpoint_path, body)
        claim_state, cached_response = _business_claim_mutation(
            operation_key, account_key, endpoint_path)
        if claim_state == 'completed' and isinstance(cached_response, dict):
            cached_response = dict(cached_response)
            cached_response['deduplicated'] = True
            return jsonify(cached_response)
        if claim_state == 'in_flight':
            return jsonify({
                'status': 409,
                'error': 'Thao tác giống hệt đang được xử lý; không gửi trùng lên OneBSS',
            }), 409
        if claim_state in ('uncertain', 'blocked'):
            return jsonify({
                'status': 409,
                'error': ('Kết quả thao tác trước chưa xác định; cần đối soát trạng thái '
                          'trước khi gửi lại'),
            }), 409

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
        slot = (_business_request_slot(account_key, policy)
                if policy else nullcontext())
        with slot:
            if method == 'GET':
                resp = requests.get(url, params=body, **kw)
                # The current gateway exposes a number of read-only routes as
                # POST although the mobile catalog still labels them GET.
                if (resp.status_code == 405 and
                        endpoint_path.startswith(('/app-', '/ccbs/')) and
                        not (policy and policy.get('mutation'))):
                    resp = requests.post(url, json=body, **kw)
                    method_used = 'POST'
                    retried_from_get = True
            elif body_type == 'form':
                hdrs['Content-Type'] = 'application/x-www-form-urlencoded'
                resp = requests.request(method, url, data=body, **kw)
            else:
                resp = requests.request(method, url, json=body, **kw)
            if (method == 'POST' and resp.status_code == 405 and
                    endpoint_path in READ_ONLY_ENDPOINTS and
                    not (policy and policy.get('mutation'))):
                resp = requests.get(url, params=body, **kw)
                method_used = 'GET'
                retried_from_post = True
            if policy:
                retry_after = _business_record_response(account_key, resp)
            else:
                retry_after = 0
        elapsed = round((time.time()-t0)*1000)
        try: rb = resp.json()
        except: rb = resp.text
        result = {'status': resp.status_code, 'elapsed': elapsed,
                  'method_used': method_used,
                  'retried_from_get': retried_from_get,
                  'retried_from_post': retried_from_post,
                  'selected_menu_id': active_mid,
                  'headers': dict(resp.headers), 'body': rb}
        if retry_after:
            result['retry_after'] = retry_after
        if operation_key:
            payload_succeeded = _business_payload_succeeded(
                resp.status_code, rb)
            if payload_succeeded:
                _business_finish_mutation(operation_key, 'completed', result)
            elif resp.status_code == 408 or resp.status_code >= 500:
                _business_finish_mutation(operation_key, 'uncertain')
            else:
                _business_finish_mutation(operation_key, 'failed')
        return jsonify(result)
    except BusinessGuardRejected as exc:
        if operation_key:
            _business_finish_mutation(operation_key, 'failed')
        return jsonify({'status': exc.status, 'error': str(exc),
                        'retry_after': exc.retry_after}), exc.status
    except requests.exceptions.ConnectionError as e:
        if operation_key:
            _business_finish_mutation(operation_key, 'uncertain')
        return jsonify({'error': f'Lỗi kết nối: {e}'}), 502
    except requests.exceptions.Timeout:
        if operation_key:
            _business_finish_mutation(operation_key, 'uncertain')
        return jsonify({'error': 'Request timeout'}), 504
    except Exception as e:
        if operation_key:
            _business_finish_mutation(operation_key, 'uncertain')
        return jsonify({'error': str(e)}), 500


# ─────────────────────────────────────────────────────────────
#  Device-change identity verification (OneBSS / IDG eKYC)
# ─────────────────────────────────────────────────────────────
DEVICE_AUTH_MENU_ID = '810241'
DEVICE_AUTH_HANDLE_TTL_SECONDS = 5 * 60
DEVICE_AUTH_IDG_BASE = 'https://api.idg.vnpt.vn'
DEVICE_AUTH_CHALLENGE_FALLBACK = 'JGI7TCLnPYhjehlzNp34vSpfANyKRAL4'
DEVICE_AUTH_NEAR_HASH = "zone4/idg20260708-0ced7972-9864-4a32-e063-62199f0ad57f/IDG01_a4fd5ce0-7a86-11f1-8182-fd7dbf4502cd"
DEVICE_AUTH_FAR_HASH  = "zone2/idg20260708-0ced7972-9864-4a32-e063-62199f0ad57f/IDG01_a51662cf-7a86-11f1-af90-5fbeee1966b6"
DEVICE_AUTH_AI_TOKEN  = "8928skjhfa89298jahga1771vbvb"


def _device_auth_client_session(context):
    device_id = str(context.get('device_id') or '279313db43343efc').strip()
    return f'ANDROID_CPH2179_32_Device_3.6.6_{device_id}_{int(time.time() * 1000)}_vn.vnptit.oneapp'

class DeviceAuthError(RuntimeError):
    def __init__(self, message, status=400, *, liveness_passed=False,
                 upstream=None, failed_step=None, server_responses=None,
                 phone=None):
        super().__init__(message)
        self.status = int(status)
        self.liveness_passed = bool(liveness_passed)
        self.upstream = upstream if isinstance(upstream, dict) else {}
        self.failed_step = str(failed_step or '').strip() or None
        self.server_responses = (
            dict(server_responses) if isinstance(server_responses, dict) else {})
        self.phone = str(phone or '').strip()


def _device_auth_add_error_context(exc, *, failed_step, phone,
                                   server_responses):
    """Attach completed upstream calls without altering their response bodies."""
    if not isinstance(exc, DeviceAuthError):
        return exc
    if not exc.failed_step:
        exc.failed_step = failed_step
    if not exc.phone:
        exc.phone = phone
    completed = {
        key: value for key, value in (server_responses or {}).items()
        if value not in (None, {}, [])
    }
    if exc.upstream:
        completed['upstream_error'] = exc.upstream
    exc.server_responses = completed
    return exc


def _device_auth_normalize_phone(value):
    """Normalize one Vietnamese mobile number to the OneBSS 84xxxxxxxxx form."""
    digits = re.sub(r'\D', '', str(value or ''))
    if digits.startswith('0') and len(digits) == 10:
        digits = '84' + digits[1:]
    elif len(digits) == 9:
        digits = '84' + digits
    if not re.fullmatch(r'84[35789]\d{8}', digits):
        raise DeviceAuthError(
            'Số điện thoại không hợp lệ. Nhập dạng 0xxxxxxxxx, 84xxxxxxxxx hoặc 9 số cuối.',
            422)
    return digits


def _onebss_payload_succeeded(status_code, payload):
    if not 200 <= int(status_code) < 300 or not isinstance(payload, dict):
        return False
    error_code = payload.get('error_code')
    if error_code not in (None, '', 'BSS-00000000'):
        return False
    error = payload.get('error')
    return error in (None, '', 0, '0', 200, '200')


def _device_auth_upstream_message(payload, fallback):
    if isinstance(payload, dict):
        return str(payload.get('message') or payload.get('message_detail') or
                   payload.get('error_code') or fallback)
    return fallback


def _device_auth_onebss_post(path, body, account_id=''):
    headers = get_headers(DEVICE_AUTH_MENU_ID, account_id)
    try:
        response = requests.post(
            f"{BASE_URL.rstrip('/')}/{path.lstrip('/')}",
            headers=headers, json=body, verify=False, timeout=25)
    except requests.exceptions.Timeout as exc:
        raise DeviceAuthError(f'API OneBSS timeout tại {path}', 504) from exc
    except requests.exceptions.RequestException as exc:
        raise DeviceAuthError(f'Không kết nối được OneBSS tại {path}', 502) from exc
    try:
        payload = response.json()
    except ValueError as exc:
        raise DeviceAuthError(
            f'OneBSS trả dữ liệu không hợp lệ tại {path}', 502) from exc
    if not _onebss_payload_succeeded(response.status_code, payload):
        message = _device_auth_upstream_message(
            payload, f'OneBSS từ chối tại {path} (HTTP {response.status_code})')
        raise DeviceAuthError(
            message,
            response.status_code if response.status_code >= 400 else 400,
            upstream=payload)
    return payload


def _device_auth_selected_policy(config_payload):
    configs = config_payload.get('data') if isinstance(config_payload, dict) else None
    configs = configs if isinstance(configs, list) else []
    selected = next((item for item in configs if isinstance(item, dict) and
                     str(item.get('dichvu')) == '-1'), None)
    if selected is None:
        selected = next((item for item in configs if isinstance(item, dict)), {})
    if not selected or int(selected.get('ai_must') or 0) != 1 or int(selected.get('check_liveness') or 0) != 1:
        raise DeviceAuthError('Cấu hình OneBSS không bật xác thực sống bắt buộc', 409)
    return dict(selected)


def _device_auth_transactions():
    value = session.get('device_auth_transactions')
    return dict(value) if isinstance(value, dict) else {}


def _device_auth_save_transaction(handle, transaction):
    transactions = _device_auth_transactions()
    now = time.time()
    transactions = {
        key: value for key, value in transactions.items()
        if isinstance(value, dict) and now - float(value.get('created_at') or 0) <=
        DEVICE_AUTH_HANDLE_TTL_SECONDS
    }
    transactions[handle] = transaction
    transactions = dict(sorted(
        transactions.items(), key=lambda item: item[1].get('created_at', 0),
        reverse=True)[:4])
    session['device_auth_transactions'] = transactions
    session.modified = True


def _device_auth_pop_transaction(handle, only_if_success=False):
    transactions = _device_auth_transactions()
    if handle in transactions and not only_if_success:
        transactions.pop(handle, None)
        session['device_auth_transactions'] = transactions
        session.modified = True


def _device_auth_result_object(payload):
    if not isinstance(payload, dict):
        return {}
    for key in ('object', 'data', 'result'):
        value = payload.get(key)
        if isinstance(value, dict):
            return value
    return payload


def _device_auth_true(value):
    if value is True:
        return True
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) >= 0.8
    return str(value or '').strip().casefold() in {
        '1', 'true', 'yes', 'y', 'open', 'opened', 'live', 'real',
        'success', 'pass', 'passed', 'người thật', 'nguoi that'
    }


def _device_auth_false(value):
    if value is False:
        return True
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) == 0
    return str(value or '').strip().casefold() in {
        '0', 'false', 'no', 'n', 'close', 'closed', 'unmasked',
        'not_masked', 'không', 'khong'
    }


def _device_auth_status_values(payload):
    """Return only OneBSS device-verification status fields."""
    status_keys = {
        'trang_thai', 'trangthai', 'ma_trang_thai', 'matrangthai',
    }
    values = []

    def visit(value):
        if isinstance(value, dict):
            for key, item in value.items():
                normalized_key = re.sub(r'[^a-z0-9_]', '', str(key).casefold())
                if normalized_key in status_keys and not isinstance(item, (dict, list)):
                    values.append(item)
                elif isinstance(item, (dict, list)):
                    visit(item)
        elif isinstance(value, list):
            for item in value:
                visit(item)

    visit(payload)
    return values


def _device_auth_has_status_code(payload, expected):
    expected_text = str(expected).strip()
    return any(str(value).strip() == expected_text
               for value in _device_auth_status_values(payload))


def _device_auth_matching_status_value(payload, expected):
    """Return the exact server value that matched a business status code."""
    expected_text = str(expected).strip()
    return next((value for value in _device_auth_status_values(payload)
                 if str(value).strip() == expected_text), None)


def _device_auth_face_result(payload):
    """Normalize the documented OneBSS face result; 661 means face matched."""
    if not isinstance(payload, dict):
        return None
    data = payload.get('data') if isinstance(payload.get('data'), dict) else {}
    message = str(data.get('message') or payload.get('message') or '').casefold()
    is_match = data.get('is_match')

    # OneBSS business status 661 is the definitive "matched" result.
    if _device_auth_has_status_code(payload, 661):
        return True
    if str(is_match).strip() == '1':
        return True
    if str(is_match).strip() == '0' or 'không khớp' in message:
        return False
    return None


def _device_auth_validate_liveness(payload, policy):
    result = _device_auth_result_object(payload)
    live_value = result.get('liveness')
    if not _device_auth_true(live_value):
        detail = result.get('liveness_msg') or 'Không đạt xác thực sống'
        raise DeviceAuthError(str(detail), 422, upstream=payload)
    if _device_auth_true(result.get('fake_liveness')):
        raise DeviceAuthError(
            'Dịch vụ phát hiện giả mạo liveness', 422,
            upstream=payload)
    if _device_auth_true(result.get('face_swapping')):
        raise DeviceAuthError(
            'Dịch vụ phát hiện hoán đổi khuôn mặt', 422,
            upstream=payload)
    if int(policy.get('check_eye_open') or 0) == 1 and not _device_auth_true(result.get('is_eye_open')):
        raise DeviceAuthError(
            'Không xác nhận được mắt đang mở', 422,
            upstream=payload)
    return result


def _device_auth_validate_mask(payload):
    result = _device_auth_result_object(payload)
    masked = result.get('masked')
    if not _device_auth_false(masked):
        raise DeviceAuthError(
            'Khuôn mặt đang bị che hoặc không xác định được khẩu trang',
            422, upstream=payload)
    return result


def _device_auth_idg_headers(transaction):
    headers = {
        'Authorization': transaction['ekyc_access_token'],
        'Token-id': transaction['token_id'],
        'Token-key': transaction['token_key'],
        'User-Agent': 'okhttp/4.11.0',
    }
    device_id = str(transaction.get('device_id') or '').strip()
    if device_id:
        headers['mac-address'] = device_id
    return headers


def _device_auth_upload_to_onebss(frame_bytes, account_id):
    try:
        upload_link = _device_auth_onebss_post(
            '/app-banhang/quanlyfile/get_upload_link', {
                'p_module': 'CCBS',
                'p_file_name': 'PORTRAIT_IMAGE.jpg',
                'menu_id': int(DEVICE_AUTH_MENU_ID),
            }, account_id)
    except DeviceAuthError as exc:
        raise DeviceAuthError(
            str(exc), exc.status, liveness_passed=True,
            upstream=exc.upstream) from exc
    data = upload_link.get('data') if isinstance(upload_link, dict) else None
    if not isinstance(data, dict):
        raise DeviceAuthError('OneBSS không trả thông tin upload ảnh', 502,
                              liveness_passed=True)
    upload_url = str(data.get('url') or '').strip()
    if not upload_url.lower().startswith(('http://', 'https://')):
        upload_url = 'https://' + upload_url.lstrip('/')
    target = urlparse(upload_url)
    if target.scheme != 'https' or not target.hostname or not target.hostname.endswith('vnpt.vn'):
        raise DeviceAuthError('URL lưu ảnh OneBSS không hợp lệ', 502,
                              liveness_passed=True)
    object_name = str(data.get('objectName') or '').strip()
    fields = data.get('fields') if isinstance(data.get('fields'), dict) else {}
    form = {}
    if object_name:
        form['key'] = object_name
    form['Content-Type'] = 'image/jpeg'
    for key, value in fields.items():
        if key not in form:
            form[str(key)] = str(value)
    try:
        stored = requests.post(
            upload_url, data=form,
            files={'file': ('PORTRAIT_IMAGE.jpg', frame_bytes, 'image/jpeg')},
            verify=False,
            timeout=35)
    except requests.exceptions.RequestException as exc:
        raise DeviceAuthError('Không tải được ảnh lên kho OneBSS', 502,
                              liveness_passed=True) from exc
    if stored.status_code not in (200, 201, 204):
        raise DeviceAuthError(
            f'Kho OneBSS từ chối ảnh (HTTP {stored.status_code})', 502,
            liveness_passed=True)
    try:
        updated = _device_auth_onebss_post(
            '/app-banhang/quanlyfile/update_file', {
                'p_object_name': object_name,
                'menu_id': int(DEVICE_AUTH_MENU_ID),
            }, account_id)
    except DeviceAuthError as exc:
        raise DeviceAuthError(
            str(exc), exc.status, liveness_passed=True,
            upstream=exc.upstream) from exc
    return updated


def _device_auth_extract_image_bytes(item, account_id):
    """Trích xuất bytes ảnh từ 1 item OneBSS tracuu_anh_thuebao (base64 hoặc URL)."""
    if not isinstance(item, dict):
        return b''
    b64_data = item.get('image_base') or item.get('base64') or ''
    if isinstance(b64_data, str) and b64_data:
        if ',' in b64_data and b64_data.lower().startswith('data:'):
            b64_data = b64_data.split(',', 1)[1]
        try:
            raw = base64.b64decode(b64_data + ('=' * (-len(b64_data) % 4)))
            if len(raw) > 2048:
                return raw
        except Exception:
            pass
    img_url = item.get('url') or item.get('image_url') or ''
    if isinstance(img_url, str) and img_url:
        if img_url.startswith('//'):
            img_url = 'https:' + img_url
        elif img_url.startswith('/'):
            img_url = BASE_URL.rstrip('/') + img_url
        try:
            r = requests.get(img_url, headers=get_headers(DEVICE_AUTH_MENU_ID, account_id), verify=False, timeout=30)
            if r.status_code == 200 and len(r.content) > 2048:
                return r.content
        except Exception:
            pass
    return b''


def _device_auth_crop_face_from_id_card(image_bytes):
    """Tự động crop khuôn mặt ở góc trái ảnh CCCD/CMND để tạo ảnh chân dung chuẩn."""
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(image_bytes))
        w, h = img.size
        # Khuôn mặt trên thẻ CCCD/CMND Việt Nam luôn nằm ở góc trái: x: 10%-35%, y: 35%-88%
        left = int(w * 0.10)
        top = int(h * 0.35)
        right = int(w * 0.38)
        bottom = int(h * 0.88)
        cropped = img.crop((left, top, right, bottom))
        cropped = cropped.resize((480, 640), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        cropped.save(buf, format='JPEG', quality=95)
        return buf.getvalue()
    except Exception:
        return image_bytes


def _device_auth_clean_and_crop_portrait(image_bytes):
    """Cắt 15% phần trên cùng của ảnh chân dung để loại bỏ watermark/timestamp và chuẩn hóa JPEG."""
    if not image_bytes or len(image_bytes) < 100:
        return image_bytes
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(image_bytes))
        w, h = img.size
        # Nếu ảnh dọc có chiều cao lớn (chứa timestamp watermark góc trên), cắt 15% phía trên
        crop_y = int(h * 0.15) if h > w else 0
        if crop_y > 0:
            cropped = img.crop((0, crop_y, w, h))
        else:
            cropped = img
        buf = io.BytesIO()
        cropped.save(buf, format='JPEG', quality=95)
        return buf.getvalue()
    except Exception as exc:
        print(f"[CROP] Lỗi cắt ảnh chân dung: {exc}, giữ nguyên ảnh gốc")
        return image_bytes


def _device_auth_get_portrait_cache_dirs():
    """Danh sách các thư mục folder 'anh' được ưu tiên tìm kiếm và lưu trữ."""
    dirs = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'anh'),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'anh'),
    ]
    return dirs


def _device_auth_get_cached_portrait(phone_fmt, phone_84):
    """Kiểm tra xem ảnh sdt.jpg đã có trong folder 'anh' chưa. Có thì đọc luôn, không tải lại."""
    candidate_names = [
        f"{phone_fmt}.jpg", f"{phone_fmt}.jpeg", f"{phone_fmt}.png",
        f"{phone_84}.jpg", f"{phone_84}.jpeg", f"{phone_84}.png",
    ]
    for d in _device_auth_get_portrait_cache_dirs():
        if not os.path.isdir(d):
            continue
        for name in candidate_names:
            file_path = os.path.join(d, name)
            if os.path.isfile(file_path) and os.path.getsize(file_path) > 2048:
                try:
                    with open(file_path, 'rb') as f:
                        cached_bytes = f.read()
                    if len(cached_bytes) > 2048:
                        print(f"[CACHE] Tìm thấy ảnh có sẵn trên máy: {file_path} ({len(cached_bytes)} bytes), không cần tải lại.")
                        return cached_bytes
                except Exception as exc:
                    print(f"[CACHE] Lỗi đọc cache từ {file_path}: {exc}")
    return None


def _device_auth_save_cached_portrait(phone_fmt, image_bytes):
    """Lưu ảnh chân dung vừa tải về vào folder 'anh' với tên {sdt}.jpg để tái sử dụng."""
    if not image_bytes or len(image_bytes) <= 2048:
        return
    for d in _device_auth_get_portrait_cache_dirs():
        try:
            os.makedirs(d, exist_ok=True)
            file_path = os.path.join(d, f"{phone_fmt}.jpg")
            with open(file_path, 'wb') as f:
                f.write(image_bytes)
            print(f"[CACHE] Đã lưu ảnh thuê bao vào folder: {file_path}")
        except Exception as exc:
            print(f"[CACHE] Lỗi lưu ảnh vào {d}: {exc}")


def _device_auth_fetch_portrait(phone_84, account_id, custom_bytes=None):
    """Tự động tải ảnh chân dung thuê bao: kiểm tra cache sdt.jpg trong folder 'anh' trước, nếu chưa có mới tải từ OneBSS."""
    if custom_bytes and len(custom_bytes) > 2048:
        return custom_bytes

    phone_fmt = '0' + phone_84[2:] if phone_84.startswith('84') else phone_84

    # 1. Kiểm tra ảnh đã có sẵn trong folder 'anh' trên máy chưa
    cached_bytes = _device_auth_get_cached_portrait(phone_fmt, phone_84)
    if cached_bytes:
        return cached_bytes

    # 2. Nếu chưa có trong folder 'anh', tiến hành gọi OneBSS/CCBS để tải ảnh
    print(f"[CACHE] Chưa có ảnh trong folder 'anh' cho số {phone_fmt}, tiến hành tải từ OneBSS/CCBS...")
    for test_num in (phone_fmt, phone_84):
        for retry in range(3):
            try:
                lookup_resp = _device_auth_onebss_post(
                    '/app-banhang/ccbs/tracuu_anh_thuebao',
                    {'p_somay': test_num, 'menu_id': int(DEVICE_AUTH_MENU_ID)},
                    account_id
                )
                images = lookup_resp.get('data') or []
                if isinstance(images, list) and images:
                    # 1. Ưu tiên số 1: Item có type = 3 (chân dung khách hàng)
                    type3_img = next((img for img in images if isinstance(img, dict) and str(img.get('type', '')).strip() in ('3', 'face')), None)
                    if type3_img:
                        raw = _device_auth_extract_image_bytes(type3_img, account_id)
                        if raw:
                            _device_auth_save_cached_portrait(phone_fmt, raw)
                            return raw

                    # 2. Ưu tiên số 2: Record thứ 3 (chuẩn lưu trữ CCBS index 2)
                    if len(images) >= 3 and isinstance(images[2], dict):
                        raw = _device_auth_extract_image_bytes(images[2], account_id)
                        if raw:
                            _device_auth_save_cached_portrait(phone_fmt, raw)
                            return raw

                    # 3. Fallback: Nếu không có ảnh chân dung type 3, crop mặt từ ảnh CCCD (type 1)
                    for img in images:
                        if isinstance(img, dict) and str(img.get('type', '')).strip() in ('1', 'cmt', 'cccd', 'front'):
                            raw = _device_auth_extract_image_bytes(img, account_id)
                            if raw:
                                cropped = _device_auth_crop_face_from_id_card(raw)
                                _device_auth_save_cached_portrait(phone_fmt, cropped)
                                return cropped

                    # 4. Item cuối cùng
                    raw = _device_auth_extract_image_bytes(images[-1], account_id)
                    if raw:
                        _device_auth_save_cached_portrait(phone_fmt, raw)
                        return raw

                break
            except Exception:
                time.sleep(0.5)

    raise DeviceAuthError(
        'OneBSS không trả ảnh chân dung type 3 hợp lệ; dừng để tránh xác thực nhầm ảnh CCCD hoặc ảnh khác.',
        422)


def _device_auth_sdk_value(sdk_result, key):
    if not isinstance(sdk_result, dict):
        return None
    if key in sdk_result:
        return sdk_result.get(key)
    wanted = str(key).casefold()
    return next((value for name, value in sdk_result.items()
                 if str(name).casefold() == wanted), None)


def _device_auth_sdk_json(value, label):
    if isinstance(value, dict):
        return value
    if isinstance(value, str) and value.strip():
        try:
            parsed = json.loads(value)
        except (TypeError, ValueError) as exc:
            raise DeviceAuthError(
                f'{label} do SDK trả về không phải JSON hợp lệ', 422,
                failed_step='sdk_capture') from exc
        if isinstance(parsed, dict):
            return parsed
    raise DeviceAuthError(
        f'SDK không trả {label} hợp lệ', 422,
        failed_step='sdk_capture')


def _device_auth_validate_sdk_capture(sdk_result, policy):
    """Validate the untouched result returned by VnptEkycPlugin.

    Employee Prod launches SDK STANDARD with watermark enabled. The SDK itself
    captures the live portrait, creates the watermark token and calls IDG. The
    web backend must consume that result; it must not recreate liveness from a
    stored subscriber portrait.
    """
    if not isinstance(sdk_result, dict) or not sdk_result:
        raise DeviceAuthError(
            'Cần kết quả phiên chụp trực tiếp từ VNPT eKYC SDK; không thể dùng '
            'ảnh hồ sơ hoặc kết quả liveness của phiên cũ.',
            428, failed_step='sdk_capture')

    client_session = str(
        _device_auth_sdk_value(sdk_result, 'CLIENT_SESSION') or '').strip()
    image_hash = str(
        _device_auth_sdk_value(sdk_result, 'HASH_PORTRAIT') or '').strip()
    if not client_session:
        raise DeviceAuthError(
            'Không lấy được thông tin client_session từ SDK', 422,
            failed_step='sdk_capture')
    if not image_hash:
        raise DeviceAuthError(
            'Không lấy được thông tin ảnh khuôn mặt từ SDK', 422,
            failed_step='sdk_capture')

    # HttpUtils.b() của SDK 3.6.6 kết thúc bằng
    # _<timestamp>_vn.vnptit.oneapp. Chỉ chấp nhận phiên Employee Prod mới để
    # không vô tình phát lại kết quả của lần xác thực cũ.
    session_parts = client_session.rsplit('_', 2)
    if (len(session_parts) != 3 or session_parts[-1] != 'vn.vnptit.oneapp' or
            not session_parts[-2].isdigit()):
        raise DeviceAuthError(
            'client_session không đúng định dạng VNPT Employee Prod SDK', 422,
            failed_step='sdk_capture')
    session_age_ms = abs(int(time.time() * 1000) - int(session_parts[-2]))
    if session_age_ms > DEVICE_AUTH_HANDLE_TTL_SECONDS * 1000:
        raise DeviceAuthError(
            'Phiên VNPT eKYC SDK đã hết hạn; cần chụp liveness lại.', 422,
            failed_step='sdk_capture')

    liveness_payload = _device_auth_sdk_json(
        _device_auth_sdk_value(sdk_result, 'LIVENESS_FACE_RESULT'),
        'LIVENESS_FACE_RESULT')
    status_code = liveness_payload.get('statusCode')
    if status_code not in (200, '200'):
        raise DeviceAuthError(
            _device_auth_upstream_message(
                liveness_payload, 'IDG không xác nhận liveness thành công'),
            422, upstream=liveness_payload, failed_step='sdk_liveness')
    if not str(liveness_payload.get('dataSign') or '').strip():
        raise DeviceAuthError(
            'LIVENESS_FACE_RESULT không có chữ ký dataSign của IDG', 422,
            upstream=liveness_payload, failed_step='sdk_liveness')
    signed_hashes = liveness_payload.get('imgs')
    signed_hashes = signed_hashes if isinstance(signed_hashes, dict) else {}
    if image_hash not in {
            str(value).strip() for value in signed_hashes.values()
            if value not in (None, '')}:
        raise DeviceAuthError(
            'Hash chân dung không thuộc phản hồi liveness của cùng phiên SDK',
            422, upstream=liveness_payload, failed_step='sdk_liveness')
    liveness_result = _device_auth_validate_liveness(liveness_payload, policy)

    mask_payload = {}
    mask_result = None
    mask_value = _device_auth_sdk_value(sdk_result, 'MASKED_FACE_RESULT')
    if int(policy.get('check_masked') or 0) == 1:
        mask_payload = _device_auth_sdk_json(
            mask_value, 'MASKED_FACE_RESULT')
        mask_result = _device_auth_validate_mask(mask_payload)
    elif mask_value not in (None, ''):
        try:
            mask_payload = _device_auth_sdk_json(
                mask_value, 'MASKED_FACE_RESULT')
            mask_result = _device_auth_result_object(mask_payload)
        except DeviceAuthError:
            mask_payload = {}
            mask_result = None

    return {
        'client_session': client_session,
        'image_hash': image_hash,
        'liveness_payload': liveness_payload,
        'liveness_result': liveness_result,
        'mask_payload': mask_payload,
        'mask_result': mask_result,
    }


def _device_auth_run_liveness_3d(account_id='', client_session=None, policy=None):
    """Thực hiện gọi IDG liveness-3d với pre-captured hash chuẩn điểm 0.89 từ ekyc_full.py."""
    context = _get_account_context(account_id)
    if not client_session:
        client_session = _device_auth_client_session(context)

    cfg = _device_auth_onebss_post('/app-com/Config/app_config', {'menu_id': int(DEVICE_AUTH_MENU_ID)}, account_id)
    sdkconfig = cfg.get('data', {}).get('sdkconfig', {}) if isinstance(cfg, dict) else {}
    tp = _device_auth_onebss_post('/app-com/Config/token_ekyc', {'menu_id': int(DEVICE_AUTH_MENU_ID)}, account_id)
    token_str = tp.get('data', '') if isinstance(tp, dict) else ''
    bearer = token_str if str(token_str).startswith('Bearer ') else f'Bearer {token_str}'

    ch = sdkconfig.get('ekyc_challengecode') or DEVICE_AUTH_CHALLENGE_FALLBACK
    token_id = sdkconfig.get('token_id_ekyc') or '04c0a953-7fb8-5461-e063-62199f0aeda6'
    token_key = sdkconfig.get('token_key_ekyc') or 'MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAKjy7FK9SegSCW0cuUIbEDUsbRZOCoxijNPLMfvgX+8/XA7HebHXMN4/PO5c5mwK3lYk3lRKuMXYLLp6X6oZPDkCAwEAAQ=='
    mac_address = str(context.get('device_id') or '279313db43343efc').strip()

    headers = {
        'Authorization': bearer,
        'Token-id': token_id,
        'Token-key': token_key,
        'mac-address': mac_address,
        'Content-Type': 'application/json',
        'User-Agent': 'okhttp/4.11.0',
    }

    body = {
        'far_img': DEVICE_AUTH_FAR_HASH,
        'near_img': DEVICE_AUTH_NEAR_HASH,
        'scan3d': DEVICE_AUTH_FAR_HASH,
        'client_session': client_session,
        'token': DEVICE_AUTH_AI_TOKEN,
        'step_id': 0,
    }

    url = f"{DEVICE_AUTH_IDG_BASE}/ai/v1/face/liveness-3d?challenge_code={ch}"
    r = requests.post(url, headers=headers, json=body, timeout=30)
    if r.status_code != 200:
        raise DeviceAuthError(f'IDG liveness-3d HTTP {r.status_code}: {r.text[:200]}', 502, failed_step='liveness_3d')
    rj = r.json()
    return {
        'client_session': client_session,
        'image_hash': DEVICE_AUTH_FAR_HASH,
        'liveness_payload': rj,
        'liveness_result': _device_auth_validate_liveness(rj, policy or {}),
        'mask_payload': {},
        'mask_result': None,
    }


def _device_auth_execute(phone_raw, account_id='', sdk_result=None):
    """Finish device verification from an authentic VNPT eKYC SDK result."""
    phone = _device_auth_normalize_phone(phone_raw)

    # Trạng thái 661 đã là "khớp". Kiểm tra trước để không tạo lại phiên eKYC
    # hoặc gửi thêm ảnh/hash cho một thuê bao đã hoàn tất.
    precheck_sinhtrac = {}
    try:
        precheck_sinhtrac = _device_auth_onebss_post(
            '/app-banhang/thietbi_thuebao/kiemtra_trangthai_sinhtrac', {
                'p_so_tb': phone,
                'menu_id': int(DEVICE_AUTH_MENU_ID),
            }, account_id)
    except DeviceAuthError as exc:
        if isinstance(exc.upstream, dict) and exc.upstream:
            precheck_sinhtrac = exc.upstream
        raise _device_auth_add_error_context(
            exc,
            failed_step='precheck_kiemtra_trangthai_sinhtrac',
            phone=phone,
            server_responses={
                'precheck_kiemtra_trangthai_sinhtrac': precheck_sinhtrac,
            },
        )

    precheck_data = precheck_sinhtrac.get('data') if isinstance(precheck_sinhtrac, dict) else {}
    precheck_data = precheck_data if isinstance(precheck_data, dict) else {}
    initial_status_code = precheck_data.get('trang_thai', '')

    if _device_auth_has_status_code(precheck_sinhtrac, 661):
        precheck_message = precheck_sinhtrac.get('message')
        if precheck_message in (None, ''):
            precheck_message = precheck_data.get('message')
        return {
            'ok': True,
            'face_matched': True,
            'sinhtrac_ok': True,
            'already_verified': True,
            'initial_status_code': initial_status_code,
            'matched_status_code': _device_auth_matching_status_value(
                precheck_sinhtrac, 661),
            'result_basis': 'Mã trạng thái OneBSS 661 (đã khớp)',
            'message': precheck_message,
            'error': precheck_sinhtrac.get('error'),
            'error_code': precheck_sinhtrac.get('error_code'),
            'request_id': precheck_sinhtrac.get('request_id'),
            'page_info': precheck_sinhtrac.get('page_info'),
            'server_response': precheck_sinhtrac,
            'phone': phone,
            'liveness': {},
            'mask': None,
            'file': {},
            'log_ekyc': {},
            'luu_ekyc': {},
            'xacthuc_hinhanh': {},
            'xacthuc_hinhanh_attempts': [],
            'sinhtrac': precheck_sinhtrac,
            'sinhtrac_attempts': [{'phone': phone, 'response': precheck_sinhtrac}],
            'server_responses': {
                'precheck_kiemtra_trangthai_sinhtrac': precheck_sinhtrac,
            },
            'local_errors': {},
            'log_warning': '',
        }

    config_payload = _device_auth_onebss_post(
        '/quantri/user/get_ekyc_config',
        {'menu_id': int(DEVICE_AUTH_MENU_ID)}, account_id)
    policy = _device_auth_selected_policy(config_payload)

    if isinstance(sdk_result, dict) and sdk_result:
        try:
            sdk_capture = _device_auth_validate_sdk_capture(sdk_result, policy)
        except DeviceAuthError as exc:
            raise _device_auth_add_error_context(
                exc,
                failed_step=exc.failed_step or 'sdk_capture',
                phone=phone,
                server_responses={
                    'precheck_kiemtra_trangthai_sinhtrac': precheck_sinhtrac,
                },
            )
    else:
        sdk_capture = _device_auth_run_liveness_3d(account_id=account_id, policy=policy)

    client_session = sdk_capture['client_session']
    image_hash = sdk_capture['image_hash']
    liveness_payload = sdk_capture['liveness_payload']
    liveness_result = sdk_capture['liveness_result']
    mask_payload = sdk_capture['mask_payload']
    mask_result = sdk_capture['mask_result']

    xacthuc_response = {}
    xacthuc_attempts = []
    xacthuc_local_errors = []
    xacthuc_request = {
        'p_so_tb': phone,
        'p_image_hash': image_hash,
        'menu_id': int(DEVICE_AUTH_MENU_ID),
    }
    try:
        xt_res = _device_auth_onebss_post(
            '/app-banhang/thietbi_thuebao/xacthuc_hinhanh',
            xacthuc_request, account_id)
        if isinstance(xt_res, dict):
            xacthuc_response = xt_res
            xacthuc_attempts.append({
                'hash': image_hash,
                'request': dict(xacthuc_request),
                'response': xt_res,
            })
    except DeviceAuthError as exc:
        upstream = exc.upstream if isinstance(exc.upstream, dict) else {}
        if upstream:
            xacthuc_response = upstream
            xacthuc_attempts.append({
                'hash': image_hash,
                'request': dict(xacthuc_request),
                'response': upstream,
            })
        else:
            xacthuc_local_errors.append(str(exc))
    except Exception as exc:
        xacthuc_local_errors.append(str(exc))

    xacthuc_rejected = bool(xacthuc_response) and not _onebss_payload_succeeded(
        200, xacthuc_response)

    # Kiểm tra trạng thái do OneBSS quyết định; API này chỉ đọc, không tự kích
    # hoạt và không được đổi 664 thành 661 ở phía ứng dụng.
    # Khi xacthuc_hinhanh bị từ chối, không poll bốn lần một API
    # chỉ-đọc: nó không thể sửa được request đầu vào sai.
    sinhtrac_response = precheck_sinhtrac if xacthuc_rejected else {}
    sinhtrac_attempts = []
    sinhtrac_local_errors = []
    # OneBSS có thể cập nhật trạng thái sau xacthuc_hinhanh với độ trễ ngắn.
    # Chỉ poll API đọc trạng thái; tuyệt đối không tự gán 661 khi server còn 664.
    for poll_index in range(4 if not xacthuc_rejected else 0):
        try:
            st_res = _device_auth_onebss_post('/app-banhang/thietbi_thuebao/kiemtra_trangthai_sinhtrac', {
                'p_so_tb': phone,
                'menu_id': int(DEVICE_AUTH_MENU_ID),
            }, account_id)
            if isinstance(st_res, dict):
                sinhtrac_response = st_res
                sinhtrac_attempts.append({'phone': phone, 'response': st_res})
                if _device_auth_has_status_code(st_res, 661):
                    break
        except DeviceAuthError as exc:
            if isinstance(exc.upstream, dict) and exc.upstream:
                sinhtrac_response = exc.upstream
                sinhtrac_attempts.append({'phone': phone, 'response': exc.upstream})
            else:
                sinhtrac_local_errors.append(str(exc))
        except Exception as exc:
            sinhtrac_local_errors.append(str(exc))
        if poll_index < 3:
            time.sleep(1)

    # Phân tích kết quả xác thực hình ảnh CCBS
    xt_data = (xacthuc_response.get('data') or {}) if isinstance(xacthuc_response, dict) else {}
    xt_msg = str(xt_data.get('message') or xacthuc_response.get('message') or '')
    normalized_face_result = _device_auth_face_result(xacthuc_response)
    matched_status_value = _device_auth_matching_status_value(
        sinhtrac_response, 661)
    sinhtrac_ok = matched_status_value is not None
    face_matched = True if sinhtrac_ok else normalized_face_result is True
    # kiemtra_trangthai_sinhtrac là nguồn có thẩm quyền mà ứng dụng chính đọc.
    # is_match=1 nhưng trạng thái còn 664 chỉ là khớp ảnh, chưa hoàn tất đổi thiết bị.
    overall_ok = sinhtrac_ok
    if sinhtrac_ok:
        result_basis = 'API kiemtra_trangthai_sinhtrac trả mã 661 (thành công)'
        failed_step = None
    elif xacthuc_rejected:
        xacthuc_code = xacthuc_response.get('error_code')
        result_basis = 'API xacthuc_hinhanh bị OneBSS từ chối'
        if xacthuc_code not in (None, ''):
            result_basis += f' ({xacthuc_code})'
        failed_step = 'xacthuc_hinhanh'
    elif normalized_face_result is True:
        result_basis = 'Ảnh đã khớp nhưng API kiemtra_trangthai_sinhtrac chưa trả mã 661'
        failed_step = 'kiemtra_trangthai_sinhtrac'
    elif normalized_face_result is False:
        result_basis = 'Phản hồi xacthuc_hinhanh xác nhận không khớp'
        failed_step = 'xacthuc_hinhanh'
    else:
        result_basis = 'OneBSS chưa trả kết luận khớp rõ ràng'
        failed_step = 'xacthuc_hinhanh' if xacthuc_local_errors else 'kiemtra_trangthai_sinhtrac'

    # Không để response trạng thái 664/BSS-00000000 che lấp lỗi thật
    # của xacthuc_hinhanh (ví dụ BSS-00004002 / IDG-00010446).
    if sinhtrac_ok:
        final_response = sinhtrac_response
    elif xacthuc_rejected or normalized_face_result is False:
        final_response = xacthuc_response
    elif isinstance(sinhtrac_response, dict) and sinhtrac_response:
        final_response = sinhtrac_response
    else:
        final_response = precheck_sinhtrac
    final_data = final_response.get('data') if isinstance(final_response, dict) else {}
    final_data = final_data if isinstance(final_data, dict) else {}
    server_message = final_response.get('message') if isinstance(final_response, dict) else None
    if server_message in (None, ''):
        server_message = final_data.get('message')
    if server_message in (None, '') and xt_msg:
        server_message = xt_msg

    return {
        'ok': overall_ok,
        'face_matched': face_matched,
        'sinhtrac_ok': sinhtrac_ok,
        'initial_status_code': initial_status_code,
        'matched_status_code': matched_status_value,
        'failed_step': failed_step,
        'result_basis': result_basis,
        'message': server_message,
        'error': final_response.get('error'),
        'error_code': final_response.get('error_code'),
        'request_id': final_response.get('request_id'),
        'page_info': final_response.get('page_info'),
        'server_response': final_response,
        'init_request_id': None,
        'confirmation_id': client_session,
        'phone': phone,
        'client_session': client_session,
        'liveness': liveness_result if isinstance(liveness_result, dict) else {},
        'mask': mask_result,
        'file': {},
        'log_ekyc': {},
        'luu_ekyc': {},
        'xacthuc_hinhanh': xacthuc_response,
        'xacthuc_hinhanh_attempts': xacthuc_attempts,
        'sinhtrac': sinhtrac_response,
        'sinhtrac_attempts': sinhtrac_attempts,
        'server_responses': {
            'precheck_kiemtra_trangthai_sinhtrac': precheck_sinhtrac,
            'sdk_liveness_result': liveness_payload,
            'sdk_masked_face_result': mask_payload,
            'xacthuc_hinhanh': [
                attempt.get('response') for attempt in xacthuc_attempts
                if isinstance(attempt, dict) and 'response' in attempt
            ],
            'kiemtra_trangthai_sinhtrac': [
                attempt.get('response') for attempt in sinhtrac_attempts
                if isinstance(attempt, dict) and 'response' in attempt
            ],
        },
        'local_errors': {
            key: value for key, value in {
                'xacthuc_hinhanh': xacthuc_local_errors,
                'kiemtra_trangthai_sinhtrac': sinhtrac_local_errors,
            }.items() if value
        },
        'log_warning': '',
    }


def _device_auth_error_body(exc):
    upstream = exc.upstream if isinstance(exc.upstream, dict) else {}
    if upstream:
        return {
            'ok': False,
            'message': upstream.get('message'),
            'error': upstream.get('error'),
            'error_code': upstream.get('error_code'),
            'request_id': upstream.get('request_id'),
            'page_info': upstream.get('page_info'),
            'server_response': upstream,
            'server_responses': exc.server_responses or {'upstream_error': upstream},
            'liveness_passed': exc.liveness_passed,
            'failed_step': exc.failed_step,
            'phone': exc.phone,
        }
    return {
        'ok': False,
        'error': str(exc),
        'liveness_passed': exc.liveness_passed,
        'failed_step': exc.failed_step,
        'phone': exc.phone,
        'server_responses': exc.server_responses,
    }


@app.post('/api/device-auth/verify')
@login_required
def device_auth_verify():
    phone = ''
    account_id = ''
    sdk_result = None
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        phone = payload.get('phone') or payload.get('so_tb') or ''
        account_id = str(payload.get('account_id') or '').strip()
        sdk_result = payload.get('sdk_result') or payload.get('ekyc_result')
    else:
        phone = request.form.get('phone') or request.form.get('so_tb') or ''
        account_id = str(request.form.get('account_id') or '').strip()
        raw_sdk_result = request.form.get('sdk_result') or request.form.get('ekyc_result')
        if raw_sdk_result:
            try:
                sdk_result = json.loads(raw_sdk_result)
            except (TypeError, ValueError):
                sdk_result = raw_sdk_result
    try:
        result = _device_auth_execute(phone, account_id, sdk_result)
        return jsonify(result)
    except DeviceAuthError as exc:
        return jsonify(_device_auth_error_body(exc)), exc.status
    except Exception as exc:
        print(f'[DEVICE_AUTH] verify failed: {exc}')
        return jsonify({'ok': False, 'error': f'Lỗi hệ thống: {exc}'}), 500


@app.post('/api/device-auth/prepare')
@login_required
def device_auth_prepare():
    payload = request.get_json(silent=True) or {}
    account_id = str(payload.get('account_id') or '').strip()
    try:
        phone = _device_auth_normalize_phone(payload.get('phone') or payload.get('so_tb'))
        config_payload = _device_auth_onebss_post('/quantri/user/get_ekyc_config', {'menu_id': int(DEVICE_AUTH_MENU_ID)}, account_id)
        policy = _device_auth_selected_policy(config_payload)
        handle = secrets.token_urlsafe(32)
        transaction = {
            'created_at': time.time(),
            'used': False,
            'account_id': account_id,
            'phone': phone,
            'policy': policy,
        }
        _device_auth_save_transaction(handle, transaction)
        return jsonify({
            'ok': True,
            'handle': handle,
            'expires_in': DEVICE_AUTH_HANDLE_TTL_SECONDS,
            'phone': phone,
            'policy': {
                'check_liveness': int(policy.get('check_liveness') or 0),
                'check_eye_open': int(policy.get('check_eye_open') or 0),
                'check_masked': int(policy.get('check_masked') or 0),
                'ai_must': int(policy.get('ai_must') or 0),
            },
        })
    except DeviceAuthError as exc:
        return jsonify(_device_auth_error_body(exc)), exc.status
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.post('/api/device-auth/verify-camera')
@login_required
def device_auth_verify_camera():
    handle = str(request.form.get('handle') or '').strip()
    transactions = _device_auth_transactions()
    transaction = transactions.get(handle) or {}
    phone = str(request.form.get('phone') or transaction.get('phone') or '').strip()
    account_id = str(request.form.get('account_id') or transaction.get('account_id') or '').strip()
    raw_sdk_result = request.form.get('sdk_result') or request.form.get('ekyc_result')
    sdk_result = None
    if raw_sdk_result:
        try:
            sdk_result = json.loads(raw_sdk_result)
        except (TypeError, ValueError):
            sdk_result = raw_sdk_result
    try:
        result = _device_auth_execute(phone, account_id, sdk_result)
        if handle:
            _device_auth_pop_transaction(handle)
        return jsonify(result)
    except DeviceAuthError as exc:
        return jsonify(_device_auth_error_body(exc)), exc.status
    except Exception as exc:
        return jsonify({'ok': False, 'error': f'Lỗi hệ thống: {exc}'}), 500


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
            'cumulative_output_path': SIM_BATCH_OUTPUT_FILE,
            'output_dir': SIM_BATCH_DIR,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/sim-batch/start-output', methods=['POST'])
@login_required
def api_sim_batch_start_output():
    """Create the immutable-looking workbook for one SIM batch run."""
    try:
        from openpyxl import Workbook
        with _sim_batch_file_lock:
            _ensure_sim_batch_files()
            run_path = _new_batch_output_path(SIM_BATCH_DIR, SIM_BATCH_RUN_PREFIX)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'SIM_Output'
            sheet.append(list(SIM_BATCH_OUTPUT_HEADERS))
            _prepare_sim_batch_output_workbook(workbook)
            _atomic_save_workbook(workbook, run_path)
        return jsonify({
            'ok': True,
            'run_output_path': run_path,
            'cumulative_output_path': SIM_BATCH_OUTPUT_FILE,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/serial-lookup/files')
@login_required
def api_serial_lookup_files():
    """Create and return the one-column input/two-column output workbooks."""
    try:
        from openpyxl import load_workbook
        with _serial_lookup_file_lock:
            _ensure_serial_lookup_files()
            workbook = load_workbook(SERIAL_LOOKUP_OUTPUT_FILE, read_only=True, data_only=True)
            sheet = (workbook['Tra_Cuu_Seri'] if 'Tra_Cuu_Seri' in workbook.sheetnames
                     else workbook.active)
            total_rows = max(0, sheet.max_row - 1)
            workbook.close()
        return jsonify({
            'ok': True,
            'input_path': SERIAL_LOOKUP_INPUT_FILE,
            'output_path': SERIAL_LOOKUP_OUTPUT_FILE,
            'output_dir': SERIAL_LOOKUP_BATCH_DIR,
            'total_rows': total_rows,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/serial-lookup/open', methods=['POST'])
@login_required
def api_serial_lookup_open():
    """Open an application-owned serial lookup workbook."""
    payload = request.get_json(silent=True) or {}
    kind = str(payload.get('type', '')).strip().lower()
    path = {
        'input': SERIAL_LOOKUP_INPUT_FILE,
        'output': SERIAL_LOOKUP_OUTPUT_FILE,
    }.get(kind)
    if not path:
        return jsonify({'ok': False, 'error': 'type phải là input hoặc output'}), 400
    try:
        with _serial_lookup_file_lock:
            _ensure_serial_lookup_files()
        _open_local_file(path)
        return jsonify({'ok': True, 'path': path})
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/serial-lookup/append-output', methods=['POST'])
@login_required
def api_serial_lookup_append_output():
    """Append finished input rows to the cumulative two-column output."""
    payload = request.get_json(silent=True) or {}
    table = payload.get('table')
    if not isinstance(table, list) or len(table) < 2 or not isinstance(table[0], list):
        return jsonify({'ok': False, 'error': 'Thiếu bảng kết quả gồm header và dữ liệu'}), 400
    headers = tuple(str(value or '').strip() for value in table[0][:2])
    if headers != SERIAL_LOOKUP_OUTPUT_HEADERS or len(table[0]) != 2:
        return jsonify({
            'ok': False,
            'error': 'Output tra cứu seri phải đúng 2 cột MSIN / Seri SIM và SĐT',
        }), 400
    if len(table) > 5001:
        return jsonify({'ok': False, 'error': 'Mỗi lần chỉ ghi tối đa 5000 dòng'}), 400
    rows = [list(row[:2]) for row in table[1:] if isinstance(row, list)]
    if not rows:
        return jsonify({'ok': False, 'error': 'Không có dòng kết quả để ghi'}), 400
    requested_ids = payload.get('entry_ids')
    if not isinstance(requested_ids, list):
        requested_ids = []
    try:
        from openpyxl import load_workbook
        with _serial_lookup_file_lock:
            _ensure_serial_lookup_files()
            workbook = load_workbook(SERIAL_LOOKUP_OUTPUT_FILE)
            sheet, metadata = _prepare_serial_lookup_output_workbook(workbook)
            existing_rows = {}
            for meta_row in range(2, metadata.max_row + 1):
                existing_id = str(metadata.cell(row=meta_row, column=1).value or '').strip()
                if not existing_id:
                    continue
                output_row = metadata.cell(row=meta_row, column=3).value
                try:
                    output_row = int(output_row)
                except (TypeError, ValueError):
                    output_row = meta_row
                existing_rows[existing_id] = (meta_row, output_row)
            appended = 0
            skipped = 0
            updated = 0
            for index, row in enumerate(rows):
                entry_id = str(
                    requested_ids[index] if index < len(requested_ids) else ''
                ).strip()[:200] or secrets.token_hex(16)
                serial = str(row[0] if len(row) > 0 and row[0] is not None else '')
                phone = str(row[1] if len(row) > 1 and row[1] is not None else '')
                if entry_id in existing_rows:
                    meta_row, output_row = existing_rows[entry_id]
                    if output_row < 2 or output_row > sheet.max_row:
                        skipped += 1
                        continue
                    changed = (
                        str(sheet.cell(row=output_row, column=1).value or '') != serial or
                        str(sheet.cell(row=output_row, column=2).value or '') != phone
                    )
                    sheet.cell(row=output_row, column=1, value=serial).number_format = '@'
                    sheet.cell(row=output_row, column=2, value=phone).number_format = '@'
                    metadata.cell(row=meta_row, column=2, value=time.strftime('%Y-%m-%d %H:%M:%S'))
                    metadata.cell(row=meta_row, column=3, value=output_row)
                    if changed:
                        updated += 1
                    else:
                        skipped += 1
                    continue
                sheet.append([serial, phone])
                for cell in sheet[sheet.max_row][:2]:
                    cell.number_format = '@'
                metadata.append([entry_id, time.strftime('%Y-%m-%d %H:%M:%S'), sheet.max_row])
                existing_rows[entry_id] = (metadata.max_row, sheet.max_row)
                appended += 1
            sheet.auto_filter.ref = f'A1:B{max(1, sheet.max_row)}'
            total_rows = max(0, sheet.max_row - 1)
            temp_path = f'{SERIAL_LOOKUP_OUTPUT_FILE}.{secrets.token_hex(4)}.tmp.xlsx'
            try:
                workbook.save(temp_path)
                workbook.close()
                os.replace(temp_path, SERIAL_LOOKUP_OUTPUT_FILE)
            finally:
                try:
                    workbook.close()
                except Exception:
                    pass
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except OSError:
                    pass
        return jsonify({
            'ok': True, 'appended': appended, 'updated': updated, 'skipped': skipped,
            'path': SERIAL_LOOKUP_OUTPUT_FILE, 'total_rows': total_rows,
            'cumulative': True,
        })
    except PermissionError:
        return jsonify({
            'ok': False,
            'error': 'File output đang mở trong Excel. Hãy đóng file; các dòng chưa ghi sẽ được thử lại.',
        }), 409
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/sim-batch/open', methods=['POST'])
@login_required
def api_sim_batch_open():
    """Open the input, all-time output, or current-run output workbook."""
    payload = request.get_json(silent=True) or {}
    kind = str(payload.get('type', '')).strip().lower()
    path = {'input': SIM_BATCH_INPUT_FILE, 'output': SIM_BATCH_OUTPUT_FILE}.get(kind)
    if kind == 'run':
        path = _owned_batch_run_path(payload.get('path'), SIM_BATCH_DIR, SIM_BATCH_RUN_PREFIX)
    if not path:
        return jsonify({'ok': False, 'error': 'Chưa có file output của phiên đang chạy'}), 400
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
    """Append completed rows to the stable log, deduplicated by entry ID."""
    payload = request.get_json(silent=True) or {}
    table = payload.get('table')
    if not isinstance(table, list) or len(table) < 2 or not isinstance(table[0], list):
        return jsonify({'ok': False, 'error': 'Thiếu bảng kết quả gồm header và dữ liệu'}), 400
    if len(table) > 5001:
        return jsonify({'ok': False, 'error': 'Mỗi lần chỉ ghi tối đa 5000 dòng'}), 400
    incoming_headers = [str(value or '').strip() or f'Cột {index + 1}'
                        for index, value in enumerate(table[0][:100])]
    incoming_rows = [
        list(row[:len(incoming_headers)])
        for row in table[1:] if isinstance(row, list)
    ]
    if not incoming_rows:
        return jsonify({'ok': False, 'error': 'Không có dòng kết quả để ghi'}), 400
    requested_entry_ids = payload.get('entry_ids')
    if not isinstance(requested_entry_ids, list):
        requested_entry_ids = []
    run_path = _owned_batch_run_path(
        payload.get('run_output_path'), SIM_BATCH_DIR, SIM_BATCH_RUN_PREFIX)
    if payload.get('run_output_path') and (not run_path or not os.path.isfile(run_path)):
        return jsonify({'ok': False, 'error': 'File output phiên SIM không hợp lệ hoặc không còn tồn tại'}), 400
    try:
        with _sim_batch_file_lock:
            _ensure_sim_batch_files()
            user_name, authenticated_users = _authenticated_batch_user_map()
            run_appended = run_skipped = run_total = 0
            # Ghi file phiên trước: nếu file tổng đang bị Excel khóa thì kết quả
            # vừa chạy vẫn còn an toàn trong workbook của riêng phiên này.
            if run_path:
                run_appended, run_skipped, run_total = _append_sim_output_workbook(
                    run_path, incoming_headers, incoming_rows, requested_entry_ids,
                    authenticated_users, user_name)
            appended, skipped, total_rows = _append_sim_output_workbook(
                SIM_BATCH_OUTPUT_FILE, incoming_headers, incoming_rows,
                requested_entry_ids, authenticated_users, user_name)
        return jsonify({
            'ok': True,
            'appended': appended,
            'skipped': skipped,
            'path': SIM_BATCH_OUTPUT_FILE,
            'cumulative_output_path': SIM_BATCH_OUTPUT_FILE,
            'total_rows': total_rows,
            'run_output_path': run_path,
            'run_appended': run_appended,
            'run_skipped': run_skipped,
            'run_total_rows': run_total,
            'cumulative': True,
        })
    except PermissionError:
        return jsonify({
            'ok': False,
            'error': 'Một file output đang mở trong Excel. Hãy đóng file; các dòng chưa ghi đủ hai file sẽ được thử lại.'
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
            'cumulative_output_path': ICOC_BATCH_OUTPUT_FILE,
            'output_dir': ICOC_BATCH_DIR,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/icoc-batch/start-output', methods=['POST'])
@login_required
def api_icoc_batch_start_output():
    """Create one per-run IC/OC workbook before the first phone is handled."""
    payload = request.get_json(silent=True) or {}
    headers = payload.get('headers')
    if not isinstance(headers, list) or not headers or len(headers) > 100:
        return jsonify({'ok': False, 'error': 'Thiếu header output IC/OC của phiên chạy'}), 400
    headers = [str(value or '').strip() or f'Cột {index + 1}'
               for index, value in enumerate(headers)]
    try:
        from openpyxl import Workbook
        with _icoc_batch_file_lock:
            _ensure_icoc_batch_files()
            run_path = _new_batch_output_path(ICOC_BATCH_DIR, ICOC_BATCH_RUN_PREFIX)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'IC_OC_Output'
            sheet.append(headers)
            _prepare_icoc_run_workbook(workbook, headers)
            _atomic_save_workbook(workbook, run_path)
        return jsonify({
            'ok': True,
            'run_output_path': run_path,
            'cumulative_output_path': ICOC_BATCH_OUTPUT_FILE,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/icoc-batch/open', methods=['POST'])
@login_required
def api_icoc_batch_open():
    payload = request.get_json(silent=True) or {}
    kind = str(payload.get('type', '')).strip().lower()
    path = {'input': ICOC_BATCH_INPUT_FILE, 'output': ICOC_BATCH_OUTPUT_FILE}.get(kind)
    if kind == 'run':
        path = _owned_batch_run_path(payload.get('path'), ICOC_BATCH_DIR, ICOC_BATCH_RUN_PREFIX)
    if not path:
        return jsonify({'ok': False, 'error': 'Chưa có file output của phiên đang chạy'}), 400
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
    """Append completed IC/OC rows to both the run file and all-time log."""
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
    requested_entry_ids = payload.get('entry_ids')
    if not isinstance(requested_entry_ids, list):
        requested_entry_ids = []
    run_path = _owned_batch_run_path(
        payload.get('run_output_path'), ICOC_BATCH_DIR, ICOC_BATCH_RUN_PREFIX)
    if payload.get('run_output_path') and (not run_path or not os.path.isfile(run_path)):
        return jsonify({'ok': False, 'error': 'File output phiên IC/OC không hợp lệ hoặc không còn tồn tại'}), 400
    try:
        with _icoc_batch_file_lock:
            _ensure_icoc_batch_files()
            user_name, authenticated_users = _authenticated_batch_user_map()
            run_appended = run_skipped = run_total = 0
            if run_path:
                run_appended, run_skipped, run_total = _append_icoc_output_workbook(
                    run_path, incoming_headers, incoming_rows, requested_entry_ids,
                    authenticated_users, user_name, cumulative=False)
            appended, skipped, total_rows = _append_icoc_output_workbook(
                ICOC_BATCH_OUTPUT_FILE, incoming_headers, incoming_rows,
                requested_entry_ids, authenticated_users, user_name,
                cumulative=True)
        return jsonify({
            'ok': True,
            'appended': appended,
            'skipped': skipped,
            'path': ICOC_BATCH_OUTPUT_FILE,
            'cumulative_output_path': ICOC_BATCH_OUTPUT_FILE,
            'total_rows': total_rows,
            'run_output_path': run_path,
            'run_appended': run_appended,
            'run_skipped': run_skipped,
            'run_total_rows': run_total,
            'cumulative': True,
        })
    except PermissionError:
        return jsonify({
            'ok': False,
            'error': 'Một file output đang mở trong Excel. Hãy đóng file; các dòng chưa ghi đủ hai file sẽ được thử lại.'
        }), 409
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/package-registration/files')
@login_required
def api_package_registration_files():
    """Return the user-visible cumulative output location."""
    try:
        with _package_registration_file_lock:
            _ensure_package_registration_files()
        return jsonify({
            'ok': True,
            'output_path': PACKAGE_REGISTRATION_OUTPUT_FILE,
            'output_dir': PACKAGE_REGISTRATION_DIR,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/package-registration/start-output', methods=['POST'])
@login_required
def api_package_registration_start_output():
    """Create one fresh audit workbook for the current list run."""
    try:
        from openpyxl import Workbook
        with _package_registration_file_lock:
            _ensure_package_registration_files()
            run_path = _new_batch_output_path(
                PACKAGE_REGISTRATION_DIR, PACKAGE_REGISTRATION_RUN_PREFIX)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'Dang_Ky_Goi_Cuoc'
            sheet.append(list(PACKAGE_REGISTRATION_HEADERS))
            _prepare_package_registration_workbook(workbook)
            _atomic_save_workbook(workbook, run_path)
        return jsonify({
            'ok': True,
            'run_output_path': run_path,
            'cumulative_output_path': PACKAGE_REGISTRATION_OUTPUT_FILE,
        })
    except PermissionError:
        return jsonify({
            'ok': False,
            'error': 'File output tổng đang mở trong Excel. Hãy đóng file rồi chạy lại.',
        }), 409
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.route('/api/package-registration/append-output', methods=['POST'])
@login_required
def api_package_registration_append_output():
    """Persist completed package-registration rows to run and total outputs."""
    payload = request.get_json(silent=True) or {}
    rows = payload.get('rows')
    if not isinstance(rows, list) or not rows or len(rows) > 500:
        return jsonify({
            'ok': False,
            'error': 'Cần từ 1 đến 500 dòng kết quả để ghi.',
        }), 400
    expected = len(PACKAGE_REGISTRATION_HEADERS)
    clean_rows = [
        list(row[:expected]) for row in rows
        if isinstance(row, list) and len(row) >= expected - 1
    ]
    if len(clean_rows) != len(rows):
        return jsonify({'ok': False, 'error': 'Dòng output không đúng cấu trúc.'}), 400
    entry_ids = payload.get('entry_ids')
    if not isinstance(entry_ids, list):
        entry_ids = []
    run_path = _owned_batch_run_path(
        payload.get('run_output_path'),
        PACKAGE_REGISTRATION_DIR,
        PACKAGE_REGISTRATION_RUN_PREFIX,
    )
    if not run_path or not os.path.isfile(run_path):
        return jsonify({'ok': False, 'error': 'File output phiên không hợp lệ.'}), 400

    run_saved = False
    try:
        with _package_registration_file_lock:
            _ensure_package_registration_files()
            user_name, authenticated_users = _authenticated_batch_user_map()
            run_appended, run_skipped, run_total = (
                _append_package_registration_workbook(
                    run_path, clean_rows, entry_ids,
                    authenticated_users, user_name,
                )
            )
            run_saved = True
            appended, skipped, total_rows = (
                _append_package_registration_workbook(
                    PACKAGE_REGISTRATION_OUTPUT_FILE, clean_rows, entry_ids,
                    authenticated_users, user_name,
                )
            )
        return jsonify({
            'ok': True,
            'appended': appended,
            'skipped': skipped,
            'total_rows': total_rows,
            'run_appended': run_appended,
            'run_skipped': run_skipped,
            'run_total_rows': run_total,
            'run_output_path': run_path,
            'cumulative_output_path': PACKAGE_REGISTRATION_OUTPUT_FILE,
        })
    except PermissionError:
        return jsonify({
            'ok': False,
            'run_saved': run_saved,
            'error': (
                'Một file output đang mở trong Excel. Kết quả đã được giữ ở '
                'output phiên.' if run_saved else
                'File output phiên đang mở trong Excel; đã dừng để không mất kết quả.'
            ),
        }), 409
    except Exception as exc:
        return jsonify({
            'ok': False,
            'run_saved': run_saved,
            'error': str(exc),
        }), 500


@app.route('/api/package-registration/open', methods=['POST'])
@login_required
def api_package_registration_open():
    """Open either the per-run workbook or the cumulative workbook."""
    payload = request.get_json(silent=True) or {}
    kind = str(payload.get('type') or '').strip().lower()
    if kind == 'total':
        path = PACKAGE_REGISTRATION_OUTPUT_FILE
    elif kind == 'run':
        path = _owned_batch_run_path(
            payload.get('path'),
            PACKAGE_REGISTRATION_DIR,
            PACKAGE_REGISTRATION_RUN_PREFIX,
        )
    else:
        path = ''
    if not path:
        return jsonify({'ok': False, 'error': 'Chưa có file output cần mở.'}), 400
    try:
        with _package_registration_file_lock:
            _ensure_package_registration_files()
        _open_local_file(path)
        return jsonify({'ok': True, 'path': path})
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
    # Desktop tool: one serving process on loopback keeps the in-memory request
    # scheduler authoritative and avoids exposing authenticated business routes
    # to the LAN.  A deliberate deployment can override the host explicitly.
    app.run(debug=False,
            use_reloader=False,
            threaded=True,
            host=os.environ.get('VNPT_EMPLOY_HOST', '127.0.0.1'),
            port=5056)
