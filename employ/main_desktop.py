import os
import sys
import threading
import webview
import time
import socket
import urllib.request
import traceback
import ctypes
import multiprocessing

# Ensure the correct paths are used for PyInstaller bundles
if getattr(sys, 'frozen', False):
    sys.path.insert(0, sys._MEIPASS)

from app import (
    app,
    _ensure_application_storage,
    _remove_legacy_default_accounts_once,
)

def get_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(('127.0.0.1', 0))
    port = s.getsockname()[1]
    s.close()
    return port

port = get_free_port()

def start_server():
    # Run the Flask app on the free port
    app.run(host='127.0.0.1', port=port, use_reloader=False, debug=False)

def wait_for_server(timeout=20):
    deadline = time.time() + timeout
    url = f'http://127.0.0.1:{port}/'
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2) as response:
                return response.status == 200
        except Exception:
            time.sleep(0.2)
    return False

def report_startup_error(error):
    log_root = os.path.join(
        os.environ.get('LOCALAPPDATA') or os.path.dirname(sys.executable),
        'VNPTEmploy'
    )
    os.makedirs(log_root, exist_ok=True)
    log_path = os.path.join(log_root, 'startup_error.log')
    with open(log_path, 'w', encoding='utf-8') as handle:
        handle.write(traceback.format_exc())
    if os.name == 'nt':
        ctypes.windll.user32.MessageBoxW(
            0,
            f'Không thể khởi động VNPT Employee.\n\n{error}\n\nChi tiết: {log_path}',
            'VNPT Employee',
            0x10,
        )

if __name__ == '__main__':
    multiprocessing.freeze_support()
    try:
        # Tạo toàn bộ thư mục/file cần thiết ngay khi mở EXE, kể cả khi người
        # dùng chưa đăng nhập và dashboard chưa được hiển thị.
        _ensure_application_storage()
        _remove_legacy_default_accounts_once()

        # Start Flask server in a daemon thread
        t = threading.Thread(target=start_server, name='vnpt-local-server', daemon=True)
        t.start()
        if not wait_for_server():
            raise RuntimeError('Máy chủ nội bộ không phản hồi sau 20 giây')

        # Chế độ dùng để xác minh EXE sau khi build mà không mở cửa sổ.
        if '--smoke-test' in sys.argv:
            raise SystemExit(0)

        # Create the native desktop window containing the Flask app. PyWebView
        # tự dùng Edge WebView2 và tự fallback về MSHTML trên Windows cũ.
        webview.create_window(
            'VNPT Employee',
            f'http://127.0.0.1:{port}',
            width=1280,
            height=800,
            min_size=(1024, 680),
        )
        webview.start(private_mode=False)
    except SystemExit:
        raise
    except Exception as error:
        report_startup_error(error)
        raise
