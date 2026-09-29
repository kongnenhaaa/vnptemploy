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
from urllib.parse import urlsplit

# Ensure the correct paths are used for PyInstaller bundles
if getattr(sys, 'frozen', False):
    sys.path.insert(0, sys._MEIPASS)

from app import (
    app,
    _ensure_application_storage,
    _remove_legacy_default_accounts_once,
)


_WALLET_HOST_SUFFIXES = (
    '.vnpt.vn',
    '.vnptmedia.vn',
    '.vnptpay.vn',
)


def _wallet_origin(value):
    """Return a trusted HTTPS wallet origin, or an empty string."""
    try:
        parsed = urlsplit(str(value or '').strip())
    except ValueError:
        return ''
    hostname = str(parsed.hostname or '').rstrip('.').lower()
    trusted = any(
        hostname == suffix[1:] or hostname.endswith(suffix)
        for suffix in _WALLET_HOST_SUFFIXES)
    if parsed.scheme.lower() != 'https' or not trusted or parsed.username or parsed.password:
        return ''
    try:
        port = parsed.port
    except ValueError:
        return ''
    return f'https://{hostname}{f":{port}" if port else ""}'


class WalletDesktopBridge:
    """Keep the short-lived VNPT Pay token inside the desktop process.

    Android's original InAppWebView sends ``WalletToken`` on wallet requests.
    A normal HTML iframe cannot attach a header to later navigations, which is
    why opening transaction history asks for the wallet password again.
    """

    def __init__(self):
        self._lock = threading.RLock()
        self._token = ''
        self._origin = ''
        self._expires_at = 0.0
        self._resource_handler = None

    def set_wallet_context(self, wallet_token, wallet_url):
        token = str(wallet_token or '').strip()
        origin = _wallet_origin(wallet_url)
        if not token or not origin or len(token) > 16384:
            self.clear_wallet_context()
            return {'ok': False, 'error': 'Phiên hoặc địa chỉ Ví VNPT Pay không hợp lệ'}
        with self._lock:
            self._token = token
            self._origin = origin
            self._expires_at = time.monotonic() + 10 * 60
        return {'ok': True}

    def clear_wallet_context(self):
        with self._lock:
            self._token = ''
            self._origin = ''
            self._expires_at = 0.0
        return {'ok': True}

    def _token_for_request(self, request_url):
        origin = _wallet_origin(request_url)
        with self._lock:
            if time.monotonic() >= self._expires_at:
                self._token = ''
                self._origin = ''
                return ''
            if not origin or origin != self._origin:
                return ''
            return self._token


# The source app registers an InAppWebView handler named ``setToken``.  This
# document-start shim is installed into child frames by WebView2 and mirrors
# that handler for VNPT Pay.  The token is delivered only by the local parent
# dashboard and kept in sessionStorage for same-origin wallet navigations.
WALLET_DOCUMENT_BRIDGE_SCRIPT = r"""
(() => {
  const storageKey = '__vnpt_employee_wallet_token';
  let walletToken = '';
  try { walletToken = sessionStorage.getItem(storageKey) || ''; } catch (_) {}

  window.addEventListener('message', event => {
    const message = event && event.data;
    if (!message || message.type !== 'vnpt-employee-wallet-token' || event.source !== window.parent) return;
    try {
      const parentUrl = new URL(event.origin);
      if (!['127.0.0.1', 'localhost'].includes(parentUrl.hostname)) return;
    } catch (_) { return; }
    walletToken = String(message.token || '');
    try {
      if (walletToken) sessionStorage.setItem(storageKey, walletToken);
      else sessionStorage.removeItem(storageKey);
    } catch (_) {}
  });

  const previousBridge = window.flutter_inappwebview;
  const previousCallHandler = previousBridge && typeof previousBridge.callHandler === 'function'
    ? previousBridge.callHandler.bind(previousBridge)
    : null;
  const bridge = previousBridge || {};
  bridge.callHandler = function(name, ...args) {
    if (name === 'setToken') {
      if (args.length && args[0] != null && String(args[0])) {
        walletToken = String(args[0]);
        try { sessionStorage.setItem(storageKey, walletToken); } catch (_) {}
        return Promise.resolve(true);
      }
      return Promise.resolve(walletToken);
    }
    return previousCallHandler ? previousCallHandler(name, ...args) : Promise.resolve(null);
  };
  window.flutter_inappwebview = bridge;
})();
"""


def install_wallet_desktop_hooks(window, bridge):
    """Install WebView2 hooks after the desktop window has finished loading."""
    if not window.events.loaded.wait(20):
        return
    try:
        from System import Func, Object
        from webview.platforms import winforms

        form = winforms.BrowserView.instances.get(window.uid)
        if not form or not getattr(form, 'browser', None):
            return

        def configure_webview():
            core = form.browser.webview.CoreWebView2
            if core is None:
                return None

            def add_wallet_header(_sender, args):
                token = bridge._token_for_request(str(args.Request.Uri))
                if token:
                    args.Request.Headers.SetHeader('WalletToken', token)

            bridge._resource_handler = add_wallet_header
            core.WebResourceRequested += bridge._resource_handler
            core.AddScriptToExecuteOnDocumentCreatedAsync(
                WALLET_DOCUMENT_BRIDGE_SCRIPT)
            return None

        form.browser.webview.Invoke(Func[Object](configure_webview))
    except Exception as error:
        # The app can still show the wallet page in a regular browser/fallback
        # renderer; only password-free navigation is unavailable there.
        print(f'[WARN] Could not install VNPT Pay desktop bridge: {error}')

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
        wallet_bridge = WalletDesktopBridge()
        window = webview.create_window(
            'VNPT Employee',
            f'http://127.0.0.1:{port}',
            js_api=wallet_bridge,
            width=1280,
            height=800,
            min_size=(1024, 680),
        )
        webview.start(
            install_wallet_desktop_hooks,
            (window, wallet_bridge),
            private_mode=False)
    except SystemExit:
        raise
    except Exception as error:
        report_startup_error(error)
        raise
