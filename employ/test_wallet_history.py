import os
import time
import unittest
from contextlib import nullcontext
from unittest.mock import patch

import app as app_module
import main_desktop


class WalletHistoryRoutingTests(unittest.TestCase):
    ENDPOINTS = (
        '/app-thuno/VnptPay/kiemTraAuToLoginViVnptPay',
        '/app-thuno/VnptPay/getWalletInfo/9',
        '/app-thuno/VnptPay/kiemTraViVnptPay',
    )

    def test_wallet_history_endpoints_use_captured_menu_10281(self):
        self.assertEqual(app_module.WALLET_HISTORY_MENU_ID, '10281')
        self.assertEqual(set(app_module.WALLET_HISTORY_ENDPOINTS), set(self.ENDPOINTS))
        for endpoint in self.ENDPOINTS:
            with self.subTest(endpoint=endpoint):
                self.assertEqual(
                    app_module.endpoint_menu_id(
                        endpoint, '699161', {'menu_id': 10281}),
                    '10281',
                )

    def test_shared_wallet_endpoint_keeps_existing_sim_menu(self):
        endpoint = '/app-thuno/VnptPay/kiemTraViVnptPay'
        self.assertEqual(
            app_module.endpoint_menu_id(
                endpoint, '10281', {'menu_id': 810641}),
            '810641',
        )
        self.assertEqual(
            app_module.endpoint_menu_id(
                endpoint, '10281', {'menu_id': 699161}),
            '699161',
        )

    def test_proxy_sends_menu_10281_in_headers_and_body(self):
        class Response:
            status_code = 200
            headers = {}

            @staticmethod
            def json():
                return {
                    'error': '0',
                    'error_code': 'BSS-00000000',
                    'data': True,
                }

        client = app_module.app.test_client()
        with client.session_transaction() as active_session:
            active_session.update({
                'username': 'wallet.user',
                'access_token': 'wallet-access-token',
                'expires_in': 3600,
                'token_time': time.time(),
                'active_menu_id': '699161',
            })

        with patch.object(
                app_module, '_business_request_slot',
                return_value=nullcontext()), patch.object(
                app_module, '_save_persistent_session') as save_session, patch.object(
                app_module.requests, 'request', return_value=Response()) as request_call:
            response = client.post('/proxy', json={
                'endpoint': '/app-thuno/VnptPay/kiemTraAuToLoginViVnptPay',
                'method': 'POST',
                'body': {'menu_id': 10281},
            })

        payload = response.get_json()
        self.assertEqual(response.status_code, 200)
        self.assertEqual(payload['selected_menu_id'], '10281')
        _, _, request_kwargs = request_call.mock_calls[0]
        self.assertEqual(request_kwargs['json'], {'menu_id': 10281})
        self.assertEqual(request_kwargs['headers']['SelectedMenuId'], '10281')
        self.assertEqual(request_kwargs['headers']['selectedmenuid'], '10281')
        save_session.assert_called_once()


class WalletHistoryTemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        template_path = os.path.join(
            os.path.dirname(__file__), 'templates', 'dashboard.html')
        with open(template_path, encoding='utf-8') as template_file:
            cls.source = template_file.read()

    def test_sidebar_and_panel_are_transaction_history_not_wallet_management(self):
        self.assertIn('data-panel="wallet-history"', self.source)
        self.assertIn('id="panel-wallet-history"', self.source)
        self.assertIn('Lịch sử giao dịch ví', self.source)
        self.assertNotIn('Quản lý ví tiền', self.source)

    def test_source_wallet_history_flow_authenticates_before_wallet_info_type_9(self):
        flow_start = self.source.index('async function walletHistoryLoad')
        flow_end = self.source.index(
            'async function walletHistoryInitialize', flow_start)
        flow = self.source[flow_start:flow_end]
        self.assertLess(
            flow.index("'/app-thuno/VnptPay/kiemTraAuToLoginViVnptPay'"),
            flow.index("'/app-thuno/VnptPay/kiemTraViVnptPay'"),
        )
        self.assertLess(
            flow.index("'/app-thuno/VnptPay/kiemTraViVnptPay'"),
            flow.index("'/app-thuno/VnptPay/getWalletInfo/9'"),
        )
        self.assertIn('{menu_id:WALLET_HISTORY_MENU_ID}', flow)
        self.assertIn('{WalletToken:walletSession.token}', flow)
        self.assertIn('if (!walletSession.token)', flow)
        self.assertIn('walletHistorySafeUrl', flow)

    def test_wallet_tokens_are_not_rendered_or_written_to_local_storage(self):
        flow_start = self.source.index('// ── Lịch sử giao dịch Ví VNPT Pay')
        flow_end = self.source.index('// ── Ảnh hồ sơ thuê bao', flow_start)
        flow = self.source[flow_start:flow_end]
        self.assertNotIn('localStorage', flow)
        self.assertNotIn('id="wallet-history-token"', flow)
        self.assertIn('walletHistoryPrimeDesktopBridge', flow)
        self.assertIn("type:'vnpt-employee-wallet-token'", flow)


class WalletDesktopBridgeTests(unittest.TestCase):
    def test_wallet_origin_only_accepts_https_vnpt_hosts(self):
        self.assertEqual(
            main_desktop._wallet_origin(
                'https://wallet.vnptmedia.vn/account?ticket=secret'),
            'https://wallet.vnptmedia.vn')
        self.assertEqual(
            main_desktop._wallet_origin('https://pay.vnpt.vn:8443/home'),
            'https://pay.vnpt.vn:8443')
        self.assertEqual(main_desktop._wallet_origin('http://pay.vnpt.vn'), '')
        self.assertEqual(main_desktop._wallet_origin('https://vnpt.vn.evil.test'), '')

    def test_wallet_token_is_scoped_to_one_origin_and_can_be_cleared(self):
        bridge = main_desktop.WalletDesktopBridge()
        result = bridge.set_wallet_context(
            'short-lived-wallet-token',
            'https://wallet.vnptmedia.vn/account')
        self.assertTrue(result['ok'])
        self.assertEqual(
            bridge._token_for_request(
                'https://wallet.vnptmedia.vn/transaction/history'),
            'short-lived-wallet-token')
        self.assertEqual(
            bridge._token_for_request('https://other.vnptmedia.vn/asset.js'),
            '')
        bridge.clear_wallet_context()
        self.assertEqual(
            bridge._token_for_request('https://wallet.vnptmedia.vn/account'),
            '')

    def test_document_bridge_matches_source_set_token_handler(self):
        script = main_desktop.WALLET_DOCUMENT_BRIDGE_SCRIPT
        self.assertIn("name === 'setToken'", script)
        self.assertIn("sessionStorage.setItem", script)
        self.assertIn("vnpt-employee-wallet-token", script)


if __name__ == '__main__':
    unittest.main()
