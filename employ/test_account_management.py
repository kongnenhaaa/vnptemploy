import os
import tempfile
import time
import unittest
from unittest.mock import patch

import app as app_module


class AccountManagementTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix='vnpt_account_management_test_')
        self.saved_file = os.path.join(self.temp_dir.name, 'employee_accounts.dat')
        self.session_file = os.path.join(self.temp_dir.name, 'session_state.dat')
        self.patchers = [
            patch.object(app_module, 'SAVED_ACCOUNTS_FILE', self.saved_file),
            patch.object(app_module, 'PERSISTENT_SESSION_FILE', self.session_file),
        ]
        for patcher in self.patchers:
            patcher.start()
        self.client = app_module.app.test_client()

    def tearDown(self):
        for patcher in reversed(self.patchers):
            patcher.stop()
        self.temp_dir.cleanup()

    def _login_primary(self, extras=None):
        with self.client.session_transaction() as active_session:
            active_session['username'] = 'primary.user'
            active_session['access_token'] = 'live-access-token'
            active_session['refresh_token'] = 'live-refresh-token'
            active_session['expires_in'] = 3600
            active_session['token_time'] = time.time()
            active_session['multi_accounts'] = extras or {}

    def test_live_token_is_not_refreshed_early(self):
        self._login_primary()
        with patch.object(app_module.requests, 'post') as post:
            response = self.client.post('/api/accounts/refresh', json={
                'account_id': app_module._account_id('primary.user'),
            })
        payload = response.get_json()
        self.assertEqual(response.status_code, 200)
        self.assertTrue(payload['ok'])
        self.assertTrue(payload['skipped'])
        post.assert_not_called()

    def test_delete_removes_only_selected_saved_and_active_extra_account(self):
        app_module.save_employee_account('extra.user', 'encrypted-password-source')
        app_module.save_employee_account('keep.user', 'keep-password-source')
        extra_id = app_module._account_id('extra.user')
        self._login_primary({
            extra_id: {
                'id': extra_id,
                'username': 'extra.user',
                'access_token': 'extra-token',
                'refresh_token': 'extra-refresh',
                'expires_in': 3600,
                'token_time': time.time(),
            },
        })

        response = self.client.post('/api/accounts/delete', json={'account_id': extra_id})
        payload = response.get_json()
        self.assertEqual(response.status_code, 200)
        self.assertTrue(payload['deleted'])
        self.assertFalse(payload['logout_required'])
        self.assertIsNone(app_module.get_saved_employee_account(extra_id))
        self.assertIsNotNone(app_module.get_saved_employee_account(
            app_module._account_id('keep.user')))
        with self.client.session_transaction() as active_session:
            self.assertNotIn(extra_id, active_session.get('multi_accounts', {}))

    def test_pasted_token_replaces_primary_session_without_otp(self):
        self._login_primary()
        app_module.save_employee_account('token.user', 'encrypted-password-source')
        account_id = app_module._account_id('token.user')

        class ProfileResponse:
            status_code = 200

            @staticmethod
            def json():
                return {
                    'error_code': 'BSS-00000000',
                    'data': {'sdt': '0901234567'},
                }

        with patch.object(app_module.requests, 'post', return_value=ProfileResponse()):
            response = self.client.post('/api/accounts/token-login', json={
                'account_id': account_id,
                'token': 'pasted-access-token',
            })

        payload = response.get_json()
        self.assertEqual(response.status_code, 200)
        self.assertTrue(payload['ok'])
        self.assertEqual(payload['account']['username'], 'token.user')
        self.assertEqual(payload['account']['phone'], '0901234567')
        with self.client.session_transaction() as active_session:
            self.assertEqual(active_session['username'], 'token.user')
            self.assertEqual(active_session['access_token'], 'pasted-access-token')
            self.assertEqual(active_session.get('multi_accounts'), {})

    def test_login_page_has_paste_token_form(self):
        response = self.client.get('/login')

        self.assertEqual(response.status_code, 200)
        self.assertIn(b'action="/login/token"', response.data)
        self.assertIn(b'name="token"', response.data)

    def test_token_login_from_login_page_does_not_require_existing_session(self):
        class ProfileResponse:
            status_code = 200

            @staticmethod
            def json():
                return {
                    'error_code': 'BSS-00000000',
                    'data': {
                        'username': 'clipboard.user',
                        'sdt': '0912345678',
                    },
                }

        with patch.object(app_module.requests, 'post', return_value=ProfileResponse()):
            response = self.client.post('/login/token', data={
                'token': 'Bearer clipboard-access-token',
            })

        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers['Location'].endswith('/dashboard'))
        with self.client.session_transaction() as active_session:
            self.assertEqual(active_session['username'], 'clipboard.user')
            self.assertEqual(active_session['access_token'], 'clipboard-access-token')
            self.assertEqual(active_session['account_phone'], '0912345678')

    def test_recent_signed_jwt_can_login_when_profile_rejects_device_metadata(self):
        def encoded(value):
            raw = app_module.json.dumps(value, separators=(',', ':')).encode('utf-8')
            return app_module.base64.urlsafe_b64encode(raw).decode('ascii').rstrip('=')

        token = '.'.join((
            encoded({'alg': 'RS256', 'typ': 'JWT'}),
            encoded({
                'user_name': 'fresh.user',
                'id_thietbi': '123',
                'exp': int(time.time()) + 3600,
            }),
            'signed_part',
        ))

        class RejectedProfileResponse:
            status_code = 401

            @staticmethod
            def json():
                return {'error_code': 'BSS-401', 'message': 'device mismatch'}

        with patch.object(app_module.requests, 'post', return_value=RejectedProfileResponse()):
            response = self.client.post('/login/token', data={
                # Also cover Markdown escaping of base64url underscores.
                'token': token.replace('_', r'\_'),
            })

        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers['Location'].endswith('/dashboard'))
        with self.client.session_transaction() as active_session:
            self.assertEqual(active_session['username'], 'fresh.user')
            self.assertEqual(active_session['access_token'], token)
            self.assertEqual(active_session['device_id'], '123')


if __name__ == '__main__':
    unittest.main()
