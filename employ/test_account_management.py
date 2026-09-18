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


if __name__ == '__main__':
    unittest.main()
