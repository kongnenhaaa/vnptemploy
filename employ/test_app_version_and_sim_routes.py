import base64
import json
import os
import tempfile
import unittest
from contextlib import nullcontext
from unittest.mock import patch

import app as app_module


class AppVersionTests(unittest.TestCase):
    def test_default_version_and_app_secret_are_current(self):
        self.assertEqual(app_module.DEFAULT_APP_VERSION, '1.5.41.130')
        with patch.dict(app_module.APP_CFG, {'APP_VERSION': '1.5.41.130'}):
            encoded = app_module._build_app_secret_value(device_id='test-device')
        decoded = json.loads(base64.b64decode(encoded).decode('utf-8'))
        self.assertEqual(decoded['app_version'], '1.5.41.130')
        self.assertEqual(decoded['device_id'], 'test-device')

    def test_saved_version_promotes_old_value_and_preserves_newer_value(self):
        with tempfile.TemporaryDirectory(prefix='vnpt_version_test_') as temp_dir:
            settings_path = os.path.join(temp_dir, 'app_settings.json')
            with patch.object(app_module, 'APP_SETTINGS_FILE', settings_path), \
                    patch.object(app_module, '_credential_root', temp_dir):
                app_module._save_app_version('1.5.41.090')
                self.assertEqual(
                    app_module._load_saved_app_version(), '1.5.41.130')
                app_module._save_app_version('1.5.41.131')
                self.assertEqual(
                    app_module._load_saved_app_version(), '1.5.41.131')

    def test_settings_endpoint_updates_runtime_version(self):
        original_config = dict(app_module.APP_CFG)
        try:
            with tempfile.TemporaryDirectory(prefix='vnpt_version_route_test_') as temp_dir:
                settings_path = os.path.join(temp_dir, 'app_settings.json')
                ekyc_patch = (
                    patch.object(app_module._ekyc, 'set_app_version')
                    if app_module.EKYC_AVAILABLE else nullcontext()
                )
                with patch.object(app_module, 'APP_SETTINGS_FILE', settings_path), \
                        patch.object(app_module, '_credential_root', temp_dir), \
                        patch.object(app_module, '_save_persistent_session'), \
                        ekyc_patch:
                    with app_module.app.test_request_context(
                            '/settings/update', method='POST',
                            json={'app_version': '1.5.41.131'}):
                        app_module.session['access_token'] = 'test-token'
                        response = app_module.update_settings.__wrapped__()
                        payload = response.get_json()
                self.assertTrue(payload['ok'])
                self.assertEqual(payload['app_version'], '1.5.41.131')
                self.assertEqual(app_module.APP_CFG['APP_VERSION'], '1.5.41.131')
        finally:
            app_module.APP_CFG.clear()
            app_module.APP_CFG.update(original_config)
            if app_module.EKYC_AVAILABLE and hasattr(app_module._ekyc, 'set_app_version'):
                app_module._ekyc.set_app_version(original_config['APP_VERSION'])


class SimKitRoutingTests(unittest.TestCase):
    SIM_READ_ENDPOINTS = (
        '/ccbs/chonSo/app_ds_dauso',
        '/ccbs/chonSo/search_isdn',
        '/ccbs/chonSo/checkSimStatus',
        '/app-com/danhmuc/get_danhmuc',
        '/app-banhang/donhang_simkit/danhsach_goicuoc',
        '/app-banhang/donhang_simkit/listdsdonhang_v2',
        '/app-thuno/VnptPay/kiemTraViVnptPay',
        '/app-thuno/VnptPay/getBalance',
    )
    SIM_MUTATION_ENDPOINTS = (
        '/app-banhang/donhang_simkit/chonso_kit_v2',
        '/app-banhang/donhang_simkit/huy_donhang',
        '/app-banhang/donhang_simkit/dangky_goicuoc',
        '/app-banhang/donhang_simkit/nhap_thongtin_khachhang_v3',
        '/app-banhang/donhang_simkit/xacnhan_thanhtoan',
        '/app-banhang/donhang_simkit/khoitao_thuebao',
    )

    def test_every_sim_endpoint_uses_menu_699161_and_guard(self):
        for endpoint in self.SIM_READ_ENDPOINTS + self.SIM_MUTATION_ENDPOINTS:
            with self.subTest(endpoint=endpoint):
                body = {'menu_id': 699161}
                self.assertEqual(
                    app_module.endpoint_menu_id(endpoint, '0', body), '699161')
                policy = app_module._business_endpoint_policy(endpoint, body)
                self.assertIsNotNone(policy)
                self.assertEqual(policy['workflow'], 'sim')
                self.assertEqual(
                    policy['mutation'], endpoint in self.SIM_MUTATION_ENDPOINTS)

    def test_mutation_registry_matches_full_sim_order_flow(self):
        self.assertEqual(
            set(app_module._SIM_MUTATION_ENDPOINTS),
            set(self.SIM_MUTATION_ENDPOINTS),
        )


if __name__ == '__main__':
    unittest.main()
