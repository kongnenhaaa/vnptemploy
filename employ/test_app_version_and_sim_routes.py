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
    LEGACY_SIM_READ_ENDPOINTS = (
        '/ccbs/chonSo/search_isdn',
        '/app-com/danhmuc/get_danhmuc',
        '/app-banhang/donhang_simkit/danhsach_goicuoc',
        '/app-banhang/donhang_simkit/listdsdonhang_v2',
        '/app-thuno/VnptPay/kiemTraViVnptPay',
        '/app-thuno/VnptPay/getBalance',
    )
    LEGACY_SIM_MUTATION_ENDPOINTS = (
        '/app-banhang/donhang_simkit/chonso_kit_v2',
        '/app-banhang/donhang_simkit/huy_donhang',
        '/app-banhang/donhang_simkit/dangky_goicuoc',
        '/app-banhang/donhang_simkit/nhap_thongtin_khachhang_v3',
        '/app-banhang/donhang_simkit/xacnhan_thanhtoan',
        '/app-banhang/donhang_simkit/khoitao_thuebao',
    )
    CUSTOMER_SELF_REG_READ_ENDPOINTS = (
        '/ccbs/chonSo/app_ds_dauso',
        '/ccbs/chonSo/checkSimStatus',
        '/app-banhang/kenhban-simkit/search_isdn_shop_mid',
        '/app-banhang/kenhban-simkit/step_donhang_simkit',
        '/app-banhang/kenhban-simkit/danhsach_goicuoc',
        '/app-banhang/kenhban-simkit/kiemtra_soluong_tb',
        '/app-banhang/kenhban-simkit/listdsdonhang_v2',
        '/app-com/danhmuc/get_danhmuc',
        '/web-quantri/danhmuc-chung/lay_tt_ts_diadanh_moi',
        '/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen',
        '/app-ccdv/vietqr/check_donhang',
    )
    CUSTOMER_SELF_REG_MUTATION_ENDPOINTS = (
        '/app-banhang/kenhban-simkit/chonso_kit_v2',
        '/app-banhang/kenhban-simkit/dangky_goicuoc',
        '/app-banhang/kenhban-simkit/nhap_thongtin_khachhang_v3',
        '/app-banhang/kenhban-simkit/khoitao_thuebao',
        '/app-banhang/kenhban-simkit/xacnhan_thanhtoan',
        '/app-banhang/kenhban-simkit/hoanthanh_donhang_tratruoc',
    )

    def test_legacy_sim_endpoints_keep_menu_699161_and_guard(self):
        endpoints = (self.LEGACY_SIM_READ_ENDPOINTS +
                     self.LEGACY_SIM_MUTATION_ENDPOINTS)
        for endpoint in endpoints:
            with self.subTest(endpoint=endpoint):
                body = {'menu_id': 699161}
                self.assertEqual(
                    app_module.endpoint_menu_id(endpoint, '0', body), '699161')
                policy = app_module._business_endpoint_policy(endpoint, body)
                self.assertIsNotNone(policy)
                self.assertEqual(policy['workflow'], 'sim')
                self.assertEqual(
                    policy['mutation'],
                    endpoint in self.LEGACY_SIM_MUTATION_ENDPOINTS)

    def test_customer_self_registration_uses_captured_menu_and_guard(self):
        endpoints = (self.CUSTOMER_SELF_REG_READ_ENDPOINTS +
                     self.CUSTOMER_SELF_REG_MUTATION_ENDPOINTS)
        for endpoint in endpoints:
            with self.subTest(endpoint=endpoint):
                body = {'menu_id': 810641}
                if endpoint.endswith('/mhddm_kiemtra_maquyen'):
                    body['ma_quyen'] = 'KHOITAOTB'
                self.assertEqual(
                    app_module.endpoint_menu_id(endpoint, '0', body), '810641')
                policy = app_module._business_endpoint_policy(endpoint, body)
                self.assertIsNotNone(policy)
                self.assertEqual(policy['workflow'], 'sim')
                self.assertEqual(
                    policy['mutation'],
                    endpoint in self.CUSTOMER_SELF_REG_MUTATION_ENDPOINTS)

    def test_customer_self_registration_wallet_calls_keep_menu_810641(self):
        for endpoint in (
            '/app-thuno/VnptPay/kiemTraViVnptPay',
            '/app-thuno/VnptPay/getBalance',
        ):
            with self.subTest(endpoint=endpoint):
                self.assertEqual(
                    app_module.endpoint_menu_id(
                        endpoint, '699161', {'menu_id': 810641}),
                    '810641',
                )

    def test_mutation_registry_matches_full_sim_order_flow(self):
        self.assertEqual(
            set(app_module._SIM_MUTATION_ENDPOINTS),
            set(self.LEGACY_SIM_MUTATION_ENDPOINTS) |
            set(self.CUSTOMER_SELF_REG_MUTATION_ENDPOINTS),
        )


class SimKitCustomerSelfRegistrationTemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        template_path = os.path.join(
            os.path.dirname(__file__), 'templates', 'dashboard.html')
        with open(template_path, encoding='utf-8') as template_file:
            cls.source = template_file.read()

    def _function_source(self, name, next_name):
        start = self.source.index(f'async function {name}')
        end = self.source.index(f'async function {next_name}', start)
        return self.source[start:end]

    def test_manual_flow_uses_captured_module_and_menu(self):
        self.assertIn("const SIM_SELF_REG_MENU_ID = 810641;", self.source)
        self.assertIn(
            "const SIM_SELF_REG_PREFIX = '/app-banhang/kenhban-simkit';",
            self.source,
        )
        self.assertIn('search_isdn_shop_mid', self.source)
        self.assertIn('step_donhang_simkit', self.source)
        self.assertIn('kiemtra_soluong_tb', self.source)

    def test_customer_payload_is_locked_to_self_registration(self):
        customer = self._function_source(
            'simConfirmCustomer', 'simPrepareVnptWallet')
        self.assertIn('p_id_hinhthuc_dk_tttb: 2', customer)
        self.assertIn('p_loai_giayto: null', customer)
        self.assertIn('p_so_gt: citizenId', customer)
        self.assertIn("expectedStatusId", self.source)
        self.assertIn("simCapturedStateMatches(step, simWizardState.orderId, '4')", customer)

    def test_initialization_precedes_payment_and_completion(self):
        initialization = self.source[
            self.source.index('async function simCheckSimStatus()'):
            self.source.index('// ── Khởi tạo SIM hàng loạt từ Excel')
        ]
        self.assertLess(
            initialization.index("ma_quyen: 'KHOITAOTB'"),
            initialization.index("'/ccbs/chonSo/checkSimStatus'"),
        )
        self.assertLess(
            initialization.index("'/ccbs/chonSo/checkSimStatus'"),
            initialization.index('`${SIM_SELF_REG_PREFIX}/khoitao_thuebao`'),
        )
        self.assertIn('simWizardState.completed = false', initialization)
        self.assertIn('simRenderPaymentPending(initialized)', initialization)

    def test_batch_is_enabled_after_completed_payment_capture(self):
        self.assertIn('const SIM_SELF_REG_BATCH_ENABLED = true;', self.source)

    def test_manual_payment_uses_completed_captured_sequence(self):
        payment = self.source[
            self.source.index('async function simConfirmPayment'):
            self.source.index('function simTodayApiDate')
        ]
        self.assertIn('await simRefreshVnptWalletCredentials()', payment)
        self.assertIn('`${SIM_SELF_REG_PREFIX}/xacnhan_thanhtoan`', payment)
        self.assertIn('p_thongtin_hoadon: null', payment)
        self.assertNotIn('menu_id: 699161', payment)
        self.assertIn("paidState.statusId !== '7'", payment)
        self.assertIn("paidState.paymentStatusId !== '1'", payment)
        self.assertIn('await simCompletePaidOrder(r)', payment)

        completion = self.source[
            self.source.index('async function simCompletePaidOrder'):
            self.source.index('function simRenderPaidSession')
        ]
        self.assertIn('`${SIM_SELF_REG_PREFIX}/hoanthanh_donhang_tratruoc`', completion)
        self.assertIn("p_ma_donhang:''", completion)
        self.assertIn('simFindSelfRegOrderInHistory', completion)
        self.assertIn('simSelfRegHistoryIsComplete', completion)
        self.assertLess(
            completion.index('hoanthanh_donhang_tratruoc'),
            completion.index('simFindSelfRegOrderInHistory'),
        )
        self.assertLess(
            completion.index('simFindSelfRegOrderInHistory'),
            completion.index('simWizardState.completed = true'),
        )

    def test_batch_input_requires_cccd_and_maps_it_to_self_registration(self):
        self.assertIn('citizenIdColumn: 2', self.source)
        self.assertIn('function simBatchNormalizeCitizenId(value)', self.source)
        self.assertIn("'so giay to','p so gt','pidnumber'", self.source)
        self.assertIn("missing.push('CCCD')", self.source)
        self.assertIn('citizenId = simBatchNormalizeCitizenId(rawCitizenId)', self.source)
        batch_process = self.source[
            self.source.index('async function simBatchProcessRow'):
            self.source.index('function simHistoryTodayIso')
        ]
        self.assertIn('pidnumber:row.citizenId', batch_process)
        self.assertIn('p_id_hinhthuc_dk_tttb:2', batch_process)
        self.assertIn('p_loai_giayto:null', batch_process)
        self.assertIn('p_so_gt:row.citizenId', batch_process)

    def test_excel_batch_uses_self_registration_flow_in_captured_order(self):
        batch_process = self.source[
            self.source.index('async function simBatchProcessRow'):
            self.source.index('function simHistoryTodayIso')
        ]
        expected_markers = [
            '`${SIM_SELF_REG_PREFIX}/chonso_kit_v2`',
            "await requireStep('2')",
            '`${SIM_SELF_REG_PREFIX}/dangky_goicuoc`',
            "await requireStep('3')",
            '`${SIM_SELF_REG_PREFIX}/nhap_thongtin_khachhang_v3`',
            "await requireStep('4')",
            "ma_quyen:'KHOITAOTB'",
            "'/ccbs/chonSo/checkSimStatus'",
            '`${SIM_SELF_REG_PREFIX}/khoitao_thuebao`',
            "await requireStep('6')",
            'await simBatchRefreshWalletCredentials',
            '`${SIM_SELF_REG_PREFIX}/xacnhan_thanhtoan`',
            '`${SIM_SELF_REG_PREFIX}/hoanthanh_donhang_tratruoc`',
            'simBatchFindSelfRegOrderInHistory',
        ]
        positions = [batch_process.index(marker) for marker in expected_markers]
        self.assertEqual(positions, sorted(positions))
        self.assertIn('p_thongtin_hoadon:null', batch_process)
        self.assertIn('simSelfRegHistoryIsComplete(historyRow, orderId, row.serial)', batch_process)
        self.assertNotIn("'/app-banhang/donhang_simkit/", batch_process)

    def test_excel_batch_derives_country_prefix_from_the_last_seven_digits(self):
        helper = self.source[
            self.source.index('function simBatchSearchParts'):
            self.source.index('function simBatchNormalizeSerial')
        ]
        search = self.source[
            self.source.index('async function simBatchSearchNumber'):
            self.source.index('async function simBatchProcessRow')
        ]
        self.assertIn("const prefix = normalized.slice(0, -7);", helper)
        self.assertIn("const suffix = normalized.slice(-7);", helper)
        self.assertIn("/^84\\d{2}$/.test(prefix)", helper)
        self.assertIn("const {prefix, suffix} = simBatchSearchParts(row.msisdn);", search)
        self.assertIn("p_prefix:prefix, p_isdn:suffix", search)
        self.assertNotIn("const prefix = nationalNumber", helper + search)


if __name__ == '__main__':
    unittest.main()
