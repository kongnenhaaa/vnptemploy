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
        '/app-banhang/kenhban-simkit/xac_thuc_khuonmat',
        '/app-banhang/kenhban-simkit/add_update_khachhang',
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


class SimAssistedPortraitTests(unittest.TestCase):
    def test_portrait_is_resolved_by_exact_citizen_id(self):
        with tempfile.TemporaryDirectory(prefix='vnpt_portrait_') as temp_dir:
            expected = os.path.join(temp_dir, '066204001088.jpg')
            with open(expected, 'wb') as portrait:
                portrait.write(b'\xff\xd8\xff' + b'x' * 4096)
            with patch.object(
                    app_module, '_device_auth_get_portrait_cache_dirs',
                    return_value=[temp_dir]):
                image_bytes, path = app_module._sim_assisted_find_portrait(
                    '066204001088')
        self.assertEqual(path, expected)
        self.assertTrue(image_bytes.startswith(b'\xff\xd8\xff'))

    def test_captured_portrait_sequence_uses_menu_810641(self):
        post_calls = []

        def onebss_post(path, body, account_id='', menu_id=None):
            post_calls.append((path, body, str(menu_id)))
            if path.endswith('/init_log_uuid'):
                return {'request_id': 'real-request-id'}
            return {'error': 200, 'error_code': 'BSS-00000000'}

        with (
            patch.object(
                app_module, '_sim_assisted_find_portrait',
                return_value=(b'\xff\xd8\xff' + b'x' * 4096,
                              r'C:\portrait\066204001088.jpg')),
            patch.object(
                app_module, '_sim_assisted_onebss_get',
                return_value={'error': 200}),
            patch.object(
                app_module, '_device_auth_onebss_post',
                side_effect=onebss_post),
            patch.object(
                app_module, '_device_auth_upload_to_onebss',
                return_value={'data': {'id_taptin': 12345}}),
        ):
            result = app_module._sim_assisted_prepare_portrait(
                '066204001088', 'account-a')

        self.assertEqual(result['portrait_file_id'], 12345)
        self.assertEqual(result['portrait_file_name'], '066204001088.jpg')
        self.assertEqual(
            [call[0] for call in post_calls],
            [
                '/app-com/Config/token_ekyc',
                '/quantri/user/get_ekyc_config',
                '/app-banhang/Ekyc/init_log_uuid',
                '/app-banhang/Ekyc/log_ekyc',
            ],
        )
        self.assertTrue(all(call[2] == '810641' for call in post_calls))
        log_body = post_calls[-1][1]
        self.assertEqual(log_body['p_step'], 'OTHER_-1')
        self.assertEqual(log_body['requestId'], 'real-request-id')

    def test_assisted_completion_uses_face_then_add_update_payloads(self):
        calls = []
        identity = {
            'data': {
                'uuid_customer': 'customer-uuid',
                'customer_cards': [{
                    'id': '066204001088',
                    'name': 'OLD NAME',
                    'birth_day': '19/11/2004',
                    'gender': 'male',
                    'nationality': 'Việt Nam',
                    'extra_info': json.dumps({
                        'loai_gt': '45',
                        'loai_gt_name': 'CĂN CƯỚC CÔNG DÂN',
                        'nationalityid': '232',
                    }),
                }],
                'customer_faces': [{
                    'channel': '36',
                    'verify_status': 1,
                    'image_url': 'stored-face',
                }],
            }
        }

        def onebss_post(path, body, account_id='', menu_id=None):
            calls.append((path, body, str(menu_id)))
            if path.endswith('/xac_thuc_khuonmat'):
                return {'data': None, 'error': None,
                        'error_code': 'BSS-00000000'}
            return {
                'data': {
                    'id_donhang': 19436,
                    'id_kbsk_kh': 20066,
                    'id_trangthai': 4,
                    'id_hinhthuc_dk_tttb': 1,
                },
                'error': None,
                'error_code': 'BSS-00000000',
            }

        with (
            patch.object(
                app_module, '_sim_assisted_onebss_get',
                return_value=identity),
            patch.object(
                app_module, '_sim_assisted_prepare_portrait',
                return_value={
                    'portrait_file_id': 14720153,
                    'portrait_file_name': '066204001088.jpg',
                    'request_id': 'client-session',
                }),
            patch.object(
                app_module, '_device_auth_onebss_post',
                side_effect=onebss_post),
        ):
            result = app_module._sim_assisted_complete_customer({
                'citizen_id': '066204001088',
                'order_id': 19436,
                'customer_id': 20066,
                'customer_name': 'Dương Đức Châu',
                'phone': '84941021019',
                'subscriber_type': 21,
            }, 'account-a')

        self.assertEqual(result['portrait_file_id'], 14720153)
        self.assertEqual(
            [call[0] for call in calls],
            [
                '/app-banhang/kenhban-simkit/xac_thuc_khuonmat',
                '/app-banhang/kenhban-simkit/add_update_khachhang',
            ],
        )
        self.assertTrue(all(call[2] == '810641' for call in calls))
        face_body = calls[0][1]
        self.assertEqual(face_body['id_anh_chandung'], 14720153)
        self.assertEqual(face_body['id_donhang'], 19436)
        self.assertEqual(face_body['id_kbsk_kh'], 20066)
        add_body = calls[1][1]
        self.assertEqual(add_body['full_name'], 'Dương Đức Châu')
        self.assertEqual(add_body['customer_card']['full_name'],
                         'Dương Đức Châu')
        self.assertEqual(add_body['so_tb'], '84941021019')


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

    def test_customer_payload_supports_both_registration_methods(self):
        customer = self._function_source(
            'simConfirmCustomer', 'simPrepareVnptWallet')
        self.assertIn("const methodId = String(simWizardState.tttbMethodId", customer)
        self.assertIn("if (methodId === '1')", customer)
        self.assertIn('simCompleteAssistedCustomer({', customer)
        self.assertIn('customerName:name', customer)
        self.assertIn('r = {status:200, body:completion.order_response}', customer)
        payload = self.source[
            self.source.index('function simCustomerPayload'):
            self.source.index('async function simConfirmCustomer')
        ]
        self.assertIn('p_id_hinhthuc_dk_tttb: Number(methodId)', payload)
        self.assertIn("p_loai_giayto: methodId === '1' ? 2 : null", payload)
        self.assertIn('p_so_gt: citizenId', payload)
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

    def test_batch_input_requires_cccd_and_maps_it_to_selected_method(self):
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
        self.assertIn('p_id_hinhthuc_dk_tttb:Number(methodId)', batch_process)
        self.assertIn("p_loai_giayto:methodId === '1' ? 2 : null", batch_process)
        self.assertIn('p_so_gt:row.citizenId', batch_process)
        self.assertIn('simCompleteAssistedCustomer({', batch_process)
        self.assertIn('customerName:profile.customerName', batch_process)
        self.assertIn('customer = {status:200, body:completion.order_response}', batch_process)

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
            'simCompleteAssistedCustomer({',
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
