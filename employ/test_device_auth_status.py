import json
import time
import unittest
from unittest.mock import patch

import app as app_module


class DeviceAuthStatusTests(unittest.TestCase):
    def _sdk_result(self):
        return {
            'CLIENT_SESSION': (
                f'ANDROID_CPH2179_32_Device_3.6.6_device_'
                f'{int(time.time() * 1000)}_vn.vnptit.oneapp'),
            'HASH_PORTRAIT': 'zone2/server-image-hash',
            'LIVENESS_FACE_RESULT': json.dumps({
                'statusCode': 200,
                'message': 'IDG-00000000',
                'dataSign': 'server-signature',
                'imgs': {
                    'far_img': 'zone2/server-image-hash',
                    'near_img': 'zone2/server-image-hash',
                },
                'object': {
                    'liveness': 'success',
                    'liveness_msg': 'Người thật',
                    'liveness_prob': 0.949,
                    'is_eye_open': 'yes',
                },
            }),
        }

    def _execute_with_statuses(self, postcheck_statuses, xacthuc_error=None):
        calls = []
        statuses = iter(['664', *postcheck_statuses])

        def onebss_post(path, body, account_id=''):
            calls.append((path, body))
            if path.endswith('/kiemtra_trangthai_sinhtrac'):
                status = next(statuses)
                return {
                    'error': 200,
                    'error_code': f'REAL-{status}',
                    'message': f'Server status {status}',
                    'data': {'Status': 0, 'trang_thai': status},
                }
            if path.endswith('/get_ekyc_config'):
                return {'data': [{
                    'dichvu': '-1', 'ai_must': 1, 'check_liveness': 1,
                    'check_eye_open': 0, 'check_masked': 0,
                }]}
            if path.endswith('/xacthuc_hinhanh'):
                if xacthuc_error is not None:
                    raise app_module.DeviceAuthError(
                        xacthuc_error['message'], 400,
                        upstream=xacthuc_error)
                return {
                    'error': 200,
                    'error_code': 'BSS-00000000',
                    'message': 'Server says face matched',
                    'data': {'is_match': 1, 'Status': 0},
                }
            raise AssertionError(f'Unexpected OneBSS path: {path}')

        with (
            patch.object(
                app_module, '_device_auth_onebss_post',
                side_effect=onebss_post),
            patch.object(
                app_module.requests, 'post',
                side_effect=AssertionError(
                    'backend must not call IDG liveness')),
            patch.object(app_module.time, 'sleep', return_value=None),
        ):
            result = app_module._device_auth_execute(
                '0834518167', sdk_result=self._sdk_result())
        return result, calls

    def test_664_is_never_rewritten_as_661(self):
        result, calls = self._execute_with_statuses(
            ['664', '664', '664', '664'])

        self.assertFalse(result['ok'])
        self.assertTrue(result['face_matched'])
        self.assertFalse(result['sinhtrac_ok'])
        self.assertIsNone(result['matched_status_code'])
        self.assertEqual(result['sinhtrac']['data']['trang_thai'], '664')
        self.assertEqual(result['error_code'], 'REAL-664')
        self.assertNotIn('661', json.dumps(result['server_responses']))
        self.assertFalse(any(
            path.endswith('/add_subs_inprogress') for path, _ in calls))

    def test_real_661_from_status_api_is_success(self):
        result, calls = self._execute_with_statuses(['664', '661'])

        self.assertTrue(result['ok'])
        self.assertTrue(result['face_matched'])
        self.assertTrue(result['sinhtrac_ok'])
        self.assertEqual(result['matched_status_code'], '661')
        self.assertEqual(result['message'], 'Server status 661')
        self.assertEqual(result['error_code'], 'REAL-661')
        raw_statuses = result['server_responses'][
            'kiemtra_trangthai_sinhtrac']
        self.assertEqual(
            [item['data']['trang_thai'] for item in raw_statuses],
            ['664', '661'])
        xacthuc_calls = [
            body for path, body in calls
            if path.endswith('/xacthuc_hinhanh')]
        self.assertEqual(
            xacthuc_calls, [{'p_image_hash': 'zone2/server-image-hash'}])

    def test_xacthuc_rejection_is_not_hidden_by_status_response(self):
        rejected = {
            'error': '400',
            'error_code': 'BSS-00004002',
            'message': 'IDG-00010446: Lỗi đầu vào không hợp lệ',
            'data': None,
        }
        result, calls = self._execute_with_statuses(
            [], xacthuc_error=rejected)

        self.assertFalse(result['ok'])
        self.assertEqual(result['failed_step'], 'xacthuc_hinhanh')
        self.assertEqual(result['error_code'], 'BSS-00004002')
        self.assertEqual(result['message'], rejected['message'])
        self.assertEqual(result['server_response'], rejected)
        status_calls = [
            path for path, _ in calls
            if path.endswith('/kiemtra_trangthai_sinhtrac')]
        self.assertEqual(
            len(status_calls), 1,
            'chỉ có precheck, không poll sau lỗi xác thực')

    def test_preexisting_661_keeps_exact_server_fields(self):
        response = {
            'error': 200,
            'message': 'Exact server message',
            'data': {'trang_thai': 661},
        }
        with patch.object(
                app_module, '_device_auth_onebss_post',
                return_value=response):
            result = app_module._device_auth_execute('84834518167')

        self.assertTrue(result['ok'])
        self.assertEqual(result['matched_status_code'], 661)
        self.assertEqual(result['message'], 'Exact server message')
        self.assertIsNone(result['error_code'])
        self.assertEqual(
            result['server_responses'][
                'precheck_kiemtra_trangthai_sinhtrac'],
            response,
        )

    def test_upstream_error_body_does_not_add_default_success_fields(self):
        upstream = {'error': 400, 'message': 'Exact rejection'}
        body = app_module._device_auth_error_body(
            app_module.DeviceAuthError(
                'fallback', 400, upstream=upstream)
        )

        self.assertIsNone(body['error_code'])
        self.assertEqual(body['message'], 'Exact rejection')
        self.assertEqual(
            body['server_responses'], {'upstream_error': upstream})

    def test_upstream_error_body_preserves_completed_steps(self):
        upstream = {'statusCode': 400, 'message': 'IDG-00010446'}
        exc = app_module.DeviceAuthError(
            'invalid', 400, upstream=upstream)
        app_module._device_auth_add_error_context(
            exc,
            failed_step='sdk_liveness',
            phone='84834518167',
            server_responses={
                'precheck_kiemtra_trangthai_sinhtrac': {
                    'request_id': 'real-status-id'},
            },
        )

        body = app_module._device_auth_error_body(exc)

        self.assertEqual(body['failed_step'], 'sdk_liveness')
        self.assertEqual(body['phone'], '84834518167')
        self.assertEqual(body['server_response'], upstream)
        self.assertEqual(
            body['server_responses'],
            {
                'precheck_kiemtra_trangthai_sinhtrac': {
                    'request_id': 'real-status-id'},
                'upstream_error': upstream,
            },
        )

    def test_664_without_current_sdk_capture_stops_before_mutation(self):
        response = {
            'error': '200',
            'error_code': 'BSS-00000000',
            'message': 'Server still requires verification',
            'request_id': 'status-request-id',
            'data': {'Status': 1, 'trang_thai': '664'},
        }
        calls = []

        def onebss_post(path, body, account_id=''):
            calls.append((path, body))
            return response

        with patch.object(
                app_module, '_device_auth_onebss_post',
                side_effect=onebss_post):
            result = app_module._device_auth_execute('0834518167')

        self.assertFalse(result['ok'])
        self.assertTrue(result['requires_sdk_capture'])
        self.assertEqual(result['failed_step'], 'sdk_capture')
        self.assertEqual(result['initial_status_code'], '664')
        self.assertEqual(result['server_response'], response)
        self.assertEqual(len(calls), 1)
        self.assertTrue(
            calls[0][0].endswith('/kiemtra_trangthai_sinhtrac'))

    def test_sdk_capture_must_be_current_employee_session(self):
        sdk_result = {
            'CLIENT_SESSION': 'old-or-invented-session',
            'HASH_PORTRAIT': 'zone2/server-image-hash',
            'LIVENESS_FACE_RESULT': json.dumps({
                'statusCode': 200,
                'dataSign': 'server-signature',
                'imgs': {'far_img': 'zone2/server-image-hash'},
                'object': {'liveness': 'success'},
            }),
        }
        with self.assertRaises(app_module.DeviceAuthError) as raised:
            app_module._device_auth_validate_sdk_capture(
                sdk_result,
                {'check_masked': 0, 'check_eye_open': 0})

        self.assertEqual(raised.exception.failed_step, 'sdk_capture')

    def test_sdk_liveness_hash_must_match_portrait_hash(self):
        sdk_result = self._sdk_result()
        payload = json.loads(sdk_result['LIVENESS_FACE_RESULT'])
        payload['imgs'] = {'far_img': 'zone2/different-session-hash'}
        sdk_result['LIVENESS_FACE_RESULT'] = json.dumps(payload)

        with self.assertRaises(app_module.DeviceAuthError) as raised:
            app_module._device_auth_validate_sdk_capture(
                sdk_result,
                {'check_masked': 0, 'check_eye_open': 0})

        self.assertEqual(raised.exception.failed_step, 'sdk_liveness')
        self.assertIn('không thuộc', str(raised.exception))

    def test_precheck_error_is_not_hidden_as_sdk_requirement(self):
        upstream = {
            'error': '400',
            'error_code': 'BSS-TEST-ERROR',
            'message': 'Exact precheck failure',
        }
        with patch.object(
                app_module, '_device_auth_onebss_post',
                side_effect=app_module.DeviceAuthError(
                    upstream['message'], 400, upstream=upstream)):
            with self.assertRaises(app_module.DeviceAuthError) as raised:
                app_module._device_auth_execute('0834518167')

        self.assertEqual(
            raised.exception.failed_step,
            'precheck_kiemtra_trangthai_sinhtrac')
        self.assertEqual(raised.exception.upstream, upstream)


if __name__ == '__main__':
    unittest.main()
