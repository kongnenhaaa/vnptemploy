import os
import time
import unittest
from unittest.mock import patch

import app as app_module


class _Response:
    def __init__(self, status_code, payload, headers=None):
        self.status_code = status_code
        self._payload = payload
        self.headers = headers or {}
        self.text = ''

    def json(self):
        return self._payload


class IcocRateLimitGuardTests(unittest.TestCase):
    def tearDown(self):
        with app_module._business_guard_lock:
            app_module._business_account_paused_until.pop('icoc-test', None)

    def test_onebss_business_throttle_is_recognised_even_on_http_500(self):
        self.assertTrue(app_module._business_payload_rate_limited({
            'error_code': 'BSS-00000500',
            'message': '13137567: gọi quá nhiều',
        }))
        self.assertFalse(app_module._business_payload_rate_limited({
            'error_code': 'BSS-00004002',
            'message': 'Sai dữ liệu',
        }))

    def test_business_throttle_pauses_account_and_returns_retry_after(self):
        response = _Response(500, {
            'error_code': 'BSS-00000500',
            'message': '13137567: gọi quá nhiều',
        })
        with patch.object(time, 'monotonic', return_value=100.0):
            retry_after = app_module._business_record_response(
                'icoc-test', response)
        self.assertEqual(
            retry_after, app_module.BUSINESS_RATE_LIMIT_RETRY_SECONDS)
        self.assertEqual(
            app_module._business_account_paused_until['icoc-test'],
            100.0 + app_module.BUSINESS_RATE_LIMIT_RETRY_SECONDS,
        )


class IcocRateLimitTemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        template_path = os.path.join(
            os.path.dirname(__file__), 'templates', 'dashboard.html')
        with open(template_path, encoding='utf-8') as template_file:
            cls.source = template_file.read()

    def test_rate_limit_is_not_classified_as_unknown_mutation(self):
        helper_start = self.source.index(
            'function isBusinessMutationUncertain')
        helper_end = self.source.index(
            'async function fetchMobileSubscriberState', helper_start)
        helper = self.source[helper_start:helper_end]
        self.assertIn('!isApiRateLimited(response)', helper)
        self.assertIn("code === 'BSS-00000500'", self.source)

    def test_icoc_retries_throttle_and_keeps_delay_across_inputs(self):
        self.assertIn(
            'async function icocBatchCallWithRateLimitRetry', self.source)
        self.assertIn('retryAfter || 30', self.source)
        self.assertIn('accountReadyAt: new Map()', self.source)
        self.assertIn(
            'icocBatchState.accountReadyAt.set(', self.source)


if __name__ == '__main__':
    unittest.main()
