"""backoff_request 模块的单元测试(使用 mock,不发真实请求)。"""

import unittest
from unittest import mock

from backoff_request import MaxRetriesExceededError, make_request_with_backoff


def _response(status_code):
    resp = mock.Mock()
    resp.status_code = status_code
    return resp


class MakeRequestWithBackoffTest(unittest.TestCase):
    @mock.patch("backoff_request.time.sleep")
    @mock.patch("backoff_request.requests.get")
    def test_success_on_first_attempt(self, mock_get, mock_sleep):
        mock_get.return_value = _response(200)

        result = make_request_with_backoff("https://example.com", headers={})

        self.assertEqual(result.status_code, 200)
        self.assertEqual(mock_get.call_count, 1)
        mock_sleep.assert_not_called()

    @mock.patch("backoff_request.time.sleep")
    @mock.patch("backoff_request.requests.get")
    def test_retries_on_429_with_exponential_waits(self, mock_get, mock_sleep):
        mock_get.side_effect = [
            _response(429),
            _response(429),
            _response(429),
            _response(200),
        ]

        result = make_request_with_backoff("https://example.com", headers={})

        self.assertEqual(result.status_code, 200)
        self.assertEqual(mock_get.call_count, 4)
        # 等待时间应为 1、2、4 秒
        self.assertEqual(
            [call.args[0] for call in mock_sleep.call_args_list], [1, 2, 4]
        )

    @mock.patch("backoff_request.time.sleep")
    @mock.patch("backoff_request.requests.get")
    def test_raises_after_max_retries(self, mock_get, mock_sleep):
        mock_get.return_value = _response(429)

        with self.assertRaises(MaxRetriesExceededError):
            make_request_with_backoff("https://example.com", headers={}, max_retries=5)

        self.assertEqual(mock_get.call_count, 5)
        self.assertEqual(
            [call.args[0] for call in mock_sleep.call_args_list], [1, 2, 4, 8, 16]
        )

    @mock.patch("backoff_request.time.sleep")
    @mock.patch("backoff_request.requests.get")
    def test_non_429_errors_returned_without_retry(self, mock_get, mock_sleep):
        mock_get.return_value = _response(500)

        result = make_request_with_backoff("https://example.com", headers={})

        # 只有 429 触发重试,其他状态码直接返回
        self.assertEqual(result.status_code, 500)
        self.assertEqual(mock_get.call_count, 1)
        mock_sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
