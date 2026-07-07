"""带指数退避的 HTTP 请求工具。

当服务端返回 429(速率受限)时,按 1、2、4、8、16 秒的间隔
指数退避后重试,超过最大重试次数则抛出异常。
"""

import time

import requests


class MaxRetriesExceededError(Exception):
    """超过最大重试次数后抛出的异常。"""


def make_request_with_backoff(url, headers=None, max_retries=5):
    """发送 GET 请求;遇到 429 时按指数退避重试。

    Args:
        url: 请求地址。
        headers: 可选的请求头。
        max_retries: 最大尝试次数(含首次请求)。

    Returns:
        requests.Response: 第一个非 429 的响应。

    Raises:
        MaxRetriesExceededError: 连续 max_retries 次都被速率限制。
    """
    for attempt in range(max_retries):
        response = requests.get(url, headers=headers)

        if response.status_code == 429:
            # 指数退避:1秒、2秒、4秒、8秒、16秒
            wait_time = 2 ** attempt
            print(f"速率受限。等待 {wait_time} 秒后重试...")
            time.sleep(wait_time)
            continue

        return response

    raise MaxRetriesExceededError("已超过最大重试次数")
