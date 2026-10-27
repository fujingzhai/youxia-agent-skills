# -*- coding: utf-8 -*-
"""批次停止条件与有限重试。不会自动解除服务端限流或持续挑战。"""
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import urlparse
import captcha


class PauseBatch(RuntimeError):
    def __init__(self, reason, retry_after=None):
        super().__init__(reason)
        self.retry_after = retry_after


class PageTimeout(RuntimeError):
    pass


def retry_seconds(value):
    if not value:
        return None
    try:
        return max(0, int(value))
    except ValueError:
        try:
            return max(0, int((parsedate_to_datetime(value) - datetime.now(timezone.utc)).total_seconds()))
        except (ValueError, TypeError):
            return None


class Guard:
    def __init__(self, page):
        self.reason = None
        self.retry_after = None
        page.on('response', self.observe)

    def observe(self, response):
        # 字形图单个 403/404 不表示机构会话失效；正文和检索请求才触发批次暂停。
        host = urlparse(response.url).hostname or ''
        if not (host == 'cnki.net' or host.endswith('.cnki.net')):
            return
        if 'attachment' in response.url:
            return
        if response.request.resource_type not in ('document', 'xhr', 'fetch'):
            return
        if response.status in (401, 403, 429):
            self.reason = f'知网返回 HTTP {response.status}，暂停批次，需检查权限/限流状态'
            self.retry_after = retry_seconds(response.headers.get('retry-after'))

    def check(self):
        if self.reason:
            raise PauseBatch(self.reason, self.retry_after)


def ensure_access(page, guard):
    guard.check()
    if captcha.present(page):
        # 只给一次自动处理机会；失败后交给人工，避免整批不断刷新挑战。
        if not captcha.solve(page, tries=1):
            raise PauseBatch('验证码未通过，已暂停批次；请用 cnki --open 人工处理')
    guard.check()
    try:
        text = page.locator('body').inner_text(timeout=3000)[:1500]
    except Exception:
        text = ''
    if any(t in text for t in ('操作过于频繁', '访问过于频繁', '访问异常', '异常访问', '请依次点击')):
        raise PauseBatch('页面提示访问受限或点选验证，已暂停批次')
    if any(t in text for t in ('登录已失效', '登录已过期', '机构权限已失效')):
        raise PauseBatch('机构登录失效，已暂停批次')


def wait_content(page, selector, guard, timeout=20000, empty_selector=None):
    """等待结果、明确空结果或挑战；超时不是“无结果”。"""
    ensure_access(page, guard)
    # 结果与空提示一起等：知网无结果 1 秒内就出「暂无数据」，只等结果会白等到超时、误报成加载失败。
    # 滑块也一起等：它晚于 domcontentloaded 才渲染（新 profile 首次必出），过了再等一轮结果。
    target = f"{selector}, {empty_selector}" if empty_selector else selector
    for _ in range(2):
        try:
            page.wait_for_selector(f"{target}, .verify-move-block", timeout=timeout)
        except Exception as e:
            ensure_access(page, guard)
            raise PageTimeout('页面未在时限内加载完整，保留为待重试') from e
        if not captcha.present(page):
            break
        ensure_access(page, guard)
    ensure_access(page, guard)
    return not (empty_selector and not page.locator(selector).count())
