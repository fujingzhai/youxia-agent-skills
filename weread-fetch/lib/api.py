#!/usr/bin/env python3
"""微信读书网页版的 JSON 接口封装（书架、书目、章节表）。

正文不走这里：正文响应是分块 base64，块边界不对齐、还要客户端签名，
自己复刻既脆又容易错位。正文一律由 capture.py 在浏览器里取（见 references/protocol.md）。
"""

from __future__ import annotations

import gzip
import json
import time
import urllib.error
import urllib.parse
import urllib.request
import zlib
from pathlib import Path

from .paths import BROWSER_PY
from .zen_cookies import read_zen

SKILL_DIR = Path(__file__).resolve().parents[1]
STATE_DIR = SKILL_DIR / ".state"
COOKIE_FILE = STATE_DIR / "cookies.json"

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

# 登录态里真正起作用的两个：会话票据与续期票据
ESSENTIAL_COOKIES = ("wr_vid", "wr_skey")


class WereadError(RuntimeError):
    pass


class SessionExpired(WereadError):
    pass


class NotLoggedIn(WereadError):
    pass


# 微信读书用响应体里的 errCode 报鉴权问题，HTTP 状态照样是 200
SESSION_ERRORS = {-2012: "登录超时", -2013: "鉴权失败", -2010: "会话无效"}
CLOAK_PY = BROWSER_PY


def refresh_session(headed: bool = False, wait: int = 0) -> bool:
    """回浏览器续会话（wr_skey 只有浏览器自己续得动），再把 Cookie 抄回来。"""
    import subprocess
    cmd = [str(CLOAK_PY), str(SKILL_DIR / "lib" / "session.py")]
    if headed:
        cmd.append("--headed")
    if wait:
        cmd += ["--wait", str(wait)]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=max(180, wait + 120))
    except Exception:
        return False
    return proc.returncode == 0


# ---------------------------------------------------------------- cookies

def sync_cookies_from_browser(browser: str = "auto") -> list[dict]:
    """从本机浏览器读微信读书的登录 Cookie，落到 .state/cookies.json。

    只取 weread.qq.com 一个域，不碰其他站点的 Cookie。
    """
    try:
        import rookiepy
    except ImportError as exc:  # pragma: no cover
        raise WereadError("缺少 rookiepy，无法读浏览器 Cookie") from exc

    order = ["zen", "chrome", "edge", "brave", "firefox"] if browser == "auto" else [browser]
    errors = []
    for name in order:
        reader = read_zen if name == "zen" else getattr(rookiepy, name, None)
        if reader is None:
            errors.append(f"{name}: 不支持")
            continue
        try:
            cookies = reader(["weread.qq.com"]) if name == "zen" else reader(domains=["weread.qq.com"])
        except Exception as exc:
            errors.append(f"{name}: {type(exc).__name__}")
            continue
        names = {c["name"] for c in cookies}
        if not set(ESSENTIAL_COOKIES) <= names:
            errors.append(f"{name}: 无 wr_skey（未登录）")
            continue
        STATE_DIR.mkdir(exist_ok=True)
        payload = {"source": name, "synced": int(time.time()), "cookies": cookies}
        COOKIE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=1, default=str))
        COOKIE_FILE.chmod(0o600)
        return cookies
    raise NotLoggedIn("没在本机浏览器里找到微信读书登录态；" + "；".join(errors))


def load_cookies(auto_sync: bool = True) -> list[dict]:
    if COOKIE_FILE.exists():
        data = json.loads(COOKIE_FILE.read_text())
        cookies = data.get("cookies") or []
        if {c["name"] for c in cookies} >= set(ESSENTIAL_COOKIES):
            return cookies
    if not auto_sync:
        raise NotLoggedIn("本地无可用 Cookie，先跑 weread login")
    return sync_cookies_from_browser()


def cookie_header(cookies: list[dict]) -> str:
    return "; ".join(f"{c['name']}={c['value']}" for c in cookies)


# ---------------------------------------------------------------- client

class WereadAPI:
    def __init__(self, cookies: list[dict] | None = None):
        self.cookies = cookies if cookies is not None else load_cookies()
        self.vid = next((c["value"] for c in self.cookies if c["name"] == "wr_vid"), "")

    def _headers(self, json_body: bool = False) -> dict:
        h = {
            "Cookie": cookie_header(self.cookies),
            "User-Agent": UA,
            "Referer": "https://weread.qq.com/",
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate",
        }
        if json_body:
            h["Content-Type"] = "application/json;charset=UTF-8"
        return h

    def _request(self, url: str, body: dict | None = None, retries: int = 2):
        """带一次自动续会话的重试。"""
        try:
            return self._raw_request(url, body, retries)
        except SessionExpired:
            if not refresh_session():
                raise NotLoggedIn("登录态已过期，且浏览器续期失败；跑 weread login --qr 扫码")
            self.cookies = load_cookies(auto_sync=False)
            self.vid = next((c["value"] for c in self.cookies if c["name"] == "wr_vid"), self.vid)
            return self._raw_request(url, body, retries)

    def _raw_request(self, url: str, body: dict | None = None, retries: int = 2):
        data = json.dumps(body).encode() if body is not None else None
        last = None
        for attempt in range(retries + 1):
            req = urllib.request.Request(url, headers=self._headers(body is not None), data=data)
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    raw = resp.read()
                    enc = resp.headers.get("Content-Encoding", "")
                    if enc == "gzip":
                        raw = gzip.decompress(raw)
                    elif enc == "deflate":
                        raw = zlib.decompress(raw, -zlib.MAX_WBITS)
                    data = json.loads(raw.decode("utf-8"))
                    code = data.get("errCode") if isinstance(data, dict) else None
                    if code in SESSION_ERRORS:
                        raise SessionExpired(f"{SESSION_ERRORS[code]}（errCode {code}）")
                    return data
            except SessionExpired:
                raise
            except urllib.error.HTTPError as exc:
                if exc.code in (401, 403):
                    raise NotLoggedIn(f"接口返回 {exc.code}，登录态已失效，重跑 weread login")
                last = exc
            except Exception as exc:
                last = exc
            time.sleep(1.2 * (attempt + 1))
        raise WereadError(f"请求失败 {url}: {last}")

    # ---- 书架 / 书目

    def shelf(self) -> dict:
        url = ("https://weread.qq.com/web/shelf/sync"
               f"?userVid={self.vid}&synckey=0&lectureSynckey=0")
        data = self._request(url)
        if "books" not in data:
            raise NotLoggedIn("书架接口没返回书目，登录态可能已失效")
        return data

    def book_info(self, book_id: str) -> dict:
        return self._request(f"https://weread.qq.com/web/book/info?bookId={book_id}")

    def chapter_infos(self, book_id: str) -> list[dict]:
        data = self._request("https://weread.qq.com/web/book/chapterInfos",
                             {"bookIds": [str(book_id)], "synckeys": [0], "teenmode": 0})
        items = (data.get("data") or [{}])[0]
        chapters = items.get("updated") or []
        if not chapters:
            raise WereadError(f"取不到章节表：{book_id}（书可能已下架，或无阅读权限）")
        return chapters

    def search(self, keyword: str, count: int = 20) -> list[dict]:
        url = ("https://weread.qq.com/web/search/global?"
               + urllib.parse.urlencode({"keyword": keyword, "maxIdx": 0,
                                         "fragmentSize": 120, "count": count}))
        data = self._request(url)
        books = []
        for group in data.get("books", []) or []:
            info = group.get("bookInfo") or group
            books.append(info)
        return books


def find_in_shelf(shelf: dict, query: str) -> list[dict]:
    """按书名/作者在书架里模糊找书，繁简都能命中。"""
    try:
        import opencc  # 可选
        conv = opencc.OpenCC("t2s").convert
    except Exception:
        conv = lambda s: s  # noqa: E731

    q = conv(query.lower())
    hits = []
    for b in shelf.get("books", []) or []:
        hay = conv((b.get("title", "") + " " + b.get("author", "")).lower())
        if q in hay or query in (b.get("title", "") + b.get("author", "")):
            hits.append(b)
    return hits
