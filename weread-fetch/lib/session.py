#!/usr/bin/env python3
"""用浏览器维持微信读书登录态，并把 Cookie 导回 .state/cookies.json。

wr_skey 是短时票据，几十分钟就过期；官方的 /web/login/renewal 只认浏览器
当下那份 wr_rt，从别处抄来的快照续不动。所以会话以 cloakbrowser 的
weread profile 为准：打开一次首页让页面自己续，再把 Cookie 抄出来给
HTTP 接口用。

跑在装有 cloakbrowser 的解释器里（默认即本技能 .venv，见 lib/paths.py）。
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_DIR))
from lib.paths import PROFILE_DIR  # noqa: E402

STATE = SKILL_DIR / ".state"
COOKIE_FILE = STATE / "cookies.json"


def seed_cookies_from_browser() -> list[dict]:
    """本机浏览器（Zen/Chrome/Edge/Brave/Firefox）里如果已登录，拿来给 profile 打底。"""
    try:
        import rookiepy
    except ImportError:
        return []
    from zen_cookies import read_zen  # 本脚本由 cloak 的解释器直接运行，lib 目录在 sys.path 上
    for name in ("zen", "chrome", "edge", "brave", "firefox"):
        reader = read_zen if name == "zen" else getattr(rookiepy, name, None)
        if reader is None:
            continue
        try:
            cookies = reader(["weread.qq.com"]) if name == "zen" else reader(domains=["weread.qq.com"])
        except Exception:
            continue
        if any(c["name"] == "wr_skey" and c["value"] for c in cookies):
            return [{
                "name": c["name"], "value": c["value"],
                "domain": c.get("domain") or ".weread.qq.com",
                "path": c.get("path") or "/",
                "secure": bool(c.get("secure")),
                "httpOnly": bool(c.get("http_only")),
                "expires": int(float(c["expires"])) if c.get("expires") else -1,
                "sameSite": "Lax",
            } for c in cookies]
    return []


def export(ctx) -> list[dict]:
    out = []
    for c in ctx.cookies():
        if "weread.qq.com" in (c.get("domain") or ""):
            out.append({"name": c["name"], "value": c["value"], "domain": c["domain"],
                        "path": c.get("path", "/"), "secure": c.get("secure", False),
                        "http_only": c.get("httpOnly", False),
                        "expires": c.get("expires", -1)})
    return out


def probe(page) -> dict:
    """在页面里同源请求书架，顺带确认登录态。"""
    return page.evaluate("""async () => {
      try {
        const vid = (document.cookie.match(/wr_vid=(\\d+)/) || [])[1] || '';
        const r = await fetch(`/web/shelf/sync?userVid=${vid}&synckey=0&lectureSynckey=0`,
                              {credentials: 'include'});
        const j = await r.json();
        return {ok: Array.isArray(j.books), errCode: j.errCode || 0,
                errMsg: j.errMsg || '', count: (j.books || []).length, vid};
      } catch (e) { return {ok: false, errMsg: String(e)}; }
    }""")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--headed", action="store_true", help="开窗口，供人工扫码登录")
    ap.add_argument("--wait", type=int, default=0, help="开窗后等待秒数，留给扫码")
    args = ap.parse_args()

    import cloakbrowser
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    ctx = cloakbrowser.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR), headless=not args.headed)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    try:
        page.goto("https://weread.qq.com/", wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3500)
        status = probe(page)

        if not status.get("ok"):
            seed = seed_cookies_from_browser()
            if seed:
                try:
                    ctx.add_cookies(seed)
                    page.goto("https://weread.qq.com/", wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(3500)
                    status = probe(page)
                except Exception:
                    pass

        if not status.get("ok") and args.wait:
            print(f"扫码       浏览器已打开，请用微信扫码登录（最多等 {args.wait} 秒）", file=sys.stderr)
            deadline = time.time() + args.wait
            while time.time() < deadline:
                page.wait_for_timeout(3000)
                status = probe(page)
                if status.get("ok"):
                    break

        if status.get("ok"):
            STATE.mkdir(exist_ok=True)
            payload = {"source": "cloakbrowser", "synced": int(time.time()),
                       "cookies": export(ctx)}
            COOKIE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=1))
            COOKIE_FILE.chmod(0o600)
        print(json.dumps(status, ensure_ascii=False))
        return 0 if status.get("ok") else 4
    finally:
        try:
            ctx.close()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
