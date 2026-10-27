#!/usr/bin/env python3
"""在浏览器里逐章取回微信读书的原始 EPUB XHTML。

为什么走浏览器而不是直接打接口：
  /web/book/chapter/e_* 的响应是「32 位 MD5 + 一段 base64 切片」，切片边界
  不按 4 字节对齐，且需要客户端算签名。自己拼块极易错位（实测会把「时代」
  解成「时c」这种坏字），而阅读器自己会把所有块拼好后调一次 window.atob。
  于是这里只做一件事：挂钩 window.atob，把它还原出来的整章 XHTML 抄走。

本文件跑在装有 cloakbrowser 的解释器里（默认即本技能 .venv，见 lib/paths.py），只用标准库 + cloakbrowser。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lib.ids import reader_url  # noqa: E402
from lib.paths import PROFILE_DIR  # noqa: E402

# 把 atob 的产物（latin-1 二进制串）还原成 UTF-8 文本存起来
HOOK = r"""
window.__wr = {items: [], t0: Date.now()};
(function () {
  const native = window.atob;
  window.atob = function (s) {
    const raw = native(s);
    try {
      const bytes = new Uint8Array(raw.length);
      for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
      window.__wr.items.push(new TextDecoder('utf-8').decode(bytes));
    } catch (e) { /* 不影响页面本身 */ }
    return raw;
  };
})();
"""

IS_XHTML = re.compile(r"<\s*(html|body|p|h[1-6]|div)\b", re.I)


def looks_like_body(text: str) -> bool:
    head = text[:600].lstrip()
    if head.startswith("<?xml") or "<html" in head.lower():
        return True
    return bool(IS_XHTML.search(head)) and "{" not in head[:80]


def looks_like_css(text: str) -> bool:
    head = text[:400]
    return ("{" in head and "}" in text[:2000]) and not looks_like_body(text)


class Capturer:
    def __init__(self, headless: bool = True, cookies: list[dict] | None = None):
        import cloakbrowser
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        self.ctx = cloakbrowser.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR), headless=headless)
        if cookies:
            try:
                self.ctx.add_cookies(cookies)
            except Exception as exc:
                print(f"警告       Cookie 注入失败：{exc}", file=sys.stderr)
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        self.page.add_init_script(HOOK)

    def close(self):
        try:
            self.ctx.close()
        except Exception:
            pass

    def grab(self, book_id: str, chapter_uid: int, timeout_s: float = 30.0):
        """打开一章，返回 (正文 XHTML 列表, CSS 列表, 页面标题)。"""
        url = reader_url(book_id, chapter_uid)
        self.page.goto(url, wait_until="domcontentloaded", timeout=60000)

        deadline = time.time() + timeout_s
        body: list[str] = []
        while time.time() < deadline:
            self.page.wait_for_timeout(400)
            items = self.page.evaluate("() => window.__wr ? window.__wr.items : []")
            body = [t for t in items if looks_like_body(t)]
            if body:
                break
        # 正文到手后再等一拍，接住可能晚到的样式表/续块
        self.page.wait_for_timeout(1200)
        items = self.page.evaluate("() => window.__wr ? window.__wr.items : []")
        body = [t for t in items if looks_like_body(t)]
        css = [t for t in items if looks_like_css(t)]
        return body, css, (self.page.title() or "")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--work-dir", required=True, type=Path, help="任务目录（含 book.json）")
    ap.add_argument("--headed", action="store_true", help="显示窗口，便于排查")
    ap.add_argument("--delay", type=float, default=0.8, help="章间隔（秒）")
    ap.add_argument("--limit", type=int, default=0, help="只抓前 N 章（调试用）")
    args = ap.parse_args()

    work = args.work_dir
    meta = json.loads((work / "book.json").read_text(encoding="utf-8"))
    book_id = str(meta["bookId"])
    chapters = meta["chapters"]
    select = meta.get("select")
    if select:
        keep = set(select)
        chapters = [c for c in chapters if c["chapterUid"] in keep]
    if args.limit:
        chapters = chapters[:args.limit]

    xdir = work / "xhtml"; xdir.mkdir(parents=True, exist_ok=True)
    cdir = work / "css"; cdir.mkdir(parents=True, exist_ok=True)
    prog_path = work / "progress.json"
    progress = json.loads(prog_path.read_text()) if prog_path.exists() else {}

    cookie_file = Path(__file__).resolve().parents[1] / ".state" / "cookies.json"
    cookies = None
    if cookie_file.exists():
        data = json.loads(cookie_file.read_text())
        cookies = []
        for c in data.get("cookies", []):
            exp = c.get("expires")
            cookies.append({
                "name": c["name"], "value": c["value"],
                "domain": c.get("domain") or ".weread.qq.com",
                "path": c.get("path") or "/",
                "secure": bool(c.get("secure")),
                "httpOnly": bool(c.get("http_only") or c.get("httpOnly")),
                "expires": int(float(exp)) if exp else -1,
                "sameSite": "Lax",
            })

    cap = Capturer(headless=not args.headed, cookies=cookies)
    todo = [c for c in chapters if progress.get(str(c["chapterUid"]), {}).get("ok") is not True]
    print(f"章节       共 {len(chapters)} 章，待抓 {len(todo)} 章")

    failed = 0
    try:
        for n, ch in enumerate(todo, 1):
            uid = ch["chapterUid"]
            idx = ch.get("chapterIdx", uid)
            title = (ch.get("title") or "").strip()
            stem = f"{idx:04d}_{uid}"
            body, css, page_title = [], [], ""
            for attempt in range(2):
                try:
                    body, css, page_title = cap.grab(book_id, uid)
                except Exception as exc:
                    print(f"  ! {title[:24]} 打开失败：{type(exc).__name__} {exc}"[:160], file=sys.stderr)
                if body:
                    break
                time.sleep(2.0)

            if not body:
                failed += 1
                progress[str(uid)] = {"ok": False, "title": title, "reason": "未捕获到正文"}
                print(f"  [{n}/{len(todo)}] 失败 {title[:30]}")
            else:
                merged = "\n".join(body)
                (xdir / f"{stem}.xhtml").write_text(merged, encoding="utf-8")
                if css:
                    (cdir / f"{stem}.css").write_text("\n".join(css), encoding="utf-8")
                progress[str(uid)] = {"ok": True, "title": title, "file": f"xhtml/{stem}.xhtml",
                                      "chars": len(merged), "parts": len(body)}
                print(f"  [{n}/{len(todo)}] {title[:30]:<32} {len(merged):>7} 字符")
            prog_path.write_text(json.dumps(progress, ensure_ascii=False, indent=1), encoding="utf-8")
            time.sleep(args.delay)
    finally:
        cap.close()

    print(f"完成       成功 {sum(1 for v in progress.values() if v.get('ok'))} 章，失败 {failed} 章")
    return 1 if failed and failed == len(todo) else 0


if __name__ == "__main__":
    raise SystemExit(main())
