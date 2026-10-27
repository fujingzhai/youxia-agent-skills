"""读 Zen 默认档案里指定域名的 Cookie（明文 cookies.sqlite），返回与 rookiepy 相同的字典格式。"""
from __future__ import annotations

import configparser
import shutil
import sqlite3
import tempfile
from pathlib import Path

ZEN_ROOT = Path.home() / "Library/Application Support/zen"


def read_zen(domains: list[str]) -> list[dict]:
    ini = configparser.ConfigParser()
    ini.read(ZEN_ROOT / "installs.ini")
    rel = next((ini[s]["Default"] for s in ini.sections() if "Default" in ini[s]), None)
    if not rel:
        raise FileNotFoundError("找不到 Zen 的默认档案")
    src = ZEN_ROOT / rel
    with tempfile.TemporaryDirectory() as tmp:  # Zen 开着时库被锁、新改动在 -wal 里，连同 -wal 复制出来读
        for n in ("cookies.sqlite", "cookies.sqlite-wal"):
            if (src / n).exists():
                shutil.copy(src / n, Path(tmp) / n)
        db = sqlite3.connect(Path(tmp) / "cookies.sqlite")
        rows = db.execute("select name, value, host, path, expiry, isSecure, isHttpOnly from moz_cookies"
                          " where originAttributes = ''").fetchall()
        db.close()
    out = []
    for name, value, host, path, exp, secure, http_only in rows:
        h = host.lstrip(".")
        if any(h == d or h.endswith("." + d) for d in domains):
            out.append({"name": name, "value": value, "domain": host, "path": path,
                        "secure": bool(secure), "http_only": bool(http_only),
                        "expires": exp // 1000 if exp > 1e11 else exp})  # 新版 Firefox 以毫秒存
    return out
