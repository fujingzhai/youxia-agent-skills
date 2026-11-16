"""启动带登录档案的浏览器（Camoufox）。只用标准库＋camoufox，由浏览器那侧的解释器导入。

本机有 browser 技能（与本技能同级的 browser/scripts/engine.py，或 WEREAD_ENGINE_DIR 指定）就走它的统一入口：
内核共用、指纹按档案固定。没有就直接用 camoufox（分享版自行 pip install "camoufox[geoip]" 并 camoufox fetch）。
网页主环境里的脚本（挂钩 atob、读 window 变量）都要加 "mw:" 前缀，所以这里总是开 main_world_eval。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


_ENGINE = False  # 未查过


def _engine():
    """本机 browser 技能的 engine（只查一次）；导入本模块时就接上，赶在别处导入 camoufox 之前。"""
    global _ENGINE
    if _ENGINE is not False:
        return _ENGINE
    _ENGINE = _find_engine()
    return _ENGINE


def _find_engine():
    here = os.environ.get("WEREAD_ENGINE_DIR") or str(Path(__file__).resolve().parents[2] / "browser" / "scripts")
    if (Path(here) / "engine.py").is_file():
        if here not in sys.path:
            sys.path.insert(0, here)
        import engine
        return engine
    return None


def persistent_context(profile_dir, headless: bool = True):
    """返回 BrowserContext，用完 ctx.close()。"""
    eng = _engine()
    if eng is not None:
        return eng.persistent_context(profile_dir, headless=headless)
    from camoufox.sync_api import Camoufox
    from playwright.sync_api._context_manager import PlaywrightContextManager
    Path(profile_dir).mkdir(parents=True, exist_ok=True)
    mgr = Camoufox(persistent_context=True, user_data_dir=str(profile_dir), headless=headless,
                   main_world_eval=True, locale=["zh-CN", "zh", "en-US", "en"],
                   os={"darwin": "macos", "win32": "windows"}.get(sys.platform, "linux"))
    ctx = mgr.__enter__()
    original = ctx.close

    def close(*a, **k):
        try:
            original(*a, **k)
        finally:
            PlaywrightContextManager.__exit__(mgr, None, None, None)

    ctx.close = close
    return ctx


def version() -> str:
    """自检用：版本＋内核是否已下载。内核没下时抛错，doctor 会报缺失。"""
    import importlib.metadata as md
    eng = _engine()
    if eng is not None:
        eng.check()
        return f"camoufox {md.version('camoufox')}（browser 技能内核）"
    from camoufox.pkgman import camoufox_path
    try:
        camoufox_path(download_if_missing=False)
    except Exception:
        raise SystemExit("camoufox 内核还没下载：.venv/bin/python -m camoufox fetch")
    return f"camoufox {md.version('camoufox')}"


_engine()
