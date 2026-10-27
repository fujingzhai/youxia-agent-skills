"""本机路径与浏览器运行时。都可用环境变量覆盖，入口脚本会先读技能目录下的 .env。

只用标准库：capture.py、session.py 由浏览器那侧的解释器直接运行，也要能导入。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]


def _env_path(name: str, default: Path) -> Path:
    value = os.environ.get(name)
    return Path(os.path.expandvars(value)).expanduser() if value else default


# 跑 capture.py / session.py 的解释器，须装有 cloakbrowser；默认就是本技能的 .venv
BROWSER_PY = _env_path("WEREAD_BROWSER_PY", Path(sys.executable))
# 浏览器登录态（持久化 profile）
PROFILE_DIR = _env_path("WEREAD_PROFILE_DIR", SKILL_DIR / "profiles" / "weread")
# 任务目录（每本书一个子目录，可丢的缓存）
WORK_ROOT = _env_path("WEREAD_WORK_DIR", SKILL_DIR / "work")
# 成品 docx / epub
OUTPUT_DIR = _env_path("WEREAD_OUTPUT_DIR", Path.home() / "Downloads" / "微信读书")
