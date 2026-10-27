#!/usr/bin/env python3
"""按字符挑字体：正文字体缺哪个字，就让它从备用字体里借字形。

中华书局宋体这类古籍字体是**按 Unicode 平面拆成好几个文件**的
（00平面＝基本多文种平面，02平面＝扩展B/C，15平面＝私用区），
每个平面在系统里是一个独立字族。所以「把正文换成中华书局宋体」不是改个名字
就完事——只写 00 平面的话，扩展B 的字照样落到 Word 的默认回退上，等于没换。

这里的做法是：为每个字族建一份码位覆盖表，落盘文字时把一段文字按
「谁能显示这个字」切成若干 run，分别指定字体。常用字仍走正文字体，
只有它缺的字才借用备用字体，观感不受影响。
"""

from __future__ import annotations

import glob
import json
import os
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

STATE = Path(__file__).resolve().parents[1] / ".state"
INDEX_CACHE = STATE / "font_index.json"
COVER_CACHE = STATE / "font_coverage.json"

FONT_DIRS = ["/System/Library/Fonts", "/Library/Fonts",
             os.path.expanduser("~/Library/Fonts")]
FONT_EXT = (".ttf", ".otf", ".ttc", ".otc")

# 成套的古籍字体：正文用哪个、缺字时按顺序找谁借
FONT_PRESETS: dict[str, tuple[str, list[str]]] = {
    "中华书局宋体": ("中华书局宋体00平面",
                 ["中华书局宋体02平面", "中华书局宋体15平面", "中华书局字体16平面"]),
}
# 兜底：上面都没有的字，再找这些通用字体（覆盖面宽且形态接近）
LAST_RESORT = ["Songti SC", "霞鹜文楷", "Hiragino Sans GB", "Heiti SC"]


# 只对汉字做逐字配字体：拉丁、希腊、标点这些 Word 是按 w:ascii 选字体的，
# 改 w:eastAsia 不起作用，硬借反而会把「楷体没有 ṛ」这种误报成缺字。
_CJK_RANGES = ((0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xF900, 0xFAFF),
               (0x20000, 0x323AF))


def _is_cjk_char(cp: int) -> bool:
    return any(lo <= cp <= hi for lo, hi in _CJK_RANGES)


def _is_pua(cp: int) -> bool:
    return (0xE000 <= cp <= 0xF8FF or 0xF0000 <= cp <= 0xFFFFD
            or 0x100000 <= cp <= 0x10FFFD)


def _iter_font_files():
    seen = set()
    for d in FONT_DIRS:
        for path in glob.glob(os.path.join(d, "**", "*"), recursive=True):
            # 扩展名大小写都要认：中华书局那几个是 .TTF
            if path in seen or not path.lower().endswith(FONT_EXT):
                continue
            seen.add(path)
            yield path


def build_index(force: bool = False) -> dict[str, str]:
    """字族名（中英文都收）-> 字体文件路径。"""
    if not force and INDEX_CACHE.exists():
        try:
            return json.loads(INDEX_CACHE.read_text(encoding="utf-8"))
        except Exception:
            pass
    from fontTools.ttLib import TTCollection, TTFont

    index: dict[str, str] = {}
    for path in _iter_font_files():
        try:
            fonts = (TTCollection(path, lazy=True).fonts
                     if path.lower().endswith((".ttc", ".otc"))
                     else [TTFont(path, lazy=True)])
        except Exception:
            continue
        for i, font in enumerate(fonts):
            try:
                records = font["name"].names
            except Exception:
                continue
            for rec in records:
                if rec.nameID != 1:
                    continue
                try:
                    name = rec.toUnicode().strip()
                except Exception:
                    continue
                if name and name not in index:
                    index[name] = f"{path}#{i}"
    STATE.mkdir(exist_ok=True)
    INDEX_CACHE.write_text(json.dumps(index, ensure_ascii=False, indent=0), encoding="utf-8")
    return index


def _ranges(codepoints) -> list[list[int]]:
    out: list[list[int]] = []
    for cp in sorted(codepoints):
        if out and cp == out[-1][1] + 1:
            out[-1][1] = cp
        else:
            out.append([cp, cp])
    return out


def coverage(family: str) -> set[int] | None:
    """某字族能显示的码位集合；字体不在本机则返回 None。"""
    cache = {}
    if COVER_CACHE.exists():
        try:
            cache = json.loads(COVER_CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
    if family in cache:
        if cache[family] is None:
            return None
        return {cp for lo, hi in cache[family] for cp in range(lo, hi + 1)}

    index = build_index()
    ref = index.get(family)
    result = None
    if ref:
        from fontTools.ttLib import TTCollection, TTFont
        path, _, idx = ref.rpartition("#")
        try:
            if path.lower().endswith((".ttc", ".otc")):
                font = TTCollection(path, lazy=True).fonts[int(idx)]
            else:
                font = TTFont(path, lazy=True)
            result = set(font.getBestCmap())
        except Exception:
            result = None
    STATE.mkdir(exist_ok=True)
    cache[family] = _ranges(result) if result else None
    COVER_CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    return result


class FontChain:
    """正文字体 + 备用字体链，按字符决定用谁。"""

    def __init__(self, fallbacks: list[str]):
        self.fallbacks = [f for f in fallbacks if f]
        self._cov: dict[str, set[int] | None] = {}
        self.missing: set[str] = set()      # 本机没装的备用字体
        self.borrowed: dict[str, int] = {}  # 字族 -> 借出的字符数
        self.uncovered: set[int] = set()    # 谁都没有的码位

    def _cov_of(self, family: str) -> set[int] | None:
        if family not in self._cov:
            cov = coverage(family)
            self._cov[family] = cov
            if cov is None:
                self.missing.add(family)
        return self._cov[family]

    def available(self) -> list[str]:
        return [f for f in self.fallbacks if self._cov_of(f)]

    def split(self, text: str, primary: str):
        """把文字切成 [(片段, 该用的字族或 None＝沿用正文字体)]。"""
        if not self.fallbacks:
            return [(text, None)]
        base = self._cov_of(primary)
        if base is None:
            return [(text, None)]

        segments: list[tuple[str, str | None]] = []
        buf: list[str] = []
        cur: str | None = None
        for ch in text:
            cp = ord(ch)
            pick: str | None = None
            # 私用区绝不借字：各家 PUA 是各自的造字约定，借来会渲出别的字。
            # 这类字只能靠 lib/pua.py 的 cmap 等价关系查证，查不到就留空并高亮。
            if _is_cjk_char(cp) and cp not in base and not _is_pua(cp):
                for fam in self.fallbacks:
                    cov = self._cov_of(fam)
                    if cov and cp in cov:
                        pick = fam
                        break
                if pick is None:
                    self.uncovered.add(cp)
            if pick != cur and buf:
                segments.append(("".join(buf), cur))
                buf = []
            cur = pick
            buf.append(ch)
            if pick:
                self.borrowed[pick] = self.borrowed.get(pick, 0) + 1
        if buf:
            segments.append(("".join(buf), cur))
        return segments


def resolve_preset(name: str) -> tuple[str, list[str]]:
    """把「中华书局宋体」这种成套字体展开成 (正文字体, 备用链)。"""
    if name in FONT_PRESETS:
        primary, fallbacks = FONT_PRESETS[name]
        return primary, fallbacks + LAST_RESORT
    return name, list(LAST_RESORT)


def main() -> int:
    import sys
    if len(sys.argv) < 2:
        print("用法：fontcover.py <字族名> [要检查的文字]")
        return 2
    fam = sys.argv[1]
    primary, fallbacks = resolve_preset(fam)
    chain = FontChain(fallbacks)
    cov = coverage(primary)
    print(f"正文字体 {primary}：" + (f"{len(cov)} 个码位" if cov else "本机未安装"))
    print("备用链：")
    for f in fallbacks:
        c = coverage(f)
        print(f"  {'✓' if c else '×'} {f}" + (f"　{len(c)} 个码位" if c else "　未安装"))
    if len(sys.argv) > 2:
        for seg, pick in chain.split(sys.argv[2], primary):
            print(f"  {seg!r} -> {pick or primary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
