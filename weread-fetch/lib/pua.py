#!/usr/bin/env python3
"""把私用区（PUA）码位还原成真正的 Unicode 字符。

有些书把生僻字编在私用区，靠嵌入字体显示。微信读书只发常用字体子集，
带 PUA 字形的扩展面子集（ext15 那种）一概不给，于是这些字到手就是个孤零零
的码位，任何阅读器和 Word 里都是空白。

还原办法不是比对字形轮廓，而是查 cmap 的等价关系：不少中文字体会把
「私用区码位」和「该字真正的 Unicode 码位」映到同一个字形 id。命中即是确证。
实例：Hiragino Sans GB 里 U+4D19（䴙）与 U+E862 同为 cid22124，
于是《治史三书》「▮鹈泉」还原为「䴙鹈泉」。

只认 CJK 区间的结果，且共享同一字形的码位不能太多（多了说明那是个兜底字形）。
"""

from __future__ import annotations

import glob
import json
import os
from pathlib import Path

STATE = Path(__file__).resolve().parents[1] / ".state"
CACHE = STATE / "pua_map.json"
MANUAL = STATE / "pua_manual.json"      # 全书共用一份，按 bookId 归档

FONT_DIRS = ["/System/Library/Fonts", "/Library/Fonts",
             os.path.expanduser("~/Library/Fonts")]

# 认得的真字区间：CJK 基本区、扩展 A、兼容汉字、扩展 B 及以后
CJK_RANGES = ((0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xF900, 0xFAFF),
              (0x20000, 0x323AF))

BAD_GLYPHS = {".notdef", "space", "nonmarkingreturn", "glyph0", None, ""}


def is_pua(cp: int) -> bool:
    return (0xE000 <= cp <= 0xF8FF or 0xF0000 <= cp <= 0xFFFFD
            or 0x100000 <= cp <= 0x10FFFD)


def _is_cjk(cp: int) -> bool:
    return any(lo <= cp <= hi for lo, hi in CJK_RANGES)


_FONT_EXT = (".ttf", ".otf", ".ttc", ".otc")


def _iter_fonts():
    """遍历本机字体。扩展名要大小写都认——中华书局宋体那几个是 .TTF，
    只匹配小写会整套漏掉，而它们恰恰是私用区字形最多的。"""
    seen = set()
    for d in FONT_DIRS:
        for path in glob.glob(os.path.join(d, "**", "*"), recursive=True):
            if path in seen or not path.lower().endswith(_FONT_EXT):
                continue
            seen.add(path)
            yield path


def _cmaps(path: str):
    """产出 (字体标识, cmap)。ttc 里每个子字体都算一个。"""
    from fontTools.ttLib import TTCollection, TTFont
    try:
        if path.lower().endswith((".ttc", ".otc")):
            coll = TTCollection(path, lazy=True)
            for i, font in enumerate(coll.fonts):
                yield f"{os.path.basename(path)}#{i}", font.getBestCmap()
        else:
            yield os.path.basename(path), TTFont(path, lazy=True).getBestCmap()
    except Exception:
        return


def resolve(codepoints: set[int], use_cache: bool = True) -> dict[int, dict]:
    """返回 {pua码位: {"to": 真码位, "char": 字符, "font": 证据字体}}。"""
    wanted = {cp for cp in codepoints if is_pua(cp)}
    if not wanted:
        return {}

    cache: dict = {}
    if use_cache and CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
    out = {cp: cache[str(cp)] for cp in wanted if str(cp) in cache}
    todo = wanted - {cp for cp in out}
    if not todo:
        return out

    for path in _iter_fonts():
        if not todo:
            break
        for tag, cmap in _cmaps(path):
            if not cmap:
                continue
            # 先反向索引：字形名 -> 指向它的所有码位
            hit = {cp: cmap[cp] for cp in todo if cp in cmap}
            if not hit:
                continue
            reverse: dict = {}
            for cp, gname in cmap.items():
                reverse.setdefault(gname, []).append(cp)
            for cp, gname in list(hit.items()):
                if gname in BAD_GLYPHS:
                    continue
                group = reverse.get(gname, [])
                if len(group) > 3:      # 一堆码位共用，多半是兜底字形
                    continue
                real = [c for c in group if _is_cjk(c)]
                if len(real) != 1:
                    continue
                out[cp] = {"to": real[0], "char": chr(real[0]), "font": tag}
                todo.discard(cp)

    if use_cache:
        STATE.mkdir(exist_ok=True)
        cache.update({str(k): v for k, v in out.items()})
        CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return out


def load_manual(book_id: str) -> dict[int, dict]:
    """读人工判读表。

    机器只认 cmap 等价这一种证据；私用区码位若用的是出版社自编的约定，本机没有
    任何字体能作证，机器就该拒绝还原。这时由人按上下文判读，结论记在这里。

    **一份文件管所有书**（`.state/pua_manual.json`，按 bookId 归档），不放任务目录——
    任务目录是临时区，清掉判读就没了；也不放成品目录，免得每本书多一个文件要管。

        {"45699582": {"E19A": {"char": "蘋", "note": "晏几道《小山词》自序…"},
                      "E1A3": "巘"}}
    """
    if not MANUAL.exists():
        return {}
    try:
        raw = json.loads(MANUAL.read_text(encoding="utf-8")).get(str(book_id)) or {}
    except Exception:
        return {}
    out: dict[int, dict] = {}
    for key, val in raw.items():
        try:
            cp = int(str(key).lstrip("Uu+"), 16)
        except ValueError:
            continue
        item = {"char": val} if isinstance(val, str) else dict(val or {})
        if item.get("char"):
            out[cp] = item
    return out


def save_manual(book_id: str, entries: dict) -> Path:
    """写入某本书的判读表，保留其他书的记录。"""
    data = {}
    if MANUAL.exists():
        try:
            data = json.loads(MANUAL.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    data[str(book_id)] = entries
    STATE.mkdir(exist_ok=True)
    MANUAL.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return MANUAL


def scan_text(text: str) -> set[int]:
    return {ord(ch) for ch in text if is_pua(ord(ch))}


def main() -> int:
    import sys
    cps = set()
    for arg in sys.argv[1:]:
        if arg.upper().startswith("U+"):
            cps.add(int(arg[2:], 16))
        else:
            cps |= scan_text(Path(arg).read_text(encoding="utf-8")) if Path(arg).exists() \
                else scan_text(arg)
    if not cps:
        print("用法：pua.py U+E862 | pua.py <文件> | pua.py '<含私用区字的文本>'")
        return 2
    res = resolve(cps)
    for cp in sorted(cps):
        r = res.get(cp)
        if r:
            print(f"  U+{cp:04X} → {r['char']} (U+{r['to']:04X})　证据：{r['font']}")
        else:
            print(f"  U+{cp:04X} → 未能还原")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
