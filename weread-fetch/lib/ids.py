#!/usr/bin/env python3
"""bookId / chapterUid 与网页阅读器 URL 之间的换算。

微信读书网页版的阅读地址形如 /web/reader/<bookHash>k<chapterHash>，
两段哈希用同一个算法从 bookId、chapterUid 算出。算法已用书架上 7 本
（含 CB_ 开头的非数字 bookId）逐一比对验证。
"""

from __future__ import annotations

import hashlib

_ALPHABET_TAG_DIGIT = "3"
_ALPHABET_TAG_TEXT = "4"


def encode_id(raw_id: str) -> str:
    """把 bookId 或 chapterUid 编成阅读器 URL 里的那段哈希。"""
    raw_id = str(raw_id)
    digest = hashlib.md5(raw_id.encode()).hexdigest()
    out = digest[:3]

    if raw_id.isdigit():
        # 纯数字按 9 位一组转十六进制
        parts = [format(int(raw_id[i:i + 9]), "x") for i in range(0, len(raw_id), 9)]
        tag = _ALPHABET_TAG_DIGIT
    else:
        parts = ["".join(format(ord(ch), "x") for ch in raw_id)]
        tag = _ALPHABET_TAG_TEXT

    out += tag + "2" + digest[-2:]
    for index, part in enumerate(parts):
        length = format(len(part), "x")
        out += length.rjust(2, "0")
        out += part
        if index < len(parts) - 1:
            out += "g"

    if len(out) < 20:
        out += digest[:20 - len(out)]
    return out + hashlib.md5(out.encode()).hexdigest()[:3]


def reader_url(book_id: str, chapter_uid: int | str | None = None) -> str:
    """拼出网页阅读器地址；不给 chapter_uid 就是书的首页。"""
    url = "https://weread.qq.com/web/reader/" + encode_id(book_id)
    if chapter_uid is not None:
        url += "k" + encode_id(str(chapter_uid))
    return url
