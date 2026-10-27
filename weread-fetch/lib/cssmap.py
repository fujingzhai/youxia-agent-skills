#!/usr/bin/env python3
"""把每章自带的 EPUB 样式表解析成「元素 -> 排版意图」。

不同出版社的 class 名完全不一样（kindle-cn-kai / preface-text / content_101 …），
硬编码类名换本书就废。但微信读书每章都会连原书 CSS 一起下发，所以这里改成
读 CSS 求实际样式，类名只在 CSS 缺失时兜底。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace

# ----------------------------------------------------------------- 字体族归类

_FAMILY_PATTERNS = [
    ("kai", re.compile(r"楷|kaiti|kai\b|stkaiti|hyqts", re.I)),
    ("fangsong", re.compile(r"仿宋|fangsong|fzfs|stfangsong|hyfs", re.I)),
    ("hei", re.compile(r"黑体|黑體|[兰蘭旗悠冬青大标准]黑|黑$|heiti|simhei|pingfang|"
                       r"microsoft\s*yahei|yahei|sourcehansans|notosans(cjk|sc)|"
                       r"hiragino\s*sans|sans-serif", re.I)),
    ("song", re.compile(r"宋|songti|simsun|stsong|mincho|sourcehanserif|"
                        r"notoserif|serif", re.I)),
]

# 归一后的族名 -> (中文字体, 西文字体)
FONT_TABLE = {
    "song": ("宋体", "Times New Roman"),
    "kai": ("楷体", "Times New Roman"),
    "fangsong": ("仿宋", "Times New Roman"),
    "hei": ("黑体", "Arial"),
}


def classify_family(value: str) -> str | None:
    for name, pat in _FAMILY_PATTERNS:
        if pat.search(value):
            return name
    return None


# ----------------------------------------------------------------- CSS 解析

_COMMENT = re.compile(r"/\*.*?\*/", re.S)
_RULE = re.compile(r"([^{}]+)\{([^{}]*)\}", re.S)

_ABS_SIZE = {"xx-small": 0.6, "x-small": 0.75, "small": 0.85, "medium": 1.0,
             "large": 1.15, "x-large": 1.4, "xx-large": 1.7}


def _parse_len_em(value: str) -> float | None:
    """把长度折算成相对正文的倍数；px 按 16px = 1em 折。"""
    value = value.strip().lower()
    if value in _ABS_SIZE:
        return _ABS_SIZE[value]
    m = re.match(r"^(-?[\d.]+)\s*(em|rem|%|px|pt)?$", value)
    if not m:
        return None
    try:
        num = float(m.group(1))
    except ValueError:
        return None
    unit = m.group(2) or ""
    if unit in ("em", "rem", ""):
        return num
    if unit == "%":
        return num / 100.0
    if unit == "px":
        return num / 16.0
    if unit == "pt":
        return num / 12.0
    return None


@dataclass
class Rule:
    parts: list                 # [(tag, (class, …)), …]，从祖先到自身
    specificity: int
    order: int
    decls: dict


_TOKEN = re.compile(r"^([a-zA-Z][\w-]*)?((?:\.[\w-]+)*)$")


def _parse_selector(sel: str) -> list | None:
    """只认「标签 + 类名」的后代选择器；带属性、伪类、通配的一律放弃。"""
    sel = sel.strip().replace(">", " ")
    if not sel or sel.startswith("@") or any(ch in sel for ch in "[]:*+~"):
        return None
    parts = []
    for token in sel.split():
        m = _TOKEN.match(token)
        if not m:
            return None
        tag = m.group(1).lower() if m.group(1) else None
        classes = tuple(c for c in (m.group(2) or "").split(".") if c)
        if not tag and not classes:
            return None
        parts.append((tag, classes))
    return parts or None


def _part_matches(el, part) -> bool:
    tag, classes = part
    name = getattr(el, "name", None)
    if not name:
        return False
    if tag and name.lower() != tag:
        return False
    if classes and not set(classes) <= set(el.get("class") or []):
        return False
    return True


_FONT_FACE = re.compile(r"@font-face\s*\{([^{}]*)\}", re.S)


class Stylesheet:
    """够用就好的 CSS 索引。

    早先图省事，把后代选择器截成最后一节当裸标签用，结果 `.cover p {text-indent:0}`
    这种规则会把全书 <p> 的首行缩进抹平（在《东汉时代的政治与社会》上实测到）。
    现在按祖先链逐级核对。
    """

    def __init__(self, css_text: str = ""):
        self.rules: list[Rule] = []
        self.aliases: dict[str, str] = {}   # @font-face 自定义族名 -> 归一族名
        self._order = 0
        self.parse(css_text)

    def _collect_font_faces(self, css_text: str) -> None:
        """原书常把字体声明成 "ht"/"kt"/"fs" 这类代号，光看名字认不出。
        但 @font-face 的 src 里带着真名（SourceHanSans、方正兰亭黑、汉仪仿宋…），
        据此建别名表，否则标题会被当成宋体，白丢一层字体层次。"""
        for body in _FONT_FACE.findall(css_text):
            decls = {}
            for item in body.split(";"):
                if ":" in item:
                    prop, _, val = item.partition(":")
                    decls[prop.strip().lower()] = val.strip()
            name = decls.get("font-family", "").strip().strip("\'\"").lower()
            if not name:
                continue
            fam = classify_family(name) or classify_family(decls.get("src", ""))
            if fam:
                self.aliases.setdefault(name, fam)

    def parse(self, css_text: str) -> None:
        if not css_text:
            return
        css_text = _COMMENT.sub(" ", css_text)
        self._collect_font_faces(css_text)
        for sel_group, body in _RULE.findall(css_text):
            decls = {}
            for item in body.split(";"):
                if ":" not in item:
                    continue
                prop, _, val = item.partition(":")
                decls[prop.strip().lower()] = val.strip()
            if not decls:
                continue
            for sel in sel_group.split(","):
                parts = _parse_selector(sel)
                if not parts:
                    continue
                spec = sum(10 * len(cls) + (1 if tag else 0) for tag, cls in parts)
                self.rules.append(Rule(parts, spec, self._order, decls))
                self._order += 1

    def match(self, el) -> dict:
        hits = []
        for rule in self.rules:
            if not _part_matches(el, rule.parts[-1]):
                continue
            if len(rule.parts) > 1:
                index = len(rule.parts) - 2
                node = getattr(el, "parent", None)
                while node is not None and index >= 0:
                    if _part_matches(node, rule.parts[index]):
                        index -= 1
                    node = getattr(node, "parent", None)
                if index >= 0:
                    continue
            hits.append(rule)
        merged: dict = {}
        for rule in sorted(hits, key=lambda r: (r.specificity, r.order)):
            merged.update(rule.decls)
        return merged


# ----------------------------------------------------------------- 计算样式

@dataclass
class Computed:
    size_em: float = 1.0            # 相对正文字号的倍数
    family: str = "song"            # song / kai / fangsong / hei
    italic: bool = False
    bold: bool = False
    underline: bool = False
    align: str | None = None        # left / center / right / justify
    indent_em: float | None = None  # 首行缩进（字符数）
    left_em: float = 0.0            # 整段左缩进（引文常用 margin-left）
    height_em: float | None = None  # 元素高度，行内生僻字图靠它定尺寸
    space_before_em: float = 0.0
    space_after_em: float = 0.0
    superscript: bool = False
    subscript: bool = False
    color: str | None = None
    family_explicit: bool = False   # 字族是否由 CSS 明确指定（标题据此决定要不要盖掉样式字体）
    size_explicit: bool = False     # 字号同理；原书 CSS 常不给标题定字号
    inherited_from_tag: str = ""
    extra_classes: list[str] = field(default_factory=list)

    def child(self) -> "Computed":
        """继承给子元素：块级几何属性不继承，字体相关继承。"""
        return replace(self, align=self.align, indent_em=None, left_em=0.0,
                       height_em=None, space_before_em=0.0, space_after_em=0.0,
                       extra_classes=[])


_ALIGN_OK = {"left", "center", "right", "justify"}


def resolve_family(value: str, aliases: dict | None = None) -> str | None:
    fam = classify_family(value)
    if fam:
        return fam
    for token in value.split(","):
        key = token.strip().strip("'\"").lower()
        if aliases and key in aliases:
            return aliases[key]
    return None


def apply_declarations(base: Computed, decls: dict, tag: str, classes: list[str],
                       aliases: dict | None = None) -> Computed:
    st = replace(base, extra_classes=list(classes))
    st.inherited_from_tag = tag

    if "font-size" in decls:
        v = _parse_len_em(decls["font-size"])
        if v and 0.2 < v < 4:
            st.size_em = base.size_em * v
            if abs(v - 1.0) > 1e-6:
                st.size_explicit = True

    if "height" in decls:
        v = _parse_len_em(decls["height"])
        if v is not None and 0 < v < 12:
            st.height_em = v

    if "font-family" in decls:
        fam = resolve_family(decls["font-family"], aliases)
        if fam:
            st.family = fam
            st.family_explicit = True

    if "font-style" in decls:
        st.italic = decls["font-style"].strip().lower() in ("italic", "oblique")
    if "font-weight" in decls:
        w = decls["font-weight"].strip().lower()
        st.bold = w in ("bold", "bolder") or (w.isdigit() and int(w) >= 600)
    if "text-decoration" in decls:
        st.underline = "underline" in decls["text-decoration"].lower()

    if "vertical-align" in decls:
        # 出版社常用 span.super / span.sub 排上下标（注号、指数、化学式），不用 <sup>。
        # 不认的话「10⁸」会印成「108」。Word 的上下标自己会缩小，CSS 里配套的
        # font-size: 0.7em 再叠上去就小得看不清，所以字号退回父级。
        va = decls["vertical-align"].strip().lower()
        if va in ("super", "sub"):
            st.superscript, st.subscript = va == "super", va == "sub"
            st.size_em = base.size_em

    if "text-align" in decls:
        a = decls["text-align"].strip().lower()
        if a in _ALIGN_OK:
            st.align = a

    if "text-indent" in decls:
        v = _parse_len_em(decls["text-indent"])
        if v is not None:
            st.indent_em = v

    for prop in ("margin-left", "padding-left"):
        if prop in decls:
            v = _parse_len_em(decls[prop])
            if v is not None and 0 < v < 12:
                st.left_em = max(st.left_em, v)

    for prop, attr in (("margin-top", "space_before_em"), ("margin-bottom", "space_after_em")):
        if prop in decls:
            v = _parse_len_em(decls[prop])
            if v is not None and 0 <= v < 6:
                setattr(st, attr, v)
    if "margin" in decls:
        parts = decls["margin"].split()
        vals = [_parse_len_em(p) for p in parts]
        vals = [v if v is not None else 0.0 for v in vals]
        if len(vals) == 1:
            top = bot = vals[0]
        elif len(vals) == 2:
            top = bot = vals[0]
        elif len(vals) >= 3:
            top, bot = vals[0], vals[2]
        else:
            top = bot = 0.0
        if 0 <= top < 6:
            st.space_before_em = top
        if 0 <= bot < 6:
            st.space_after_em = bot

    if "color" in decls:
        c = decls["color"].strip()
        if re.match(r"^#[0-9a-fA-F]{6}$", c) and c.lower() not in ("#000000", "#000"):
            st.color = c
    return st


# 没有 CSS 时按标签兜底
TAG_FALLBACK = {
    "b": {"font-weight": "bold"}, "strong": {"font-weight": "bold"},
    "i": {"font-style": "italic"}, "em": {"font-style": "italic"},
    "u": {"text-decoration": "underline"},
    "small": {"font-size": "0.85em"},
    "blockquote": {"text-indent": "0", "margin-top": "0.5em", "margin-bottom": "0.5em"},
}


def compute(el, sheet: Stylesheet, base: Computed) -> Computed:
    tag = el.name.lower()
    classes = [c for c in (el.get("class") or [])]
    decls = dict(TAG_FALLBACK.get(tag, {}))
    decls.update(sheet.match(el))

    style_attr = el.get("style")
    if style_attr:
        for item in style_attr.split(";"):
            if ":" in item:
                p, _, v = item.partition(":")
                decls[p.strip().lower()] = v.strip()

    st = apply_declarations(base, decls, tag, classes, sheet.aliases)
    if tag == "sup":
        st.superscript = True
    elif tag == "sub":
        st.subscript = True
    return st
