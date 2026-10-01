#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2docx.py — Markdown 文稿 → 学术版式 DOCX。

宋体正文、首行缩进两字符、Word 真脚注、导航窗格大纲。

标题采用「正文段落 + 直接 outlineLvl」：外观完全由 run 的直接格式控制，
不套 Heading 样式。这样既能进 Word 导航窗格跳转，事后又能在
「段落 → 大纲级别」里自由升降，不会被标题样式锁住外观。

行内写法
  **粗体**                                          加粗
  ==高亮==                                          Word 黄色高亮
  ^[脚注内容]                                       真脚注
  <sup>(备注)</sup>                                 真脚注（旧稿兼容）
  <span data-type="inline-memo"
        data-inline-memo-content="注">正文</span>   正文 + 真脚注（旧稿兼容）
  <span style="color: #d23f31;">文字</span>         彩色字
  `代码`                                            等宽
  [文字](url)                                       超链接
  ![alt](url)                                       图片（需 embed_images=True）

块级写法
  # 至 ######            大纲 1-6 级
  > …（文首连续若干行）  元信息块，居中楷体灰
  > …（正文中）          引用块
  - / * / 1.             列表
  | a | b |              表格
  [00:15:30] 独占一行    时间戳标记行
  ---                    跳过
"""

from __future__ import annotations

import re
from pathlib import Path

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor

import footnotes as fn

FONT_MAIN = "宋体"
FONT_SUB = "楷体"
FONT_MONO = "Courier New"

SIZE_TITLE = 20
SIZE_H1 = 15
SIZE_H2 = 13
SIZE_H3 = 12
SIZE_BODY = 12
SIZE_AUX = 10.5
SIZE_MARK = 10
SIZE_FOOTNOTE = 9

GRAY_META = RGBColor(80, 80, 80)
GRAY_QUOTE = RGBColor(60, 60, 60)
GRAY_MARK = RGBColor(100, 100, 100)
BLUE_LINK = RGBColor(0x1F, 0x4E, 0x79)

HIGHLIGHT_YELLOW = "yellow"


# ---------------------------------------------------------------- 底层格式

def set_run_font(run, font_name=FONT_MAIN, size_pt=SIZE_BODY,
                 bold=False, italic=False, color=None):
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font_name)
    rFonts.set(qn("w:hint"), "eastAsia")


def set_highlight(run, color=HIGHLIGHT_YELLOW):
    rPr = run._element.get_or_add_rPr()
    el = rPr.find(qn("w:highlight"))
    if el is None:
        el = OxmlElement("w:highlight")
        rPr.append(el)
    el.set(qn("w:val"), color)


def set_outline_level(paragraph, level):
    """标大纲级别（level 从 1 起），不套标题样式。"""
    if not 1 <= int(level) <= 6:
        raise ValueError(f"不支持的大纲级别：{level}")
    pPr = paragraph._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:outlineLvl")):
        pPr.remove(old)
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(int(level) - 1))
    pPr.append(el)
    paragraph.paragraph_format.keep_with_next = True


def verify_outline(path, expected):
    """从最终 DOCX 回读逐条核对标题文字和级别，避免仅有字体外观。"""
    saved = docx.Document(str(path))
    actual = []
    for p in saved.paragraphs:
        levels = p._p.xpath("./w:pPr/w:outlineLvl/@w:val")
        if levels:
            actual.append((p.text, int(levels[0]) + 1))
    if actual != expected:
        raise ValueError(f"DOCX 大纲校验失败：预期 {expected!r}，实际 {actual!r}")
    return len(actual)


def set_first_line_indent(paragraph, chars=200, dxa=480):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars))
    ind.set(qn("w:firstLine"), str(dxa))


def set_spacing(paragraph, before_pt=0, after_pt=0, line_spacing=1.35):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    pf.line_spacing = line_spacing


def add_hyperlink(paragraph, url, text, size_pt=SIZE_BODY):
    """插入真超链接（python-docx 没有现成接口，手工建关系项）。"""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run = paragraph.add_run(text)
    set_run_font(run, size_pt=size_pt, color=BLUE_LINK)
    run.underline = True
    link.append(run._element)
    paragraph._p.append(link)
    return run


def create_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.18)
        s.right_margin = Inches(1.18)
    return doc


def content_width(doc):
    s = doc.sections[0]
    return Emu(s.page_width - s.left_margin - s.right_margin)


# ---------------------------------------------------------------- 行内解析

INLINE_RE = re.compile(
    r"""
    (?P<memo><span[^>]*data-type="inline-memo"[^>]*>.*?</span>)
  | (?P<color><span[^>]*style="[^"]*color\s*:[^"]*"[^>]*>.*?</span>)
  | (?P<image>!\[(?P<alt>[^\]]*)\]\((?P<isrc>[^)\s]+)[^)]*\))
  | (?P<link>(?<!!)\[(?P<ltext>[^\]]+)\]\((?P<lurl>https?://[^)\s]+)[^)]*\))
  | (?P<fnote>\^\[(?P<ftext>[^\]]+)\])
  | (?P<sup><sup>[（(](?P<suptext>.+?)[）)]</sup>)
  | (?P<mark>==(?P<marktext>.+?)==)
  | (?P<bold>\*\*(?P<boldtext>.+?)\*\*)
  | (?P<code>`(?P<codetext>[^`]+)`)
    """,
    re.X | re.S,
)

MEMO_CONTENT_RE = re.compile(r'data-inline-memo-content="(?P<c>[^"]*)"')
COLOR_VAL_RE = re.compile(r"color\s*:\s*(?P<v>#[0-9a-fA-F]{3,6}|[a-zA-Z]+)")
SPAN_INNER_RE = re.compile(r"^<span[^>]*>(?P<inner>.*)</span>$", re.S)
TAG_RE = re.compile(r"<[^>]+>")

NAMED_COLORS = {
    "red": RGBColor(0xD2, 0x3F, 0x31),
    "blue": RGBColor(0x1F, 0x4E, 0x79),
    "green": RGBColor(0x2E, 0x7D, 0x32),
    "orange": RGBColor(0xE6, 0x7E, 0x22),
}


def _parse_color(text):
    m = COLOR_VAL_RE.search(text)
    if not m:
        return None
    v = m.group("v")
    if v.startswith("#"):
        h = v[1:]
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if len(h) == 6:
            try:
                return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            except ValueError:
                return None
        return None
    return NAMED_COLORS.get(v.lower())


def _strip_tags(text):
    return TAG_RE.sub("", text)


class Renderer:
    def __init__(self, doc, *, embed_images=False, image_fetch=None):
        self.doc = doc
        self.fnotes = fn.FootnoteManager(doc)
        fn.ensure_styles(doc, FONT_MAIN, FONT_MAIN, SIZE_FOOTNOTE)
        self.embed_images = embed_images
        self.image_fetch = image_fetch          # (url) -> 本地路径 或 None
        self.n_images = 0
        self.n_images_failed = 0

    # ---- 脚注

    def add_footnote(self, paragraph, text):
        text = _strip_tags(text).strip()
        if not text:
            return
        nid = self.fnotes.new_note()
        para = self.doc.add_paragraph(style=fn.STYLE_TEXT)
        run = para.add_run(text)
        set_run_font(run, font_name=FONT_MAIN, size_pt=SIZE_FOOTNOTE)
        self.fnotes.attach(nid, para)
        self.fnotes.add_reference(paragraph, nid)

    # ---- 行内

    def add_inline(self, paragraph, text, *, size_pt=SIZE_BODY,
                   font_name=FONT_MAIN, color=None, bold=False, highlight=False):
        """把一段带行内标记的文本写进段落。

        各分支都把 highlight/color/bold 透传下去递归解析，别在这里把内层标记
        当纯文本拍平——高亮里嵌脚注、嵌粗体是常见写法，拍平会把 ^[...] 原样漏进正文。
        """

        def plain(s, **kw):
            if not s:
                return
            opts = dict(font_name=font_name, size_pt=size_pt, color=color, bold=bold)
            opts.update(kw)
            run = paragraph.add_run(s)
            set_run_font(run, **opts)
            if highlight:
                set_highlight(run)
            return run

        pos = 0
        for m in INLINE_RE.finditer(text):
            plain(text[pos:m.start()])
            pos = m.end()

            if m.group("memo"):
                raw = m.group("memo")
                note = MEMO_CONTENT_RE.search(raw)
                inner = SPAN_INNER_RE.match(raw)
                plain(_strip_tags(inner.group("inner")) if inner else "")
                if note:
                    self.add_footnote(paragraph, note.group("c"))

            elif m.group("color"):
                raw = m.group("color")
                c = _parse_color(raw)
                inner = SPAN_INNER_RE.match(raw)
                self.add_inline(paragraph,
                                inner.group("inner") if inner else _strip_tags(raw),
                                size_pt=size_pt, font_name=font_name,
                                color=c or color, bold=bold, highlight=highlight)

            elif m.group("image"):
                self.add_image(m.group("isrc"), m.group("alt"))

            elif m.group("link"):
                add_hyperlink(paragraph, m.group("lurl"), m.group("ltext"), size_pt)

            elif m.group("fnote"):
                self.add_footnote(paragraph, m.group("ftext"))

            elif m.group("sup"):
                self.add_footnote(paragraph, m.group("suptext"))

            elif m.group("mark"):
                self.add_inline(paragraph, m.group("marktext"), size_pt=size_pt,
                                font_name=font_name, color=color, bold=bold,
                                highlight=True)

            elif m.group("bold"):
                inner = m.group("boldtext")
                ts = re.fullmatch(r"\[?(\d{1,2}:\d{2}(?::\d{2})?)\]?", inner.strip())
                if ts:
                    run = paragraph.add_run(f"〔{ts.group(1)}〕")
                    set_run_font(run, font_name=FONT_SUB, size_pt=SIZE_MARK,
                                 bold=True, color=GRAY_MARK)
                    if highlight:
                        set_highlight(run)
                else:
                    self.add_inline(paragraph, inner, size_pt=size_pt,
                                    font_name=font_name, color=color, bold=True,
                                    highlight=highlight)

            elif m.group("code"):
                plain(m.group("codetext"), font_name=FONT_MONO)

        plain(text[pos:])

    # ---- 图片

    def add_image(self, url, alt=""):
        if not self.embed_images or self.image_fetch is None:
            if alt:
                p = self.doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_spacing(p, before_pt=4, after_pt=4, line_spacing=1.2)
                run = p.add_run(f"［图：{alt}］")
                set_run_font(run, font_name=FONT_SUB, size_pt=SIZE_MARK, color=GRAY_MARK)
            return
        path = self.image_fetch(url)
        if not path:
            self.n_images_failed += 1
            return
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before_pt=6, after_pt=2, line_spacing=1.0)
        try:
            run = p.add_run()
            run.add_picture(str(path), width=content_width(self.doc))
            self.n_images += 1
        except Exception:
            self.n_images_failed += 1
            self.doc.paragraphs[-1]._p.getparent().remove(self.doc.paragraphs[-1]._p)
            return
        if alt and not re.fullmatch(r"(image|图片|img)?\d*", alt.strip(), re.I):
            cap = self.doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_spacing(cap, before_pt=0, after_pt=6, line_spacing=1.2)
            run = cap.add_run(alt.strip())
            set_run_font(run, font_name=FONT_SUB, size_pt=SIZE_MARK, color=GRAY_MARK)


# ---------------------------------------------------------------- 块级解析

RE_TS_LINE = re.compile(r"^(?:\*\*)?[\[【]?(\d{1,2}:\d{2}(?::\d{2})?)[\]】]?(?:\*\*)?$")
RE_LIST = re.compile(r"^\s*(?P<m>[-*+]|\d+[.)])\s+(?P<t>.*)$")
RE_TABLE_SEP = re.compile(r"^\|?[\s:|-]+\|[\s:|-]*$")

HEAD_STYLE = {
    1: dict(size=SIZE_TITLE, before=18, after=12, align=WD_ALIGN_PARAGRAPH.CENTER),
    2: dict(size=SIZE_H1, before=18, after=8, align=WD_ALIGN_PARAGRAPH.LEFT),
    3: dict(size=SIZE_H2, before=12, after=6, align=WD_ALIGN_PARAGRAPH.LEFT),
    4: dict(size=SIZE_H3, before=10, after=4, align=WD_ALIGN_PARAGRAPH.LEFT),
    5: dict(size=SIZE_H3, before=8, after=4, align=WD_ALIGN_PARAGRAPH.LEFT),
    6: dict(size=SIZE_H3, before=6, after=4, align=WD_ALIGN_PARAGRAPH.LEFT),
}


def _split_row(line):
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def render_markdown(md_text, out_path: Path, *, embed_images=False,
                    image_fetch=None, meta_head_lines=20):
    doc = create_document()
    r = Renderer(doc, embed_images=embed_images, image_fetch=image_fetch)

    lines = md_text.splitlines()
    i, n = 0, len(lines)
    body_started = False
    in_code = False
    expected_headings = []

    while i < n:
        raw = lines[i]
        line = raw.strip()

        if line.startswith("```"):
            in_code = not in_code
            i += 1
            continue

        if in_code:
            p = doc.add_paragraph()
            set_spacing(p, before_pt=0, after_pt=0, line_spacing=1.15)
            p.paragraph_format.left_indent = Pt(18)
            run = p.add_run(raw)
            set_run_font(run, font_name=FONT_MONO, size_pt=SIZE_MARK)
            i += 1
            continue

        if not line:
            i += 1
            continue

        if line in ("---", "***", "___"):
            i += 1
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            st = HEAD_STYLE[level]
            p = doc.add_paragraph()
            p.alignment = st["align"]
            set_spacing(p, before_pt=st["before"], after_pt=st["after"], line_spacing=1.25)
            run = p.add_run(_strip_tags(m.group(2)).strip())
            set_run_font(run, font_name=FONT_MAIN, size_pt=st["size"], bold=True)
            set_outline_level(p, level)
            expected_headings.append((p.text, level))
            if level >= 2:
                body_started = True
            i += 1
            continue

        # 引用 / 元信息块
        if line.startswith(">"):
            text = line.lstrip(">").strip()
            is_meta = (not body_started) and i < meta_head_lines
            p = doc.add_paragraph()
            if is_meta:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_spacing(p, before_pt=2, after_pt=2, line_spacing=1.25)
                r.add_inline(p, text, font_name=FONT_SUB, size_pt=SIZE_AUX, color=GRAY_META)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                set_first_line_indent(p)
                set_spacing(p, before_pt=3, after_pt=3, line_spacing=1.3)
                p.paragraph_format.left_indent = Pt(16)
                p.paragraph_format.right_indent = Pt(16)
                r.add_inline(p, text, font_name=FONT_SUB, size_pt=SIZE_AUX, color=GRAY_QUOTE)
            i += 1
            continue

        # 表格
        if line.startswith("|") and i + 1 < n and RE_TABLE_SEP.match(lines[i + 1].strip()):
            header = _split_row(line)
            rows = []
            j = i + 2
            while j < n and lines[j].strip().startswith("|"):
                rows.append(_split_row(lines[j]))
                j += 1
            t = doc.add_table(rows=1 + len(rows), cols=len(header))
            t.style = "Table Grid"
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for c, txt in enumerate(header):
                cell = t.cell(0, c)
                cell.text = ""
                r.add_inline(cell.paragraphs[0], txt, size_pt=SIZE_AUX, bold=True)
            for ri, row in enumerate(rows, 1):
                for c in range(len(header)):
                    cell = t.cell(ri, c)
                    cell.text = ""
                    r.add_inline(cell.paragraphs[0],
                                 row[c] if c < len(row) else "", size_pt=SIZE_AUX)
            doc.add_paragraph()
            i = j
            continue

        # 独占一行的时间戳
        m = RE_TS_LINE.match(line)
        if m:
            p = doc.add_paragraph()
            set_spacing(p, before_pt=10, after_pt=3, line_spacing=1.2)
            run = p.add_run(f"【{m.group(1)}】")
            set_run_font(run, font_name=FONT_SUB, size_pt=SIZE_MARK, bold=True, color=GRAY_MARK)
            i += 1
            continue

        # 整行就是一张图
        m = re.fullmatch(r"!\[(?P<alt>[^\]]*)\]\((?P<src>[^)\s]+)[^)]*\)", line)
        if m:
            r.add_image(m.group("src"), m.group("alt"))
            body_started = True
            i += 1
            continue

        # 列表
        m = RE_LIST.match(line)
        if m:
            p = doc.add_paragraph()
            set_spacing(p, before_pt=2, after_pt=2, line_spacing=1.3)
            p.paragraph_format.left_indent = Pt(24)
            mk = m.group("m")
            bullet = p.add_run("· " if mk in "-*+" else f"{mk} ")
            set_run_font(bullet, font_name=FONT_MAIN, size_pt=SIZE_BODY, bold=True)
            r.add_inline(p, m.group("t"))
            body_started = True
            i += 1
            continue

        # 正文段落
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_first_line_indent(p)
        set_spacing(p, before_pt=0, after_pt=0, line_spacing=1.35)

        text = line
        m = re.match(r"^[\[【](\d{1,2}:\d{2}(?::\d{2})?)[\]】]\s*", text)
        if m:
            run = p.add_run(f"〔{m.group(1)}〕")
            set_run_font(run, font_name=FONT_SUB, size_pt=SIZE_MARK, bold=True, color=GRAY_MARK)
            sp = p.add_run(" ")
            set_run_font(sp, size_pt=SIZE_BODY)
            text = text[m.end():]

        r.add_inline(p, text)
        body_started = True
        i += 1

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    n_notes = r.fnotes.save(doc, out_path)
    n_headings = verify_outline(out_path, expected_headings)
    if len(md_text) >= 2000 and not any(level >= 2 for _, level in expected_headings):
        import warnings
        warnings.warn("长文未识别到正文小节大纲；请核对源文章节。加粗和独占时间戳不等于标题。",
                      UserWarning, stacklevel=2)
    return {"footnotes": n_notes, "images": r.n_images, "images_failed": r.n_images_failed,
            "headings": n_headings}


def render_file(md_path, out_path, **kw):
    return render_markdown(Path(md_path).read_text(encoding="utf-8"), Path(out_path), **kw)


def main():
    import argparse
    ap = argparse.ArgumentParser(description="Markdown → 学术版式 DOCX")
    ap.add_argument("input", help="输入 Markdown 文件")
    ap.add_argument("-o", "--output", required=True, help="输出 DOCX 路径")
    args = ap.parse_args()
    st = render_file(args.input, args.output)
    print(f"✓ {args.output}  脚注 {st['footnotes']}  大纲标题 {st['headings']}（已回读核验）")


if __name__ == "__main__":
    main()
