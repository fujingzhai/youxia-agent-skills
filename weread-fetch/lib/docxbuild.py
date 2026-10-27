#!/usr/bin/env python3
"""把抓回来的 EPUB XHTML 拼成一份可用于研究的 Word 文档。

处理的复杂结构：
  图片   —— 块图（按 CSS 里的 alone50/80/100 定宽）、随文小图、生僻字图，
            图说单独成段用楷体；下载失败会在原位留记号而不是静默丢掉。
  表格   —— 还原 colspan/rowspan 合并，加细边框，表头加粗并跨页重复。
  注释   —— 章末 p.fnote + 正文锚点，以及微信读书自己的 img.qqreader-footnote
            （注文藏在 alt 里），统一转成 Word 真脚注。
  样式   —— 字号/字体/对齐/缩进从本章 CSS 求值，不认死出版社的 class 名。
"""

from __future__ import annotations

import hashlib
import io
from itertools import groupby
from collections import Counter
import json
import os
import re
import sys
import tempfile
import urllib.request
from dataclasses import dataclass, replace as dc_replace
from pathlib import Path

from bs4 import BeautifulSoup, Comment, NavigableString, Tag
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lib.cssmap import FONT_TABLE, Computed, Stylesheet, compute  # noqa: E402
from lib import footnotes as fn  # noqa: E402
from lib import pua  # noqa: E402
from lib import fontcover  # noqa: E402
from lib.titles import clean_title  # noqa: E402

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

BODY_SIZE_PT = 11.0
FOOTNOTE_SIZE_PT = 9.0
MAX_IMAGE_HEIGHT_CM = 17.5

ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}

INLINE_TAGS = {"span", "a", "b", "strong", "i", "em", "u", "sup", "sub", "small",
               "code", "font", "s", "strike", "cite", "abbr", "ruby", "rt", "big", "tt", "var"}
SKIP_TAGS = {"style", "script", "head", "title", "link", "meta", "svg"}

# 边界要认空白：调用方把多个类名用空格拼起来，「footnote moofs14」原先就认不出（读墨制的书）
RE_NOTE_CLASS = re.compile(r"(^|[\s_-])(f?notes?|footnotes?|zhu(shi)?|annot(ation)?|endnote)([\s_-]|$)", re.I)
# EPUB3 的语义标注：注文块 epub:type="footnote"、注号 epub:type="noteref"（role 是 ARIA 的同义写法）
RE_NOTE_TYPE = re.compile(r"\b(footnote|endnote|rearnote|note)\b", re.I)
RE_NOTEREF_TYPE = re.compile(r"\bnoteref\b|doc-noteref", re.I)
RE_BACKLINK_TEXT = re.compile(r"^[↺↩↵⏎⤴↑^返回]+$")
RE_EPUB_NOTE_NUMBER = re.compile(r"^\s*\d+\s*[．.、]\s*")
RE_CAPTION_CLASS = re.compile(
    r"(picture[-_]?txt|pic[-_]?txt|img[-_]?txt|img[-_]?title|pic[-_]?title"
    r"|caption|图说|图题|tushuo|imgnote)", re.I)
RE_POEM_CLASS = re.compile(r"(poem|shi[-_]?ci|verse|quotation|quote|yinwen)", re.I)
RE_SIGNATURE_CLASS = re.compile(r"(signature|right[-_]?info|author[-_]?sign|luokuan)", re.I)
# 各家标注图片宽度的类名不同：kindle-cn 用 alone50/80/100，别家用 width70/width90，
# 数字都是相对版心的百分比
RE_ALONE = re.compile(r"(?:alone|width|pic|img)[-_]?(\d{2,3})(?![\d])")
RE_UI_IMAGE = re.compile(r"(^|/)(note|footnote|annot)\w*\.(png|gif|svg)$", re.I)
RE_CJK = re.compile(r"[\u3400-\u9fff]")
# 繁体底本的书（如上海古籍《金明館叢稿初編》）会把繁简两套正文一起下发，
# 用这两个 class 区分，阅读器靠 CSS 只显示一套。不挑的话整本书会出两遍。
VARIANT_CLASS = {"traditional": "wr-traditional", "simplified": "wr-simplified"}


def to_traditional(text: str) -> str:
    """借本机 opencc 转繁。只用于书名作者这类**我们自己生成**的文字，
    正文一个字都不动。"""
    import shutil
    import subprocess
    exe = shutil.which("opencc")
    if not exe or not text:
        return text
    try:
        out = subprocess.run([exe, "-c", "s2t"], input=text, capture_output=True,
                             text=True, timeout=15)
        return out.stdout.strip() or text
    except Exception:
        return text
RE_SLUG = re.compile(r"^[A-Za-z][\w.-]*$")
# 平台自己塞进版权页的发行声明，不属于书的内容
RE_PLATFORM_NOTICE = re.compile(
    r"(授权微信读书|微信读书[^。]{0,12}(制作|发行)|由.{0,30}授权.{0,10}制作与发行"
    r"|本(电子)?书.{0,20}微信读书.{0,20}(制作|发行|独家))")


def norm(text: str) -> str:
    return re.sub(r"[\s　·．.，,、：:—\-–—_〔〕\[\]()（）]", "", text or "")


# ---------------------------------------------------------------- 图片

class ImageStore:
    def __init__(self, work_dir: Path, cookies: str = ""):
        self.dir = work_dir / "images"
        self.dir.mkdir(parents=True, exist_ok=True)
        self.cookies = cookies
        self.failed: list[str] = []
        self._cache: dict[str, Path | None] = {}

    def get(self, url: str) -> Path | None:
        if url in self._cache:
            return self._cache[url]
        name = hashlib.md5(url.encode()).hexdigest()[:16]
        existing = list(self.dir.glob(name + ".*"))
        if existing:
            self._cache[url] = existing[0]
            return existing[0]

        headers = {"User-Agent": UA, "Referer": "https://weread.qq.com/"}
        if self.cookies:
            headers["Cookie"] = self.cookies
        path = None
        for _ in range(2):
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=40) as resp:
                    data = resp.read()
                    ctype = (resp.headers.get("Content-Type") or "").lower()
                ext = (".png" if "png" in ctype else ".gif" if "gif" in ctype
                       else ".webp" if "webp" in ctype else ".jpg")
                if data[:8] == b"\x89PNG\r\n\x1a\n":
                    ext = ".png"
                elif data[:3] == b"GIF":
                    ext = ".gif"
                elif data[:2] == b"\xff\xd8":
                    ext = ".jpg"
                path = self.dir / (name + ext)
                path.write_bytes(data)
                break
            except Exception:
                path = None
        if path is None:
            self.failed.append(url)
        self._cache[url] = path
        return path

    @staticmethod
    def size(path: Path) -> tuple[int, int] | None:
        try:
            from PIL import Image
            with Image.open(path) as im:
                return im.size
        except Exception:
            return None

    # 标准 JFIF APP0 段：版本 1.1、无单位、像素比 1:1、无缩略图
    JFIF_APP0 = b"\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"

    @classmethod
    def for_docx(cls, path: Path):
        """交给 add_picture 的东西：认得的给原路径，认不得的转成认得的字节流。

        python-docx 只认 PNG、带 JFIF/Exif 头的 JPEG、GIF、TIFF、BMP，认 JPEG 靠的是
        文件头第 6 字节起的 JFIF/Exif 标记。微信读书封面图床给的 JPEG 常常没有这个标记
        （开头直接是量化表 FF D8 FF DB），像素完全正常，它却抛 UnrecognizedImageError，
        以前被 except 吞掉，21 本书的扉页封面就这样丢了、验收照样通过。
        按损失从小到大试：原样 → JPEG 补一个 JFIF 头（一个像素不动）→ Pillow 重编码
        （webp、Exif 段坏掉的 JPEG）。都不行才抛错，由调用方留记号。缓存里的原图不动。
        """
        from docx.image.image import Image as DocxImage
        try:
            DocxImage.from_file(str(path))
            return str(path)
        except Exception:
            pass
        data = path.read_bytes()
        if data[:2] == b"\xff\xd8":
            fixed = data[:2] + cls.JFIF_APP0 + data[2:]
            try:
                DocxImage.from_blob(fixed)
                return io.BytesIO(fixed)
            except Exception:
                pass
        import warnings
        from PIL import Image, ImageOps
        buf = io.BytesIO()
        with warnings.catch_warnings(), Image.open(path) as src:
            warnings.simplefilter("ignore")          # 坏 Exif 会刷 UserWarning
            fmt = (src.format or "").upper()
            try:
                im = ImageOps.exif_transpose(src)    # 丢 Exif 之前先把旋转落到像素上
            except Exception:
                im = src
            if fmt == "JPEG" and im.mode in ("RGB", "L"):
                im.save(buf, "JPEG", quality=95)
            else:
                im.convert("RGBA").save(buf, "PNG")
        buf.seek(0)
        return buf


# ---------------------------------------------------------------- 构建器

@dataclass
class Options:
    title: str = ""
    author: str = ""
    toc: bool = False               # 书前目录默认不排：成品靠侧栏大纲导航，--toc 才排
    page_break_chapters: bool = True
    body_size_pt: float = BODY_SIZE_PT
    keep_cover: bool = True
    body_font: str = "中华书局宋体"   # 正文中文字体；成套字体会自动展开成备用链
    variant: str = "traditional"     # 繁体底本的书会同时下发繁简两套，取哪一套


class BookBuilder:
    def __init__(self, work_dir: Path, meta: dict, opts: Options, cookies: str = ""):
        self.work = work_dir
        self.meta = meta
        self.opts = opts
        self.images = ImageStore(work_dir, cookies)
        self.fonts = dict(FONT_TABLE)
        self.dual_variant = False        # 这本书是否同时下发繁简两套
        self.body_text: list[str] = []   # 供繁体本查简体残留
        self.chain: fontcover.FontChain | None = None
        self.font_note = ""
        if opts.body_font:
            en = self.fonts["song"][1]
            primary, fallbacks = fontcover.resolve_preset(opts.body_font)
            if fontcover.coverage(primary) is None:
                self.font_note = f"本机没装「{primary}」，正文退回宋体"
                primary = FONT_TABLE["song"][0]
            self.fonts["song"] = (primary, en)
            # 正文若换成楷体系（如霞鹜文楷），引文再用楷体就没有对比了，改仿宋
            if re.search(r"楷|kai", primary, re.I):
                self.fonts["kai"] = (FONT_TABLE["fangsong"][0], en)
            self.chain = fontcover.FontChain(fallbacks)
        self.doc = Document()
        self.fnotes = fn.FootnoteManager(self.doc)
        self.family_remap: dict[str, str] = {}
        self.pua_table: dict[int, str] = {}      # 私用区码位 -> 真字
        self.pua_manual_cps: set[int] = set()    # 其中哪些是人工判读的
        self.pua_unresolved: set[int] = set()
        self.cross_notes = {"markers": {}, "blocks": {}}   # 跨文件注的配对，build_book 先扫全书填好
        self.body_hei = False            # 整本正文是否黑体，同上
        self.stats = {"chapters": 0, "paragraphs": 0, "images": 0, "tables": 0,
                      "footnotes": 0, "chars": 0, "short_chapters": [],
                      "pua_fixed": 0, "pua_left": [], "pua_map": {},
                      "notices_dropped": 0, "font_borrowed": {}, "font_missing": [],
                      "font_note": "", "variant_dropped": 0, "simplified_left": [],
                      "pua_manual": {}, "pua_fixed_manual": 0, "continuations": 0,
                      "source_fonts": [], "title_fields_cleaned": 0,
                      "chapter_heads": [], "cross_file_notes": 0, "book_fonts": {},
                      "cover": "无", "toc_entries": 0, "image_embed_failures": 0,
                      "platform_dups": [], "unmatched_noterefs": []}
        self._setup()

    # ---- 版式与样式

    def _setup(self):
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(3.0)
        sec.top_margin = sec.bottom_margin = Cm(2.5)
        self.text_width_cm = 21.0 - 6.0

        base = self.opts.body_size_pt
        normal = self.doc.styles["Normal"]
        self._font(normal, "song", base)
        pf = normal.paragraph_format
        pf.first_line_indent = Pt(base * 2)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5

        self._mkstyle("WR 书名", base * 1.9, "hei", bold=True, align="center",
                      indent=0, before=0, after=base * 1.2)
        self._mkstyle("WR 作者", base * 1.15, "kai", align="center",
                      indent=0, before=0, after=base * 0.6)
        self._mkstyle("WR 说明", base * 0.85, "song", align="center",
                      indent=0, before=base, after=0, color="808080")

        for lvl, (name, mult) in enumerate([
                ("WR 标题一", 1.45), ("WR 标题二", 1.25),
                ("WR 标题三", 1.12), ("WR 标题四", 1.0)]):
            st = self._mkstyle(name, base * mult, "hei", bold=True, align=None,
                               indent=0, before=base * (1.2 - lvl * 0.2), after=base * 0.6)
            st.paragraph_format.keep_with_next = True
            self._outline(st, lvl)

        self._mkstyle("WR 正文无缩进", base, "song", indent=0)
        self._mkstyle("WR 引文", base * 0.95, "kai", indent=0, left=base * 2,
                      before=base * 0.4, after=base * 0.4)
        self._mkstyle("WR 图", base, "song", align="center", indent=0,
                      before=base * 0.5, after=base * 0.2)
        self._mkstyle("WR 图说", base * 0.85, "kai", align="center", indent=0,
                      before=0, after=base * 0.6)
        self._mkstyle("WR 落款", base * 0.95, "kai", align="right", indent=0,
                      before=base * 0.3, after=base * 0.3)
        self._mkstyle("WR 表格文字", base * 0.9, "song", indent=0, before=0, after=0)
        self._mkstyle("WR 尾注", base * 0.85, "song", indent=0, before=0, after=0)
        self._mkstyle("WR 分隔", base, "song", align="center", indent=0,
                      before=base * 0.6, after=base * 0.6)
        for lvl in range(4):
            st = self._mkstyle(f"WR 目录{lvl + 1}", base, "song", indent=0,
                               left=base * 2 * lvl, before=0, after=base * 0.15)
            st.paragraph_format.line_spacing = 1.3

        cn, en = self.fonts["song"]
        fn.ensure_styles(self.doc, cn, en, FOOTNOTE_SIZE_PT)

    def _font(self, style, family: str, size_pt: float, bold=False, italic=False):
        cn, en = self.fonts.get(family, self.fonts["song"])
        f = style.font
        f.name = en
        f.size = Pt(size_pt)
        f.bold = bold
        f.italic = italic
        f.color.rgb = RGBColor(0, 0, 0)
        rpr = style.element.get_or_add_rPr()
        rpr.get_or_add_rFonts().set(qn("w:eastAsia"), cn)
        return style

    def _mkstyle(self, name, size_pt, family, *, bold=False, align=None, indent=None,
                 left=None, before=None, after=None, color=None):
        styles = self.doc.styles
        try:
            st = styles[name]
        except KeyError:
            st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            st.base_style = styles["Normal"]
        self._font(st, family, size_pt, bold=bold)
        if color:
            st.font.color.rgb = RGBColor.from_string(color)
        pf = st.paragraph_format
        if align:
            pf.alignment = ALIGN[align]
        if indent is not None:
            pf.first_line_indent = Pt(indent)
        if left is not None:
            pf.left_indent = Pt(left)
        if before is not None:
            pf.space_before = Pt(before)
        if after is not None:
            pf.space_after = Pt(after)
        pf.line_spacing = 1.5
        st.quick_style = True
        return st

    @staticmethod
    def _outline(style, level: int):
        ppr = style.element.get_or_add_pPr()
        el = ppr.find(qn("w:outlineLvl"))
        if el is None:
            el = OxmlElement("w:outlineLvl")
            ppr.append(el)
        el.set(qn("w:val"), str(level))

    # ---- 扉页与目录

    def title_page(self, cover: Path | None = None):
        d = self.doc
        # 书名作者照搬接口原值，不做繁简转换——那是平台的元数据，转它就是改数据。
        # 繁体本的接口书名确实是简体（《金明館叢稿初編》就是），但保留原样与文件名
        # 一致，也符合「保留而非修改」的原则。
        if self.opts.title:
            d.add_paragraph(self.opts.title, style="WR 书名")
        if self.opts.author:
            d.add_paragraph(self.opts.author, style="WR 作者")
        if cover and cover.exists():
            p = d.add_paragraph(style="WR 图")
            try:
                p.add_run().add_picture(ImageStore.for_docx(cover),
                                        width=Cm(min(9.0, self.text_width_cm)))
                self.stats["cover"] = "嵌入"
            except Exception as exc:
                # 扉页不留空段，缺封面进验收警告。这里原先是 except: pass。
                p._p.getparent().remove(p._p)
                self.stats["cover"] = f"嵌不进 Word（{type(exc).__name__}）"

    def toc_anchor(self):
        """书前目录先占个位，正文排完再按真实标题回填（见 fill_toc）。"""
        return self.doc.add_paragraph(style="WR 正文无缩进")

    def fill_toc(self, anchor, depth: int) -> int:
        """书前目录（默认不排，--toc 才排），返回条目数。

        条目取正文里实际排出来的标题，不取 chapterInfos：接口只列它切开的那几个文件，
        同一文件里的小节不单列（《失控》每章的「x.1」都和章名在一个文件里，照接口排的
        旧目录整整少了 24 条，侧栏大纲里却有）。深度跟章节表走——接口列到几级就排到
        几级，免得节下的小标题把目录撑得比原书还长。做成纯文字（不用 Word 目录域、
        不做跳转），省得每次打开都要「更新域」。
        """
        levels = {self.doc.styles[name].style_id: lvl for name, lvl in
                  (("WR 标题一", 1), ("WR 标题二", 2), ("WR 标题三", 3), ("WR 标题四", 4))}
        entries, after = [], False
        for para in self.doc.paragraphs:
            if para._p is anchor._p:
                after = True
                continue
            lvl = levels.get(para._p.style) if after else None
            text = " ".join(para.text.split()) if lvl and lvl <= depth else ""
            if text:
                entries.append((lvl, text))
        conv = (to_traditional if (self.dual_variant
                                   and self.opts.variant.startswith(("t", "繁")))
                else (lambda x: x))
        anchor.insert_paragraph_before(conv("目　录"), style="WR 标题一")
        for lvl, text in entries:
            self.add_text(anchor.insert_paragraph_before(style=f"WR 目录{lvl}"),
                          text, Computed())
        anchor.add_run().add_break(WD_BREAK.PAGE)
        return len(entries)

    def page_break(self):
        self.doc.add_paragraph(style="WR 正文无缩进").add_run().add_break(WD_BREAK.PAGE)

    # ---- 段落与文字

    def _apply_para(self, para, st: Computed, role: str):
        base = self.opts.body_size_pt
        pf = para.paragraph_format
        if st.align in ALIGN:
            pf.alignment = ALIGN[st.align]
        if st.indent_em is not None:
            pf.first_line_indent = Pt(base * st.size_em * st.indent_em)
        if st.left_em:
            pf.left_indent = Pt(base * st.size_em * min(st.left_em, 8.0))
            if st.indent_em is None:
                pf.first_line_indent = Pt(0)
        if st.space_before_em:
            pf.space_before = Pt(base * min(st.space_before_em, 2.0))
        if st.space_after_em:
            pf.space_after = Pt(base * min(st.space_after_em, 2.0))
        if role == "body" and st.indent_em is None and st.align == "center":
            pf.first_line_indent = Pt(0)

    def _apply_run(self, run, st: Computed, set_font: bool = True, set_size: bool = True,
                   remap: bool = True, font_override: str | None = None):
        if set_size:
            run.font.size = Pt(round(self.opts.body_size_pt * st.size_em, 1))
        if font_override or set_font:
            family = self.family_remap.get(st.family, st.family) if remap else st.family
            cn, en = self.fonts.get(family, self.fonts["song"])
            run.font.name = en
            rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
            rfonts.set(qn("w:eastAsia"), font_override or cn)
            if font_override:
                # 借来的字段单独指定，三个属性都写上，免得 Word 按脚本挑错字体
                rfonts.set(qn("w:ascii"), font_override)
                rfonts.set(qn("w:hAnsi"), font_override)
        run.font.bold = st.bold or None
        run.font.italic = st.italic or None
        run.font.underline = st.underline or None
        if st.superscript:
            run.font.superscript = True
        if st.subscript:
            run.font.subscript = True
        if st.color:
            try:
                run.font.color.rgb = RGBColor.from_string(st.color.lstrip("#").upper())
            except Exception:
                pass

    # ---- 图片

    def add_text(self, para, text: str, st: Computed, set_font: bool = True,
                 set_size: bool = True, remap: bool = True) -> list:
        """落一段文字。正文字体缺的字会切出来单独指定备用字体。"""
        # 所有文字入口共用：正文、锚点脚注、alt 脚注、目录等不能各自处理缺字。
        for ch in text:
            cp = ord(ch)
            if cp in self.pua_manual_cps:
                self.stats["pua_fixed_manual"] += 1
            elif cp in self.pua_table:
                self.stats["pua_fixed"] += 1
        text = text.translate(self.pua_table)
        family = self.family_remap.get(st.family, st.family) if remap else st.family
        cn = self.fonts.get(family, self.fonts["song"])[0]
        runs = []
        # 先按是否待核分段，再按字体覆盖切分。高亮只落在未知码位本身。
        for unknown, chars in groupby(text, key=lambda c: pua.is_pua(ord(c))):
            chunk = "".join(chars)
            parts = self.chain.split(chunk, cn) if self.chain else [(chunk, None)]
            for seg, override in parts:
                if not seg:
                    continue
                run = para.add_run(seg)
                self._apply_run(run, st, set_font=set_font, set_size=set_size,
                                remap=remap, font_override=override)
                if unknown:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                runs.append(run)
        return runs

    def _image_width(self, el: Tag, path: Path) -> tuple[float, float]:
        classes = " ".join(el.get("class") or [])
        pct = None
        m = RE_ALONE.search(classes)
        if m:
            pct = int(m.group(1)) / 100.0
        elif re.search(r"bleed|fullpage|full[-_]?width", classes, re.I):
            pct = 1.0
        if pct is None:
            raw = el.get("data-w-new") or el.get("data-w") or ""
            num = re.sub(r"[^\d.]", "", str(raw))
            if num:
                try:
                    pct = min(1.0, float(num) / 750.0)
                except ValueError:
                    pct = None
        size = ImageStore.size(path)
        if pct is None:
            pct = min(1.0, (size[0] / 900.0)) if size else 0.7
        pct = max(0.15, min(1.0, pct))

        width = self.text_width_cm * pct
        ratio = None
        if size and size[0]:
            ratio = size[1] / size[0]
        else:
            try:
                ratio = float(el.get("data-ratio"))
            except (TypeError, ValueError):
                ratio = None
        if ratio:
            height = width * ratio
            if height > MAX_IMAGE_HEIGHT_CM:
                width *= MAX_IMAGE_HEIGHT_CM / height
                height = MAX_IMAGE_HEIGHT_CM
        else:
            height = 0.0
        return width, height

    def _src(self, el: Tag) -> str | None:
        """取图片地址。

        实测六本书里真内容图一律是 res.weread.qq.com 的绝对地址，相对路径
        （../Images/note.png 之类）都是阅读器自带素材、且多半取不到。但阅读器 JS 里
        确有把 `../images/x.png` 改写成 `wrepub/web/<bookId>/x.png` 的规则（抓它的
        实际请求实测过，路径是拍平的、只留文件名），所以相对路径也按这条解析，
        取不到就记进 image_failures，不静默丢。
        """
        src = (el.get("src") or el.get("data-src") or el.get("xlink:href") or "").strip()
        if not src:
            return None
        if src.startswith("//"):
            return "https:" + src
        if src.startswith("http"):
            return src
        book_id = str(self.meta.get("bookId") or "")
        if not book_id:
            self.images.failed.append(src)      # book.json 坏了才会走到这
            return None
        import urllib.parse
        return (f"https://res.weread.qq.com/wrepub/web/{book_id}/"
                + urllib.parse.quote(src.split("/")[-1]))

    def add_block_image(self, container, el: Tag):
        src = self._src(el)
        if not src:
            return False
        path = self.images.get(src)
        if not path:
            p = container.add_paragraph("〔图片下载失败〕", style="WR 图说")
            return True
        width, _ = self._image_width(el, path)
        p = container.add_paragraph(style="WR 图")
        try:
            p.add_run().add_picture(ImageStore.for_docx(path), width=Cm(round(width, 2)))
            self.stats["images"] += 1
        except Exception:
            p.add_run("〔图片无法嵌入〕")
            self.stats["image_embed_failures"] += 1
        # alt 常常只是制作时的文件标识（copyright-1 / q1 / c-1），不是图说；
        # 真图说一般另有 p.picture-txt。所以只认像人话的 alt。
        alt = (el.get("alt") or "").strip()
        if len(alt) > 6 and RE_CJK.search(alt) and not RE_SLUG.match(alt):
            container.add_paragraph(alt, style="WR 图说")
        return True

    def add_inline_image(self, para, el: Tag, st: Computed,
                         img_style: Computed | None = None, in_text: bool = False):
        """随文图。

        生僻字在很多书里是拍成图片嵌在行内的（陈寅恪《金明館叢稿初編》里的「黄法𣰋」
        就是这么处理的）。这类图必须**按高度**定尺寸——按 data-w 的宽度算的话，
        一个字会变成五厘米宽的大图横在句子中间。判据有三层：CSS 给了 height、
        类名像生僻字图、或者它就嵌在有文字的段落里。
        """
        src = self._src(el)
        if not src:
            return
        path = self.images.get(src)
        if not path:
            # 随文图多半是生僻字（s-pic 一类），静默丢等于丢字，留个高亮记号
            run = para.add_run("〔?〕")
            self._apply_run(run, st)
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            return
        classes = " ".join(el.get("class") or [])
        height_em = img_style.height_em if img_style else None
        looks_glyph = bool(re.search(r"inline[-_]?char|img[-_]?h\d|inline[-_]?image|zi\b",
                                     classes, re.I))
        run = para.add_run()
        try:
            size = ImageStore.size(path)
            ratio = (size[1] / size[0]) if size and size[0] else 1.0
            square = 0.75 <= ratio <= 1.35
            if height_em or looks_glyph or in_text:
                # 方块字图按一个字高排；CSS 里的 height（常写 1.8em）是给阅读器
                # 留白用的，照搬到 Word 会比周围的字明显大一圈。
                em = 1.05 if (square and in_text) else min(height_em or 1.2, 2.0)
                run.add_picture(ImageStore.for_docx(path),
                                height=Pt(round(self.opts.body_size_pt * st.size_em * em, 1)))
            else:
                width, _ = self._image_width(el, path)
                run.add_picture(ImageStore.for_docx(path),
                                width=Cm(round(min(width, self.text_width_cm), 2)))
            self.stats["images"] += 1
        except Exception:
            # 嵌不进去跟下载失败一样处理：随文图多半是生僻字，静默丢掉就是丢字
            run.text = "〔?〕"
            self._apply_run(run, st)
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            self.stats["image_embed_failures"] += 1


# ---------------------------------------------------------------- 章节渲染

BLOCK_TAGS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "table", "ul", "ol",
              "li", "blockquote", "hr", "pre", "section", "article", "figure",
              "figcaption", "center", "dl", "dt", "dd", "tr", "td", "th", "tbody",
              "thead", "caption"}


def build_note_map(soup) -> tuple[dict, list]:
    """把「正文里的注号」对上「章末的注文块」。

    各家 EPUB 的写法不统一，见过这些：
      1. 注号 <a href="#ch1" id="ch1-back">，注文 <p class="fnote"><a id="ch1">
      2. 只有 href 没有 id：注号 <a href="#ch1">，注文里回跳 <a href="#ch1-back">
      3. 注号与注文两个 <a> 互指对方 id，class 毫无线索
      4. 注号的 id 挂在另一个空 <a> 上：<a id="w48"></a><a href="#m48">[48]</a>
      5. class 名各异（fnote / note / content2a / content3 …）
    先按 class、回跳后缀、段首锚点、互指四条路收集注文块，再按 id 与
    「回跳 href 去掉 -back」配对，配不上的若数量正好相等就按出现顺序配。
    """
    cands: list[Tag] = []
    for el in soup.find_all(["p", "div", "li", "aside"]):
        first_a = el.find("a")
        href = (first_a.get("href") or "") if first_a is not None else ""
        backish = href.startswith("#") and re.search(r"[-_]?(back|ref)$", href[1:] or "")
        if _is_note_block(el) or backish:
            cands.append(el)

    # 注文侧的锚点：一个 <a> 同时带 id（供注号指过来）与 href（回跳），且在段首。
    # 中华书局那家就是这么做的——注号那侧的 id 挂在**另一个空 <a> 上**
    # （`<a id="w48"></a><a href="#m48">[48]</a>`），所以严格的「两个 a 互指」
    # 判据认不出，得放宽成「有 id、有回跳、在段首、且确有注号指向它」。
    for a in soup.find_all("a"):
        own, href = a.get("id"), (a.get("href") or "")
        if not own or not href.startswith("#"):
            continue
        block = a.find_parent(["p", "div", "li", "td"])
        if block is None or not _starts_block(a, block):
            continue
        if soup.find("a", href="#" + own) is None:
            continue        # 没有注号指向它，那就不是注文
        if not any(x is block for x in cands):
            cands.append(block)

    # 还有一种写法：注号与注文各是一个 <a>，互相指向对方的 id，class 却毫无线索
    # （孙英刚《神文时代》：注号 <a href="#b1" id="bt1">注1</a>，
    #   注文 <p class="content2a"><a href="#bt1" id="b1">注1</a>正文…）。
    # 互指本身就是确证，段首的那个是注文。
    by_id = {el["id"]: el for el in soup.find_all(id=True)}
    for a in soup.find_all("a"):
        href, own = (a.get("href") or ""), a.get("id")
        if not href.startswith("#") or not own:
            continue
        other = by_id.get(href[1:])
        if other is None or other.name != "a" or (other.get("href") or "") != "#" + own:
            continue
        for side in (a, other):
            block = side.find_parent(["p", "div", "li", "td"])
            if block is not None and _starts_block(side, block) and block not in cands:
                cands.append(block)
    # 只留最内层，别把包着一堆注的外层 div 当成一条注
    notes = [el for el in cands
             if not any(other is not el and other in el.descendants for other in cands)]

    by_key: dict[str, Tag] = {}
    for el in notes:
        keys = []
        if el.get("id"):
            keys.append(el["id"])
        for a in el.find_all("a"):
            if a.get("id"):
                keys.append(a["id"])
            href = (a.get("href") or "")
            if href.startswith("#"):
                keys.append(re.sub(r"[-_]?(back|ref)$", "", href[1:]))
        for k in keys:
            if k:
                by_key.setdefault(k, el)

    note_ids = {id(n) for n in notes}
    markers: list[Tag] = []
    for a in soup.find_all("a"):
        href = (a.get("href") or "")
        if not href.startswith("#"):
            continue
        if any(id(anc) in note_ids for anc in a.parents):
            continue
        if a.find("sup") or len(a.get_text(strip=True)) <= 6:
            markers.append(a)

    mapping: dict[int, Tag] = {}
    used: set[int] = set()
    pending: list[Tag] = []
    for a in markers:
        key = (a.get("href") or "")[1:]
        target = by_key.get(key)
        if target is not None and id(target) not in used:
            mapping[id(a)] = target
            used.add(id(target))
        else:
            pending.append(a)
    leftover = [n for n in notes if id(n) not in used]
    if pending and len(pending) == len(leftover):
        for a, n in zip(pending, leftover):
            mapping[id(a)] = n
    # 再兑一遍漏锚点的注文：注号的 href 指不到任何 id，而文末有同编号的块。
    still = [a for a in markers if id(a) not in mapping]
    if still:
        orphans = _orphan_note_blocks(soup)
        mine = {id(n) for n in notes}
        for a in still:
            m = RE_NOTE_NUMBER.match(a.get_text(strip=True))
            if not m:
                continue
            block = orphans.get(m.group(1))
            if block is None or id(block) in mine:
                continue
            mapping[id(a)] = block
            mine.add(id(block))
            used.add(id(block))

    # 同一条注被多次引用（本书《对《毛主席诗词》若干词句的解释》里 [14]/[16]/[17]/[20]
    # 各被引两到五次，《在西北野战军前委扩大会议上的讲话》的 [6] 也是）。
    # 多个注号不能指向同一个脚注：LibreOffice 实测会把两个引用号都编上、注文只印一份，
    # 后面的编号就漂了（也过不了「引用与注文一一对应」这道验收）。一个引用一条脚注、
    # 注文各印一份，是通用阅读器都能正确渲染的做法，宁可注文重复也不让编号错位。
    for a in markers:
        if id(a) in mapping:
            continue
        m = RE_NOTE_NUMBER.match(a.get_text(strip=True))
        if not m:
            continue
        for b in markers:
            if id(b) not in mapping:
                continue
            mb = RE_NOTE_NUMBER.match(b.get_text(strip=True))
            if mb and mb.group(1) == m.group(1):
                mapping[id(a)] = BeautifulSoup(str(mapping[id(b)]), "html.parser").find(True)
                break
    return mapping, notes


def _is_note_label(el: Tag) -> bool:
    """「注释」/「注」这类统辖注文的小标题，不是书的内容。

    注文都挪进 Word 脚注了，这行标签下面空空如也，留着只会在正文里印出一串
    无所指的「注释」（本书 766 处）。排版常把字距拉成全角空格（「注　　释」），
    所以比对时先把空白一律抹掉。只认短到只有一个词的（「注释」「注」、
    「註釋」这类繁体写法），有正文的长段落一概不动。
    """
    if el.name not in ("p", "div", "h1", "h2", "h3", "h4", "h5", "h6"):
        return False
    text = re.sub(r"\s+", "", el.get_text(" ", strip=True))
    return text in ("注释", "注", "註釋", "注釋") and not el.find("img")


RE_NOTE_NUMBER = re.compile(r"^\s*[\[〔(（]\s*(\d+)\s*[\]〕)）]")


def _orphan_note_blocks(soup) -> dict[str, Tag]:
    """按开头编号认出的注文块（§数字），给「注号锚点指向不存在的 id」的章兑底。

    中华书局这套书里偶有几条注文漏了回跳锚点：注号是 <a href="#wz_3_301">[3]</a>，
    可全篇没有 id="wz_3_301"，注文却实实在在地以「〔3〕雇农支部…」开头印在文末。
    这种注号在锚点那条路上永远配不上，只能按编号认领；不认的话，注文会当作
    正文印在章末（本书 2 处，如《长冈乡调查》，另有一处在章标题里）。
    只收「没有可跳转锚点、且以编号开头」的块：有锚点的走正常配对，不在这里抢。
    """
    found: dict[str, Tag] = {}
    for el in soup.find_all(["p", "div", "li"]):
        if el.find(["p", "div"]):      # 只认最内层
            continue
        m = RE_NOTE_NUMBER.match(el.get_text(" ", strip=True))
        if not m:
            continue
        has_anchor = any((a.get("href") or "").startswith("#") for a in el.find_all("a"))
        if has_anchor:
            continue
        found.setdefault(m.group(1), el)
    return found


class Ctx:
    """渲染一章时的临时上下文。"""

    def __init__(self, sheet: Stylesheet, notes: dict, heading_rank: dict,
                 chapter_level: int, chapter_title: str = ""):
        self.sheet = sheet
        self.notes = notes            # id(注号 <a>) -> 注文块
        self.consumed: set[int] = set()
        self.heading_rank = heading_rank
        self.chapter_level = chapter_level
        self.chapter_title = chapter_title
        self.in_footnote = False
        self.after_break = False   # 刚打过换行，下一段文字要去掉行首空白


def _has_block_child(el: Tag) -> bool:
    return any(isinstance(c, Tag) and c.name.lower() in BLOCK_TAGS for c in el.children)


def is_note_image(el: Tag) -> bool:
    """注释图：注文藏在 alt 里，必须留在行内交给脚注逻辑，不能当装饰图挪走。"""
    classes = " ".join(el.get("class") or []).lower()
    return bool("footnote" in classes or "comment" in classes
                or RE_UI_IMAGE.search(el.get("src") or ""))


def _image_only(el: Tag) -> list:
    """整块只有图没有正文时，返回其中的图；否则空表。"""
    imgs = el.find_all("img")
    if not imgs:
        return []
    if any(is_note_image(im) for im in imgs):
        return []  # 短段落里的注释图仍是脚注入口，不能当封面/装饰图。
    text = el.get_text(" ", strip=True)
    return imgs if len(text) <= 2 else []


def _starts_block(node: Tag, block: Tag) -> bool:
    """node 是不是 block 里第一个有内容的东西（注文的回跳链接就在段首）。"""
    for item in block.descendants:
        if item is node:
            return True
        if isinstance(item, NavigableString) and item.strip():
            return False
        if isinstance(item, Tag) and item.name == "img":
            return False
    return False


def _only_consumed_after(node: Tag, ctx: "Ctx") -> bool:
    """分隔线后面只剩已挪进脚注的注文块时，这条线隔开的东西已经不在了。"""
    seen = False
    for sib in node.next_siblings:
        if isinstance(sib, NavigableString):
            if sib.strip():
                return False
        elif isinstance(sib, Tag):
            if _is_note_label(sib):
                continue        # 「注释」这个小标题也跟着注文一起走
            if id(sib) not in ctx.consumed:
                return False
            seen = True
    return seen


def _is_epub_note(el: Tag) -> bool:
    """EPUB3 标准注文块。注号那侧也可能误标成 doc-noteref（读墨就把注文 div 标成了它），
    所以只认 epub:type，且 <a> 不算——<a> 只会是注号。"""
    return el.name != "a" and bool(RE_NOTE_TYPE.search(el.get("epub:type") or ""))


def _is_noteref(a: Tag) -> bool:
    return bool(RE_NOTEREF_TYPE.search(" ".join([a.get("epub:type") or "", a.get("role") or ""])))


def _is_note_block(el: Tag) -> bool:
    return (bool(RE_NOTE_CLASS.search(" ".join(el.get("class") or [])))
            or _is_epub_note(el)
            or (el.get("role") or "") in ("doc-footnote", "doc-endnote"))


class ChapterRenderer:
    """一章 XHTML -> 文档段落。逻辑挂在 BookBuilder 上使用。"""

    def __init__(self, book: BookBuilder):
        self.b = book

    # ---------------------------------------------------- 注释登记

    def build_note_map(self, soup: BeautifulSoup) -> tuple[dict, list]:
        return build_note_map(soup)

    def is_footnote_marker(self, a: Tag, ctx: Ctx) -> Tag | None:
        return ctx.notes.get(id(a))

    # ---------------------------------------------------- 脚注生成

    def make_footnote_from_block(self, block: Tag, ctx: Ctx, style: Computed) -> int:
        note_id = self.b.fnotes.new_note()
        ctx.consumed.add(id(block))
        prev = ctx.in_footnote
        ctx.in_footnote = True

        clone = BeautifulSoup(str(block), "html.parser")
        root = clone.find(True) or clone
        # 去掉行首的编号回跳链接，Word 会自己编号
        first_a = root.find("a")
        href = (first_a.get("href") or "") if first_a is not None else ""
        # 跨文件的注，段首回跳链接带着文件名（../Text/x.xhtml#noteBack_1）
        if href.startswith("#") or ("#" in href and _starts_block(first_a, root)):
            first_a.decompose()
        # EPUB3 标准写法的回跳链接在注文末尾（「…表示。<a role="doc-backlink">↺</a>」），
        # 不去掉的话每条脚注都拖一个箭头
        for a in list(root.find_all("a", href=True)):
            if (a.get("role") == "doc-backlink"
                    or ("#" in a["href"] and RE_BACKLINK_TEXT.match(a.get_text(strip=True) or "x"))):
                a.decompose()
        # 同一写法的注文以纯文字编号开头（「1．編注…」），编号交给 Word 脚注，原样留着就成了「⁸ 1．」
        if _is_epub_note(block) or _is_epub_note(root):
            lead = root.find(string=re.compile(r"\S"))
            if lead is not None and RE_EPUB_NOTE_NUMBER.match(str(lead)):
                lead.replace_with(RE_EPUB_NOTE_NUMBER.sub("", str(lead), count=1))

        # 一些书将整段 XHTML 转义后塞进注文本身；仅解包完整段落标记，
        # 不把普通讨论中的 <a> 等字样当标签删掉。
        for node in list(root.find_all(string=True)):
            value = str(node).strip()
            if re.fullmatch(r"<(p|div)\b[^>]*>.*</\1>", value, re.I | re.S):
                fragment = BeautifulSoup(value, "html.parser")
                for element in list(fragment.contents):
                    node.insert_before(element)
                node.extract()

        nested = root.find_all(["p", "div"], recursive=False)
        blocks = nested if nested else [root]
        for el in blocks:
            para = self.b.doc.add_paragraph(style=fn.STYLE_TEXT)
            st = Computed(size_em=FOOTNOTE_SIZE_PT / self.b.opts.body_size_pt)
            self.emit_inline(para, el, ctx, st)
            if not para.text.strip() and not para.runs:
                continue
            self.b.fnotes.attach(note_id, para)
        ctx.in_footnote = prev
        self.b.stats["footnotes"] += 1
        return note_id

    def make_footnote_from_text(self, text: str, ctx: Ctx) -> int:
        if re.search(r"<(?:p|div|span|br|img|b|i|em|strong)\b", text, re.I):
            fragment = BeautifulSoup("<div>" + text + "</div>", "html.parser")
            return self.make_footnote_from_block(fragment.div, ctx, Computed())
        note_id = self.b.fnotes.new_note()
        para = self.b.doc.add_paragraph(style=fn.STYLE_TEXT)
        st = Computed(size_em=FOOTNOTE_SIZE_PT / self.b.opts.body_size_pt)
        self.b.add_text(para, text.strip(), st)
        self.b.fnotes.attach(note_id, para)
        self.b.stats["footnotes"] += 1
        return note_id

    def emit_parenthetical(self, para, text: str, ctx: Ctx, style: Computed):
        text = text.strip()
        self.b.add_text(para, "（", style)
        if re.search(r"<(?:p|div|span|br|img|b|i|em|strong)\b", text, re.I):
            holder = BeautifulSoup("<span>" + text + "</span>", "html.parser").span
            self.emit_inline(para, holder, ctx, style)
        else:
            self.b.add_text(para, text, style)
        self.b.add_text(para, "）", style)

    # ---------------------------------------------------- 行内

    def emit_inline(self, para, node, ctx: Ctx, style: Computed, set_font: bool = True,
                    set_size: bool = True, remap: bool = True):
        for child in node.children:
            if isinstance(child, NavigableString):
                text = str(child)
                if not text.strip() and "\n" in text:
                    continue
                text = re.sub(r"[ \t\r\n]+", " ", text)
                if ctx.after_break:
                    text = text.lstrip()
                if not text:
                    continue
                ctx.after_break = False
                if self.b.dual_variant:
                    self.b.body_text.append(text)
                self.b.add_text(para, text, style,
                                       set_font=set_font or style.family_explicit,
                                       set_size=set_size or style.size_explicit,
                                       remap=remap)
                self.b.stats["chars"] += len(text.strip())
                continue
            if not isinstance(child, Tag):
                continue

            name = child.name.lower()
            if name in SKIP_TAGS:
                continue
            if name == "br":
                # 换行前后的空白要吃掉：源码里 <br/> 常跟着缩进换行，
                # 归一成空格后每行会平白多一个行首空格——词、诗这种讲对齐的最明显
                if para.runs and para.runs[-1].text.endswith(" "):
                    para.runs[-1].text = para.runs[-1].text.rstrip()
                para.add_run().add_break()
                ctx.after_break = True
                continue

            if name == "img":
                alt = (child.get("alt") or "").strip()
                is_note_img = is_note_image(child)
                if is_note_img and alt and not ctx.in_footnote:
                    nid = self.make_footnote_from_text(alt, ctx)
                    self.b.fnotes.add_reference(para, nid)
                    continue
                if is_note_img and alt:
                    # 脚注里不能再套脚注。注中之注多是英文原词或文献出处，按括注接在原处；
                    # 不接住的话会把 note.png 当随文图去取，取不到就成了〔?〕
                    self.emit_parenthetical(para, alt, ctx, style)
                    continue
                if is_note_img and not alt:
                    continue      # 只是个注释小图标，没内容
                # 注文里也可能夹生僻字图（孙英刚《神文时代》就有），
                # footnotes.xml 现在有自己的关系分部，能正常嵌图
                self.b.add_inline_image(
                    para, child, style,
                    img_style=compute(child, ctx.sheet, style.child()),
                    in_text=len(node.get_text(strip=True)) > 2)
                continue

            child_style = compute(child, ctx.sheet, style.child())

            if name == "a" and not ctx.in_footnote:
                target = self.is_footnote_marker(child, ctx)
                if target is not None:
                    nid = self.make_footnote_from_block(target, ctx, child_style)
                    self.b.fnotes.add_reference(para, nid)
                    continue

            self.emit_inline(para, child, ctx, child_style, set_font, set_size, remap)

    # ---------------------------------------------------- 块

    def _para_style_name(self, el: Tag, st: Computed) -> tuple[str, str]:
        classes = " ".join(el.get("class") or [])
        if RE_NOTE_CLASS.search(classes):
            return "note", "WR 尾注"
        if RE_CAPTION_CLASS.search(classes):
            return "caption", "WR 图说"
        if RE_SIGNATURE_CLASS.search(classes):
            return "signature", "WR 落款"
        if RE_POEM_CLASS.search(classes) or el.name.lower() == "blockquote":
            return "quote", "WR 引文"
        if st.indent_em is not None and abs(st.indent_em) < 0.2:
            return "body", "WR 正文无缩进"
        return "body", "Normal"

    def emit_paragraph(self, container, nodes, ctx: Ctx, style: Computed,
                       style_name: str, role: str, source: Tag | None = None):
        wrapper = source
        para = container.add_paragraph(style=style_name)
        ctx.after_break = False
        self.b._apply_para(para, style, role)
        if wrapper is not None:
            self.emit_inline(para, wrapper, ctx, style)
        else:
            holder = BeautifulSoup("<span></span>", "html.parser").span
            for n in nodes:
                holder.append(n.__copy__() if isinstance(n, Tag) else NavigableString(str(n)))
            self.emit_inline(para, holder, ctx, style)
        if not para.runs or RE_PLATFORM_NOTICE.search(para.text):
            if para.runs and RE_PLATFORM_NOTICE.search(para.text):
                self.b.stats["notices_dropped"] += 1
            p = para._p
            p.getparent().remove(p)
            return None
        self.b.stats["paragraphs"] += 1
        return para

    def emit_heading(self, container, el: Tag, ctx: Ctx, style: Computed):
        depth = ctx.heading_rank.get(el.name.lower(), 1)
        level = max(0, min(3, ctx.chapter_level - 1 + depth))
        name = ["WR 标题一", "WR 标题二", "WR 标题三", "WR 标题四"][level]

        # 有些书的标题是一张美术字图片，真标题在 title 属性里（章节表里也有）。
        # 把图当标题的话，导航窗格会多出一堆没有文字的条目，所以图挪出去单独成段。
        images = [im for im in el.find_all("img") if not is_note_image(im)]
        has_text = bool(el.get_text(strip=True))
        para = container.add_paragraph(style=name)
        if has_text:
            for im in images:
                im.extract()
            self.emit_inline(para, el, ctx, style, set_font=style.family_explicit,
                             set_size=style.size_explicit, remap=False)
        else:
            text = (el.get("title") or "").strip() or ctx.chapter_title
            if text:
                self.b.add_text(para, text, style, set_font=style.family_explicit,
                                set_size=style.size_explicit, remap=False)

        if not para.runs:
            para._p.getparent().remove(para._p)
        else:
            if style.align in ALIGN:
                para.paragraph_format.alignment = ALIGN[style.align]
            self.b.stats["paragraphs"] += 1

        for im in images:
            self.b.add_block_image(container, im)

    def walk(self, container, node: Tag, ctx: Ctx, style: Computed):
        buffer: list = []

        def flush():
            if not buffer:
                return
            if any((isinstance(n, NavigableString) and n.strip())
                   or isinstance(n, Tag) for n in buffer):
                self.emit_paragraph(container, list(buffer), ctx, style, "Normal", "body")
            buffer.clear()

        for child in node.children:
            if isinstance(child, NavigableString):
                if child.strip():
                    buffer.append(child)
                continue
            if not isinstance(child, Tag):
                continue
            name = child.name.lower()
            if name in SKIP_TAGS or id(child) in ctx.consumed:
                continue
            # 注文都进了脚注，「注释」这个小标题已经无所指，不印。
            if _is_note_label(child):
                continue
            if name in INLINE_TAGS:
                buffer.append(child)
                continue
            flush()

            child_style = compute(child, ctx.sheet, style.child())

            if name in ("h1", "h2", "h3", "h4", "h5", "h6"):
                self.emit_heading(container, child, ctx, child_style)
            elif name == "table":
                self.emit_table(container, child, ctx, child_style)
            elif name == "hr":
                if not _only_consumed_after(child, ctx):
                    container.add_paragraph("＊　　＊　　＊", style="WR 分隔")
            elif name == "img":
                self.b.add_block_image(container, child)
            elif name in ("p", "figcaption", "dt", "dd", "pre", "caption"):
                only = _image_only(child)
                if only:
                    for im in only:
                        self.b.add_block_image(container, im)
                    continue
                role, sname = self._para_style_name(child, child_style)
                if _has_block_child(child):
                    self.walk(container, child, ctx, child_style)
                else:
                    self.emit_paragraph(container, [], ctx, child_style, sname, role, source=child)
            elif name == "li":
                role, sname = self._para_style_name(child, child_style)
                if _has_block_child(child):
                    self.walk(container, child, ctx, child_style)
                else:
                    self.emit_paragraph(container, [], ctx, child_style,
                                        "WR 正文无缩进", role, source=child)
            elif name in ("div", "section", "article", "blockquote", "figure",
                          "center", "ul", "ol", "dl", "body"):
                only = _image_only(child)
                if only and not _has_block_child(child):
                    for im in only:
                        self.b.add_block_image(container, im)
                    continue
                if _has_block_child(child):
                    self.walk(container, child, ctx, child_style)
                else:
                    role, sname = self._para_style_name(child, child_style)
                    if child.get_text(strip=True) or child.find("img"):
                        self.emit_paragraph(container, [], ctx, child_style, sname,
                                            role, source=child)
            else:
                if _has_block_child(child):
                    self.walk(container, child, ctx, child_style)
                else:
                    buffer.append(child)
        flush()

    # ---------------------------------------------------- 表格

    @staticmethod
    def _table_grid(table: Tag):
        rows = table.find_all("tr")
        occupied: dict[tuple[int, int], str] = {}
        origins: list[tuple[int, int, int, int, Tag]] = []   # r, c, rowspan, colspan, cell
        ncols = 0
        for r, tr in enumerate(rows):
            c = 0
            for cell in tr.find_all(["td", "th"], recursive=False):
                while (r, c) in occupied:
                    c += 1
                try:
                    cs = max(1, int(cell.get("colspan", 1)))
                except ValueError:
                    cs = 1
                try:
                    rs = max(1, int(cell.get("rowspan", 1)))
                except ValueError:
                    rs = 1
                for dr in range(rs):
                    for dc in range(cs):
                        occupied[(r + dr, c + dc)] = "x"
                origins.append((r, c, rs, cs, cell))
                c += cs
            ncols = max(ncols, c)
        return len(rows), max(1, ncols), origins

    @staticmethod
    def _set_borders(table):
        tbl_pr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "595959")
            borders.append(el)
        tbl_pr.append(borders)

    def emit_table(self, container, el: Tag, ctx: Ctx, style: Computed):
        nrows, ncols, origins = self._table_grid(el)
        if not nrows or not origins:
            return
        try:
            table = container.add_table(rows=nrows, cols=ncols) if hasattr(container, "add_table") \
                else self.b.doc.add_table(rows=nrows, cols=ncols)
        except Exception:
            return
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        self._set_borders(table)

        for r, c, rs, cs, _cell in origins:
            if rs > 1 or cs > 1:
                try:
                    table.cell(r, c).merge(table.cell(min(r + rs - 1, nrows - 1),
                                                      min(c + cs - 1, ncols - 1)))
                except Exception:
                    pass

        header_rows = 1 if el.find("thead") or (origins and origins[0][4].name == "th") else 0
        for r, c, _rs, _cs, cell in origins:
            try:
                dc = table.cell(r, c)
            except Exception:
                continue
            for p in list(dc.paragraphs):
                p._p.getparent().remove(p._p)
            cell_style = compute(cell, ctx.sheet, style.child())
            if _has_block_child(cell):
                self.walk(dc, cell, ctx, cell_style)
            else:
                para = dc.add_paragraph(style="WR 表格文字")
                self.emit_inline(para, cell, ctx, cell_style)
            if not dc.paragraphs:
                dc.add_paragraph(style="WR 表格文字")
            for p in dc.paragraphs:
                try:
                    p.style = self.b.doc.styles["WR 表格文字"]
                except Exception:
                    pass
                p.paragraph_format.first_line_indent = Pt(0)
                if cell.name == "th" or (header_rows and r == 0):
                    for run in p.runs:
                        run.font.bold = True
                    if cell_style.align is None:
                        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if header_rows:
            tr_pr = table.rows[0]._tr.get_or_add_trPr()
            th = OxmlElement("w:tblHeader")
            th.set(qn("w:val"), "true")
            tr_pr.append(th)

        self.b.stats["tables"] += 1
        try:
            container.add_paragraph(style="WR 正文无缩进")
        except Exception:
            pass


# ---------------------------------------------------------------- 整书组装

HEADING_TAGS = ["h1", "h2", "h3", "h4", "h5", "h6"]


RE_BOOKNAME_FIELD = re.compile(r"^\s*书名\s*[：:]\s*(.+?)\s*$")


def strip_marketing_from_biblio(body: Tag) -> int:
    """版权页的「书名：…」一行按题名清理规则去营销话术与册数，返回改了几处。

    微信读书版的版权页是平台自己排的（「品牌方：…」「本书由…授权微信读书制作与
    发行」都在同一页），它把营销话术一起写进了书名：「智能：AI时代的商业、组织与
    战略的本质（曾鸣教授全新商业思考）」。扉页与文件名用的是清理过的题名，版权页
    再留着同一串就显得自相矛盾，所以这里跟着清理。只动这一种字段行：出版社、
    ISBN、出版时间、版权所有这些一个字不动。没有可清的成分时原样返回。
    """
    n = 0
    for el in body.find_all(["p", "div"]):
        if el.find(["p", "div", "table", "ul", "ol"]):
            continue
        text = el.get_text(" ", strip=True)
        m = RE_BOOKNAME_FIELD.match(text)
        if not m:
            continue
        cleaned = clean_title(m.group(1))
        if cleaned == m.group(1).strip():
            continue
        for node in list(el.children):
            node.extract()
        el.append(text[:m.start(1)] + cleaned)
        n += 1
    return n


def is_continuation(ch: dict) -> bool:
    """这一章是不是「上一章的后半截」。

    超长章节会被 weread 切成两条目录项，后一条没有标题，title 直接就是它自己的
    序号（《三家〈诗〉辑佚史》第六章 14.7 万字，被切成 uid12 与 uid13，
    后者 title 为 "13"）。这种不该单独立标题，更不该占一级大纲。
    """
    title = (ch.get("title") or "").strip()
    return bool(title) and title.isdigit() and title in (
        str(ch.get("chapterIdx", "")), str(ch.get("chapterUid", "")))


RE_NOTE_MARKER = re.compile(r"^[\[〔(（]?\s*\d+\s*[\]〕)）]?$")


def _heading_text(h: Tag) -> str:
    return norm(h.get_text(" ", strip=True) or (h.get("title") or ""))


def _heading_text_no_marker(h: Tag) -> str:
    """标题文字，去掉夹在标题里的注号。

    中华书局这套书把注号直接嵌进章标题，注号还常落在标题中间：
    「给蔡和森<a><sup>[1]</sup></a>的信」。注号会先被抽成 Word 脚注（只留正文标记），
    于是标题文字成了「给蔡和森的信」，不再等于章节表里的「给蔡和森的信」，
    认不出章标题——补出来的目录标题把标题又印了一遍，一、二级各一条。
    普通书里注号多在标题末尾（「在新民学会长沙会员大会上的发言[1]」），
    norm() 去不掉方括号，同样认不出。所以标题比对只认「纯数字注号」的 <a>/<sup>。
    """
    clone = BeautifulSoup(str(h), "html.parser")
    root = clone.find(True) or clone
    for a in list(root.find_all("a")):
        text = a.get_text(strip=True)
        if text and RE_NOTE_MARKER.match(text):
            a.decompose()
    for sup in list(root.find_all("sup")):
        if RE_NOTE_MARKER.match(sup.get_text(strip=True)):
            sup.decompose()
    return norm(root.get_text(" ", strip=True) or (root.get("title") or ""))


def _no_text_between(a: Tag, b: Tag) -> bool:
    """a、b 之间没有正文文字（只隔着装饰图、空段），才算同一条标题拆成了几行。"""
    for item in a.next_elements:
        if item is b:
            return True
        if (isinstance(item, NavigableString) and item.strip()
                and not any(anc is a for anc in item.parents)):
            return False
    return False


def match_title_heads(heads: list[Tag], title: str) -> list[Tag]:
    """章首哪一个（或哪几个相邻的）标题对应章节表里的这条 title。

    多数书一个 h 标签就是整条标题。也有书把「第1章」和「人际关系的构成」拆成两个 h1
    （《亲密关系（第8版）》14 章全这样）。拆开的每一行都只是 title 的一截，按「包含」去认
    会随字数长短认错：四个字的「第10章」被当成整条、章名降成二级；「第3章」「吸引力」都不足
    四个字认不上，补上章节表标题后两截又各成二级；「人际关系的构成」被提成一级，「第1章」
    反倒排到它下面。所以依次认：整条相等 → 相邻几个拼起来相等 → 原来的包含判据。
    """
    nt = norm(title)
    if not nt:
        return []
    # 标题里的注号先去干净（见 _heading_text_no_marker）；纯装饰用的空 <h> 也一并跳过，
    # 否则「相邻标题拼起来相等」会把这些空壳算进去。
    first = [h for h in heads[:2] if _heading_text_no_marker(h)]
    for h in first:
        if _heading_text_no_marker(h) == nt:
            return [h]
    near = [h for h in heads[:4] if _heading_text_no_marker(h)]
    for size in (2, 3):
        for i in range(len(near) - size + 1):
            group = near[i:i + size]
            if ("".join(_heading_text_no_marker(h) for h in group) == nt
                    and all(_no_text_between(x, y) for x, y in zip(group, group[1:]))):
                return group
    for h in first:
        ht = _heading_text_no_marker(h)
        if (len(ht) > 3 and ht in nt) or (len(nt) > 3 and nt in ht):
            return [h]
    return []


def _href_file(href: str) -> str:
    """锚点里的文件名（不带目录）；同文件锚点（#xxx）返回空串。"""
    return (href or "").split("#", 1)[0].rsplit("/", 1)[-1]


RE_CROSS_FILE_ANCHOR = re.compile(r'href="[^"#]+\.x?html?#')


def link_cross_file_notes(chapters: list[dict], soups: dict) -> dict:
    """配对跨文件的注：注号在正文那一节，注文集中在另一个文件里（常是「本章小结」）。

    《亲密关系（第8版）》：正文 `<a href="../Text/Section1_0009.xhtml#note_1" id="noteBack_1">[1]</a>`，
    注文在另一章 `<p><a href="../Text/Section1_0002.xhtml#noteBack_1" id="note_1">[1]</a>…</p>`。
    渲染是一章一章来的，章内配对看不见别的文件，于是注号印成「[1]」、注文成了章末普通段落。
    所以先全书扫一遍：注号指向注文锚点、锚点又回指注号的 id，两边文件名各自归属的章也对得上，
    互指即确证；锚点还得在段首。id 在那一章里不唯一就不配，宁可维持原状。

    chapters 取章节表的 files 认文件归属；soups 是 {chapterUid: 解析好的章}。返回
      markers {uid: {(注号 href, 注号 id): 注文块副本}}：渲染注号那章时并入注释表；
      blocks  {uid: {注文锚点 id}}：渲染注文那章时跳过这些块。
    """
    owner: dict[str, int] = {}
    for ch in chapters:
        for f in ch.get("files") or []:
            owner.setdefault(_href_file(f), ch["chapterUid"])
    markers: dict[int, dict] = {}
    blocks: dict[int, set] = {}
    if not owner:
        return {"markers": markers, "blocks": blocks}

    by_id: dict[int, dict[str, list[Tag]]] = {}
    for uid, soup in soups.items():
        table: dict[str, list[Tag]] = {}
        for el in soup.find_all(id=True):
            table.setdefault(el["id"], []).append(el)
        by_id[uid] = table

    used: set[int] = set()
    for uid, soup in soups.items():
        for a in soup.find_all("a", href=True):
            href, own = a["href"], a.get("id")
            target, _, frag = href.partition("#")
            target = _href_file(target)
            if not (own and frag and target):
                continue
            if not (a.find("sup") or len(a.get_text(strip=True)) <= 6):
                continue
            target_uid = owner.get(target)
            if target_uid == uid:
                continue    # 同一章的两个文件：渲染时去掉锚点文件名，交给章内配对
            found = by_id.get(target_uid, {}).get(frag, [])
            if len(found) != 1 or found[0].name != "a" or id(found[0]) in used:
                continue
            anchor = found[0]
            back, _, back_frag = (anchor.get("href") or "").partition("#")
            if back_frag != own or owner.get(_href_file(back)) != uid:
                continue
            block = anchor.find_parent(["p", "div", "li", "td"])
            if (block is None or not _starts_block(anchor, block)
                    or any(x is a for x in block.descendants)):
                continue
            used.add(id(anchor))
            markers.setdefault(uid, {})[(href, own)] = (
                BeautifulSoup(str(block), "html.parser").find(True))
            blocks.setdefault(target_uid, set()).add(frag)
    return {"markers": markers, "blocks": blocks}


def _note_chars(book: BookBuilder, since: int) -> int:
    """自第 since 条起新加的脚注有多少字（直接数落下的 w:t，不进字符计数器）。

    用于章字数对账：接口的 wordCount 把注文算在内，而正文字符计数不含脚注，
    注比正文还长的短章不加回这一块就会误报偏少。与正文相不相等要分开算：
    每个 <w:t> 两端去空白，免得把 run 之间的缩进换行数进来。
    """
    total = 0
    for paras in book.fnotes.notes[since:]:
        for p_el in paras:
            for t in p_el.iter(qn("w:t")):
                if t.text:
                    total += len(t.text.strip())
    return total


def chapter_sheet(soup, css_text: str) -> Stylesheet:
    sheet = Stylesheet(css_text)
    for st_el in soup.find_all("style"):
        sheet.parse(st_el.get_text())
    return sheet


def render_chapter(book: BookBuilder, renderer: ChapterRenderer, ch: dict,
                   xhtml: str, css_text: str, is_first: bool):
    soup = BeautifulSoup(xhtml, "html.parser")
    # 制书方留在正文里的编辑注释（中信这本第一章等处有 <!-- replace below with real src -->）
    # 不是书的内容，阅读器也不显示。但 bs4 的 Comment 是 NavigableString 的子类，
    # walk() 会把它当裸文本收进段落印出来，所以渲染前先清掉。
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        c.extract()
    drop = VARIANT_CLASS.get(
        "simplified" if book.opts.variant.startswith(("t", "繁")) else "traditional")
    if drop and soup.select("." + drop):
        for el in soup.select("." + drop):
            el.decompose()
        book.stats["variant_dropped"] += 1
        book.dual_variant = True
    # 一章由几个原文件拼成时每个文件各有一个 <body>（《亲密关系（第8版）》的章首 =
    # 标题页 + 导语页）。下面凡是「看全章」的判断都要看全部 body：只看第一个的话，
    # 标题页里没有长段落，黑体转宋体不触发，导语就留在了黑体里。
    bodies = soup.find_all("body") or [soup]
    book.stats["title_fields_cleaned"] = (
        book.stats.get("title_fields_cleaned", 0)
        + sum(strip_marketing_from_biblio(part) for part in bodies))
    sheet = chapter_sheet(soup, css_text)

    title = (ch.get("title") or "").strip()
    level = int(ch.get("level") or 1)

    heads = [h for part in bodies for h in part.find_all(HEADING_TAGS)]
    title_heads = match_title_heads(heads, title)

    rest = [h.name.lower() for h in heads if not any(h is t for t in title_heads)]
    ranks = {tag: i + 1 for i, tag in enumerate(sorted(set(rest)))}

    # Kindle 转制常把整本正文设成某种黑体（如「汉仪旗黑」）。Word 里整本黑体没法读，
    # 所以正文里的黑体按宋体落地，黑体让给标题与强调。
    #
    # 判据用占比而不是「谁最多」：陈寅恪这类书里成段的楷体引文很多，按众数判会被
    # 稀释成「主体是楷体」，于是他自己的正文就留在了黑体里。
    book.family_remap = {}
    all_paras = [el for part in bodies for el in part.find_all("p")]
    fams = [compute(el, sheet, Computed()).family
            for el in all_paras
            if len(el.get_text(strip=True)) > 20]
    # 留下来源证据；按字体比例只能提示，不能把导读、附录的仿宋擅改成宋体。
    samples = {}
    for el in all_paras:
        if len(el.get_text(strip=True)) <= 20:
            continue
        family = compute(el, sheet, Computed()).family
        samples.setdefault(family, {"class": el.get("class", []),
            "css_font_family": sheet.match(el).get("font-family", "继承/默认"),
            "text": el.get_text(strip=True)[:60]})
    book.stats["source_fonts"].append({"chapter": title,
        "long_paragraph_families": dict(Counter(fams)), "samples": samples})
    # 整本正文是黑体时每章都转：按章判会被稀释——章首导语后面跟一长串楷体要点，黑体
    # 只占一成多，那一章就漏了（《亲密关系（第8版）》第14章、《何枝可依》自序）；
    # 索引一类全是短条目，一个长段落都没有，也判不出来。
    if (fams and fams.count("hei") / len(fams) >= 0.2) or book.body_hei:
        book.family_remap = {"hei": "song"}

    # 一章由几个原文件拼成、注号与注文各在一个文件里时，锚点带着文件名
    # （《波斯笔记》第四章：`<a href="../Text/Chapter1-13.html#ch1">`）。章内配对只认 #xxx，
    # 所以指向本章自己文件的锚点先去掉文件名。别章的文件留给 link_cross_file_notes。
    own_files = {_href_file(f) for f in ch.get("files") or []}
    for a in soup.find_all("a", href=True) if own_files else []:
        target, sep, frag = a["href"].partition("#")
        if sep and target and _href_file(target) in own_files:
            a["href"] = "#" + frag

    note_map, _note_blocks = renderer.build_note_map(soup)
    ctx = Ctx(sheet, note_map, ranks, level, title)
    uid = ch.get("chapterUid")
    external = book.cross_notes["markers"].get(uid)
    if external:
        for a in soup.find_all("a", href=True):
            block = external.get((a["href"], a.get("id")))
            if block is not None and id(a) not in note_map:
                note_map[id(a)] = block
    # 原书明说是注号（epub:type/role）却没配上注文的，注文会当普通段落留在章末、不成脚注。
    # 验收只查「已有的脚注对不对」，查不出这种整批漏配，所以在这里点数上报。
    lost = [a for a in soup.find_all("a") if _is_noteref(a) and id(a) not in note_map]
    if lost:
        book.stats["unmatched_noterefs"].append({"title": title, "count": len(lost)})
    skip = book.cross_notes["blocks"].get(uid)
    moved = 0       # 挪去别章当脚注的注文字数，字数偏少的判断里要算回来
    if skip:
        for a in soup.find_all("a", id=True):
            block = a.find_parent(["p", "div", "li", "td"]) if a["id"] in skip else None
            if block is not None and _starts_block(a, block):
                ctx.consumed.add(id(block))
                moved += len(block.get_text(strip=True))

    if is_continuation(ch):
        # 接排上一章，不出标题、不分页
        for part in bodies:
            renderer.walk(book.doc, part, ctx, Computed())
        book.stats["chapters"] += 1
        book.stats["continuations"] += 1
        return 0

    if book.opts.page_break_chapters and level <= 1 and not is_first:
        book.page_break()

    base = Computed()
    style_name = ["WR 标题一", "WR 标题二", "WR 标题三", "WR 标题四"][max(0, min(3, level - 1))]
    # 字数对账的起点。章标题里也可能夹着注号（中华书局那套书就有，且常落在标题中间），
    # 那一条注文会在这里被抽进脚注，字数得算本章的账，所以必须在排标题之前取值——
    # 原先取在后面，标题里的注文就成了「不计入本章」的漏账，整章又不单独补。
    before = book.stats["chars"]
    before_notes = book.stats["footnotes"]
    if title_heads:
        para = book.doc.add_paragraph(style=style_name)
        head_images = []
        for th in title_heads:
            ctx.consumed.add(id(th))
            th_style = compute(th, sheet, base)
            if len(title_heads) > 1:
                # 拆成几行的标题各有字号（小字「第1章」、大字章名），拼回一行后交给标题样式
                th_style = dc_replace(th_style, size_explicit=False)
            # 章标题常常是一张美术字图片，真文字在 title 属性或章节表里。
            # 拿图当标题的话，导航窗格会出现一串没有文字的条目，所以图挪出来单独成段。
            images = [im for im in th.find_all("img") if not is_note_image(im)]
            head_images += images
            if para.runs:
                para.add_run(" ")
            if th.get_text(strip=True):
                for im in images:
                    im.extract()
                renderer.emit_inline(para, th, ctx, th_style,
                                     set_font=th_style.family_explicit,
                                     set_size=th_style.size_explicit, remap=False)
            elif len(title_heads) > 1 and (th.get("title") or "").strip():
                para.add_run(th["title"].strip())
        if not para.runs:
            # 标题是张美术字图片时用真文字补上；字体字号交给标题样式，别在 run 上写死
            para.add_run((title_heads[0].get("title") or "").strip() or title)
        for im in head_images:
            book.add_block_image(book.doc, im)
        book.stats["chapter_heads"].append((title, level))
    elif title:
        book.doc.add_paragraph(title, style=style_name)
        book.stats["chapter_heads"].append((title, level))

    for part in bodies:
        renderer.walk(book.doc, part, ctx, base)
    # 正文的字符统计不含脚注（add_text 在 build_note_map 之外只记正文），
    # 而 wordCount 把注文也算在内。短章（注比正文还长，如「民盟的教训」）
    # 不加回注文字数就会误报偏少。注文只在被渲染时才计入，不会重复。
    got = (book.stats["chars"] - before + moved
           + _note_chars(book, before_notes))

    wc = int(ch.get("wordCount") or 0)
    if wc > 300 and got < wc * 0.6:
        book.stats["short_chapters"].append(
            {"title": title, "expect": wc, "got": got})
    book.stats["chapters"] += 1
    return got


def _scan_simplified(pieces: list[str]) -> list[str]:
    """繁体本里若混着简体字，多半是平台转换没转干净。

    只报告不改原文——擅自转换等于改动古籍。而且 opencc 的 s2t 偏激进，
    会把繁体里本就通用的异体字（迹/蹟、周/週、搜/蒐、别/別）也判成简体，
    所以这里的结果是「疑似」，要由人来甄别。"""
    text = "".join(pieces)
    if not text:
        return []
    converted = to_traditional(text)
    if len(converted) != len(text):
        return []      # 长度对不上就没法逐字比，宁可不报
    samples, seen = [], set()
    for i, (a, b) in enumerate(zip(text, converted)):
        if a != b and a not in seen:
            seen.add(a)
            samples.append(f"{a}→{b}　…{text[max(0, i - 12):i]}【{a}】{text[i + 1:i + 12]}…")
            if len(samples) >= 8:
                break
    return samples


def build_book(work_dir: Path, out_path: Path, opts: Options, cookies: str = "") -> dict:
    meta = json.loads((work_dir / "book.json").read_text(encoding="utf-8"))
    prog_path = work_dir / "progress.json"
    progress = json.loads(prog_path.read_text(encoding="utf-8")) if prog_path.exists() else {}

    chapters = [c for c in meta["chapters"]
                if progress.get(str(c["chapterUid"]), {}).get("ok")]
    if not chapters:
        raise RuntimeError("没有已抓好的章节，先跑 weread fetch")

    opts.title = opts.title or meta.get("title", "")
    opts.author = opts.author or meta.get("author", "")

    book = BookBuilder(work_dir, meta, opts, cookies)
    renderer = ChapterRenderer(book)

    # 先把全书的私用区码位一次性还原：靠本机字体的 cmap 等价关系查真字
    # （某些字体会把私用区码位和真码位映到同一字形）。换不出来的原样保留并高亮。
    found: set[int] = set()
    for ch in chapters:
        rec = progress[str(ch["chapterUid"])]
        xp = work_dir / rec["file"]
        if xp.exists():
            raw = xp.read_text(encoding="utf-8")
            found |= pua.scan_text(raw)
            # 扉页在渲染正文之前就要落笔，所以繁简版本得先探明
            if "wr-simplified" in raw:
                book.dual_variant = True
    if found:
        resolved = pua.resolve(found)
        book.pua_table = {cp: r["char"] for cp, r in resolved.items()}

        # 人工指定：机器找不到字形证据时，由人判读后记进 .state/pua_manual.json
        # （一份文件管所有书，按 bookId 归档）。决定权在人、结论留痕、重抓也不会丢。
        manual = pua.load_manual(meta.get("bookId", ""))
        for cp, item in manual.items():
            book.pua_table[cp] = item["char"]
            book.pua_manual_cps.add(cp)
            book.stats["pua_manual"][f"U+{cp:04X}"] = item["char"] + (
                f"（{item['note']}）" if item.get("note") else "")
        book.pua_unresolved = found - set(book.pua_table)
        book.stats["pua_map"] = {f"U+{cp:04X}": f"{r['char']} (U+{r['to']:04X})"
                                 for cp, r in sorted(resolved.items())}
        for cp in sorted(book.pua_unresolved):
            book.stats["pua_left"].append(f"U+{cp:04X}")

    cover_path = None
    if meta.get("cover"):
        cover_path = book.images.get(meta["cover"])
        if cover_path is None:
            book.stats["cover"] = "下载失败"
    book.title_page(cover_path)

    listed = [c for c in chapters
              if not (opts.keep_cover and (c.get("title") or "").strip()
                      in ("封面", "Cover", "cover") and cover_path)]
    # 书前目录要照正文实际排出的标题，所以先占位，正文排完再回填
    toc_at = book.toc_anchor() if (opts.toc and listed) else None

    # 渲染前先全书扫一遍，有两件事一章一章看不出来：
    #   注号与注文分在两章的注，得先配好，渲染时各章才看得见对面；
    #   正文是不是整本黑体（判据同 render_chapter，只是放到全书的长段落上算）。
    drop = VARIANT_CLASS.get("simplified" if opts.variant.startswith(("t", "繁")) else "traditional")
    soups, book_fams, last_css = {}, Counter(), ""
    for ch in listed:
        xp = work_dir / progress[str(ch["chapterUid"])]["file"]
        if not xp.exists():
            continue
        cpath = work_dir / "css" / (xp.stem + ".css")
        css_text = cpath.read_text(encoding="utf-8") if cpath.exists() else last_css
        if cpath.exists():
            last_css = css_text
        raw = xp.read_text(encoding="utf-8")
        soup = BeautifulSoup(raw, "html.parser")
        for el in soup.select("." + drop) if drop else []:
            el.decompose()
        sheet = chapter_sheet(soup, css_text)
        book_fams.update(compute(el, sheet, Computed()).family for el in soup.find_all("p")
                         if len(el.get_text(strip=True)) > 20)
        if RE_CROSS_FILE_ANCHOR.search(raw):
            soups[ch["chapterUid"]] = soup
    book.cross_notes = link_cross_file_notes(listed, soups)
    book.stats["cross_file_notes"] = sum(len(v) for v in book.cross_notes["markers"].values())
    del soups
    total = sum(book_fams.values())
    book.body_hei = bool(total) and book_fams["hei"] / total >= 0.2
    book.stats["book_fonts"] = dict(book_fams)

    last_css = ""
    first_body = True
    # 有的章平台不下发原文，给的是别章的 XHTML（《发展的自我》「版权声明」下发的是「版权信息」，
    # 阅读器上也这么显示，重抓无用）。原样排会重复一遍，所以略去并报出来。
    seen: dict[str, str] = {}
    for ch in chapters:
        uid = ch["chapterUid"]
        rec = progress[str(uid)]
        xpath = work_dir / rec["file"]
        if not xpath.exists():
            continue
        title = (ch.get("title") or "").strip()
        if opts.keep_cover and title in ("封面", "Cover", "cover") and cover_path:
            continue
        raw = xpath.read_text(encoding="utf-8")
        if raw in seen:
            book.stats["platform_dups"].append({"title": title, "same_as": seen[raw]})
            continue
        seen[raw] = title
        cpath = work_dir / "css" / (xpath.stem + ".css")
        css_text = cpath.read_text(encoding="utf-8") if cpath.exists() else last_css
        if cpath.exists():
            last_css = css_text
        render_chapter(book, renderer, ch, raw, css_text, first_body)
        first_body = False
    if toc_at is not None:
        depth = max([int(c.get("level") or 1) for c in listed if not is_continuation(c)] or [1])
        book.stats["toc_entries"] = book.fill_toc(toc_at, max(1, min(4, depth)))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    # 先验收候选文件，通过后才替换成品；失败不会覆盖已有成果。
    fd, staged_name = tempfile.mkstemp(prefix=".weread-pending-", suffix=".docx", dir=out_path.parent)
    os.close(fd)
    staged_path = Path(staged_name)
    n = book.fnotes.save(book.doc, staged_path)
    book.stats["footnotes"] = n
    if book.chain:
        book.stats["font_borrowed"] = dict(book.chain.borrowed)
        book.stats["font_missing"] = sorted(book.chain.missing)
    if book.dual_variant and opts.variant.startswith(("t", "繁")):
        book.stats["simplified_left"] = _scan_simplified(book.body_text)
    book.stats["font_note"] = book.font_note
    book.stats["image_failures"] = len(book.images.failed)
    # 检查真实落盘的 OOXML，不能把进程退出或统计数字当成验收。
    from lib.docxcheck import audit, heading_mismatches
    report = audit(staged_path)
    report["bookId"] = str(meta.get("bookId", ""))
    report["source_chapters"] = len(meta["chapters"])
    report["captured_chapters"] = len(chapters)
    report["rendered_chapters"] = book.stats["chapters"]
    report["source_fonts"] = book.stats["source_fonts"]
    report["book_fonts"] = book.stats["book_fonts"]
    report["cover"] = book.stats["cover"]
    if meta.get("cover") and book.stats["cover"] != "嵌入":
        report["warnings"].append(f"扉页没有封面：{book.stats['cover']}")
    for ch in chapters:
        if not (work_dir / progress[str(ch["chapterUid"])]["file"]).exists():
            report["errors"].append(f"成功台账对应的 XHTML 缺失：{ch['title']}")
    if len(chapters) < len(meta["chapters"]):
        report["warnings"].append(f"仅有 {len(chapters)}/{len(meta['chapters'])} 章缓存成功；不能按整书交付")
    if book.stats["image_failures"]:
        report["errors"].append(f"图片下载失败 {book.stats['image_failures']} 处")
    for item in book.stats["short_chapters"]:
        report["warnings"].append(f"章节字数偏少：{item['title']}")
    for item in book.stats["unmatched_noterefs"]:
        report["warnings"].append(
            f"注号没配上注文 {item['count']} 处：{item['title']}（注文按普通段落留在章末，未成脚注）")
    for item in book.stats["platform_dups"]:
        report["warnings"].append(
            f"平台未下发原文：{item['title']}（给的是「{item['same_as']}」，已略去）")
    mismatched = heading_mismatches(staged_path, book.stats["chapter_heads"])
    if mismatched:
        report["warnings"].append(
            f"章节标题与章节表对不上 {len(mismatched)} 处（{'；'.join(mismatched[:3])}"
            f"{'…' if len(mismatched) > 3 else ''}）；侧栏大纲会跟着错，须核实")
        report["heading_mismatches"] = mismatched
    if not report["errors"]:
        staged_path.replace(out_path)
        report["path"] = str(out_path.resolve())
    else:
        # 候选件挪回任务目录留着查。原先就地留在成品目录，是个隐藏的 .weread-pending-*.docx，
        # 失败一次剩一个，谁也看不见。
        import shutil
        rejected = work_dir / "docx-rejected.docx"
        shutil.move(str(staged_path), str(rejected))
        report["path"] = str(rejected.resolve())
    report_path = work_dir / "docx-audit.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    book.stats["audit"] = report
    book.stats["audit_path"] = str(report_path)
    if report["errors"]:
        raise RuntimeError("DOCX 验收失败：" + "；".join(report["errors"])
                           + f"；报告 {report_path}；候选件 {report['path']}")
    return book.stats
