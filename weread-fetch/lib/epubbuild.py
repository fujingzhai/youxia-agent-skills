#!/usr/bin/env python3
"""把逐章取回的材料重新压回一个 EPUB。

抓下来的东西本来就是书在服务端那份 EPUB 的零件：xhtml 是包里的原文件，
css 是原书样式表，图片在 res.weread.qq.com 上。缺的只是外壳——
mimetype、container.xml、content.opf、目录。这里就是把壳补上。

EPUB 是归档件，力求贴近原书；docx 是工作副本，那边才做脚注、图宽这类改写。

一处补不了：原书嵌入的字体（@font-face 指向的 ../Fonts/*.ttf）服务端不提供，
微信读书自己的阅读器也不下载它们。后果是靠私用区码位（PUA）编码的生僻字
在任何阅读器里都会是豆腐块。生成时会把这类字符点出来。
"""

from __future__ import annotations

import hashlib
import html
import json
import re
import urllib.parse
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from bs4 import BeautifulSoup, NavigableString

RES_BASE = "https://res.weread.qq.com/wrepub/web"

# QQ阅读体系把夹注做成一张小图标、注文塞在 alt 里。这个约定除了微信读书自己
# 没人认得，而且那张 note.png 对普通出版书根本取不到（实测 404）。所以打包时
# 把它还原成 EPUB3 标准脚注，阅读器就能正常弹注了。
RE_NOTE_IMG_CLASS = re.compile(r"(qqreader-)?footnote|comment|annot", re.I)
RE_NOTE_IMG_SRC = re.compile(r"(^|/)(note|footnote|comment|annot)\w*\.(png|gif|jpg|svg)$", re.I)


def _is_note_image(img) -> bool:
    classes = " ".join(img.get("class") or [])
    return bool(RE_NOTE_IMG_CLASS.search(classes)
                or RE_NOTE_IMG_SRC.search(img.get("src") or ""))

# 私用区与罕用扩展区：没有原书字体就渲染不出来
_PUA_RANGES = ((0xE000, 0xF8FF), (0xF0000, 0xFFFFD), (0x100000, 0x10FFFD))


def is_pua(cp: int) -> bool:
    return any(lo <= cp <= hi for lo, hi in _PUA_RANGES)


@dataclass
class EpubStats:
    chapters: int = 0
    images: int = 0
    css_files: int = 0
    fonts_embedded: int = 0
    fonts_missing: list = field(default_factory=list)
    footnotes: int = 0
    anchors_repaired: int = 0
    image_failures: list = field(default_factory=list)
    pua: list = field(default_factory=list)     # [(字符, 章标题, 上下文)]


def _safe_name(text: str, fallback: str) -> str:
    name = re.sub(r"[^\w.\-]+", "_", text or "")
    return name[:60] or fallback


class EpubBuilder:
    def __init__(self, work_dir: Path, meta: dict, images):
        self.work = work_dir
        self.meta = meta
        self.images = images          # 复用 docxbuild.ImageStore，带 Cookie 与缓存
        self.book_id = str(meta["bookId"])
        self.stats = EpubStats()
        self.files: dict[str, bytes] = {}          # zip 内路径 -> 内容
        self.manifest: list[tuple[str, str, str]] = []   # (id, href, media-type)
        self.spine: list[str] = []
        self._css_by_hash: dict[str, str] = {}
        self._img_by_url: dict[str, str] = {}

    # ---------------------------------------------------------- 资源

    def _media_type(self, name: str) -> str:
        ext = name.rsplit(".", 1)[-1].lower()
        return {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                "gif": "image/gif", "svg": "image/svg+xml", "webp": "image/webp",
                "css": "text/css", "xhtml": "application/xhtml+xml",
                "ncx": "application/x-dtbncx+xml"}.get(ext, "application/octet-stream")

    def _add(self, href: str, data: bytes, item_id: str | None = None) -> str:
        self.files["OEBPS/" + href] = data
        item_id = item_id or "id_" + hashlib.md5(href.encode()).hexdigest()[:10]
        self.manifest.append((item_id, href, self._media_type(href)))
        return item_id

    def resolve_image(self, src: str) -> str | None:
        """把正文里的图片地址换成包内路径，顺便把图download下来。

        绝对地址直接下；相对地址（../Images/x.jpg）按阅读器自己的规则拼成
        res.weread.qq.com/wrepub/web/<bookId>/x.jpg —— 这条是抓阅读器请求实测出来的。
        """
        src = (src or "").strip()
        if not src:
            return None
        if src.startswith("//"):
            url = "https:" + src
        elif src.startswith("http"):
            url = src
        else:
            url = f"{RES_BASE}/{self.book_id}/{src.split('/')[-1]}"

        if url in self._img_by_url:
            return self._img_by_url[url]   # 含失败的 None，不重复请求
        path = self.images.get(url)
        if path is None:
            if url not in self.stats.image_failures:
                self.stats.image_failures.append(url)
            self._img_by_url[url] = None
            return None
        href = "Images/" + _safe_name(path.name, "img") 
        if "OEBPS/" + href not in self.files:
            self._add(href, path.read_bytes())
            self.stats.images += 1
        self._img_by_url[url] = href
        return href

    _FONT_FACE = re.compile(r"@font-face\s*\{[^{}]*\}", re.S)
    # 字体名里可能带括号（FZKaiS(SIP)B.ttf），所以先认带引号的整体，别按 ) 截断
    _URL = re.compile(r"""url\(\s*(?:"([^"]*)"|'([^']*)'|([^)\s]*))\s*\)""")

    @classmethod
    def _urls_in(cls, text: str) -> list[str]:
        return [next(g for g in m.groups() if g) for m in cls._URL.finditer(text)
                if any(m.groups())]

    def embed_fonts(self, css_text: str) -> str:
        """把原书 @font-face 引用的字体取回来嵌进包里。

        字体和相对图片同一个基址：res.weread.qq.com/wrepub/web/<bookId>/<文件名>。
        但服务端只放了常用子集——名字带 ext02/ext15 的扩展面子集一律 404，
        而第 15 面正是私用区，所以靠 PUA 编码的生僻字仍然无解（见 SKILL.md）。
        取不到的 @font-face 整块删掉，留着只会变成坏引用。
        """
        def replace_block(m):
            block = m.group(0)
            body = block[block.index("{") + 1: block.rindex("}")]
            decls, order = {}, []
            for item in body.split(";"):
                if ":" not in item:
                    continue
                key, _, val = item.partition(":")
                key = key.strip().lower()
                if key not in decls:
                    order.append(key)
                decls[key] = val.strip()

            refs = [u for u in self._urls_in(decls.get("src", ""))
                    if u.lower().rsplit(".", 1)[-1] in ("ttf", "otf", "woff", "woff2", "ttc")]
            if not refs:
                return block

            local = []
            for ref in refs:
                name = ref.split("/")[-1]
                path = self.images.get(
                    f"{RES_BASE}/{self.book_id}/{urllib.parse.quote(name)}")
                if path is None:
                    if name not in self.stats.fonts_missing:
                        self.stats.fonts_missing.append(name)
                    continue
                href = "Fonts/" + _safe_name(name, "font.ttf")
                if "OEBPS/" + href not in self.files:
                    self._add(href, path.read_bytes())
                    self.stats.fonts_embedded += 1
                local.append(href)

            if not local:
                return ""     # 一个都没取到就整块删掉，留着只是坏引用
            decls["src"] = ", ".join(f'url("../{h}")' for h in local)
            return ("@font-face {\n"
                    + "".join(f"  {k}: {decls[k]};\n" for k in order)
                    + "}\n")

        return self._FONT_FACE.sub(replace_block, css_text)

    def resolve_css_images(self, css_text: str) -> str:
        """样式表里也可能引用图片（background-image 之类），一并取回改成包内路径。"""
        def repl(m):
            ref = next((g for g in m.groups() if g), "")
            if not ref or ref.startswith(("http", "data:", "../Fonts/")):
                return m.group(0)
            if ref.lower().rsplit(".", 1)[-1] not in ("png", "jpg", "jpeg", "gif", "svg", "webp"):
                return m.group(0)
            href = self.resolve_image(ref)
            return f'url("../{href}")' if href else "none"

        return self._URL.sub(repl, css_text)

    def add_css(self, css_text: str) -> str:
        """同样内容的样式表只存一份。"""
        css_text = self.resolve_css_images(self.embed_fonts(css_text))
        key = hashlib.md5(css_text.encode("utf-8")).hexdigest()[:12]
        if key in self._css_by_hash:
            return self._css_by_hash[key]
        href = f"Styles/style_{key}.css"
        self._add(href, css_text.encode("utf-8"))
        self._css_by_hash[key] = href
        self.stats.css_files += 1
        return href

    # ---------------------------------------------------------- 章节

    def _scan_pua(self, text: str, title: str) -> None:
        for i, ch in enumerate(text):
            if is_pua(ord(ch)):
                ctx = text[max(0, i - 18):i] + "▮" + text[i + 1:i + 18]
                self.stats.pua.append((f"U+{ord(ch):04X}", title, ctx.replace("\n", " ")))

    def add_chapter(self, idx: int, title: str, xhtml: str, css_text: str,
                    variant: str = "traditional") -> str:
        soup = BeautifulSoup(xhtml, "html.parser")
        # 繁体底本的书会把繁简两套一起下发，只留一套
        drop = "wr-simplified" if variant.startswith(("t", "繁")) else "wr-traditional"
        for el in soup.select("." + drop):
            el.decompose()
        bodies = soup.find_all("body")

        # 一章可能由原书好几个 xhtml 拼来，合成一个 body
        if len(bodies) > 1:
            main = bodies[0]
            for extra in bodies[1:]:
                for child in list(extra.contents):
                    main.append(child.extract())
                extra.decompose()
            bodies = [main]
        body = bodies[0] if bodies else soup

        notes = self._convert_alt_footnotes(soup, body, idx)
        notes += self._mark_anchor_footnotes(body)

        for img in body.find_all("img"):
            href = self.resolve_image(img.get("src") or img.get("data-src") or "")
            if href:
                img["src"] = "../" + href
            else:
                img.decompose()      # 取不到的图直接摘掉，别留下坏 src
        self.stats.footnotes += notes

        self._scan_pua(body.get_text(" ", strip=True), title)

        css_href = self.add_css(css_text) if css_text.strip() else None
        head = ["<title>" + html.escape(title or f"第 {idx} 章") + "</title>"]
        if css_href:
            head.append(f'<link rel="stylesheet" type="text/css" href="../{css_href}"/>')

        doc = ('<?xml version="1.0" encoding="utf-8"?>\n'
               '<!DOCTYPE html>\n'
               '<html xmlns="http://www.w3.org/1999/xhtml" '
               'xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="zh">\n'
               '<head><meta charset="utf-8"/>' + "".join(head) + '</head>\n'
               + str(body).replace("<body", "<body", 1) + '\n</html>\n')

        href = f"Text/ch{idx:04d}.xhtml"
        item_id = f"ch{idx:04d}"
        self._add(href, doc.encode("utf-8"), item_id)
        self.spine.append(item_id)
        self.stats.chapters += 1
        return href

    def _convert_alt_footnotes(self, soup, body, chapter_idx: int) -> int:
        """把「注文藏在 img alt 里」的夹注换成 EPUB3 的 noteref + aside。"""
        marks = [im for im in body.find_all("img")
                 if _is_note_image(im) and (im.get("alt") or "").strip()]

        # 没有注文的注释图标：若它本身就是可点的注号（外面套着 <a href="#…">），
        # 不能直接删——删了注号就成了空锚点，点不动。换成可见的序号。
        plain = [im for im in body.find_all("img")
                 if _is_note_image(im) and not (im.get("alt") or "").strip()]
        seq = 0
        for im in plain:
            link = im.find_parent("a")
            if link is not None and (link.get("href") or "").startswith("#"):
                seq += 1
                sup = soup.new_tag("sup")
                sup.string = f"({seq})"
                im.replace_with(sup)
            else:
                im.decompose()      # 纯装饰图标，原书那张 note.png 多数还取不到

        if not marks:
            return 0

        asides = []
        for n, im in enumerate(marks, 1):
            note_id = f"fn{chapter_idx}_{n}"
            ref_id = f"fnref{chapter_idx}_{n}"
            text = (im.get("alt") or "").strip()

            a = soup.new_tag("a", href=f"#{note_id}", id=ref_id)
            a["epub:type"] = "noteref"
            sup = soup.new_tag("sup")
            sup.string = f"({n})"
            a.append(sup)
            im.replace_with(a)

            aside = soup.new_tag("aside", id=note_id)
            aside["epub:type"] = "footnote"
            para = soup.new_tag("p")
            back = soup.new_tag("a", href=f"#{ref_id}")
            back.string = f"({n})"
            para.append(back)
            para.append(NavigableString(" " + text))
            aside.append(para)
            asides.append(aside)

        holder = soup.new_tag("section")
        holder["epub:type"] = "footnotes"
        holder["class"] = "wr-footnotes"
        for aside in asides:
            holder.append(aside)
        body.append(holder)
        return len(asides)

    @staticmethod
    def _anchor_ids(body) -> set:
        """整章已有的 id 集合。判断注号是否指空必须看全章，
        只看注文块内部会误判——注文的 id 常挂在外层 <li> 上。"""
        return {el.get("id") for el in body.find_all(True) if el.get("id")}

    def _mark_anchor_footnotes(self, body) -> int:
        """处理锚点式注（注号 <a href="#x"> + 章末注文块）。

        做两件事：
        1. 打上 epub:type，EPUB3 阅读器才会做成弹注，不必来回翻页。
        2. 补断掉的锚点。有些原书的注文块只写了回跳 href、没写 id，
           注号点过去落空（《波斯笔记·波斯王陵》整章 21 条都是这样）。
           配对关系在上一步已经算出来了，这里顺手把 id 补回去。

        配对逻辑与 docx 那边共用同一个函数，免得两套规则走偏。"""
        from lib.docxbuild import build_note_map

        mapping, _ = build_note_map(body)
        existing = self._anchor_ids(body)
        marked = repaired = 0
        for marker in body.find_all("a"):
            target = mapping.get(id(marker))
            if target is None:
                continue
            want = (marker.get("href") or "#")[1:]
            if want and want not in existing:
                if not target.get("id"):
                    target["id"] = want
                else:
                    inner = target.find("a")
                    if inner is not None and not inner.get("id"):
                        inner["id"] = want
                    else:
                        continue
                existing.add(want)
                repaired += 1
            if not marker.get("id"):
                back = next((a.get("href", "")[1:] for a in target.find_all("a")
                             if (a.get("href") or "").startswith("#")), "")
                new_id = back or (want + "-back")
                if new_id not in existing:
                    marker["id"] = new_id
                    existing.add(new_id)
            marker["epub:type"] = "noteref"
            target["epub:type"] = "footnote"
            marked += 1
        self.stats.anchors_repaired += repaired
        return marked

    # ---------------------------------------------------------- 外壳

    def _nav(self, toc: list[dict]) -> str:
        def render(items, depth=1):
            out = ["  " * depth + "<ol>"]
            for it in items:
                out.append("  " * depth + f'<li><a href="{it["href"]}">'
                                          f'{html.escape(it["title"])}</a>')
                if it["children"]:
                    out.append(render(it["children"], depth + 1))
                out.append("  " * depth + "</li>")
            out.append("  " * depth + "</ol>")
            return "\n".join(out)

        return ('<?xml version="1.0" encoding="utf-8"?>\n'
                '<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" '
                'xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="zh">\n'
                '<head><meta charset="utf-8"/><title>目录</title></head>\n<body>\n'
                '<nav epub:type="toc" id="toc"><h1>目录</h1>\n'
                + render(toc) + '\n</nav>\n</body>\n</html>\n')

    def _ncx(self, toc: list[dict], uid: str) -> str:
        counter = [0]

        def render(items, depth=1):
            out = []
            for it in items:
                counter[0] += 1
                n = counter[0]
                out.append(f'<navPoint id="np{n}" playOrder="{n}">'
                           f'<navLabel><text>{html.escape(it["title"])}</text></navLabel>'
                           f'<content src="{it["href"]}"/>')
                if it["children"]:
                    out.append(render(it["children"], depth + 1))
                out.append("</navPoint>")
            return "".join(out)

        return ('<?xml version="1.0" encoding="utf-8"?>\n'
                '<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">\n'
                f'<head><meta name="dtb:uid" content="{uid}"/></head>\n'
                f'<docTitle><text>{html.escape(self.meta.get("title", ""))}</text></docTitle>\n'
                '<navMap>' + render(toc) + '</navMap>\n</ncx>\n')

    def _opf(self, uid: str, cover_id: str | None) -> str:
        items = "\n".join(
            f'    <item id="{i}" href="{h}" media-type="{m}"'
            + (' properties="cover-image"' if i == cover_id else "") + "/>"
            for i, h, m in self.manifest)
        spine = "\n".join(f'    <itemref idref="{i}"/>' for i in self.spine)
        meta_cover = f'    <meta name="cover" content="{cover_id}"/>\n' if cover_id else ""
        return ('<?xml version="1.0" encoding="utf-8"?>\n'
                '<package xmlns="http://www.idpf.org/2007/opf" version="3.0" '
                'unique-identifier="bookid">\n'
                '  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">\n'
                f'    <dc:identifier id="bookid">{uid}</dc:identifier>\n'
                f'    <dc:title>{html.escape(self.meta.get("title", ""))}</dc:title>\n'
                f'    <dc:creator>{html.escape(self.meta.get("author", ""))}</dc:creator>\n'
                '    <dc:language>zh</dc:language>\n'
                f'    <dc:source>weread:{self.book_id}</dc:source>\n'
                + meta_cover +
                '  </metadata>\n  <manifest>\n'
                '    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" '
                'properties="nav"/>\n'
                '    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>\n'
                + items + '\n  </manifest>\n  <spine toc="ncx">\n' + spine +
                '\n  </spine>\n</package>\n')

    def write(self, out_path: Path, toc: list[dict], cover_href: str | None) -> EpubStats:
        uid = "urn:uuid:" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"weread:{self.book_id}"))
        cover_id = None
        if cover_href:
            cover_id = next((i for i, h, _ in self.manifest if h == cover_href), None)

        self.files["OEBPS/nav.xhtml"] = self._nav(toc).encode("utf-8")
        self.files["OEBPS/toc.ncx"] = self._ncx(toc, uid).encode("utf-8")
        self.files["OEBPS/content.opf"] = self._opf(uid, cover_id).encode("utf-8")
        self.files["META-INF/container.xml"] = (
            '<?xml version="1.0" encoding="utf-8"?>\n'
            '<container version="1.0" '
            'xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n'
            '  <rootfiles><rootfile full-path="OEBPS/content.opf" '
            'media-type="application/oebps-package+xml"/></rootfiles>\n'
            '</container>\n').encode("utf-8")

        out_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            # mimetype 必须是第一个条目且不压缩，否则部分阅读器不认
            z.writestr(zipfile.ZipInfo("mimetype"), "application/epub+zip",
                       compress_type=zipfile.ZIP_STORED)
            for name, data in self.files.items():
                z.writestr(name, data)
        return self.stats


def build_epub(work_dir: Path, out_path: Path, cookies: str = "",
               variant: str = "traditional") -> EpubStats:
    from lib.docxbuild import ImageStore

    meta = json.loads((work_dir / "book.json").read_text(encoding="utf-8"))
    prog_path = work_dir / "progress.json"
    progress = json.loads(prog_path.read_text(encoding="utf-8")) if prog_path.exists() else {}
    chapters = [c for c in meta["chapters"]
                if progress.get(str(c["chapterUid"]), {}).get("ok")]
    if not chapters:
        raise RuntimeError("没有已抓好的章节，先跑 weread fetch")

    builder = EpubBuilder(work_dir, meta, ImageStore(work_dir, cookies))

    cover_href = None
    if meta.get("cover"):
        cover_href = builder.resolve_image(meta["cover"])

    toc: list[dict] = []
    stack: list[tuple[int, dict]] = []
    last_css = ""
    for n, ch in enumerate(chapters, 1):
        rec = progress[str(ch["chapterUid"])]
        xpath = work_dir / rec["file"]
        if not xpath.exists():
            continue
        cpath = work_dir / "css" / (xpath.stem + ".css")
        css_text = cpath.read_text(encoding="utf-8") if cpath.exists() else last_css
        if cpath.exists():
            last_css = css_text
        title = (ch.get("title") or f"第 {n} 章").strip()
        href = builder.add_chapter(n, title, xpath.read_text(encoding="utf-8"),
                                   css_text, variant=variant)

        node = {"title": title, "href": href, "children": []}
        level = max(1, int(ch.get("level") or 1))
        while stack and stack[-1][0] >= level:
            stack.pop()
        (stack[-1][1]["children"] if stack else toc).append(node)
        stack.append((level, node))

    return builder.write(out_path, toc, cover_href)
