#!/usr/bin/env python3
"""给 python-docx 补上真正的 Word 脚注。

python-docx 不支持脚注。这里的做法是：脚注段落先照常用 python-docx 建在正文
末尾（这样字体、样式、缩进全部复用正文那套渲染代码），登记好之后再从
document.xml 里摘走，单独写进 word/footnotes.xml，并补齐 content-type、
关系项和 settings 里的 footnotePr。

脚注分部原本没有自己的关系分部，图片会失效；现在会照着 document.xml.rels
另建一份 word/_rels/footnotes.xml.rels（图片本体是共用的，只需复制关系项），
所以注文里的生僻字图能保住。w:hyperlink 仍会拆成普通文字。
"""

from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import nsmap, qn
from docx.oxml import parse_xml
from docx.shared import Pt

W = nsmap["w"]

_NS_DECL = " ".join(
    f'xmlns:{p}="{u}"' for p, u in (
        ("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
        ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
    )
)

FOOTNOTES_CT = ("application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.footnotes+xml")
FOOTNOTES_REL = ("http://schemas.openxmlformats.org/officeDocument"
                 "/2006/relationships/footnotes")

STYLE_TEXT = "Footnote Text"
STYLE_REF = "Footnote Reference"


def ensure_styles(doc, body_font_cn: str, body_font_en: str, size_pt: float):
    """建好脚注正文段落样式与脚注号字符样式。"""
    styles = doc.styles

    def get_or_add(name, kind):
        try:
            return styles[name]
        except KeyError:
            return styles.add_style(name, kind)

    text = get_or_add(STYLE_TEXT, WD_STYLE_TYPE.PARAGRAPH)
    text.font.size = Pt(size_pt)
    text.font.name = body_font_en
    text.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), body_font_cn)
    pf = text.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.left_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15

    ref = get_or_add(STYLE_REF, WD_STYLE_TYPE.CHARACTER)
    ref.font.superscript = True
    return text, ref


class FootnoteManager:
    def __init__(self, doc):
        self.doc = doc
        self.notes: list[list] = []      # 每条脚注 = 若干 <w:p> 元素
        self.enabled = True

    # ---- 建脚注

    def new_note(self) -> int:
        self.notes.append([])
        return len(self.notes)           # 正文里用的 w:id，从 1 起

    def attach(self, note_id: int, paragraph) -> None:
        self.notes[note_id - 1].append(paragraph._p)

    def add_reference(self, paragraph, note_id: int) -> None:
        """在正文段落末尾插入脚注引用号。"""
        xml = (f'<w:r {_NS_DECL}><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
               f'<w:footnoteReference w:id="{note_id}"/></w:r>')
        paragraph._p.append(parse_xml(xml))

    # ---- 摘出并落盘

    def _sanitize(self, p_el):
        # 图保留：关系项会在 _inject 里补进 footnotes.xml.rels
        for link in list(p_el.findall(f".//{{{W}}}hyperlink")):
            parent = link.getparent()
            index = list(parent).index(link)
            for child in list(link):
                parent.insert(index, child)
                index += 1
            parent.remove(link)
        return p_el

    def _embed_ids(self) -> set[str]:
        """脚注里用到的图片关系 id，用来给 footnotes.xml 建自己的关系分部。"""
        R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        ids = set()
        for paras in self.notes:
            for p_el in paras:
                for el in p_el.iter():
                    for attr in (f"{{{R}}}embed", f"{{{R}}}link", f"{{{R}}}id"):
                        val = el.get(attr)
                        if val:
                            ids.add(val)
        return ids

    def _footnote_xml(self) -> str:
        from lxml import etree

        parts = [f'<w:footnotes {_NS_DECL}>']
        parts.append('<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>'
                     '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                     '<w:r><w:separator/></w:r></w:p></w:footnote>')
        parts.append('<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>'
                     '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                     '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')

        for i, paras in enumerate(self.notes, 1):
            parts.append(f'<w:footnote w:id="{i}">')
            if not paras:
                parts.append('<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
                             '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                             '<w:footnoteRef/></w:r></w:p>')
            for j, p_el in enumerate(paras):
                self._sanitize(p_el)
                if j == 0:
                    marker = parse_xml(
                        f'<w:r {_NS_DECL}><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                        f'<w:footnoteRef/></w:r>')
                    spacer = parse_xml(f'<w:r {_NS_DECL}><w:t xml:space="preserve"> </w:t></w:r>')
                    pPr = p_el.find(f"{{{W}}}pPr")
                    at = list(p_el).index(pPr) + 1 if pPr is not None else 0
                    p_el.insert(at, spacer)
                    p_el.insert(at, marker)
                xml = etree.tostring(p_el, encoding="unicode")
                xml = re.sub(r'\sxmlns:[a-zA-Z0-9]+="[^"]*"', "", xml)
                parts.append(xml)
            parts.append("</w:footnote>")
        parts.append("</w:footnotes>")
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + "".join(parts)

    def detach_from_body(self):
        """把脚注段落从正文里摘掉（内容已被 _footnote_xml 引用）。"""
        for paras in self.notes:
            for p_el in paras:
                parent = p_el.getparent()
                if parent is not None:
                    parent.remove(p_el)

    # ---- 存盘

    def save(self, doc, path: Path):
        if not self.notes:
            doc.save(str(path))
            return 0
        embed_ids = self._embed_ids()
        xml = self._footnote_xml()
        self.detach_from_body()

        # 中间文件名必须唯一：同一本书若有两个生成进程并行（比如 fetch 自带的生成
        # 阶段撞上手动 build），共用一个临时名会互相写坏，产出坏 zip
        tmp = path.with_name(f"{path.stem}.{os.getpid()}.raw.docx")
        doc.save(str(tmp))
        try:
            _inject(tmp, path, xml, embed_ids)
        finally:
            tmp.unlink(missing_ok=True)
        return len(self.notes)


def _inject(src: Path, dst: Path, footnotes_xml: str, embed_ids: set[str] | None = None):
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        blobs = {n: zin.read(n) for n in names}

    # 1) content types
    ct = blobs["[Content_Types].xml"].decode("utf-8")
    if "footnotes+xml" not in ct:
        ct = ct.replace("</Types>",
                        f'<Override PartName="/word/footnotes.xml" ContentType="{FOOTNOTES_CT}"/></Types>')
        blobs["[Content_Types].xml"] = ct.encode("utf-8")

    # 2) document 关系
    rels_name = "word/_rels/document.xml.rels"
    rels = blobs[rels_name].decode("utf-8")
    if "footnotes.xml" not in rels:
        used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
        rid = f"rId{max(used) + 1 if used else 1}"
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{rid}" Type="{FOOTNOTES_REL}" '
                            f'Target="footnotes.xml"/></Relationships>')
        blobs[rels_name] = rels.encode("utf-8")

    # 3) settings 里声明分隔符脚注
    st_name = "word/settings.xml"
    if st_name in blobs:
        st = blobs[st_name].decode("utf-8")
        if "footnotePr" not in st:
            block = '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
            m = re.search(r"(<w:settings[^>]*>)", st)
            if m:
                st = st[:m.end()] + block + st[m.end():]
                blobs[st_name] = st.encode("utf-8")

    # 4) 脚注自己的关系分部：图片本体共用，只把用到的关系项复制过来
    if embed_ids:
        doc_rels = blobs["word/_rels/document.xml.rels"].decode("utf-8")
        picked = [m.group(0) for m in re.finditer(r"<Relationship\b[^>]*/>", doc_rels)
                  if (re.search(r'Id="([^"]+)"', m.group(0)) or [None])
                  and re.search(r'Id="([^"]+)"', m.group(0)).group(1) in embed_ids]
        if picked:
            blobs["word/_rels/footnotes.xml.rels"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + "".join(picked) + "</Relationships>").encode("utf-8")

    blobs["word/footnotes.xml"] = footnotes_xml.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in blobs.items():
            zout.writestr(name, data)
