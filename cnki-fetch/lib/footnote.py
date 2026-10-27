# -*- coding: utf-8 -*-
"""给 python-docx 补上真·页下脚注：手工写 word/footnotes.xml 部件并挂关系。

python-docx 本身没有脚注 API，只能在 doc.save() 之后重打包 docx。
"""
import os, shutil, zipfile
from xml.sax.saxutils import escape
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

FONT = "STSong"


def _rpr(sz_half, sup=False):
    return (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}"/>'
            f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/>'
            + ('<w:vertAlign w:val="superscript"/>' if sup else '') + '</w:rPr>')


def add_ref(paragraph, fid, sz_half=21, mark=None):
    """在段落当前末尾插入脚注引用。mark 非空则用自定义标记（如 *），不占自动编号。"""
    body = (f'<w:footnoteReference w:customMarkFollows="1" w:id="{fid}"/><w:t>{escape(mark)}</w:t>'
            if mark else f'<w:footnoteReference w:id="{fid}"/>')
    paragraph._p.append(parse_xml(f'<w:r {nsdecls("w")}>{_rpr(sz_half, sup=True)}{body}</w:r>'))


def set_continuous(doc):
    """脚注全篇连续编号，不按页、按节重排。"""
    sectPr = doc.sections[0]._sectPr
    sectPr.insert(0, parse_xml(f'<w:footnotePr {nsdecls("w")}><w:numFmt w:val="decimal"/>'
                               f'<w:numRestart w:val="continuous"/></w:footnotePr>'))


def _one(fid, text, sz_half, mark=None):
    ref = f'<w:t>{escape(mark)}</w:t>' if mark else '<w:footnoteRef/>'
    return (f'<w:footnote w:id="{fid}"><w:p><w:pPr>'
            f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:jc w:val="both"/>'
            f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}"/>'
            f'<w:sz w:val="{sz_half}"/></w:rPr></w:pPr>'
            f'<w:r>{_rpr(sz_half, sup=True)}{ref}</w:r>'
            f'<w:r>{_rpr(sz_half)}<w:t xml:space="preserve"> {escape(text)}</w:t></w:r>'
            f'</w:p></w:footnote>')


def inject(path, notes, sz_half=18):
    """notes: [(fid, text, mark_or_None)]，写入 footnotes.xml 并补 content-type 与关系。"""
    seps = ''.join(
        f'<w:footnote w:type="{t}" w:id="{i}"><w:p><w:pPr>'
        f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        f'<w:r><w:{e}/></w:r></w:p></w:footnote>'
        for t, i, e in (("separator", -1, "separator"),
                        ("continuationSeparator", 0, "continuationSeparator")))
    fx = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          f'<w:footnotes {nsdecls("w")}>{seps}'
          + ''.join(_one(f, t, sz_half, m) for f, t, m in notes) + '</w:footnotes>')

    tmp = str(path) + ".tmp"
    shutil.move(str(path), tmp)
    zin, zout = zipfile.ZipFile(tmp), zipfile.ZipFile(str(path), 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        raw = zin.read(item.filename)
        if item.filename == '[Content_Types].xml':
            raw = raw.decode().replace('</Types>',
                '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxml'
                'formats-officedocument.wordprocessingml.footnotes+xml"/></Types>').encode()
        elif item.filename == 'word/_rels/document.xml.rels':
            raw = raw.decode().replace('</Relationships>',
                '<Relationship Id="rIdFn1" Type="http://schemas.openxmlformats.org/officeDocument'
                '/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>').encode()
        zout.writestr(item, raw)
    zout.writestr('word/footnotes.xml', fx.encode())
    zout.close(); zin.close(); os.remove(tmp)
