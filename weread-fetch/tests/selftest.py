#!/usr/bin/env python3
"""转换器自检：拿一份人造章节跑通表格合并、随文图、三种注释与样式映射。

微信读书大多数书里的表格在制作阶段就被拍成了图片，真 <table> 很少见，
所以表格这条路只能靠夹具守着，别等遇到再发现坏了。
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL))

from docx import Document                       # noqa: E402
from docx.oxml.ns import qn                     # noqa: E402
from lib import docxbuild, epubbuild, epubcheck  # noqa: E402

FIXTURE = SKILL / "tests" / "fixture"


def check(cond, label, detail=""):
    print(("通过     " if cond else "失败     ") + label + (f"　{detail}" if detail else ""))
    return bool(cond)


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="weread-selftest-"))
    work = tmp / "work"
    shutil.copytree(FIXTURE, work)
    out = tmp / "自检.docx"

    stats = docxbuild.build_book(work, out, docxbuild.Options(toc=False), "")
    print(f"统计     段 {stats['paragraphs']} / 图 {stats['images']} / 表 {stats['tables']} "
          f"/ 脚注 {stats['footnotes']}")

    ok = True
    import unittest
    from test_quality import QualityTests, StructureTests, TitleTests
    loader = unittest.defaultTestLoader
    suite = unittest.TestSuite([loader.loadTestsFromTestCase(c)
                                for c in (QualityTests, StructureTests, TitleTests)])
    result = unittest.TextTestRunner().run(suite)
    ok &= result.wasSuccessful()
    d = Document(str(out))
    z = zipfile.ZipFile(out)
    doc_xml = z.read("word/document.xml").decode()

    ok &= check(stats["footnotes"] == 6, "六种注释写法都转成脚注", f"实得 {stats['footnotes']}")
    ok &= check("word/footnotes.xml" in z.namelist(), "footnotes.xml 已写入")
    fx = z.read("word/footnotes.xml").decode()
    ok &= check(doc_xml.count("footnoteReference") == 6, "正文里有 6 处脚注引用")
    ok &= check("互指" in fx, "互指式注文进了脚注")
    ok &= check("空 a 上" in fx, "空锚点式注文进了脚注")
    ok &= check(fx.count("<w:drawing>") >= 1, "注文里的图保住了")
    ok &= check("word/_rels/footnotes.xml.rels" in z.namelist(), "脚注有自己的关系分部")
    if "word/_rels/footnotes.xml.rels" in z.namelist():
        import re as _re
        need = set(_re.findall(r'r:embed="([^"]+)"', fx))
        have = set(_re.findall(r'Id="([^"]+)"',
                               z.read("word/_rels/footnotes.xml.rels").decode()))
        ok &= check(need <= have, "脚注图片的关系 id 都能解析", f"{sorted(need)} vs {sorted(have)}")
    ok &= check("藏在 alt 里" in fx, "img 夹注的 alt 文字进了脚注")
    ok &= check("按 id 配对" in fx and "反推配对" in fx, "两种锚点写法都配上了")
    ok &= check("己式注文" in fx, "EPUB3 标准写法（多类名、epub:type）配上了")
    ok &= check("↺" not in fx and "1．己式" not in fx, "回跳箭头与注文自带编号已去掉")
    ok &= check(not stats["unmatched_noterefs"], "没有漏配的注号", str(stats["unmatched_noterefs"]))

    ok &= check(len(d.tables) == 1, "表格已生成", f"实得 {len(d.tables)}")
    if d.tables:
        t = d.tables[0]
        ok &= check(len(t.rows) == 4 and len(t.columns) == 4, "表格 4×4")
        # colspan：首行「制度」跨两列
        grid = t.rows[0].cells
        ok &= check(grid[1].text.strip() == grid[2].text.strip() == "制度", "colspan 合并生效")
        # rowspan：「汉」跨两行
        ok &= check(t.cell(1, 0).text.strip() == t.cell(2, 0).text.strip() == "汉", "rowspan 合并生效")
        ok &= check(t._tbl.tblPr.find(qn("w:tblBorders")) is not None, "表格有边框")

    styles = [p.style.name for p in d.paragraphs]
    ok &= check("WR 图说" in styles, "图说用了图说样式")
    ok &= check("WR 落款" in styles, "落款右对齐样式命中")
    ok &= check("WR 分隔" in styles, "hr 变成分隔行")

    quote = next((p for p in d.paragraphs if p.text.startswith("此段为引文")), None)
    ok &= check(quote is not None, "找到引文段")
    if quote:
        ok &= check(quote.paragraph_format.left_indent is not None
                    and quote.paragraph_format.left_indent.pt > 10, "引文左缩进已应用")
        fonts = set()
        for r in quote.runs:
            rpr = r._element.find(qn("w:rPr"))
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None:
                    fonts.add(rf.get(qn("w:eastAsia")))
        ok &= check("楷体" in fonts, "引文用了楷体", str(sorted(f for f in fonts if f)))
        ok &= check(any(r.font.bold for r in quote.runs), "行内加粗保留")
        ok &= check(any(r.font.superscript for r in quote.runs), "上标保留")
        ok &= check(any(r.font.subscript for r in quote.runs), "下标保留")
        sizes = {round(r.font.size.pt, 1) for r in quote.runs if r.font.size}
        ok &= check(len(sizes) > 1, "小字字号与正文不同", str(sorted(sizes)))

    imgs = [p for p in d.paragraphs if p._p.findall(".//" + qn("w:drawing"))]
    ok &= check(len(imgs) >= 2, "块图与随文图都嵌入了", f"实得 {len(imgs)}")
    inline = next((p for p in d.paragraphs if p.text.startswith("随文生僻字")), None)
    ok &= check(inline is not None and inline._p.findall(".//" + qn("w:drawing")),
                "生僻字图留在正文行内")

    caption = next((p for p in d.paragraphs if p.text.startswith("图一")), None)
    ok &= check(caption is not None and len(caption.runs) >= 3, "图说里的换行与署名小字保留")

    pua_para = next((x for x in d.paragraphs if "鹈泉" in x.text), None)
    ok &= check(pua_para is not None, "私用区字所在段落存在")
    if pua_para:
        from docx.enum.text import WD_COLOR_INDEX
        text = pua_para.text
        restored = "\ue862" not in text
        highlighted = any(r.font.highlight_color == WD_COLOR_INDEX.YELLOW
                          for r in pua_para.runs)
        ok &= check(restored or highlighted,
                    "私用区字要么还原成真字要么高亮标出，不静默丢失",
                    f"还原={restored} 高亮={highlighted} 文本={text[-12:]!r}")
        if restored:
            ok &= check("䴙鹈泉" in text, "U+E862 还原为䴙", text[-10:])

    poem = next((x for x in d.paragraphs if x.text.startswith("長恨此身非我有")), None)
    ok &= check(poem is not None, "诗词段落存在")
    if poem:
        lines = poem.text.split("\n")
        ok &= check(len(lines) == 3, "换行保留成三行", str(len(lines)))
        ok &= check(all(not ln.startswith((" ", "\u3000")) for ln in lines[1:]),
                    "换行后没有多余行首空格", repr(lines[1:]))

    ext = next((x for x in d.paragraphs if x.text.startswith("扩展区生僻字")), None)
    ok &= check(ext is not None, "扩展区生僻字段落存在")
    if ext:
        fonts = set()
        for r in ext.runs:
            rpr = r._element.find(qn("w:rPr"))
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None and rf.get(qn("w:eastAsia")):
                    fonts.add(rf.get(qn("w:eastAsia")))
        rare = [r.text for r in ext.runs if "𠄀" in r.text]
        ok &= check(len(fonts) > 1, "扩展区字被切出来单独配了字体", str(sorted(fonts)))
        ok &= check(bool(rare) and all(len(t) <= 6 for t in rare),
                    "只把缺的字切出去，没有整段换字体", str(rare))

    heads = [x for x in d.paragraphs if x.style.name.startswith("WR 标题")]
    ok &= check(all(x.text.strip() for x in heads), "没有无文字的标题（图片不当大纲）")
    ok &= check(not any(x._p.findall(".//" + qn("w:drawing")) for x in heads),
                "标题段里不夹图")

    # ---- EPUB 侧
    epub_out = tmp / "自检.epub"
    est = epubbuild.build_epub(work, epub_out, "")
    print(f"EPUB     {est.chapters} 章 / 图 {est.images} / 脚注 {est.footnotes} "
          f"/ 补锚点 {est.anchors_repaired}")
    probs = epubcheck.validate(epub_out)
    ok &= check(not probs, "EPUB 结构自查通过", "；".join(probs[:2]))
    ok &= check(est.footnotes >= 3, "EPUB 里的注都成了 epub:type 注", f"实得 {est.footnotes}")
    ez = zipfile.ZipFile(epub_out)
    ch = ez.read("OEBPS/Text/ch0001.xhtml").decode()
    ok &= check('epub:type="noteref"' in ch, "注号带 noteref")
    ok &= check('epub:type="footnote"' in ch, "注文带 footnote")
    ok &= check("<table" in ch, "表格在 EPUB 里原样保留")
    ok &= check(ez.infolist()[0].filename == "mimetype", "mimetype 是第一个条目")
    ok &= check("藏在 alt 里" in ch, "img 夹注的注文在 EPUB 正文里可见")

    for name in ("word/document.xml", "word/footnotes.xml", "word/styles.xml"):
        from lxml import etree
        try:
            etree.fromstring(z.read(name))
        except Exception as exc:
            ok = False
            print(f"失败     {name} XML 不合法 {exc}")

    print(("\n自检     全部通过　" if ok else "\n自检     有失败项　") + str(out))
    if ok:
        shutil.rmtree(tmp, ignore_errors=True)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
