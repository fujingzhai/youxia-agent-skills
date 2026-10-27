"""离线检查真实 DOCX：缺字高亮范围、脚注闭环、字体分布、标签泄漏与标题结构。

这是结构验收，不替代 Word 的视觉检查，也不证明与底本逐字相同。
"""
from collections import Counter
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from lib.cssmap import FONT_TABLE
from lib.pua import is_pua

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"

OUTLINE_STYLES = {"WR 标题一": 1, "WR 标题二": 2, "WR 标题三": 3, "WR 标题四": 4}
# 不算正文的段落样式：标题、扉页、书前目录
NON_BODY_PREFIX = ("WR 标题", "WR 书名", "WR 作者", "WR 目录")
# 标题里的注号：[1]（1）〔1〕①
RE_NOTE_MARK = re.compile(r"[\[［(（〔]\d{1,3}[\]］)）〕]|[①-⑳]")


def audit(path: Path) -> dict:
    report = {"path": str(path.resolve()), "errors": [], "warnings": [],
              "highlights": [], "fonts": {}, "scope": "OOXML structure, not visual or source equivalence"}
    errors = report["errors"]
    fonts = Counter()
    with ZipFile(path) as z:
        roots = {name: ET.fromstring(z.read(name)) for name in
                 ("word/document.xml", "word/footnotes.xml") if name in z.namelist()}
        doc = roots["word/document.xml"]
        refs = [r.get(W + "id") for r in doc.findall(".//w:footnoteReference", NS)]
        notes = [n.get(W + "id") for n in roots.get("word/footnotes.xml", ET.Element("empty"))
                 if n.tag == W + "footnote" and n.get(W + "type") is None]
        report["footnotes"] = len(notes)
        if Counter(refs) != Counter(notes):
            errors.append("正文脚注引用与注文 ID 不一一对应")
        for part, root in roots.items():
            for p in root.findall(".//w:p", NS):
                text = "".join(n.text or "" for n in p.findall(".//w:t", NS))
                if re.search(r"<(?:p|div|span)\s+(?:class|style)\s*=", text, re.I):
                    errors.append(f"{part} 出现未解析的 HTML 标签：{text[:80]}")
                for run in p.findall(".//w:r", NS):
                    t = "".join(n.text or "" for n in run.findall(".//w:t", NS))
                    hi = run.find("w:rPr/w:highlight", NS)
                    unknown = [c for c in t if is_pua(ord(c))]
                    yellow = hi is not None and hi.get(W + "val") == "yellow"
                    if unknown and not yellow:
                        errors.append(f"{part} 存在未高亮的私用区字符")
                    if hi is not None and hi.get(W + "val") != "none":
                        report["highlights"].append({"part": part, "text": t,
                            "characters": len(t), "pua": [f"U+{ord(c):04X}" for c in unknown]})
                        if t != "〔?〕" and (not unknown or len(unknown) != len(t)):
                            errors.append(f"{part} 高亮超出待核字范围（{len(t)} 字符）")
                        if t == "〔?〕":
                            errors.append(f"{part} 存在缺失的随文图片")
                    rf = run.find("w:rPr/w:rFonts", NS)
                    font = rf.get(W + "eastAsia") if rf is not None else None
                    if t and part == "word/document.xml":
                        fonts[font or "继承段落样式"] += len(t)
        report["fonts"] = dict(fonts)
        report["highlighted_characters"] = sum(h["characters"] for h in report["highlights"])
        unknown_count = sum(len(h["pua"]) for h in report["highlights"])
        if unknown_count:
            report["warnings"].append(f"仍有 {unknown_count} 处私用区字符待人工核定；不能宣称已还原")
        total = sum(fonts.values())
        if total and fonts["仿宋"] / total >= 0.3:
            report["warnings"].append(f"仿宋文字占 {fonts['仿宋'] / total:.1%}；需对照原书 CSS，不能仅凭占比判错或自动改宋体")

        paras = list(_paragraphs(doc, _style_names(z)))
        hei = hei_body_paragraphs(paras)
        if hei:
            where = list(dict.fromkeys(ch for ch, _ in hei))
            report["warnings"].append(
                f"正文长段整段仍是黑体 {len(hei)} 段，在 {len(where)} 个一级标题下"
                f"（{'、'.join(where[:5])}{'…' if len(where) > 5 else ''}）；"
                "可能是正文黑体转宋体漏了，也可能是原书刻意设计，须对照 CSS 核实")
            report["hei_body_samples"] = [f"{ch}：{t}" for ch, t in hei[:10]]
        dups = duplicated_headings(paras)
        if dups:
            report["warnings"].append(
                f"标题后面紧跟着把它拆开重复的标题 {len(dups)} 处（{'；'.join(dups[:3])}）；"
                "多半是一条标题被拆成几行后又各自成了标题，侧栏大纲会重复")
        # 图没进 Word 的两种样子：空的图段落（异常被吞掉，扉页封面就这样丢过 21 本），
        # 和块图留下的「〔图片无法嵌入〕」。随文图的〔?〕在上面按高亮报错。
        empty_figs = sum(1 for style, text, _, has_img in paras
                         if style == "WR 图" and not has_img and not text.strip())
        if empty_figs:
            report["warnings"].append(f"空的图段落 {empty_figs} 处：图没嵌进 Word，多半是格式没认出来")
        unembedded = sum(text.count("〔图片无法嵌入〕") for _, text, _, _ in paras)
        if unembedded:
            report["warnings"].append(f"「〔图片无法嵌入〕」{unembedded} 处：图已下载但 Word 认不出")
        report["errors"] = list(dict.fromkeys(errors))
    return report


def _style_names(z: ZipFile) -> dict:
    if "word/styles.xml" not in z.namelist():
        return {}
    root = ET.fromstring(z.read("word/styles.xml"))
    names = {}
    for s in root.findall("w:style", NS):
        n = s.find("w:name", NS)
        if n is not None:
            names[s.get(W + "styleId")] = n.get(W + "val")
    return names


def _paragraphs(doc, names: dict):
    """依次产出 (样式名, 段落文字, 其中落在黑体里的字数, 是否有图)。"""
    hei_font = FONT_TABLE["hei"][0]
    for p in doc.iter(W + "p"):
        ps = p.find("w:pPr/w:pStyle", NS)
        style = names.get(ps.get(W + "val"), ps.get(W + "val")) if ps is not None else "Normal"
        text, hei = "", 0
        for run in p.iter(W + "r"):
            t = "".join(n.text or "" for n in run.findall(W + "t"))
            text += t
            rf = run.find("w:rPr/w:rFonts", NS)
            if t and rf is not None and rf.get(W + "eastAsia") == hei_font:
                hei += len(t)
        yield style, text, hei, p.find(".//" + W + "drawing") is not None


def hei_body_paragraphs(paras) -> list[tuple[str, str]]:
    """整段落在黑体里的正文长段：(所在一级标题, 段首文字)。

    Kindle 转制书整本正文设成黑体时，构建器会把正文黑体落成宋体；某一段漏网
    （《亲密关系（第8版）》14 个章首导语就因为只看了第一个文件而留在黑体）
    在统计里只是黑体字数略多，肉眼翻 Word 才看得出，所以逐段查。"""
    found, chapter = [], ""
    for style, text, hei, _ in paras:
        if style == "WR 标题一":
            chapter = text.strip()
        if style.startswith(NON_BODY_PREFIX):
            continue
        body = text.strip()
        if len(body) >= 20 and hei >= 0.8 * len(text):
            found.append((chapter or "书前", body[:30]))
    return found


def duplicated_headings(paras) -> list[str]:
    """标题后面紧跟着（中间只隔着图或空段）几个更深的标题，拼起来正是它自己。

    典型是章节表的标题补了一条，原书拆成两行的「第3章」「吸引力」又各成一个标题。
    只认「拼起来相等」，不认「是它的一截」：子章节本就常以父标题开头（「智能复利和
    黑洞效应」下的「智能复利」），「第一章」下的「一」也不是重复。"""
    from lib.docxbuild import norm
    seq = []        # 标题记 (级别, 规整文字, 原文)；有字的正文段记 None
    for style, text, _, _ in paras:
        level = OUTLINE_STYLES.get(style)
        if level is not None:
            seq.append((level, norm(text), text.strip()))
        elif text.strip():
            seq.append(None)
    found = []
    for i, head in enumerate(seq):
        if head is None or not head[1]:
            continue
        joined, parts = "", []
        for nxt in seq[i + 1:i + 4]:
            if nxt is None or nxt[0] <= head[0] or not nxt[1]:
                break
            joined += nxt[1]
            parts.append(nxt[2])
            if joined == head[1]:
                found.append(f"{head[2]} → {' + '.join(parts)}")
                break
            if not head[1].startswith(joined):
                break
    return found


def heading_mismatches(path: Path, expected: list[tuple[str, int]]) -> list[str]:
    """成品标题与章节表是否对得上：去空白标点后同名、同级、同序。

    章节表是侧栏大纲的骨架。正文里的章标题若只认出半条（「第10章」）、
    与章名顺序颠倒或级别错位，脚注和高亮检查都看不出来，得拿章节表逐条对。
    成品标题比章节表多出字（副标题）算对上，少了字就是没认全。
    用最长公共子序列对齐，某条对不上时不会连累后面同名的「本章小结」「请你思考」。
    注号不算标题文字：章节表里的「回首与重访[1]」，成品里 [1] 已经成了脚注引用。"""
    from lib.docxbuild import norm

    def key(text: str) -> str:
        return norm(RE_NOTE_MARK.sub("", text or ""))

    with ZipFile(path) as z:
        doc = ET.fromstring(z.read("word/document.xml"))
        names = _style_names(z)
    heads = [(OUTLINE_STYLES[s], key(t), t.strip())
             for s, t, _, _ in _paragraphs(doc, names) if s in OUTLINE_STYLES]
    want = [(max(1, min(4, int(level or 1))), key(title), title)
            for title, level in expected if key(title)]
    n, m = len(want), len(heads)

    def match(i, j):
        return want[i][0] == heads[j][0] and want[i][1] in heads[j][1]

    table = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        row, nxt = table[i], table[i + 1]
        for j in range(m - 1, -1, -1):
            row[j] = nxt[j + 1] + 1 if match(i, j) else max(nxt[j], row[j + 1])
    problems, i, j = [], 0, 0
    while i < n:
        if j < m and match(i, j) and table[i][j] == table[i + 1][j + 1] + 1:
            i, j = i + 1, j + 1
        elif j < m and table[i][j + 1] >= table[i + 1][j]:
            j += 1
        else:
            near = heads[j][2] if j < m else "后面没有标题了"
            problems.append(f"{want[i][2]}（成品此处是：{near}）")
            i += 1
    return problems
