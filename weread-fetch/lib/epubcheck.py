#!/usr/bin/env python3
"""EPUB 结构自查。

没装 epubcheck（要 Java），所以自己查那几条真会让阅读器打不开或掉资源的规则：
mimetype 的位置与压缩方式、container/opf 的指向、清单与 spine 是否自洽、
id 有无重复、每个 xhtml 是不是合法 XML（EPUB 要求良构 XML，HTML 那套宽容不算数），
以及 xhtml 与 css 里的引用是否都能在包内落地。
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree

OPF_NS = "{http://www.idpf.org/2007/opf}"
CONTAINER_NS = "{urn:oasis:names:tc:opendocument:xmlns:container}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"

# 字体名里可能带括号（FZKaiS(SIP)B.ttf），带引号的要整体认
URL_RE = re.compile(r"""url\(\s*(?:"([^"]*)"|'([^']*)'|([^)\s]*))\s*\)""")
EXTERNAL = ("http:", "https:", "data:", "mailto:")


def _first_group(m) -> str:
    return next((g for g in m.groups() if g), "")


def validate(path: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        names = z.namelist()

        # 1. mimetype 必须是第一条、且不压缩
        if not infos or infos[0].filename != "mimetype":
            problems.append("mimetype 不是压缩包里的第一个条目")
        elif infos[0].compress_type != zipfile.ZIP_STORED:
            problems.append("mimetype 被压缩了，必须以 STORED 方式存")
        elif z.read("mimetype") != b"application/epub+zip":
            problems.append("mimetype 内容不对")

        # 2. container.xml -> opf
        if "META-INF/container.xml" not in names:
            return problems + ["缺 META-INF/container.xml"]
        try:
            root = etree.fromstring(z.read("META-INF/container.xml"))
            opf_path = root.find(f".//{CONTAINER_NS}rootfile").get("full-path")
        except Exception as exc:
            return problems + [f"container.xml 解析失败：{exc}"]
        if opf_path not in names:
            return problems + [f"container.xml 指向的 {opf_path} 不存在"]

        # 3. opf 自洽
        try:
            opf = etree.fromstring(z.read(opf_path))
        except Exception as exc:
            return problems + [f"content.opf 解析失败：{exc}"]
        base = posixpath.dirname(opf_path)

        uid_ref = opf.get("unique-identifier")
        if uid_ref not in {e.get("id") for e in opf.iter(f"{DC_NS}identifier")}:
            problems.append(f"unique-identifier «{uid_ref}» 没有对应的 dc:identifier")

        manifest = {}
        for item in opf.iter(f"{OPF_NS}item"):
            full = posixpath.normpath(posixpath.join(base, item.get("href")))
            manifest[item.get("id")] = full
            if full not in names:
                problems.append(f"清单里的 {item.get('href')} 在包里不存在")
        if not any("nav" in (e.get("properties") or "") for e in opf.iter(f"{OPF_NS}item")):
            problems.append('清单里没有 properties="nav" 的导航文件')

        spine = [e.get("idref") for e in opf.iter(f"{OPF_NS}itemref")]
        if not spine:
            problems.append("spine 是空的")
        for idref in spine:
            if idref not in manifest:
                problems.append(f"spine 引用了清单里没有的 {idref}")

        # 4. 每个 xhtml 必须良构、id 不重复、内部引用能落地
        parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True)
        for name in [p for p in manifest.values() if p.endswith((".xhtml", ".html"))]:
            try:
                doc = etree.fromstring(z.read(name), parser)
            except Exception as exc:
                problems.append(f"{name} 不是良构 XML：{str(exc)[:110]}")
                continue
            here = posixpath.dirname(name)
            all_ids = [e.get("id") for e in doc.iter() if e.get("id")]
            have_ids = set(all_ids)
            if len(all_ids) != len(have_ids):
                dup = sorted({i for i in all_ids if all_ids.count(i) > 1})
                problems.append(f"{name} 里 id 重复：{dup[:5]}")
            for el in doc.iter():
                for attr in ("href", "src"):
                    ref = el.get(attr)
                    if not ref or ref.startswith(EXTERNAL):
                        continue
                    target, _, frag = ref.partition("#")
                    if not target:
                        if frag and frag not in have_ids:
                            problems.append(f"{name} 里的锚点 #{frag} 找不到")
                        continue
                    full = posixpath.normpath(posixpath.join(here, target))
                    if full not in names:
                        problems.append(f"{name} 引用了不存在的 {target}")

        # 5. CSS 里的 url() 同样要落地，否则会留下指向包外的坏引用
        for name in [p for p in manifest.values() if p.endswith(".css")]:
            here = posixpath.dirname(name)
            text = z.read(name).decode("utf-8", "replace")
            for m in URL_RE.finditer(text):
                ref = _first_group(m)
                if not ref or ref.startswith(EXTERNAL):
                    continue
                full = posixpath.normpath(posixpath.join(here, ref))
                if full not in names:
                    problems.append(f"{name} 里的 url({ref}) 在包里不存在")
    return problems


def main() -> int:
    import sys
    if len(sys.argv) < 2:
        print("用法：epubcheck.py <file.epub> …")
        return 2
    bad = 0
    for arg in sys.argv[1:]:
        probs = validate(Path(arg))
        print(f"=== {Path(arg).name}")
        if not probs:
            print("  结构自查通过")
            continue
        bad = 1
        grouped: dict[str, list] = {}
        for p in probs:
            key = re.sub(r"[^ ]*\d[^ ]*", "…", p)
            grouped.setdefault(key, [0, p])
            grouped[key][0] += 1
        for cnt, sample in grouped.values():
            print(f"  ×{cnt:<4} {sample[:150]}")
    return bad


if __name__ == "__main__":
    raise SystemExit(main())
