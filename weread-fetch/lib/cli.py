#!/usr/bin/env python3
"""微信读书抓取入口。用法见 SKILL.md。"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lib import api, paths  # noqa: E402
from lib.ids import reader_url  # noqa: E402
# 题名/作者清理规则在 lib/titles.py：成品内页（docxbuild）也要用同一套。
from lib.titles import clean_author, clean_title, title_needs_review  # noqa: E402,F401

SKILL_DIR = Path(__file__).resolve().parents[1]
CLOAK_PY = paths.BROWSER_PY          # 路径都可用环境变量或 .env 覆盖，见 lib/paths.py
TMP_ROOT = paths.WORK_ROOT
DEFAULT_OUT = paths.OUTPUT_DIR

BAD_FS = re.compile(r'[/\\:*?"<>|\x00-\x1f]')


def slug(text: str, limit: int = 40) -> str:
    return BAD_FS.sub("_", (text or "").strip())[:limit] or "book"


def _wide(text: str) -> int:
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)


def say(label: str, msg: str = ""):
    print(label + " " * max(1, 11 - _wide(label)) + msg)


# ---------------------------------------------------------------- 找书

def resolve_book(client: api.WereadAPI, query: str) -> dict:
    """书名/作者/bookId -> 书目信息。优先书架，找不到再搜书城。"""
    query = query.strip()
    if query.startswith("MP_WXS_"):
        raise api.WereadError(
            "公众号（MP_WXS_ 开头）不走图书接口，本 skill 抓不了；"
            "公众号推文请换用别的工具")
    if re.fullmatch(r"\d{5,}", query) or query.startswith("CB_"):
        info = client.book_info(query)
        if info.get("title"):
            return info
        raise api.WereadError(f"查不到 bookId {query}")

    shelf = client.shelf()
    hits = api.find_in_shelf(shelf, query)
    if len(hits) == 1:
        return client.book_info(hits[0]["bookId"])
    if len(hits) > 1:
        say("多本匹配", "请用 bookId 指定：")
        for b in hits[:15]:
            print(f"           {b['bookId']:<28} {b.get('title','')}　—　{b.get('author','')}")
        raise SystemExit(2)

    found = client.search(query, count=10)
    if not found:
        raise api.WereadError(f"书架和书城都没找到「{query}」")
    say("书架未命中", "以下是书城结果，确认后用 bookId 重跑：")
    for b in found[:10]:
        print(f"           {b.get('bookId',''):<28} {b.get('title','')}　—　{b.get('author','')}")
    raise SystemExit(2)


def out_stem(meta: dict) -> str:
    """成品文件名：书名_作者（都用清理过的）。"""
    title = clean_title(meta.get("title") or "")
    author = clean_author(meta.get("author") or "")
    return slug(f"{title}_{author}" if author else title, 90)


def work_dir_for(info: dict) -> Path:
    return TMP_ROOT / f"{info['bookId']}-{slug(info.get('title',''), 30)}"


def write_book_meta(client: api.WereadAPI, info: dict, work: Path, chapter_filter=None) -> dict:
    chapters = client.chapter_infos(info["bookId"])
    # book.json 永远存全量章节表；筛选只写进 select，决定这次抓哪些。
    # 否则分几次按 --chapters 抓完，最后 build 只会认得最后一次那几章。
    select = None
    if chapter_filter:
        select = [c["chapterUid"] for c in chapters
                  if c.get("chapterIdx", c["chapterUid"]) in chapter_filter]
    meta = {
        "bookId": info["bookId"],
        "title": info.get("title", ""),
        "author": info.get("author", ""),
        "cover": info.get("cover", ""),
        "format": info.get("format", ""),
        "readerUrl": reader_url(info["bookId"]),
        "chapters": chapters,
        "select": select,
    }
    work.mkdir(parents=True, exist_ok=True)
    (work / "book.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    return meta


def parse_range(spec: str) -> set[int] | None:
    if not spec or spec.lower() == "all":
        return None
    out: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, _, b = part.partition("-")
            out.update(range(int(a), int(b) + 1))
        elif part:
            out.add(int(part))
    return out or None


# ---------------------------------------------------------------- 命令

def cmd_login(args):
    """会话以浏览器 profile 为准，profile 无效时自动去本机浏览器（Zen/Chrome 等）取种。"""
    ok = api.refresh_session(headed=args.qr, wait=args.wait if args.qr else 0)
    if not ok:
        say("未登录", "浏览器里没有可用的微信读书会话")
        say("处理", "跑 weread login --qr，用微信扫码登录后会自动存下会话")
        return 4
    client = api.WereadAPI()
    shelf = client.shelf()
    from urllib.parse import unquote
    name = next((c["value"] for c in client.cookies if c["name"] == "wr_name"), "")
    say("登录态", f"有效，用户 {unquote(name) or '(未知)'}")
    say("书架", f"{shelf.get('bookCount', 0)} 本")
    return 0


def cmd_shelf(args):
    client = api.WereadAPI()
    shelf = client.shelf()
    books = shelf.get("books", [])
    if args.keyword:
        books = api.find_in_shelf(shelf, args.keyword)
    say("书架", f"{len(books)} 本" + (f"（筛选：{args.keyword}）" if args.keyword else ""))
    for b in books:
        print(f"           {b.get('bookId',''):<28} {b.get('title','')[:34]:<36} {b.get('author','')[:18]}")


def cmd_search(args):
    client = api.WereadAPI()
    for b in client.search(args.keyword, count=args.count):
        print(f"           {b.get('bookId',''):<28} {b.get('title','')[:34]:<36} {b.get('author','')[:18]}")


def cmd_info(args):
    client = api.WereadAPI()
    info = resolve_book(client, args.book)
    chapters = client.chapter_infos(info["bookId"])
    say("书名", f"{info.get('title','')}　{info.get('author','')}")
    say("bookId", info["bookId"])
    say("阅读页", reader_url(info["bookId"]))
    say("章节", f"{len(chapters)} 章，合计 {sum(int(c.get('wordCount') or 0) for c in chapters)} 字")
    for c in chapters:
        idx = c.get("chapterIdx", c["chapterUid"])
        pad = "  " * (int(c.get("level") or 1) - 1)
        print(f"           {idx:>4}  {pad}{c.get('title','')[:40]:<44} {c.get('wordCount', 0):>7} 字")


def cmd_capture(args, info=None, work=None):
    client = api.WereadAPI()
    if info is None:
        info = resolve_book(client, args.book)
        work = work_dir_for(info)
    meta = write_book_meta(client, info, work, parse_range(getattr(args, "chapters", None)))
    say("书名", f"{meta['title']}　{meta['author']}")
    say("任务目录", str(work))
    say("待抓", f"{len(meta.get('select') or meta['chapters'])} 章"
                + (f"（全书 {len(meta['chapters'])} 章）" if meta.get("select") else ""))

    # -u：否则抓取进度会被块缓冲，管道里几分钟看不到一行
    cmd = [str(CLOAK_PY), "-u", str(SKILL_DIR / "lib" / "capture.py"), "--work-dir", str(work)]
    if getattr(args, "headed", False):
        cmd.append("--headed")
    if getattr(args, "limit", 0):
        cmd += ["--limit", str(args.limit)]
    if getattr(args, "delay", None) is not None:
        cmd += ["--delay", str(args.delay)]
    rc = subprocess.call(cmd)
    return info, work, rc


def cmd_epub(args, info=None, work=None):
    from lib import epubbuild
    if info is None:
        info = resolve_book(api.WereadAPI(), args.book)
        work = work_dir_for(info)
    meta = json.loads((work / "book.json").read_text(encoding="utf-8"))
    out = (Path(args.out).with_suffix(".epub") if getattr(args, "out", None)
           else DEFAULT_OUT / f"{out_stem(meta)}.epub")
    cookies = api.cookie_header(api.load_cookies(auto_sync=False)) if api.COOKIE_FILE.exists() else ""
    st = epubbuild.build_epub(work, out, cookies,
                              variant=getattr(args, "variant", "traditional"))
    from lib import epubcheck
    problems = epubcheck.validate(out)
    say("EPUB", str(out))
    if problems:
        say("结构问题", f"{len(problems)} 处")
        for pr in problems[:6]:
            print(f"           {pr[:130]}")
    else:
        say("结构自查", "通过（mimetype/清单/spine/锚点/XML 良构）")
    say("统计", f"{st.chapters} 章 / 图 {st.images} / 样式表 {st.css_files} / 脚注 {st.footnotes}")
    if st.fonts_embedded or st.fonts_missing:
        say("嵌入字体", f"取到 {st.fonts_embedded} 个"
                        + (f"，{len(st.fonts_missing)} 个服务端没有（对应 @font-face 已删除，回退系统字体）"
                           if st.fonts_missing else ""))
        for n in st.fonts_missing[:4]:
            print(f"           缺 {n}")
    if st.anchors_repaired:
        say("锚点修复", f"{st.anchors_repaired} 处原书注号指空，已按配对关系补回")
    if st.image_failures:
        say("图片缺失", f"{len(st.image_failures)} 个地址取不到（已从正文摘除）")
    if st.pua:
        uniq = sorted({c for c, _, _ in st.pua})
        say("生僻字", f"{len(st.pua)} 处私用区码位（{', '.join(uniq[:6])}）——原书靠嵌入字体显示，"
                      "而微信读书不提供字体，任何阅读器里都会是空白")
        for c, t, ctx in st.pua[:5]:
            print(f"           {c}  {t[:16]}　…{ctx}…")
    return out


def cmd_build(args, info=None, work=None):
    from lib import docxbuild
    client = None
    if info is None:
        client = api.WereadAPI()
        info = resolve_book(client, args.book)
        work = work_dir_for(info)
    if not (work / "book.json").exists():
        raise SystemExit(f"任务目录里没有 book.json：{work}")

    meta = json.loads((work / "book.json").read_text(encoding="utf-8"))
    out = (Path(args.out) if getattr(args, "out", None)
           else DEFAULT_OUT / f"{out_stem(meta)}.docx")
    raw_title = (meta.get("title") or "").strip()
    raw_author = (meta.get("author") or "").strip()
    title, author = clean_title(raw_title), clean_author(raw_author)
    if title != raw_title or author != raw_author:
        say("书目清理", f"{raw_title}／{raw_author}　→　{title}／{author}"
                        "（去掉营销标签、装帧、册数与署名角色词；版次与丛书名不动）")
    left = title_needs_review(title)
    if left:
        say("题名待核", f"仍留着 {left}——若非题名成分，用 --out 指定文件名。"
                        "（丛书编号一类是书目信息，留着是对的）")

    opts = docxbuild.Options(
        title=title, author=author,
        toc=getattr(args, "toc", False) and not getattr(args, "no_toc", False),
        page_break_chapters=not getattr(args, "no_page_break", False),
        body_size_pt=getattr(args, "size", 11.0),
        body_font=getattr(args, "font", "") or "中华书局宋体",
        variant=getattr(args, "variant", "traditional"),
    )
    cookies = api.cookie_header(api.load_cookies(auto_sync=False)) if api.COOKIE_FILE.exists() else ""
    stats = docxbuild.build_book(work, out, opts, cookies)

    say("成品", str(out))
    cover = stats.get("cover", "无")
    say("扉页", "书名、作者、封面" if cover == "嵌入"
                else "书名、作者（书目没有封面图）" if cover == "无"
                else f"书名、作者；缺封面——{cover}")
    if stats.get("toc_entries"):
        say("书前目录", f"{stats['toc_entries']} 条（照正文实际排出的标题）")
    warnings = stats["audit"]["warnings"]
    say("结构验收", "通过" + (f"，另有 {len(warnings)} 条待核" if warnings else "")
                    + "（高亮范围、脚注引用、HTML 标签泄漏、图片嵌入；不等于视觉或逐字校验）")
    say("验收报告", stats["audit_path"])
    for warning in stats["audit"]["warnings"]:
        say("待核", warning)
    say("统计", f"{stats['chapters']} 章 / {stats['paragraphs']} 段 / {stats['chars']} 字 / "
                f"图 {stats['images']} / 表 {stats['tables']} / 脚注 {stats['footnotes']}")
    if stats.get("variant_dropped"):
        kind = "繁体" if str(getattr(args, "variant", "traditional")).startswith("t") else "简体"
        say("繁简", f"{stats['variant_dropped']} 章同时含繁简两套，已取{kind}本")
    if stats.get("simplified_left"):
        say("繁简存疑", f"繁体正文里有 {len(stats['simplified_left'])} 种字被 opencc 判为简体。"
                        "原文未作改动；注意 opencc 会把繁体通用的异体字（迹/蹟、周/週、"
                        "搜/蒐、别/別）也算进来，需自行甄别：")
        for line in stats["simplified_left"][:6]:
            print(f"           {line}")
    if stats.get("font_note"):
        say("字体", stats["font_note"])
    if stats.get("font_borrowed"):
        say("借字", "；".join(f"{k} {v} 个字" for k, v in stats["font_borrowed"].items())
                    + "（正文字体缺这些字，已逐字改用备用字体）")
    if stats.get("notices_dropped"):
        say("已滤除", f"{stats['notices_dropped']} 处平台发行声明")
    if stats.get("cross_file_notes"):
        say("跨章注释", f"{stats['cross_file_notes']} 条注号与注文分在两章，已配成脚注")
    if stats.get("title_fields_cleaned"):
        say("题名清理", f"版权页「书名：」按题名规则清理 {stats['title_fields_cleaned']} 处")
    if stats.get("pua_map"):
        say("生僻字", f"自动还原 {stats['pua_fixed']} 处（有字形证据）：" +
                      "；".join(f"{k}→{v}" for k, v in list(stats["pua_map"].items())[:6]))
    if stats.get("pua_manual"):
        say("人工指定", f"{len(stats['pua_manual'])} 个码位、{stats['pua_fixed_manual']} 处按人工判读替换"
                        "（无字形证据，凭文本判读；依据见 .state/pua_manual.json）：")
        for k, v in list(stats["pua_manual"].items())[:8]:
            print(f"           {k} → {v}")
    if stats.get("pua_left"):
        say("待核字", f"{len(stats['pua_left'])} 个私用区码位查不到真字，已在正文高亮："
                      + "、".join(stats["pua_left"][:8]))
    if stats.get("continuations"):
        say("接排", f"{stats['continuations']} 章是超长章被切开的后半截（标题只是序号），已接排不另立标题")
    if stats.get("image_failures"):
        say("图片缺失", f"{stats['image_failures']} 张取不到，已在原位留记号（块图留文字说明、"
                        "随文图留高亮〔?〕）。这类多半平台侧就没有——阅读器自己也显示不出来")
    for s in stats.get("short_chapters", [])[:8]:
        say("字数偏少", f"{s['title'][:28]} 期望 {s['expect']} 实得 {s['got']}")
    for s in stats.get("platform_dups", []):
        say("平台缺章", f"{s['title'][:28]} 平台给的是「{s['same_as'][:20]}」的原文，"
                        "阅读器上也如此，重抓无用，已略去")
    return out


def cmd_fetch(args):
    client = api.WereadAPI()
    info = resolve_book(client, args.book)
    work = work_dir_for(info)
    if not args.build_only:
        info, work, rc = cmd_capture(args, info, work)
    formats = [f.strip().lower() for f in (args.format or "docx").split(",") if f.strip()]
    out = None
    if "epub" in formats:
        out = cmd_epub(args, info, work)
    if "docx" in formats:
        out = cmd_build(args, info, work)
    return out


def cmd_selftest(args):
    import subprocess
    return subprocess.call([sys.executable, str(SKILL_DIR / "tests" / "selftest.py")])


def cmd_doctor(args):
    """环境自检：依赖、浏览器运行时、字体、登录态与路径。缺必需项时退出码 1。"""
    import importlib
    import shutil
    say("Python", sys.version.split()[0])
    bad = 0
    for mod, need in (("docx", True), ("bs4", True), ("lxml", True), ("PIL", True),
                      ("fontTools", True), ("rookiepy", False), ("opencc", False)):
        try:
            importlib.import_module(mod)
            say(mod, "ok")
        except Exception:
            bad += need
            say(mod, "缺失，重装 requirements.txt" if need else "缺失（可选：从浏览器取 Cookie／繁简找书）")
    probe = subprocess.run([str(CLOAK_PY), "-c", "import cloakbrowser; print(cloakbrowser.__version__)"],
                           capture_output=True, text=True)
    if probe.returncode == 0:
        say("浏览器", f"cloakbrowser {probe.stdout.strip()}（{CLOAK_PY}）")
    else:
        bad += 1
        say("浏览器", f"缺失：{CLOAK_PY} 导入不了 cloakbrowser，正文抓不了")
    from lib import fontcover
    primary, _ = fontcover.resolve_preset("中华书局宋体")
    say("正文字体", "中华书局宋体 已装" if fontcover.coverage(primary) is not None
                    else "中华书局宋体 未装，默认会退回宋体（可用 --font 另指定）")
    say("opencc命令", "有" if shutil.which("opencc") else "无（可选，仅用于书名转繁）")
    say("登录态", "有 .state/cookies.json" if api.COOKIE_FILE.exists() else "无，先跑 weread login")
    say("浏览器档案", str(paths.PROFILE_DIR))
    say("任务目录", str(TMP_ROOT))
    say("成品目录", str(DEFAULT_OUT))
    return 1 if bad else 0


def cmd_check(args):
    from lib.docxcheck import audit
    report = audit(Path(args.path))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report["errors"] else 0


def main() -> int:
    ap = argparse.ArgumentParser(prog="weread", description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("login", help="建立/续上微信读书登录态")
    p.add_argument("--qr", action="store_true", help="开窗口用微信扫码登录")
    p.add_argument("--wait", type=int, default=180, help="扫码等待秒数，默认 180")
    p.set_defaults(func=cmd_login)

    p = sub.add_parser("shelf", help="列书架")
    p.add_argument("keyword", nargs="?", help="按书名/作者过滤")
    p.set_defaults(func=cmd_shelf)

    p = sub.add_parser("search", help="搜书城")
    p.add_argument("keyword")
    p.add_argument("--count", type=int, default=20)
    p.set_defaults(func=cmd_search)

    p = sub.add_parser("info", help="看书的章节表")
    p.add_argument("book")
    p.set_defaults(func=cmd_info)

    def add_common(sp):
        sp.add_argument("book")
        sp.add_argument("--chapters", help="只要部分章，如 3-20 或 1,5,9（按章节表的序号）")
        sp.add_argument("--out", help="输出 docx 路径")
        sp.add_argument("--no-page-break", action="store_true", help="章间不分页")
        sp.add_argument("--size", type=float, default=11.0, help="正文字号 pt，默认 11")
        sp.add_argument("--variant", default="traditional", choices=["traditional", "simplified"],
                        help="繁体底本的书会同时下发繁简两套，取哪一套（默认繁体）")
        sp.add_argument("--font", default="",
                        help="正文中文字体，默认「中华书局宋体」（按平面拆开的成套字体，"
                             "缺字会自动逐字向 02/15 平面借）。也可填 宋体 / 霞鹜文楷 等；"
                             "填楷体系时引文会自动改用仿宋以保持对比")
        sp.add_argument("--headed", action="store_true", help="显示浏览器窗口")
        sp.add_argument("--limit", type=int, default=0, help="只抓前 N 章（调试）")
        sp.add_argument("--delay", type=float, default=0.8,
                        help="章与章之间多等几秒，放慢节奏用（默认 0.8）")
        sp.add_argument("--format", default="docx",
                        help="出什么：docx（默认）/ epub / docx,epub")
        sp.add_argument("--toc", action="store_true",
                        help="在书前排一份纯文字目录（默认不排，靠侧栏大纲导航）")
        sp.add_argument("--no-toc", action="store_true", help=argparse.SUPPRESS)  # 旧参数，已是默认

    p = sub.add_parser("fetch", help="抓正文并生成 docx（可续跑）")
    add_common(p)
    p.add_argument("--build-only", action="store_true", help="跳过抓取，直接用缓存生成")
    p.set_defaults(func=cmd_fetch)

    p = sub.add_parser("capture", help="只抓正文，不生成 docx")
    add_common(p)
    p.set_defaults(func=cmd_capture)

    p = sub.add_parser("build", help="只用已抓缓存生成 docx")
    add_common(p)
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("epub", help="只用已抓缓存打包 EPUB（归档件，贴近原书）")
    add_common(p)
    p.set_defaults(func=cmd_epub)

    p = sub.add_parser("selftest", help="跑转换器自检（表格/图/注释/样式）")
    p.set_defaults(func=cmd_selftest)

    p = sub.add_parser("doctor", help="环境自检：依赖、浏览器、字体、登录态与路径")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("check", help="离线检查已生成 DOCX，不登录、不重抓")
    p.add_argument("path")
    p.set_defaults(func=cmd_check)

    args = ap.parse_args()
    try:
        result = args.func(args)
        if isinstance(result, int):
            return result
    except api.NotLoggedIn as exc:
        say("未登录", str(exc))
        say("处理", "在本机浏览器里登录 weread.qq.com 后跑 weread login，或直接 weread login --qr 扫码")
        return 3
    except api.WereadError as exc:
        say("失败", str(exc))
        return 1
    except RuntimeError as exc:
        say("失败", str(exc))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
