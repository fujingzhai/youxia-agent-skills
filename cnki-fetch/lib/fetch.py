# -*- coding: utf-8 -*-
"""知网检索、按缺项续抓、原版与 DOCX 双落盘及批次暂停。"""
import argparse, fcntl, io, json, re, sys, time, os
from PIL import Image
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).parent))
import cloakbrowser
import captcha
import search as SR
import docx_builder as DB
import quality as QA
import records as REC
import network as NET

SKILL = Path(__file__).resolve().parent.parent
PROFILE = SKILL / "profiles/cnki"
CACHE = SKILL / "cache"
DEFAULT_OUT = REC.DEFAULT_ROOT      # 环境变量 CNKI_OUTPUT_DIR，或 .env 里设
SEARCH = "https://kns.cnki.net/kns8s/defaultresult/index?kw={}&korder=TI"


_lock_fd = None


def acquire_lock():
    """cloakbrowser 免费版只允许一个会话，两个 cnki 同时跑会抢同一个 profile 直接崩。
    浏览器起来之前先拿独占锁。"""
    global _lock_fd
    _lock_fd = open(SKILL / ".lock", "w")
    try:
        fcntl.flock(_lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        log("忙", "另一个 cnki 正在运行，等它结束再跑（浏览器只能开一个会话）")
        sys.exit(2)
    for stale in PROFILE.glob("Singleton*"):     # 上次异常退出留下的锁
        try:
            stale.unlink()
        except OSError:
            pass


def log(tag, msg):
    w = sum(2 if ord(c) > 127 else 1 for c in tag)   # 中文按两格宽对齐
    print(tag + " " * max(1, 11 - w) + str(msg), flush=True)


def safe(name, limit=120):
    return re.sub(r'[/:*?"<>|\\]', '_', name).strip()[:limit]


def parse_meta(md, row):
    """知网节 '文史 . 2026 (01) : 5-27' → 引用行、刊名、年、期"""
    out = {}
    tip = re.sub(r'\s+', ' ', (md.get('tip') or '')).replace('查看该刊数据库收录来源', '').strip()
    # 形如「文史 . 2026 (01) : 5-27」；也有带卷号的「中国语言文学研究 . 2024 ,36 (01) : 66-73」
    m = re.match(r'^(.+?)\s*\.\s*(\d{4})\s*(?:[,，]\s*(\d+))?\s*(?:\((\d+)\))?'
                 r'\s*(?::\s*([\d\-–—]+))?', tip)
    if m:
        jn, yr, vol, iss, pg = (m.group(1).strip(), m.group(2),
                                m.group(3), m.group(4), m.group(5))
        cite = f"《{jn}》{yr}年"
        if iss:
            cite += f"第{int(iss)}期"
        elif vol:
            cite += f"第{int(vol)}卷"
        if pg:
            cite += "，第" + pg.replace('-', '—').replace('–', '—') + "页"
        out.update(cite=cite, journal=jn, year=yr,
                   issue=(str(int(iss)) if iss else (str(int(vol)) if vol else "")))
    else:
        jn, yr = row.get('source', ''), (row.get('date') or '')[:4]
        out.update(cite=(f"《{jn}》{yr}年" if jn else ""), journal=jn, year=yr, issue="")
    if md.get('doi'):
        out['doi'] = md['doi']
    au = [a for a in (md.get('au') or []) if a]
    # 作者名取搜索结果行的字段（形如「金立;于翔」，分隔明确）。详情页的 .author 会把多个
    # <a> 的文字直接拼成「金立于翔」，中间没有分隔符，据此判断人数会把两个人当成一个。
    names = re.sub(r'\s*[;；,，]\s*', '、', (row.get('author') or '').strip()).strip(' 、')
    if au or names:
        if not names:
            names = re.sub(r'\s*\d+(\s*[,，]\s*\d+)*\s*$', '', au[0]).strip()
            names = re.sub(r'\s*[,，]\s*', '、', re.sub(r'\s+', ' ', names)).strip(' 、')
        orgs = re.sub(r'\s*\d+\.\s*', '　', au[1]).strip() if len(au) > 1 else ''
        single = '、' not in names
        if single and orgs:
            out['author_line'] = f"{names}（{orgs.replace('　', '、').strip('、')}）"
        else:
            out['author_line'] = names
            if orgs:
                out['orgn_line'] = orgs
    return out


def filename(row, meta):
    """篇名_作者_刊名年-期.docx"""
    authors = re.sub(r'\s+', '', row.get('author') or '').replace(',', '、').replace('，', '、')
    parts = [p for p in authors.split('、') if p]
    if len(parts) > 3:
        authors = parts[0] + "等"
    src = meta.get('journal') or row.get('source') or ''
    tail = f"{src}{meta.get('year','')}"
    if meta.get('issue'):
        tail += f"-{meta['issue']}"
    return safe("_".join(x for x in (row['name'], authors, tail) if x)) + ".docx"


def search(ctx, kw):
    page = ctx.new_page()
    try:    # 等结果表真的渲染出来；固定 sleep 慢一拍就会误报「无结果」
        guard = NET.Guard(page)
        page.goto(SEARCH.format(quote(kw)), wait_until="domcontentloaded")
        if not NET.wait_content(page, 'table.result-table-list tbody tr', guard,
                                empty_selector='.no-content, .noresult, .no-result, #noResult, .no-data'):
            return []
        rows = page.evaluate("""()=>Array.from(document.querySelectorAll('table.result-table-list tbody tr')).map(tr=>({
        href: tr.querySelector('.name a')?.href,
        name: tr.querySelector('.name a')?.innerText?.trim(),
        author: tr.querySelector('.author')?.innerText?.trim().replace(/\\s+/g,' '),
        source: tr.querySelector('.source')?.innerText?.trim(),
        date: tr.querySelector('.date')?.innerText?.trim(),
        ai: tr.querySelector('.operat a.icon-airead')?.href}))""")
        return [r for r in rows if r.get('name')]
    finally:
        page.close()


def detail_meta(ctx, row, page=None):
    p = page or ctx.new_page()
    try:
        guard = NET.Guard(p)
        p.goto(row['href'], wait_until="domcontentloaded")
        NET.wait_content(p, '.top-tip', guard)
        md = p.evaluate("""()=>({tip: document.querySelector('.top-tip')?.innerText?.trim(),
        au: Array.from(document.querySelectorAll('.author')).map(e=>e.innerText.trim()),
        doi: Array.from(document.querySelectorAll('li')).filter(li=>li.querySelector('.rowtit')?.innerText?.includes('DOI'))
              .map(li=>li.innerText.replace(/^\\s*DOI[：:]\\s*/,'').trim())[0]})""")
        return parse_meta(md, row)
    finally:
        if page is None:
            p.close()


IMG_SEL = 'div.ChapterWrap.Content img[src*="attachment"]'


def _load_once(ctx, url):
    """加载一次 AI 阅读页，返回 (html, {正文中的图片序号: bytes})。

    字形图 URL 带一次性 nonce，页面渲染时用掉就作废，事后重新 GET 只拿到
    {"code":706}，所以必须在加载时监听响应收下 body。少数图页面自己也不会去请求，
    靠调用方重载补齐——两次加载的正文结构一致，按序号对得上。
    """
    p = ctx.new_page()
    got = {}
    try:
        guard = NET.Guard(p)
        p.on("response", lambda r: _grab(r, got))
        p.goto(url, wait_until="domcontentloaded")
        NET.wait_content(p, 'div.ChapterWrap.Content', guard, timeout=30000)
        p.wait_for_timeout(1200)
        order = []
        order = p.evaluate("""async(sel)=>{
            const els=[...document.querySelectorAll(sel)];
            for(const e of els){ e.scrollIntoView({block:'center'}); await new Promise(r=>setTimeout(r,60)); }
            window.scrollTo(0,0); return els.map(e=>e.src); }""", IMG_SEL)
        for _ in range(10):
            if sum(1 for u in order if u in got) >= len(order):
                break
            p.wait_for_timeout(600)
        html = p.content()
        return html, {i: got[u] for i, u in enumerate(order) if u in got}
    finally:
        p.close()


def _grab(r, got):
    if 'xml/attachment' not in r.url:
        return
    try:
        b = r.body()
        if b[:3] == b'\xff\xd8\xff' or b[:4] == b'\x89PNG':
            got[r.url] = b
    except Exception:
        pass


def full_text(ctx, row):
    html, imgs = _load_once(ctx, row['ai'])
    if not html or "ChapterWrap" not in html:
        return None, {}
    srcs = DB.parse(DB.BeautifulSoup(html, 'html.parser'))[2]
    # 页面并不会去请求每一个字形图（约两成从不发请求）。这些 URL 的一次性 nonce
    # 反而还没被用掉，事后直接 GET 正好能取到——与截获互补，两步凑齐。
    miss = [i for i in range(len(srcs)) if i not in imgs]
    if miss:
        log("补图", f"页面截获 {len(imgs)}/{len(srcs)}，直取剩余 {len(miss)} 个")
        for i in miss:
            for _ in range(3):
                try:
                    response = ctx.request.get(srcs[i])
                    if response.status in (401, 403, 429):
                        raise NET.PauseBatch(f"字形图请求返回 HTTP {response.status}，暂停批次")
                    b = response.body()
                    if b[:3] == b'\xff\xd8\xff' or b[:4] == b'\x89PNG':
                        imgs[i] = b; break
                except NET.PauseBatch:
                    raise
                except Exception:
                    pass
                time.sleep(1.2)
    return html, imgs


def warn_notes(st):
    """知网 XML 的脚注完整性因篇而异，漏标时正文照旧、注号连同注文一起消失。
    正文很长却几乎没有脚注，多半就是漏标，提醒去核对原刊。"""
    if st['notes'] < 5 and st['paras'] > 25 and not st.get('refs'):
        log("！注意", f"全文 {st['paras']} 段只有 {st['notes']} 条脚注。可能本就用文中夹注体例，"
                      f"也可能知网 XML 漏标——对照原版 PDF 核对")


def write_progress(outdir, items, current=None, status="running", reason=None):
    """批量状态写入独立隐藏文件；中断后能准确知道停在哪里，不依赖日志滚屏。"""
    data = {"updated_at": time.strftime('%Y-%m-%dT%H:%M:%S%z'), "status": status,
            "items": items}
    if current is not None:
        data["current"] = current
    if reason:
        data["reason"] = str(reason)
        retry_after = getattr(reason, "retry_after", None)
        if retry_after is not None:
            data["retry_after_seconds"] = retry_after
    try:
        QA.write_json(Path(outdir) / ".cnki-progress.json", data)
    except OSError as e:
        log("！", f"进度文件写入失败：{e}")


def download_file(ctx, row, meta, outdir, fmt="pdf"):
    """下原版 PDF／CAJ。走详情页的 .btn-dlpdf / .btn-dlcaj（不是 a 标签，选择器别加 a），
    这两个入口不经过 bar.cnki.net，无验证码；撞点选码的是「HTML阅读」那个链接。"""
    name = meta.get('_filename', row['name'] + '.docx')
    out = REC.existing_path(outdir / (re.sub(r'\.docx$', '', name) + '.' + fmt))
    out.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(3):        # 这个接口会偶发卡住，重试比干等管用
        d = ctx.new_page()
        try:
            guard = NET.Guard(d)
            d.goto(row['href'], wait_until="domcontentloaded")
            NET.wait_content(d, f'.btn-dl{fmt}', guard)
            el = d.query_selector(f'.btn-dl{fmt}')
            if not el:
                log("！", f"详情页没有 {fmt.upper()} 下载按钮")
                return None
            with d.expect_download(timeout=90000) as di:
                el.click()
            di.value.save_as(str(out))
            # 下载完成不等于拿到原版：登录失效时知网偶尔会把错误页保存成 .pdf。
            # 立即检查文件头和 EOF，避免把损坏文件当作已完成成果，下一次也能续抓。
            if not QA.valid_original(out, fmt):
                try:
                    out.unlink()
                except OSError:
                    pass
                raise ValueError(f"下载内容不是有效的 {fmt.upper()} 文件")
            REC.record(out, {**row, **meta})
            return out
        except NET.PauseBatch:
            raise
        except Exception as e:
            if attempt == 2:
                log("！", f"{fmt.upper()} 下载失败（试了 3 次）：{str(e)[:50]}")
                REC.record(out, {**row, **meta}, error=f'{fmt.upper()} 下载失败：{e}')
                return None
            log("重试", f"{fmt.upper()} 下载第 {attempt + 1} 次未成，5 秒后再来")
            time.sleep(5)
        finally:
            d.close()
    return None



def note_manual(kw, row, reason, ctx=None, meta=None, outdir=None):
    """没有 XML 正文时记一笔，并自动降级去下原版 PDF——总比空手好。"""
    f = SKILL / "待手工下载.md"
    line = f"- [ ] {kw}｜{row.get('name','')}｜{row.get('source','')} {row.get('date','')}｜{reason}\n"
    old = f.read_text(encoding='utf-8') if f.exists() else "# 无 HTML 全文，已自动下 PDF\n\n"
    if line not in old:
        f.write_text(old + line, encoding='utf-8')
    log("回退", f"{reason}，改下原版 PDF")
    if ctx is not None and row.get('href'):
        meta = meta or {}
        meta.setdefault('_filename', filename(row, meta))
        outdir.mkdir(parents=True, exist_ok=True)
        got = download_file(ctx, row, meta, outdir, "pdf")
        if got:
            log("落盘", f"{got}　{got.stat().st_size//1024} KB")
            return got
    return None


def fetch_one(ctx, kw, outdir, index=None, use_cache=False, pdf=None, pdf_only=False,
              redo=False):
    key = re.sub(r'[^\w]', '_', kw)[:60]
    if use_cache and (CACHE / f"{key}.html").exists():
        meta = json.loads((CACHE / f"{key}.json").read_text(encoding='utf-8'))
        html = (CACHE / f"{key}.html").read_text(encoding='utf-8')
        n = len(DB.parse(DB.BeautifulSoup(html, 'html.parser'))[2])
        imgs = {i: (CACHE / f"{key}_img{i}.png").read_bytes()
                for i in range(n) if (CACHE / f"{key}_img{i}.png").exists()}
        cached_n = meta.get('_img_count')
        if cached_n is not None and cached_n != n:
            # 图片按「正文中第几张」存盘，解析规则一改索引就对不上，重排会张冠李戴
            log("！注意", f"缓存图片 {cached_n} 张，当前解析出 {n} 张，索引可能错位，"
                          f"建议去掉 --cache 联网重抓")
        out = REC.existing_path(outdir / meta['_filename'])
        st = DB.build(html, meta, imgs, out)
        warn_notes(st)
        log("缓存重排", f"{out.name}　{st['paras']} 段 · 脚注 {st['notes']} 条"
                        + ("＋地脚星号注" if st['fund'] else "")
                        + (f" · 参考文献 {st['refs']} 条" if st.get('refs') else "")
                        + f" · 字形图 {st['imgs']} 个")
        return out

    if use_cache:      # --cache 但没有对应缓存，直说，别拿 ctx=None 往下走
        log("无缓存", f"{kw}　缓存里没有这一篇，去掉 --cache 联网抓")
        return None

    log("检索", kw)
    rows = search(ctx, kw)
    if not rows:
        log("无结果", "知网篇名检索无此词：可能未收录，或换个写法（简繁、更短）再试"); return None
    if len(rows) > 1 and index is None:
        exact = [i for i, r in enumerate(rows) if r['name'].strip() == kw.strip()]
        if len(exact) == 1:
            index = exact[0]                    # 篇名精确命中，不必人工挑
    if len(rows) > 1 and index is None:
        log("命中多条", "用 --index N 指定（从 0 起）：")
        for i, r in enumerate(rows[:10]):
            print(f"  [{i}] {r['name']}｜{r.get('author','')}｜{r.get('source','')} {r.get('date','')}")
        return None
    row = rows[index or 0]
    log("命中", f"{row['name']}｜{row.get('author','')}｜{row.get('source','')} {row.get('date','')}")
    if not row.get('ai') and not pdf_only:
        return note_manual(kw, row, "无 AI 阅读入口", ctx, None, outdir)

    meta = detail_meta(ctx, row)
    meta.update({k: row[k] for k in ('name', 'author', 'href') if row.get(k)})
    log("元数据", meta.get('cite') or "(未解析到刊期)")
    docx_p = REC.existing_path(outdir / filename(row, meta))
    pdf_p = docx_p.with_suffix('.' + (pdf or 'pdf'))
    meta['_filename'] = filename(row, meta)
    want_docx, want_original = not pdf_only, bool(pdf)
    docx_ready = REC.docx_ready(docx_p, DB.BUILD_VERSION)
    original_ready = QA.valid_original(pdf_p, pdf or 'pdf')
    if not redo and (not want_docx or docx_ready) and (not want_original or original_ready):
        REC.record(docx_p if docx_p.exists() else pdf_p, meta)
        log("已有", f"{(pdf_p if pdf_only else docx_p).name}　格式均完整，跳过")
        return pdf_p if pdf_only else docx_p
    if pdf_only:
        outdir.mkdir(parents=True, exist_ok=True)
        if not redo and original_ready:
            return pdf_p
        f = download_file(ctx, row, meta, outdir, pdf or "pdf")
        if f: log("落盘", f"{f}　{f.stat().st_size//1024} KB")
        return f
    if not redo and docx_ready:
        log("续抓", "DOCX 已有验收记录，只补缺失的原版文件")
        if want_original and not original_ready:
            f = download_file(ctx, row, meta, outdir, pdf)
            if f:
                log("原版", f"{f.name}　{f.stat().st_size//1024} KB")
        return docx_p
    html, imgs = full_text(ctx, row)
    if not html:
        return note_manual(kw, row, "AI 阅读页无 XML 正文", ctx, meta, outdir)

    total = len(DB.parse(DB.BeautifulSoup(html, 'html.parser'))[2])
    if len(imgs) < total:
        log("警告", f"{total - len(imgs)}/{total} 个生僻字图取不到，正文用 ▯ 占位")

    meta['_img_count'] = total
    CACHE.mkdir(parents=True, exist_ok=True)
    QA.write_text(CACHE / f"{key}.html", html)
    QA.write_text(CACHE / f"{key}.json", json.dumps(meta, ensure_ascii=False))
    for i, b in imgs.items():
        QA.write_bytes(CACHE / f"{key}_img{i}.png", b)

    outdir.mkdir(parents=True, exist_ok=True)
    out = docx_p
    st = DB.build(html, meta, imgs, out)
    warn_notes(st)
    log("正文", f"{st['paras']} 段 · 脚注 {st['notes']} 条"
              + ("＋基金星号注" if st['fund'] else "")
              + (f" · 参考文献 {st['refs']} 条" if st.get('refs') else "")
              + f" · 表格 {st.get('tables', 0)} 个 · 生僻字图 {st['imgs']} 个 · {'繁体' if st['trad'] else '简体'}"
              + f" · 验收 {st['quality']['status']}")
    log("落盘", str(out))
    if want_original and (redo or not original_ready):
        f = download_file(ctx, row, meta, outdir, pdf)
        if f:
            log("原版", f"{f.name}　{f.stat().st_size//1024} KB")
    return out


def main():
    ap = argparse.ArgumentParser(prog="cnki", description="知网全文抓取，默认只下原版 PDF，--docx 另出 DOCX")
    ap.add_argument("keyword", nargs="?", help="篇名关键词")
    ap.add_argument("-f", "--file", nargs="?", const=str(SKILL / "待抓.md"),
                    help="批量：每行一个篇名（默认读 skill 目录下 待抓.md）")
    ap.add_argument("-o", "--out", default=str(DEFAULT_OUT), help=f"输出目录（默认 {DEFAULT_OUT}）")
    ap.add_argument("-i", "--index", type=int, help="命中多条时取第 N 条（从 0 起）")
    ap.add_argument("--search-only", action="store_true", help="只检索列候选，不抓")
    g = ap.add_argument_group("检索（给了下面任一项就是列表模式，只列不抓）")
    g.add_argument("-t", "--title", help="按篇名检索（只列不抓）")
    g.add_argument("-a", "--author", help="按作者检索")
    g.add_argument("-j", "--journal", help="按文献来源（刊名）检索")
    g.add_argument("-k", "--subject", help="按主题检索")
    g.add_argument("--latest", action="store_true", help="只列最新一期（配合 -j）")
    g.add_argument("--since", help="起始年月，如 2021 或 2021-06")
    g.add_argument("--until", help="截止年月")
    g.add_argument("-n", "--limit", type=int, default=20, help="最多列出多少条（默认 20）")
    g.add_argument("--field", dest="subject_facet", help="按学科筛，如「中国古代史」（先用 --facets 看有哪些）")
    g.add_argument("--core", help="按来源类别筛：CSSCI / 北大核心 / AMI / CSCD")
    g.add_argument("--type", default="论文", choices=["论文", "期刊", "辑刊", "学位", "全部"],
                   help="文献类型：论文=期刊+辑刊（默认）/ 期刊 / 辑刊 / 学位 / 全部")
    g.add_argument("--fanti", action="store_true",
                   help="强制做繁简变体检索（默认只有作者检索做——篇名等字段知网自己会繁简归一）")
    g.add_argument("--no-fanti", action="store_true", help="关掉繁简变体检索")
    g.add_argument("--facets", action="store_true", help="列出本次检索可用的学科等分面选项")
    g.add_argument("--source", help="只保留「文献来源」含该词的（同名作者混杂时用）")
    g.add_argument("--to-list", action="store_true", help="把结果篇名写进 待抓.md，供 cnki -f 批量抓")
    ap.add_argument("--docx", action="store_true", help="另出 DOCX，与原版双落盘（默认只下原版 PDF）")
    ap.add_argument("--no-pdf", action="store_true", help="不下原版，只出 DOCX")
    # 旧默认是双落盘时留下的开关，现在默认即只下原版，保留免得旧脚本报错
    ap.add_argument("--pdf", "--pdf-only", "--no-docx", dest="pdf_compat", action="store_true",
                    help=argparse.SUPPRESS)
    ap.add_argument("--caj", action="store_true", help="原版下 CAJ 而不是 PDF")
    ap.add_argument("--redo", action="store_true", help="已抓过的也重抓（默认跳过）")
    ap.add_argument("--cache", action="store_true", help="用本地缓存重排，不联网（调格式用）")
    ap.add_argument("--open", action="store_true", help="弹出有头窗口，人工过验证/登录一次")
    ap.add_argument("--doctor", action="store_true", help="环境自检")
    a = ap.parse_args()
    outdir = Path(a.out).expanduser()

    if a.doctor:
        log("Python", sys.version.split()[0])
        for m in ("cloakbrowser", "cv2", "docx", "bs4", "PIL"):
            try:
                __import__(m); log(m, "ok")
            except Exception as e:
                log(m, f"缺失 {e}")
        log("登录态", "有" if (PROFILE / "Default").exists() else "无（首次运行会自动建）")
        log("输出目录", str(outdir))
        return

    if a.open:
        acquire_lock()
        ctx = cloakbrowser.launch_persistent_context(user_data_dir=str(PROFILE), headless=False)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        pg.goto("https://kns.cnki.net/kns8s/AdvSearch")
        log("有头模式", "请在窗口内完成验证/登录，完成后回车关闭")
        input()
        ctx.close(); return

    field = ("title" if a.title else "author" if a.author else
             "journal" if a.journal else "subject" if a.subject else None)
    if field:
        term = a.title or a.author or a.journal or a.subject
        PROFILE.mkdir(parents=True, exist_ok=True)
        acquire_lock()
        ctx = cloakbrowser.launch_persistent_context(user_data_dir=str(PROFILE), headless=True,
                                                accept_downloads=True)
        try:
            rows, total = SR.query(ctx, term, field, a.limit, a.since, a.until, a.latest,
                                   a.subject_facet, a.core, a.facets,
                                   getattr(a, 'type'),
                                   (not a.no_fanti) and (a.fanti or field == "author"))
        except (NET.PauseBatch, NET.PageTimeout) as e:
            extra = f"；建议等待 {e.retry_after} 秒" if getattr(e, 'retry_after', None) else ""
            log("暂停", str(e) + extra); return
        finally:
            ctx.close()
        head = {"title": "篇名", "author": "作者", "journal": "来源", "subject": "主题"}[field]
        if a.facets:
            log("分面", f"{head}＝{term}")
            for grp, items in rows:
                print(f"  {grp}：" + "　".join(items[:20]))
            return
        if a.source:
            rows = [r for r in rows if a.source in (r.get('source') or '')]
        span = (f"　{a.since or ''}—{a.until or ''}" if (a.since or a.until) else "")
        log("检索", f"{head}＝{term}{span}　类型＝{getattr(a, 'type')}")
        if a.latest and rows:
            log("最新一期", f"{rows[0]['date']}　{len(rows)} 篇")
        for i, r in enumerate(rows):
            print(f"  [{i:>2}] {r['name']}｜{r.get('author','')}｜{r.get('source','')} "
                  f"{r.get('date','')}｜{'有AI阅读入口' if r.get('ai') else '仅PDF'}")
        log("小计", f"{len(rows)} 条")
        if a.to_list:
            f = SKILL / "待抓.md"
            old = f.read_text(encoding='utf-8') if f.exists() else "# 待抓篇名清单\n\n"
            add = "".join(f"- [ ] {r['name']}\n" for r in rows if r['name'] not in old)
            f.write_text(old + add, encoding='utf-8')
            log("写入", f"{add.count(chr(10))} 条进 待抓.md（跳过已在清单里的）")
        return

    if not a.keyword and not a.file:
        ap.print_help(); return

    PROFILE.mkdir(parents=True, exist_ok=True)
    if a.cache:
        todo = [a.keyword] if a.keyword else []
        if a.file:
            for ln in Path(a.file).read_text(encoding='utf-8').splitlines():
                ln = re.sub(r'^[-*]\s*(\[[ x]\]\s*)?', '', ln).strip()
                if ln and not ln.startswith("#"):
                    todo.append(ln)
        for kw in todo:
            try:
                fetch_one(None, kw, outdir, a.index, use_cache=True)
            except Exception as e:
                log("失败", f"{kw}：{e}")
        return

    acquire_lock()
    ctx = cloakbrowser.launch_persistent_context(user_data_dir=str(PROFILE), headless=True,
                                                accept_downloads=True)
    try:
        if a.search_only:
            try:
                rows = search(ctx, a.keyword)
            except (NET.PauseBatch, NET.PageTimeout) as e:
                log("暂停", str(e)); return
            if not rows:
                log("无结果", "知网篇名检索无此词：可能未收录，或换个写法（简繁、更短）再试")
            for i, r in enumerate(rows):
                print(f"  [{i}] {r['name']}｜{r.get('author','')}｜{r.get('source','')} "
                      f"{r.get('date','')}｜{'有AI阅读入口' if r.get('ai') else '仅PDF'}")
            return
        todo = [a.keyword] if a.keyword else []
        if a.file:
            for ln in Path(a.file).read_text(encoding='utf-8').splitlines():
                ln = re.sub(r'^[-*]\s*(\[[ x]\]\s*)?', '', ln).strip()
                if ln and not ln.startswith("#"):
                    todo.append(ln)
        ok = 0
        progress = [{"keyword": kw, "status": "pending"} for kw in todo]
        write_progress(outdir, progress, status="running")
        paused = False
        for n, kw in enumerate(todo, 1):
            if len(todo) > 1:
                print(f"\n—— {n}/{len(todo)} ——")
            progress[n - 1]["status"] = "running"
            write_progress(outdir, progress, current=n, status="running")
            try:
                if fetch_one(ctx, kw, outdir, a.index if len(todo) == 1 else None,
                             pdf=(None if a.no_pdf else ("caj" if a.caj else "pdf")),
                             pdf_only=not (a.docx or a.no_pdf), redo=a.redo):
                    ok += 1
                    progress[n - 1]["status"] = "done"
                else:
                    progress[n - 1]["status"] = "failed"
            except (NET.PauseBatch, NET.PageTimeout) as e:
                progress[n - 1]["status"] = "paused"
                write_progress(outdir, progress, current=n, status="paused", reason=e)
                extra = f"；建议等待 {e.retry_after} 秒" if getattr(e, 'retry_after', None) else ""
                log("暂停", str(e) + extra)
                paused = True
                break
            except Exception as e:
                progress[n - 1]["status"] = "failed"
                progress[n - 1]["error"] = str(e)[:300]
                log("失败", f"{kw}：{e}")
            write_progress(outdir, progress, current=n, status="running")
            if n < len(todo):
                time.sleep(3)
        write_progress(outdir, progress, current=(n if todo else None),
                       status="paused" if paused else "done")
        if len(todo) > 1:
            log("\n合计", f"{ok}/{len(todo)} 篇成功" + ("，已暂停" if paused else ""))
    finally:
        ctx.close()


if __name__ == '__main__':
    main()
