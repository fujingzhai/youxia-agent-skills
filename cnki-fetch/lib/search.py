# -*- coding: utf-8 -*-
"""知网检索：按篇名/作者/来源/主题查，支持学科等分面、时间范围与翻页。

结果按发表时间倒序，所以 --since 可以边翻边停。
左侧分面（学科、来源类别、研究层次…）结构是 dl[groupid] > dt b + dd li > a + span(计数)。
"""
import re
from urllib.parse import quote
from opencc import OpenCC
import captcha
import network as NET

FIELDS = {"title": "TI", "author": "AU", "journal": "LY", "subject": "SU", "keyword": "KY"}
URL = "https://kns.cnki.net/kns8s/defaultresult/index?kw={}&korder={}"

# 实测的库代码：CJFQ 期刊、CCJD 辑刊（《文史》这类）、CDMD 学位论文（博+硕）
DBCODES = {"论文": "CJFQ,CCJD", "期刊": "CJFQ", "辑刊": "CCJD", "学位": "CDMD", "全部": ""}

_S2T, _T2S = OpenCC('s2t'), OpenCC('t2s')


def variants(term, cap=8):
    """繁简变体。知网是字面精确匹配，不做繁简归一，而且「全繁」并不等于 opencc 的 s2t：

    国家的繁体字形规范里，有些字的繁体正字就是那个简体字形——「吴」的规范繁体仍是「吴」而非「吳」。
    所以繁体文章写的是「吴揚廣」，查「吳揚廣」一条也出不来。这与姓/名的位置无关，纯是字形规范问题。
    这类字不多但无法预先穷举，所以逐字取简/繁做笛卡尔积，全部查一遍再合并。
    """
    t, s = _S2T.convert(term), _T2S.convert(term)
    if len(t) != len(term) or len(s) != len(term):      # 词级转换改了长度，退化
        return list(dict.fromkeys([term, t, s]))
    per = [sorted({term[i], t[i], s[i]}) for i in range(len(term))]
    n = 1
    for c in per:
        n *= len(c)
    if n > cap:
        return list(dict.fromkeys([term, t, s]))
    out = [""]
    for c in per:
        out = [a + b for a in out for b in c]
    return list(dict.fromkeys([term] + out))

ROWS_JS = r"""()=>Array.from(document.querySelectorAll('table.result-table-list tbody tr')).map(tr=>({
    href: tr.querySelector('.name a')?.href,
    name: tr.querySelector('.name a')?.innerText?.trim(),
    author: tr.querySelector('.author')?.innerText?.trim().replace(/\s+/g,' '),
    source: tr.querySelector('.source')?.innerText?.trim(),
    date: tr.querySelector('.date')?.innerText?.trim(),
    ai: tr.querySelector('.operat a.icon-airead')?.href})).filter(r=>r.name)"""

# 分页认 #PageNext。别用 a.cur——那是左侧「主要主题/次要主题」的分组切换。
# 也别信 .search-page 里的 <em>：那不是总页数（6327 条的刊也显示 10），据此判停会空翻。
NEXT_JS = r"""()=>{const a=document.querySelector('#PageNext')||document.querySelector('.search-page .but-r');
    if(!a || a.className.includes('disabled')) return false;
    a.click(); return true;}"""

FIRST_JS = "()=>document.querySelector('table.result-table-list tbody tr .name a')?.href||''"

FACET_JS = r"""()=>{const box=document.querySelector('#divGroup')||document.querySelector('#ModuleGroupFilter');
    if(!box) return [];
    const out=[];
    for(const dl of box.querySelectorAll('dl[groupid]')){
        const g=(dl.querySelector('dt b')?.innerText||'').trim();
        const items=[...dl.querySelectorAll('dd li')].map(li=>{
            const a=li.querySelector('a'), s=li.querySelector('span');
            return ((a?.innerText||'').trim()+(s?s.innerText.trim():'')).replace(/\s+/g,'');
        }).filter(Boolean);
        if(g && items.length) out.push([g, items]);
    }
    return out;}"""

EXPAND_JS = r"""(g)=>{const box=document.querySelector('#divGroup')||document.querySelector('#ModuleGroupFilter');
    for(const dl of box.querySelectorAll('dl[groupid]')){
        if(((dl.querySelector('dt b')?.innerText)||'').trim()!==g) continue;
        if(!dl.querySelector('dd li')) dl.querySelector('dt')?.click();
        return true;}
    return false;}"""

APPLY_JS = r"""([g,v])=>{const box=document.querySelector('#divGroup')||document.querySelector('#ModuleGroupFilter');
    for(const dl of box.querySelectorAll('dl[groupid]')){
        if(((dl.querySelector('dt b')?.innerText)||'').trim()!==g) continue;
        for(const li of dl.querySelectorAll('dd li')){
            const a=li.querySelector('a'); const t=(a?.innerText||'').trim();
            if(t===v||t.startsWith(v)){ a.click(); return t; }
        }
    }
    return null;}"""


def list_facets(page):
    try:
        return page.evaluate(FACET_JS)
    except Exception as e:
        print("分面读取失败:", e)
        return []


def apply_facet(page, group, value):
    """点左侧分面。有些组是懒加载的，先点标题展开一次。"""
    try:
        page.evaluate(EXPAND_JS, group)
        page.wait_for_timeout(1200)
        hit = page.evaluate(APPLY_JS, [group, value])
    except Exception:
        hit = None
    if hit:
        page.wait_for_timeout(3500)
    return hit


def query(ctx, kw, field="title", limit=20, since=None, until=None, latest=False,
          subject_facet=None, core=None, show_facets=False, dbtype="论文", fanti=True):
    """→ (rows, total)。繁简变体逐个查后合并去重。"""
    # 实测：篇名 TI 繁简互通（论衡/論衡 结果一致），作者 AU 不通。所以默认只对作者展开变体
    terms = variants(kw) if (fanti and not show_facets) else [kw]
    if len(terms) > 1:
        print(f"繁简变体    {len(terms)} 个：" + "　".join(terms), flush=True)
    all_rows, seen, total = [], set(), None
    for i, t in enumerate(terms, 1):
        rows, tot = _one(ctx, t, field, limit, since, until, latest,
                         subject_facet, core, show_facets, dbtype)
        if show_facets:
            return rows, tot
        total = total or tot
        new = 0
        for r in rows:
            k = (r.get('name'), r.get('date'))
            if k in seen:
                continue
            seen.add(k); all_rows.append(r); new += 1
        if len(terms) > 1:
            print(f"  变体{i}/{len(terms)} {t}：{len(rows)} 条，新增 {new}", flush=True)
    all_rows.sort(key=lambda r: r.get('date') or '', reverse=True)
    return _trim(all_rows, limit, latest), total


def _one(ctx, kw, field, limit, since, until, latest, subject_facet, core, show_facets, dbtype):
    page = ctx.new_page()
    url = URL.format(quote(kw), FIELDS.get(field, "TI"))
    db = DBCODES.get(dbtype, "")
    if db:
        url += "&crossDbcodes=" + db
    guard = NET.Guard(page)
    page.goto(url, wait_until="domcontentloaded")
    if not NET.wait_content(page, 'table.result-table-list tbody tr', guard,
                            empty_selector='.no-content, .noresult, .no-result, #noResult, .no-data'):
        page.close()
        return [], None

    if show_facets:
        f = list_facets(page)
        page.close()
        return f, None

    for grp, val in (("学科", subject_facet), ("来源类别", core)):
        if val:
            hit = apply_facet(page, grp, val)
            print("分面        " + (f"{grp}＝{hit}" if hit else f"{grp}＝{val} 未命中，已忽略"), flush=True)

    out, seen, pages, prev_first = [], set(), 0, None
    total = None
    while pages < 40:
        guard.check()
        pages += 1
        page_rows = page.evaluate(ROWS_JS)
        first = page_rows[0].get('href') if page_rows else None
        if first and first == prev_first:      # 翻页没生效，别再转
            break
        prev_first = first
        for r in page_rows:
            k = r.get('href') or r['name']
            if k in seen:
                continue
            seen.add(k)
            d = r.get('date') or ''
            if until and d[:len(until)] > until:
                continue
            if since and d[:len(since)] < since:      # 倒序，越界即可收工
                page.close()
                return _trim(out, limit, latest), total
            out.append(r)
        if latest and out:
            if any(r['date'] != out[0]['date'] for r in out):
                break
        elif len(out) >= limit:
            break
        if not page.evaluate(NEXT_JS):
            break
        try:    # 等首条真的换掉，别用固定 sleep——没加载完就取会漏掉整页
            page.wait_for_function(
                "prev => {const a=document.querySelector('table.result-table-list tbody tr .name a');"
                "return a && a.href !== prev;}", arg=first, timeout=10000)
        except Exception:
            break
        NET.ensure_access(page, guard)
    page.close()
    return _trim(out, limit, latest), total


def _trim(rows, limit, latest):
    if latest and rows:
        top = rows[0]['date']
        return [r for r in rows if r['date'] == top]
    return rows[:limit]
