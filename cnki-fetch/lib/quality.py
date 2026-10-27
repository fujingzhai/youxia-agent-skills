# -*- coding: utf-8 -*-
"""本地成果验收。验证 DOM → DOCX，不把它冒充 PDF → XML 的全文校勘。"""
import collections
import hashlib
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from bs4 import BeautifulSoup

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'


def norm(text):
    return re.sub(r'\s+', '', text).replace(':', '：').replace('關鍵詞', '关键词')


def fingerprint(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for part in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(part)
    return h.hexdigest()


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent,
                                     prefix='.' + path.name, suffix='.tmp', delete=False) as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        tmp = f.name
    os.replace(tmp, path)


def write_bytes(path, data):
    """原子写入缓存二进制，避免中断留下半张字形图。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(mode='wb', dir=path.parent,
                                     prefix='.' + path.name, suffix='.tmp', delete=False) as f:
        f.write(data)
        tmp = f.name
    os.replace(tmp, path)


def write_text(path, data):
    """原子写入缓存文本，避免 HTML/JSON 在断电后处于半写状态。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent,
                                     prefix='.' + path.name, suffix='.tmp', delete=False) as f:
        f.write(data)
        tmp = f.name
    os.replace(tmp, path)


def inspect_docx(path, html=None, dropped=()):
    """dropped：首部有意丢弃的西文，逐字核查要放行，否则每篇都误报缺文"""
    issues = []
    try:
        with zipfile.ZipFile(path) as z:
            if z.testzip():
                raise ValueError('ZIP 校验失败')
            body = ET.fromstring(z.read('word/document.xml'))
            foot = ET.fromstring(z.read('word/footnotes.xml'))
            # XML、关系目标均应能读取；图片由 python-docx 写入 media 部件。
            rel = ET.fromstring(z.read('word/_rels/document.xml.rels'))
            names = set(z.namelist())
            for r in rel:
                if r.get('TargetMode') != 'External' and r.get('Type', '').endswith(('/image', '/footnotes')):
                    target = r.get('Target', '')
                    if 'word/' + target not in names:
                        issues.append('关系目标缺失：' + target)
    except Exception as e:
        return {'status': 'invalid', 'issues': ['DOCX 结构无效：' + str(e)],
                'pdf_compared': False}
    body_text = ''.join(x.text or '' for x in body.iter(W + 't'))
    foot_text = ''.join(x.text or '' for x in foot.iter(W + 't'))
    joined = norm(body_text + foot_text)
    refs = collections.Counter(x.get(W + 'id') for x in body.iter(W + 'footnoteReference'))
    notes = {x.get(W + 'id'): x for x in foot.findall('w:footnote', NS)
             if int(x.get(W + 'id', '0')) > 0}
    if set(refs) != set(notes) or any(n != 1 for n in refs.values()):
        issues.append('脚注引用与脚注部件不对应')
    if any(not norm(''.join(t.text or '' for t in x.iter(W + 't'))).strip('*') for x in notes.values()):
        issues.append('存在空脚注')
    placeholders = {m: body_text.count(m) for m in ('▯', '［图］', '［未支持的图片］') if m in body_text}
    placeholders['原始图片引用'] = len(re.findall(r'［图片：[^］]*］', body_text))
    if not placeholders['原始图片引用']:
        placeholders.pop('原始图片引用')
    if placeholders:
        issues.append('存在缺图/缺字占位符')
    table_count = len(list(body.iter(W + 'tbl')))
    image_count = len(list(body.iter(W + 'drawing')))
    missing = []
    source = {}
    if html is not None:
        soup = BeautifulSoup(html, 'html.parser')
        containers = soup.select('div.ChapterWrap.Content')
        source = {'tables': sum(len(c.find_all('table')) for c in containers),
                  'images': len(soup.select('div.ChapterWrap.Content img[src*="attachment"]')),
                  'footnotes': len(soup.select('div.ChapterWrap.Content sup.footNote'))}
        if not containers or not any(c.get_text(strip=True) for c in containers):
            issues.append('原始正文容器为空')
        if table_count != source['tables']:
            issues.append('原生表格数量不符')
        if image_count != source['images']:
            issues.append('图片数量不符')
        # 独立于转换器的逐文本节点检查，能捕获未识别标签、表格单元格等漏项。
        skip = [norm(d) for d in dropped]
        for c in containers:
            for t in c.find_all(string=True):
                if t.find_parent(['script', 'style']):
                    continue
                sup = t.find_parent('sup')
                previous = sup.previous_sibling if sup else None
                while previous is not None and not getattr(previous, 'name', None) and not str(previous).strip():
                    previous = previous.previous_sibling
                ref_page = (sup is not None and getattr(previous, 'name', None) == 'sup'
                            and previous.select_one('a[type="reference"]'))
                if sup and ('footNote' in (sup.get('class') or []) or
                            sup.select_one('a[type="reference"]') or ref_page or
                            re.match(r'^\([Pp][^)]+\)', sup.get_text().strip())):
                    continue
                val = norm(re.sub(r'^\[\d+\]', '', str(t)))
                if len(val) >= 4 and val not in joined and not any(val in d for d in skip):
                    missing.append(val[:160])
        if missing:
            issues.append('原始文本片段未在 DOCX 找到')
        for n in soup.select('div.ChapterWrap.Content sup.footNote'):
            val = re.sub(r'^\(\d+\)\s*', '', n.get('footnotetitle', '').strip())
            if not val or norm(val) not in norm(foot_text):
                issues.append('原始脚注正文缺失')
                break
    return {'status': 'review_required' if issues else 'verified', 'issues': issues,
            'source': source, 'docx': {'tables': table_count, 'images': image_count,
                                     'footnotes': len(notes), 'text_chars': len(body_text)},
            'placeholders': placeholders, 'missing_text_count': len(missing),
            'missing_text_samples': missing[:20], 'pdf_compared': False,
            'scope': '缓存 DOM 到 DOCX 的本地验收'}


def valid_original(path, fmt='pdf'):
    p = Path(path)
    if not p.is_file() or p.stat().st_size < 100:
        return False
    with p.open('rb') as f:
        start = f.read(16)
        if fmt == 'pdf':
            if not start.startswith(b'%PDF-'):
                return False
            f.seek(max(0, p.stat().st_size - 2048))
            return b'%%EOF' in f.read()
        return start.startswith((b'CAJ', b'HN', b'KDH', b'%PDF-'))
