# -*- coding: utf-8 -*-
"""统一抓取台账：UTF-8 BOM CSV，锁定读改写并原子替换；完整验收报告保存在单元格内。"""
import csv
import fcntl
import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import quality as QA

DEFAULT_ROOT = Path(os.environ.get('CNKI_OUTPUT_DIR') or Path.home() / 'Downloads/知网文献').expanduser()
CSV_NAME = '抓取记录.csv'
FIELDS = ['记录标识', '篇名', '作者', '刊物', '年份', '期号', 'DOI', '来源链接',
          'DOCX路径', 'PDF路径', 'CAJ路径', '原版状态', '验收状态', '问题说明',
          '抓取时间', '记录更新时间', 'DOCX校验值', '生成版本', '验收详情']
RETIRED = ('精简模式',)     # 首部体例已定死，旧台账读到就丢掉，下次落盘时列自动消失
STATES = {'verified': '通过', 'review_required': '待核对', 'invalid': '失败'}


def root_for(path):
    directory = Path(path).expanduser().resolve().parent
    default = DEFAULT_ROOT.resolve()
    if directory == default or default in directory.parents:
        return default
    for parent in (directory, *directory.parents):
        if (parent / CSV_NAME).is_file():
            return parent
    return directory


def read_rows(root):
    path = Path(root) / CSV_NAME
    if not path.exists():
        return []
    with path.open(encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        if [c for c in (reader.fieldnames or []) if c not in RETIRED] != FIELDS:
            raise ValueError(f'台账列名不匹配，保留原文件：{path}')
        rows = [{k: v for k, v in row.items() if k not in RETIRED} for row in reader]
    if any(None in row or any(v is None for v in row.values()) for row in rows):
        raise ValueError(f'台账行损坏，保留原文件：{path}')
    return rows


def _write_rows(root, rows):
    # 同一目录锁不会因 CSV 原子替换失效，也不会增加用户可见的锁文件。
    data = io.StringIO(newline='')
    writer = csv.DictWriter(data, fieldnames=FIELDS)
    writer.writeheader()
    writer.writerows(rows)
    QA.write_bytes(Path(root) / CSV_NAME, data.getvalue().encode('utf-8-sig'))


def _key(path, root):
    return Path(path).resolve().relative_to(root).with_suffix('').as_posix()


def _find(rows, key):
    exact = [r for r in rows if r['记录标识'] == key]
    if exact:
        return exact[0]
    # 分类移动后允许唯一同名成果继续使用原记录，避免根目录重复落盘。
    same = [r for r in rows if Path(r['记录标识']).name == Path(key).name]
    return same[0] if len(same) == 1 else None


def existing_path(path):
    path = Path(path)
    if path.exists():
        return path
    root = root_for(path)
    rows = read_rows(root)
    key = _key(path, root)
    row = _find(rows, key)
    if row is None and sum(Path(r['记录标识']).name == Path(key).name for r in rows) > 1:
        raise ValueError(f'台账中存在多份同名成果，请用 -o 指定所属刊物目录：{path.name}')
    if row:
        for field in ('DOCX路径', 'PDF路径', 'CAJ路径'):
            if row[field] and (root / row[field]).is_file():
                return (root / row[field]).with_suffix(path.suffix)
    return path


def docx_ready(path, build_version):
    path = Path(path)
    if not path.is_file():
        return False
    root = root_for(path)
    row = _find(read_rows(root), _key(path, root))
    return bool(row and row['验收状态'] in ('通过', '待核对')
                and row['生成版本'] == str(build_version)
                and row['DOCX校验值'] == QA.fingerprint(path))


def record(path, meta=None, report=None, error=None, timestamp=None):
    path = Path(path).expanduser().resolve()
    root = root_for(path)
    root.mkdir(parents=True, exist_ok=True)
    meta = meta or {}
    key = _key(path, root)
    lock = os.open(root, os.O_RDONLY)
    try:
        fcntl.flock(lock, fcntl.LOCK_EX)
        rows = read_rows(root)
        row = next((r for r in rows if r['记录标识'] == key), None)
        if row is None:
            row = dict.fromkeys(FIELDS, '')
            rows.append(row)
        row['记录标识'] = key
        parts = path.stem.rsplit('_', 2)
        title, author, tail = parts if len(parts) == 3 else (path.stem, '', '')
        match = re.fullmatch(r'(.*?)(\d{4})(?:-(\d+))?', tail)
        values = {'篇名': meta.get('name') or title,
                  '作者': meta.get('author') or re.split('[（(]', meta.get('author_line', ''))[0] or author,
                  '刊物': meta.get('journal') or (match[1] if match else ''),
                  '年份': meta.get('year') or (match[2] if match else ''),
                  '期号': meta.get('issue') or (match[3] if match else ''),
                  'DOI': meta.get('doi'), '来源链接': meta.get('href')}
        row.update({k: str(v) for k, v in values.items() if v})
        now = datetime.now(ZoneInfo('Asia/Shanghai')).isoformat(timespec='seconds')
        row['记录更新时间'] = now
        if path.exists() and not row['抓取时间']:
            row['抓取时间'] = timestamp or now
        for suffix, field in (('.docx', 'DOCX路径'), ('.pdf', 'PDF路径'), ('.caj', 'CAJ路径')):
            candidate = path.with_suffix(suffix)
            row[field] = candidate.relative_to(root).as_posix() if candidate.is_file() else ''
        originals = [(suffix, QA.valid_original(path.with_suffix('.' + suffix), suffix))
                     for suffix in ('pdf', 'caj') if path.with_suffix('.' + suffix).exists()]
        row['原版状态'] = '；'.join(f'{fmt.upper()} ' + ('通过' if ok else '无效')
                                     for fmt, ok in originals) or '缺原版'
        if report is not None:
            row.update({'验收状态': STATES.get(report.get('status'), '待核对'),
                        '问题说明': '；'.join(report.get('issues', [])),
                        'DOCX校验值': report.get('sha256', ''),
                        '生成版本': str(report.get('build_version', '')),
                        '验收详情': json.dumps(report, ensure_ascii=False, separators=(',', ':'))})
        elif not row['验收状态']:
            row['验收状态'] = '未验收' if row['DOCX路径'] else '仅原版'
        if error:
            row['问题说明'] = str(error)
            if not row['DOCX路径']:
                row['验收状态'] = '失败'
        _write_rows(root, rows)
        return dict(row)
    finally:
        os.close(lock)
