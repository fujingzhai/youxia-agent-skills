#!/usr/bin/env python3
"""B站登录态字幕探测、下载和清洗。

Arc/Chrome Cookie 通过 yt-dlp 的 Chromium cookie reader 进入内存；本脚本不会
打印、导出或保存 Cookie 值。没有可用字幕时保留 meta.json 并返回状态码 2。
"""

from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse

try:
    from yt_dlp import YoutubeDL
    from yt_dlp import cookies as ytdlp_cookies
    from yt_dlp.extractor.bilibili import BiliBiliIE
    from yt_dlp.networking import Request
except ImportError as exc:
    print(
        "缺少 yt-dlp 的 Python 库（bili 脚本需要 import yt_dlp）。\n"
        f"当前解释器：{sys.executable}\n"
        "安装命令：python3 -m pip install -U yt-dlp\n"
        "装好后重跑；或先执行 `yt check` 检查环境。",
        file=sys.stderr,
    )
    raise SystemExit(2) from exc

SUITE_ROOT = Path(__file__).resolve().parents[3]
SHARED_SCRIPTS = SUITE_ROOT / "scripts"
sys.path.insert(0, str(SHARED_SCRIPTS))

from clean import parse_cues, to_paragraphs  # noqa: E402
from common import TMP_ROOT  # noqa: E402
BVID_RE = re.compile(r"BV[0-9A-Za-z]{10}")
ZH_RE = re.compile(r"(?:^|[-_])(zh|yue)(?:$|[-_])", re.I)
INTERVIEW_RE = re.compile(
    r"访谈|訪談|采访|採訪|对话|對話|播客|podcast|演讲|演講|圆桌|圓桌|讲座|講座|"
    r"论坛|論壇|口述|实录|實錄|商业访谈录|商業訪談錄",
    re.I,
)
TRANSCRIPT_HINT_RE = re.compile(
    r"逐字稿|文字稿|文字版|全文|文稿|实录|實錄|访谈稿|訪談稿|公众号|公眾號|"
    r"transcript|show\s*notes",
    re.I,
)
URL_RE = re.compile(r"https?://[^\s<>()\[\]，。、]+")
# 播客平台默认区分说话人，B站 AI 字幕不区分。描述栏挂了这些链接，
# 说明同一期内容在播客平台也有，底本质量通常更高，值得停下来问一句。
PODCAST_HOSTS = {"xiaoyuzhoufm.com": "小宇宙"}

NOISE_HOSTS = {
    "bilibili.com", "weibo.com", "x.com", "twitter.com", "douyin.com",
    "tiktok.com", "instagram.com", "youtube.com", "youtu.be", "music.163.com",
    "spotify.com", "apple.com",
}


class BiliError(RuntimeError):
    pass


class NoSubtitle(BiliError):
    pass


def die(message: str, code: int = 1) -> None:
    sys.stdout.flush()
    print(message, file=sys.stderr)
    raise SystemExit(code)


def parse_target(url: str) -> tuple[str, int]:
    match = BVID_RE.search(url)
    if not match:
        die("无法从链接中识别 BV 号。请使用 bilibili.com/video/BV... 链接。")
    query = parse_qs(urlparse(url).query)
    try:
        page = max(1, int((query.get("p") or ["1"])[0]))
    except ValueError:
        page = 1
    return match.group(0), page


def workdir(bvid: str, page: int, page_count: int | None = None) -> Path:
    suffix = f"_p{page}" if (page_count or 1) > 1 or page > 1 else ""
    path = TMP_ROOT / f"{bvid}{suffix}"
    path.mkdir(parents=True, exist_ok=True)
    return path


def browser_config() -> tuple[str, str | None]:
    browser = os.environ.get("BILI_BROWSER", "arc").strip().lower()
    profile_override = os.environ.get("BILI_BROWSER_PROFILE")
    if browser == "arc":
        if sys.platform != "darwin":
            die("Arc 的 Cookie 读取依赖 macOS 钥匙串；当前不是 macOS。请设置 BILI_BROWSER=chrome 并先登录 Chrome。")
        profile = profile_override or str(
            Path.home() / "Library" / "Application Support" / "Arc" / "User Data" / "Default"
        )
        return browser, profile
    if browser == "chrome":
        return browser, profile_override
    die("BILI_BROWSER 目前只支持 arc 或 chrome。")


@contextlib.contextmanager
def logged_in_ydl():
    """Yield a YoutubeDL whose cookie jar contains the browser's Bilibili login."""
    browser, profile = browser_config()
    opts = {"quiet": True, "no_warnings": True, "skip_download": True}
    original_settings = ytdlp_cookies._get_chromium_based_browser_settings
    try:
        with YoutubeDL(opts) as ydl:
            if browser == "arc":
                arc_root = str(Path(profile).parent.parent)

                def arc_settings(_browser_name):
                    return {
                        "browser_dir": arc_root,
                        "keyring_name": "Arc",
                        "supports_profiles": True,
                    }

                ytdlp_cookies._get_chromium_based_browser_settings = arc_settings
                jar = ytdlp_cookies._extract_chrome_cookies(
                    "chrome", profile, None, ytdlp_cookies.YDLLogger(ydl)
                )
            else:
                jar = ytdlp_cookies._extract_chrome_cookies(
                    "chrome", profile, None, ytdlp_cookies.YDLLogger(ydl)
                )
            ydl.cookiejar = jar
            yield ydl, browser
    finally:
        ytdlp_cookies._get_chromium_based_browser_settings = original_settings


def browser_label(browser: str) -> str:
    return "Arc" if browser == "arc" else "Chrome"


def get_json(ydl: YoutubeDL, url: str, referer: str = "https://www.bilibili.com/") -> dict:
    try:
        with ydl.urlopen(Request(url, headers={"Referer": referer})) as response:
            return json.load(response)
    except Exception as exc:
        raise BiliError(f"B站接口请求失败：{type(exc).__name__}: {exc}") from exc


def api(ydl: YoutubeDL, path: str, params: dict, referer: str) -> dict:
    payload = get_json(ydl, f"https://api.bilibili.com{path}?{urlencode(params)}", referer)
    if payload.get("code") != 0:
        raise BiliError(f"B站接口返回 code={payload.get('code')}：{payload.get('message')}")
    return payload.get("data") or {}


def validate_login(ydl: YoutubeDL, browser: str) -> dict:
    nav = get_json(ydl, "https://api.bilibili.com/x/web-interface/nav")
    data = nav.get("data") or {}
    if nav.get("code") != 0 or not data.get("isLogin"):
        raise BiliError(
            f"{browser_label(browser)} 中没有有效的B站登录 Cookie。请先登录B站后重试。"
        )
    return {"browser": browser_label(browser), "is_login": True}


def is_chinese_track(track: dict) -> bool:
    label = f"{track.get('lan') or ''} {track.get('lan_doc') or ''}"
    return bool(ZH_RE.search(f"-{label.lower()}-") or "中文" in label or "粤语" in label)


def track_kind(track: dict) -> str:
    lan = str(track.get("lan") or "")
    if track.get("type") == 0 and track.get("ai_status") in (None, 0):
        return "manual"
    if track.get("type") == 1 or track.get("ai_status") == 2 or lan.startswith("ai-"):
        return "ai-translated" if track.get("ai_type") == 1 else "ai-original"
    return "unknown"


def usable_tracks(raw_tracks: list[dict]) -> list[dict]:
    tracks = []
    for raw in raw_tracks:
        if not raw.get("subtitle_url"):
            continue
        item = {
            "lan": raw.get("lan"),
            "lan_doc": raw.get("lan_doc"),
            "type": raw.get("type"),
            "ai_type": raw.get("ai_type"),
            "ai_status": raw.get("ai_status"),
            "subtitle_url": raw.get("subtitle_url"),
        }
        item["kind"] = track_kind(item)
        item["is_chinese"] = is_chinese_track(item)
        tracks.append(item)
    return tracks


def pick_track(tracks: list[dict]) -> tuple[dict, str]:
    buckets = [
        (lambda t: t["kind"] == "manual" and t["is_chinese"], "人工 CC（中文）"),
        (lambda t: t["kind"] == "manual", "人工 CC"),
        (lambda t: t["kind"] == "ai-original" and t["is_chinese"], "B站平台 AI 字幕（原生中文）"),
        (lambda t: t["kind"] == "ai-original", "B站平台 AI 字幕（原生语言）"),
    ]
    for predicate, why in buckets:
        for track in tracks:
            if predicate(track):
                return track, why
    translated = [track for track in tracks if track["kind"] == "ai-translated"]
    if translated:
        raise NoSubtitle(
            "只有平台 AI 翻译轨，没有人工 CC 或原生 AI 字幕；为避免双重误差，不自动采用。"
        )
    raise NoSubtitle("没有带有效 subtitle_url 的人工 CC 或平台 AI 字幕。")


def clean_url(url: str) -> str:
    return url.rstrip(".,;，。；)]}")


def is_noise_url(url: str) -> bool:
    host = urlparse(url).netloc.lower().removeprefix("www.")
    return any(host == item or host.endswith(f".{item}") for item in NOISE_HOSTS)


def transcript_candidates(description: str) -> tuple[list[dict], list[str]]:
    lines = [line.strip() for line in (description or "").splitlines()]
    candidates, hints, seen = [], [], set()
    for index, line in enumerate(lines):
        if not TRANSCRIPT_HINT_RE.search(line):
            continue
        hints.append(line[:240])
        urls = URL_RE.findall(line)
        if not urls:
            for following in lines[index + 1:index + 3]:
                urls = URL_RE.findall(following)
                if urls:
                    break
        for url in urls:
            url = clean_url(url)
            if url in seen or is_noise_url(url):
                continue
            seen.add(url)
            candidates.append({"url": url, "context": line[:240]})
    return candidates, hints


def normalize_title(title: str) -> str:
    title = re.sub(r"^[【〖\[].{0,30}?[】〗\]]\s*", "", title or "")
    return re.sub(r"\s+", " ", title).strip()


def podcast_links(description: str) -> list[dict]:
    """从描述栏挖播客平台链接。

    和 transcript_candidates 不同，这里**不要求**所在行带"文字版/逐字稿"之类的字眼——
    播客链接通常是光秃秃挂在那里的。这是判断"同一期是不是也在播客平台"最可靠的信号，
    比拿标题去搜准得多（小宇宙的搜索接口是模糊匹配，噪音极大，不可用来做自动判定）。
    """
    found, seen = [], set()
    for url in URL_RE.findall(description or ""):
        url = clean_url(url)
        if url in seen:
            continue
        host = urlparse(url).netloc.lower().removeprefix("www.")
        for domain, label in PODCAST_HOSTS.items():
            if host == domain or host.endswith(f".{domain}"):
                seen.add(url)
                found.append({"platform": label, "url": url})
                break
    return found


def search_plan(title: str, owner: str, description: str, duration: int, subtitle_kind: str | None):
    candidates, hints = transcript_candidates(description)
    haystack = f"{title} {owner} {description[:1200]}"
    interview = bool(INTERVIEW_RE.search(haystack))
    long_ai = duration >= 1200 and subtitle_kind == "auto"
    podcasts = podcast_links(description)
    recommended = bool(candidates or hints or podcasts or interview or long_ai)
    reasons = []
    if podcasts:
        reasons.append("描述栏挂了播客平台链接")
    if candidates:
        reasons.append("描述栏有现成文稿候选链接")
    if hints:
        reasons.append("描述栏有文字版/全文线索")
    if interview:
        reasons.append("内容属于访谈/播客/演讲等高命中类型")
    if long_ai:
        reasons.append("长视频且当前底本是平台 AI 字幕")
    clean_title = normalize_title(title)
    queries = [
        f'"{clean_title}" 文字版 全文 实录 逐字稿',
        f'"{owner}" "{clean_title[:36]}" 文稿',
        f'site:mp.weixin.qq.com "{clean_title[:42]}"',
        f'site:xiaoyuzhoufm.com "{owner}"',
    ]
    return candidates, hints, podcasts, recommended, "；".join(dict.fromkeys(reasons)), queries


def format_duration(seconds: int) -> str:
    hours, rem = divmod(int(seconds or 0), 3600)
    minutes, secs = divmod(rem, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"


def srt_timestamp(seconds: float) -> str:
    millis = max(0, round(float(seconds) * 1000))
    hours, millis = divmod(millis, 3_600_000)
    minutes, millis = divmod(millis, 60_000)
    secs, millis = divmod(millis, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def write_subtitle_files(
    ydl: YoutubeDL,
    bvid: str,
    page: int,
    track: dict,
    outdir: Path,
    referer: str,
    make_transcript: bool,
):
    url = track["subtitle_url"]
    if url.startswith("//"):
        url = f"https:{url}"
    payload = get_json(ydl, url, referer)
    cues = payload.get("body") or []
    if not cues or not any(str(cue.get("content") or "").strip() for cue in cues):
        raise NoSubtitle("字幕 URL 可访问，但正文为空。")

    stem = f"{bvid}{f'_p{page}' if page > 1 else ''}.{track.get('lan') or 'subtitle'}"
    raw_path = outdir / "subtitle.json"
    srt_path = outdir / f"{stem}.srt"
    transcript_path = outdir / "transcript.txt"
    raw_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    blocks = []
    for index, cue in enumerate(cues, 1):
        content = str(cue.get("content") or "").replace("\r", " ").replace("\n", " ").strip()
        if not content:
            continue
        blocks.append(
            f"{index}\n{srt_timestamp(cue.get('from') or 0)} --> {srt_timestamp(cue.get('to') or 0)}\n{content}\n"
        )
    srt_path.write_text("\n".join(blocks), encoding="utf-8")

    paragraph_count = None
    if make_transcript:
        parsed = parse_cues(srt_path.read_text(encoding="utf-8"))
        paragraphs = to_paragraphs(parsed, "manual", timestamps=True)
        transcript_path.write_text("\n\n".join(paragraphs) + "\n", encoding="utf-8")
        paragraph_count = len(paragraphs)
    else:
        transcript_path = None
    return raw_path, srt_path, transcript_path, len(cues), paragraph_count


def save_meta(outdir: Path, meta: dict) -> None:
    (outdir / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="B站登录态字幕探测与文稿底稿生成")
    parser.add_argument("mode", choices=["probe", "subs", "doc"])
    parser.add_argument("url")
    parser.add_argument("--json", action="store_true", help="输出完整结果 JSON")
    args = parser.parse_args()

    bvid, requested_page = parse_target(args.url)
    referer = f"https://www.bilibili.com/video/{bvid}/"

    try:
        with logged_in_ydl() as (ydl, browser):
            login = validate_login(ydl, browser)
            video = api(ydl, "/x/web-interface/view", {"bvid": bvid}, referer)
            pages = video.get("pages") or []
            if not pages:
                raise BiliError("视频元数据没有分P信息。")
            if requested_page > len(pages):
                raise BiliError(f"请求 p={requested_page}，但该视频只有 {len(pages)} 个分P。")
            page = pages[requested_page - 1]
            cid = page.get("cid")
            # 字幕必须走当前 yt-dlp 提取器使用的 WBI 签名接口。普通
            # /x/player/v2 在 2026-08 实测可能返回属于其它视频的陈旧字幕 URL。
            # 这种错配表面上仍是 200 + 非空正文，不能靠“下载成功”识别。
            extractor = BiliBiliIE(ydl)
            signed = extractor._sign_wbi({"bvid": bvid, "cid": cid}, bvid)
            player_payload = extractor._download_json(
                "https://api.bilibili.com/x/player/wbi/v2",
                bvid,
                query=signed,
                headers=extractor._HEADERS,
                note=f"读取B站字幕元数据 {cid}",
            )
            if player_payload.get("code") != 0:
                raise BiliError(
                    f"B站字幕接口返回 code={player_payload.get('code')}：{player_payload.get('message')}"
                )
            player = player_payload.get("data") or {}
            raw_tracks = ((player.get("subtitle") or {}).get("subtitles") or [])
            tracks = usable_tracks(raw_tracks)
            outdir = workdir(bvid, requested_page, len(pages))

            selected = None
            why = None
            no_subtitle_reason = None
            try:
                selected, why = pick_track(tracks)
            except NoSubtitle as exc:
                no_subtitle_reason = str(exc)

            subtitle_kind = None
            if selected:
                subtitle_kind = "manual" if selected["kind"] == "manual" else "auto"
            description = video.get("desc") or ""
            owner = (video.get("owner") or {}).get("name") or ""
            title = video.get("title") or bvid
            duration = int(page.get("duration") or video.get("duration") or 0)
            candidates, hints, podcasts, recommended, reason, queries = search_plan(
                title, owner, description, duration, subtitle_kind
            )

            meta = {
                "platform": "bilibili",
                "video_id": outdir.name,
                "bvid": bvid,
                "aid": video.get("aid"),
                "cid": cid,
                "page": requested_page,
                "page_count": len(pages),
                "part_title": page.get("part"),
                "title": title,
                "channel": owner,
                "channel_url": f"https://space.bilibili.com/{(video.get('owner') or {}).get('mid')}" if (video.get("owner") or {}).get("mid") else None,
                "upload_date": dt.datetime.fromtimestamp(video.get("pubdate") or 0).strftime("%Y%m%d"),
                "duration": format_duration(duration),
                "duration_seconds": duration,
                "url": f"{referer}?p={requested_page}" if len(pages) > 1 else referer,
                "login": login,
                "subtitle_available": bool(selected),
                "subtitle_kind": subtitle_kind,
                "subtitle_lang": selected.get("lan") if selected else None,
                "subtitle_label": selected.get("lan_doc") if selected else None,
                "subtitle_why": why,
                "no_subtitle_reason": no_subtitle_reason,
                "subtitle_tracks": [
                    {key: value for key, value in track.items() if key != "subtitle_url"}
                    for track in tracks
                ],
                "metadata_only_track_count": sum(1 for track in raw_tracks if not track.get("subtitle_url")),
                "subtitle_endpoint": "/x/player/wbi/v2",
                "needs_translation": bool(selected and not selected["is_chinese"]),
                "description": description,
                "transcript_candidates": candidates,
                "transcript_hints": hints,
                "podcast_versions": podcasts,
                "search_recommended": recommended,
                "search_reason": reason,
                "search_queries": queries,
                "workdir": str(outdir),
            }

            if args.mode in {"subs", "doc"} and selected:
                raw_path, srt_path, transcript_path, cue_count, paragraph_count = write_subtitle_files(
                    ydl,
                    bvid,
                    requested_page,
                    selected,
                    outdir,
                    referer,
                    make_transcript=args.mode == "doc",
                )
                meta.update({
                    "subtitle_json_path": str(raw_path),
                    "subtitle_path": str(srt_path),
                    "subtitle_cue_count": cue_count,
                })
                if args.mode == "doc":
                    meta.update({
                        "transcript_path": str(transcript_path),
                        "transcript_paragraph_count": paragraph_count,
                    })

            save_meta(outdir, meta)
            if args.json:
                print(json.dumps(meta, ensure_ascii=False, indent=2))
            else:
                print(f"标题       {title}")
                print(f"UP主       {owner}　{meta['upload_date']}　{meta['duration']}")
                if len(pages) > 1:
                    print(f"分P        {requested_page}/{len(pages)}　{page.get('part') or ''}")
                print(f"登录态     {login['browser']}（有效，Cookie 仅驻留内存）")
                if selected:
                    print(f"选用轨     {why}　{selected.get('lan_doc') or selected.get('lan')}")
                    if args.mode in {"subs", "doc"}:
                        print(f"字幕       {meta['subtitle_path']}")
                    if args.mode == "doc":
                        print(f"文稿       {meta['transcript_path']}")
                else:
                    print(f"字幕       没有可用字幕：{no_subtitle_reason}")
                for item in podcasts:
                    print(f"播客版     {item['platform']}　{item['url']}")
                    print("           ↑ 播客平台分说话人，底本通常更好。先告诉用户，由用户决定是否改用播客版，不要自己转线。")
                if recommended:
                    print(f"现成文稿   必查：{reason}")
                    for candidate in candidates:
                        print(f"候选       {candidate['url']}")
                    for query in queries:
                        print(f"检索       {query}")
                else:
                    print("现成文稿   未发现直接线索；按内容类型可跳过主动检索")
                print(f"元数据     {outdir / 'meta.json'}")
                print(f"工作区     {outdir}")

            if not selected:
                die(
                    "没有可用字幕。请只按上面的候选和检索式检查优质现成文稿；"
                    "若仍未找到，立即停止，不下载音频、不做 ASR。",
                    code=2,
                )
    except BiliError as exc:
        die(str(exc), code=1)


if __name__ == "__main__":
    main()
