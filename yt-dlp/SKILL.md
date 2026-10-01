---
name: yt-dlp
description: 把 YouTube、B站视频链接整理成文稿（docx）；开工先 yt check，再按平台读子技能。
---

# yt-dlp 套件

围绕 `yt-dlp` 的视频文稿能力，按平台拆成子技能，共享脚本都在本目录。支持子技能发现的 agent 可直接匹配 `youtube-transcript` 或 `bilibili-transcript`；其它 agent 按本文统一入口调用，不依赖某个单一 agent 平台的工作流。

**铁律：用户说"总结"也要出全文。** 仍按子技能完整流程走（探测来源 → 找现成稿 → 下载清洗 → 校错翻译 → 成稿），产出**开头 500 字摘要、正文完整逐字稿**的文稿。只有用户明确限定"只要一段摘要、不要全文"时才只写摘要，拿不准就出全文。

| 子技能 | 能做什么 | 状态 |
|---|---|---|
| [youtube-transcript](skills/youtube-transcript/SKILL.md) | YouTube 视频出文稿：下最优字幕 → 清洗 → 非中文翻译 → 渲染 docx | ✅ 可用 |
| [bilibili-transcript](skills/bilibili-transcript/SKILL.md) | B站：内存复用浏览器登录态，取人工/AI字幕，查优质现成稿，渲染 docx | ✅ 可用 |
| media-download | 下载视频/音频本体 | 🚧 未建 |

## 依赖检查（开工第一步）

接到任务后、下载之前先跑 `yt check`，缺什么它会打印安装命令，按提示装好再继续。

| 依赖 | 用途 | 安装 |
|---|---|---|
| Python 3.9+ | 全部脚本 | 系统 Python、Homebrew 或官网 |
| `yt-dlp` 命令 | 下载字幕 | `brew install yt-dlp` 或 `python3 -m pip install -U yt-dlp` |
| `yt_dlp` 库 | 仅 B站子技能 | `python3 -m pip install -U yt-dlp` |
| `python-docx` | 仅 `yt render` 出 docx | `python3 -m pip install -U python-docx` |
| 浏览器登录态 | YouTube 用 Chrome/Safari，B站用 Arc/Chrome | 在对应浏览器登录后重试 |

可选环境变量：

| 变量 | 作用 |
|---|---|
| `YT_TMP_ROOT` | 临时工作区，默认 `~/.cache/yt-transcript/`（Windows 为 `%LOCALAPPDATA%\yt-transcript`） |
| `YT_OUTPUT_DIR` | docx 成稿目录，默认当前目录下 `yt-transcript-output/` |
| `YTDLP_PATH` | 指定 yt-dlp 可执行文件 |
| `YT_COOKIES_FROM` | YouTube cookies 来源浏览器，默认 `chrome`，可设 `safari` / `none` |
| `BILI_BROWSER` | B站 cookies 来源浏览器，默认 `arc`，可设 `chrome`（非 macOS 必须设 chrome） |
| `BILI_BROWSER_PROFILE` | 浏览器 profile 路径或名称，不设用默认 profile |

## 统一入口

在本目录执行 `./yt <子命令>`，或用绝对路径 / `python3 <本目录>/yt <子命令>`；下文以 `yt` 指代。

| 命令 | 用途 |
|---|---|
| `yt check` | 检查依赖与运行环境 |
| `yt doc <url>` | 一步到位：探测 → 下最优字幕 → 洗成文稿 |
| `yt subs <url>` | 只下字幕（加 `--json` 输出结构化元信息） |
| `yt clean <file>` | 只清洗一个已有的 vtt/srt |
| `yt bili probe <url>` | B站只探测登录态、字幕轨和现成文稿线索 |
| `yt bili subs <url>` | B站下载最佳字幕 |
| `yt bili doc <url>` | B站下载最佳字幕并生成带时间戳底稿 |
| `yt render <md>` | 把整理好的成稿 Markdown 渲染为 docx，落成稿目录 |
| `yt rm <id>` / `yt rm --all` | 把临时工作目录移入废纸篓（可恢复） |

抓取与清洗是纯标准库实现，只有 `yt render` 需要 python-docx。yt-dlp 按「`YTDLP_PATH` → PATH → 常见安装路径」查找并校验版本，过旧直接报错并给升级命令。

## 共用约定

- **临时工作区** `<YT_TMP_ROOT>/<video_id>/`；成果写入目的地后务必 `yt rm <video_id>`。
- **成果落点** `<YT_OUTPUT_DIR>/<标题>.docx`，平铺不分子目录，归档交给用户自己做。
- **cookies**：YouTube **必须带 cookie**——匿名访问会间歇触发 `Sign in to confirm you're not a bot`（同机同 IP 有的视频能过、有的不能），带浏览器 cookie 则全部正常。所以默认从 Chrome 读登录态，**别改成匿名**。B站 cookie 只驻留内存，不落盘不打印。
- **大纲验收**：成稿用 `#` 至 `######` 表示章节，渲染成 Word 大纲 1–6 级，导航窗格可跳转；不能只用加粗或时间戳代替标题。抓取底稿不是已分节成稿；采用现成网页稿时应恢复原有章节，不把正文强调句批量升为标题。`yt render` 保存后会回读 docx 逐条核对标题文字与级别；还需对照源文确认小节覆盖，只有总标题不算通过。

## 扩展本套件

新增子技能＝在 `skills/<技能名>/SKILL.md` 写文件，并在上面的表里加一行。公共能力（yt-dlp 定位、临时目录、语言判定）在 `scripts/common.py`，Markdown→docx 在 `scripts/render_docx.py`，复用它们，别各写一套。目录结构：`skills/<技能名>/SKILL.md`（各平台子技能）、`scripts/`（共享脚本）、`tests/`（渲染器离线回归）、`yt`（跨平台统一入口）。
