---
name: weread-fetch
disable-model-invocation: true
description: 把微信读书的书抓成 DOCX。
---

# 微信读书抓取

本目录若有 `local.md`（本机部署约定：路径、配套工具、存档），先读它；与本文冲突时以它为准。

在 cloakbrowser 里挂钩阅读器 `window.atob` 取服务端 EPUB 原文，按原样还原图、表、注、字体层次后重排成 DOCX。成品当完整的书，不写出处、不留平台声明。只抓当前账号有权读的，不绕付费墙、登录、验证码。

- 改抓取/转换代码前读 [references/protocol.md](references/protocol.md)；各类内容怎么落 Word、规则来由读 [references/output-spec.md](references/output-spec.md)。

## 交付

- 「抓这本书」＝ `weread fetch <书>`，只交 DOCX；用户没要 EPUB 不得加 `--format docx,epub`。
- 等 fetch 结束看它的退出码与输出，不要轮询成品文件名（文件名会经题名清理，轮询会永远等不到）。
- `fetch`/`build` 自动验收，报告在任务目录 `docx-audit.json`；非零退出＝失败，不覆盖旧成品，候选件留作诊断。**失败不能宣布完成。**
- 如实汇报：待核字（不凭上下文猜字当确证）、字体比例异常、字数偏少、平台缺章、注号没配上注文（注文会留在章末不成脚注），及三项结构告警：章节标题与章节表对不上（`heading_mismatches`）、标题后紧跟拆开重复的标题、正文长段仍是黑体（`hei_body_samples`，对照 CSS 判断漏转还是原书设计）。
- `weread check <docx>` 离线复查，不查标题与章节表对应。结构通过≠逐字、视觉、字形已核；没看过排版不说「排版已验证」。
- 排版核查：`soffice --headless --convert-to pdf` 转 PDF 再看。LibreOffice 用中华书局宋体渲染时书名号会显示成空的《》，是已知假象，docx 没丢字，不必另找渲染器求证。**无人值守时别让 Word/WPS 打开或导出**（macOS 沙盒弹授权窗、导出 PDF 走在线服务，会卡死任务）。
- 核查用的临时文件放任务目录，不放 `/tmp`，完事清掉。

## 命令

入口 `weread`（安装见文末）。成品默认 `~/Downloads/微信读书/<书名>_<作者>.docx`，任务目录默认技能目录下 `work/<bookId>-<书名>/`（可丢缓存，留着能 `build` 重生成）；都可在 `.env` 改，见 `.env.example`。

| 命令 | 说明 |
| --- | --- |
| `weread doctor` | 环境自检：依赖、浏览器、字体、登录态、路径 |
| `weread login [--qr]` | 续登录态；`--qr` 开窗扫码 |
| `weread shelf [词]` / `search <词>` | 列书架 / 搜书城 |
| `weread info <书>` | 章节表：序号、层级、字数 |
| `weread fetch <书>` | 抓正文＋生成 docx，可续跑 |
| `weread capture <书>` | 只抓不生成 |
| `weread build <书>` | 用缓存重生成 docx（改样式用） |
| `weread epub <书>` | 备用：打 EPUB 并结构自查 |
| `weread selftest` | 转换器自检，改完代码必跑 |

书可用书名（繁简都认，先书架后书城）或 bookId；同名多本会列出 bookId 让你挑。

参数：`--chapters 3-20|1,5,9`（分次抓，`build` 按全书顺序拼）、`--out <路径>`、`--toc`（默认不排目录，靠侧栏大纲）、`--size 11`、`--font 中华书局宋体`（默认；未装则退回宋体；楷体系时引文改仿宋）、`--variant simplified`（繁简双份的书改取简体，默认繁体）、`--no-page-break`、`--headed`（排查用）。

## 规则

- **私用区生僻字**：只用 `lib/pua.py` 查本机字体 cmap 等价还原；查不到原码位保留，只高亮该字，不连带整 run/整段；**绝不借字、绝不丢字**。人工判读记 `.state/pua_manual.json`。黄色高亮只用于此和图片失败的 `〔?〕`。其他缺字借备用字体并报数（`lib/fontcover.py <字体> "<文字>"`）；没装指定字体退回宋体并说明。
- **繁体底本**默认取繁体；接口给的书名作者照搬不转；简体残留只报告不改（清单是「疑似」）。
- **字体层次照原书 CSS**，唯一改写是正文整片黑体转宋体（标题黑体保留）。仿宋、楷体段落再多也不改，除非用户明确要求。
- **文件名**只去营销标签、装帧、合集册数、单人署名角色词；版次、丛书名与编号、单册册次保留。没把握报「题名待核」，宁可漏删。版权页只按同规则清「书名：」一行，平台发行声明滤掉并报数。
- 书架上 `MP_WXS_` 开头的是公众号，抓不了。
- 按章单线程、有间隔，**不改并发**。遇无权限、下架、验证码就停下报告。
- 登录态只在 `.state/cookies.json`（0600）与浏览器 profile（默认 `profiles/weread`），不进代码、不提交、不外传。

## 前置条件与安装

| 条件 | 说明 |
| --- | --- |
| 账号 | 微信读书账号；只能抓该账号有权阅读的书（免费、已购或会员可读） |
| 系统 | macOS 上开发与验证；Linux 应可用但未实测（字体目录只扫 macOS 路径）；Windows 不支持 |
| Python | 3.10+（推荐 3.12）。推荐用 [uv](https://github.com/astral-sh/uv)：环境装在本技能 `.venv`，包取自 uv 全局缓存，装过的不再下载，macOS／Linux 上以写时复制克隆，几乎不占额外磁盘 |
| 浏览器 | [cloakbrowser](https://github.com/CloakHQ/CloakBrowser)（随 requirements 安装）：反检测 Chromium，Playwright 接口。内核缓存在 `~/.cloakbrowser`，本机所有 cloakbrowser 共用：已有对应版本就直接沿用，没有才在首次启动时自动下载（约 350 MB）；免费版同时只能开一个会话，别与其他用它的任务并发 |
| 字体 | 默认正文「中华书局宋体」（按 Unicode 平面拆成几个字族的古籍字库，需自行安装），未装退回宋体；生僻字还原靠本机字体 cmap，装得越全还原越多 |
| 可选 | `opencc` 命令行（书名作者转繁）；LibreOffice（排版核查） |
| 使用边界 | 成品仅供个人学习研究，勿传播；遵守微信读书用户协议与著作权法 |

安装（在技能目录里执行）：

```bash
uv venv --python 3.12 .venv && uv pip install --python .venv/bin/python -r requirements.txt
#（没有 uv：python3 -m venv .venv && .venv/bin/pip install -r requirements.txt，依赖会另下一份）
ln -s "$PWD/weread" /usr/local/bin/weread    # 或 ~/.local/bin 等任一 PATH 目录；不链也可用 ./weread
cp .env.example .env                          # 可选：改成品与缓存目录
weread doctor                                 # 自检
weread selftest                               # 转换器自检，应全部通过
weread login                                  # 先从本机浏览器取登录态；不行就 weread login --qr 扫码
```

登录态取种顺序：浏览器 profile 自身 → 本机 Zen／Chrome／Edge／Brave／Firefox 里已登录的 weread.qq.com Cookie（经 rookiepy 读取，macOS 读 Chrome 系时可能弹钥匙串授权）→ `--qr` 开窗扫码。

## 排障

- `未登录`/`登录超时`：先 `weread login`，不行 `weread login --qr`。
- 某章失败（`progress.json` 里 `ok:false`）：重跑 `fetch` 只补失败章；连续失败用 `--headed` 看页面。
- `字数偏少`（不足接口 `wordCount` 六成）：多半没加载完，重跑该章。
- 调样式：改 `lib/docxbuild.py` → `weread build <书>` → `weread selftest`。
