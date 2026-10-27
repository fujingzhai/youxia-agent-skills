# youxia-agent-skills

可分享、可复用的 AI agent 技能合集。每个技能都是自包含目录：自带 `SKILL.md` 说明与全部脚本，复制即用，不绑定任何单一 agent 平台（Claude Code、Codex 等能读 `SKILL.md` 的 agent 都能用）。

## 目录

| 技能 | 说明 | 入口 | 自检 | 主要前提 |
|---|---|---|---|---|
| [cnki-fetch](cnki-fetch/SKILL.md) | 知网论文检索与全文下载：按篇名抓原版 PDF／CAJ，可另出带真·页下脚注的 DOCX；按作者、刊物、主题列文章清单 | `cnki` | `cnki --doctor` | 机构知网权限、国内网络出口 |
| [weread-fetch](weread-fetch/SKILL.md) | 把微信读书里有权阅读的书抓成 DOCX：还原图、表、注、字体层次与生僻字 | `weread` | `weread doctor` | 微信读书账号 |

两个技能都设为**手动触发**（`SKILL.md` 头部的 `disable-model-invocation: true`，Codex 对应 `agents/openai.yaml`）：在 Claude Code 里用 `/cnki-fetch`、`/weread-fetch` 调用，agent 不会自行选用。想让 agent 按需自动调用，删掉这两处设置即可。

## 安装（按技能单独安装，不必整库安装）

每个技能都是自包含目录。**你只需要其中一个技能时，只取那个文件夹即可，不需要 clone 整个仓库。**

分享单个技能给他人时，直接发对应的子目录链接，例如：

```text
https://github.com/fujingzhai/youxia-agent-skills/tree/main/cnki-fetch
```

对方如果是让 agent 安装，直接把下面这段话连同链接一起发给他的 agent 即可（`<技能名>` 换成 `cnki-fetch` 或 `weread-fetch`）：

```text
请安装这个 skill：https://github.com/fujingzhai/youxia-agent-skills/tree/main/<技能名>
只安装其中的 <技能名> 子目录（复制该文件夹到你的 skills 目录），不要安装仓库里的其他内容。
装好后按它 SKILL.md 的「前置条件与安装」建虚拟环境、装依赖，再运行自检命令，缺什么按输出补齐。
```

获取单个技能的方法（任选其一）：

- 在 GitHub 打开上面的子目录链接，点右上角 **Code → Download ZIP**，解压后只把里面对应的技能文件夹复制到自己的 skills 目录；
- 或者 `git clone https://github.com/fujingzhai/youxia-agent-skills.git` 后，只复制自己需要的那个子目录（整个仓库很轻量）。

复制出来的技能目录放到你的 agent 所识别的 skills 目录（不同 agent 目录不同，按其文档说明，常见的如 `~/.claude/skills/` 或项目内 `.claude/skills/`）。

## 依赖

完整步骤和自检命令都在各技能 `SKILL.md` 的「前置条件与安装」一节。agent 应主动执行自检，不要假设电脑上已装好任何工具。两个技能共同的前提：

- Python 3.10+，推荐用 [uv](https://github.com/astral-sh/uv) 在技能目录里建 `.venv`；
- [cloakbrowser](https://github.com/CloakHQ/CloakBrowser)（反检测 Chromium，随 `requirements.txt` 安装）：首次启动会自动下载约 350 MB 的浏览器内核到 `~/.cloakbrowser`；免费版同一时间只能开一个浏览器会话，**两个技能不要同时跑**；
- 系统：macOS 上开发与验证；Linux 应可用但未实测；Windows 不支持。

## 本机定制

- 输出目录等路径写在技能目录下的 `.env`（模板见各自的 `.env.example`），入口脚本会自动读取。
- 本机专属的约定（网络环境、配套工具、个人习惯）可以写进技能目录下的 `local.md`：`SKILL.md` 开头会让 agent 先读它，与正文冲突时以它为准。
- `.env`、登录态（`profiles/`、`.state/`）、缓存都已列入 `.gitignore`，不要提交或外传。

## 使用边界

只抓你的账号或所在机构有权访问的内容，不绕付费墙；成品仅供个人学习研究，遵守各平台的使用条款与著作权法。

## 添加新技能

在仓库根目录新建一个文件夹，例如 `my-skill/`，内含：

```text
my-skill/
├── SKILL.md       # 技能说明：用途、命令、前置条件与安装、自检命令
├── .env.example   # 可配置项（路径等）的模板
└── ...            # 技能需要的脚本与资源
```

保持每个技能自包含：不写死本机路径（用 `.env` 或环境变量），不引用系统外散落的配置，不提交登录态、缓存和个人清单。
