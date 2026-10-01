# youxia-agent-skills

可分享、可复用的 AI agent 技能合集。每个技能都是自包含目录：自带 `SKILL.md` 说明与全部脚本，复制即用，不绑定任何单一 agent 平台。

## 目录

| 技能 | 说明 |
|---|---|
| [yt-dlp](yt-dlp/SKILL.md) | YouTube / B站视频文稿整理：探测来源 → 找现成稿 → 下载最佳字幕 → 清洗校错 → 翻译 → 渲染 DOCX（真脚注、导航窗格大纲） |

## 安装（按技能单独安装，不必整库安装）

每个技能都是自包含目录。**你只需要其中一个技能时，只取那个文件夹即可，不需要 clone 整个仓库。**

分享单个技能给他人时，直接发对应的子目录链接，例如：

```text
https://github.com/fujingzhai/youxia-agent-skills/tree/main/yt-dlp
```

对方如果是让 agent 安装，直接把下面这段话连同链接一起发给他的 agent 即可：

```text
请安装这个 skill：https://github.com/fujingzhai/youxia-agent-skills/tree/main/yt-dlp
只安装其中的 yt-dlp 子目录（复制该文件夹到你的 skills 目录），不要安装仓库里的其他内容。
装好后先运行 yt check 检查依赖，缺什么按输出提示安装。
```

对方获取该单个技能的方法（任选其一）：

- 在 GitHub 打开上面的子目录链接，点右上角 **Code → Download ZIP**，解压后只把里面的 `yt-dlp/` 文件夹复制到自己的 skills 目录；
- 或者 `git clone https://github.com/fujingzhai/youxia-agent-skills.git` 后，只复制自己需要的那个子目录（整个仓库很轻量）。

复制出来的技能目录放到你的 agent 所识别的 skills 目录（不同 agent 目录不同，按其文档说明，常见的如 `~/.claude/skills/` 或项目内 `.claude/skills/`）。任何 agent 只要能读到该目录，按其中的 `SKILL.md` 调用脚本即可，不依赖特定平台。

## 依赖

每个技能所需的系统依赖都在各自 `SKILL.md` 的「依赖与环境检查」章节声明，
并带有自检命令。以 `yt-dlp` 为例：

- Python 3.9+
- `yt-dlp` 可执行文件（YouTube / B站字幕下载）
- `yt_dlp` Python 库（仅 B站子技能需要）
- `python-docx`（仅 `yt render` 生成 DOCX 需要）

安装后先运行 `yt check`，它会逐项检查并给出缺失项的安装命令。agent 应主动执行自检，
不要假设用户电脑上已装好任何工具。

## 添加新技能

在仓库根目录新建一个文件夹，例如 `my-skill/`，内含：

```text
my-skill/
├── SKILL.md      # 技能说明：用途、依赖、流程、自检命令
└── ...           # 技能需要的脚本与资源
```

提交并推送即可。保持每个技能自包含：不写死本机路径，不引用系统外散落的配置。
