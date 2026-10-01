# YouTube 文稿整理：实例、反例与缘由

SKILL.md 只留规则；本文是支撑这些规则的真实踩坑、判错实例和范文，拿不准怎么执行时对照。

## 为什么先跑 yt doc 再找现成稿

⓪ 编号在 ① 后面不是笔误：**先跑 `yt doc`**，它顺手把描述栏里的文稿链接挖出来给你，再据此决定要不要找、值不值得找。反过来先搜是浪费。实测 Lex Fridman、Dwarkesh、Lenny's 都在描述里明写 `Transcript: <url>`。

## 为什么绝不用 YouTube 的自动翻译字幕

`automatic_captions` 里除原生轨外的几百种语言，全是 YouTube 拿原生轨机翻的结果，等于在语音识别误差上再叠一层翻译误差。用户看到的"[music] 被当成句子成分翻进正文"就是这么来的。脚本只取原生轨，翻译交给 agent——至少能看见上下文。

## 现成稿分档的缘由

**B 档为什么不能当原话引述**：英文引述块的唯一作用是让读者核对**原话**。a16z 的文稿页自己写着"condensed and edited for readability"，Google 的 I/O keynote 文稿是"adapted to include more of what was announced"——都不是他说过的话。拿这种稿子填引述块，性质等同于 AI 重写，与"英文引述块 = 逐字照录"那条死规矩直接冲突。

**C 档同理**：第三方站多数就是把 YouTube 自动字幕换个皮重发，甚至更差（丢时间戳、Whisper 漏句），换过去等于白忙。

**付费墙绕行实例**：Lenny's 官网文章付费，描述栏却另给了一个公开 Dropbox「Transcripts Archive [public]」，按嘉宾名存 txt，全站免费。

## 专名误识实例

| 识别成 | 实际是 | 怎么发现的 |
|---|---|---|
| Lisa Doll | Lee Sedol（李世乭） | 讲围棋的视频里冒出没听过的女性名字 |
| Opening Eye | OpenAI | 读音接近，但语境里只可能是公司名 |
| AlphaChepan | AlphaChip | 前后都在讲 Alpha 系列，这个拼不出来 |
| gnome | GNoME | 讲材料科学时冒出"地精" |
| VO | Veo | 与 Nano Banana、Gemini 并列，只能是产品名 |
| The Big Orange Spot | The Big Orange Splot | 童书名，查得到正确拼写 |

同音词如 `in silicon` ← `in silico`、`fair play` 之类。受访者口误例：讲疟疾时说 `malaria virus`（疟疾由寄生虫引起，非病毒），保留原文加备注，不要替他改。

## 引述块被压缩的事故（2026-08-10）

有 agent 把每段英文压到原来的 1/4（14092 词 → 3233 词），读者以为那是原话，实际是 AI 重写的——**比不放英文更糟，因为它有欺骗性**。更严重的是丢了语义：原话 `There are no roles anymore. We haven't seen that.` 被压成 `There's been a lot written about role collapse.`，说话人对"岗位消失论"的**否认**没了，中文顺势写成"我们确实看到了……"，**把反驳变成了认同**。

## 备注的写法

能百分百确定的识别错误直接改正，不加备注：`Opening Eye` → OpenAI、`AlphaChepan` → AlphaChip、`Lisa Doll` → 李世乭，这类改动没有判断风险，加备注只是在向读者汇报工作过程。读者要的是一篇能读的文稿，不是校对日志——每个备注都是一次打断，得省着用。

脚注挂在名字上时别把名字又写一遍：
- 错：`嘉宾：Andrew Ambrosino（Andrew Ambrosino^[职位]）` → 出来是"Andrew Ambrosino（Andrew Ambrosino）"
- 对：`嘉宾：Andrew Ambrosino^[职位]`

脚注是挂在文字上的，不需要另起括号。

## 摘要范例

判断标准：把摘要单独拿给没看过视频的人，他应该能知道这期讲了哪些**具体主张**，而不只是知道"讲了什么话题"。

❌ 空洞（禁止这样写）
> 本期节目中，嘉宾与主持人围绕 AI 对产品团队的影响展开了深入探讨，分享了许多有价值的见解，并就未来发展趋势进行了展望。

✅ 有实质
> OpenAI 全公司 90% 的员工在用 Codex——不是 90% 的工程师，是全员。Ambrosino 认为产品流程被倒转了：实现不再是成本瓶颈，品味才是。他反对"岗位消失论"，但承认 Codex 组的角色融合程度高于公司其他部门，原因是这本身就是给工程师做的技术产品，设计师懂工程、PM 会写代码。他警告直接裁撤 PM 岗位的公司会掉进"区域防守"陷阱……

## 高亮与元信息块示例

```markdown
[00:00:50] **Andrew**：==实现不再是高昂的成本所在，真正的核心变成了品味==。
```

底本说明例：`底本：作者官网逐字文稿（lexfridman.com/demis-hassabis-2-transcript），已核对 YouTube 人工字幕，说话人标注沿用原稿。`
