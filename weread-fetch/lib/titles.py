#!/usr/bin/env python3
"""书名的清理规则：剥掉平台/出版方塞进题名的营销话术，留住真正的书目信息。

单独成模块是因为两处都要用同一套规则——文件名（lib/cli.py）和成品内页
（lib/docxbuild.py 的扉页与版权页）。放在 cli 里会让 docxbuild 反向依赖 CLI。
本模块只依赖 re，没有副作用。
"""

from __future__ import annotations

import re

_MARKETING = (r"(?:出品|赠|正版|包邮|全新|官方|旗舰|自营|独家|特价|限时|礼盒|随书|"
              r"附赠|畅销|典藏|珍藏|首发|热销|钜惠|重磅|力荐|推荐|倾情|诚意|揭秘|解密)")
RE_MARKETING = re.compile(_MARKETING)
# 平台/出品方的营销标签。只认带营销词的，不是见【】就删——【译丛067】那种是丛书
# 编号，属于书目信息要留着。
RE_TITLE_TAG = re.compile(r"[【\[［][^】\]］]*" + _MARKETING + r"[^】\]］]*[】\]］]")
# 圆括号既用来放营销话术（「智能：AI时代的商业、组织与战略的本质（曾鸣教授全新商业思考）」），
# 也用来放版次、丛书编号（「（增订版）」「（第8版）」「（译丛067）」），所以两头都得判。
RE_TITLE_PAREN = re.compile(r"[（(][^（()）]{1,40}[）)]")
RE_BIBLIO = re.compile(r"版|册|冊|卷|本|辑|輯|集|译|譯|注|校|编|編|第|丛书|叢書|"
                       r"修订|增訂|增订|影印|插图|插圖|精装|精裝|平装|平裝|简体|繁体|"
                       r"全集|文集|选集|選集")

# 书目字段里混着的非题名成分：装帧、册数、套装。这些是「这次印成几本、什么装帧」，
# 不是作品的题名（《三家〈诗〉辑佚史（精）全二册》）；微信读书又把几册合成了一本
# 电子书，册数更无从谈起。版次、丛书名不动——「增订版」「第8版」是有意义的版本信息。
_NUM = r"[一二两兩三四五六七八九十百\d]+"
# 「上下」「上中下」连着出现才是「整套几册」（《失控：…（上、下）》）。单独一个「（上）」
# 是只收了上册的那本书，删了会和下册撞成同一个文件名，所以不动。
_VOLUMES = r"[上中下](?:\s*[、·・/／,，]?\s*[上中下]){1,2}"
# 括在括号里的册数可以认得宽一些；不带括号的只认有「全/共/套装」领头的，
# 否则「某某第三册」会被削成「某某第」。
_PAREN_COUNT = (r"(?:套装\s*)?(?:(?:全|共)\s*)?" + _NUM + r"\s*[册冊]"
                r"|(?:套装\s*)?(?:全|共)\s*" + _NUM + r"\s*[卷本]"
                r"|(?:套装\s*)?" + _VOLUMES + r"\s*[册冊卷]?"
                r"|套装")
_BARE_COUNT = (r"(?:套装\s*)?(?:全|共)\s*" + _NUM + r"\s*[册冊卷本]"
               r"|套装(?:\s*" + _NUM + r"\s*[册冊])?"
               r"|上下[册冊]")
RE_TITLE_JUNK = re.compile(
    r"[（(](?:精|平|精装|精裝|平装|平裝|精装本|平装本|函套)[)）]"
    r"|[（(]\s*(?:" + _PAREN_COUNT + r")\s*[)）]"
    # 不带括号的那种，两边都不能挨着括号：「（插图珍藏版·全三册）」里的「全三册」
    # 删掉会留下半个括号，宁可不动，交给 title_needs_review 亮出来。
    r"|(?<![（(])(?:" + _BARE_COUNT + r")(?![)）])")
# 清完还带册数痕迹、又不是「单独一册」的括注：「（1-3册）」是全套还是其中三册，
# 光看题名判断不了。
RE_VOLUME_LEFT = re.compile(r"[册冊]|" + _VOLUMES + r"|"
                            + _NUM + r"\s*[-—–~～至]\s*" + _NUM + r"\s*卷")
RE_ONE_VOLUME = re.compile(r"第?\s*" + _NUM + r"\s*[册冊卷]|[上中下]\s*[册冊卷]?")
# 作者字段末尾的角色词。只处理单人署名（没有空格），多人合署原样不动
RE_AUTHOR_ROLE = re.compile(r"(编著|編著|主编|主編|校注|点校|點校|译注|譯注|注解|"
                            r"著|编|編|撰|辑|輯|译|譯|注)$")


def _strip_marketing_paren(text: str) -> str:
    """删掉纯营销话术的圆括号注，留下版次、丛书编号一类书目信息。

    「（曾鸣教授全新商业思考）」是平台加的宣传语，该删；
    「（增订版）」「（译丛067）」是题名的组成部分，一个字都不能动（册数另由
    RE_TITLE_JUNK 处理）。判据是「含营销词」且「不含书目词」。两边都保守：宁可漏删，
    让 title_needs_review 把可疑括号亮出来，也不误伤版次。
    """
    def repl(m: re.Match) -> str:
        inner = m.group(0)[1:-1]
        if RE_BIBLIO.search(inner) or not RE_MARKETING.search(inner):
            return m.group(0)
        return ""
    return RE_TITLE_PAREN.sub(repl, text)


def clean_title(raw: str) -> str:
    out = RE_TITLE_TAG.sub("", (raw or "").strip())
    out = _strip_marketing_paren(out)
    out = RE_TITLE_JUNK.sub("", out).strip(" 　·—-")
    return out or (raw or "").strip()


def title_needs_review(title: str) -> str:
    """清理之后仍留着可疑括号时提醒一句。

    营销标签平台随时会造新词，正则永远追不全；与其假装干净，不如把没认出来的
    括号亮出来，由跑这一趟的人（或 agent）决定要不要 --out 另起文件名。
    【】一类的括号一律亮出来；圆括号里的书目信息（增订版、译丛编号）太常见，只亮
    两种：带营销词的，和清完还带册数痕迹、又不是单独一册的。《失控》的「（上、下）」
    当初就是规则没认出、这里也没出声，才一路带进了文件名。
    """
    left = re.findall(r"[【\[［][^】\]］]{1,20}[】\]］]", title or "")
    for m in RE_TITLE_PAREN.finditer(title or ""):
        inner = m.group(0)[1:-1].strip()
        if RE_MARKETING.search(inner) or (RE_VOLUME_LEFT.search(inner)
                                          and not RE_ONE_VOLUME.fullmatch(inner)):
            left.append(m.group(0))
    return "、".join(left)


def clean_author(raw: str) -> str:
    author = (raw or "").strip()
    if not author or re.search(r"\s", author):
        return author          # 多人合署（如「汤用彤 汤一介等导读」）不碰
    out = RE_AUTHOR_ROLE.sub("", author).strip()
    return out if len(out) >= 2 else author
