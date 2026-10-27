# 微信读书网页版：接口与正文格式

本文是 2026-09-04 实地侦察的结论，改代码前先看这里，能省掉重复踩坑。

## 一、阅读页地址怎么算

`/web/reader/<bookHash>k<chapterHash>`，两段哈希同一个算法（`lib/ids.py`）：

```
h = md5(id)
out = h[:3] + tag + '2' + h[-2:]         # 纯数字 id 用 tag='3'，否则 '4'
纯数字：按 9 位一组转十六进制，多组之间用 'g' 隔开
非数字：整串逐字符转十六进制
每组前面补两位长度（十六进制）
不足 20 位用 h 补齐，最后再接 md5(out)[:3]
```

已用书架上 7 本核对通过，含 `CB_` 开头的个人上传书。`chapterUid` 走同一个函数，
例如 uid=12 → `c20321001cc20ad4d76f5ae`。

## 二、JSON 接口（带 Cookie 直接 HTTP 就能打）

| 用途 | 方法 | 地址 |
| --- | --- | --- |
| 书架 | GET | `/web/shelf/sync?userVid=&synckey=0&lectureSynckey=0` |
| 书目 | GET | `/web/book/info?bookId=` |
| 章节表 | POST | `/web/book/chapterInfos`，体 `{"bookIds":[id],"synckeys":[0],"teenmode":0}` |
| 搜书 | GET | `/web/search/global?keyword=&maxIdx=0&count=20` |

`chapterInfos` 每章给的字段里，这几个有用：

- `chapterUid` / `chapterIdx`：前者用于算地址，后者是给人看的序号。
- `level`：原书目录层级，直接拿来定 Word 的大纲级别。
- `wordCount`：拿来和抓下来的字数对账，少于六成就该报警。
- `files`：原 EPUB 里的文件名，能看出这一章是由几个 xhtml 拼的。
- `tar`：`https://res.weread.qq.com/wrco/tar_<bookId>_<uid>`，是这一章**图片资源**的
  打包（成员形如 `epub_<bookId>_<n>`），**不含正文**。正文里的 `<img src>` 本来就是
  `https://res.weread.qq.com/wrepub/epub_<bookId>_<n>` 这样的绝对地址，直接下就行，
  所以现在没用到 tar。

## 三、正文为什么必须走浏览器

正文接口是 `POST /web/book/chapter/e_{i}`，请求体：

```json
{"b":"<bookHash>","c":"<chapterHash>","r":32729841,"st":0,
 "ct":1788520805,"ps":"...","pc":"...","sc":0,"s":"c77c8224"}
```

`s` 是客户端算的签名，`e_{i}` 里的 i 是请求序号不是章号。响应体是纯文本：

```
[32 位大写十六进制 MD5][一段 base64]
```

MD5 校验的是后面那段 base64 字符串本身（已验证四份样本全对）。坑在于：

1. **那段 base64 是整份文件 base64 之后被切出来的一片**，切点不按 4 字符对齐，
   所以单块解码要先猜 0–3 的偏移；`st=1` 的那次返回的是 CSS，其余是正文。
2. 分块拼接顺序不在响应里，得自己推。硬拼很容易错位——实测把「魏晋时**代**思想」
   解成了「魏晋时**c**思想」这种坏字，而且只坏零星几个字节，肉眼极难发现。

阅读器自己会把所有块拼好、整体 `window.atob` 一次。所以正确做法是挂钩 `atob`：

```js
const native = window.atob;
window.atob = function (s) {
  const raw = native(s);
  const bytes = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
  window.__wr.items.push(new TextDecoder('utf-8').decode(bytes));
  return raw;
};
```

一章通常正好两次调用：一次 CSS，一次完整正文 XHTML。拿到的是**原书 EPUB 的
XHTML 原文**，class 名、`<h1>`、`id` 锚点、`data-ratio` 全在，UTF-8 干净无坏字。

有的章平台根本不下发原文，给的是别章的 XHTML：《发展的自我》（3300018874）「版权声明」请求的是本章哈希，回来的却是「版权信息」，阅读器上显示的也是版权信息（2026-09-18 实测，重抓结果不变）。构建时与前面某章原文完全相同的章略去不排，报「平台缺章」。

正文渲染在 canvas 上（`.readerChapterContent` 里两个 `<canvas>`），所以读 DOM 文本
拿不到正文，别往那个方向走。

## 四、样式随章下发

每章的 `<style>` 里注入了原书 CSS（`p.bodytext`、`h3.text-title-3-c`、`.fnote`、
`.font1` …）。各出版社类名完全不同，硬编码类名换本书就废，所以 `lib/cssmap.py`
是读 CSS 求值，类名只在 CSS 缺失时兜底。

## 五、注释的几种写法

```html
<!-- 甲：注号与注文都有 id -->
<a href="#ch1" id="ch1-back"><sup>(1)</sup></a>
<p class="fnote"><a href="#ch1-back" id="ch1">(1)</a> 注文…</p>

<!-- 乙：只有 href，没有 id（同一本书里都可能混用） -->
<a href="#ch1"><sup>(1)</sup></a>
<p class="fnote"><a href="#ch1-back">(1)</a> 注文…</p>

<!-- 丙：微信读书自己插的夹注，注文在 alt 里 -->
<img class="qqreader-footnote" src="../Images/note.png" alt="注文全在这里。" />

<!-- 丁：一章由两个原文件拼成，注号在前一个、注文在后一个，锚点带文件名（《波斯笔记》第四章） -->
<a href="../Text/Chapter1-13.html#ch1" id="ch1-back"><sup>(1)</sup></a>          <!-- Chapter1-12.html -->
<p class="fnote"><a href="../Text/Chapter1-12.html#ch1-back">(1)</a> 注文…</p>   <!-- Chapter1-13.html -->

<!-- 戊：注号与注文分在两章，注文集中在「本章小结」（《亲密关系（第8版）》，36 条） -->
<span class="super"><a href="../Text/Section1_0009.xhtml#note_1" id="noteBack_1">[1]</a></span>
<p class="quotation"><a href="../Text/Section1_0002.xhtml#noteBack_1" id="note_1">[1]</a>注文…</p>
```

甲乙两式靠「id 配对 → 回跳 href 去掉 `-back` 再配 → 数量相等就按顺序配」三级兜底。
丁式在渲染前把指向本章自己文件（章节表 `files`）的锚点去掉文件名，就成了甲乙式。
戊式一章一章渲染时看不见对面，构建前全书扫一遍（`link_cross_file_notes`）：注号指向
注文锚点、锚点又回指注号 id，且两边文件名各自归属的章对得上，才算确证。
注文里若又夹着丙式夹注，Word 脚注不能嵌套，按括注接在原处。

## 六、图片

- 正文块图：`<img class="…image-alone80-withnote" data-w data-ratio data-w-new>`，
  类名里的 `alone50/80/100` 就是相对版心的宽度百分比，直接拿来定宽最准。
- 图说紧跟在图后，类名含 `picture-txt`，里面常有 `<br/>` + `span.font1` 放摄影署名。
- 生僻字是 `img.kindle-cn-inline-character`，必须按行内、与正文等高嵌入。
- `src` 是相对路径（如 `../Images/note.png`）的都是阅读器自带素材，不是书里的图。
- **表格多数在制书阶段就被拍成了图片**（十五个抽样章里一个真 `<table>` 都没有），
  但 CSS 里留着 `td` 规则，说明确有真表格的书。表格代码路径靠 `weread selftest`
  的夹具守着。

## 六之二、能不能直接下整本 EPUB（结论：不能）

书在服务端确实是 EPUB，但**没有整包下载的口子**。2026-09-04 实测：

| 探测 | 结果 |
| --- | --- |
| `res.weread.qq.com/wrepub/epub_<bookId>_<n>` | 只有图片。非图片编号返回 `errcode -2003` |
| `res.weread.qq.com/wrco/tar_<bookId>_<uid>` | 该章**图片**打包，不含正文 |
| `wrco/epub_<id>`、`wrco/txt_<id>_<uid>`、`wrbook/<id>`、`wrepub/<id>.epub` | 204 空 / 404 / -2003 |
| `/web/book/download`、`/web/book/epub`、`/web/book/bookDownload` | 404（返回站点首页 HTML） |
| 阅读器 4 个 JS 包（7.8 MB）里搜 `epub`/`download`/`tar_` | 没有整本下载路径 |

正文只从 `POST /web/book/chapter/e_*` 出，且要签名。所以「一次拿整本」不存在，
只能逐章取——这也是本 skill 必须按章跑浏览器的原因。

好在逐章取到的就是 EPUB 里那几个 xhtml 文件的原文（`chapterInfos.files` 会列出
它们的原始路径，如 `Text/Chapter_2_1.xhtml`），加上随章下发的原书 CSS 和
`res.weread.qq.com` 上的图片，等于把 EPUB 拆开一份份拿回来了。

## 六之三、相对路径图片：基址是 wrepub/web/<bookId>/

阅读器 JS 里有这么两条改写：

```js
html.replace(/\.\.\/video\/(.*?\.(mp4|wmv|3gp|rm|rmvb|mov|m4v|avi))/gi,
             'https://res.weread.qq.com/wrepub/web/' + X + '/$1')
    .replace(/\.\.\/images\/(.*?\.(png|jpg|jpeg|gif))/gi,
             'https://res.weread.qq.com/wrepub/web/' + X + '/$1')
```

`X` **就是 bookId**——抓阅读器实际请求实测到：

```
302  https://res.weread.qq.com/wrepub/web/CB_7DXA55A5DCnP74J75P3t88AW/note.png
```

注意路径是拍平的：`../Images/note.png` → `web/<bookId>/note.png`，中间目录丢掉，
只留文件名。这条只对确实上传了相对资源的书有效，普通出版书（如 22247793、
32608344）取 `note.png` 是 404。EPUB 与 docx 两侧都按这条规则解析相对图片，取不到就记进失败清单，不静默丢。

抽样五本书 72 章共 583 个相对路径 `<img>`，**全部是注释小图标**
（`qqreader-footnote` / `note.png`），真内容图一律是绝对地址。

## 六之四、原书嵌入字体：拿不到

有的书 CSS 里有 `@font-face`，`src` 指向 `../Fonts/汉仪仿宋B.ttf` 之类
（《东汉时代的政治与社会》有 11 个）。**这些字体文件服务端不提供，微信读书自己的
网页阅读器也根本不下载它们**——实测整页只下 KaTeX 的数学字体，正文用系统字体渲染。

**订正（同日实测）：字体其实能取到一部分。** 和相对图片同一个基址：

```
200  https://res.weread.qq.com/wrepub/web/<bookId>/汉仪仿宋B.ttf        773664B
200  https://res.weread.qq.com/wrepub/web/<bookId>/HYQTS.ttf            57884B
404  .../文鼎书苑宋…ext02.ttf     404  .../文鼎书苑宋…ext15.ttf
```

《东汉时代的政治与社会》9 个字体取到 4 个。EPUB 打包会把取到的嵌进 `Fonts/`，
取不到的整块 `@font-face` 删掉（留着就是坏引用，校验器会报）。

**但 PUA 依旧无解，两条原因：**

1. 服务端只放常用子集。取不到的几个偏偏叫 `ext02`/`ext15`——第 15 面就是私用区。
   用 fontTools 查过能取到的 4 个字体：字形数 177 / 84 / 34 / 1590，**PUA 数均为 0**。
2. 有的书连 `@font-face` 都没有。《治史三书》CSS 长 135093 字符、`font-family`
   出现 686 次，`@font-face` 为 **0**——字体全靠系统名（汉仪旗黑、汉仪楷体）。
   那 2 处 U+E862 就孤零零挂在正文里：「回纥入唐先取▮鹈泉」，即**䴙鹈泉**。

所以经典古籍库那套「读字形轮廓反查真字」的思路在这里用不上——那套的前提是
**手里有字形**（字体文件或录文图），要解的是码位被打乱；这里是反过来，码位没乱，
缺的是字形本身，没有可比对的东西。微信读书自己的阅读器也是空白，不是抓取的损失。
EPUB 打包会把这类字符连同上下文列出来供人工核定。

## 七、登录态

- `wr_skey` 是短时票据，几十分钟就过期，接口会返回 `errCode -2012 登录超时`
  （HTTP 状态仍是 200，只看状态码会漏判）。
- `/web/login/renewal` 只认浏览器当下那份 `wr_rt`，从其他浏览器抄来的快照续不动，
  会直接返回 `-2013 鉴权失败` 并把 Cookie 清空。
- 所以会话以 cloakbrowser 的 `weread` profile 为准（路径见 `lib/paths.py`）：打开一次首页让页面自己续，
  再把 Cookie 抄给 HTTP 接口用（`lib/session.py`）。首次可从本机 Zen/Chrome 等浏览器取种。
