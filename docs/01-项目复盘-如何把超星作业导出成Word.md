# 这个项目是怎么做出来的：把超星「作业」页导出成 Word

> 写给你看的：你是爬虫初学者，所以我不只给结论，还把**当时脑子里在想什么、为什么这么试、哪一次试错了**都写出来。
> 全文对应工作区 `D:\xxt`，所有文件名都是真实存在的（清理过一轮，见文末）。

---

## 0. 一句话总览

你最初问的是「`111.py` 是干什么的、怎么用」。我发现它是爬**考试答案页**的脚本，对你手上的**作业页**完全不适用，于是重写了一个新脚本，最终把三份作业（599 道题）连题干、选项、答案一起导出成 `作业题目导出.docx`。

整件事的骨架只有五步，后面的万言都是在这五步里的细节：

```
① 读懂现成脚本 → ② 确认必须登录 → ③ 探测页面结构 → ④ 写解析+导出 → ⑤ 回读核验
                                             ↑
                                    90% 的时间花在这里
```

**记住这句话**：爬虫的难点从来不是「发请求」，而是「在几百 KB 的 HTML 里，把你要的那几段稳定地切出来」。请求只有一行代码，切数据要几十行。

---

## 1. 先建立词汇表（不看这段后面会懵）

| 名词 | 说人话 | 在这个项目里的具体样子 |
|---|---|---|
| **HTTP 请求** | 浏览器向服务器要一个网页 | `requests.get(url, headers=..., verify=False)` |
| **URL 参数** | 问号后面 `a=1&b=2` 那些键值对 | `workId=57144512`（作业编号）、`enc=324a9c...`（服务端给的校验串） |
| **Cookie** | 服务器发给你、之后每次请求都带回去的一串「身份证」 | `_uid=123456789; fid=1234; p_auth_token=eyJhbGci...`（已脱敏） |
| **User-Agent（UA）** | 声明「我是哪个浏览器」 | `Mozilla/5.0 ... Edg/155.0.0.0` |
| **Referer** | 声明「我是从哪个页面点过来的」 | `https://mooc1.chaoxing.com/` |
| **登录态** | Cookie 有效 = 服务器认得出你是谁 | 少了它，一切请求都被踢到 `passport2.chaoxing.com/login` |
| **HTML DOM** | 网页在内存里的树：`div` 套 `div` 再套 `span` | 一道题 = 一个 `div.questionLi` |
| **选择器 / XPath** | 「在这棵树里定位某类节点」的语法 | `//div[contains(@class,'questionLi')]` |
| **lxml** | 把 HTML 解析成树、并支持 XPath 的库 | `from lxml import etree` → `etree.HTML(html)` |
| **python-docx** | 用代码生成 Word 的库 | `Document()` → `add_paragraph()` → `save()` |
| **分页（pagination）** | 一次拿不完，要按页拿 | 本项目是「一份作业一页」，不是列表分页 |
| **反爬** | 站点故意为难脚本：频率限制、加密参数、打乱选项 | 本站的 `randomOptions=true` 会打乱选项字母 |

---

## 2. 第 0 步：先读懂你手上的 `111.py`

初学者最常见的错误是「拿到一个脚本就改 URL 硬套」。所以第一步不是写代码，是**搞清楚这个脚本假设了什么**。

### 2.1 它到底在爬哪个页面

打开 `D:\xxt\111.py`，关键信息在 URL 里：

```
exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?...
                              ^^^^^^^^^^^^^^^^^^^^^^^
```

- `exm-` = exam（考试）
- `mayuan-ans` = 麦源 / 答案
- `look-detail` = 查看详情

所以它爬的是**「考试已结束后查看答案」的手机版页面**，参数是 `examId`、`examAnswerId`。

### 2.2 为什么它对你的作业页 100% 无效（三条证据）

| 证据 | `111.py` 的假设 | 你的作业页实际情况 |
|---|---|---|
| 页面模板 | `div[2]/div[3n-2]` 这种**按位置数出来的死路径** | 作业页每题是 `div.questionLi`，层级完全不同 |
| 选项数量 | 写死 `range(4)`（只抓 4 个选项） | 多选题也是 4 个，但判断题只有 2 个 → 判断题会抓空 |
| 答案来源 | 页面里直接有「正确答案：」这段文字 | 作业**作答页**里答案是空的（`<input id="answer..." value="">`） |

结论：**改 URL 没用，只能重写。** 这也是我当时的判断依据——不是「懒」，是结构假设全部不成立。

### 2.3 从它身上学到的 4 个有用招式（我确实沿用了）

1. **`verify=False` + `urllib3.disable_warnings()`**：有些站点的证书链不完整，忽略 SSL 报错。
2. **`docum.styles['Normal']` 里设 `eastAsia` 字体**：中文 Word 必须把「东亚字体」也设成宋体，否则中文会回落成 Calibri，看起来很难受。
3. **`answer[2:]` 砍前缀**的思路：答案前面常带「正确答案：」这类固定字，需要切掉——只是后来我省略了「按长度砍」这种脆做法，改成「找到标记再切」。
4. **先存 HTML 再解析**：出问题时有原件可比对。

### 2.4 它埋的坑（也是给你的提醒）

- **第 11 行是一段注释掉的 base64**。解码后才是 `Marxism1=[...]` 和 `xi2=[...]` 两个 URL 列表，而且 `xi2` 根本没被使用。→ *看别人的脚本，注释里常藏着真正要用的东西。*
- `head = {}` 是空的：不填 Cookie 必然拿到登录页。作者留了坑，你要会填。
- `range(4)` 写死选项数：换成「有多少抓多少」才通用。
- `/html/body/div[2]/div[3n-2]` 这种路径：页面一改版就全废。→ *永远优先用 `class`/`id` 做锚点。*

---

## 3. 第 1 步：先问「不登录能不能拿到」——最省事的验证

写代码前先花 30 秒验证，能省掉几小时白工。

我用工具直接请求了你给的作业链接（不带任何 Cookie），结果：

```
302 → passport2.chaoxing.com/login?...
```

**结论：必须带登录态。** 这一步的意义是定下大方向：

- 如果 302 到登录页 → 老老实实准备 Cookie（本项目就是这条）
- 如果 200 且内容对 → 直接抓，连 Cookie 都不需要
- 如果 200 但内容是个空壳（比如只有 `<div id="app"></div>`）→ 数据是 JS 动态加载的，要去 **F12 → Network → XHR** 里找 JSON 接口（见文档 `02`）

---

## 4. 第 2 步：把登录态变成文件（`curl.txt`）

### 4.1 怎么拿 Cookie（推荐「Copy as cURL」）

1. Chrome/Edge 里正常打开那个作业页（已登录状态）。
2. 按 `F12` → 切到 **Network（网络）** 面板 → 按 `Ctrl+R` 刷新。
3. 在列表里点第一个文档请求（名字通常是 `view?...` 或 `dowork?...`）。
4. 右键 → **Copy → Copy as cURL (bash)**。
5. 把复制到的一整行粘贴进 `D:\xxt\curl.txt`。

为什么不手工抄 Cookie？因为它有十几项、几百个字符，抄错一个字符就变成登录页。

### 4.2 为什么我写了个 `load_from_curl()` 兼容两种格式

Windows 上没有 bash，`Copy as cURL (bash)` 出来的东西你没法直接执行——所以我们只把它当**头信息的搬运工**，用正则把 `-H 'Name: value'` 抠出来：

```python
for quote, body in re.findall(r"(?:-H|--header)\s+(['\"])(.*?)\1", raw):
    name, value = body.split(":", 1)
    headers[name.strip()] = value.strip()
```

同时兼容**另一种格式**：有人会把浏览器 F12 里「Copy value」的**原始 HTTP 报文**（第一行 `GET /path HTTP/1.1`，后面逐行 `Name: value`）整段贴进来。所以又加了一层兜底：

```python
if not headers.get("Cookie"):                       # 没解析到 Cookie 就说明是报文格式
    for line in raw.splitlines():
        m = re.match(r"^([A-Za-z0-9-]+):\s*(\S.*)$", line.strip())
        if m:
            headers[m.group(1)] = m.group(2).strip()
```

> 你后来几次贴给我的就是第二种格式。这个兼容不是我聪明，是被真实情况逼出来的。

### 4.3 安全三条（重要）

1. **Cookie 等同于账号密码**。谁拿到它，谁就是你。
2. **不要贴到聊天、截图、GitHub**。我提醒过两次，还是收到了你的 Cookie 明文——最稳妥的做法是：只放在 `curl.txt` 里，用完删掉，并去超星**退出登录**让 token 失效。
3. 本项目特意从 `curl.txt` 读 Cookie，而不是硬编码进脚本，就是为了方便你「删文件即撤销」。

---

## 5. 第 3 步：探测页面结构（全项目最关键的一步）

### 5.1 为什么不能猜 XPath

初学者最容易犯的错：在浏览器里右键「检查」，看到 `<div class="questionLi">`，就直接写 `//div[@class='questionLi']`。

问题在于：

- 页面可能有 **299 道题**，class 里还混着别的（`marBom60 questionLi singleQuesId`）→ 必须用 `contains()` 而不是 `=`
- 不同题型（单选/多选/填空/判断）**DOM 不一样**
- 同一个站的不同页面（`dowork` 作答页 vs `view` 查看页）**DOM 也不一样**

所以必须先「探测」——把真实页面抓下来，用程序问它：**你到底长什么样？**

### 5.2 探测脚本的四个动作

`D:\xxt\probe_dowork.py`（168 行）就是干这个的，它只做四件事：

**① 存原件**

```python
with open(os.path.join(OUT_DIR, "dowork_debug.html"), "w", encoding="utf-8") as f:
    f.write(html)
```

之后所有疑问都可以在这份 HTML 里 `grep` 找答案，不用反复请求站点。

**② 候选 XPath 逐个数一数（最关键的一招）**

```python
for xp in ["//div[contains(@class,'questionLi')]",
           "//h3[contains(@class,'mark_name')]",
           "//ul[contains(@class,'mark_ulTop')]",
           "//div[contains(@class,'mark_answer')]",
           "//dl[contains(@class,'mark_fill')]",
           "//input[@type='radio']"]:
    print("%-55s -> %d" % (xp, len(doc.xpath(xp))))
```

这不是为了「找一个能用的」，而是**用命中数量反推结构**：

- 命中 **299** → 这就是「每题一个」的容器 ✅
- 命中 **0** → 这个词压根不在这个页面里 ❌
- 命中 **1196** → 是「每题 4 个」的东西（选项），可用来交叉验证题数

**③ 统计 class 频次 TOP30**

```python
freq = {}
for el in doc.iter():
    for c in (el.get("class") or "").split():
        freq[c] = freq.get(c, 0) + 1
```

真实运行时输出里 `questionLi` 出现 299 次、`mark_letter` 出现 299 次……**频次≈题数**的 class 就是好锚点。这招专治「我不知道该用什么 class 定位」。

**④ 打「结构树」（骨架）**

```python
def walk(node, depth, out, max_depth=16):
    ...
    line = "  " * depth + "<" + tag
    if nid: line += " id=" + nid
    if cls: line += " class=" + cls
    own = re.sub(r"\s+", " ", (node.text or "")).strip()
    if own: line += "  TEXT=" + own[:80]      # 顺手带出这段文字，方便认出「这是题干」
    out.append(line)
```

它会输出一份缩进文本：

```
<div class="marBom60 questionLi singleQuesId" id="question215650089"  TEXT=1.
  <div class="aiArea">
    <div class="aiAreaContent">
      <h3 class="mark_name colorDeep">  TEXT=1.
      <ul class="mark_letter colorDeep qtDetail">
      <div class="mark_answer">
```

**比在浏览器里点来点去快十倍**，而且能直接看到「哪些节点里真的有文字」。

> 小技巧：把 `max_depth` 限制在 16，并且 `len(out) > 4000` 就 `return`——否则 299 道题会把输出刷成几十万行。

### 5.3 第一次探测（`dowork` 作答页）发现了什么

| 发现 | 具体值 | 含义 |
|---|---|---|
| 题型标题 | `<h2 class="type_tit">一. 单选题（共299题，1913分）</h2>` | 只有 1 种题型，299 题 |
| 题目容器 | `div.padBom50.questionLi.fontLabel`，带 `typeName="单选题"`、`id="question215650089"`、`data="215650089"` | 锚点确定 |
| 题干 | `h3.mark_name...workTextWrap` 里，`h3.text` 是 `1.`，`span.colorShallow` 是 `(单选题, 5分)`，**正文是 span 之后的裸文本节点** | 取文字要小心，`h3.text` 只有序号 |
| 选项 | `div.clearfix.answerBg.workTextWrap` → `span.num_option.fl`（显示字母，带 `data`=原始字母）+ `div.fl.answer_p > p`（内容） | 有 `data` 就能还原顺序 |
| 答案 | `<input type="hidden" id="answer215650089" value="" />` **全空** | 这页没有答案 |
| 打乱 | `<input id="randomOptions" value="true"/>` | **选项被打乱了**：`data="B"` 的显示成 A、`data="C"` 的显示成 B…… |

**两个要命的结论：**

1. 「作答页」抓不到答案 → 必须等作业提交、批改后，换成 `view` 页（本项目后来就是这么走的）。
2. 选项被打乱 → 如果答案只给了字母（A/B/C/D），**必须用 `data` 映射回显示字母**，否则答案全错位。

> 这就是为什么脚本里选项字典要同时存 `show`（页面上显示的字母）和 `orig`（服务器原始字母）。

### 5.4 五分钟学会 XPath（够你改本项目）

| 写法 | 意思 |
|---|---|
| `//div` | 全文档任意位置的 `div` |
| `./div` | **当前节点下面一层**的 `div`（`.` 表示当前） |
| `//div[@id='x']` | id 等于 x |
| `//div[contains(@class,'questionLi')]` | class **包含** questionLi（class 有多个值，所以用 contains） |
| `//div[contains(@class,'a')][contains(@class,'b')]` | 同时包含 a 和 b（两个括号并列 = 并且） |
| `//ul/li` | ul 的直接子 li |
| `//h3/text()` | 取 h3 的直接文字 |
| `//h3//text()` | 取 h3 里所有后代文字 |
| `string(//h3)` | 把 h3 里所有文字拼成一个串 |

在 Python + lxml 里：

```python
nodes = doc.xpath("//div[contains(@class,'questionLi')]")   # 返回列表
node  = nodes[0]
node.get("id")            # 属性
node.xpath("./h3")[0]     # 继续往下找
text  = "".join(node.itertext())   # 节点内全部文字
```

---

## 6. 第 4 步：第一版脚本 + 三个真实翻车

### 6.1 翻车 ①：答案抓出来是「我的答案 / 0分 / (1) / 111」

第一版我用「找到答案区域的文字，从『正确答案』后面切」的思路。当时定位答案是：

```python
dls = q.xpath(".//dl[contains(@class,'mark_fill')]")   # ❌ 取第一个
```

结果填空题输出变成：

```
答案：我的答案：0分(1)111
```

**病因**：一道填空题有两个 `dl.mark_fill`——**「我的答案」在前，「正确答案」在后**。`xpath()` 返回列表，我取了 `[0]`，正好取到「我的答案」那块（里面还有我乱填的 `111` 和 `0分`）。

**修法**：不按顺序，按**内容**找——遍历所有候选节点，只认含「正确答案」字样的那一个，再从那个字后面切：

```python
def _cut_correct(t):
    for kw in ("我的答案", "答案解析", "AI讲解", "知识点", "点评"):
        t = t.split(kw)[0]                # 把尾巴全砍掉
    t = re.sub(r"[（(]\s*\d+(?:\.\d+)?\s*分\s*[)）]", " ", t)   # 去掉 (5分)
    t = re.sub(r"^\s*0?\s*分", " ", t)
    return clean(t)
```

> **教训**：XPath 返回的是**列表**。取 `[0]` 之前先问自己「一定只有一个吗？」拿不准就遍历 + 条件过滤。

### 6.2 翻车 ②：单选题、多选题的选项全丢了

你当时的原话是「填空题和判断题很好地爬取了，但单选题和多选题题目选项没了」。

**病因**：我第一版只认 `dowork` 页的选项结构 `div.clearfix.answerBg`。而 `view` 页的选项是：

```html
<ul class="mark_letter colorDeep qtDetail">
  <li tabindex="0" class="workTextWrap">A. 朱杨柱、张志远、黎家盈</li>
```

**完全没有** `answerBg`、没有 `span.num_option`、没有 `data` 属性——字母就是 `li` 里的行内文字 `"A. "`。

**修法**：把选项定位改成「**候选列表，按顺序试，谁先命中用谁**」：

```python
OPTION_XPATHS = [
    ".//ul[contains(@class,'mark_letter')]/li",                  # view 页（首选）
    "./div[contains(@class,'stem_answer')]//div[contains(@class,'answerBg')]",  # dowork 页
    ".//div[contains(@class,'answerBg')]",
    ".//div[contains(@class,'answer')]//li",
    ".//ul[contains(@class,'answer')]/li",
    ...
]
```

同时 `parse_options()` 要能**两种都吃**：

- 有 `span.num_option` → 字母从 `span` 取，`orig` 从 `data` 取
- 没有 → 用 `_leading_letter()` 从文本开头剥出 `A` / `A.` / `A、`，剩下的才是选项内容

> **教训**：写爬虫的选项/字段解析，默认要「多形态兼容」。同一个网站的不同页面，模板经常不一样。

### 6.3 翻车 ③：`AttributeError: 'NoneType' object has no attribute 'rFonts'`

设置宋体时我写了：

```python
docum.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')   # ❌
```

`111.py` 里这行能跑，是因为它构造文档后恰好有 `rPr`；这里 `rPr` 是 `None`，一访问就崩。

**修法**（python-docx 的正确姿势）：

```python
rpr = style._element.get_or_add_rPr()        # 没有就创建一个
rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")   # 没有 rFonts 也创建
```

> **教训**：`get_or_add_xxx()` 系列是 python-docx 的「安全访问器」，改样式时优先用它，别直接链式点属性。

### 6.4 从三次翻车得到的三个习惯

1. **先报数，再解析**：每页都打印「题型 N 个，题目 M 道，有选项 X 题，没选项 Y 题」。数字不对，立刻就知道哪层错了。
2. **出问题就打结构**：`dump_struct(qs[0])` 打印首题 3 层标签树。比瞎猜快。
3. **XPath 永远给候选**：`find_questions()` 遍历 `QUEST_XPATHS`，并打印这次用的是哪条——下次页面改版，日志会直接告诉你。

---

## 7. 第 5 步：换到 `view` 页（有答案的那个页面）

### 7.1 `dowork` 和 `view` 差在哪

| | `/mooc2/work/dowork`（作答页） | `/mooc2/work/view`（查看页） |
|---|---|---|
| 用途 | 你还没提交时答题用 | 提交/批改后回看 |
| 答案 | `<input id="answer...">` **为空** | `<span class="rightAnswerContent">B</span>` **有值** |
| 选项容器 | `div.answerBg` + `span.num_option[data]` | `ul.mark_letter > li`，字母在文字里 |
| 选项是否打乱 | 是（`randomOptions=true`） | 仍是页面上的顺序，但字母就是显示字母 |
| 题型标题 | `h2.type_tit` | 同 |
| 题目容器 | `div.questionLi`（带 `typeName`） | `div.questionLi.singleQuesId`（没有 `typeName`，题型要从 `(单选题, 5分)` 里取） |

> 一句话：**作答页给的是「空卷子」，查看页给的是「批改好的卷子」**。要答案就必须用后者。

### 7.2 真实 DOM 全解剖

一份 `view` 页里，一道单选题长这样（删掉了无关属性，保留结构）：

```html
<div class="marBom60 questionLi singleQuesId" id="question215650089" data="215650089">
  <div class="aiArea">
    <div class="aiAreaContent">
      <h3 class="mark_name colorDeep">1.
        <span class="colorShallow">(单选题, 5分)</span>
        <span class="qtContent workTextWrap">从2025年10月31日奔赴星河到2026年5月29日返回地球…由以下哪三位航天员组成？</span>
      </h3>
      <ul class="mark_letter colorDeep qtDetail">
        <li tabindex="0" class="workTextWrap">A. 朱杨柱、张志远、黎家盈</li>
        <li tabindex="0" class="workTextWrap">B. 张陆、武飞、张洪章</li>
        <li tabindex="0" class="workTextWrap">C. 陈冬、刘洋、蔡旭哲</li>
        <li tabindex="0" class="workTextWrap">D. 景海鹏、朱杨柱、桂海潮</li>
      </ul>
      <div class="mark_answer">
        <div class="mark_key clearfix">
          <span class="colorDeep marginRight40 fl">
            <i class="fontWeight custom-style">我的答案:</i><span class="stuAnswerContent">B</span>
          </span>
          <span class="element-invisible-hidden colorDeep marginRight40 fl">:张陆、武飞、张洪章; </span>
          <span class="colorGreen marginRight40 fl">
            <i>正确答案:</i><span class="rightAnswerContent workTextWrap">B</span>
          </span>
          <span class="element-invisible-hidden colorGreen marginRight40 fl">:张陆、武飞、张洪章; </span>
        </div>
        <div class="mark_score">
          <span class="marking_dui"></span>
          <div class="totalScore fr"><i class="custom-style">5</i>分</div>
        </div>
        <div class="analysisDiv">
          <div class="analysis"><i>答案解析：</i><span class="qtAnalysis workTextWrap">…</span></div>
        </div>
      </div>
    </div>
    <a class="aiAssistant">…</a>
  </div>
</div>
```

逐行读它，能得出 6 条对我们有用的事实：

1. 题目锚点 = `div.questionLi`（299 个）。
2. 题干在 `span.qtContent` 里；`h3.text` 只有 `1.`，`span.colorShallow` 是 `(单选题, 5分)`——**题型、分值只能从这里抠**。
3. 选项在 `ul.mark_letter > li`，字母是文字前缀 `"A. "`。
4. 正确答案在 `span.rightAnswerContent`（多选是 `ACD`，无分隔符）。
5. **「我的答案」和「正确答案」是兄弟节点，而且各有一份隐藏的「文字版答案」**（`element-invisible-hidden`，`:张陆、武飞、张洪章;`）。这解释了为什么早期版本会串味。
6. 解析在 `span.qtAnalysis`（本项目没导出它，但你要的话加一行就行）。

### 7.3 三种题型的差异（这是本项目最花时间的部分）

| | 单选题 | 多选题 | 判断题 | 填空题 |
|---|---|---|---|---|
| 选项容器 | `ul.mark_letter > li`（4 个） | 同左（4 个） | 同左（**2 个**：`A. 对` / `B. 错`） | **没有选项** |
| 正确答案形态 | `B` | `ACD` | **`对`**（中文！不是 `A`） | `(1) 开源`（原文，含编号） |
| 答案位置 | `span.rightAnswerContent` | 同左 | 同左 | `dl.mark_fill.colorGreen > dd.rightAnswerContent` |
| 额外坑 | — | 无分隔符，别按字母个数数 | 必须把「对」映射成 A | 没有选项就无法映射字母，只能原样输出 |

判断题的映射逻辑（`normalize_answer()` 里）：

```python
if len(options) <= 2:                    # 判断题只有 2 个选项
    for o in options:
        if o["text"] in key:             # "对" 命中 "对" → 返回显示字母 A
            return o["show"]
```

### 7.4 为什么「先探测、后写码」值得

`dowork` 页和 `view` 页的 DOM 差异，**只能靠抓下来看**，猜不出来。我第一版因为跳过了「对比两个页面」这一步，白写了一轮选项解析。

真实时间分配大概是：

- 发请求、搞 Cookie：10%
- 探测结构（存 HTML、数命中、打结构树、grep 找答案节点）：**60%**
- 写解析与导出：20%
- 核验与修 bug：10%

---

## 8. 第 6 步：答案提取的三次迭代

### v1：读隐藏 input 的 value

```python
ins = q.xpath("./input[starts-with(@id,'answer') and not(starts-with(@id,'answertype'))]")
raw_ans = clean(ins[0].get("value")) if ins else ""
```

在 `dowork` 页（未提交）里它是空的 → 失败。
注意那个 `not(starts-with(@id,'answertype'))`：页面上还有 `answertype215650089` 这种干扰 id，不排除就会读错。

### v2：从「正确答案」后面的文字切片

```python
def extract_correct_answer(q):
    for n in q.xpath(".//*[contains(@class,'mark_answer')]//*[contains(@class,'mark_fill')]"
                     "| .//dl[contains(@class,'mark_fill')]"):
        t = text_of(n)
        if "正确答案" in t:
            return _cut_correct(t)
    return ""
```

能用了，但会带尾巴：`B :张陆、武飞、张洪章; `。

### v3（定稿）：直取 `rightAnswerContent` + 专门做归一化

```python
def extract_raw_answer(q):
    # view 页：<span class="rightAnswerContent">B</span> / <dd class="rightAnswerContent">(1) 开源</dd>
    for n in q.xpath(".//*[contains(@class,'rightAnswerContent')]"):
        t = text_of(n)
        if t:
            return t
    # dowork 页：<input type="hidden" id="answer..." value="B">
    ...
```

拿到原始值后统一交给 `normalize_answer(raw, options)`：

```python
def normalize_answer(raw, options):
    t = clean((raw or "").replace("\xa0", " "))
    t = re.sub(r"[（(]\s*\d+(?:\.\d+)?\s*分\s*[)）]", " ", t)   # 砍 (5分)
    t = clean(t)

    m = re.match(r"^([A-Ha-h][A-Ha-h\s,，、]*)\s*[:：]", t)      # "B :张陆…" → "B"
    if m:
        t = clean(m.group(1))

    by_orig, by_text = {}, {}
    for o in options:
        by_orig[(o["orig"] or o["show"]).upper()] = o["show"]    # 原始字母 → 显示字母
        by_text[clean(o["text"])] = o["show"]                    # 选项文字 → 显示字母

    if options and t and not re.search(r"[^\sA-Ha-h,，、]", t):    # 纯字母
        letters = re.sub(r"[^A-Ha-h]", "", t).upper()
        out = "".join(by_orig.get(c, c) for c in letters)
        if out:
            return out

    if options and t:                                            # 判断题「对/错」
        key = clean(t)
        if key in by_text:
            return by_text[key]
        if len(options) <= 2:
            for o in options:
                if clean(o["text"]) and clean(o["text"]) in key:
                    return o["show"]

    return re.sub(r"\s*[;；]\s*$", "", t)                        # 填空题原样返回
```

**三个分支对应三种题型**：纯字母（单选/多选/判断题给字母时）→ 中文选项文字（判断题）→ 其他（填空题）。

`parse_question()` 里的优先级（三路兜底）：

```python
raw = extract_raw_answer(q)                       # ① 直取 rightAnswerContent
if raw: ... return item
flags = [o["show"] for o in options if o.get("correct")]   # ② 选项上的对错标记
if flags: ... return item
ans = extract_correct_answer(q)                   # ③ 文本切片兜底
```

> **教训**：答案的「获取」和「清洗」要分成两个函数。获取负责找地方，清洗负责去杂质——混在一起，改一处就崩另一处。

---

## 9. 第 7 步：写 Word 的那些坑

### 9.1 中文必须是宋体，而且要点到 `eastAsia`

```python
def setup_doc():
    docum = Document()
    style = docum.styles["Normal"]
    style.font.name = "宋体"            # 只设这个不够
    style.font.size = Pt(10.5)          # 五号
    rpr = style._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")   # 关键：东亚字体
    return docum
```

Word 的字体分「西文 (ascii)」和「东亚 (eastAsia)」两套。只设 `style.font.name` 时，中文会用主题默认字体（常见是 Calibri 回落），看起来不对。

### 9.2 排版细节

```python
head = "%s. %s" % (no, stem)                                  # "1. 题干…"
if it["type"] and it["type"] not in head:
    head = "%s. [%s] %s" % (no, it["type"], stem)             # 题干里没题型就补一个

for o in it["options"]:
    docum.add_paragraph("%s、%s" % (o["show"], o["text"]))    # "A、选项内容"

line = "答案：%s" % it["answer"]
texts = [o["text"] for o in it["options"] if o["show"] in it["answer"]]
if texts:
    line += "（%s）" % "；".join(texts)                        # "答案：B（张陆、武飞、张洪章）"
```

- 每份作业之间 `docum.add_page_break()`，并加粗一行 `作业 57144512（共 299 题）`，否则三份题会糊成一片。
- 答案末尾补选项文字，是给人看的——复习时不用回头对选项。

### 9.3 写文件被占用

`docum.save(out)` 在 docx 被 Word/WPS 打开时会抛 `PermissionError`。修法不是「让用户关掉」，而是**兜底改存**：

```python
try:
    docum.save(out)
except PermissionError:
    alt = out[:-5] + "_" + time.strftime("%H%M%S") + ".docx"
    docum.save(alt)
    print("\n!! 写 %s 失败（文件被占用，多半是你还开着它）\n   已改存为：%s" % (OUT_DOCX, alt))
```

---

## 10. 第 8 步：怎么证明「抓对了」

**导出成功 ≠ 抓对了。** 一定要回读核验。

`D:\xxt\check_docx.py`（20 行）做的事：用 python-docx 把生成的 docx 读回来，打印每份作业前 13 段 + 一道判断题样本：

```python
d = Document("作业题目导出.docx")
ps = [p.text for p in d.paragraphs]
idx = [i for i, t in enumerate(ps) if t.startswith("作业 ")]
for i in idx:
    print(ps[i])
    for t in ps[i+1:i+14]:
        print("  " + t[:70])
```

核验结果（`D:\xxt\check_out.txt` 全文）：

| 作业 | 题型 | 选项 | 答案 |
|---|---|---|---|
| 57144512 | 一. 单选题（299题, 1913分） | A、B、C、D 齐全 | `答案：B（张陆、武飞、张洪章）` |
| 57144559 | 一. 多选题（120题, 600分） | A、B、C、D 齐全 | `答案：ACD（要自觉践行以人民为中心…）` |
| 57144602 | 一. 填空题（90题）+ 二. 判断题（90题） | 判断题 `A、对` / `B、错` | 填空 `答案：(1) 开源`；判断 `答案：A（对）` |

- 段落总数 3062
- 判断题序号从 **91** 接着往下排（与网页一致）
- 曾经的污染 `我的答案：/0分/(1)/111` 已消失

**核验的三个层次**（建议每次都做）：

1. **数量对不对**：题目数 = 网页上的题数？（299 / 120 / 180）
2. **字段全不全**：每道选择题都有选项？每道题都有答案？
3. **抽 3~5 条人工对照**：挑最短、最长、最刁钻（判断题、填空题）各一条，跟网页比。

### 顺带解释你遇到过的两次「运行失败」

| 现象 | 真实原因 | 排除方法 |
|---|---|---|
| `ModuleNotFoundError: No module named 'requests'/'lxml'/'docx'` | 用了系统 Python，不是虚拟环境 | 必须 `D:\xxt\.venv\Scripts\python.exe dowork_export.py` |
| `PermissionError: [Errno 13] ...作业题目导出.docx` | 文档被 Word/WPS 占用 | 现在会自动改存带时间戳的文件 |
| `IndentationError` / `SyntaxError` | 恰好在改文件时运行了 | 重跑一次 |

同一份代码我在本机连续跑通多次，所以「运行失败」基本都在**环境**这一层，不在逻辑层。

---

## 11. 代码地图：每个文件/函数是干什么的

```
D:\xxt\
├── 111.py                     你原来的脚本（爬考试答案页，本项目未使用，留作对照）
├── dowork_export.py           ★主脚本 507 行：抓取 → 解析 → 写 docx
├── dowork_export_old_copy.py  旧版本（选项/答案解析更弱的那个，留作对照，见第 12 节说明）
├── probe_dowork.py            探测脚本：存 HTML / 数 XPath 命中 / class 频次 / 打结构树
│                             （dowork_export.py 复用了它的 load_from_curl()）
├── probe2.py                  二级探测：逐题解析后落盘，看每道题解析成了什么
├── check_docx.py              核验：把 docx 读回来打印抽样
├── check_out.txt              check_docx.py 的输出（核验证据）
├── curl.txt                   Cookie / UA / Referer（敏感！用完删）
├── 作业题目导出.docx           成果
├── 页面样本\                  抓下来的原始 HTML（debug_57144512/57144559/57144602.html）
└── .venv\                     Python 虚拟环境（requests / lxml / python-docx）
```

`dowork_export.py` 的函数表：

| 函数 | 职责 |
|---|---|
| `build_head()` | 从 `curl.txt` 读 Cookie/UA/Referer，缺 Cookie 时告警 |
| `load_from_curl(path)` | 解析 curl 文本（两种格式）→ (headers, url)，来自 probe_dowork.py |
| `clean(s)` | 压缩空白、去首尾空格 |
| `text_of(node)` | 取节点内全部文字并 clean |
| `_cls(node)` / `is_correct_option(bg)` | 判断选项是否被标成「正确」（class 含 dui/right/correct…） |
| `_leading_letter(text)` | 从 `"A. xxx"` 里剥出 `"A"` |
| `parse_options(q)` | 按 `OPTION_XPATHS` 找选项，产出 `[{show, orig, text, correct}]` |
| `_in_answer_box(node)` | 判断节点是否位于答案区（避免把答案文字当选项） |
| `dump_struct(node)` | 打印节点 3 层结构树（调试用） |
| `_cut_correct(t)` | 砍掉答案后面的杂质 |
| `extract_correct_answer(q)` | 从「正确答案」文字后切片（兜底路径） |
| `extract_raw_answer(q)` | 直取 `rightAnswerContent`，退化为隐藏 input |
| `normalize_answer(raw, options)` | 字母映射 / 判断题中文映射 / 填空题原样 |
| `parse_question(q, idx)` | 单题解析 → `{idx,qid,type,no,stem,options,answer,has_answer}` |
| `setup_doc()` | 建文档 + 宋体 + 10.5pt + eastAsia |
| `write_question(docum, it)` | 一题写 1 + N + 1 段（题干/选项/答案） |
| `get_work_id(url)` | 从 URL 抠 `workId`（用于命名和日志） |
| `find_questions(doc)` | 遍历 `QUEST_XPATHS`，返回 (题目列表, 命中的表达式) |
| `main()` | 遍历 URLS：请求 → 存 HTML → 解析 → 写 docx → 统计 → 保存 |

**几个可以单独摘出来复用的函数**：`clean` / `text_of` / `_leading_letter` / `dump_struct` / `load_from_curl`。以后爬别的站，这几个几乎不用改。

---

## 12. 踩坑总表（按踩到的顺序）

| # | 坑 | 现象 | 根因 | 修法 |
|---|---|---|---|---|
| 1 | 直接套用 111.py | 什么都抓不到 | 它是考试页脚本，DOM 完全不同 | 重写 |
| 2 | 不带 Cookie | 302 到 `passport2.chaoxing.com` | 需要登录态 | curl.txt + `load_from_curl` |
| 3 | curl 文本格式不对 | headers 解析成空 | Windows 没有 bash / 贴的是原始报文 | 正则兼容两种格式 |
| 4 | 猜 XPath | 命中 0 或命中几千 | 没探测结构 | 候选 XPath 计数 + class 频次 |
| 5 | 取 `dls[0]` | 答案是「我的答案：0分」 | 列表里有多个 `dl.mark_fill`，「我的答案」在前 | 遍历 + 认「正确答案」关键字 |
| 6 | 只认 `answerBg` | 单选/多选选项全空 | `dowork` 与 `view` 模板不同 | `ul.mark_letter > li` 优先 + 多候选 |
| 7 | 直接点 `rPr.rFonts` | `AttributeError: NoneType` | `rPr` 可能不存在 | `get_or_add_rPr()` |
| 8 | 选项字母错位 | 答案与选项对不上 | `randomOptions=true` 打乱 | 用 `data`(orig) 映射到 show |
| 9 | 题干为空 | 只剩 `1.` | `h3.text` 只有序号 | 取 `span.qtContent` |
| 10 | 判断题答案错 | 输出「对」而非 A | 答案给的是中文 | `by_text` 映射 |
| 11 | 写 docx 报权限错 | `PermissionError` | 文档被 Word 打开 | 改存带时间戳的文件 |
| 12 | 用错 Python | `ModuleNotFoundError` | 没走 `.venv` | 用绝对路径调解释器 |
| 13 | 控制台中文乱码 | 输出是 `锟斤拷` | Windows 控制台 GBK | 把结果写进 UTF-8 文件再读，别信控制台 |

---

## 13. 你可以自己重跑的步骤

```powershell
# 1) 进工作区
cd D:\xxt

# 2) 装依赖（只需一次，虚拟环境已建好，这一步其实不用跑了）
D:\xxt\.venv\Scripts\python.exe -m pip install requests lxml python-docx

# 3) 更新登录态：浏览器 F12 → Network → Copy as cURL → 整段粘进 curl.txt

# 4) 跑主脚本（约 3 秒/页 + 1 秒间隔）
D:\xxt\.venv\Scripts\python.exe dowork_export.py

# 5) 核验导出的 docx
D:\xxt\.venv\Scripts\python.exe check_docx.py

# 6) 想单独探测某个页面的结构时
D:\xxt\.venv\Scripts\python.exe probe_dowork.py
#   产物在 .\页面样本\dowork_debug.html 与 dowork_skeleton.txt
```

**换一批作业怎么改**：只改 `dowork_export.py` 顶部 `URLS` 列表（三条 `.../work/view?...` 链接），别动别的地方。

**常见报错对照**：

| 报错 | 意思 |
|---|---|
| `被踢到登录页：Cookie 过期或不属于该账号` | 换 Cookie；或该链接属于别的 `_uid` |
| `没抓到题目，页面结构不是 dowork/view 这两种已知模板` | 看 `页面样本\debug_<workId>.html`，重新探测结构 |
| `题目抓到了但都没答案` | 还是作答页，换成 `/work/view` |
| `没选项 N 题（其中选择/判断题 M 题）` M>0 | 选项 XPath 要加候选 |

---

## 14. 三个小练习（从易到难）

1. **改输出**：在 `write_question()` 里把「答案解析」也导出来（`span.qtAnalysis`）。提示：解析答案要放进 `parse_question()` 的返回字典。
2. **加题型**：如果某份作业出现「简答题」，它会没有选项、答案是长文本。你猜现在的代码会输出什么？怎么改才好看？（提示：`has_answer` 已有；问题是长文本挤在一行。）
3. **换站点**：找一个你常看的、不需要登录的网页（比如某个新闻列表），用 `probe_dowork.py` 的三招探测结构，把标题+链接抓出来存成 Excel。这是最典型的爬虫入门题，见文档 `02`。

---

## 附：两条方法论（比代码更值钱）

1. **「先看见，再动手」**：把页面存成本地文件、把结构打成文本、把命中数打出来。这三个动作能消灭 80% 的瞎猜。
2. **「每一步都留证据」**：报数（多少题）、打样本（首题结构）、存原件（debug html）、回读核验（docx）。出问题时，这些证据直接指出是哪一层坏了；没有证据，就只能重头再来。
