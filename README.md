# chaoxing-work-export

把超星学习通（泛雅 / mooc1）的**作业查看页**导出成一份排好版的 Word 文档，题干、选项、正确答案、题型、分值一次性带走。

> ⚠️ 仅供**导出自己账号下、自己已提交的作业**用于个人复习。不要用它批量搬运他人内容、不要公开传播题库，也不要高频请求（脚本默认每页间隔 1 秒）。

---

## 特性

- **答案可靠**：从 `span.rightAnswerContent` 直取正确答案，另有两条兜底路径（隐藏 `input`、选项对错 class、`mark_answer` 文本切片）
- **选项不错位**：作答页的 `randomOptions=true` 会打乱选项字母，脚本用 `data` 属性把「原始字母 → 显示字母」映射回来
- **多题型兼容**：单选、多选、判断（答案是「对/错」中文）、填空（答案是 `(1) xxx`），各自归一化成统一格式
- **多页合并**：`URLS` 里写几条就导出几份，每份自动分页并加粗标题
- **中文排版**：宋体 + 五号 + `w:eastAsia`，避免中文回落成 Calibri
- **自带探测工具**：`probe_dowork.py` 能把页面存本地、数 XPath 命中、统计 class 频次、打印结构树 —— 站点改版时靠它重新定位
- **自带核验工具**：`check_docx.py` 把生成的 docx 读回来抽样打印，导出成功 ≠ 抓对了

## 目录结构

```
chaoxing-work-export/
├── dowork_export.py      ★ 主脚本：抓取 → 解析 → 写 docx
├── probe_dowork.py       一级探测：存 HTML / 数 XPath 命中 / class 频次 / 结构树
├── probe2.py             二级探测：逐题解析后落盘，看每道题解析成了什么
├── check_docx.py         核验：把 docx 读回来打印抽样
├── requirements.txt      依赖
├── examples/
│   └── curl.example.txt  登录态文件模板（照这个格式填）
└── legacy/               历史脚本，仅供对照，不参与运行
    ├── 111.py            最初的考试答案页脚本（DOM 假设与本项目不同）
    └── dowork_export_old_copy.py   早期版本（选项/答案处理较弱）
```

## 快速开始

```bash
# 1. 装依赖（建议用虚拟环境）
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

**2. 准备登录态** —— 新建 `curl.txt`（已被 `.gitignore` 忽略），内容二选一：

- **推荐**：浏览器已登录状态下打开那个作业页 → `F12` → **Network** → `Ctrl+R` → 点第一个文档请求 → 右键 **Copy → Copy as cURL (bash)** → 整条粘进 `curl.txt`。
- 或者按 `examples/curl.example.txt` 的格式，把原始 HTTP 请求报文（第一行 `GET /path HTTP/1.1`，接着逐行 `Name: value`）整段贴进去。

脚本会用正则从里面抠出 `Cookie` / `User-Agent` / `Referer`，两种格式都吃。

**3. 填你要导出的链接** —— 编辑 `dowork_export.py` 顶部 `URLS`：

```python
URLS = [
    "https://mooc1.chaoxing.com/mooc-ans/mooc2/work/view"
    "?courseId=...&classId=...&cpi=...&workId=...&answerId=...&enc=...",
]
```

必须是 **`/mooc2/work/view`**（已提交/已批改页），不能是 `/work/dowork`（作答页的答案是空的）。

**4. 跑起来**

```bash
python dowork_export.py     # 产出 作业题目导出.docx，原始 HTML 存到 页面样本/
python check_docx.py        # 核验：抽样打印，结果写到 check_out.txt
```

## 输出长什么样

```
作业 57144512（共 299 题）
一. 单选题（共299题，1913分）
1. [单选题] 从2025年10月31日奔赴星河到2026年5月29日返回地球……由以下哪三位航天员组成？
A、朱杨柱、张志远、黎家盈
B、张陆、武飞、张洪章
C、陈冬、刘洋、蔡旭哲
D、景海鹏、朱杨柱、桂海潮
答案：B（张陆、武飞、张洪章）
```

判断题会输出 `答案：A（对）`，填空题保留原文编号 `答案：(1) 开源`。

## 核心逻辑一览

| 步骤 | 做什么 | 关键代码 |
|---|---|---|
| 1 | 读登录态，拼 headers | `build_head()` / `load_from_curl()` |
| 2 | 请求并把原始 HTML 落盘（出问题时有原件可比对） | `main()` |
| 3 | 定位题目容器，多个候选 XPath 按顺序试 | `find_questions()` + `QUEST_XPATHS` |
| 4 | 解析单题：题干 / 选项 / 答案 | `parse_question()` |
| 5 | 选项多形态兼容（`ul.mark_letter>li` 或 `div.answerBg`） | `parse_options()` + `OPTION_XPATHS` |
| 6 | 答案三路兜底 + 归一化 | `extract_raw_answer()` / `normalize_answer()` |
| 7 | 写 Word（宋体、eastAsia、分页） | `setup_doc()` / `write_question()` |
| 8 | 回读核验 | `check_docx.py` |

站点改版时，通常只需要调 `QUEST_XPATHS` / `OPTION_XPATHS` —— 先用 `probe_dowork.py` 看新结构，再往候选列表里加一条。

## 常见问题

| 现象 | 原因 | 处理 |
|---|---|---|
| 被踢到 `passport2.chaoxing.com/login` | Cookie 过期，或链接不属于该 `_uid` | 重新 Copy as cURL 覆盖 `curl.txt` |
| `ModuleNotFoundError: No module named 'requests'` | 用的是系统 Python，不是虚拟环境 | 确认 `pip list` 里有依赖，或用解释器绝对路径 |
| `PermissionError: ...作业题目导出.docx` | 文档正被 Word / WPS 打开 | 脚本会自动改存带时间戳的文件；关掉文档再跑更省事 |
| 抓到题目但答案全空 | 用的是 `/work/dowork` 作答页 | 换成 `/work/view` 查看页 |
| `没抓到题目，页面结构不是 dowork/view 这两种已知模板` | 站点改版 | 看 `页面样本/debug_<workId>.html`，用 `probe_dowork.py` 重新探测 |
| 选项字母和答案对不上 | 忽略了 `randomOptions` 打乱 | 用选项的 `data` 属性（`orig`）映射回显示字母 |
| 控制台中文乱码 | Windows 控制台 GBK | 结果已写入 UTF-8 文件，直接看文件 |

## 安全提示

1. `curl.txt` 里的 Cookie **等同于账号密码**。不要提交、不要贴聊天、不要截图。`.gitignore` 已把它排除，但请自己再确认一次 `git status`。
2. 用完可以去超星**退出登录**，让 token 失效。
3. 仓库里的 `URLS` 已替换成 `<COURSE_ID>` 之类的占位符，填自己的链接时注意别把含课程/班级 ID 的完整链接提交上去。
4. 尊重站点使用条款与 `robots.txt`，控制请求频率，只取自己有权查看的内容。

## 参考文档

[学习通题库 - 打死不参团哼 - 博客园](https://www.cnblogs.com/Flin/articles/16996969.html)

## License

MIT