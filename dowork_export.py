# -*- coding: utf-8 -*-
"""
泛雅(mooc1) 作业页 -> Word 导出
适用：
  .../mooc2/work/dowork?...   作答页（可能没答案）
  .../mooc2/work/view?...     已提交/已批改页（通常有答案）

特性：
  * 选项字母按 data 属性还原，randomOptions 打乱时不会错位
  * 答案三路兜底：隐藏 input -> 选项对错标记 -> mark_fill/mark_answer 文本
  * 支持多条 URL 连续导出，每份单独分页
  * 解析完把 HTML 存到 debug_<workId>.html，并打印首题 XML 片段（方便我改 XPath）
  * 宋体排版（含 eastAsia，避免中文回落成 Calibri）

用法：
  python dowork_export.py
登录态：从同目录 curl.txt 读取 Cookie / UA / Referer（参考 examples/curl.example.txt）
"""

import os
import re
import time
import traceback
import urllib3
import requests
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from probe_dowork import load_from_curl

urllib3.disable_warnings()

# ============ 配置区 ============
# 把你要导出的「查看页」链接整条粘到下面（浏览器地址栏直接复制即可）。
# 必须是 /mooc2/work/view（已提交/已批改，才有答案），/work/dowork 是作答页，答案为空。
# 原始链接形如：
#   https://mooc1.chaoxing.com/mooc-ans/mooc2/work/view
#     ?courseId=1234567&classId=1234567&cpi=12345678
#     &workId=57144512&answerId=56682801&enc=0123456789abcdef0123456789abcdef
# 换成几个就写几条，脚本会依次抓取、每份单独分页。
URLS = [
    # 作业 1
    "https://mooc1.chaoxing.com/mooc-ans/mooc2/work/view"
    "?courseId=<COURSE_ID>&classId=<CLASS_ID>&cpi=<CPI>"
    "&workId=<WORK_ID>&answerId=<ANSWER_ID>&enc=<ENC>",

    # 作业 2（不需要就删掉）
    # "https://mooc1.chaoxing.com/mooc-ans/mooc2/work/view"
    # "?courseId=<COURSE_ID>&classId=<CLASS_ID>&cpi=<CPI>"
    # "&workId=<WORK_ID>&answerId=<ANSWER_ID>&enc=<ENC>",
]
OUT_DOCX = "作业题目导出.docx"
SLEEP = 1.0          # 多个 URL 之间的间隔，别把站点惹毛
SHOW_SNIPPET = True  # 解析后打印首题 XML 片段（改 XPath 用），不要就改 False
# ================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CURL_FILE = os.path.join(BASE_DIR, "curl.txt")
HTML_DIR = os.path.join(BASE_DIR, "页面样本")   # 抓下来的原始 HTML 存这里

# 题目容器候选 XPath，按顺序试
QUEST_XPATHS = [
    "//div[contains(@class,'questionLi')]",
    "//div[contains(@class,'TiMu')]",
    "//div[contains(@class,'stem_answer')]/parent::div",
]

DEFAULT_HEAD = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/155.0.0.0 Safari/537.36 Edg/155.0.0.0"),
    "Referer": "https://mooc1.chaoxing.com/",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


def build_head():
    head = dict(DEFAULT_HEAD)
    if os.path.exists(CURL_FILE):
        ch, _ = load_from_curl(CURL_FILE)
        for k in ("Cookie", "User-Agent", "Referer"):
            if ch.get(k):
                head[k] = ch[k]
    if not head.get("Cookie"):
        print("!! 没有 Cookie，请求会落到登录页。先准备 curl.txt")
    return head


def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def text_of(node):
    return clean("".join(node.itertext())) if node is not None else ""


def _cls(node):
    return (node.get("class") or "").lower() if node is not None else ""


CORRECT_TOKENS = {"dui", "rightanswer", "right_answer", "correct", "right"}
WRONG_TOKENS = {"cuo", "wrong", "error"}


def is_correct_option(bg):
    """选项 div 是否被标成「正确」"""
    tokens = set(_cls(bg).split())
    if tokens & WRONG_TOKENS:
        return False
    if tokens & CORRECT_TOKENS:
        return True
    if "right" in _cls(bg) or "dui" in _cls(bg):
        return True
    # 选项里带图标/文字标记
    for xp in (".//*[contains(@class,'dui')]",
               ".//*[contains(@class,'right')]",
               ".//img[contains(@src,'dui')]",
               ".//img[contains(@src,'right')]"):
        if bg.xpath(xp):
            return True
    return False


OPTION_XPATHS = [
    # view 页：<ul class="mark_letter colorDeep qtDetail"><li>A. xxx</li>
    ".//ul[contains(@class,'mark_letter')]/li",
    "./div[contains(@class,'stem_answer')]//div[contains(@class,'answerBg')]",
    ".//div[contains(@class,'answerBg')]",
    ".//div[contains(@class,'answer')]//li",
    ".//ul[contains(@class,'answer')]/li",
    ".//div[contains(@class,'mark_choose')]//li",
    ".//div[contains(@class,'mark_choose')]//div[contains(@class,'clearfix')]",
    "./ul/li",
    ".//li",
    ".//p",
]

ANSWER_BOX_CLASS = ("mark_answer", "markanswer", "mark_fill")


def _in_answer_box(node):
    """节点是否位于「我的答案/正确答案」区块内（避免把答案当选项）"""
    for anc in node.iterancestors():
        cls = _cls(anc)
        if any(k in cls for k in ANSWER_BOX_CLASS):
            return True
    return False


def _leading_letter(text):
    m = re.match(r"^\s*([A-Ha-h])\s*[、.．)）:：]", text or "")
    return m.group(1).upper() if m else ""


def parse_options(q):
    """返回 [{show, orig, text, correct}]，多套 XPath 依次尝试"""
    for xp in OPTION_XPATHS:
        nodes = q.xpath(xp)
        if not nodes:
            continue
        opts = []
        for bg in nodes:
            if _in_answer_box(bg):
                continue
            sp = bg.xpath(".//span[contains(@class,'num_option')]")
            show = text_of(sp[0]) if sp else ""
            full = text_of(bg)
            if not show:
                # 找单独一个字母的短节点当标号
                for cand in bg.xpath(".//i | .//em | .//span | .//label | .//b"):
                    t = text_of(cand)
                    if re.fullmatch(r"[A-Ha-h]", t or ""):
                        show = t.upper()
                        break
            if not show:
                show = _leading_letter(full)
            text = full
            if show and text.startswith(show):
                text = clean(text[len(show):])
            text = re.sub(r"^[、.．)）:：]\s*", "", text)
            if not text:
                continue
            orig = (sp[0].get("data") or show) if sp else show
            opts.append({
                "show": show,
                "orig": clean(orig) or show,
                "text": text,
                "correct": is_correct_option(bg),
            })
        if opts:
            for i, o in enumerate(opts):
                if not o["show"]:
                    o["show"] = chr(65 + i)
                if not o["orig"]:
                    o["orig"] = o["show"]
            return opts
    return []


def dump_struct(node, depth=3, prefix="   "):
    """打印节点的标签/class/id/首段文本，用来定位选项到底藏在哪"""
    lines = []

    def walk(n, d):
        if d > depth:
            return
        tag = n.tag if isinstance(n.tag, str) else "?"
        cls = (n.get("class") or "").strip()
        ident = (n.get("id") or "").strip()
        own = clean("".join(t for t in n.xpath("./text()") if t))[:70]
        lines.append("%s<%s%s%s>%s%s" % (
            prefix + "  " * d, tag,
            (' class="%s"' % cls) if cls else "",
            (' id="%s"' % ident) if ident else "",
            (" " + own) if own else "", ""))
        for c in n:
            if isinstance(c.tag, str):
                walk(c, d + 1)

    walk(node, 0)
    return "\n".join(lines)


def _cut_correct(t):
    """从一段文本里截出「正确答案」后面的内容，丢掉我的答案/得分/解析"""
    tail = t.split("正确答案", 1)[1]
    tail = re.sub(r"^[\s:：]+", "", tail)
    for stop in ("我的答案", "答案解析", "答案分析", "AI讲解", "知识点", "点评"):
        if stop in tail:
            tail = tail.split(stop, 1)[0]
    tail = re.sub(r"[（(]\s*\d+(?:\.\d+)?\s*分\s*[)）]", " ", tail)
    tail = re.sub(r"(?<![\d.])0\s*分", " ", tail)
    return clean(tail)


def extract_correct_answer(q):
    """只返回「正确答案」之后的内容；找不到返回空串"""
    # 优先在答案容器里找写着「正确答案」的那个节点（而不是整个容器）
    nodes = q.xpath(".//*[contains(@class,'mark_answer') or "
                    "contains(@class,'markAnswer') or "
                    "contains(@class,'mark_fill')]")
    if not nodes:
        nodes = q.xpath(".//dl | .//dd | .//dt | .//p | .//span | .//td")
    for n in nodes:
        t = text_of(n)
        if "正确答案" in t:
            return _cut_correct(t)
    # 兜底：整题文本里找
    t = text_of(q)
    if "正确答案" in t:
        return _cut_correct(t)
    return ""


def extract_raw_answer(q):
    """页面上「正确答案」的原始值（未做格式化）"""
    # view 页：<span class="rightAnswerContent">B</span> / <dd class="rightAnswerContent">(1) 开源</dd>
    for n in q.xpath(".//*[contains(@class,'rightAnswerContent')]"):
        t = text_of(n)
        if t:
            return t
    # dowork 页：<input type="hidden" id="answer215650089" value="B">
    ins = q.xpath("./input[starts-with(@id,'answer') "
                  "and not(starts-with(@id,'answertype'))]")
    if ins:
        v = clean(ins[0].get("value"))
        if v:
            return v
    return ""


def normalize_answer(raw, options):
    """把原始答案整理成「A」/「AB」/「(1) 开源」这种可读形式"""
    t = clean((raw or "").replace("\xa0", " "))
    t = re.sub(r"[（(]\s*\d+(?:\.\d+)?\s*分\s*[)）]", " ", t)
    t = clean(t)

    # "B :张陆、武飞、张洪章;" -> "B"
    m = re.match(r"^([A-Ha-h][A-Ha-h\s,，、]*)\s*[:：]", t)
    if m:
        t = clean(m.group(1))

    by_orig, by_text = {}, {}
    for o in options:
        by_orig[(o["orig"] or o["show"]).upper()] = o["show"]
        by_text[clean(o["text"])] = o["show"]

    # 纯字母答案：原始字母 -> 页面显示字母
    if options and t and not re.search(r"[^\sA-Ha-h,，、]", t):
        letters = re.sub(r"[^A-Ha-h]", "", t).upper()
        out = "".join(by_orig.get(c, c) for c in letters)
        if out:
            return out

    if options and t:
        # 判断题：答案是「对/错」这类选项文字
        key = clean(t)
        if key in by_text:
            return by_text[key]
        if len(options) <= 2:
            for o in options:
                ot = clean(o["text"])
                if ot and ot in key:
                    return o["show"]

    return re.sub(r"\s*[;；]\s*$", "", t)


def parse_question(q, idx):
    """把一个题目 div 解析成字典"""
    item = {
        "idx": idx,
        "qid": q.get("id", ""),
        "type": q.get("typeName") or "",
        "no": "",
        "stem": "",
        "options": [],
        "answer": "",
        "has_answer": False,
    }

    # 题干
    heads = q.xpath("./h3[contains(@class,'mark_name')]")
    if not heads:
        heads = q.xpath("./h3") or q.xpath("./h2") or q.xpath(".//h3")
    if heads:
        h3 = heads[0]
        no = clean(h3.text)
        m = re.match(r"^\s*(\d+)\s*[.、．]?\s*$", no)
        item["no"] = m.group(1) if m else no
        full = text_of(h3)
        spans = h3.xpath("./span[contains(@class,'colorShallow')]")
        label = text_of(spans[0]) if spans else ""
        if not item["type"] and label:
            item["type"] = re.sub(r"[(（].*?[)）]", "", label).strip("()（） ")
        stem = full.replace(label, "") if label else full
        qt = h3.xpath(".//span[contains(@class,'qtContent')]")
        if qt:
            stem = text_of(qt[0])
        stem = re.sub(r"^\s*\d+\s*[.、．]\s*", "", stem)
        # 万一「我的答案」挤进了题干，砍掉
        stem = re.split(r"我的答案|正确答案", stem)[0]
        item["stem"] = clean(stem.replace("\xa0", " "))

    item["options"] = parse_options(q)

    # 答案 1：页面给的「正确答案」
    raw = extract_raw_answer(q)
    if raw:
        item["answer"] = normalize_answer(raw, item["options"])
        item["has_answer"] = True
        return item

    # 答案 2：选项上的对错标记
    flags = [o["show"] for o in item["options"] if o.get("correct")]
    if flags:
        item["answer"] = "".join(flags)
        item["has_answer"] = True
        return item

    # 答案 3：兜底，从「正确答案」文字后面截
    ans = extract_correct_answer(q)
    if ans:
        item["answer"] = normalize_answer(ans, item["options"])
        item["has_answer"] = True
    return item


def setup_doc():
    docum = Document()
    style = docum.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)
    rpr = style._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    return docum


def write_question(docum, it):
    no = it["no"] or str(it["idx"])
    stem = it["stem"] or "（题干解析失败）"
    head = "%s. %s" % (no, stem)
    if it["type"] and it["type"] not in head:
        head = "%s. [%s] %s" % (no, it["type"], stem)
    docum.add_paragraph(head)

    for o in it["options"]:
        docum.add_paragraph("%s、%s" % (o["show"], o["text"]))

    if it["has_answer"]:
        line = "答案：%s" % it["answer"]
        texts = []
        for ch in it["answer"]:
            hit = [o for o in it["options"] if o["show"] == ch]
            if hit:
                texts.append(hit[0]["text"])
        if texts:
            line += "（%s）" % "；".join(texts)
        docum.add_paragraph(line)


def get_work_id(url):
    m = re.search(r"[?&]workId=(\d+)", url)
    return m.group(1) if m else "unknown"


def find_questions(doc):
    for xp in QUEST_XPATHS:
        qs = doc.xpath(xp)
        if qs:
            return qs, xp
    return [], ""


def main():
    head = build_head()
    docum = setup_doc()
    os.makedirs(HTML_DIR, exist_ok=True)   # 抓下来的 HTML 统一放 页面样本\
    total, with_ans = 0, 0

    for i, url in enumerate(URLS, 1):
        work_id = get_work_id(url)
        print("\n[%d/%d] workId=%s 抓取 %s" % (i, len(URLS), work_id, url[:100]))
        try:
            r = requests.get(url, headers=head, verify=False, timeout=30,
                             allow_redirects=True)
            print("   状态 %s，最终地址 %.90s" % (r.status_code, r.url))
            if "passport2" in r.url:
                print("   !! 被踢到登录页：Cookie 过期或不属于该账号，跳过")
                continue

            html_file = os.path.join(HTML_DIR, "debug_%s.html" % work_id)
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(r.text)
            print("   原始 HTML 已存：%s" % html_file)

            doc = etree.HTML(r.text)
            if doc is None:
                print("   !! HTML 解析失败（返回的不是网页）")
                continue

            titles = [text_of(x) for x in
                      doc.xpath("//h2[contains(@class,'type_tit')]")]
            qs, used_xp = find_questions(doc)
            print("   题型 %d 个，题目 %d 道（用 %s）"
                  % (len(titles), len(qs), used_xp or "无匹配"))

            if not qs:
                print("   !! 没抓到题目，页面结构不是 dowork/view 这两种已知模板")
                continue

            parsed = [parse_question(q, k) for k, q in enumerate(qs, 1)]
            lost = [p for p in parsed if not p["options"]
                    and ("单选" in p["type"] or "多选" in p["type"]
                         or "判断" in p["type"])]
            print("   有选项 %d 题，没选项 %d 题（其中选择/判断题 %d 题）"
                  % (sum(1 for p in parsed if p["options"]),
                     sum(1 for p in parsed if not p["options"]),
                     len(lost)))

            if SHOW_SNIPPET:
                print("   ---- 第 1 题结构 ----")
                print(dump_struct(qs[0]))
                if lost:
                    bi = parsed.index(lost[0])
                    print("   ---- 第 %d 题（%s，没抓到选项）结构 ----"
                          % (lost[0]["idx"], lost[0]["type"] or "?"))
                    print(dump_struct(qs[bi]))
                print("   ---- 结构输出结束 ----")

            if total > 0:
                docum.add_page_break()
            p_head = docum.add_paragraph("作业 %s（共 %d 题）" % (work_id, len(qs)))
            if p_head.runs:
                p_head.runs[0].bold = True

            for t in titles:
                docum.add_paragraph(t)

            for it in parsed:
                write_question(docum, it)
                total += 1
                if it["has_answer"]:
                    with_ans += 1

        except Exception:
            print("   !! 这个 URL 处理时抛异常：")
            traceback.print_exc()
            continue

        time.sleep(SLEEP)

    out = os.path.join(BASE_DIR, OUT_DOCX)
    try:
        docum.save(out)
    except PermissionError:
        # 最常见的原因：这个 docx 正被 Word / WPS / 查看器打开，文件被锁住
        alt = out[:-5] + "_" + time.strftime("%H%M%S") + ".docx"
        docum.save(alt)
        print("\n!! 写 %s 失败（文件被占用，多半是你还开着它）" % OUT_DOCX)
        print("   已改存为：%s" % alt)
        out = alt
    print("\n完成：%d 题，其中带答案 %d 题" % (total, with_ans))
    print("已保存：%s" % out)
    if total == 0:
        print("提示：一题都没抓到，看上面提示的 HTML 文件名，比对该页结构。")
    elif with_ans == 0:
        print("提示：题目抓到了但都没答案 —— 该页没提供答案（多半还是作答页）。")


if __name__ == "__main__":
    main()
