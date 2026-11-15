# -*- coding: utf-8 -*-
"""
第二级探测：在已知结构上验证抽取规则，并把结果落盘成 probe2_result.txt。
重点确认三件事：
  1. 题干文本到底取不取得完整（h3 里 span 之后的 tail）
  2. 答案在哪：input#answerXXX 的 value？还是别处
  3. 有多少种题型、每种题型的块结构差异
"""

import re
import json
import urllib3
import requests
from lxml import etree
from probe_dowork import URL, HEAD, CURL_FILE, load_from_curl
import os

urllib3.disable_warnings()


def clean(s):
    if not s:
        return ""
    return re.sub(r"\s+", " ", s).strip()


def fetch():
    url, head = URL, dict(HEAD)
    if os.path.exists(CURL_FILE):
        ch, cu = load_from_curl(CURL_FILE)
        if cu:
            url = cu
        for k in ("Cookie", "User-Agent", "Referer"):
            if ch.get(k):
                head[k] = ch[k]
    r = requests.get(url, headers=head, verify=False, timeout=30)
    return r, etree.HTML(r.text)


def parse_question(q):
    """返回一个题目的结构化字典"""
    info = {"qid": q.get("id"), "classes": q.get("class")}
    h3s = q.xpath("./h3[contains(@class,'mark_name')]")
    h3 = h3s[0] if h3s else None
    info["h3_text"] = clean(h3.text) if h3 is not None else ""
    if h3 is not None:
        full = clean("".join(h3.itertext()))
        info["h3_full"] = full
        spans = h3.xpath("./span[contains(@class,'colorShallow')]")
        typ = clean("".join(spans[0].itertext())) if spans else ""
        info["type_label"] = typ
        stem = full
        if typ:
            stem = stem.replace(typ, "")
        stem = re.sub(r"^\s*\d+\s*[.、．]\s*", "", stem)
        info["stem"] = clean(stem)

    # 答案 input
    ins = q.xpath("./input[starts-with(@id,'answer') and not(starts-with(@id,'answertype'))]")
    info["answer_input_id"] = ins[0].get("id") if ins else None
    info["answer_value"] = ins[0].get("value") if ins else None
    tin = q.xpath("./input[starts-with(@id,'answertype')]")
    info["answertype_value"] = tin[0].get("value") if tin else None

    # 选项
    opts = []
    for bg in q.xpath("./div[contains(@class,'stem_answer')]/div[contains(@class,'answerBg')]"):
        letter = bg.xpath("./span[contains(@class,'num_option')]")
        body = bg.xpath("./div[contains(@class,'answer_p')]")
        opts.append({
            "letter": clean("".join(letter[0].itertext())) if letter else "",
            "text": clean("".join(body[0].itertext())) if body else clean("".join(bg.itertext())),
        })
    info["options"] = opts

    # 有没有答案展示区（已批改/已提交时的样式）
    for cls in ("mark_answer", "mark_fill", "answer_p", "mark_key"):
        if q.xpath(".//*[contains(@class,'%s')]" % cls):
            info.setdefault("has", []).append(cls)
    return info


def main():
    r, doc = fetch()
    out = []
    out.append("HTTP %s | 最终URL %s" % (r.status_code, r.url))

    tits = [clean("".join(x.itertext())) for x in doc.xpath("//h2[contains(@class,'type_tit')]")]
    out.append("\n=== 题型标题 type_tit (%d 个) ===" % len(tits))
    out.extend(tits)

    qs = doc.xpath("//div[contains(@class,'questionLi')]")
    out.append("\n=== questionLi 总数: %d ===" % len(qs))

    infos = [parse_question(q) for q in qs]

    out.append("\n=== 每题摘要: idx | 题型 | 答案值 | 选项数 | 题干前40字 ===")
    for i, inf in enumerate(infos, 1):
        out.append("%3d | %-14s | ans=%-6s | %d选项 | %s" % (
            i, inf["type_label"], str(inf["answer_value"])[:6], len(inf["options"]),
            inf["stem"][:40]))

    out.append("\n=== 前 3 题完整明细 (JSON) ===")
    out.append(json.dumps(infos[:3], ensure_ascii=False, indent=1))

    # 每种题型挑一个，导出原始 XML，便于核对嵌套
    out.append("\n=== 各题型样本原始 XML（截断 1800 字）===")
    seen = set()
    for q, inf in zip(qs, infos):
        key = inf["type_label"]
        if key in seen:
            continue
        seen.add(key)
        xml = etree.tostring(q, encoding="unicode", pretty_print=True)
        out.append("\n----- 题型: %s (qid=%s) -----" % (key, inf["qid"]))
        out.append(xml[:1800])

    text = "\n".join(out)
    with open("probe2_result.txt", "w", encoding="utf-8") as f:
        f.write(text)
    print(text[:3000])
    print("\n... 完整结果已写入 probe2_result.txt (%d 字符)" % len(text))


if __name__ == "__main__":
    main()
