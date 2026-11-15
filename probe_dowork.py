# -*- coding: utf-8 -*-
"""
探测超星泛雅(mooc1) 作业页 dowork 的 HTML 结构。

两种填登录态的方式，任选其一：
  A) 在 D:\\xxt\\curl.txt 里粘贴浏览器「Copy as cURL (bash)」的整条内容（推荐，
     会自动取出 Cookie / User-Agent / Referer，还能顺带用 curl 里的真实 URL）
  B) 直接改下面 HEAD["Cookie"]，手填 Cookie 字符串

用法：
  D:\\xxt\\.venv\\Scripts\\python.exe probe_dowork.py
产出：
  dowork_debug.html      原始HTML
  dowork_skeleton.txt    精简结构树
"""

import os
import re
import urllib3
import requests
from lxml import etree

urllib3.disable_warnings()

URL = ("https://mooc1.chaoxing.com/mooc-ans/mooc2/work/dowork"
       "?courseId=267701950&classId=156254095&cpi=408927812"
       "&workId=57144512&answerId=56682801"
       "&standardEnc=664a0a6083ef2bd1e0ff8d344912aeea"
       "&enc=324a9c7860cdf49783e87df2d4ea8c07")

HEAD = {
    "Cookie": "",          # 方式B：把 _uid=...; fid=... 整串粘进来（不要带 "Cookie: " 前缀）
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/155.0.0.0 Safari/537.36 Edg/155.0.0.0"),
    "Referer": "https://mooc1.chaoxing.com/",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

SKIP = {"script", "style", "noscript", "svg", "iframe", "link", "meta"}
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CURL_FILE = os.path.join(BASE_DIR, "curl.txt")
# 探测产物统一放这里，免得把工作区根目录堆满
OUT_DIR = os.path.join(BASE_DIR, "页面样本")


def load_from_curl(path):
    """从 Copy as cURL 文本里解析 headers 和 url"""
    with open(path, encoding="utf-8", errors="ignore") as f:
        raw = f.read()
    raw = raw.replace("\\\n", " ").replace("^\n", " ").replace("`\n", " ")

    headers = {}
    for quote, body in re.findall(r"(?:-H|--header)\s+(['\"])(.*?)\1", raw):
        if ":" not in body:
            continue
        name, value = body.split(":", 1)
        headers[name.strip()] = value.strip()

    # 兼容直接把浏览器「Copy value」的原始 HTTP 请求报文贴进来的情况：
    #   GET /path HTTP/1.1
    #   Host: ...
    #   Cookie: ...
    if not headers.get("Cookie"):
        for line in raw.splitlines():
            m = re.match(r"^([A-Za-z0-9-]+):\s*(\S.*)$", line.strip())
            if m:
                headers[m.group(1)] = m.group(2).strip()

    url = None
    m = re.search(r"--url\s+(['\"])(.*?)\1", raw)
    if m:
        url = m.group(2).strip()
    else:
        m = re.search(r"curl(?:\.exe)?\s+(?:-\S+\s+)*?['\"](https?://[^'\"]+)['\"]", raw)
        if m:
            url = m.group(1).strip()
        else:
            m = re.search(r"(https?://mooc1\.chaoxing\.com/[^\s'\"]+)", raw)
            if m:
                url = m.group(1).strip()
            else:
                # 原始 HTTP 报文格式：请求行里是相对路径，需要配 Host 头拼回绝对 URL
                #   GET /mooc-ans/mooc2/work/view?... HTTP/1.1
                #   Host: mooc1.chaoxing.com
                req = re.search(r"^\s*(?:GET|POST)\s+(\S+)\s+HTTP/", raw, re.M)
                host = re.search(r"^\s*Host:\s*(\S+)", raw, re.M | re.I)
                if req and host:
                    path = req.group(1)
                    url = "https://%s%s" % (host.group(1).strip(),
                                            path if path.startswith("/") else "/" + path)

    return headers, url


def walk(node, depth, out, max_depth=16):
    if len(out) > 4000:
        return
    tag = node.tag
    if not isinstance(tag, str) or tag in SKIP:
        return
    cls = (node.get("class") or "").strip()
    nid = (node.get("id") or "").strip()
    line = "  " * depth + "<" + tag
    if nid:
        line += " id=" + nid
    if cls:
        line += " class=" + cls
    own = re.sub(r"\s+", " ", (node.text or "")).strip()
    if own:
        line += "  TEXT=" + own[:80]
    out.append(line)
    if depth >= max_depth:
        return
    for ch in node:
        walk(ch, depth + 1, out, max_depth)


def main():
    url, head = URL, dict(HEAD)

    if os.path.exists(CURL_FILE):
        ch, cu = load_from_curl(CURL_FILE)
        if cu:
            url = cu
        for k in ("Cookie", "User-Agent", "Referer"):
            if ch.get(k):
                head[k] = ch[k]
        print("已从 curl.txt 读取：URL用%s，Cookie长度=%d，UA=%s"
              % ("curl里的" if cu else "脚本内置", len(head.get("Cookie", "")),
                 head.get("User-Agent", "")[:40]))
    else:
        print("未发现 curl.txt，使用脚本内 HEAD（Cookie 长度=%d）"
              % len(head.get("Cookie", "")))

    if not head.get("Cookie"):
        print(">>> Cookie 是空的，八成会拿到登录页。先按 A 或 B 填好登录态。")

    os.makedirs(OUT_DIR, exist_ok=True)
    r = requests.get(url, headers=head, verify=False, timeout=30, allow_redirects=True)
    html = r.text
    print("HTTP", r.status_code, "| 长度", len(html), "| 最终URL", r.url)

    with open(os.path.join(OUT_DIR, "dowork_debug.html"), "w", encoding="utf-8") as f:
        f.write(html)

    doc = etree.HTML(html)
    title = " ".join(doc.xpath("//title/text()")) if doc is not None else ""
    print("title:", title.strip())
    if "passport2" in r.url or "登录" in title:
        print(">>> 仍是登录页：Cookie 没生效/已过期")

    for xp in ["//div[contains(@class,'questionLi')]",
               "//div[contains(@class,'mark_item')]",
               "//div[contains(@class,'TiMu')]",
               "//h3[contains(@class,'mark_name')]",
               "//ul[contains(@class,'mark_ulTop')]",
               "//div[contains(@class,'mark_answer')]",
               "//dl[contains(@class,'mark_fill')]",
               "//span[contains(@class,'colorShallow')]",
               "//div[contains(@class,'Zy_TItle')]",
               "//input[@type='radio']"]:
        print("%-55s -> %d" % (xp, len(doc.xpath(xp))))

    freq = {}
    for el in doc.iter():
        for c in (el.get("class") or "").split():
            freq[c] = freq.get(c, 0) + 1
    print("\n--- class 频次 TOP30 ---")
    for c, n in sorted(freq.items(), key=lambda kv: -kv[1])[:30]:
        print("%5d  %s" % (n, c))

    out = []
    body = doc.xpath("//body")
    walk(body[0] if body else doc, 0, out)
    with open(os.path.join(OUT_DIR, "dowork_skeleton.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print("\n已写出 页面样本\\dowork_debug.html / dowork_skeleton.txt（骨架 %d 行）" % len(out))


if __name__ == "__main__":
    main()
