# -*- coding: utf-8 -*-
import io
from docx import Document
d = Document("作业题目导出.docx")
ps = [p.text for p in d.paragraphs]
out = io.StringIO()
out.write("段落总数 %d\n" % len(ps))
idx = [i for i, t in enumerate(ps) if t.startswith("作业 ")]
for i in idx:
    out.write("=" * 20 + " " + ps[i] + "\n")
    for t in ps[i + 1:i + 14]:
        out.write("  " + t[:70] + "\n")
# 判断题样本
for i, t in enumerate(ps):
    if "人工智能赋能传统产业转型升级" in t:
        out.write("=" * 20 + " 判断题样本\n")
        for x in ps[i:i + 4]:
            out.write("  " + x[:80] + "\n")
        break
open("check_out.txt", "w", encoding="utf-8").write(out.getvalue())
print("done")
