# -*- coding: utf-8 -*-
"""量 PDF 里正文文本与表格线的真实坐标，验证页面版式（不依赖光栅预览）。"""
import re
import sys
import zlib

path = sys.argv[1]
data = open(path, "rb").read()
PT2MM = 25.4 / 72.0

mbox = re.findall(rb"/MediaBox\s*\[([^\]]*)\]", data)
for mb in mbox[:2]:
    vals = [float(v) for v in mb.split()]
    if len(vals) == 4:
        print("MediaBox: %.1f x %.1f pt = %.1f x %.1f mm"
              % (vals[2] - vals[0], vals[3] - vals[1],
                 (vals[2] - vals[0]) * PT2MM, (vals[3] - vals[1]) * PT2MM))

text_x = []
line_x = []
streams = 0
for m in re.finditer(rb"stream\r?\n", data):
    start = m.end()
    end = data.find(b"endstream", start)
    if end < 0:
        continue
    raw = data[start:end]
    try:
        s = zlib.decompress(raw)
    except Exception:
        continue
    streams += 1
    body = s.decode("latin-1")
    # 文本定位：BT ... ET 里的 Tm / Td
    for bt in re.finditer(r"BT(.*?)ET", body, re.S):
        block = bt.group(1)
        ctm = None
        for tok in re.finditer(r"([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm", block):
            ctm = (float(tok.group(5)), float(tok.group(6)))
        if ctm:
            text_x.append(ctm[0])
        for tok in re.finditer(r"([-\d.]+)\s+([-\d.]+)\s+Td", block):
            if ctm:
                text_x.append(ctm[0] + float(tok.group(1)))
    # 表格线：re（细长矩形）与 l（线段），取 x
    for tok in re.finditer(r"([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+re", body):
        x, y, w, h = (float(tok.group(i)) for i in range(1, 5))
        if h > 10 and w < 3:
            line_x.append(x)
        if w > 10 and h < 3:
            line_x.append(x)
    for tok in re.finditer(r"([-\d.]+)\s+([-\d.]+)\s+m\s+([-\d.]+)\s+([-\d.]+)\s+l", body):
        x1, y1, x2, y2 = (float(tok.group(i)) for i in range(1, 5))
        if abs(x1 - x2) < 0.5 and abs(y1 - y2) > 5:
            line_x.append(x1)

print("解压出内容流: %d 个" % streams)
if text_x:
    print("正文文本 x 坐标: 最小 %.1f pt = %.1f mm，最大 %.1f pt = %.1f mm（共 %d 处定位）"
          % (min(text_x), min(text_x) * PT2MM, max(text_x), max(text_x) * PT2MM, len(text_x)))
if line_x:
    uniq = sorted({round(v, 1) for v in line_x})
    print("竖线/横线 x 位置(pt):", uniq[:14])
    print("             换算(mm):", [round(v * PT2MM, 1) for v in uniq[:14]])
print("预期：左页边距 20mm = 56.7pt；正文右边界 190mm = 538.6pt")
