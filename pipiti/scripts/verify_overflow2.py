# -*- coding: utf-8 -*-
"""verify_overflow.py 확장판: 슬라이드마다 구분선 위치가 다르므로 layout.json을 읽어
(1) 구분선 바로 위 띠, (2) 해설 패널 아래(슬라이드 하단) 띠에 글자가 있는지 검사한다.
사용법: python3 verify_overflow2.py deck.pdf deck.pptx.layout.json"""
import sys, subprocess, glob, os, json
from PIL import Image
import numpy as np

pdf, layout = sys.argv[1], json.load(open(sys.argv[2], encoding="utf-8"))
outdir = pdf + ".qa_png"
os.makedirs(outdir, exist_ok=True)
for f in glob.glob(os.path.join(outdir, "*.jpg")):
    os.remove(f)
dpi = 100
subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), pdf, os.path.join(outdir, "p")], check=True)
files = sorted(glob.glob(os.path.join(outdir, "p-*.jpg")))
assert len(files) == len(layout), (len(files), len(layout))
bad = []
for f, lay in zip(files, layout):
    if "divider" not in lay:
        continue
    arr = np.array(Image.open(f).convert("L"))
    x0, x1 = int(0.5 * dpi), int(12.83 * dpi)
    checks = {
        "above_divider": (lay["divider"] - 0.10, lay["divider"] - 0.02),
        "below_divider": (lay["divider"] + 0.03, lay["divider"] + 0.12),
        "bottom": (7.33, 7.48),
    }
    for name, (a, b) in checks.items():
        strip = arr[int(a * dpi):int(b * dpi), x0:x1]
        dark = int((strip < 150).sum())
        if dark > 15:
            bad.append((os.path.basename(f), name, dark, lay.get("font")))
if bad:
    print(f"OVERFLOW 의심 {len(bad)}건:")
    for b in bad:
        print(" ", b)
    sys.exit(1)
print("OK: overflow 없음")
