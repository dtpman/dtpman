# 이 블록 전체를 scripts/verify_overflow.py 로 저장한다.
# -*- coding: utf-8 -*-
"""deck.pptx -> PDF 렌더 후, 문항 슬라이드의 '구분선~정답 배너 사이' 여백대에
텍스트가 삐져나왔는지(overflow) 자동으로 검사한다. generate_quiz_pptx.py의
레이아웃 상수(divider=4.70in, 정답 배너 시작=4.85in)와 반드시 맞춰서 쓸 것.
사용법: soffice --headless --convert-to pdf deck.pptx && python3 verify_overflow.py deck.pdf
"""
import sys
import subprocess
import glob
import os
from PIL import Image
import numpy as np

def main():
    pdf_path = sys.argv[1]
    outdir = pdf_path + ".qa_png"
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["pdftoppm", "-jpeg", "-r", "100", pdf_path, os.path.join(outdir, "p")], check=True)
    dpi = 100
    bad = []
    for f in sorted(glob.glob(os.path.join(outdir, "p-*.jpg"))):
        im = Image.open(f).convert("L")
        arr = np.array(im)
        y0, y1 = int(4.72 * dpi), int(4.82 * dpi)
        x0, x1 = int(0.5 * dpi), int(12.83 * dpi)
        strip = arr[y0:y1, x0:x1]
        dark = int((strip < 150).sum())
        if dark > 15:
            bad.append((f, dark))
    if bad:
        print(f"OVERFLOW 의심 슬라이드 {len(bad)}건:")
        for b in bad:
            print(" ", b)
        sys.exit(1)
    print("OK: overflow 없음")

if __name__ == "__main__":
    main()
