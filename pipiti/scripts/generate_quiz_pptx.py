# 이 블록 전체를 scripts/generate_quiz_pptx.py 로 저장한다.
# -*- coding: utf-8 -*-
"""
문제 -> (클릭) -> 정답 -> (클릭) -> 정답 이유  순서로 공개되는 퀴즈형 PPTX 생성기.

입력: items.json  { "book": str, "deck_title": str, "items": [ {...}, ... ] }
  item 필드:
    id           문항 번호/식별자 (표시용)
    label        슬라이드 상단 배지에 쓸 라벨 (없으면 id 사용)
    passage      지문/대화 원문 (없으면 None) - 문항 위에 회색 박스로 표시
    prompt       발문 (필수)
    choices      선택지 리스트 (객관식이면 문자열 리스트, 아니면 None)
    kind         "choice" | "tf" | "text"
    answer       정답 텍스트 (그대로 표시)
    explanation  정답 이유/해설 텍스트

출력: <output>.pptx  (파워포인트 네이티브 클릭 애니메이션 내장, 별도 매크로/동영상 없음)

사용법:
    python3 generate_quiz_pptx.py items.json output.pptx
"""
import json
import sys
import zipfile
import shutil
import os
from lxml import etree

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------
# 색상 팔레트
# ----------------------------------------------------------------------
NAVY = RGBColor(0x1E, 0x2A, 0x4A)
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY_BG = RGBColor(0xF2, 0xF3, 0xF5)
GRAY_TXT = RGBColor(0x55, 0x55, 0x55)
ANSWER_BG = RGBColor(0xFF, 0xE9, 0x8A)
ANSWER_BORDER = RGBColor(0xE0, 0xB4, 0x00)
ANSWER_TXT = RGBColor(0x5A, 0x42, 0x00)
EXPL_BG = RGBColor(0xE6, 0xF0, 0xFB)
EXPL_BORDER = RGBColor(0x6C, 0x8E, 0xBF)
EXPL_TXT = RGBColor(0x15, 0x2A, 0x4A)
ACCENT = RGBColor(0xC0, 0x3A, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT_KR = "맑은 고딕"


# ----------------------------------------------------------------------
# 글자수 기반 폰트 크기 추정 (라이브러리에 실제 레이아웃 엔진이 없으므로
# 대략적 추정 후 word_wrap + auto-fit 플래그로 보정한다)
# ----------------------------------------------------------------------
def _hangul_frac(s):
    if not s:
        return 0.0
    n = sum(1 for ch in s if "가" <= ch <= "힣")
    return n / max(1, len(s))


def _em_per_char(s):
    """실측 보정치: 순수 영문 ~0.65em/자, 순수 한글(단어 사이 공백 포함) ~1.2em/자.
    (LibreOffice + 맑은 고딕 대체 폰트, 150dpi PDF 렌더 실측 기준, 12.333in 폭 박스)"""
    frac = _hangul_frac(s)
    return frac * 1.20 + (1 - frac) * 0.65


def estimate_lines(text, pt, box_w_in, safety=1.10):
    """word_wrap 텍스트 한 문단이 차지할 줄 수를 추정."""
    text = text if text else " "
    em = _em_per_char(text) * safety
    char_w_in = em * pt / 72.0
    chars_per_line = max(1, int(box_w_in / char_w_in))
    return max(1, -(-len(text) // chars_per_line))  # ceil div


def fit_font_size(text_blocks, box_w_in, box_h_in, max_pt=20, min_pt=10, line_gap=1.26):
    """text_blocks: list[str] (각 문단, 실제 렌더 폰트가 모두 동일하다고 가정할 때의 추정용).
    대략적인 줄 수를 추정해 box_h를 넘지 않는 가장 큰 폰트를 고른다. 실측 보정치 기반이며
    안전 마진(safety factor)을 포함하므로 약간 작게 잡히는 쪽으로 치우친다(넘치는 것보다 안전)."""
    for pt in range(max_pt, min_pt - 1, -1):
        n_lines = sum(estimate_lines(b, pt, box_w_in) for b in text_blocks)
        line_h_in = pt * line_gap / 72.0
        if n_lines * line_h_in <= box_h_in:
            return pt
    return min_pt


def add_textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    return box


import re as _re
_UL = _re.compile(r"\[\[(.*?)\]\]")


def set_paragraph(p, text, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT, font=FONT_KR):
    """[[...]] 로 감싼 부분은 밑줄(원문 밑줄 친 부분)로 그린다."""
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.alignment = align
    p.line_spacing = 1.05
    p.space_before = Pt(0)
    pos = 0
    parts = []
    for m in _UL.finditer(text):
        if m.start() > pos:
            parts.append((text[pos:m.start()], False))
        parts.append((m.group(1), True))
        pos = m.end()
    if pos < len(text) or not parts:
        parts.append((text[pos:], False))
    for chunk, ul in parts:
        if chunk == "":
            continue
        run = p.add_run()
        run.text = chunk
        run.font.size = Pt(size)
        run.font.bold = bold or ul
        run.font.italic = italic
        run.font.underline = ul
        run.font.color.rgb = ACCENT if ul else color
        run.font.name = font


def plain(text):
    return _UL.sub(lambda m: m.group(1), text)


def fill_multiline(tf, lines, size, color, bold=False, italic=False, space_after=2):
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        set_paragraph(p, line if line != "" else " ", size, color, bold=bold, italic=italic)
        p.space_after = Pt(space_after)


def rounded_box(slide, x, y, w, h, fill_color, line_color, line_w=Pt(1.25)):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.adjustments[0] = 0.06
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    shp.line.color.rgb = line_color
    shp.line.width = line_w
    shp.shadow.inherit = False
    return shp


def set_shape_name(shape, name):
    shape._element.nvSpPr.cNvPr.set("name", name)


# ----------------------------------------------------------------------
# 슬라이드 한 장 생성
# ----------------------------------------------------------------------
LAYOUT_LOG = []  # 슬라이드별 구분선 y(in) 기록 -> verify 단계에서 사용


def add_table(slide, rows, x, y, w, font_pt=14):
    n_r, n_c = len(rows), max(len(r) for r in rows)
    row_h = Inches(0.42)
    gt = slide.shapes.add_table(n_r, n_c, x, y, w, row_h * n_r)
    tbl = gt.table
    for ri, row in enumerate(rows):
        for ci in range(n_c):
            cell = tbl.cell(ri, ci)
            txt = row[ci] if ci < len(row) else ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if ri == 0 else (GRAY_BG if ci == 0 else WHITE)
            tf = cell.text_frame
            tf.word_wrap = True
            cell.margin_left = cell.margin_right = Inches(0.05)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            set_paragraph(p, txt, font_pt, WHITE if ri == 0 else DARK, bold=(ri == 0 or ci == 0),
                          align=PP_ALIGN.CENTER)
    return gt


def build_item_slide(prs, item, index, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # 배경
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE
    bg.line.fill.background()
    bg.shadow.inherit = False
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)

    margin = Inches(0.5)
    content_w = SLIDE_W - margin * 2

    # ---- 상단 배지 ----
    badge = rounded_box(slide, margin, Inches(0.32), Inches(3.2), Inches(0.5), NAVY, NAVY)
    btf = badge.text_frame
    btf.word_wrap = True
    btf.margin_left = Inches(0.1)
    btf.margin_right = Inches(0.1)
    p = btf.paragraphs[0]
    set_paragraph(p, f"{item.get('label', '문항 ' + str(item['id']))}", 15, WHITE, bold=True)
    p.alignment = PP_ALIGN.CENTER
    btf.vertical_anchor = MSO_ANCHOR.MIDDLE

    if item.get("topic"):
        tb = add_textbox(slide, margin + Inches(3.35), Inches(0.32), Inches(6.5), Inches(0.5), anchor=MSO_ANCHOR.MIDDLE)
        set_paragraph(tb.text_frame.paragraphs[0], item["topic"], 13, GRAY_TXT)

    counter = add_textbox(slide, SLIDE_W - margin - Inches(2.2), Inches(0.32), Inches(2.2), Inches(0.5),
                           anchor=MSO_ANCHOR.MIDDLE)
    p = counter.text_frame.paragraphs[0]
    set_paragraph(p, f"{index} / {total}", 13, GRAY_TXT, align=PP_ALIGN.RIGHT)

    line_gap = 1.26
    full_w_in = 12.333

    # ---- 해설 패널 높이 먼저 결정 (긴 해설이면 패널을 키우고 위 영역을 줄인다) ----
    expl_lines = item["explanation"].split("\n")
    expl_inner_w = full_w_in - 0.5
    def expl_need(pt):
        n = sum(estimate_lines(plain(("정답 이유   " if i == 0 else "") + l), pt, expl_inner_w) for i, l in enumerate(expl_lines))
        return n * pt * line_gap / 72.0 + len(expl_lines) * 2 / 72.0 + 0.25
    expl_pt = 15
    expl_h_in = 1.35
    for cand in range(15, 9, -1):
        expl_pt = cand
        if expl_need(cand) <= 1.35:
            break
    else:
        expl_pt = 11
        expl_h_in = min(2.6, max(1.35, expl_need(11)))
        if expl_need(11) > 2.6:
            for cand in range(11, 8, -1):
                expl_pt = cand
                if expl_need(cand) <= 2.6:
                    break
            expl_h_in = 2.6
    bottom = 7.30
    expl_top = bottom - expl_h_in
    ans_h_in = 0.75
    ans_top = expl_top - 0.12 - ans_h_in
    divider_y = ans_top - 0.15

    # ---- 컨텐츠 영역 ----
    content_top_in = 0.95
    content_h_in = divider_y - 0.05 - content_top_in
    content_top = Inches(content_top_in)

    side = item.get("table") or item.get("image")
    content_w_in = 6.9 if side else full_w_in
    text_w = Inches(content_w_in)

    passage = item.get("passage")
    passage_lines = passage.split("\n") if passage else []
    prompt_lines = item["prompt"].split("\n")
    choice_lines = []
    for c in (item.get("choices") or []):
        choice_lines.extend(c.split("\n"))

    def required_height_in(pt):
        passage_pt = pt
        prompt_pt = pt + 1
        h = 0.0
        for line in prompt_lines:
            h += estimate_lines(plain(line), prompt_pt, content_w_in) * prompt_pt * line_gap / 72.0 + 3 / 72.0
        if passage_lines:
            h += 6 / 72.0 * 1.3
            for line in passage_lines:
                h += estimate_lines(plain(line), passage_pt, content_w_in - 0.3) * passage_pt * line_gap / 72.0 + 2 / 72.0
        if choice_lines:
            h += 8 / 72.0 * 1.3
            for c in choice_lines:
                h += estimate_lines(plain(c), pt, content_w_in) * pt * line_gap / 72.0 + 3 / 72.0
        return h

    usable_h_in = content_h_in - 0.20
    MAX_PT = 20

    def best_pt(fn):
        for cand in range(MAX_PT, 8, -1):
            if fn(cand) <= usable_h_in:
                return cand
        return 9

    max_choice = max((len(plain(c)) for c in choice_lines), default=0)
    RIGHT_W = 3.2 if max_choice <= 26 else 4.85
    LEFT_W = full_w_in - RIGHT_W - 0.15

    def split_height(pt, left_w=LEFT_W, right_w=RIGHT_W - 0.2):
        prompt_pt = pt + 1
        h = sum(estimate_lines(plain(l), prompt_pt, full_w_in) * prompt_pt * line_gap / 72.0 + 3 / 72.0 for l in prompt_lines)
        hp = sum(estimate_lines(plain(l), pt, left_w - 0.3) * pt * line_gap / 72.0 + 2 / 72.0 for l in passage_lines)
        hc = sum(estimate_lines(plain(c), pt, right_w) * pt * line_gap / 72.0 + 3 / 72.0 for c in choice_lines)
        return h + 0.15 + max(hp, hc)

    font_pt = best_pt(required_height_in)
    split = False
    if not side and passage_lines and choice_lines and font_pt < 18:
        sp = best_pt(split_height)
        if sp > font_pt + 1:
            split, font_pt = True, sp
    item["_font_pt"] = font_pt

    def write_prompt(tf):
        first = True
        for line in prompt_lines:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            set_paragraph(p, line, font_pt + 1, NAVY, bold=True)
            p.space_after = Pt(3)

    def write_passage(tf, first):
        for line in passage_lines:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            set_paragraph(p, line if line.strip() else " ", font_pt, DARK)
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(0.25))))
            p.space_after = Pt(2)

    def write_choices(tf, first):
        for c in item["choices"]:
            for k, cl in enumerate(c.split("\n")):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                set_paragraph(p, cl if k == 0 else "      " + cl.strip(), font_pt, DARK)
                p.space_after = Pt(3)

    if split:
        prompt_h_in = sum(estimate_lines(plain(l), font_pt + 1, full_w_in) * (font_pt + 1) * line_gap / 72.0 + 3 / 72.0
                          for l in prompt_lines) + 0.15
        pb = add_textbox(slide, margin, content_top, Inches(full_w_in), Inches(prompt_h_in))
        write_prompt(pb.text_frame)
        lower_top = content_top_in + prompt_h_in
        lower_h = content_h_in - prompt_h_in
        lb = add_textbox(slide, margin, Inches(lower_top), Inches(LEFT_W), Inches(lower_h))
        write_passage(lb.text_frame, True)
        rb = rounded_box(slide, margin + Inches(LEFT_W + 0.15), Inches(lower_top), Inches(RIGHT_W), Inches(lower_h), GRAY_BG, GRAY_BG)
        rtf = rb.text_frame
        rtf.word_wrap = True
        rtf.vertical_anchor = MSO_ANCHOR.TOP
        rtf.margin_left = rtf.margin_right = Inches(0.08)
        rtf.margin_top = Inches(0.08)
        write_choices(rtf, True)
    else:
        box = add_textbox(slide, margin, content_top, text_w, Inches(content_h_in), anchor=MSO_ANCHOR.TOP)
        tf = box.text_frame
        write_prompt(tf)
        if passage_lines:
            spacer = tf.add_paragraph()
            set_paragraph(spacer, " ", 6, GRAY_TXT)
            write_passage(tf, False)
        if choice_lines:
            spacer = tf.add_paragraph()
            set_paragraph(spacer, " ", 8, GRAY_TXT)
            write_choices(tf, False)

    if item.get("table"):
        add_table(slide, item["table"], margin + Inches(7.1), content_top + Inches(0.6), Inches(5.2), font_pt=14)
    if item.get("image"):
        img_h = min(Inches(content_h_in - 0.2), Inches(3.2))
        slide.shapes.add_picture(item["image"], margin + Inches(8.0), content_top + Inches(0.1), height=img_h)

    # 구분선
    line = slide.shapes.add_connector(1, margin, Inches(divider_y), SLIDE_W - margin, Inches(divider_y))
    line.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    line.line.width = Pt(1)

    # ---- 정답 배너 (클릭 1) ----
    ans_box = rounded_box(slide, margin, Inches(ans_top), content_w, Inches(ans_h_in), ANSWER_BG, ANSWER_BORDER)
    set_shape_name(ans_box, "reveal_1_en")
    atf = ans_box.text_frame
    atf.word_wrap = True
    atf.vertical_anchor = MSO_ANCHOR.MIDDLE
    atf.margin_left = Inches(0.2)
    atf.margin_right = Inches(0.2)
    ans_text = item["answer"]
    n = len(plain(ans_text))
    ans_font = 20 if n < 60 else (16 if n < 95 else 13)
    p = atf.paragraphs[0]
    set_paragraph(p, f"정답:  {ans_text}", ans_font, ANSWER_TXT, bold=True, align=PP_ALIGN.CENTER)

    # ---- 해설 패널 (클릭 2) ----
    expl_box = rounded_box(slide, margin, Inches(expl_top), content_w, Inches(expl_h_in), EXPL_BG, EXPL_BORDER)
    set_shape_name(expl_box, "reveal_2_en")
    etf = expl_box.text_frame
    etf.word_wrap = True
    etf.vertical_anchor = MSO_ANCHOR.TOP
    etf.margin_left = Inches(0.25)
    etf.margin_right = Inches(0.25)
    etf.margin_top = Inches(0.1)
    first = True
    for i, line in enumerate(expl_lines):
        p = etf.paragraphs[0] if first else etf.add_paragraph()
        first = False
        set_paragraph(p, line if line.strip() else " ", expl_pt, EXPL_TXT, bold=False)
        if i == 0:
            lab = p.add_run()
            lab.text = "정답 이유   "
            lab.font.size = Pt(expl_pt); lab.font.bold = True; lab.font.name = FONT_KR
            lab.font.color.rgb = EXPL_BORDER
            r_el = lab._r
            r_el.getparent().remove(r_el)
            pPr = p._p.find(qn("a:pPr"))
            (pPr.addnext(r_el) if pPr is not None else p._p.insert(0, r_el))
        p.space_after = Pt(2)

    LAYOUT_LOG.append({"slide": len(prs.slides), "divider": divider_y, "split": split, "content_top": content_top_in,
                       "expl_top": expl_top, "font": font_pt, "expl_pt": expl_pt})
    return slide


def build_section_slide(prs, sec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(1.25))
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background(); bg.shadow.inherit = False
    t = add_textbox(slide, Inches(0.5), Inches(0.2), Inches(12.3), Inches(0.9), anchor=MSO_ANCHOR.MIDDLE)
    set_paragraph(t.text_frame.paragraphs[0], sec["title"], 30, WHITE, bold=True)
    box = add_textbox(slide, Inches(0.6), Inches(1.5), Inches(12.1), Inches(5.7))
    tf = box.text_frame
    first = True
    for kind, text in sec["lines"]:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        if kind == "h":
            set_paragraph(p, text, sec.get("pt", 18), NAVY, bold=True); p.space_after = Pt(6)
        elif kind == "ex":
            set_paragraph(p, "▸ " + text, sec.get("pt", 18), DARK, bold=True); p.space_before = Pt(8)
        elif kind == "kr":
            set_paragraph(p, "    " + text, sec.get("pt", 18) - 3, GRAY_TXT)
        else:
            set_paragraph(p, text, sec.get("pt", 18) - 2, DARK); p.space_after = Pt(3)
    LAYOUT_LOG.append({"slide": len(prs.slides), "section": True})
    return slide


def build_cover_slide(prs, deck_title, book, n_items):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY
    bg.line.fill.background()
    bg.shadow.inherit = False

    box = add_textbox(slide, Inches(1.0), Inches(2.6), Inches(11.3), Inches(1.6), anchor=MSO_ANCHOR.MIDDLE)
    p = box.text_frame.paragraphs[0]
    set_paragraph(p, deck_title, 34, WHITE, bold=True)

    box2 = add_textbox(slide, Inches(1.0), Inches(4.05), Inches(11.3), Inches(0.6))
    p2 = box2.text_frame.paragraphs[0]
    set_paragraph(p2, book, 18, RGBColor(0xCF, 0xD8, 0xE8))

    box3 = add_textbox(slide, Inches(1.0), Inches(4.65), Inches(11.3), Inches(0.5))
    p3 = box3.text_frame.paragraphs[0]
    set_paragraph(p3, f"총 {n_items}문항 · 클릭 → 정답 공개 → 클릭 → 정답 이유 공개", 14, RGBColor(0xAE, 0xBB, 0xD0))
    return slide


# ----------------------------------------------------------------------
# 네이티브 클릭 애니메이션(Appear) 주입 — OOXML 레벨
# ----------------------------------------------------------------------
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def appear_block(spid, ids):
    a, b, c, d = ids
    return f'''
        <p:par>
          <p:cTn id="{a}" fill="hold">
            <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
            <p:childTnLst>
              <p:par>
                <p:cTn id="{b}" fill="hold">
                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="{c}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst>
                          <p:set>
                            <p:cBhvr>
                              <p:cTn id="{d}" dur="1" fill="hold">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                              </p:cTn>
                              <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                              <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                            </p:cBhvr>
                            <p:to><p:strVal val="visible"/></p:to>
                          </p:set>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
        </p:par>'''


def build_timing_xml(reveals):
    blocks = []
    next_id = 10
    for spid in reveals:
        ids = [next_id + i for i in range(4)]
        blocks.append(appear_block(spid, ids))
        next_id += 10
    bld = "".join(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>' for spid in reveals)
    xml = f'''<p:timing xmlns:p="{P_NS}">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    {"".join(blocks)}
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst>
                  <p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
                </p:prevCondLst>
                <p:nextCondLst>
                  <p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
                </p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      <p:bldLst>{bld}</p:bldLst>
    </p:timing>'''
    return etree.fromstring(xml.encode("utf-8"))


def find_reveal_spids(root):
    items = []
    for cNvPr in root.iter(f"{{{P_NS}}}cNvPr"):
        m = None
        name = cNvPr.get("name", "")
        if name.startswith("reveal_") and name.endswith("_en"):
            order = int(name.split("_")[1])
            items.append((order, int(cNvPr.get("id"))))
    items.sort(key=lambda t: t[0])
    return [spid for _, spid in items]


def inject_click_animations(pptx_path):
    workdir = pptx_path + ".unpacked"
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(workdir)

    slides_dir = os.path.join(workdir, "ppt", "slides")
    n_animated = 0
    for fname in sorted(os.listdir(slides_dir)):
        if not (fname.startswith("slide") and fname.endswith(".xml")):
            continue
        path = os.path.join(slides_dir, fname)
        tree = etree.parse(path, etree.XMLParser(remove_blank_text=False))
        root = tree.getroot()
        for existing in root.findall(f"{{{P_NS}}}timing"):
            root.remove(existing)
        reveals = find_reveal_spids(root)
        if not reveals:
            continue
        root.append(build_timing_xml(reveals))
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
        n_animated += 1

    out_path = pptx_path
    tmp_out = pptx_path + ".tmp"
    if os.path.exists(tmp_out):
        os.remove(tmp_out)
    with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zf:
        for base, _, files in os.walk(workdir):
            for f in files:
                full = os.path.join(base, f)
                rel = os.path.relpath(full, workdir)
                zf.write(full, rel)
    shutil.move(tmp_out, out_path)
    shutil.rmtree(workdir)
    return n_animated


# ----------------------------------------------------------------------
def main():
    if len(sys.argv) < 3:
        print("usage: python3 generate_quiz_pptx.py items.json output.pptx")
        sys.exit(1)
    items_path, out_path = sys.argv[1], sys.argv[2]

    with open(items_path, encoding="utf-8") as f:
        data = json.load(f)

    items = data["items"]
    book = data.get("book", "")
    deck_title = data.get("deck_title", "문제 풀이")

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    build_cover_slide(prs, deck_title, book, len(items))
    LAYOUT_LOG.append({"slide": 1, "cover": True})
    sections = {s["before"]: s for s in data.get("sections", [])}
    for i, item in enumerate(items, start=1):
        if item["id"] in sections:
            build_section_slide(prs, sections[item["id"]])
        build_item_slide(prs, item, i, len(items))

    prs.save(out_path)
    n_animated = inject_click_animations(out_path)
    with open(out_path + ".layout.json", "w", encoding="utf-8") as lf:
        json.dump(LAYOUT_LOG, lf, ensure_ascii=False, indent=1)
    print(f"saved {out_path}: {len(items)+1} slides, {n_animated} slides with click animations")


if __name__ == "__main__":
    main()
