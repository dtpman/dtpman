# -*- coding: utf-8 -*-
"""items_raw.json(HWP에서 추출) -> items.json (피피티 스킬 스키마)"""
import json, re
raw = json.load(open('items_raw.json', encoding='utf-8'))
CIRC = '①②③④⑤'
items = []

def norm(s):
    s = s.replace(' ', ' ')
    s = re.sub(r'(?<=\S) {3,}(?=\S)', ' ', s)  # 원문 정렬용 연속 공백 정리
    return s.rstrip()

def split_choice_line(line):
    # "① wait\t\t② waited\t③ waiting" -> 개별 선택지
    parts = re.split(r'(?=[①②③④⑤])', line)
    return [p.strip() for p in parts if p.strip()]

for n, it in enumerate(raw, 1):
    body = [l for l in it['body'].split('\n') if l.strip() != '<<IMG>>' and not l.strip().startswith('Pattern ')]
    expl = it['expl']
    expl = re.sub(r'^Pattern \d\n', '', expl).strip('\n')
    # 정답/해설 분리
    m = re.match(r'^\s*([①②③④⑤](?:\s*,\s*[①②③④⑤])*)\s*\n?\s*:?\s*', expl)
    ans_marks = re.findall(r'[①②③④⑤]', m.group(1))
    exp_text = expl[m.end():].strip()
    exp_lines = [norm(l).strip() for l in exp_text.split('\n') if l.strip()]
    if not exp_lines:
        exp_lines = ['(원문 정답지에 별도 해설 없음)']
    # 본문 분리: 첫 선택지 이전 = 지문, 이후 = 선택지
    idx = next((i for i, l in enumerate(body) if l.lstrip().startswith('①')), None)
    assert idx is not None, n
    pre = [norm(l) for l in body[:idx]]
    post = [l.replace('\u00a0', ' ').rstrip() for l in body[idx:]]
    # 표 머리글 줄 "(A) (B)" / "ⓐ ⓑ" 감지 (지문 마지막 줄)
    header = None
    while pre and not pre[-1].strip():
        pre.pop()
    if pre and re.fullmatch(r'\s*((\([A-E]\)|[ⓐ-ⓔ])\s*)+', pre[-1]):
        header = re.findall(r'\([A-E]\)|[ⓐ-ⓔ]', pre[-1])
        pre.pop()
    while pre and not pre[0].strip():
        pre.pop(0)
    choices = []
    for l in post:
        if not l.strip():
            continue
        if l.lstrip()[0] in CIRC:
            choices.extend(split_choice_line(l))
        else:  # 이어지는 줄 (→ ..., = ...)
            choices[-1] += '\n' + norm(l).strip()
    assert len(choices) == 5, (n, choices)
    if header:
        new = []
        for c in choices:
            num, rest = c[0], c[1:].strip()
            fields = [f for f in re.split(r'\t+|\s+-\s+|\s{2,}', rest) if f.strip()]
            assert len(fields) == len(header), (n, c, fields)
            new.append(num + '  ' + '   '.join(f'{h} {f.strip()}' for h, f in zip(header, fields)))
        choices = new
    else:
        choices = [norm(re.sub(r'\t+', ' ', c)) for c in choices]
    chosen = [c for c in choices if c[0] in ans_marks]
    if len(ans_marks) == 1:
        answer = chosen[0].replace('\n', '  ')
    else:
        answer = ', '.join(ans_marks)
        short = '  /  '.join(c.replace('\n', ' ') for c in chosen)
        if len(short) < 110:
            answer += '   (' + short + ')'
    pattern = 1 if n <= 97 else 2
    num = n if pattern == 1 else n - 97
    item = dict(
        id=f'{n:03d}',
        label=f'Pattern {pattern} · {num:02d}번',
        topic='현재분사' if pattern == 1 else 'as + 형용사/부사의 원급 + as',
        passage='\n'.join(pre) if pre else None,
        prompt=norm(it['prompt']),
        choices=choices,
        kind='choice',
        answer=answer,
        explanation='\n'.join(exp_lines),
    )
    if it['table'] and n != 126:  # 126번 뒤 표는 저작권 안내 footer
        cells = it['table']
        ncol = {119: 4, 120: 4, 121: 5, 122: 4, 123: 5, 124: 4, 125: 3}[n]
        rows = [cells[i:i + ncol] for i in range(0, len(cells), ncol)]
        item['table'] = rows
    if n == 126:
        item['image'] = 'BIN0004_v.jpg'
    items.append(item)

assert len(items) == 126
sections = [
    dict(before='001', title='Pattern 1.  현재분사', lines=[
        ('t', '분사는 동사의 형태를 바꾸어 형용사처럼 사용할 수 있으며, 현재분사와 과거분사 두 가지가 있다.'),
        ('h', '현재분사: 동사원형 + -ing  → 능동·진행     |     과거분사: 동사원형 + -ed  → 수동·완료'),
        ('t', '분사가 단독으로 명사를 꾸미면 명사 앞, 목적어·부사(구) 등과 함께 쓰이면 명사 뒤에 온다.'),
        ('ex', 'Pacers usually have flags or balloons [[showing]] their finish time.'),
        ('kr', '페이서들은 보통 자신들의 완주 시간을 나타내는 깃발이나 풍선들을 가지고 있다.'),
        ('ex', 'The girl [[waiting]] at the bus stop is my sister.'),
        ('kr', '버스 정류장에서 기다리고 있는 소녀는 내 여동생이다.  (who[that] is 생략으로 볼 수 있음)'),
        ('ex', 'The bird [[singing]] in the tree is very big.'),
        ('kr', '나무에서 노래하고 있는 새는 매우 크다.'),
        ('ex', 'There are many students [[studying]] in the library.'),
        ('kr', '도서관에는 공부하고 있는 학생들이 많다.'),
        ('ex', 'The woman [[wearing]] glasses is my teacher.'),
        ('kr', '안경을 쓰고 있는 여자는 우리 선생님이다.'),
    ], pt=18),
    dict(before='098', title='Pattern 2.  as + 형용사/부사의 원급 + as', lines=[
        ('t', '‘as ~ as …’는 ‘…만큼 ~한/하게’라는 뜻의 원급 비교 표현으로, 두 비교 대상이 동등함을 나타낸다.'),
        ('h', 'as와 as 사이에는 반드시 형용사·부사의 원급(원래 형태)을 쓴다.   부정: not as[so] ~ as …  (…만큼 ~하지 않은)'),
        ('ex', 'They are [[as]] important [[as]] the players.'),
        ('kr', '그들은 선수들만큼 중요하다.'),
        ('ex', 'Jason can run [[as]] fast [[as]] Mike.'),
        ('kr', 'Jason은 Mike만큼 빠르게 달릴 수 있다.'),
        ('ex', 'I am [[as]] hungry [[as]] you.'),
        ('kr', '나는 너만큼 배고프다.'),
        ('ex', 'Canada is [[as]] large [[as]] the USA.'),
        ('kr', '캐나다는 미국만큼 크다.'),
        ('ex', 'Today is [[as]] windy [[as]] yesterday.'),
        ('kr', '오늘은 어제만큼 바람이 분다.'),
    ], pt=18),
]
json.dump(dict(book='중3 동아(윤정미) Lesson 5 · Pattern 1 현재분사 / Pattern 2 as+원급+as',
               deck_title='Lesson 5  Grammar Build Up 객관식',
               items=items, sections=sections),
          open('items.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('items:', len(items))
