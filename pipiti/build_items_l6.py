# -*- coding: utf-8 -*-
"""paras.json (hwp2json.py 결과, 중2 천재(이상기) Lesson 6 Grammar Build Up 워크북)
-> items_L6.json (피피티 스킬 스키마: 문제 -> 클릭 -> 정답)

HWP 구조: Part N 머리글 뒤에 [지시문 영/한] [보기·박스] 그리고 문항들.
  문항 = (스타일5 문제문장)* + 정답 줄(스타일0, 들여쓰기 ' ') + 뒤따르는 줄(빈칸 틀·힌트)
  스타일5 문제가 없는 Part(수동태 Part 9, so~that Part 1·5·6)는 정답 줄이 문항의 시작이다.
이 워크북에는 해설이 없으므로 explanation 은 비워 둔다(정답만 공개).
"""
import json, re, sys

src = sys.argv[1]
out = sys.argv[2]
paras = json.load(open(src, encoding='utf-8'))
HAN = re.compile(r'[가-힣]')
BLANKISH = re.compile(r'^[\s_→=\-\.\(\)\d]*$')
TAG = re.compile(r'^(Pattern \d|Part \d+)$')


def clean(t):
    return re.sub(r'[ \t]*\t[ \t]*', '    ', t.replace('\xa0', ' ')).rstrip()


# ---- Part 경계 ----
starts = [i for i, (k, s, l, t) in enumerate(paras) if k == 'P' and l == 6 and re.fullmatch(r'Part \d+', t)]
END = next(i for i, (k, s, l, t) in enumerate(paras) if k == 'P' and t == 'Grammar Build Up 워크북' and i > starts[-1])
bounds = []
for st, nx in zip(starts, starts[1:] + [END]):
    # 다음 Pattern 도입부('Lesson' 머리글)가 먼저 나오면 거기서 Part 를 끝낸다
    stop = next((i for i in range(st + 1, nx) if paras[i][0] == 'P' and paras[i][3] == 'Lesson' and paras[i][2] == 6), nx)
    bounds.append((st, stop))

PATTERN = {1: '수동태', 2: 'so ~ that'}
items, stats = [], []
pat = 1
for n, (a, b) in enumerate(bounds):
    pname = f'Part {int(paras[a][3].split()[1])}'
    if paras[a][3] == 'Part 1' and n > 0:
        pat = 2
    partno = int(paras[a][3].split()[1])
    seq = []
    for k, s, l, t in paras[a + 1:b]:
        if k != 'P' or not t.strip() or TAG.match(t.strip()):
            continue
        seq.append((s, l, t))

    def is_ans(e):
        s, l, t = e
        return s == 0 and l == 3 and t.startswith(' ')

    first = next((i for i, e in enumerate(seq) if e[0] == 5 or is_ans(e)), len(seq))
    pre = [clean(t) for s, l, t in seq[:first]]
    ko_i = next((i for i, t in enumerate(pre) if HAN.search(t) and not t.startswith(('<', '•'))), None)
    instr = pre[ko_i] if ko_i is not None else (pre[0] if pre else '')
    extras = pre[ko_i + 1:] if ko_i is not None else pre[1:]

    # 문항 묶기
    body = seq[first:]
    groups, cur = [], None
    has_q = any(e[0] == 5 for e in body)
    for e in body:
        if e[0] == 5:
            if cur is None or cur['ans'] is not None:
                cur = dict(q=[], ans=None, post=[])
                groups.append(cur)
            cur['q'].append(clean(e[2]))
        elif is_ans(e):
            if cur is None or cur['ans'] is not None or (not has_q):
                cur = dict(q=[], ans=None, post=[])
                groups.append(cur)
            cur['ans'] = clean(e[2]).strip()
        else:
            if cur is None:
                continue
            cur['post'].append(clean(e[2]))
    k = 0
    for g in groups:
        if g['ans'] is None:
            print('!! 정답 없음', pat, partno, g['q'][:1]); continue
        k += 1
        post = [x for x in g['post'] if not BLANKISH.match(x)]
        body_lines = []
        if extras:
            body_lines += extras + ['']
        body_lines += g['q'] + post
        item = dict(
            id=f'{pat}-{partno:02d}-{k:02d}',
            label=f'P{pat} · Part {partno} · {k:02d}번',
            topic=f'Pattern {pat}. {PATTERN[pat]}',
            passage='\n'.join(body_lines).strip('\n'),
            prompt=instr,
            choices=None,
            kind='text',
            answer=re.sub(r'\s*\n\s*', '   /   ', g['ans']),
            explanation='',
        )
        items.append(item)
    stats.append((pat, partno, sum(1 for e in body if e[0] == 5), k))

for s in stats:
    print('P%d Part %2d  sty5=%3d  items=%3d' % s)
print('TOTAL', len(items))

# ---- 패턴 요약 슬라이드 ----
def split_en_ko(t):
    t = t.lstrip('• ').strip()
    if '\n' in t:
        en, ko = t.split('\n', 1)
        return en.strip(), ko.strip()
    m = HAN.search(t)
    return (t[:m.start()].strip(), t[m.start():].strip()) if m else (t, '')

def P(i):
    return paras[i][3]

sec1 = [('t', P(i).lstrip('• ').strip()) for i in range(18, 24)]
sec1.append(('h', '▼ 교과서 예문'))
for i in (28, 31, 32, 33):
    en, ko = split_en_ko(P(i))
    sec1 += [('ex', en), ('kr', ko)]
sec2 = [('t', P(1774))]
sec2.append(('h', '▼ 교과서 예문'))
for i in (1779, 1782, 1783, 1784):
    en, ko = split_en_ko(P(i))
    sec2 += [('ex', en), ('kr', ko)]

first1 = next(it['id'] for it in items if it['id'].startswith('1-'))
first2 = next(it['id'] for it in items if it['id'].startswith('2-'))
data = dict(
    book='중2 천재(이상기) Lesson 6 · Pattern 1 수동태 / Pattern 2 so ~ that',
    deck_title='Lesson 6  Grammar Build Up 워크북',
    items=items,
    sections=[
        dict(before=first1, title='Pattern 1.  수동태', lines=sec1, pt=18),
        dict(before=first2, title='Pattern 2.  so ~ that', lines=sec2, pt=18),
    ],
)
json.dump(data, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('saved', out)
