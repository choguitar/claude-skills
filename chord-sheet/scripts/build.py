"""코드 악보 텍스트를 2단 A4 악보 docx로 만든다.

사용법:
  python build.py SONGS_TXT OUT_DOCX                      # 새 악보 파일
  python build.py SONGS_TXT OUT_DOCX --append-to EXISTING  # 기존 악보 뒤에 곡 추가

텍스트 형식은 SKILL.md 참조. 서식 수치의 정본은 이 파일의 상수다.

배치 규칙 — 한 곡은 한 페이지 안에 담는다:
  도돌이표(|: :|)는 입력에 있어도 펼쳐서 쓰고, 한 페이지를 넘을 때만 다시 접는다.
  1. 섹션 단위로 단을 넘기지 않는다(섹션이 단 중간에서 잘리지 않음).
  2. 한 페이지를 넘으면 섹션 안의 연속 반복 구간을 도돌이표로 접는다.
  3. 그래도 넘으면 코드 줄과 그 아래 가사 묶음 단위로만 붙여 둔다.
  4. 그래도 넘으면 그 곡만 줄 간격 → 글자 크기 순으로 줄인다(SHRINK_STEPS).
  한 단에 들어가는 곡은 앞 곡이 왼쪽 단만 썼으면 오른쪽 단에, 아니면 새 페이지에서 시작한다.
"""
import argparse
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

BASE_DOCX = Path(__file__).resolve().parent.parent / 'assets' / 'base.docx'

# 단위: twip(1/20pt). w:spacing 자간 값도 같은 단위다. 글자 크기는 half-point(28 = 14pt).
FONT = 'D2Coding'
BASE_SIZE = 28            # 코드·가사 14pt
HALF_GLYPH = 140          # D2Coding 14pt 반각 폭. 한글 등 전각은 2배
MEASURES_PER_LINE = 4
MEASURE_SLOT = 1260       # 한 마디 칸 폭. 줄의 마지막 마디는 단 끝까지 쓴다
HALF_SLOT = MEASURE_SLOT // 2  # 반 마디. 코드 2개인 마디의 두 번째 코드 위치(4/4박자 3박)
BASE_SPACING = -20        # 코드·가사 기본 자간 -1pt
TIGHT_SPACING = -40       # '코드 좁게' 자간 -2pt
LYRIC_MIN_SPACING = -40   # 가사 자동 축소 하한. 이보다 더 줄여야 하면 줄바꿈을 허용하고 경고
MIN_GAP = 40              # 마디 끝과 다음 탭 사이 최소 여백
RIGHT_EDGE_GAP = 20       # 줄 끝 반복 표시(':|')를 단 오른쪽 끝에서 이만큼 안쪽에 맞춘다
LINE_SPACING = 216        # 줄 간격 0.9배 (240 = 1배)
SECTION_SPACE_BEFORE = 120  # 섹션 위 6pt — 섹션 사이 빈 줄을 대체

# 문단 높이 실측값 — 2026-09-23, Word(Microsoft 365, Windows)에서 D2Coding 기본 서식으로 측정.
# 폰트·크기·줄 간격 상수를 바꾸면 check_layout.py 결과로 다시 잰다.
LINE_HEIGHT = 386         # 코드·가사 한 줄 (14pt, 줄 간격 0.9)
SECTION_HEIGHT = 336      # 섹션 줄 (위 여백 제외)
TITLE_HEIGHT = 316
INFO_HEIGHT = 302
FIT_MARGIN = 0.98         # 높이 추정 오차 여유

# 한 페이지를 넘는 곡을 줄이는 단계: (글자 크기 half-point, 줄 간격)
SHRINK_STEPS = ((28, 204), (28, 192), (26, 192), (24, 192))

# 반 마디마다 탭. 마지막 마디의 반 마디 지점까지
TAB_STOPS = [HALF_SLOT * k for k in range(1, 2 * (MEASURES_PER_LINE - 1) + 2)]
TIGHT_STYLE = '<w:rStyle w:val="ChordTight"/>'

STYLE_IDS = ('SheetBase', 'SongTitle', 'SongInfo', 'SongSection', 'Chord', 'Lyric', 'ChordTight')


def styles_xml():
    rfonts = f'<w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
    tabs = ''.join(f'<w:tab w:val="left" w:pos="{pos}"/>' for pos in TAB_STOPS)
    return (
        f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="SheetBase"><w:name w:val="악보 기본"/><w:qFormat/><w:pPr><w:keepLines/><w:widowControl w:val="0"/><w:spacing w:after="0" w:line="{LINE_SPACING}" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:rPr>{rfonts}<w:color w:val="000000"/><w:spacing w:val="{BASE_SPACING}"/><w:sz w:val="{BASE_SIZE}"/><w:szCs w:val="{BASE_SIZE + 4}"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="SongTitle"><w:name w:val="악보 제목"/><w:basedOn w:val="SheetBase"/><w:next w:val="SongInfo"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:spacing w:val="0"/><w:sz w:val="24"/><w:szCs w:val="28"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="SongInfo"><w:name w:val="곡 정보"/><w:basedOn w:val="SheetBase"/><w:next w:val="SongSection"/><w:qFormat/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:spacing w:val="0"/><w:sz w:val="22"/><w:szCs w:val="24"/></w:rPr></w:style>'
        f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="SongSection"><w:name w:val="섹션"/><w:basedOn w:val="SheetBase"/><w:next w:val="Chord"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="{SECTION_SPACE_BEFORE}"/></w:pPr><w:rPr><w:spacing w:val="0"/><w:sz w:val="24"/><w:szCs w:val="28"/></w:rPr></w:style>'
        f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Chord"><w:name w:val="코드"/><w:basedOn w:val="SheetBase"/><w:next w:val="Lyric"/><w:qFormat/><w:pPr><w:keepNext/><w:tabs>{tabs}</w:tabs></w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Lyric"><w:name w:val="가사"/><w:basedOn w:val="SheetBase"/><w:qFormat/></w:style>'
        f'<w:style w:type="character" w:customStyle="1" w:styleId="ChordTight"><w:name w:val="코드 좁게"/><w:qFormat/><w:rPr><w:spacing w:val="{TIGHT_SPACING}"/></w:rPr></w:style>'
    )


def column_geometry(document_xml):
    """본문 sectPr의 용지·여백·단 설정에서 (단 폭, 단 높이)를 구한다."""
    sect = re.findall(r'<w:sectPr.*?</w:sectPr>', document_xml, re.S)[-1]

    def attr(tag, name):
        return int(re.search(rf'<w:{tag} [^>]*w:{name}="(\d+)"', sect).group(1))

    cols = re.search(r'<w:cols [^>]*/>', sect)
    num = int(re.search(r'w:num="(\d+)"', cols.group(0)).group(1)) if cols and 'w:num=' in cols.group(0) else 1
    space = int(re.search(r'w:space="(\d+)"', cols.group(0)).group(1)) if num > 1 else 0
    width = (attr('pgSz', 'w') - attr('pgMar', 'left') - attr('pgMar', 'right') - space * (num - 1)) // num
    height = attr('pgSz', 'h') - attr('pgMar', 'top') - attr('pgMar', 'bottom')
    return width, height


def text_width(text, spacing, size=BASE_SIZE):
    glyph = HALF_GLYPH * size / BASE_SIZE
    return sum(glyph * (2 if unicodedata.east_asian_width(c) in 'WF' else 1) + spacing for c in text)


def run(text, rpr=''):
    return f'<w:r>{f"<w:rPr>{rpr}</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def column_top_spacer(height):
    """단 나누기 + 다음 단 맨 위 빈 공간(height twip). 단 나누기 문단은 1pt로 둬서 앞 단 끝에 공간을 거의 쓰지 않는다."""
    thin = 20
    return [f'<w:p><w:pPr><w:pStyle w:val="Lyric"/><w:spacing w:line="{thin}" w:lineRule="exact"/></w:pPr>'
            '<w:r><w:br w:type="column"/></w:r></w:p>',
            f'<w:p><w:pPr><w:pStyle w:val="Lyric"/><w:keepNext/><w:spacing w:line="{height - thin}" w:lineRule="exact"/></w:pPr></w:p>']


def split_measures(line):
    """'|' 하나가 한 마디를 연다. 빈 마디('|' 뒤가 비었음)는 앞 코드가 이어지는 마디라 그대로 둔다.
    ':|'는 그 마디의 반복 끝 표시로 마디 끝에 붙는다 — 줄 끝이든 줄 중간(':|' 바로 뒤가 다음 마디)이든.
    '|:Em |A7 | |C G :|' → ['|:Em', '|A7', '|', '|C G :|'],  '|C | :|' → ['|C', '| :|'],
    '|C :|Am |F' → ['|C :|', '|Am', '|F']"""
    raw_parts = line.split('|')[1:]
    parts = [' '.join(p.split()) for p in raw_parts]
    measures, i = [], 0
    while i < len(parts):
        text = parts[i]
        # ':' 하나뿐인 칸은 '|:'로 붙어 있으면 빈 마디의 반복 시작, 띄어 있으면 빈 마디의 반복 끝
        is_start_only = text == ':' and raw_parts[i].startswith(':')
        if text.endswith(':') and not is_start_only:
            body = text[:-1].rstrip()
            measures.append(f'|{body} :|' if body else '| :|')
            i += 1
            if i == len(parts) - 1 and parts[i] == '':  # 줄 끝 ':|'의 '|' 뒤 빈칸은 마디가 아니다
                i += 1
            continue
        measures.append('|' + text)
        i += 1
    return measures


def join_measures(measures):
    """split_measures의 역. ':|' 뒤 마디는 그 '|'를 공유한다: ['|G :|', '|Am'] → '|G :|Am'."""
    line = ''
    for m in measures:
        if line.endswith(':|'):
            line += m[1:] if m != '|' else ' |'
        else:
            line += (' ' if line else '') + m
    return line


def parse(text):
    """텍스트 → [{'title', 'info', 'sections': [{'name', 'lines': [(kind, text)]}]}]"""
    songs = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if raw.startswith('# '):
            songs.append({'title': line[2:].strip(), 'info': None, 'sections': []})
            continue
        if not songs:
            sys.exit(f'첫 곡 제목(# 로 시작하는 줄)보다 앞에 내용이 있습니다: {line}')
        song = songs[-1]
        if line.startswith('['):
            song['sections'].append({'name': line, 'lines': []})
            continue
        kind = 'chord' if line.startswith('|') else 'lyric'
        if kind == 'lyric' and song['info'] is None and not song['sections']:
            song['info'] = line
            continue
        if not song['sections']:
            song['sections'].append({'name': None, 'lines': []})
        song['sections'][-1]['lines'].append((kind, line))
    return songs


def group_end(lines, i):
    """i번째 줄부터 시작한 '코드 줄 + 아래 가사' 묶음이 끝나는 위치(다음 코드 줄 또는 끝)."""
    j = i + 1
    while j < len(lines) and lines[j][0] != 'chord':
        j += 1
    return j


def expand_marks(lines):
    """|: ... :| 구간을 두 번 펼친다. 구간은 줄 맨 앞 '|:'로 여는 코드 줄부터 줄 맨 끝 ':|'로 닫는 코드 줄 아래 가사까지.
    줄 중간의 반복 표시(예: '|C :|Am')는 한 줄 안의 반복이라 펼치지 않고 그대로 둔다."""
    lines = list(lines)
    line_level = lambda t: ':|' not in t[:-2]  # 줄 중간에 반복 끝이 없다
    while True:
        start = next((i for i, (k, t) in enumerate(lines) if k == 'chord' and t.startswith('|:') and line_level(t)), None)
        if start is None:
            return lines
        end = next((i for i in range(start, len(lines))
                    if lines[i][0] == 'chord' and lines[i][1].endswith(':|') and line_level(lines[i][1])), None)
        if end is None:
            return lines
        block = lines[start:group_end(lines, end)]
        block[0] = ('chord', '|' + block[0][1][2:])
        block[end - start] = ('chord', block[end - start][1][:-2].rstrip())
        lines = lines[:start] + block + block + lines[start + len(block):]


def compress_marks(lines):
    """연속으로 두 번 나오는 묶음 구간을 |: ... :| 로 접는다. 가장 긴 구간부터 찾는다."""
    out, i = [], 0
    while i < len(lines):
        length = 0
        if lines[i][0] == 'chord':
            for size in range((len(lines) - i) // 2, 0, -1):
                mid, end = i + size, i + 2 * size
                at_boundary = all(p == len(lines) or lines[p][0] == 'chord' for p in (mid, end))
                if at_boundary and lines[i:mid] == lines[mid:end]:
                    length = size
                    break
        if not length:
            out.append(lines[i])
            i += 1
            continue
        block = lines[i:i + length]
        last_chord = max(j for j, (k, _) in enumerate(block) if k == 'chord')
        block[0] = ('chord', '|:' + block[0][1][1:])
        block[last_chord] = ('chord', block[last_chord][1] + ' :|')
        out += block
        i += 2 * length
    return out


def expand_song(song):
    return {**song, 'sections': [{**sec, 'lines': expand_marks(sec['lines'])} for sec in song['sections']]}


def compress_song(song):
    return {**song, 'sections': [{**sec, 'lines': compress_marks(sec['lines'])} for sec in song['sections']]}


class Builder:
    def __init__(self, col_width, col_height):
        self.col_width = col_width
        self.col_height = col_height * FIT_MARGIN
        self.warnings = []

    # ---- 가로: 한 줄 안에 넣기 ----

    def lyric_spacing(self, line, size):
        """가사 줄 자간. 기본으로 들어가면 None, 줄바꿈이 불가피하면 'wrap'."""
        if text_width(line, BASE_SPACING, size) <= self.col_width:
            return None
        spacing = int((self.col_width - text_width(line, 0, size)) // len(line))
        return 'wrap' if spacing < LYRIC_MIN_SPACING else spacing

    def measure_fit(self, measure, slot, size):
        """(rPr 조각, 자간). 칸에 안 들어가면 '코드 좁게', 그래도 넘치면 필요한 만큼 자간을 직접 줄인다."""
        if text_width(measure, BASE_SPACING, size) + MIN_GAP <= slot:
            return '', BASE_SPACING
        if text_width(measure, TIGHT_SPACING, size) + MIN_GAP <= slot:
            return TIGHT_STYLE, TIGHT_SPACING
        spacing = int((slot - MIN_GAP - text_width(measure, 0, size)) // len(measure))
        self.warnings.append(f'코드 마디 자간 {spacing / 20:.1f}pt로 축소: {measure}')
        return f'{TIGHT_STYLE}<w:spacing w:val="{spacing}"/>', spacing

    def measure_pieces(self, measure, i, is_last, size):
        """마디를 조각 [(탭 위치, 텍스트, rPr 조각, 자간, 추가 탭)]으로 나눈다.
        추가 탭이 None이면 스타일의 반 마디 탭에서 시작한다. 'left'/'right'면 그 위치에 문단 탭을 더한다
        ('right'는 탭 위치에서 끝난다).
        코드 2개는 반 마디씩 — 반 마디 칸에 안 들어가면 '코드 좁게', 그래도 안 되면 한 조각.
        줄의 마지막 마디는 단 끝까지 쓸 수 있다.
        반복 끝은 마디 끝에 붙인다 — 줄 끝이면 ':|'를 오른쪽 맞춤, 줄 중간이면 ':'를 경계 바로 앞에 두어
        다음 마디의 '|'와 이어 ':|'가 된다. (오른쪽 맞춤 탭은 다음 탭까지의 글자를 통째로 맞추므로 줄 중간엔 못 쓴다)"""
        start = MEASURE_SLOT * i
        slot_end = self.col_width if is_last else start + MEASURE_SLOT
        mark = []
        if measure.endswith(':|'):
            measure = measure[:-2].rstrip() or '|'
            if is_last:
                mark = [(self.col_width - RIGHT_EDGE_GAP, ':|', '', BASE_SPACING, 'right')]
                slot_end = mark[0][0] - text_width(':|', BASE_SPACING, size)
            else:
                slot_end -= text_width(':', BASE_SPACING, size)
                mark = [(slot_end, ':', '', BASE_SPACING, 'left')]
        slot = slot_end - start
        chords = measure[1:].split()
        if len(chords) == 2:
            halves = (('|' + chords[0], HALF_SLOT), (chords[1], slot - HALF_SLOT))
            for spacing, rpr in ((BASE_SPACING, ''), (TIGHT_SPACING, TIGHT_STYLE)):
                if all(text_width(t, spacing, size) + MIN_GAP <= s for t, s in halves):
                    return [(start, halves[0][0], rpr, spacing, None),
                            (start + HALF_SLOT, halves[1][0], rpr, spacing, None)] + mark
        rpr, spacing = self.measure_fit(measure, slot, size)
        return [(start, measure, rpr, spacing, None)] + mark

    def chord_bodies(self, line, size):
        """코드 줄 → 줄마다 (문단 탭 설정, run XML). 5마디 이상은 4마디씩 나눈다.
        반복 끝 표시가 있는 줄은 그 자리의 탭을 문단에 추가한다."""
        measures = split_measures(line)
        if len(measures) > MEASURES_PER_LINE:
            self.warnings.append(f'{len(measures)}마디 줄을 {MEASURES_PER_LINE}마디씩 나눔: {line}')
        size_rpr = '' if size == BASE_SIZE else f'<w:sz w:val="{size}"/><w:szCs w:val="{size + 4}"/>'
        out = []
        for first in range(0, len(measures), MEASURES_PER_LINE):
            chunk = measures[first:first + MEASURES_PER_LINE]
            pieces = [p for i, m in enumerate(chunk) for p in self.measure_pieces(m, i, i == len(chunk) - 1, size)]
            extra = [(kind, pos) for pos, *_, kind in pieces if kind]
            stops = sorted(set(TAB_STOPS) | {pos for _, pos in extra})
            body, cursor = [], 0
            for pos, text, rpr, spacing, kind in pieces:
                # Word 탭은 현재 위치보다 뒤의 첫 탭 위치로 간다 → pos까지 지나칠 탭 개수.
                # 앞 글자가 정확히 pos에서 끝나면(줄 중간 ':' 뒤의 '|') 탭 없이 바로 이어진다.
                body.append('<w:r><w:tab/></w:r>' * sum(1 for s in stops if cursor < s <= pos))
                body.append(run(text, rpr + size_rpr))
                cursor = pos if kind == 'right' else pos + text_width(text, spacing, size)
            tabs = ''.join(f'<w:tab w:val="{kind}" w:pos="{int(pos)}"/>' for kind, pos in extra)
            out.append((f'<w:tabs>{tabs}</w:tabs>' if tabs else '', ''.join(body)))
        return out

    # ---- 세로: 한 페이지 안에 넣기 ----

    def line_count(self, kind, text, size):
        if kind == 'chord':
            return -(-len(split_measures(text)) // MEASURES_PER_LINE)
        return 2 if self.lyric_spacing(text, size) == 'wrap' else 1

    def units(self, song, grouped, size, line):
        """붙여 둘 묶음들 [(높이, 섹션 제목으로 시작하는지)]. grouped=False면 섹션 단위, True면 '코드 줄 + 아래 가사' 단위."""
        line_height = LINE_HEIGHT * size * line / (BASE_SIZE * LINE_SPACING)
        out = []
        for sec in song['sections']:
            head = SECTION_HEIGHT + SECTION_SPACE_BEFORE if sec['name'] else 0
            groups = [[]]
            for kind, text in sec['lines']:
                if grouped and kind == 'chord' and groups[-1]:
                    groups.append([])
                groups[-1].append(self.line_count(kind, text, size) * line_height)
            out.append((head + sum(groups[0]), bool(sec['name'])))
            out.extend((sum(g), False) for g in groups[1:])
        return out

    def columns_needed(self, units, head, align):
        """단 순서대로 채울 때 (필요한 단 수, 둘째 단을 여는 묶음 번호). 한 묶음이 한 단보다 크면 (None, None).
        align=True면 둘째 단 첫 줄을 첫째 단의 첫 섹션 높이에 맞추는 빈 공간(top_offset)을 둘째 단 위에 둔다."""
        cols, used, split = 1, head, None  # 제목·곡 정보는 첫 묶음과 함께 움직인다
        for i, (h, starts_section) in enumerate(units):
            if h + (head if i == 0 else 0) > self.col_height:
                return None, None
            if i > 0 and used + h > self.col_height:
                cols += 1
                used = self.top_offset(head, starts_section) if align and cols == 2 else 0
                split = i if cols == 2 else split
                if used + h > self.col_height:
                    return None, None
            used += h
        return cols, split

    @staticmethod
    def top_offset(head, starts_section):
        """둘째 단 위 빈 공간. 섹션 제목으로 시작하면 그 섹션의 위 여백이 더해지므로 제목·곡 정보 높이만,
        코드 줄로 시작하면 섹션 위 여백만큼 더 둔다 — 첫째 단의 첫 섹션 줄과 같은 높이가 된다."""
        return head + (0 if starts_section else SECTION_SPACE_BEFORE)

    def plan(self, song):
        """(내용, 글자 크기, 줄 간격, 묶음 방식, 필요한 단 수, 둘째 단 맞춤) — 배치 규칙 순서대로 두 단(한 페이지) 안에 드는 첫 조합.
        둘째 단 맞춤은 (여는 묶음 번호, 빈 공간 높이) 또는 None — 맞춤은 '가능하면'이라, 넣으면 넘치는 조합에서는 뺀다."""
        head = TITLE_HEIGHT + (INFO_HEIGHT if song['info'] else 0)
        expanded = expand_song(song)
        folded = compress_song(expanded)
        options = [(expanded, BASE_SIZE, LINE_SPACING, False),
                   (folded, BASE_SIZE, LINE_SPACING, False),
                   (folded, BASE_SIZE, LINE_SPACING, True)]
        options += [(folded, size, line, grouped) for size, line in SHRINK_STEPS for grouped in (False, True)]
        align = None
        for content, size, line, grouped in options:
            units = self.units(content, grouped, size, line)
            cols, split = self.columns_needed(units, head, align=True)
            if cols and cols <= 2:
                align = (split, self.top_offset(head, units[split][1])) if split is not None else None
                break
            cols, _ = self.columns_needed(units, head, align=False)
            if cols and cols <= 2:
                self.warnings.append(f'둘째 단 첫 줄 높이 맞춤을 생략함(맞추면 한 페이지를 넘음): {song["title"]}')
                break
        else:
            self.warnings.append(f'최대로 줄여도 한 페이지를 넘음 — 섹션 생략 등 곡을 줄여야 함: {song["title"]}')
            cols = 3
        if content is folded and folded != expanded:
            self.warnings.append(f'한 페이지에 맞추려고 반복 구간을 도돌이표로 접음: {song["title"]}')
        return content, size, line, grouped, cols, align

    # ---- XML ----

    def song_paras(self, song, size, line, grouped, start, align):
        """start: None(문서 첫 곡) | 'column' | 'page'
        align: (둘째 단을 여는 묶음 번호, 빈 공간 높이) — 그 묶음 앞에서 단을 나누고 빈 공간을 둔다."""
        shrunk = (size, line) != (BASE_SIZE, LINE_SPACING)
        line_ppr = f'<w:spacing w:line="{line}" w:lineRule="auto"/>' if shrunk else ''
        size_rpr = f'<w:sz w:val="{size}"/><w:szCs w:val="{size + 4}"/>' if shrunk else ''
        if grouped:
            self.warnings.append(f'섹션째로는 한 페이지를 넘어 코드+가사 단위로 단을 나눔: {song["title"]}')
        if shrunk:
            self.warnings.append(f'한 페이지에 맞추려고 글자 {size / 2:g}pt · 줄 간격 {line / 240:.2f}배로 줄임: {song["title"]}')

        title_ppr = '<w:pageBreakBefore/>' if start == 'page' else ''
        title_brk = '<w:r><w:br w:type="column"/></w:r>' if start == 'column' else ''
        out = [f'<w:p><w:pPr><w:pStyle w:val="SongTitle"/>{title_ppr}</w:pPr>{title_brk}{run(song["title"])}</w:p>']
        if song['info']:
            out.append(f'<w:p><w:pPr><w:pStyle w:val="SongInfo"/></w:pPr>{run(song["info"])}</w:p>')

        unit = -1  # units()와 같은 순서로 센 묶음 번호
        for sec in song['sections']:
            # (스타일, 본문, 묶음 시작 여부, 문단 탭 설정)
            items = [('SongSection', run(sec['name']), True, '')] if sec['name'] else []
            has_lines = False  # 섹션 제목 바로 다음 코드 줄은 새 묶음이 아니다
            for kind, text in sec['lines']:
                if kind == 'chord':
                    bodies = self.chord_bodies(text, size)
                    items += [('Chord', b, grouped and n == 0 and has_lines, tabs) for n, (tabs, b) in enumerate(bodies)]
                    has_lines = True
                    continue
                has_lines = True
                spacing = self.lyric_spacing(text, size)
                if spacing == 'wrap':
                    self.warnings.append(f'가사가 길어 줄바꿈됨 (두 줄로 나누길 권장): {text}')
                rpr = (f'<w:spacing w:val="{spacing}"/>' if isinstance(spacing, int) else '') + size_rpr
                items.append(('Lyric', run(text, rpr), False, ''))
            for n, (style, body, starts_unit, tabs) in enumerate(items):
                if n == 0 or starts_unit:
                    unit += 1
                    if align and unit == align[0]:
                        out += column_top_spacer(align[1])
                last_in_unit = n == len(items) - 1 or items[n + 1][2]
                keep = '<w:keepNext w:val="0"/>' if last_in_unit else '<w:keepNext/>'
                ppr = line_ppr if style in ('Chord', 'Lyric') else ''
                out.append(f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{keep}{tabs}{ppr}</w:pPr>{body}</w:p>')
        return out

    def build(self, songs, starts_on_new_page):
        paras, right_column_free = [], False
        for n, song in enumerate(songs):
            content, size, line, grouped, cols, align = self.plan(song)
            if n == 0 and not starts_on_new_page:
                start = None
            elif right_column_free and cols == 1:
                start = 'column'
            else:
                start = 'page'
            paras += self.song_paras(content, size, line, grouped, start, align)
            right_column_free = start != 'column' and cols == 1
        return paras


def inject_styles(styles):
    for sid in STYLE_IDS:
        styles = re.sub(rf'<w:style [^>]*w:styleId="{sid}".*?</w:style>', '', styles, flags=re.S)
    return styles.replace('</w:styles>', styles_xml() + '</w:styles>')


def apply_settings(settings):
    if '<w:embedTrueTypeFonts' not in settings:
        settings = re.sub(r'(<w:zoom [^>]*/>)', r'\1<w:embedTrueTypeFonts/>', settings, count=1)
    return re.sub(r'(w:name="compatibilityMode" w:uri="[^"]*" w:val=")\d+"', r'\g<1>15"', settings)


def main():
    ap = argparse.ArgumentParser(description='코드 악보 텍스트 → docx')
    ap.add_argument('songs_txt')
    ap.add_argument('out_docx')
    ap.add_argument('--append-to', help='곡을 이어 붙일 기존 악보 docx (이 스크립트로 만든 파일)')
    args = ap.parse_args()

    source = Path(args.append_to) if args.append_to else BASE_DOCX
    with zipfile.ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    document = entries['word/document.xml'].decode('utf8')
    styles = entries['word/styles.xml'].decode('utf8')
    if args.append_to and 'w:styleId="SongTitle"' not in styles:
        sys.exit('기존 파일에 악보 스타일이 없습니다. extract.py로 텍스트를 뽑아 새 파일로 빌드하세요.')

    songs = parse(Path(args.songs_txt).read_text(encoding='utf8'))
    if not songs:
        sys.exit('곡이 없습니다. 제목 줄은 "# 가수 – 제목" 형식입니다.')
    builder = Builder(*column_geometry(document))
    paras = builder.build(songs, starts_on_new_page=bool(args.append_to))

    body_sect = document.rindex('<w:sectPr')
    entries['word/document.xml'] = (document[:body_sect] + ''.join(paras) + document[body_sect:]).encode('utf8')
    entries['word/styles.xml'] = inject_styles(styles).encode('utf8')
    entries['word/settings.xml'] = apply_settings(entries['word/settings.xml'].decode('utf8')).encode('utf8')

    with zipfile.ZipFile(args.out_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)

    print(f'완료: {len(songs)}곡 → {args.out_docx}')
    for w in builder.warnings:
        print(f'[경고] {w}')


if __name__ == '__main__':
    main()
