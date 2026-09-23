"""기존 악보 docx에서 곡을 뽑아 build.py 입력 텍스트로 저장한다.

사용법: python extract.py IN_DOCX OUT_TXT

- build.py로 만든 파일은 단락 스타일로 구분한다.
- 그 외 파일은 내용으로 추정한다: '[' 섹션, '|' 코드, 굵고 16pt 이상이면 제목, 나머지 가사.
- 곡 정보 줄이 없는 곡은 마지막 코드의 근음으로 Key를 추정해 넣고, 추정한 곡을 출력한다.
"""
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

TITLE_MIN_HALF_POINTS = 32  # 16pt
STYLE_KIND = {'SongTitle': 'title', 'SongInfo': 'info', 'SongSection': 'section', 'Chord': 'chord', 'Lyric': 'lyric'}


def text_of(p):
    return ''.join(t if t else ' ' for t, _ in re.findall(r'<w:t[^>]*>([^<]*)</w:t>|(<w:tab/>)', p))


def kind_of(p, text):
    style = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    if style and style.group(1) in STYLE_KIND:
        return STYLE_KIND[style.group(1)]
    if text.startswith('['):
        return 'section'
    if text.startswith('|'):
        return 'chord'
    size = re.search(r'<w:sz w:val="(\d+)"', p)
    if '<w:b/>' in p and size and int(size.group(1)) >= TITLE_MIN_HALF_POINTS:
        return 'title'
    return 'lyric'


def normalize_chord(text):
    measures = [' '.join(m.split()) for m in text.split('|')[1:]]
    line = ' '.join('|' + m for m in measures if m)
    return line + '|' if text.rstrip().endswith('|') and line.endswith(':') else line


def guess_key(chord_lines):
    tokens = ' '.join(chord_lines).replace('|', ' ').replace(':', ' ').split()
    for token in reversed(tokens):
        root = re.match(r'[A-G][#b]?m?(?!aj)', token)
        if root:
            return root.group(0)
    return '?'


def main():
    if len(sys.argv) != 3:
        sys.exit('사용법: python extract.py IN_DOCX OUT_TXT')
    with zipfile.ZipFile(sys.argv[1]) as z:
        document = z.read('word/document.xml').decode('utf8')

    songs = []  # [title, info, [(kind, text)]]
    for p in re.findall(r'<w:p[ >].*?</w:p>', document, re.S):
        text = text_of(p).strip()
        if not text:
            continue
        kind = kind_of(p, text)
        if kind == 'title':
            songs.append([text, None, []])
        elif not songs:
            continue
        elif kind == 'info':
            songs[-1][1] = text
        else:
            songs[-1][2].append((kind, normalize_chord(text) if kind == 'chord' else text))

    lines, guessed = [], []
    for title, info, body in songs:
        if info is None:
            info = f'Key {guess_key([t for k, t in body if k == "chord"])}'
            guessed.append(f'{title}: {info}')
        lines += [f'# {title}', info]
        for kind, text in body:
            lines += ['', text] if kind == 'section' else [text]
        lines.append('')
    Path(sys.argv[2]).write_text('\n'.join(lines), encoding='utf8')

    print(f'완료: {len(songs)}곡 → {sys.argv[2]}')
    if guessed:
        print(f'[확인 필요] Key를 마지막 코드로 추정한 곡 {len(guessed)}개:')
        for g in guessed:
            print(f'  {g}')


if __name__ == '__main__':
    main()
