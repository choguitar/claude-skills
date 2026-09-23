"""Word가 실제로 배치한 결과를 읽어 악보 배치 규칙 위반을 찾는다 (Windows + MS Word).

사용법: python check_layout.py IN_DOCX

위반(종료 코드 1): 곡이 한 페이지를 넘음, 코드 줄과 가사가 다른 단으로 갈림.
참고: 섹션이 단 중간에서 잘림 — build가 섹션째로는 한 페이지에 안 들어가 코드+가사 단위로 나눈 곡이면 정상.
"""
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Range.Information: 3 = 페이지 번호, 5 = 페이지 기준 가로 위치 → 페이지 절반 기준으로 단 번호(0/1)
PS = r'''
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$word = New-Object -ComObject Word.Application
try {
  $doc = $word.Documents.Open('{path}', $false, $true)
  $doc.Repaginate()
  $half = $doc.PageSetup.PageWidth / 2
  foreach ($p in $doc.Paragraphs) {
    # 문단 끝 위치로 잰다 — 제목 앞의 단 나누기는 문단 시작을 앞 단에 남긴다
    $e = $p.Range.Duplicate
    $e.MoveEnd(1, -1) | Out-Null
    $e.Collapse(0)
    $column = [int]($e.Information(5) -gt $half)
    '{0}|{1}|{2}|{3}' -f $e.Information(3), $column, $p.Style.NameLocal, $p.Range.Text.Trim()
  }
  $doc.Close($false)
} finally { $word.Quit() }
'''


def read_layout(path):
    script = PS.replace('{path}', str(Path(path).resolve()).replace("'", "''"))
    out = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script],
                         capture_output=True, check=True).stdout.decode('utf8')
    rows = []
    for line in out.splitlines():
        page, x, style, text = (line.lstrip('﻿').split('|', 3) + [''])[:4]
        if page.isdigit():
            rows.append((int(page), int(x), style, text))
    return rows


def main():
    if len(sys.argv) != 2:
        sys.exit('사용법: python check_layout.py IN_DOCX')
    rows = read_layout(sys.argv[1])
    problems, notes, title, section, prev = [], [], None, None, None
    for page, x, style, text in rows:
        if not text.replace('\x0e', '').replace('\x0c', '').strip():  # 단 나누기(\x0e)·빈 공간 문단
            continue
        column = (page, x)
        if style == '악보 제목':
            title, title_page, section = text, page, None
        elif title is None:
            continue
        if page != title_page:
            problems.append(f'한 페이지 넘음: {title} ({title_page}→{page}쪽)')
            title_page = page  # 같은 곡은 한 번만 보고
        if style == '섹션':
            section, section_column = text, column
        elif section and column != section_column:
            notes.append(f'섹션 잘림: {title} {section} (build가 코드+가사 단위로 나눈 곡이면 정상)')
            section = None
        if prev and prev[0] == '코드' and style == '가사' and prev[1] != column:
            problems.append(f'코드·가사 분리: {title} "{text}"')
        prev = (style, column)

    songs = sum(1 for r in rows if r[2] == '악보 제목')
    pages = max((r[0] for r in rows), default=0)
    print(f'{songs}곡, {pages}쪽')
    for n in notes:
        print(f'[참고] {n}')
    for p in dict.fromkeys(problems):
        print(f'[위반] {p}')
    if problems:
        sys.exit(1)
    print('배치 규칙 위반 없음')


if __name__ == '__main__':
    main()
