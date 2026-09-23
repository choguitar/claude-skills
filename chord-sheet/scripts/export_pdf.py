"""MS Word로 docx를 PDF로 내보낸다 (Windows + Word 필요). Word 렌더링 그대로라 인쇄·배포본과 같다.
곡 제목('악보 제목' 스타일, 개요 수준 1)이 PDF 책갈피로 들어간다.

사용법: python export_pdf.py IN.docx [OUT.pdf]   (OUT 생략 시 같은 이름 .pdf)
"""
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Word COM은 PowerShell로 호출한다. 경로는 작은따옴표 문자열 — ' 는 '' 로 이스케이프
PS = r'''
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
try {
  $doc = $word.Documents.Open('{src}', $false, $true)
  # 17 = PDF. 마지막 인자 1 = 제목(개요 수준)으로 책갈피 생성 → 곡 제목이 PDF 책갈피가 된다
  $doc.ExportAsFixedFormat('{dst}', 17, $false, 0, 0, 1, 1, 0, $true, $true, 1)
  $doc.Close($false)
} finally { $word.Quit() }
'''


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit('사용법: python export_pdf.py IN.docx [OUT.pdf]')
    src = Path(sys.argv[1]).resolve()
    if not src.is_file():
        sys.exit(f'파일 없음: {src}')
    dst = Path(sys.argv[2]).resolve() if len(sys.argv) == 3 else src.with_suffix('.pdf')
    script = PS.replace('{src}', str(src).replace("'", "''")).replace('{dst}', str(dst).replace("'", "''"))
    subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script], check=True)
    print(dst)


if __name__ == '__main__':
    main()
