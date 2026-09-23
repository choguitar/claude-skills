"""chord-sheet 스킬에 필요한 도구가 이 PC에 있는지 확인한다. 설치는 하지 않는다 — install_deps.py 참고.

사용법: python check_env.py
출력: 빠진 항목은 '[없음] <이름>' — install_deps.py에 그 이름을 넘기면 설치된다(word 제외).
종료 코드: 필수 도구가 빠졌으면 1
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

FONT_DIRS = (Path(os.environ.get('WINDIR', r'C:\Windows')) / 'Fonts',
             Path(os.environ.get('LOCALAPPDATA', '')) / 'Microsoft' / 'Windows' / 'Fonts')


def has_font():
    return any(d.is_dir() and any('d2coding' in f.name.lower() for f in d.iterdir()) for d in FONT_DIRS)


def has_word():
    """Word를 띄우지 않고 COM 등록 여부만 본다."""
    if not shutil.which('powershell.exe'):
        return False
    probe = "if ([type]::GetTypeFromProgID('Word.Application')) { exit 0 } else { exit 1 }"
    return subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', probe],
                          capture_output=True).returncode == 0


def main():
    missing = False
    if sys.version_info >= (3, 8):
        print(f'[OK]   python {sys.version.split()[0]}')
    else:
        print(f'[없음] python — 3.8 이상 필요(현재 {sys.version.split()[0]}). install_python.ps1 / .sh')
        missing = True

    if has_font():
        print('[OK]   font (D2Coding)')
    else:
        print('[없음] font — D2Coding 폰트. 없으면 Word가 다른 폰트로 그려 정렬이 모두 어긋남')
        missing = True

    if has_word():
        print('[OK]   word')
    else:
        print('[없음] word — MS Word. 배치 검사·PDF에 필요(docx 생성만은 없어도 됨). 자동 설치 불가 — Microsoft Office를 직접 설치')
        missing = True

    if shutil.which('pdftoppm'):
        print('[OK]   poppler (pdftoppm)')
    else:
        print('[선택] poppler — pdftoppm 없음. PDF 미리보기 이미지 단계만 생략됨')

    sys.exit(1 if missing else 0)


if __name__ == '__main__':
    main()
