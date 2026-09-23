"""check_env.py가 [없음]/[선택]으로 보고한 도구를 설치한다. 사용자 승인 뒤에만 실행한다.

사용법: python install_deps.py font poppler   (필요한 것만 나열)
모두 현재 사용자 범위로 설치 — 관리자 권한이 필요 없다.
poppler는 winget을 쓰며, 설치 시 winget 패키지 약관에 동의한다.
python은 여기서 설치할 수 없다(이 스크립트가 Python이라서) — install_python.ps1 / .sh를 쓴다.
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

SKILL_DIR = Path(__file__).resolve().parent.parent
FONT_FILE = SKILL_DIR / 'assets' / 'fonts' / 'D2Coding-Ver1.3.2-20180524-all.ttc'
FONT_REG_KEY = r'HKCU\Software\Microsoft\Windows NT\CurrentVersion\Fonts'


def install_font():
    """사용자 폰트 폴더에 복사하고 현재 사용자 레지스트리에 등록한다(Windows 10 1809+ 방식)."""
    font_dir = Path(os.environ['LOCALAPPDATA']) / 'Microsoft' / 'Windows' / 'Fonts'
    font_dir.mkdir(parents=True, exist_ok=True)
    dest = font_dir / FONT_FILE.name
    shutil.copyfile(FONT_FILE, dest)
    subprocess.run(['reg.exe', 'add', FONT_REG_KEY, '/v', 'D2Coding (TrueType)', '/t', 'REG_SZ', '/d', str(dest), '/f'],
                   check=True, capture_output=True)


def install_poppler():
    if not shutil.which('winget'):
        sys.exit('winget이 없어 설치 불가 — https://github.com/oschwartz10612/poppler-windows/releases 에서 직접 설치')
    subprocess.run(['winget', 'install', '--id', 'oschwartz10612.Poppler', '--exact', '--scope', 'user', '--silent',
                    '--accept-package-agreements', '--accept-source-agreements'], check=True)


INSTALLERS = {
    'font': install_font,
    'poppler': install_poppler,
    'word': lambda: print('Word는 자동 설치 불가 — Microsoft Office를 직접 설치하세요'),
    'python': lambda: print('python은 install_python.ps1 / install_python.sh로 설치합니다'),
}


def main():
    items = sys.argv[1:]
    unknown = [i for i in items if i not in INSTALLERS]
    if not items or unknown:
        sys.exit(f'사용법: python install_deps.py {"|".join(INSTALLERS)} ...' + (f'  (알 수 없는 항목: {unknown})' if unknown else ''))
    for item in items:
        print(f'== {item}')
        INSTALLERS[item]()
    print('\n설치 끝. poppler는 PATH가 바뀌므로 에이전트/터미널을 다시 시작해야 인식될 수 있다.')
    print('폰트는 이미 열려 있던 Word를 다시 켜야 보인다.')


if __name__ == '__main__':
    main()
