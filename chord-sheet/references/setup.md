# 필요 조건

경로는 스킬 폴더 기준이다.

## 확인

이 세션에서 스킬을 처음 쓸 때 한 번 확인한다:

```
python scripts/check_env.py
```

`python` 명령 자체가 없으면(명령을 찾을 수 없다는 오류) 아래 "Python 설치"로 간다.

| 이름 | 도구 | 쓰는 곳 | 없을 때 |
|---|---|---|---|
| `python` | Python 3.8+ | 모든 스크립트(표준 라이브러리만) | 진행 불가 |
| `font` | D2Coding 폰트 | 서식 전체 | 진행 불가 — 정렬이 모두 어긋난다 |
| `word` | MS Word (Windows) | `check_layout.py`, `export_pdf.py` | docx까지만 만들고, 배치 검사·PDF는 못 했다고 알린다 |
| `poppler` | pdftoppm | PDF 미리보기 이미지 | 눈 확인 단계만 건너뛴다 |

## 설치

빠진 도구가 있으면 **설치 전에 사용자에게 한 번 묻는다** — 프로그램 설치는 PC를 바꾸는 일이라서다. 물을 때 알릴 것:
- 설치할 항목과 각각의 용도(위 표)
- 모두 현재 사용자 범위라 관리자 권한은 필요 없음
- `python`·`poppler`는 winget으로 설치하며 winget 패키지 약관에 동의하게 됨
- `word`는 자동 설치가 안 됨 — 유료 프로그램이라 사용자가 직접 설치

### Python 설치

설치 스크립트가 Python이라 Python만은 셸 스크립트로 설치한다. 쓰는 셸에 맞는 것 하나를 실행한다:

```
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/install_python.ps1    # PowerShell·cmd
bash scripts/install_python.sh                                                    # Git Bash 등
```

설치 직후엔 PATH가 현재 셸에 반영되지 않는다. 마지막 줄에 출력된 `python.exe` 경로로 이어서 실행한다(예: `…\Python312\python.exe scripts/check_env.py`). 다음 세션부터는 `python`으로 된다.

### 나머지 도구

빠진 항목 이름만 넘겨 설치하고, 다시 확인한다:

```
python scripts/install_deps.py font poppler
python scripts/check_env.py
```

- `poppler`는 설치 뒤에도 `[선택]`으로 나올 수 있다. PATH 변경은 에이전트·터미널을 다시 시작해야 반영된다.
- 폰트는 이미 열려 있던 Word를 다시 켜야 보인다.
- 폰트 파일은 스킬 `assets/fonts/`에 들어 있다. 라이선스는 같은 폴더 `LICENSE.txt`(SIL OFL 1.1).
- Word는 대체하지 않는다. LibreOffice 등은 줄 높이·폰트 배치가 Word와 달라서 "Word에서 한 곡 한 페이지"를 보장할 수 없다.
