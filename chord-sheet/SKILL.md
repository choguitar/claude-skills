---
name: chord-sheet
description: 코드+가사 악보를 정해진 서식의 docx·PDF로 만든다. 악보 작성, 기존 악보 변환·곡 추가, 악보 PDF 변환을 요청할 때 쓴다.
---

# 코드 악보 작성

밴드 합주·버스킹용 코드 악보를 한 가지 서식으로 만든다. 한 곡은 한 페이지(A4 2단) 안에 담고, 코드는 마디·반 마디 단위로 세로 정렬한다.

악보의 정본은 **docx**다. 사용자가 Word에서 직접 고쳐도 된다. txt는 작업할 때만 쓰는 중간 파일이라, 기존 악보를 고칠 때는 항상 현재 docx에서 txt를 새로 뽑아(`extract.py`) 고친 뒤 다시 빌드한다. 서식 수치(폰트 크기, 자간, 탭 위치, 줄 간격)는 `build.py` 상단 상수에서만 정한다.

## 스크립트

경로는 이 스킬 폴더 기준이다. 모두 Python이라 셸(PowerShell·cmd·bash)과 상관없이 `python <스킬 폴더>/scripts/…`로 실행한다.

| 명령 | 언제 |
|---|---|
| `python scripts/check_env.py` | 세션에서 처음 쓸 때 한 번 — 필요 도구 설치 여부 확인. `python` 자체가 없으면 setup.md |
| `python scripts/install_deps.py 이름…` | 빠진 도구가 있고 **사용자가 설치를 승인했을 때만** |
| `python scripts/extract.py IN.docx OUT.txt` | 기존 악보 docx를 변환하거나 고칠 때 — 텍스트로 뽑기 |
| `python scripts/build.py SONGS.txt OUT.docx [--append-to EXISTING.docx]` | txt가 준비되면 — 악보 docx 만들기. `--append-to`는 기존 악보 뒤에 곡 추가 |
| `python scripts/check_layout.py OUT.docx` | 빌드 직후 — Word 실제 배치로 규칙 위반 검사 |
| `python scripts/export_pdf.py OUT.docx [OUT.pdf]` | 검사 뒤 확인용, 또는 PDF를 달라고 할 때 |

## 참고 문서

| 문서 | 언제 읽나 |
|---|---|
| `references/workflows.md` | **작업 시작 시 항상** — 상황(새 곡 텍스트 / 곡 수정 / 예전 서식 변환 / 제목만)별 진행 순서와 결과 보고 때 알릴 것 |
| `references/format.md` | txt를 쓰거나 고칠 때 — 텍스트 형식, 곡 정보 줄, 섹션명 표준 |
| `references/layout.md` | 빌드 경고·검사 결과를 해석할 때 — 한 페이지 배치 규칙과 검증 절차 |
| `references/setup.md` | `python`이 없거나 `check_env.py`가 `[없음]`·`[선택]`을 냈을 때 — 도구별 역할, 설치 전 확인할 것, 설치 방법 |
