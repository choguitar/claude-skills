# Python 3.12를 현재 사용자 범위로 설치한다(winget). 사용자 승인 뒤에만 실행한다.
# Python이 없어 다른 스크립트(.py)를 못 돌릴 때 쓰는 유일한 비-Python 단계 — install_python.sh와 내용이 같아야 한다.
# 설치 직후엔 PATH가 현재 창에 반영되지 않으므로, 마지막 줄에 출력되는 python.exe 경로를 직접 쓴다.
# 실행: powershell -NoProfile -ExecutionPolicy Bypass -File install_python.ps1
$ErrorActionPreference = 'Stop'

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
  Write-Error 'winget이 없음 — https://www.python.org/downloads/ 에서 직접 설치'
  exit 1
}
winget install --id Python.Python.3.12 --exact --scope user --silent --accept-package-agreements --accept-source-agreements
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
