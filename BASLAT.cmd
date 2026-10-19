@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Yerel Python ortami bulunamadi. Once README.md kurulum adimlarini uygulayin.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -B -u "panel.py" %*
pause
