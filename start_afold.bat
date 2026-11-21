@echo off
setlocal
cd /d "%~dp0"
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo Python 3 not found.
  pause
  exit /b 1
)
%PY% -c "import tkinter" >nul 2>nul
if errorlevel 1 (
  echo Tkinter is not available in this Python installation.
  pause
  exit /b 2
)
for /f "delims=" %%P in ('%PY% -c "import sys; print(sys.executable.replace(chr(92),chr(47)))"') do set "PYEXE=%%P"
if exist "%PYEXE:python.exe=pythonw.exe%" (
  start "AFold OS" /b "%PYEXE:python.exe=pythonw.exe%" "%~dp0afold_os.py"
) else (
  start "AFold OS" /b %PY% "%~dp0afold_os.py"
)
endlocal
