@echo off
rem TMDB monitor (read-only, does not affect start.py)
cd /d "%~dp0"

rem Use built-in commands only: System32 may be missing from PATH.
rem "python" may be the Microsoft Store stub, so also check that it really runs.
set "PYTHONW_EXE="
for %%X in (pythonw.exe) do set "PYTHONW_EXE=%%~$PATH:X"
if not defined PYTHONW_EXE goto no_python
python -c "" >nul 2>nul || goto no_python

start "" pythonw monitor_gui.py
exit /b 0

:no_python
"%SystemRoot%\System32\chcp.com" 65001 >nul
echo.
echo 找不到 Python，監控面板無法啟動。
echo.
echo 請安裝 Python 3（建議 3.13），然後再開啟 start.bat。
echo 使用安裝程式時，請務必勾選「Add python.exe to PATH」。
echo.
echo 按任意鍵開啟 Python 下載頁...
pause >nul
start "" https://www.python.org/downloads/
exit /b 1
