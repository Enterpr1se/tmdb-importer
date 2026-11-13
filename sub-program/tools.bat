@echo off
rem TMDB sub-program tools panel
cd /d "%~dp0"

rem Use built-in commands only: System32 may be missing from PATH.
set "PYTHONW_EXE="
for %%X in (pythonw.exe) do set "PYTHONW_EXE=%%~$PATH:X"
if not defined PYTHONW_EXE goto no_python
python -c "" >nul 2>nul || goto no_python

start "" pythonw tools_gui.py
exit /b 0

:no_python
"%SystemRoot%\System32\chcp.com" 65001 >nul
echo.
echo 找不到 Python，小工具面板無法啟動。
echo 請先按主程式 start.bat 的提示安裝 Python 3。
echo.
pause
exit /b 1
