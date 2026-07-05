@echo off
chcp 65001 >nul
cd /d "%~dp0"
where pythonw >nul 2>nul
if %errorlevel%==0 (
    pythonw "capnhat.py"
    start "" pythonw "gui.py"
) else (
    python "capnhat.py"
    start "" python "gui.py"
)
