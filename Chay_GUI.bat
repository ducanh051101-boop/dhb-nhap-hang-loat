@echo off
chcp 65001 >nul
cd /d "%~dp0"
where pythonw >nul 2>nul
if %errorlevel%==0 goto RUNW
where python >nul 2>nul
if %errorlevel%==0 goto RUNP

echo [!] May nay chua co Python.
echo     Hay chay file "CaiThuVien.bat" truoc (no se cai Python + thu vien),
echo     roi mo lai "Chay_GUI.bat".
echo.
pause
exit /b

:RUNW
pythonw "capnhat.py"
start "" pythonw "gui.py"
exit /b

:RUNP
python "capnhat.py"
start "" python "gui.py"
exit /b
