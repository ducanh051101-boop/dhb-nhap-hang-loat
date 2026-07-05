@echo off
chcp 65001 >nul
title Nhap / Sua hang loat bai dang - DHB Reup
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py "capnhat.py"
    py "import_bai_dang.py" %*
) else (
    python "capnhat.py"
    python "import_bai_dang.py" %*
)
