@echo off
chcp 65001 >nul
title Cai thu vien cho cong cu DHB - Nhap & Sua hang loat
echo ============================================================
echo   CAI THU VIEN CAN THIET (chi lam 1 lan tren may moi)
echo ============================================================
echo.
echo Dang kiem tra Python...
where python >nul 2>nul
if errorlevel 1 (
  echo [LOI] May nay CHUA co Python.
  echo   -^> Hay cai Python 3 tai: https://www.python.org/downloads/
  echo      QUAN TRONG: khi cai nho TICK vao o "Add Python to PATH".
  echo   Cai xong roi chay lai file nay.
  echo.
  pause
  exit /b
)
python --version
echo.
echo Dang cai thu vien (openpyxl, customtkinter)...
python -m pip install --upgrade openpyxl customtkinter
echo.
echo ============================================================
echo   XONG! Gio nhay dup "Chay_GUI.bat" de bat dau dung.
echo ============================================================
pause
