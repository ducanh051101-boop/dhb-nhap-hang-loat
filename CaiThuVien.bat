@echo off
chcp 65001 >nul
title Cai dat cong cu DHB - chay lan dau tren may moi
echo ============================================================
echo   CAI DAT (chi lam 1 lan tren may moi)
echo ============================================================
echo.

REM --- 1) Kiem tra Python ---
where python >nul 2>nul
if %errorlevel%==0 goto HAVEPY

echo [!] May nay CHUA co Python 3.
echo.
where winget >nul 2>nul
if %errorlevel%==0 (
  echo Dang thu TU DONG cai Python bang winget, vui long cho...
  winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
  echo.
  echo ^>^> Da chay lenh cai Python.
  echo    HAY DONG cua so nay, roi chay lai "CaiThuVien.bat" mot lan nua
  echo    ^(de Windows nhan Python vua cai^).
  echo.
  pause
  exit /b
)

echo Khong co winget de tu cai. Hay cai Python thu cong:
echo    1^) Trang tai Python se tu mo ra sau khi bam phim.
echo    2^) Tai Python 3, khi cai NHO TICK vao o "Add Python to PATH".
echo    3^) Cai xong, chay lai file "CaiThuVien.bat" nay.
echo.
pause
start "" https://www.python.org/downloads/
exit /b

:HAVEPY
python --version
echo.
echo Dang cai thu vien can thiet (openpyxl, customtkinter)...
python -m pip install --upgrade openpyxl customtkinter
echo.
echo ============================================================
echo   XONG! Gio nhay dup "Chay_GUI.bat" de bat dau dung.
echo ============================================================
pause
