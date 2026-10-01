@echo off
chcp 65001 >nul
title Media Report Builder
cd /d "%~dp0"
echo.
echo   Dang khoi dong Media Report Builder...
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo   [LOI] Khong tim thay python trong PATH.
  echo   Cai Python 3.10+ roi chay lai file nay.
  pause
  exit /b 1
)
python -m app.server
echo.
echo   App da tat.
pause
