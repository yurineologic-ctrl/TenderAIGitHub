@echo off
chcp 65001 >nul
echo ============================================================
echo   ASUS Notebooks Parser
echo   Source  - asus.com/ua-ua/store/laptops/
echo   Output  - Desktop\NOUT_Asus.xlsx
echo ============================================================
echo.

if exist ".venv\Scripts\activate.bat" (
  call ".venv\Scripts\activate.bat"
  python asus_parser.py %*
) else if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" asus_parser.py %*
) else (
  echo ERROR - Virtual environment not found.
  echo Create it with  python -m venv .venv
  echo Then install    .venv\Scripts\python.exe -m pip install -r requirements.txt
  echo And             .venv\Scripts\playwright.exe install chromium
  echo.
  pause
  exit /b 1
)

echo.
pause
