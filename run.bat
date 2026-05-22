@echo off
rem run.bat — activate .venv if present, otherwise use .venv python, then run script
if exist ".venv\Scripts\activate.bat" (
  call ".venv\Scripts\activate.bat"
  python lenovo_parser.py %*
) else if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" lenovo_parser.py %*
) else (
  echo Virtual environment not found.
  echo Create one with: python -m venv .venv
  echo Then install deps: .venv\Scripts\python.exe -m pip install -r requirements.txt
  pause
)