@echo off
REM Builds CleanSheet.exe on THIS Windows PC (free tools only). Run from the folder containing clean_sheet.py and clean_sheet_gui.py.
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller
python -m PyInstaller --noconfirm --windowed --name CleanSheet clean_sheet_gui.py
echo.
echo Done. Your app is in: dist\CleanSheet\CleanSheet.exe  (keep the whole CleanSheet folder together)
pause
