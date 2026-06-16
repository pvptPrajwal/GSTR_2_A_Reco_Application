@echo off
:: ============================================================
::  GSTR-2A Tool — Build Script
::  Run this ONCE on your Windows machine to create the .exe
::  Requirements: Python 3.10+ must be installed
:: ============================================================

echo.
echo  Installing required Python packages...
echo  -----------------------------------------------
pip install pandas openpyxl pyinstaller

echo.
echo  Building .exe — this takes 1-2 minutes...
echo  -----------------------------------------------

pyinstaller ^
  --onefile ^
  --windowed ^
  --name "GSTR2A_Reconciliation_Tool" ^
  --add-data "gstr2a_app.py;." ^
  gstr2a_app.py

echo.
echo  -----------------------------------------------
echo  Done! Your .exe is inside the  dist  folder.
echo  File: dist\GSTR2A_Reconciliation_Tool.exe
echo  -----------------------------------------------
echo  You can copy that .exe to any Windows computer.
echo  No Python needed on that computer.
pause