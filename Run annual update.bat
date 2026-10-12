@echo off
rem Run annual update.bat
rem
rem Updates the indicators workbook with the latest year of data.
rem Double-click this file, or drag last year's finished workbook onto it.
rem
rem It creates a folder "Indicators update <year>" next to the workbook you
rem choose, holding: 1 Data (source files), 2 Workbook (the new workbook),
rem 3 Web Content (website files) and Run report.txt (what to check).

setlocal
cd /d "%~dp0"
title Indicators annual update

echo.
echo  INDICATORS ANNUAL UPDATE
echo  ========================
echo.
echo  Before continuing:
echo    - close Excel (save any open work first);
echo    - put any manual downloads in the "1 Data" folder of last year's
echo      update (see the "read me" files there). This can also be done
echo      after a first run, and the update run again.
echo.
pause

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo  First run on this computer: setting up. This takes a few minutes.
    py -3 -m venv .venv 2>nul || python -m venv .venv
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo  Python was not found. Install Python 3.11 or later from https://www.python.org
        echo  ^(tick "Add python.exe to PATH" during installation^), then run this again.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo  Installing the required packages failed. Check the internet connection and run this again.
        pause
        exit /b 1
    )
)

set "WORKBOOK=%~1"
if "%WORKBOOK%"=="" (
    echo.
    echo  Choose LAST YEAR's finished indicators workbook in the window that opens.
    for /f "usebackq delims=" %%F in (`powershell -NoProfile -STA -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Title = 'Choose last year''s finished indicators workbook'; $d.Filter = 'Excel workbook (*.xlsx)|*.xlsx'; if ($d.ShowDialog() -eq 'OK') { $d.FileName }"`) do set "WORKBOOK=%%F"
)
if "%WORKBOOK%"=="" (
    echo  No workbook was chosen. Nothing was done.
    pause
    exit /b 1
)
for %%I in ("%WORKBOOK%") do set "WORKBOOK_FOLDER=%%~dpI"

echo.
echo  Starting workbook: %WORKBOOK%
echo  This takes about five minutes. Excel works in the background; leave it alone.
echo.

".venv\Scripts\python.exe" regional-indicators\run_end_to_end.py "%WORKBOOK%" ^
    --test-copy "%TEMP%\indicators-working-copy.xlsx" ^
    --output-root "%WORKBOOK_FOLDER%." ^
    --open-report

echo.
echo  Finished. The run report has opened; the results are in the folder
echo  "Indicators update ..." next to the starting workbook.
pause
