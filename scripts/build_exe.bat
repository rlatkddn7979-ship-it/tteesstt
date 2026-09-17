@echo off
REM Build the tool bundle into a single Windows .exe (PyInstaller).
REM Run this from inside the scripts folder (double-click works too).
REM NOTE: kept ASCII-only on purpose - Korean text in a .bat file can get
REM mangled by cmd.exe's codepage handling and break the commands below.

pip install -r requirements.txt
if errorlevel 1 (
    echo pip install failed. See the error above.
    pause
    exit /b 1
)

pyinstaller --onefile --noconsole --name doowon_product_lookup launcher.py
if errorlevel 1 (
    echo PyInstaller build failed. See the error above.
    pause
    exit /b 1
)

echo.
echo Build complete: dist\doowon_product_lookup.exe
echo Copy this single .exe anywhere (or to another PC) and run it directly.
pause
