@echo off
set SCRIPT_NAME=visualization_GUI.py
set EXE_NAME=generate_visualization_GUI.exe
set FONT_FILENAME=ipaexg.ttf

:: Clean previous build
rmdir /s /q build dist __pycache__ >nul 2>&1
del /q "%SCRIPT_NAME:.py=.spec%" >nul 2>&1

:: Bundle the font manually
set "FONT_ADD=--add-data=%FONT_FILENAME%;."

:: Build EXE
pyinstaller --onefile --windowed %FONT_ADD% "%SCRIPT_NAME%"

echo.

pause
