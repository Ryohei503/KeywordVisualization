@echo off
set SCRIPT_NAME=main.py
set EXE_NAME=JIRA_dataAnalysis_tool.exe
set FONT_FILENAME=src\ipaexg.ttf

:: Clean previous build
rmdir /s /q build dist __pycache__ >nul 2>&1
del /q "%SCRIPT_NAME:.py=.spec%" >nul 2>&1

:: Bundle the font manually
set "FONT_ADD=--add-data=%FONT_FILENAME%;src"

:: Install required packages
pip install -r requirements.txt


:: Find imblearn VERSION.txt location generically
for /f "delims=" %%i in ('python -c "import imblearn, os; print(os.path.join(os.path.dirname(imblearn.__file__), 'VERSION.txt'))"') do set IMBLEARN_VERSION=%%i

:: Build EXE (add other necessary files as needed)
pyinstaller --onefile --windowed --name=%EXE_NAME% %FONT_ADD% ^
  --add-data "%IMBLEARN_VERSION%;imblearn" ^
  --hidden-import=transformers.models.gemma3n ^
  --hidden-import=transformers.models.glm4v ^
  --hidden-import=transformers.models.smollm3 ^
  --hidden-import=transformers.models.t5gemma ^
  "%SCRIPT_NAME%"

echo.
pause
