@echo off
set SCRIPT_NAME=main.py
set EXE_NAME=Jira_dataAnalysis_tool.exe
set FONT_FILENAME=src\ipaexg.ttf
for /f "delims=" %%i in ('python -c "import ipadic, os; print(os.path.dirname(ipadic.__file__))"') do set IPADIC_BASE=%%i

:: Clean previous build
rmdir /s /q build dist __pycache__ >nul 2>&1
del /q "%SCRIPT_NAME:.py=.spec%" >nul 2>&1

:: Bundle the font manually
set "FONT_ADD=--add-data=%FONT_FILENAME%;src"
set "IPADIC_ADD=--add-data=%IPADIC_BASE%\*;ipadic"
set "SLOTHLIB_ADD=--add-data=src\slothlib.txt;src"

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
  %IPADIC_ADD% ^
  %SLOTHLIB_ADD% ^
  "%SCRIPT_NAME%"

echo.
pause
