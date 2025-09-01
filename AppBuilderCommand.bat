@echo off
REM ---- Build PdfFooterAutomation.exe using PyInstaller ----
REM Change path to your Python interpreter if needed

REM 1. Activate virtual environment (if any)
REM call "path\to\venv\Scripts\activate.bat"

REM 2. Ensure PyInstaller installed + dependencies
py -m pip install --upgrade pip
py -m pip install pyinstaller
py -m pip install PyPDF2 reportlab pywin32

REM 3. Navigate to script folder
cd /d "%~dp0"

REM 4. Run PyInstaller to generate EXE
REM --add-data is used for font file (preeti.ttf)
py -m PyInstaller ^
    --onefile ^
    --windowed ^
    --hidden-import=PyPDF2 ^
    --hidden-import=reportlab ^
    --hidden-import=win32com ^
    --add-data "preeti.ttf;." ^
    --icon=logo.ico ^
    PdfFooterAutomation.py

REM 5. Done message
echo.
echo EXE build complete! Check dist folder.

REM 6. (Optional) Build installer using Inno Setup, if iss script is present
if exist "PdfFooterAutomation.iss" (
    echo Running Inno Setup Compiler...
    "C:\Program Files\Inno Setup 6\ISCC.exe" PdfFooterAutomation.iss
)

pause
