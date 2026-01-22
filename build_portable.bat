@echo off
setlocal

set VENV_DIR=.venv
if not exist %VENV_DIR% (
    python -m venv %VENV_DIR%
)

call %VENV_DIR%\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

pyinstaller --noconfirm --clean --onedir --windowed --name PdfScanCompressor main.py

if not exist dist\PdfScanCompressor (
    echo Build failed.
    exit /b 1
)

echo Build complete: dist\PdfScanCompressor
endlocal
