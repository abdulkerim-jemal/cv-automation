@echo off
setlocal
cd /d "%~dp0"
where python >nul 2>&1 || (echo Python not found.&pause&exit /b 1)
python -c "import streamlit, streamlit_cropper" >nul 2>&1 || (echo Installing core requirements...&python -m pip install -r requirements.txt)
set "TESS=%ProgramFiles%\Tesseract-OCR\tesseract.exe"
if not exist "%TESS%" if exist "%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe" set "TESS=%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"
if exist "%TESS%" (echo Tesseract OK: %TESS%) else (echo WARNING: Tesseract 5.x executable not found - OCR will not work until it is installed.)
where soffice >nul 2>&1
if errorlevel 1 (
  if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" (echo LibreOffice OK.) else (echo WARNING: LibreOffice executable not found - PDF export needs LibreOffice or Microsoft Word.)
) else echo LibreOffice OK.
python -m streamlit run app.py
pause
