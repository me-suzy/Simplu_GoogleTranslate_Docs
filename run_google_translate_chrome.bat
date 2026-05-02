@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Google Translate Docs - Chrome

set "PYTHON_EXITCODE=0"
for /f %%I in ('powershell.exe -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "LOG_STAMP=%%I"
set "LOG_FILE=run_google_translate_chrome_%LOG_STAMP%.log"

echo ========================================== > "%LOG_FILE%"
echo LOG START: %DATE% %TIME% >> "%LOG_FILE%"
echo ========================================== >> "%LOG_FILE%"

echo ==========================================
echo   GOOGLE TRANSLATE DOCS - CHROME
echo   Log: %LOG_FILE%
echo ==========================================
echo.

set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

set "PYEXE="
where python >nul 2>&1
if not errorlevel 1 set "PYEXE=python"
if not defined PYEXE (
    where py >nul 2>&1
    if not errorlevel 1 set "PYEXE=py -3"
)
if not defined PYEXE (
    echo [EROARE] Nu gasesc Python in PATH.
    set "PYTHON_EXITCODE=1"
    goto AFTER_PY
)

echo [INFO] Folosesc: %PYEXE%
echo [INFO] Folosesc: %PYEXE% >> "%LOG_FILE%"
echo [INFO] Pornesc automat Chrome debug daca portul 9222 nu raspunde.
echo.

if "%PYEXE%"=="python" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& python -u 'google_translate_docs_chrome.py' 2>&1 | Tee-Object -FilePath '%LOG_FILE%' -Append; exit $LASTEXITCODE"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& py -3 -u 'google_translate_docs_chrome.py' 2>&1 | Tee-Object -FilePath '%LOG_FILE%' -Append; exit $LASTEXITCODE"
)
set "PYTHON_EXITCODE=%errorlevel%"

:AFTER_PY
echo. >> "%LOG_FILE%"
echo [LOG] Cod iesire Python: %PYTHON_EXITCODE% >> "%LOG_FILE%"
echo [LOG] Cod iesire Python: %PYTHON_EXITCODE%
echo.
echo Script finalizat.
echo Log salvat in: %LOG_FILE%
echo.
echo Apasa o tasta pentru a inchide...
pause >nul
endlocal
exit /b %PYTHON_EXITCODE%
