@echo off
setlocal
cd /d "%~dp0"
set SIMPLU_GT_MAX_FILES=1
call run_google_translate_chrome.bat
endlocal
