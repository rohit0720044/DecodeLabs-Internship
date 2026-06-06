@echo off
title DecodeLabs Assignment Viewer
cd /d "%~dp0"
set "BUNDLED_PY=C:\Users\rks07\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

echo.
echo ================================================
echo   DecodeLabs Assignment Viewer
echo ================================================
echo.
echo Keep this window open while using the website.
echo The app will print the local URL below.
echo.

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" app.py
  goto done
)

py -3 app.py
if errorlevel 1 (
  echo.
  echo Python launcher failed. Trying plain python...
  python app.py
)
:done
pause
