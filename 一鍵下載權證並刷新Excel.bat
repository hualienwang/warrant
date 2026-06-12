@echo off
setlocal
cd /d "%~dp0"

set "LOG=%~dp0run-refresh-excel.log"
set "BUN_EXE=bun"

echo [%date% %time%] Start > "%LOG%"
echo Working directory: %cd% >> "%LOG%"
echo.
echo Running. Log file:
echo %LOG%
echo.

where bun >nul 2>nul
if errorlevel 1 (
  echo Bun was not found.
  echo [%date% %time%] Bun was not found. >> "%LOG%"
  pause
  exit /b 1
)

echo [%date% %time%] Start download and Excel refresh. >> "%LOG%"
"%BUN_EXE%" "%~dp0download-twse-warrant-all.mjs" --refresh-excel >> "%LOG%" 2>&1
if errorlevel 1 (
  echo.
  echo Failed. See log:
  echo %LOG%
  echo.
  type "%LOG%"
  pause
  exit /b 1
)

echo [%date% %time%] Done. >> "%LOG%"
echo.
echo Done: downloaded latest CSV and refreshed Excel workbook.
echo.
echo Log file:
echo %LOG%
pause
