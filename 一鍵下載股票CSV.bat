@echo off
setlocal
cd /d "%~dp0"

set "LOG=%~dp0run-stocks.log"

echo [%date% %time%] Start > "%LOG%"
echo Working directory: %cd% >> "%LOG%"
echo.
echo Running. Log file:
echo %LOG%
echo.

where bun >nul 2>nul
if not errorlevel 1 (
  set "RUNTIME=bun"
) else (
  where node >nul 2>nul
  if errorlevel 1 (
    echo Bun and Node.js were not found.
    echo [%date% %time%] Bun and Node.js were not found. >> "%LOG%"
    pause
    exit /b 1
  )
  set "RUNTIME=node"
)

echo [%date% %time%] Using %RUNTIME%, start stock download. >> "%LOG%"
"%RUNTIME%" "%~dp0download-twse-stocks-all.mjs" --refresh-excel >> "%LOG%" 2>&1
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
echo Done: downloaded latest stock CSV. Open Excel股票處理.xlsx to refresh.
echo.
echo Log file:
echo %LOG%
pause
