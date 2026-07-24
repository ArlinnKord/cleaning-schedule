@echo off
cd /d "%~dp0"
node index.js
if %errorlevel% neq 0 (
  echo.
  pause
)
