@echo off
title Mogal Complaint System

echo.
echo  Mogal Complaint Management System
echo  ===================================
echo.

REM Kill any existing process on port 3000
echo  Freeing port 3000...
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":3000 " ^| findstr "LISTENING"') do (
  taskkill /f /pid %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul

REM Install dependencies if missing
if not exist "node_modules\" (
  echo  Installing dependencies, please wait...
  call npm install
  echo.
)

REM Import Excel data on first run
if not exist "data.db" (
  echo  First run detected - importing customer data from Excel...
  call node --experimental-sqlite scripts/import-excel.js
  echo.
)

echo  Starting server at http://localhost:3000
echo  Press Ctrl+C to stop.
echo.

start "" /b cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:3000"

node --experimental-sqlite server.js
pause
