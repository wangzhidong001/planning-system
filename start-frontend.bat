@echo off
chcp 65001 >nul
cd /d "%~dp0frontend"
echo Starting frontend at http://localhost:5173 ...
npm run dev
pause
