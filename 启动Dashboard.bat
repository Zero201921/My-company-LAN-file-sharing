@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 AI-Study Dashboard...
start "" http://127.0.0.1:8321
python Dashboard\server.py --port 8321
pause
