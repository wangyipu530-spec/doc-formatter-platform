@echo off
chcp 65001 >nul
echo 正在启动文档自动排版助手...
cd /d "%~dp0"
set PYTHONPATH=
"%~dp0venv\Scripts\python" main.py
pause
