@echo off
REM 快速启动脚本 - 启动改进版UI
chcp 65001 >nul
cd /d "%~dp0"

echo 正在启动发票归类工具...
python invoice_renamer.py --ui

if errorlevel 1 (
    echo 启动失败，请检查Python环境是否正确安装
    pause
)
