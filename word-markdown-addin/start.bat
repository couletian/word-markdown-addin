@echo off
chcp 65001 > nul
title Markdown 工具箱 - 开发服务器

echo.
echo ========================================
echo   Markdown 工具箱 · Word 加载项
echo ========================================
echo.

cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Node.js，请先安装：https://nodejs.org/
    echo.
    pause
    exit /b 1
)

echo 正在启动 HTTPS 开发服务器...
echo 启动后请勿关闭此窗口。
echo.

node server.js
pause
