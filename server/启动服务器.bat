@echo off
chcp 65001 >nul
title 云岭村 AI智能体服务器 - 端口3456
cd /d "%~dp0"

echo.
echo ========================================
echo   云岭村 AI智能体后端 v2.0
echo ========================================
echo.
echo 检查依赖...
if not exist "node_modules" (
    echo 正在安装依赖，请稍候...
    call npm install
    echo.
)

echo 启动服务器...
echo.
node server.js
echo.
echo 服务器已停止！
pause
