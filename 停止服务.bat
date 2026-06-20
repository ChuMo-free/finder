@echo off
chcp 65001 >nul
title 停止非遗地址查找工具

echo ========================================
echo    停止非遗地址查找工具
echo ========================================
echo.

echo 正在查找运行中的服务进程...

:: 查找占用5000端口的进程
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000 "') do (
    set PID=%%a
)

if "%PID%"=="" (
    echo.
    echo [提示] 未检测到运行中的服务
    echo 可能服务已经停止，或未启动
    echo.
    pause
    exit /b 0
)

echo.
echo 找到服务进程，PID: %PID%
echo.
echo 正在停止服务...

:: 终止进程
taskkill /F /PID %PID% >nul 2>&1

if %errorlevel% equ 0 (
    echo.
    echo ✓ 服务已成功停止
    echo.
) else (
    echo.
    echo [警告] 停止服务可能失败
    echo 请手动检查是否还有进程在运行
    echo.
)

echo 按任意键退出...
pause >nul
exit /b 0
