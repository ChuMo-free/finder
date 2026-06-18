@echo off
chcp 65001 >nul
title 非遗地址坐标查找工具

echo ========================================
echo    非遗地址坐标查找工具 - 启动中...
echo ========================================
echo.

cd /d "%~dp0"

echo [1/3] 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未检测到Python，请先安装Python 3.8+
    pause
    exit /b 1
)
echo ✅ Python环境正常

echo.
echo [2/3] 启动服务...
start "" python app.py

echo.
echo [3/3] 等待服务启动...
timeout /t 3 /nobreak >nul

echo.
echo ✅ 服务启动成功！
echo 🌐 正在打开浏览器...
echo.
echo 访问地址：http://localhost:5000
echo.
echo ========================================
echo   提示：请勿关闭此窗口，否则网页会打不开
echo   关闭工具：直接关闭此窗口即可
echo ========================================
echo.

start http://localhost:5000

pause
