@echo off
chcp 65001 >nul
title 非遗地址查找工具 - 运行中
echo ========================================
echo    非遗地址坐标查找工具
echo ========================================
echo.

:: 切换到脚本所在目录
cd /d "%~dp0"

:: ========================================
:: 步骤1：检测Python环境
:: ========================================
echo [1/8] 正在检测Python环境...

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [错误] 未检测到Python环境！
    echo.
    echo 请先安装 Python 3.8 或更高版本：
    echo   1. 访问 https://www.python.org/downloads/
    echo   2. 下载并安装最新版 Python
    echo   3. 安装时请勾选 "Add Python to PATH"
    echo.
    echo 安装完成后请重新运行此脚本
    echo.
    pause
    exit /b 1
)

:: 检查Python版本
for /f "tokens=2 delims= " %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo       Python版本: %PYTHON_VERSION%
echo       ✓ Python环境检测通过
echo.

:: ========================================
:: 步骤2：检查端口占用
:: ========================================
echo [2/8] 正在检查端口5000...

netstat -ano | findstr ":5000 " >nul 2>&1
if %errorlevel% equ 0 (
    echo       [警告] 端口5000可能已被占用
    echo       正在尝试启动服务，如果失败请关闭占用端口的程序
    echo.
) else (
    echo       ✓ 端口5000可用
    echo.
)

:: ========================================
:: 步骤3：创建虚拟环境（首次运行）
:: ========================================
echo [3/8] 正在检查运行环境...

if not exist "venv" (
    echo       首次运行，正在创建虚拟环境...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [错误] 创建虚拟环境失败
        pause
        exit /b 1
    )
    echo       ✓ 虚拟环境创建成功
) else (
    echo       ✓ 虚拟环境已存在
)
echo.

:: ========================================
:: 步骤4：安装依赖（首次运行）
:: ========================================
echo [4/8] 正在检查依赖包...

call venv\Scripts\activate.bat

:: 检查是否已安装依赖（通过检查是否存在flask包）
python -c "import flask" >nul 2>&1
if %errorlevel% neq 0 (
    echo       正在安装依赖包，请稍候...
    echo       （首次运行可能需要1-3分钟）
    echo.
    pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    if %errorlevel% neq 0 (
        echo.
        echo [错误] 依赖安装失败，请检查网络连接
        echo 或尝试手动执行: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    echo.
    echo       ✓ 依赖安装完成
) else (
    echo       ✓ 依赖包已安装
)
echo.

:: ========================================
:: 步骤5：启动Flask服务
:: ========================================
echo [5/8] 正在启动服务...

:: 检查是否已有服务在运行
netstat -ano | findstr ":5000 " >nul 2>&1
if %errorlevel% equ 0 (
    echo       检测到服务可能已在运行，尝试直接打开浏览器...
    goto open_browser
)

:: 最小化启动Flask服务
start /min "非遗地址查找工具" python app.py
echo       服务启动中，请稍候...

:: ========================================
:: 步骤6：等待服务启动
:: ========================================
echo.
echo [6/8] 正在等待服务启动...

set /a retry_count=0
:wait_loop
timeout /t 1 /nobreak >nul
set /a retry_count+=1

netstat -ano | findstr ":5000 " >nul 2>&1
if %errorlevel% equ 0 (
    echo       ✓ 服务启动成功
    goto open_browser
)

if %retry_count% geq 15 (
    echo.
    echo [错误] 服务启动超时（15秒）
    echo 请检查是否有错误信息，或端口被占用
    echo.
    pause
    exit /b 1
)

goto wait_loop

:: ========================================
:: 步骤7：打开浏览器
:: ========================================
:open_browser
echo.
echo [7/8] 正在打开浏览器...

start http://localhost:5000
echo       ✓ 浏览器已打开
echo.

:: ========================================
:: 步骤8：完成提示
:: ========================================
echo [8/8] 启动完成！
echo.
echo ========================================
echo    工具已成功启动！
echo ========================================
echo.
echo 访问地址: http://localhost:5000
echo.
echo 重要提示：
echo   - 请勿关闭此窗口，关闭会停止服务
echo   - 如需停止服务，请运行 "停止服务.bat"
echo   - 浏览器窗口可能需要手动切换到前台
echo.
echo 按任意键最小化此窗口（服务继续运行）...
pause >nul

:: 最小化窗口
echo 窗口已最小化，服务继续在后台运行...
exit /b 0
