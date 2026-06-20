#!/bin/bash

# 非遗地址坐标查找工具 - Linux/Mac 一键启动脚本

echo "========================================"
echo "   非遗地址坐标查找工具 - 启动中..."
echo "========================================"
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到Python3，请先安装Python 3.8或更高版本"
    echo "下载地址: https://www.python.org/downloads/"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

# 检查端口5000是否被占用
if lsof -Pi :5000 -sTCP:LISTEN -t &> /dev/null; then
    echo "[警告] 端口5000已被占用，正在尝试释放..."
    lsof -ti :5000 | xargs kill -9 2>/dev/null
    sleep 1
    echo "端口已释放"
    echo ""
fi

# 切换到脚本所在目录
cd "$(dirname "$0")"

# 检查是否已安装依赖
if [ ! -d "venv" ]; then
    echo "首次启动，正在创建虚拟环境并安装依赖..."
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
    echo "依赖安装完成！"
    echo ""
else
    source venv/bin/activate
fi

# 启动Flask应用（后台运行）
echo "正在启动服务..."
nohup python3 app.py > flask.log 2>&1 &
FLASK_PID=$!

# 等待几秒让服务启动
echo "等待服务启动..."
sleep 3

# 检查服务是否启动成功
if ! kill -0 $FLASK_PID 2>/dev/null; then
    echo "[错误] 服务启动失败，请查看 flask.log 了解详情"
    read -p "按回车键退出..."
    exit 1
fi

# 再次检查端口
if ! lsof -Pi :5000 -sTCP:LISTEN -t &> /dev/null; then
    sleep 2
    if ! lsof -Pi :5000 -sTCP:LISTEN -t &> /dev/null; then
        echo "[警告] 服务可能还在启动中，请稍候..."
    fi
fi

echo "服务启动成功！"
echo "访问地址: http://localhost:5000"
echo "进程PID: $FLASK_PID"
echo ""
echo "正在打开浏览器..."

# 自动打开浏览器
if command -v open &> /dev/null; then
    # Mac
    open http://localhost:5000
elif command -v xdg-open &> /dev/null; then
    # Linux
    xdg-open http://localhost:5000
else
    echo "请手动打开浏览器访问: http://localhost:5000"
fi

echo ""
echo "提示："
echo "  - 服务已在后台运行，关闭此终端不会停止服务"
echo "  - 如需停止服务，请运行: ./停止.sh"
echo "  - 或执行: kill $FLASK_PID"
echo ""
read -p "按回车键退出此窗口（服务继续在后台运行）..."
