#!/bin/bash
# 非遗地址坐标查找工具 - 一键启动脚本

echo "========================================"
echo "   非遗地址坐标查找工具 - 启动中..."
echo "========================================"
echo ""

# 进入脚本所在目录
cd "$(dirname "$0")"

# 检查Python
echo "[1/3] 检查Python环境..."
if ! command -v python3 &> /dev/null; then
    echo "❌ 未检测到Python3，请先安装Python"
    exit 1
fi
echo "✅ Python环境正常"

echo ""
echo "[2/3] 检查端口..."
# 释放5000端口
lsof -ti:5000 | xargs kill -9 2>/dev/null
sleep 1
echo "✅ 端口已准备好"

echo ""
echo "[3/3] 启动服务..."
echo ""
echo "========================================"
echo "  🌐 访问地址：http://localhost:5000"
echo "  💡 提示：按 Ctrl+C 停止服务"
echo "========================================"
echo ""

# 启动Flask
python3 app.py
