#!/bin/bash

# 非遗地址坐标查找工具 - 停止服务脚本

echo "正在停止非遗地址查找工具服务..."

# 查找并杀死占用5000端口的进程
if lsof -Pi :5000 -sTCP:LISTEN -t &> /dev/null; then
    PIDS=$(lsof -ti :5000)
    echo "找到进程: $PIDS"
    kill $PIDS 2>/dev/null
    sleep 1
    
    # 再次检查，如果还在就强制杀死
    if lsof -Pi :5000 -sTCP:LISTEN -t &> /dev/null; then
        echo "强制停止中..."
        lsof -ti :5000 | xargs kill -9 2>/dev/null
    fi
    
    echo "服务已停止"
else
    echo "未检测到运行中的服务"
fi

echo ""
read -p "按回车键退出..."
