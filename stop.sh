#!/bin/bash
# 微信小店对账与利润核算系统 - 停止脚本
cd "$(dirname "$0")"
PID_FILE="logs/server.pid"

if [ ! -f "$PID_FILE" ]; then
    echo "未找到 PID 文件，服务可能未在运行"
    exit 0
fi

PID=$(cat "$PID_FILE")
if kill -0 "$PID" 2>/dev/null; then
    kill "$PID"
    for i in $(seq 1 20); do
        sleep 0.3
        kill -0 "$PID" 2>/dev/null || break
    done
    if kill -0 "$PID" 2>/dev/null; then
        kill -9 "$PID"
        echo "服务已强制停止 (PID $PID)"
    else
        echo "服务已停止 (PID $PID)"
    fi
else
    echo "PID $PID 已不存在，清理 PID 文件"
fi
rm -f "$PID_FILE"
