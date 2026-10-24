#!/bin/bash
# 微信小店对账与利润核算系统 - macOS 启动脚本
set -e
cd "$(dirname "$0")"

PY=python3
command -v "$PY" >/dev/null 2>&1 || { echo "错误: 未找到 python3，请先安装 Python 3"; exit 1; }

mkdir -p logs
PID_FILE="logs/server.pid"
PORT_FILE="logs/server.port"
PORT="${WECHAT_SHOP_PORT:-8765}"

if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
    OLD_PORT=$(cat "$PORT_FILE" 2>/dev/null || echo "$PORT")
    echo "服务已在运行 (PID $(cat "$PID_FILE")): http://127.0.0.1:$OLD_PORT"
    exit 0
fi

nohup "$PY" main.py --port "$PORT" > logs/server.out 2>&1 &
PID=$!
echo "$PID" > "$PID_FILE"
echo "$PORT" > "$PORT_FILE"

for i in $(seq 1 30); do
    sleep 0.5
    if curl -s "http://127.0.0.1:$PORT/api/health" > /dev/null 2>&1; then
        # 端口可能被旧实例占用并响应健康检查：需确认新进程仍在监听
        if ! kill -0 "$PID" 2>/dev/null; then
            echo "启动失败: 端口 $PORT 已被其他服务实例占用（新进程已退出），最近日志:"
            tail -n 20 logs/server.out
            rm -f "$PID_FILE"
            exit 1
        fi
        echo "服务启动成功: http://127.0.0.1:$PORT (PID $PID)"
        echo "健康检查:    http://127.0.0.1:$PORT/api/health"
        exit 0
    fi
    if ! kill -0 "$PID" 2>/dev/null; then
        echo "服务启动失败，最近日志:"
        tail -n 20 logs/server.out
        rm -f "$PID_FILE"
        exit 1
    fi
done

echo "服务启动超时，请查看 logs/server.out"
exit 1
