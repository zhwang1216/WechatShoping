"""微信小店对账与利润核算系统 - 服务入口"""
import argparse
import sys

from app.server import ReconcileServer, ServerStartupError


def main():
    parser = argparse.ArgumentParser(description="微信小店对账与利润核算系统")
    parser.add_argument("--host", default=None, help="监听地址，默认 127.0.0.1")
    parser.add_argument("--port", type=int, default=None, help="监听端口，默认 8765")
    parser.add_argument("--init-only", action="store_true", help="仅初始化业务数据文件后退出")
    args = parser.parse_args()

    server = ReconcileServer(host=args.host, port=args.port)
    if args.init_only:
        report = server.init_all_store_files()
        if report["store_count"] == 0:
            print("系统尚无店铺，请先正常启动并完成 /setup 初始化")
            return 0
        print(f"业务数据文件初始化完成: 覆盖 {report['store_count']} 个店铺, "
              f"新建 {report['created']} 个, 修复 {report['repaired']} 个")
        print(f"数据目录: {server.stores_root}")
        return 0

    print(f"{server.host}:{server.port} 启动中 ... (Ctrl+C 停止)")
    try:
        server.start()
    except ServerStartupError as e:
        print(f"\n[启动失败] {e}", file=sys.stderr)
        print("服务未启动，原运行中的服务（如有）不受影响。", file=sys.stderr)
        return 1
    except OSError as e:
        print(f"\n[启动失败] 无法在 {server.host}:{server.port} 启动服务: {e}",
              file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
