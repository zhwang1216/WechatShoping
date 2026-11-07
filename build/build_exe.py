"""微信小店对账系统打包构建脚本。

封装 PyInstaller 打包流程：环境校验 -> 清理旧产物 -> 调用 PyInstaller ->
产物体积校验 -> 输出结果。

前置条件：
- Windows 64 位系统
- 已安装 Python 3.x 与 PyInstaller（pip install pyinstaller）
- 项目根目录存在 main.py、app/、static/、requirements.txt

用法：
    python build/build_exe.py                      # 默认单目录模式
    python build/build_exe.py --mode onefile       # 单文件模式
    python build/build_exe.py --no-console         # 无控制台窗口
    python build/build_exe.py --clean              # 构建前强制清理
"""
import argparse
import os
import shutil
import subprocess
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPEC_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "package_to_exe.spec")
DIST_DIR = os.path.join(PROJECT_ROOT, "dist")
BUILD_DIR = os.path.join(PROJECT_ROOT, "build")
MAX_SIZE_MB = 100
REQUIRED_FILES = ["main.py", "app", "static", "requirements.txt"]


def check_environment():
    """校验 PyInstaller 已安装且项目源码完整，缺失时输出指引并退出码 1。"""
    pyinstaller_bin = shutil.which("pyinstaller")
    if pyinstaller_bin is None:
        try:
            import PyInstaller  # noqa: F401
        except ImportError:
            print("[环境检查] 未检测到 PyInstaller。")
            print("请先安装：pip install pyinstaller")
            sys.exit(1)
    for name in REQUIRED_FILES:
        path = os.path.join(PROJECT_ROOT, name)
        if not os.path.exists(path):
            print(f"[环境检查] 缺少必需文件/目录：{path}")
            sys.exit(1)
    print("[环境检查] 通过")


def clean_old_artifacts():
    """清理 PyInstaller 中间产物与 dist/ 旧产物。"""
    for name in ("dist",):
        target = os.path.join(PROJECT_ROOT, name)
        if os.path.isdir(target):
            shutil.rmtree(target)
            print(f"[清理] 已删除 {target}")
    for entry in os.listdir(BUILD_DIR):
        if entry.endswith(".spec") or entry.endswith(".py") or entry.endswith(".md"):
            continue
        target = os.path.join(BUILD_DIR, entry)
        if os.path.isdir(target):
            shutil.rmtree(target)
        else:
            os.remove(target)
        print(f"[清理] 已删除 {target}")


def run_pyinstaller(mode, console):
    """调用 pyinstaller 命令执行打包。"""
    cmd = ["pyinstaller", SPEC_FILE]
    if mode == "onefile":
        cmd.append("--onefile")
    else:
        cmd.append("--onedir")
    if console:
        cmd.append("--console")
    else:
        cmd.append("--noconsole")
    cmd.extend(["--distpath", DIST_DIR, "--workpath",
                os.path.join(BUILD_DIR, "pyinstaller_work")])
    print(f"[构建] 执行：{' '.join(cmd)}")
    try:
        subprocess.run(cmd, check=True, cwd=PROJECT_ROOT)
    except subprocess.CalledProcessError as exc:
        print(f"[构建] PyInstaller 执行失败（退出码 {exc.returncode}）。")
        print("请查看上方日志定位错误。常见原因：依赖缺失、路径含中文等。")
        sys.exit(1)
    except FileNotFoundError:
        print("[构建] 未找到 pyinstaller 命令，请确认已安装并加入 PATH。")
        print("安装指引：pip install pyinstaller")
        sys.exit(1)


def validate_output(dist_dir):
    """校验产物体积并输出结果。"""
    if not os.path.isdir(dist_dir):
        print(f"[校验] 产物目录不存在：{dist_dir}")
        sys.exit(1)
    total_bytes = 0
    for root, _dirs, files in os.walk(dist_dir):
        for fname in files:
            total_bytes += os.path.getsize(os.path.join(root, fname))
    total_mb = total_bytes / (1024 * 1024)
    print(f"[校验] 产物总体积：{total_mb:.1f} MB")
    if total_mb > MAX_SIZE_MB:
        print(f"[告警] 产物体积超过 {MAX_SIZE_MB} MB 上限（当前 {total_mb:.1f} MB）。")
        print("请检查是否引入了大型第三方库。此告警不阻断构建。")
    else:
        print(f"[校验] 体积符合 ≤{MAX_SIZE_MB} MB 约束。")
    print(f"\n构建完成！产物路径：{dist_dir}")
    print("部署步骤：将 dist/ 目录整体复制到目标 Windows 机器，双击 exe 运行。")


def main():
    parser = argparse.ArgumentParser(description="微信小店对账系统打包构建脚本")
    parser.add_argument("--mode", choices=["onedir", "onefile"], default="onedir",
                        help="打包模式（默认 onedir）")
    console_group = parser.add_mutually_exclusive_group()
    console_group.add_argument("--console", dest="console", action="store_true",
                               default=True, help="保留控制台窗口（默认）")
    console_group.add_argument("--no-console", dest="console", action="store_false",
                               help="不显示控制台窗口")
    parser.add_argument("--clean", action="store_true", default=True,
                        help="构建前清理旧产物（默认开启）")
    args = parser.parse_args()

    check_environment()
    if args.clean:
        clean_old_artifacts()
    run_pyinstaller(args.mode, args.console)
    product_dir = os.path.join(DIST_DIR, "微信小店对账系统")
    validate_output(product_dir if os.path.isdir(product_dir) else DIST_DIR)


if __name__ == "__main__":
    main()
