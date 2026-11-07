"""运行时路径解析模块。

区分开发环境与 PyInstaller 打包环境，返回两类路径基准：
- runtime_root：可写数据目录基准（开发环境=项目根，打包环境=exe 同级）
- bundle_root：只读资源目录基准（开发环境=项目根，打包环境=内嵌资源根）

设计意图：打包后 ``__file__`` 指向 PyInstaller 临时解压目录，若沿用原
``config.py`` 基于 ``__file__`` 的路径计算，业务数据会写入临时目录、重启后
丢失。本模块通过 ``sys.frozen`` 标识将"可写数据目录"与"只读静态资源"解耦：
数据目录始终在 exe 同级（可写、持久），静态资源随包内嵌（只读）。
开发环境下两者退化为同一项目根目录，与原逻辑完全一致（零行为变化）。
"""
import os
import sys
from dataclasses import dataclass


@dataclass(frozen=True)
class RuntimePaths:
    """运行时路径解析结果（不可变值对象）。

    Attributes:
        runtime_root: 可写数据目录基准（开发环境=项目根, 打包环境=exe 同级）。
        bundle_root: 只读资源目录基准（开发环境=项目根, 打包环境=内嵌资源根）。
    """
    runtime_root: str
    bundle_root: str


def resolve_paths() -> RuntimePaths:
    """解析当前运行环境的路径基准，自动区分开发环境与 PyInstaller 打包环境。

    Returns:
        RuntimePaths: 含 runtime_root 与 bundle_root 的不可变值对象。

    - 打包环境（``sys.frozen`` 为真）：runtime_root 取 exe 所在目录，
      bundle_root 取 ``sys._MEIPASS``（单文件模式为临时解压目录，单目录模式
      为 _internal 目录），回退到 exe 同级。
    - 开发环境：runtime_root 与 bundle_root 均取项目根目录（config.py 所在
      目录的上级两级），与原 ``config.py`` 逻辑完全一致。
    """
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        bundle_root = getattr(sys, "_MEIPASS", None) or exe_dir
        return RuntimePaths(runtime_root=exe_dir, bundle_root=bundle_root)
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return RuntimePaths(runtime_root=project_root, bundle_root=project_root)
