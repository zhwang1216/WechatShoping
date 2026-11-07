# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：微信小店对账系统 -> Windows 可执行程序。

由 build/build_exe.py 调用，或直接执行：
    pyinstaller build/package_to_exe.spec --onedir --console

配置要点（对应 spec.md 6.1 数据约束）：
- 入口：main.py（规则 1）
- 平台：win_amd64（规则 2）
- 模式：默认单目录 onedir（规则 3），可通过 --onefile 覆盖
- 资源：static/ 内嵌（规则 4）
- 排除：data/backups/logs/uploads/tests 等运行时与开发产物（规则 5）
- 控制台：默认保留（规则 6）
"""
import os

block_cipher = None
project_root = os.path.dirname(os.path.dirname(os.path.abspath(SPEC)))

a = Analysis(
    [os.path.join(project_root, 'main.py')],
    pathex=[project_root],
    binaries=[],
    datas=[
        (os.path.join(project_root, 'static'), 'static'),
    ],
    hiddenimports=['openpyxl', 'xlrd'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tests', 'pytest', 'pytest_cache'],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='微信小店对账系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    target_arch='x86_64',
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='微信小店对账系统',
)
