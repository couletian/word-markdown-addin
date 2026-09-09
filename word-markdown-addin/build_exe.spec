# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置 — 生成单文件 Markdown工具箱.exe

用法：
  pyinstaller --clean --noconfirm build_exe.spec
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

ROOT = Path(SPECPATH).resolve()

# 把 src/、assets/、manifest.xml 全部打进 exe（运行时解压到 sys._MEIPASS）
datas = [
    (str(ROOT / "src"), "src"),
    (str(ROOT / "assets"), "assets"),
    (str(ROOT / "manifest.xml"), "."),
]

hiddenimports = (
    collect_submodules("cryptography")
    + ["ssl", "http.server", "socketserver", "webbrowser", "winreg", "ctypes"]
)

a = Analysis(
    ["launcher.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 排除明显用不到的重型库，压缩体积
    excludes=[
        "tkinter", "unittest", "pydoc", "doctest", "test",
        "numpy", "pandas", "matplotlib", "PIL", "scipy",
        "PyQt5", "PyQt6", "PySide2", "PySide6", "wx",
        "setuptools", "pip", "wheel", "distutils",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Markdown工具箱",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,          # 保留控制台，显示使用指引与运行状态
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / "assets" / "icon.ico"),
)
