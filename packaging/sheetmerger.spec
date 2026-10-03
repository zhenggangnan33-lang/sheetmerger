# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置：pyinstaller --noconfirm --clean packaging/sheetmerger.spec
# 生成单文件程序 dist/SheetMerger(.exe)；设置环境变量 SHEETMERGER_BUILD_CLI=1 时另外生成命令行版。
import os
import re

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
with open(os.path.join(ROOT, "core", "__init__.py"), encoding="utf-8") as f:
    VERSION = re.search(r'__version__ = "([^"]+)"', f.read()).group(1)

DATAS = [(os.path.join(ROOT, "core", "aliases_default.json"), "core"),
         (os.path.join(ROOT, "gui", "app_icon.png"), "gui")]

# 运行时用不到的库，排除以减小体积
EXCLUDES = [
    "tkinter", "unittest", "pydoc_data", "pytest", "_pytest", "xlwt", "IPython", "matplotlib",
    "scipy", "PIL", "pyarrow", "numexpr", "bottleneck", "sqlalchemy", "openpyxl.tests",
    "PySide6.QtNetwork", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets", "PySide6.QtMultimedia", "PySide6.QtPdf", "PySide6.Qt3DCore",
    "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtSql", "PySide6.QtTest",
]

ICON = os.path.join(SPECPATH, "app.ico")
ICON = ICON if os.path.exists(ICON) else None
VERSION_FILE = os.path.join(SPECPATH, "version_info.txt")
VERSION_FILE = VERSION_FILE if os.path.exists(VERSION_FILE) else None


def build(script, name, console):
    a = Analysis([os.path.join(ROOT, script)], pathex=[ROOT], datas=DATAS,
                 excludes=EXCLUDES, noarchive=False)
    pyz = PYZ(a.pure)
    return EXE(pyz, a.scripts, a.binaries, a.datas, name=name, console=console,
               upx=False, strip=False, debug=False, icon=ICON, version=VERSION_FILE,
               runtime_tmpdir=None, disable_windowed_traceback=False)


build("app.py", "SheetMerger", console=False)
if os.environ.get("SHEETMERGER_BUILD_CLI") == "1":
    build("cli.py", "SheetMerger-cli", console=True)
