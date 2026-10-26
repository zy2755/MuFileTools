# -*- mode: python ; coding: utf-8 -*-
#
# 小睦文件转换器 · MuFileTools 打包配置（onedir 文件夹模式）
# ---------------------------------------------------------------
# 与 MuFileTools.spec（onefile 单文件）的区别：
#   - 产物是一个文件夹 dist_onedir/MuFileTools/，里面 MuFileTools.exe + _internal/
#   - 启动时不需要把整包解压到临时目录，所以启动快很多（约 1 秒 vs 2.7 秒）
#   - 代价：不能只发一个 exe，要整个文件夹一起拷
#
# 排除项与 onefile 版保持一致。
# ---------------------------------------------------------------

from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files('pymupdf')
# 同 onefile 版：app_icon.ico 也要打进包里，否则运行时窗口图标会退回占位图
datas += [('app_icon.ico', '.')]

EXCLUDE_FILES = {
    'opengl32sw.dll',
    'qdirect2d.dll',
    'qtiff.dll',
}

EXCLUDE_DIRS = {
    'translations',
    'tls',
}

EXCLUDE_MODULES = [
    'tkinter',
    'unittest',
    'pydoc',
    'doctest',
    'test',
    'distutils',
    'lib2to3',
    'ensurepip',
    'PySide6.QtQml',
    'PySide6.QtQuick',
    'PySide6.QtQuickWidgets',
    'PySide6.QtWebEngineCore',
    'PySide6.QtWebEngineWidgets',
    'PySide6.QtMultimedia',
    'PySide6.QtMultimediaWidgets',
    'PySide6.QtSql',
    'PySide6.QtTest',
    'PySide6.QtPdf',
    'PySide6.QtOpenGL',
    'PySide6.QtOpenGLWidgets',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['win32com.client', 'pythoncom', 'PIL.ImageDraw'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDE_MODULES,
    noarchive=False,
    optimize=1,
)


def _keep(entry):
    name = entry[0].replace('\\', '/').lower()
    base = name.rsplit('/', 1)[-1]
    if base in EXCLUDE_FILES:
        return False
    parts = name.split('/')
    if len(parts) >= 2 and parts[-2] in EXCLUDE_DIRS:
        return False
    return True


a.binaries = [x for x in a.binaries if _keep(x)]
a.datas = [x for x in a.datas if _keep(x)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MuFileTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MuFileTools',
)
