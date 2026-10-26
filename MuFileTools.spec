# -*- mode: python ; coding: utf-8 -*-
#
# 小睦文件转换器 · MuFileTools 打包配置（优化版）
# ---------------------------------------------------------------
# 相比默认打包，这里剔除了运行时用不到的 Qt 组件，主要收益：
#   opengl32sw.dll   19.7 MB  软件 OpenGL 兜底（有硬件加速时永不加载）
#   translations/     6.7 MB  99 个语言包（程序未使用 QTranslator，无需任何翻译）
#   tls/              0.6 MB  TLS 网络后端（程序不走网络）
#   qdirect2d.dll     1.0 MB  Direct2D 平台插件（备用，保留 qwindows.dll 即可）
#   qtiff.dll         0.9 MB  TIFF 图片格式（只处理 PNG/JPG）
# ---------------------------------------------------------------
# 注意：不要排除 numpy。PyMuPDF 的 layout/onnx 子模块与 PIL 的可选路径会用到它，
#       虽然本程序未直接 import，但排除后存在运行时崩溃风险。
# ---------------------------------------------------------------

from PyInstaller.utils.hooks import collect_data_files

# pymupdf 的字体等资源
datas = collect_data_files('pymupdf')
# app_icon.ico 必须一起打包：icon= 参数只决定 exe 的「文件图标」，
# 而运行时窗口/任务栏图标是 ui/shell.py 里读这个文件生成的。
# 不打包会退回代码画的蓝色占位图，和文件图标对不上。
datas += [('app_icon.ico', '.')]

# 按文件名剔除（小写比较）
EXCLUDE_FILES = {
    'opengl32sw.dll',    # 软件 OpenGL 兜底，19.7 MB
    'qdirect2d.dll',     # Direct2D 平台插件（备用）
    'qtiff.dll',         # TIFF 格式插件
}

# 按所在目录剔除（目录名小写）
EXCLUDE_DIRS = {
    'translations',      # Qt 语言包，程序没加载 QTranslator
    'tls',               # TLS 后端，程序不走网络
}

# 明确不打包的 Python 模块
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
    """entry = (name, path, typecode)，按包内相对路径判断是否保留。"""
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
    a.binaries,
    a.datas,
    [],
    name='MuFileTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.ico'],
)
