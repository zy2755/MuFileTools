@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

rem ============================================================
rem  小睦文件转换器 MuFileTools 打包脚本
rem  优先使用项目自带的 .venv，避免用到系统里另一个 Python
rem  （多个 Python 并存时，依赖很容易装错解释器）
rem ============================================================

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo === 小睦文件转换器 MuFileTools 打包 ===
echo 使用解释器：%PY%
echo.

"%PY%" -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
  echo 未检测到 PyInstaller，正在安装...
  "%PY%" -m pip install pyinstaller
  if errorlevel 1 goto :err
)

echo [1] 单文件模式   产物 dist\MuFileTools.exe          启动约 2.7 秒
echo [2] 文件夹模式   产物 dist_onedir\MuFileTools\      启动约 1.3 秒
echo.
choice /c 12 /n /m "请选择打包方式 [1=单文件 / 2=文件夹]: "
if errorlevel 2 goto :onedir

echo.
echo 正在打包（单文件模式）...
"%PY%" -m PyInstaller --noconfirm --clean MuFileTools.spec
if errorlevel 1 goto :err
echo.
echo 打包完成： dist\MuFileTools.exe
echo 这个文件可以直接发给别人，双击即用。
goto :done

:onedir
echo.
echo 正在打包（文件夹模式）...
"%PY%" -m PyInstaller --noconfirm --clean --distpath dist_onedir --workpath build_onedir MuFileTools_onedir.spec
if errorlevel 1 goto :err
echo.
echo 打包完成： dist_onedir\MuFileTools\MuFileTools.exe
echo 启动快一倍，但整个文件夹要一起拷贝。

:done
echo.
pause
exit /b 0

:err
echo.
echo 打包失败，请查看上方报错。
pause
exit /b 1
