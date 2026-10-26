@echo off
cd /d "%~dp0"
setlocal enabledelayedexpansion

echo ============================================
echo  小睦文件转换器 环境安装（创建 .venv 并安装依赖）
echo ============================================

rem ---- 找一个带 venv 模块的 Python（优先 3.13，可直接复用 wheels 里的安装包）----
set "BASE="
for %%P in ("D:\Python3.13.15\python.exe" "python" "py -3.13" "py -3" "%LOCALAPPDATA%\Programs\Python\Python313\python.exe") do (
  if not defined BASE (
    %%~P -c "import venv,ensurepip" >nul 2>nul && set "BASE=%%~P"
  )
)

if not defined BASE (
  echo [错误] 没有找到可用的 Python 3.10+，请先安装 Python 后再运行本脚本。
  pause
  exit /b 1
)
echo [1/3] 使用解释器: !BASE!

if exist ".venv\Scripts\python.exe" (
  echo [2/3] .venv 已存在，跳过创建（如需重建请先删除 .venv 文件夹）
) else (
  echo [2/3] 创建虚拟环境 .venv ...
  "!BASE!" -m venv .venv
  if errorlevel 1 goto err
)

set "PIP_USER=0"
echo [3/3] 安装依赖...
if exist "wheels\pymupdf-2.0-cp311-abi3-win_amd64.whl" (
  .venv\Scripts\python.exe -m pip install --disable-pip-version-check -q --no-index --find-links wheels --no-deps pymupdf networkx numpy openpyxl et-xmlfile pywin32 shiboken6 PySide6-Essentials pillow python-docx lxml typing-extensions
) else (
  echo 本地 wheels 目录缺失，改用网络安装（需要联网，请耐心等待）...
  .venv\Scripts\python.exe -m pip install --disable-pip-version-check -q -r requirements.txt
)
if errorlevel 1 goto err

.venv\Scripts\python.exe -c "import pymupdf, openpyxl, win32com.client, docx, PIL; from PySide6.QtWidgets import QApplication; print('依赖校验通过')"
if errorlevel 1 goto err

echo.
echo 安装完成！以后双击 run.bat 启动即可。
pause
exit /b 0

:err
echo [错误] 安装失败，请把上方报错截图反馈。
pause
exit /b 1
