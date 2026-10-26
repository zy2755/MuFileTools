@echo off
cd /d "%~dp0"
setlocal

rem ---------------------------------------------------
rem  小睦文件转换器启动脚本（优先使用项目自带的 .venv 环境）
rem ---------------------------------------------------
set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

if not defined PY (
  echo [小睦文件转换器] 未找到 .venv 虚拟环境。
  echo [小睦文件转换器] 请先双击 install.bat 创建环境，或执行:
  echo [小睦文件转换器]     python -m pip install -r requirements.txt
  pause
  exit /b 1
)

"%PY%" main.py %*
if errorlevel 1 (
  echo.
  echo [小睦文件转换器] 启动出错，请把上面的报错信息截图反馈。
  pause
)
endlocal