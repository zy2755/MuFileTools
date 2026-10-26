@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === 小睦文件转换器 MuFileTools 打包脚本（PyInstaller）===
python -m pip show pyinstaller >nul 2>&1 || python -m pip install pyinstaller
if errorlevel 1 goto :err

python -m PyInstaller --noconfirm --clean --windowed --name MuFileTools ^
  --icon app_icon.ico ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --collect-data pymupdf ^
  main.py
if errorlevel 1 goto :err

echo.
echo 打包完成： dist\MuFileTools\MuFileTools.exe
echo 如需单文件模式，追加参数 --onefile（启动稍慢）。
pause
exit /b 0

:err
echo 打包失败，请查看上方报错。
pause
exit /b 1
