@echo off
chcp 65001 >nul
echo ==========================================
echo   设备库存管理工具 - EXE 打包脚本
echo ==========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.x
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Python 已检测到

:: 检查并安装 pip
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo [2/4] pip 未安装，正在安装...
    python -m ensurepip --upgrade
    if errorlevel 1 (
        echo [错误] ensurepip 失败，尝试下载 get-pip.py...
        powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py'"
        python get-pip.py
    )
) else (
    echo [2/4] pip 已安装
)

:: 安装 pyinstaller
echo [3/4] 正在安装 PyInstaller...
python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 清华源失败，尝试官方源...
    python -m pip install pyinstaller
)

:: 打包
echo [4/4] 正在打包 EXE...
cd /d "%~dp0"
python -m PyInstaller --onefile --noconsole --name 库存管理 device_inventory_gui.py

if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   打包完成！
echo   EXE 文件位置：dist\库存管理.exe
echo ==========================================
pause
