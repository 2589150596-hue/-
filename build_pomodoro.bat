@echo off
chcp 65001 >nul
echo ==========================================
echo       番茄钟 - 打包工具
echo ==========================================
echo.

REM 检查是否安装了 pyinstaller
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装 pyinstaller...
    pip install pyinstaller
)

echo [1/3] 清理旧文件...
if exist "build\pomodoro" rmdir /s /q "build\pomodoro"
if exist "dist\番茄钟.exe" del /f /q "dist\番茄钟.exe"

echo [2/3] 正在打包...
pyinstaller pomodoro.spec --noconfirm

echo [3/3] 完成！
echo.
if exist "dist\番茄钟.exe" (
    echo [成功] 可执行文件已生成:
    echo        dist\番茄钟.exe
    echo.
    echo 按任意键退出...
) else (
    echo [失败] 打包出错，请检查日志
)
pause >nul
