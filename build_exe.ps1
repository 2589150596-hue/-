#Requires -Version 5.1
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  设备库存管理工具 - EXE 打包脚本" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 检查 Python
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "[错误] 未检测到 Python，请先安装 Python 3.x" -ForegroundColor Red
    Write-Host "下载地址：https://www.python.org/downloads/" -ForegroundColor Yellow
    Read-Host "按 Enter 退出"
    exit 1
}
Write-Host "[1/4] Python 已检测到：$(python --version)" -ForegroundColor Green

# 检查 pip
$hasPip = $false
try {
    $pipVer = python -m pip --version 2>$null
    if ($pipVer) { $hasPip = $true }
} catch {}

if (-not $hasPip) {
    Write-Host "[2/4] pip 未安装，正在安装..." -ForegroundColor Yellow
    python -m ensurepip --upgrade
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ensurepip 失败，尝试下载 get-pip.py..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py' -UseBasicParsing
            python get-pip.py
        } catch {
            Write-Host "[错误] 无法安装 pip，请检查网络连接" -ForegroundColor Red
            Read-Host "按 Enter 退出"
            exit 1
        }
    }
} else {
    Write-Host "[2/4] pip 已安装" -ForegroundColor Green
}

# 安装 pyinstaller
Write-Host "[3/4] 正在安装 PyInstaller..." -ForegroundColor Cyan
$mirrors = @(
    "https://pypi.tuna.tsinghua.edu.cn/simple",
    "https://mirrors.aliyun.com/pypi/simple/",
    "https://pypi.org/simple"
)
$installed = $false
foreach ($mirror in $mirrors) {
    python -m pip install pyinstaller -i $mirror --upgrade
    if ($LASTEXITCODE -eq 0) {
        $installed = $true
        break
    }
}
if (-not $installed) {
    Write-Host "[错误] PyInstaller 安装失败，请检查网络" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    exit 1
}

# 打包
Write-Host "[4/4] 正在打包 EXE..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir
python -m PyInstaller --onefile --noconsole --name 库存管理 device_inventory_gui.py

if ($LASTEXITCODE -ne 0) {
    Write-Host "[错误] 打包失败" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    exit 1
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "  打包完成！" -ForegroundColor Green
Write-Host "  EXE 文件位置：dist\库存管理.exe" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Read-Host "按 Enter 退出"
