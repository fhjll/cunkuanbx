#Requires -Version 5.1
<#
============================================================
  存保审核工具 - Win7 64位兼容版 本地打包脚本 (PowerShell)

  使用环境: 在 Win10/Win11 电脑上运行本脚本
  前置条件: 已安装 Python 3.8 (64-bit) 并勾选 "py launcher"
            推荐 Python 3.8.10 64-bit:
            https://www.python.org/downloads/release/python-3810/

  运行方式: 在脚本上右键 -> "使用 PowerShell 运行"
            如提示禁止运行脚本，请先执行(只需一次):
              Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
            或用命令行运行:
              powershell -ExecutionPolicy Bypass -File build_win7.ps1

  原理: exe 的 Win7 兼容性取决于打包所用的 Python 版本
        (Python 3.9+ 不再支持 Win7)，与打包机器的系统无关。
        因此 Win11 + Python 3.8 打包出的 exe 可在 Win7 64位运行。
============================================================
#>

# 控制台输出使用 UTF-8，避免中文乱码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 强制 Python 子进程使用 UTF-8，避免中文 Windows 下 pip 用 GBK 解码报
# UnicodeDecodeError（系统区域为 GBK 且用户路径含中文时常见问题）
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:PIP_NO_INPUT = '1'

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

function Fail([string]$msg) {
    Write-Host ""
    Write-Host "[ERROR] $msg" -ForegroundColor Red
    Read-Host "按回车键退出"
    exit 1
}

Write-Host "============================================================"
Write-Host "  存保审核工具 - Win7 64位兼容版 打包"
Write-Host "============================================================"
Write-Host ""

# ---- 1. 检查 Python 3.8 ----
$pyVer = $null
try { $pyVer = (& py -3.8 -V 2>&1) } catch {}
if (-not $pyVer -or $LASTEXITCODE -ne 0) {
    Fail "未找到 Python 3.8。`n请先安装 Python 3.8.10 64-bit:`nhttps://www.python.org/downloads/release/python-3810/`n安装时勾选 'Add python.exe to PATH' 和 'py launcher'。"
}

# ---- 2. 确认是 64-bit Python ----
& py -3.8 -c "import struct,sys; sys.exit(0 if struct.calcsize('P')*8==64 else 1)" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Fail "检测到 Python 3.8 是 32-bit 版本。`nWin7 64位目标需要 64-bit Python，请换装 64-bit 安装包。"
}
Write-Host "[OK] 使用 $pyVer (64-bit)" -ForegroundColor Green

# ---- 3. 创建/复用虚拟环境 venv_win7 ----
if (-not (Test-Path "venv_win7")) {
    Write-Host ""
    Write-Host "[..] 创建虚拟环境 venv_win7 ..."
    & py -3.8 -m venv venv_win7
    if ($LASTEXITCODE -ne 0) { Fail "虚拟环境创建失败。" }
}
$python = Join-Path $PSScriptRoot "venv_win7\Scripts\python.exe"
if (-not (Test-Path $python)) { Fail "未找到虚拟环境 Python: $python" }

# ---- 4. 安装依赖 ----
Write-Host ""
Write-Host "[..] 安装依赖 ..."
& $python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { Fail "pip 自身升级失败，请检查网络后重试。" }
# PyInstaller 5.13.2: 稳定支持 Python 3.8 且产出兼容 Win7 的版本
# --no-cache-dir: 跳过本地缓存，避免缓存中的 GBK 解码问题
& $python -m pip install --no-cache-dir pyinstaller==5.13.2
if ($LASTEXITCODE -ne 0) { Fail "PyInstaller 安装失败，请检查网络后重试。" }
& $python -m pip install --no-cache-dir -r requirements.txt
if ($LASTEXITCODE -ne 0) { Fail "项目依赖安装失败，请检查网络后重试。" }

# ---- 5. PyInstaller 打包 ----
Write-Host ""
Write-Host "[..] 开始打包 ..."
& $python -m PyInstaller --clean --noconfirm --onefile --windowed `
    --name "存保审核工具_Win7_64bit" `
    --add-data "templates;templates" `
    gui.py
if ($LASTEXITCODE -ne 0) { Fail "打包失败，请查看上方日志。" }

# ---- 6. 结果检查 ----
$exe = Join-Path $PSScriptRoot "dist\存保审核工具_Win7_64bit.exe"
Write-Host ""
if (Test-Path $exe) {
    Write-Host "============================================================"
    Write-Host "  打包成功!" -ForegroundColor Green
    Write-Host "  输出文件: dist\存保审核工具_Win7_64bit.exe"
    Write-Host "  可直接拷贝到 Win7 64位 电脑上运行。"
    Write-Host "============================================================"
} else {
    Fail "未找到输出文件，打包可能未成功。"
}
Write-Host ""
Read-Host "按回车键退出"
