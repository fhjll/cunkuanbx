@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  存保审核工具 - Win7 64位兼容版 本地打包脚本
REM
REM  使用环境: 在 Win11 电脑上运行本脚本
REM  前置条件: 已安装 Python 3.8 (64-bit) 并勾选 "py launcher"
REM            推荐 Python 3.8.10 64-bit:
REM            https://www.python.org/downloads/release/python-3810/
REM
REM  原理: exe 的 Win7 兼容性取决于打包所用的 Python 版本
REM        (Python 3.9+ 不再支持 Win7)，与打包机器的系统无关。
REM        因此 Win11 + Python 3.8 打包出的 exe 可在 Win7 64位运行。
REM ============================================================

echo ============================================================
echo   存保审核工具 - Win7 64位兼容版 打包
echo ============================================================
echo.

REM ---- 1. 检查 Python 3.8 ----
py -3.8 -V >nul 2>&1
if errorlevel 1 (
    echo [ERROR] 未找到 Python 3.8。
    echo         请先安装 Python 3.8.10 64-bit:
    echo         https://www.python.org/downloads/release/python-3810/
    echo         安装时勾选 "Add python.exe to PATH" 和 "py launcher"。
    pause
    exit /b 1
)

REM ---- 2. 确认是 64-bit Python ----
py -3.8 -c "import struct,sys; sys.exit(0 if struct.calcsize('P')*8==64 else 1)" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] 检测到 Python 3.8 是 32-bit 版本。
    echo         Win7 64位目标需要 64-bit Python，请换装 64-bit 安装包。
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('py -3.8 -V') do echo [OK] 使用 %%v ^(64-bit^)

REM ---- 3. 创建/复用虚拟环境 venv_win7 ----
if not exist venv_win7 (
    echo.
    echo [..] 创建虚拟环境 venv_win7 ...
    py -3.8 -m venv venv_win7
    if errorlevel 1 (
        echo [ERROR] 虚拟环境创建失败。
        pause
        exit /b 1
    )
)
call venv_win7\Scripts\activate.bat

REM ---- 4. 安装依赖 ----
echo.
echo [..] 安装依赖 ...
python -m pip install --upgrade pip
REM PyInstaller 5.13.2: 稳定支持 Python 3.8 且产出兼容 Win7 的版本
pip install pyinstaller==5.13.2
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] 依赖安装失败，请检查网络后重试。
    pause
    exit /b 1
)

REM ---- 5. PyInstaller 打包 ----
echo.
echo [..] 开始打包 ...
pyinstaller --clean --noconfirm --onefile --windowed ^
    --name 存保审核工具_Win7_64bit ^
    --add-data "templates;templates" ^
    gui.py
if errorlevel 1 (
    echo [ERROR] 打包失败，请查看上方日志。
    pause
    exit /b 1
)

REM ---- 6. 结果检查 ----
echo.
if exist "dist\存保审核工具_Win7_64bit.exe" (
    echo ============================================================
    echo   打包成功!
    echo   输出文件: dist\存保审核工具_Win7_64bit.exe
    echo   可直接拷贝到 Win7 64位 电脑上运行。
    echo ============================================================
) else (
    echo [ERROR] 未找到输出文件，打包可能未成功。
)
echo.
pause
endlocal
