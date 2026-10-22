#!/usr/bin/env bash
# =============================================================================
#  在银河麒麟 / 统信UOS 本机打包为单文件程序（最稳妥的打包方式）
# -----------------------------------------------------------------------------
#  为什么建议在本机打包：PyInstaller 不能跨平台编译，而且可执行文件依赖
#  构建机的 glibc 版本；在本机打包出来的程序，一定能在本机运行。
#
#  用法：
#     bash build_linux.sh           打包
#     bash build_linux.sh --deps    先安装打包所需依赖（用户目录）
#
#  产物：dist/存保审核工具   （可直接双击/命令行运行，无需 Python 环境）
#
#  运行依赖（打包机与目标机都需要，通常系统已带）：
#     sudo apt install libxcb-xinerama0 libxcb-icccm4 libxcb-image0 \
#                      libxcb-keysyms1 libxcb-randr0 libxcb-render-util0 \
#                      libxcb-shape0 libxkbcommon-x11-0 libxcb-xkb1 \
#                      fonts-noto-cjk            # Deb 系（麒麟/UOS）
#     sudo yum install libxcb libxkbcommon-x11 wqy-zenhei-fonts   # RPM 系
# =============================================================================
set -euo pipefail
cd "$(dirname "$0")"

PYTHON="${PYTHON:-python3}"

if [ "${1:-}" = "--deps" ]; then
    echo "安装打包依赖到用户目录 ..."
    "$PYTHON" -m pip install --user --upgrade pip wheel
    "$PYTHON" -m pip install --user pyinstaller openpyxl xlrd PyQt5
fi

if ! "$PYTHON" -c "import PyInstaller" >/dev/null 2>&1; then
    echo "缺少 PyInstaller，请先执行：bash build_linux.sh --deps" >&2
    exit 1
fi

echo "开始打包（单文件，内置 templates/ 模板）..."
"$PYTHON" -m PyInstaller --onefile --noconfirm \
    --name 存保审核工具 \
    --add-data "templates:templates" \
    --add-data "bg.jpeg:." \
    gui.py

echo
echo "打包完成：dist/存保审核工具"
echo "可直接运行： ./dist/存保审核工具"
echo "如提示缺少 Qt 平台插件(libqxcb)，按文件头注释安装 xcb 相关库即可。"
