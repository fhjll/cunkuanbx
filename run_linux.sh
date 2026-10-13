#!/usr/bin/env bash
# =============================================================================
#  存款保险申报表审核工具 —— 银河麒麟 / 统信UOS / 通用 Linux 启动脚本
# -----------------------------------------------------------------------------
#  用法：
#     bash run_linux.sh            正常启动
#     bash run_linux.sh --install  缺依赖时自动用 pip 安装（用户目录，无需 root）
#
#  环境要求（任选其一）：
#    A. 系统自带的 PyQt5（推荐，最省事）：
#         sudo apt install python3-pyqt5 python3-openpyxl python3-xlrd   # Deb 系
#         sudo yum install python3-qt5 python3-openpyxl python3-xlrd     # RPM 系
#    B. 用 pip 安装到用户目录（脚本 --install 会做这一步）：
#         python3 -m pip install --user PyQt5 openpyxl xlrd
#
#  中文字体（如界面文字显示为方框，任装其一）：
#    sudo apt install fonts-noto-cjk 或 fonts-wqy-microhei / fonts-wqy-zenhei
# =============================================================================
set -u

cd "$(dirname "$0")" || exit 1

PYTHON="${PYTHON:-python3}"
if ! command -v "$PYTHON" >/dev/null 2>&1; then
    echo "未找到 python3，请先安装：sudo apt install python3" >&2
    exit 1
fi

# ---------- 依赖检查 ----------
missing=""
"$PYTHON" - <<'PY' >/dev/null 2>&1 || missing="PyQt5 openpyxl xlrd"
import importlib
for m in ("PyQt5.QtWidgets", "openpyxl", "xlrd"):
    importlib.import_module(m)
PY

if [ -n "$missing" ]; then
    if [ "${1:-}" = "--install" ]; then
        echo "正在安装依赖（$missing）到用户目录 ..."
        "$PYTHON" -m pip install --user PyQt5 openpyxl xlrd || {
            echo "pip 安装失败。请改用系统包，例如：" >&2
            echo "  sudo apt install python3-pyqt5 python3-openpyxl python3-xlrd" >&2
            exit 1
        }
    else
        echo "缺少依赖：$missing" >&2
        echo "可任选一种方式安装：" >&2
        echo "  1) 系统包（推荐）： sudo apt install python3-pyqt5 python3-openpyxl python3-xlrd" >&2
        echo "  2) 用户目录：      bash run_linux.sh --install" >&2
        exit 1
    fi
fi

# ---------- 启动 ----------
# 无显示环境（如 ssh）时给出明确提示，避免只看到 Qt 报错
if [ -z "${DISPLAY:-}" ] && [ -z "${WAYLAND_DISPLAY:-}" ]; then
    echo "未检测到图形显示环境（DISPLAY/WAYLAND_DISPLAY 为空），无法打开界面。" >&2
    echo "请在图形桌面中运行本脚本。" >&2
    exit 1
fi

exec "$PYTHON" gui.py "$@"
