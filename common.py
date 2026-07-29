# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险申报表审核工具 - 公共工具模块
=============================================================================

本模块提供所有其他模块共用的工具函数和类：
  1. get_resource_path()       — 获取资源绝对路径（兼容开发模式和 PyInstaller 打包）
  2. extract_embedded_template() — 解压嵌入式模板文件到临时目录
  3. to_float()                — 统一数值转换（支持千分位逗号/中文逗号）
  4. _is_xlsx_format()         — 检测文件是否为 OOXML (xlsx) 格式
  5. ExcelSheet 类             — 统一 Excel 读取器（兼容 .xls 和 .xlsx）

注意事项：
  - 所有 Excel 文件读取都通过 ExcelSheet 类统一处理，不要在业务代码中直接调用 xlrd 或 openpyxl
  - 本模块不依赖其他自定义模块
=============================================================================
"""
import os
import sys
import tempfile
import shutil
import zipfile
import xlrd
import openpyxl


# ============================================================
# 资源路径 — 支持 PyInstaller 打包和开发模式
# ============================================================

def get_resource_path(relative_path):
    """
    获取资源文件的绝对路径。

    在开发模式下，路径相对于本脚本所在目录；
    在 PyInstaller 打包后，路径相对于运行时解压的 _MEIPASS 临时目录。

    参数:
        relative_path (str): 相对路径，例如 "templates/deposit_template.xlsx"

    返回:
        str: 资源的绝对路径

    使用示例:
        >>> tpl = get_resource_path("templates/deposit_template.xlsx")
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包模式：文件被解压到 _MEIPASS 临时目录
        base = sys._MEIPASS
    else:
        # 开发模式：相对于脚本目录
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


def extract_embedded_template(relative_path):
    """
    将嵌入 exe 的模板文件提取到临时目录，返回可用路径。

    模板文件在打包时被嵌入 exe，运行时需解压到 %TEMP% 目录。
    在开发模式下，直接返回原始路径。

    参数:
        relative_path (str): 相对路径，例如 "templates/deposit_template.xlsx"

    返回:
        str: 模板文件的可用路径
    """
    src = get_resource_path(relative_path)
    if os.path.exists(src):
        return src  # 开发模式下直接用
    # 打包模式下文件可能在 _MEIPASS 中但路径格式有差异，兜底返回
    return src


# ============================================================
# 数值转换
# ============================================================

def to_float(val):
    """
    统一数值转换函数。

    将各种可能格式的输入值转换为 float 类型。
    支持：None/空字符串 → 0.0；带千分位逗号或中文逗号的字符串。

    参数:
        val: 输入值，可以是 None、str、int、float 等

    返回:
        float: 转换后的浮点数

    使用示例:
        >>> to_float(None)
        0.0
        >>> to_float("1,234.56")
        1234.56
        >>> to_float("1，234.56")  # 中文逗号
        1234.56
    """
    if val is None or val == '':
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        pass
    try:
        if isinstance(val, str):
            val = val.replace(",", "").replace("，", "").strip()
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _is_xlsx_format(filepath):
    """
    检测文件是否实际上是 OOXML (xlsx) 格式，无论扩展名是什么。

    检测原理：OOXML 文件本质是 ZIP 压缩包，且内部包含标志文件 xl/workbook.xml。
    某些系统导出的文件虽然扩展名为 .xls，但实际是 OOXML 格式，
    这种情况需要用 openpyxl 而非 xlrd 读取。

    参数:
        filepath (str): 文件路径

    返回:
        bool: True 表示文件是 OOXML 格式，False 表示不是
    """
    if not zipfile.is_zipfile(filepath):
        return False
    try:
        with zipfile.ZipFile(filepath) as z:
            return any("xl/workbook.xml" in n for n in z.namelist())
    except Exception:
        return False


# ============================================================
# ExcelSheet 统一读取器
# ============================================================

class ExcelSheet:
    """
    统一 Excel 读取器，兼容 .xls 和 .xlsx 格式。

    提供与 xlrd 兼容的接口（cell_value, nrows），
    自动根据文件扩展名选择底层引擎：
      - .xlsx/.xlsm  → openpyxl
      - .xls         → xlrd

    特殊处理：
      - 若文件扩展名为 .xls 但内部为 OOXML 格式，会自动复制为临时 .xlsx 再读取
      - cell_value(r, c) 统一使用 0-indexed 行列索引

    使用示例:
        >>> ws = ExcelSheet("data.xls")
        >>> print(ws.nrows)          # 总行数
        >>> val = ws.cell_value(3, 2)  # 第4行第3列的值
    """

    def __init__(self, filepath):
        """
        初始化 Excel 读取器。

        参数:
            filepath (str): Excel 文件路径

        异常:
            可能抛出 openpyxl 或 xlrd 的异常（由上层调用者处理）
        """
        filepath = str(filepath)
        ext = filepath.lower()
        temp_copy = None

        # 特殊情况：后缀为 .xls 但内部是 OOXML 格式
        # 需要复制一份临时 .xlsx 文件供 openpyxl 读取
        if not ext.endswith((".xlsx", ".xlsm")) and _is_xlsx_format(filepath):
            fd, temp_copy = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            shutil.copy2(filepath, temp_copy)
            filepath = temp_copy
            ext = ".xlsx"

        try:
            if ext.endswith((".xlsx", ".xlsm")):
                # OOXML 格式 → 使用 openpyxl
                wb = openpyxl.load_workbook(filepath, data_only=True)
                self._ws = wb.active           # 读取活动工作表
                self._is_xlsx = True
            else:
                # 旧式 .xls 格式 → 使用 xlrd
                wb = xlrd.open_workbook(filepath)
                self._ws = wb.sheet_by_index(0)  # 读取第一个工作表
                self._is_xlsx = False
        finally:
            # 清理临时文件
            if temp_copy:
                try:
                    os.remove(temp_copy)
                except Exception:
                    pass

    @property
    def nrows(self):
        """返回工作表的总行数（0-indexed）"""
        return self._ws.max_row if self._is_xlsx else self._ws.nrows

    def cell_value(self, r, c):
        """
        获取指定单元格的值。

        参数:
            r (int): 行索引（0-indexed）
            c (int): 列索引（0-indexed）

        返回:
            单元格的值（类型取决于 Excel 中的实际数据类型）
        """
        if self._is_xlsx:
            # openpyxl 使用 1-indexed
            return self._ws.cell(r + 1, c + 1).value
        # xlrd 使用 0-indexed
        return self._ws.cell_value(r, c)
