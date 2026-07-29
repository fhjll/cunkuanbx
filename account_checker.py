# -*- coding: utf-8 -*-
"""
=============================================================================
科目余额核对表生成引擎
=============================================================================

功能说明：
  从科目余额表中提取各科目期末余额，填入科目余额核对表模板中。
  同时从附件2（住房公积金明细表）、附件3（社保基金存款明细表）、
  附件4（高管人员存款统计表）中提取本金和利息合计数。

数据处理流程：
  Step 1 - 明细行填充：
      按模板中 A 列（科目代码）和 D 列（利息科目代码）匹配，
      从源数据中查找对应余额写入 C 列（本金）和 F 列（利息）。
  Step 2 - 附件数据：
      读取各附件的合计行，将本金和利息填入对应行。
  Step 3 - 合计行计算：
      按模板公式规则汇总子项，计算个人存款、单位存款、
      财政性存款的合计数，写入对应合计行。
  Step 4 - 汇总行计算：
      计算境内存款合计和被保险存款余额合计行。
      境内存款 = 个人 + 单位 + 财政性 + 非存款类（其中项
      住房公积金/社保基金存款已含在单位存款中，不重复计入）。

关于个体工商户调整值（c1/c2）：
  - 生成核对表时填入的是原始子项合计数
  - c1/c2 调整值仅在「校验」步骤中应用（从单位扣减加到个人）
  - 这样在校验时可以灵活调整，无需重新生成核对表

安全写入：
  本模块特殊处理了合并单元格（MergedCell）问题。
  模板中部分行存在合并单元格，直接写入会导致 openpyxl 异常。
  _safe_set_value / _safe_get_value 函数自动解析合并区域，
  始终向合并区域的左上角单元格写入。

依赖模块：
  - openpyxl：读写 .xlsx 文件
  - xlrd：读取旧式 .xls 文件
  - common.py：公共工具（to_float, _is_xlsx_format）
=============================================================================
"""
from pathlib import Path
import os
import tempfile
import shutil
import zipfile
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.cell import MergedCell
from common import to_float, _is_xlsx_format


# ============================================================
# 安全写入工具 — 正确处理合并单元格
# ============================================================

def _resolve_merged(ws, row, col):
    """
    解析合并单元格，返回可写入的左上角坐标。

    如果 (row, col) 属于某个合并区域，返回该区域的左上角单元格坐标；
    否则返回原坐标。

    参数:
        ws (openpyxl.worksheet.worksheet.Worksheet): 工作表对象
        row (int): 1-indexed 行号
        col (int): 1-indexed 列号

    返回:
        tuple: (row, col) — 可写入的单元格坐标
    """
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        return row, col
    for mr in ws.merged_cells.ranges:
        if (mr.min_row <= row <= mr.max_row and
                mr.min_col <= col <= mr.max_col):
            return mr.min_row, mr.min_col
    return row, col  # 不会执行到这里，兜底


def _safe_get_value(ws, row, col):
    """
    安全读取单元格值，自动解析 MergedCell。

    参数:
        ws: 工作表对象
        row (int): 1-indexed 行号
        col (int): 1-indexed 列号

    返回:
        单元格的值
    """
    r, c = _resolve_merged(ws, row, col)
    return ws.cell(r, c).value


def _safe_set_value(ws, row, col, val, number_format=None, font=None,
                    alignment=None, border=None):
    """
    安全向工作表写入值，自动处理 MergedCell 问题。

    如果指定坐标属于合并区域，会自动定位到左上角单元格再写入。

    参数:
        ws: 工作表对象
        row (int): 1-indexed 行号
        col (int): 1-indexed 列号
        val: 要写入的值
        number_format (str, optional): 数字格式字符串，如 '#,##0.00'
        font (Font, optional): 字体样式
        alignment (Alignment, optional): 对齐方式
        border (Border, optional): 边框样式
    """
    r, c = _resolve_merged(ws, row, col)
    cell = ws.cell(r, c)
    cell.value = val
    if number_format:
        cell.number_format = number_format
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border


# ============================================================
# Excel 读取器（多 sheet 支持）
# ============================================================

class _ExcelSheets:
    """
    统一的 .xls / .xlsx 读取器，支持多 sheet。

    提供与 xlrd 兼容的接口（sheet_names, sheet_by_name），
    自动根据文件格式选择底层引擎。

    与 common.ExcelSheet 的区别：
      - common.ExcelSheet 只读取第一个活动 sheet
      - _ExcelSheets 读取所有 sheet（因为科目余额表可能有多个 sheet）

    特殊处理：
      - 文件扩展名为 .xls 但内部为 OOXML 时，自动转换为临时 .xlsx 读取
    """

    def __init__(self, filepath):
        """
        初始化读取器。

        参数:
            filepath (str): Excel 文件路径
        """
        filepath = str(filepath)
        ext = filepath.lower()
        temp_copy = None

        # 后缀为 .xls 但内部是 OOXML → 复制为临时 .xlsx
        if not ext.endswith((".xlsx", ".xlsm")) and _is_xlsx_format(filepath):
            fd, temp_copy = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            shutil.copy2(filepath, temp_copy)
            filepath = temp_copy
            ext = ".xlsx"

        try:
            if ext.endswith((".xlsx", ".xlsm")):
                self._wb = openpyxl.load_workbook(filepath, data_only=True)
                self._sheet_names = self._wb.sheetnames
                self._is_xlsx = True
            else:
                self._wb_xlrd = xlrd.open_workbook(filepath)
                self._sheet_names = self._wb_xlrd.sheet_names()
                self._is_xlsx = False
        finally:
            if temp_copy:
                try:
                    os.remove(temp_copy)
                except Exception:
                    pass

    def sheet_names(self):
        """返回所有工作表名称列表"""
        return self._sheet_names

    def sheet_by_name(self, name):
        """
        按名称获取指定工作表。

        参数:
            name (str): 工作表名称

        返回:
            _XlSheet 或 xlrd Sheet 对象
        """
        if self._is_xlsx:
            return _XlSheet(self._wb[name])
        return self._wb_xlrd.sheet_by_name(name)


class _XlSheet:
    """
    xlrd-compatible wrapper for openpyxl Worksheet。

    提供与 xlrd 兼容的 nrows 属性和 cell_value 方法，
    使得业务代码可以统一使用 xlrd 接口风格调用。
    """

    def __init__(self, ws):
        """封装 openpyxl Worksheet 对象"""
        self._ws = ws

    @property
    def nrows(self):
        """返回总行数（0-indexed）"""
        return self._ws.max_row

    def cell_value(self, r, c):
        """
        获取单元格值（0-indexed）。

        参数:
            r (int): 行索引（0-indexed）
            c (int): 列索引（0-indexed）

        返回:
            float 或原始值（None 转换为 0.0）
        """
        val = self._ws.cell(r + 1, c + 1).value
        return val if val is not None else 0.0


# ============================================================
# 附件读取
# ============================================================

def _read_attachment_summary(filepath):
    """
    读取附件文件，返回（本金合计, 利息合计）。

    不同附件格式略有不同：
      - 附件2（住房公积金明细表）：合计行在第6列（本金）和第7列（利息）
      - 附件3（社保基金存款明细表）：同附件2
      - 附件4（高管人员存款统计表）：合计行在第7列（本金）和第8列（利息）

    查找策略：
      1. 从最后一行向前查找"合计"行
      2. 若没找到，取倒数第2行作为兜底

    参数:
        filepath (str): 附件文件路径

    返回:
        tuple: (principal, interest) — 本金合计和利息合计（单位：元）
    """
    xls = _ExcelSheets(str(filepath))
    ws = xls.sheet_by_name(xls.sheet_names()[0])
    name_lower = str(filepath).lower()

    # 高管存款（附件4）列位置不同
    if "高管" in name_lower or "高级" in name_lower:
        last_row = ws.nrows - 1
        return to_float(ws.cell_value(last_row, 6)), to_float(ws.cell_value(last_row, 7))

    # 附件2/3：从后向前找"合计"行
    principal = 0.0
    interest = 0.0
    found = False
    for r in range(ws.nrows - 1, -1, -1):
        cell0 = str(ws.cell_value(r, 0)).strip()
        if cell0 in ('合计', '合\xa0计'):
            vp = to_float(ws.cell_value(r, 5))   # 第6列 = 本金
            vi = to_float(ws.cell_value(r, 6))   # 第7列 = 利息
            if vp != 0.0 or vi != 0.0:
                principal, interest, found = vp, vi, True
                break
            elif not found:
                principal, interest, found = vp, vi, True

    # 没找到合计行，取倒数第2行
    if principal == 0.0 and interest == 0.0 and ws.nrows > 3:
        for r in range(ws.nrows - 2, -1, -1):
            vp = to_float(ws.cell_value(r, 5))
            vi = to_float(ws.cell_value(r, 6))
            if vp != 0.0 or vi != 0.0:
                principal, interest = vp, vi
                break
    return principal, interest


# ============================================================
# 主引擎
# ============================================================

class AccountChecker:
    """
    科目余额核对表生成器。

    功能：
      读取科目余额表的各项科目余额，填入科目余额核对表模板，
      同时从附件中提取公积金、社保基金、高管存款的本金和利息数据。

    使用步骤：
      1. checker = AccountChecker()
      2. checker.load_account_balance(科目余额表路径)
      3. checker.compute(模板路径, 输出路径, 附件路径...)

    关于个体工商户调整值（c1/c2）：
      - 本生成器仅填入子项原始合计
      - c1/c2 在后续「校验」步骤中动态应用，不在此处处理
    """

    def __init__(self):
        """初始化科目余额核对表生成器"""
        self.acct_data = {}       # 科目余额数据 {科目代码: 余额}
        self.not_found = []       # 未匹配到的科目代码列表
        self.messages = []        # 运行日志

    def load_account_balance(self, filepath):
        """
        加载源科目余额数据。

        支持 .xls / .xlsx 格式，自动遍历所有 sheet。
        读取规则：
          - 数据从第4行开始（0-indexed）
          - A列 = 科目代码，G列 = 借方余额，H列 = 贷方余额
          - 余额取借方和贷方中非零的一个

        参数:
            filepath (str): 科目余额表文件路径
        """
        xls = _ExcelSheets(str(filepath))
        self.acct_data.clear()
        for sname in xls.sheet_names():
            ws = xls.sheet_by_name(sname)
            for r in range(3, ws.nrows):
                code = str(ws.cell_value(r, 0)).strip()
                if not code:
                    continue
                debit = to_float(ws.cell_value(r, 6))   # G列：借方余额
                credit = to_float(ws.cell_value(r, 7))  # H列：贷方余额
                balance = debit if debit != 0 else credit
                self.acct_data[code] = balance
        self.messages.append(
            f"科目余额表 ({filepath}): {len(self.acct_data)} 条记录, "
            f"{len(xls.sheet_names())} 个 sheet")

    def compute(self, template_path, output_path=None,
                attachment2=None, attachment3=None, attachment4=None):
        """
        填入科目余额核对表并输出。

        参数:
            template_path (str): 模板文件路径
            output_path (str): 输出文件路径（必填）
            attachment2 (str, optional): 附件2（住房公积金明细表）
            attachment3 (str, optional): 附件3（社保基金存款明细表）
            attachment4 (str, optional): 附件4（高管人员存款统计表）

        返回:
            list[str]: 运行日志消息列表

        异常:
            ValueError: 如果 output_path 为空

        注意:
          - c1/c2/c3/c4 调整值仅在「校验」时应用，生成时不写入
          - 合计行始终覆盖为静态值（而非保留公式），避免公式缓存丢失问题
        """
        if not output_path:
            raise ValueError("output_path 是必填参数")

        self.not_found.clear()

        wb = openpyxl.load_workbook(str(template_path))
        ws = wb.active

        data_font = Font(name='Arial', size=10)
        border = Border(left=Side('thin'), right=Side('thin'),
                        top=Side('thin'), bottom=Side('thin'))
        num_fmt = '#,##0.00'

        # --------------------------------------------------
        # Step 1: 按模板中的科目代码填入余额
        # --------------------------------------------------
        # 模板列位置：
        #   A列 = 科目代码（用于查找本金，写入C列）
        #   B列 = 科目名称
        #   C列 = 本金
        #   D列 = 利息科目代码（用于查找利息，写入F列）
        #   F列 = 利息
        # --------------------------------------------------
        written_a = 0    # C 列写入计数
        written_d = 0    # F 列写入计数
        skipped_merge = 0

        for r in range(4, ws.max_row + 1):
            code_a_raw = ws.cell(r, 1).value   # A列：科目代码
            code_d_raw = ws.cell(r, 4).value   # D列：利息科目代码

            # 跳过合并单元格（header/label 行）
            if isinstance(code_a_raw, MergedCell) or isinstance(code_d_raw, MergedCell):
                skipped_merge += 1
                continue

            code_a = str(code_a_raw or "").strip()
            code_d = str(code_d_raw or "").strip()

            # 清理 .0 后缀（Excel 导出时常见问题：数字显示为 "2003.0"）
            if code_a and '.' in code_a:
                try:
                    code_a = str(int(float(code_a)))
                except:
                    pass
            if code_d and '.' in code_d:
                try:
                    code_d = str(int(float(code_d)))
                except:
                    pass

            # 写入本金（C列）
            if code_a:
                if code_a in self.acct_data:
                    _safe_set_value(ws, r, 3, self.acct_data[code_a],
                                    number_format=num_fmt, font=data_font, border=border,
                                    alignment=Alignment(horizontal='right', vertical='center'))
                    written_a += 1
                else:
                    self.not_found.append(f"A:{code_a}")

            # 写入利息（F列）
            if code_d:
                if code_d in self.acct_data:
                    _safe_set_value(ws, r, 6, self.acct_data[code_d],
                                    number_format=num_fmt, font=data_font, border=border,
                                    alignment=Alignment(horizontal='right', vertical='center'))
                    written_d += 1
                else:
                    self.not_found.append(f"D:{code_d}")

        self.messages.append(
            f"明细写入: C列={written_a} 行, F列={written_d} 行"
            + (f", 跳过合并={skipped_merge}" if skipped_merge else ""))

        # --------------------------------------------------
        # Step 2: 附件数据填入
        # --------------------------------------------------
        # 根据模板中的科目名称关键词定位行，填入本金和利息
        # --------------------------------------------------

        # 附件2：住房公积金（本金、利息分别填入C列和F列）
        if attachment2:
            principal2, interest2 = _read_attachment_summary(str(attachment2))
            self.messages.append(f"附件2(住房公积金): 本金={principal2:,.2f}, 利息={interest2:,.2f}")
            for r in range(4, ws.max_row + 1):
                cell_name = str(ws.cell(r, 2).value or "").strip()
                if "住房公积金" in cell_name and "明细" not in cell_name:
                    _safe_set_value(ws, r, 3, principal2, number_format=num_fmt,
                                    font=data_font, border=border,
                                    alignment=Alignment(horizontal='right', vertical='center'))
                    _safe_set_value(ws, r, 6, interest2, number_format=num_fmt,
                                    font=data_font, border=border,
                                    alignment=Alignment(horizontal='right', vertical='center'))
                    break

        # 附件3：社保基金存款
        if attachment3:
            principal3, interest3 = _read_attachment_summary(str(attachment3))
            self.messages.append(f"附件3(社保基金): 本金={principal3:,.2f}, 利息={interest3:,.2f}")
            for r in range(4, ws.max_row + 1):
                cell_name = str(ws.cell(r, 2).value or "").strip()
                if "社保基金" in cell_name:
                    _safe_set_value(ws, r, 3, principal3, number_format=num_fmt,
                                    font=data_font, border=border,
                                    alignment=Alignment(horizontal='right', vertical='center'))
                    _safe_set_value(ws, r, 6, interest3, number_format=num_fmt,
                                    font=data_font, border=border,
                                    alignment=Alignment(horizontal='right', vertical='center'))
                    break

        # 附件4：高管人员存款
        if attachment4:
            principal4, interest4 = _read_attachment_summary(str(attachment4))
            self.messages.append(f"附件4(高管人员): 本金={principal4:,.2f}, 利息={interest4:,.2f}")
            found = False
            for r in range(4, ws.max_row + 1):
                for nc in [1, 2, 4]:  # 在 A/B/D 列中查找
                    cell_name = str(ws.cell(r, nc).value or "").strip()
                    if "高级管理人员" in cell_name or "高管" in cell_name:
                        _safe_set_value(ws, r, 3, principal4, number_format=num_fmt,
                                        font=data_font, border=border,
                                        alignment=Alignment(horizontal='right', vertical='center'))
                        _safe_set_value(ws, r, 6, interest4, number_format=num_fmt,
                                        font=data_font, border=border,
                                        alignment=Alignment(horizontal='right', vertical='center'))
                        found = True
                        break
                if found:
                    break
            if not found:
                self.messages.append("  警告：未找到'高级管理人员'行，附件4数据未填入")

        # 调整列宽
        for col in ['C', 'F']:
            ws.column_dimensions[col].width = 20

        # --------------------------------------------------
        # Step 3: 计算并写入合计行
        # --------------------------------------------------
        # 个人存款合计 = SUM(2003, 2004, 200503, 200504, 保证金个人, 22410401)
        # 单位存款合计 = SUM(2001, 2002, 200501, 200502, 200601, 2011, 保证金单位, 22410402, 630108)
        #
        # 此处仅计算子项原始合计，c1/c2/c3/c4 调整值在校验时应用
        # --------------------------------------------------
        gr_total_p = 0.0
        gr_total_i = 0.0
        dw_total_p = 0.0
        dw_total_i = 0.0

        for r in range(4, ws.max_row + 1):
            av = str(_safe_get_value(ws, r, 1) or '').strip()
            if '.' in av:
                try:
                    av = str(int(float(av)))
                except:
                    pass
            if not av:
                continue
            # 个人存款相关代码
            if av in ('2003', '2004', '200503', '200504') or \
               (av.startswith('2014') and av.endswith('2')) or \
               av == '22410401':
                gr_total_p += to_float(_safe_get_value(ws, r, 3))
                gr_total_i += to_float(_safe_get_value(ws, r, 6))
            # 单位存款相关代码
            elif av in ('2001', '2002', '200501', '200502', '200601', '2011') or \
                 (av.startswith('2014') and av.endswith('1')) or \
                 av in ('22410402', '630108'):
                dw_total_p += to_float(_safe_get_value(ws, r, 3))
                dw_total_i += to_float(_safe_get_value(ws, r, 6))

        # 写入个人存款合计行
        gr_found = False
        dw_found = False
        for r in range(4, ws.max_row + 1):
            b_val = str(_safe_get_value(ws, r, 2) or '').strip()
            if '个人存款' in b_val and not gr_found:
                gr_found = True
                _safe_set_value(ws, r, 3, gr_total_p, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                _safe_set_value(ws, r, 6, gr_total_i, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                self.messages.append(
                    f"个人存款合计 (R{r}): 本金={gr_total_p:,.2f}, "
                    f"利息={gr_total_i:,.2f} (子项合计，调整值在校验时应用)")
            elif '单位存款' in b_val and not dw_found:
                dw_found = True
                _safe_set_value(ws, r, 3, dw_total_p, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                _safe_set_value(ws, r, 6, dw_total_i, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                self.messages.append(
                    f"单位存款合计 (R{r}): 本金={dw_total_p:,.2f}, "
                    f"利息={dw_total_i:,.2f} (子项合计，调整值在校验时应用)")

        if not gr_found:
            self.messages.append("  ⚠ 警告：未找到「个人存款」合计行")
        if not dw_found:
            self.messages.append("  ⚠ 警告：未找到「单位存款」合计行")

        # 财政性存款合计（国库定期 200602 + 代理财政 2007）
        fiscal_p = 0.0
        fiscal_i = 0.0
        for r in range(4, ws.max_row + 1):
            av = str(_safe_get_value(ws, r, 1) or '').strip()
            if '.' in av:
                try:
                    av = str(int(float(av)))
                except:
                    pass
            if not av:
                continue
            if av == '200602':
                fiscal_p += to_float(_safe_get_value(ws, r, 3))
                fiscal_i += to_float(_safe_get_value(ws, r, 6))
            elif av == '2007':
                fiscal_p += to_float(_safe_get_value(ws, r, 3))
                fiscal_i += to_float(_safe_get_value(ws, r, 6))

        fiscal_found = False
        for r in range(4, ws.max_row + 1):
            b_val = str(_safe_get_value(ws, r, 2) or '').strip()
            if '财政性存款' in b_val and '（三）' in b_val and not fiscal_found:
                fiscal_found = True
                _safe_set_value(ws, r, 3, fiscal_p, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                _safe_set_value(ws, r, 6, fiscal_i, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                self.messages.append(
                    f"财政性存款合计 (R{r}): 本金={fiscal_p:,.2f}, 利息={fiscal_i:,.2f}")

        if not fiscal_found:
            self.messages.append("  ⚠ 警告：未找到「（三）财政性存款」合计行")

        # --------------------------------------------------
        # Step 4: 境内存款合计 和 被保险存款余额
        # --------------------------------------------------

        # 一、境内存款 = 个人 + 单位 + 财政性 + 非存款类
        #   （模板公式 C5=C6+C7+C10+C11）
        # 注意：「其中：住房公积金存款」和「社保基金存款」是（二）单位存款
        #   下的其中项，已包含在单位存款中，境内存款汇总时不再重复计入。
        # 非存款类金融机构存款 = 明细行（201705 + 201799）合计
        nondep_p = 0.0
        nondep_i = 0.0
        for r in range(4, ws.max_row + 1):
            av = str(_safe_get_value(ws, r, 1) or '').strip()
            if '.' in av:
                try:
                    av = str(int(float(av)))
                except:
                    pass
            if av in ('201705', '201799'):
                nondep_p += to_float(_safe_get_value(ws, r, 3))
                nondep_i += to_float(_safe_get_value(ws, r, 6))

        # 写入（四）非存款类金融机构存款合计行（覆盖模板公式，固化数值）
        for r in range(4, ws.max_row + 1):
            b_val = str(_safe_get_value(ws, r, 2) or '').strip()
            if '非存款类' in b_val:
                _safe_set_value(ws, r, 3, nondep_p, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                _safe_set_value(ws, r, 6, nondep_i, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                self.messages.append(
                    f"非存款类金融机构存款合计 (R{r}): 本金={nondep_p:,.2f}, "
                    f"利息={nondep_i:,.2f}")
                break

        dom_p = gr_total_p + dw_total_p + fiscal_p + nondep_p
        dom_i = gr_total_i + dw_total_i + fiscal_i + nondep_i

        # 写入境内存款合计行
        dom_found = False
        for r in range(4, ws.max_row + 1):
            a_val = str(_safe_get_value(ws, r, 1) or '').strip()
            b_val = str(_safe_get_value(ws, r, 2) or '').strip()
            if '境内存款' in (a_val or b_val) and \
               '（一）' not in (a_val or b_val) and \
               '（二）' not in (a_val or b_val):
                dom_found = True
                _safe_set_value(ws, r, 3, dom_p, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                _safe_set_value(ws, r, 6, dom_i, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                self.messages.append(
                    f"境内存款合计 (R{r}): 本金={dom_p:,.2f}, 利息={dom_i:,.2f}")
                break
        if not dom_found:
            self.messages.append("  ⚠ 警告：未找到「一、境内存款」合计行")

        # 被保险存款余额 = C6+C7+C10-C14-C15（按模板公式计算）
        ins_p = to_float(_safe_get_value(ws, 6, 3)) + \
                to_float(_safe_get_value(ws, 7, 3)) + \
                to_float(_safe_get_value(ws, 10, 3)) - \
                to_float(_safe_get_value(ws, 14, 3)) - \
                to_float(_safe_get_value(ws, 15, 3))
        ins_i = to_float(_safe_get_value(ws, 6, 6)) + \
                to_float(_safe_get_value(ws, 7, 6)) + \
                to_float(_safe_get_value(ws, 10, 6)) - \
                to_float(_safe_get_value(ws, 14, 6)) - \
                to_float(_safe_get_value(ws, 15, 6))
        for r in range(4, ws.max_row + 1):
            a_val = str(_safe_get_value(ws, r, 1) or '').strip()
            if '被保险存款余额' in a_val or ('合计' in a_val and r >= 15):
                _safe_set_value(ws, r, 3, ins_p, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                _safe_set_value(ws, r, 6, ins_i, number_format=num_fmt,
                                font=data_font, border=border,
                                alignment=Alignment(horizontal='right', vertical='center'))
                self.messages.append(
                    f"被保险存款余额 (R{r}): 本金={ins_p:,.2f}, 利息={ins_i:,.2f}")
                break

        # --------------------------------------------------
        # 保存
        # --------------------------------------------------
        wb.save(str(output_path))

        # --------------------------------------------------
        # 诊断输出
        # --------------------------------------------------
        if not self.acct_data:
            self.messages.append("  ❌ 源数据为空！请检查科目余额表文件路径及格式")
        if self.not_found:
            uq = list(dict.fromkeys(self.not_found))
            self.messages.append(f"未匹配代码 ({len(uq)}): {', '.join(uq[:10])}")
        else:
            self.messages.append("全部代码匹配成功 ✓")
        self.messages.append(f"输出: {output_path}")
        return self.messages
