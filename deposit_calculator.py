# -*- coding: utf-8 -*-
"""
=============================================================================
存款归属表计算引擎 v9
=============================================================================

功能说明：
  根据 A1411 和 A2411 两张监管月报表中的数据，结合存款归属表模板，
  自动计算并填充存款归属表的合计数。

数据来源：
  - A1411 月报表（资产负债表）：各类存款的期末余额（单位：元）
  - A2411 月报表（补充报表）：补充科目数据（单位：元）
  - 附件2：住房公积金明细表（用于核对住房公积金存款）
  - 附件3：社保基金存款明细表（用于核对社保基金存款）

计算逻辑：
  1. 从 A1411/A2411 读取各科目代码对应的余额数据
  2. 非计算行：按模板中的科目代码映射查表取数，各行独立填列，
     不进行子项数据间的汇总
  3. 计算行：按模板 F 列公式计算（如 F8=SUM(F40:F48)）。
     住房公积金存款、社保基金存款等计算行均按公式对子项求和，
     不直接取附件数据（附件数据仅作为参考记录到日志）
  4. 输出填充完成的存款归属表（.xlsx），所有数值固化为静态值

关键概念：
  - "计算行"：模板中备注列为"计算行"的行（F 列带公式），其值按公式计算填充
  - 非计算行：按各自科目代码独立取数填列，不参与任何子项汇总

单位说明：
  - A1411/A2411：元
  - 存款归属表：万元（输出时已转换）
  - 附件：元（读取后除以 10000 转换为万元）

依赖模块：
  - openpyxl：读写 .xlsx 文件
  - common.py：公共工具（to_float, ExcelSheet）
=============================================================================
"""
from pathlib import Path
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from common import to_float, ExcelSheet


# ============================================================
# 科目代码解析
# ============================================================

def _parse_sub_code(code_str):
    """
    解析科目代码字符串，拆分为 A1411 代码和 A2411 代码。

    科目代码格式有两种：
      1. 单代码："2003" → 仅从 A1411 查询
      2. 双代码："1001/2001" → 斜杠前为 A1411 代码，斜杠后为 A2411 代码
      3. 以 '22' 开头的代码 → 仅从 A2411 查询

    参数:
        code_str (str): 科目代码字符串，可能包含 '/' 分隔符

    返回:
        tuple: (a1411_code, a2411_code) — 两个代码均为字符串，
               不存在的部分返回空字符串
    """
    if not code_str or not str(code_str).strip():
        return "", ""
    code_str = str(code_str).strip()
    if '/' in code_str:
        # 双代码格式：A1411代码 / A2411代码
        parts = code_str.split('/', 1)
        return parts[0].strip(), parts[1].strip()
    if code_str.startswith('22'):
        # 以 22 开头 → 仅存在于 A2411
        return "", code_str
    return code_str, ""


# ============================================================
# 模板公式计算（计算行按公式求和）
# ============================================================

_SUM_RANGE_RE = re.compile(
    r'SUM\(\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)\)', re.IGNORECASE)
_CELL_REF_RE = re.compile(r'\$?([A-Za-z]+)\$?(\d+)')


def _formula_refs(formula):
    """
    提取公式中引用的所有单元格坐标。

    参数:
        formula (str): Excel 公式字符串，如 '=F12-SUM(F13:F21)+F22'

    返回:
        set[tuple]: {(列字母大写, 行号), ...}
    """
    refs = set()
    body = str(formula)
    for m in _SUM_RANGE_RE.finditer(body):
        col = m.group(1).upper()
        r1, r2 = int(m.group(2)), int(m.group(4))
        for rr in range(min(r1, r2), max(r1, r2) + 1):
            refs.add((col, rr))
    for m in _CELL_REF_RE.finditer(_SUM_RANGE_RE.sub('', body)):
        refs.add((m.group(1).upper(), int(m.group(2))))
    return refs


def _eval_excel_formula(formula, values):
    """
    计算简化版 Excel 公式。

    仅支持模板中出现的公式形态：单元格引用、SUM(单列区域)、加减运算。
    未在 values 中出现的引用按 0 处理。

    参数:
        formula (str): Excel 公式字符串（可带前导 '='）
        values (dict): {(列字母大写, 行号): float} 已知单元格数值

    返回:
        float: 公式计算结果
    """
    expr = str(formula).strip()
    if expr.startswith('='):
        expr = expr[1:]

    def _sum_repl(m):
        col = m.group(1).upper()
        r1, r2 = int(m.group(2)), int(m.group(4))
        total = sum(values.get((col, rr), 0.0)
                    for rr in range(min(r1, r2), max(r1, r2) + 1))
        return repr(total)

    expr = _SUM_RANGE_RE.sub(_sum_repl, expr)
    expr = _CELL_REF_RE.sub(
        lambda m: repr(values.get((m.group(1).upper(), int(m.group(2))), 0.0)),
        expr)
    return float(eval(expr, {'__builtins__': {}}, {}))


# ============================================================
# 附件读取
# ============================================================

def _read_attachment_total(filepath):
    """
    读取附件（附件2/附件3）的本金合计数。

    附件2/3的格式约定：
      - 表头在第3行之后
      - 最后一行为"合计"行（可能存在多个"合计"行，取最后一个）
      - 金额在第6列（0-indexed 列索引 = 5）

    参数:
        filepath (str): 附件文件路径（支持 .xls 和 .xlsx）

    返回:
        float: 本金合计数（单位：元）

    注意：
      返回的单位是元，调用方需要自行转换为万元（除以 10000）
    """
    ws = ExcelSheet(str(filepath))
    # 从后往前寻找最后一个"合计"行（总合计）
    for r in range(ws.nrows - 1, -1, -1):
        cell0 = str(ws.cell_value(r, 0)).strip()
        if cell0 in ('合计', '合\xa0计'):
            return to_float(ws.cell_value(r, 5))
    # 若没找到合计行，取倒数第二行作为兜底
    if ws.nrows > 3:
        return to_float(ws.cell_value(ws.nrows - 2, 5))
    return 0.0


# ============================================================
# 存款归属表计算器
# ============================================================

class DepositCalculator:
    """
    存款归属表计算器。

    功能：
      读取 A1411/A2411 月报表及附件数据，按模板结构计算存款归属表。

    使用步骤：
      1. 实例化：calc = DepositCalculator()
      2. 加载数据：calc.load_a1411(path), calc.load_a2411(path)
      3. （可选）加载附件：calc.load_attachments(att2, att3)
      4. 计算并输出：calc.compute(template_path, output_path, ...)

    计算流程：
      Step 1: 逐行读取模板，解析科目代码和 F 列公式
      Step 2: 非计算行按科目代码从 A1411/A2411 独立取数（不做子项汇总）
      Step 3: 计算行按模板 F 列公式计算（公积金/社保按公式求和，不取附件数）
      Step 4: 写入 Excel 并保存
    """

    def __init__(self):
        """初始化计算器，清空所有数据容器"""
        self.a1411_data = {}          # A1411 数据字典 {科目代码: 余额}
        self.a2411_data = {}          # A2411 数据字典 {科目代码: 余额}
        self.not_found_a1411 = []     # 未在 A1411 中找到的代码列表
        self.not_found_a2411 = []     # 未在 A2411 中找到的代码列表
        self.messages = []            # 运行日志消息列表
        self.attach2_value = None     # 附件2（住房公积金）本金合计（万元）
        self.attach3_value = None     # 附件3（社保基金）本金合计（万元）

    # ----------------------------------------------------------
    # 数据加载
    # ----------------------------------------------------------

    def load_a1411(self, filepath):
        """
        加载 A1411 月报表数据。

        读取规则：
          - 数据从第4行开始（0-indexed 行索引 = 3）
          - A列 = 科目代码，C列 = 期末余额

        参数:
            filepath (str): A1411 月报表文件路径
        """
        ws = ExcelSheet(str(filepath))
        self.a1411_data.clear()
        for r in range(3, ws.nrows):
            code = str(ws.cell_value(r, 0)).strip()
            if code:
                self.a1411_data[code] = to_float(ws.cell_value(r, 2))
        self.messages.append(f"A1411: {len(self.a1411_data)} records")

    def load_a2411(self, filepath):
        """
        加载 A2411 月报表数据。

        读取规则同 load_a1411：
          - 数据从第4行开始
          - A列 = 科目代码，C列 = 期末余额

        参数:
            filepath (str): A2411 月报表文件路径
        """
        ws = ExcelSheet(str(filepath))
        self.a2411_data.clear()
        for r in range(3, ws.nrows):
            code = str(ws.cell_value(r, 0)).strip()
            if code:
                self.a2411_data[code] = to_float(ws.cell_value(r, 2))
        self.messages.append(f"A2411: {len(self.a2411_data)} records")

    def load_attachments(self, attachment2=None, attachment3=None):
        """
        加载附件数据（住房公积金和社保基金存款）。

        注意：
          - 附件中单位为元，存款归属表单位为万元，自动除以 10000 转换。
          - 附件值仅作为参考记录到日志；计算行（公积金/社保基金存款）
            按模板公式对子项求和，不直接取附件数据。

        参数:
            attachment2 (str, optional): 附件2（住房公积金明细表）路径
            attachment3 (str, optional): 附件3（社保基金存款明细表）路径
        """
        if attachment2:
            val_yuan = _read_attachment_total(str(attachment2))
            self.attach2_value = val_yuan / 10000.0
            self.messages.append(
                f"附件2(住房公积金)本金合计: {val_yuan:,.2f} 元 "
                f"→ {self.attach2_value:,.2f} 万元")
        if attachment3:
            val_yuan = _read_attachment_total(str(attachment3))
            self.attach3_value = val_yuan / 10000.0
            self.messages.append(
                f"附件3(社保基金存款)本金合计: {val_yuan:,.2f} 元 "
                f"→ {self.attach3_value:,.2f} 万元")

    # ----------------------------------------------------------
    # 数据查询
    # ----------------------------------------------------------

    def lookup(self, code_a1411, code_a2411):
        """
        查询科目代码对应的余额值。

        同时查询 A1411 和 A2411，返回两者相加的结果。
        如果代码不存在，记录到 not_found 列表中以便诊断。

        参数:
            code_a1411 (str): A1411 科目代码（可空）
            code_a2411 (str): A2411 科目代码（可空）

        返回:
            float: 查询到的余额合计值
        """
        v1 = self.a1411_data.get(code_a1411, 0.0) if code_a1411 else 0.0
        v2 = self.a2411_data.get(code_a2411, 0.0) if code_a2411 else 0.0
        if code_a1411 and code_a1411 not in self.a1411_data:
            self.not_found_a1411.append(code_a1411)
        if code_a2411 and code_a2411 not in self.a2411_data:
            self.not_found_a2411.append(code_a2411)
        return v1 + v2

    # ----------------------------------------------------------
    # 核心计算
    # ----------------------------------------------------------

    def compute(self, template_path, output_path,
                attachment2=None, attachment3=None):
        """
        计算存款归属表并输出。

        参数:
            template_path (str): 存款归属表模板路径（.xlsx）
            output_path (str): 输出路径（必填，不再支持 None 回退）
            attachment2 (str, optional): 附件2：住房公积金明细表路径
            attachment3 (str, optional): 附件3：社保基金存款明细表路径

        返回:
            list[str]: 运行日志消息列表

        异常:
            ValueError: 如果 output_path 为空

        计算流程说明:
          Step 1 - 解析模板：读取科目代码、备注标记及 C/F 列公式
          Step 2 - 非计算行取数：按科目代码从 A1411/A2411 查表，
                   各行独立填列，不进行子项数据间的汇总
          Step 3 - 计算行求值：按模板公式计算（含公积金/社保基金
                   计算行，按公式对子项求和，不直接取附件数据）
          Step 4 - 写入 Excel：填充 F 列（及 C 列合计公式）并格式化
        """
        if not output_path:
            raise ValueError("output_path 是必填参数")

        self.not_found_a1411.clear()
        self.not_found_a2411.clear()

        # 加载附件数据（仅作为参考记录到日志，计算行不直接取附件数）
        self.load_attachments(attachment2, attachment3)

        # 打开模板
        wb = openpyxl.load_workbook(str(template_path))
        ws = wb.active

        # --------------------------------------------------
        # Step 1: 读取模板结构
        # --------------------------------------------------
        # 每条记录包含：
        #   sub_code: 科目代码（C列；以 '=' 开头的是合计公式，非科目代码）
        #   sub_name: 科目名称（D列）
        #   remark:   备注标记（G列，"计算行"等）
        # 同时收集 C/F 列中的 Excel 公式（计算行的计算依据）
        # --------------------------------------------------

        rows_data = []
        formulas = {}   # {(列字母, 行号): 公式字符串}
        values = {}     # {(列字母, 行号): float} 已确定的单元格数值
        for r in range(5, ws.max_row + 1):
            sub_code = str(ws.cell(r, 3).value or "").strip()   # C列：科目代码
            sub_name = str(ws.cell(r, 4).value or "").strip()   # D列：科目名称
            remark = str(ws.cell(r, 7).value or "").strip()     # G列：备注

            for col_idx, col_letter in ((3, 'C'), (6, 'F')):
                cell_v = ws.cell(r, col_idx).value
                if isinstance(cell_v, str) and cell_v.strip().startswith('='):
                    formulas[(col_letter, r)] = cell_v.strip()

            rows_data.append({'row': r, 'sub_code': sub_code,
                              'sub_name': sub_name, 'remark': remark,
                              'value': None})

        # --------------------------------------------------
        # Step 2: 非计算行按科目代码独立取数（不做子项间汇总）
        # --------------------------------------------------
        for rd in rows_data:
            code = rd['sub_code']
            if code and not code.startswith('='):
                c1, c2 = _parse_sub_code(code)
                rd['value'] = self.lookup(c1, c2)
                values[('F', rd['row'])] = rd['value']

        # --------------------------------------------------
        # Step 3: 计算行按模板公式求值
        # --------------------------------------------------
        # 公式引用的可能是其他公式单元格（如 F10=F58+F59），
        # 因此按依赖关系迭代求值，直至全部解析。
        # 住房公积金/社保基金存款计算行同样按公式（SUM 子项）求和，
        # 不直接取附件数据。
        # --------------------------------------------------
        pending = dict(formulas)
        while pending:
            progressed = False
            for key in list(pending):
                refs = _formula_refs(pending[key])
                if all(ref in values or ref not in formulas for ref in refs):
                    values[key] = _eval_excel_formula(pending[key], values)
                    del pending[key]
                    progressed = True
            if not progressed:
                # 存在无法解析的引用（如循环引用），按当前已知值兜底计算
                for key, formula in pending.items():
                    values[key] = _eval_excel_formula(formula, values)
                self.messages.append(
                    "WARNING: 部分公式引用无法解析，按当前已知值计算")
                break

        # 将公式计算结果回填到对应行
        for rd in rows_data:
            f_key = ('F', rd['row'])
            if f_key in formulas:
                rd['value'] = values.get(f_key)
            elif rd['remark'] == '计算行' and rd['value'] is None:
                self.messages.append(
                    f"WARNING: 第{rd['row']}行标记为计算行但模板中无公式，未填充")

        # 公积金/社保基金：输出公式计算值与附件参考值，便于人工核对
        for rd in rows_data:
            if rd['remark'] != '计算行' or rd['value'] is None:
                continue
            name = rd['sub_name']
            if '公积金' in name and self.attach2_value is not None:
                self.messages.append(
                    f"住房公积金存款: 公式计算={rd['value']:,.2f} 万元 "
                    f"(附件2参考: {self.attach2_value:,.2f} 万元)")
            elif ('社保' in name and '公积金' not in name
                    and self.attach3_value is not None):
                self.messages.append(
                    f"社保基金存款: 公式计算={rd['value']:,.2f} 万元 "
                    f"(附件3参考: {self.attach3_value:,.2f} 万元)")

        # --------------------------------------------------
        # Step 4: 写入 Excel
        # --------------------------------------------------
        # 始终写入计算值覆盖单元格（包括覆盖模板中可能存在的公式）。
        # 生成工具的职责就是计算并固化数值，供后续校验读取。
        # --------------------------------------------------
        data_font = Font(name='Arial', size=10)
        bold_font = Font(name='Arial', size=10, bold=True)
        border = Border(left=Side('thin'), right=Side('thin'),
                        top=Side('thin'), bottom=Side('thin'))
        num_fmt = '#,##0.00'

        for rd in rows_data:
            cell = ws.cell(rd['row'], 6)  # F列 = 金额列
            if rd['value'] is not None:
                cell.value = round(rd['value'], 2)
                cell.number_format = num_fmt
                cell.font = bold_font if rd['remark'] == '计算行' else data_font
                cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

        # C 列合计公式（如 C6=F6+F7+F10+F11 境内存款合计）同样固化为数值
        for (col_letter, r) in formulas:
            if col_letter == 'C' and ('C', r) in values:
                cell = ws.cell(r, 3)
                cell.value = round(values[('C', r)], 2)
                cell.number_format = num_fmt
                cell.font = bold_font
                cell.alignment = Alignment(horizontal='right', vertical='center')

        ws.column_dimensions['F'].width = 20

        # 保存输出
        wb.save(str(output_path))

        # --------------------------------------------------
        # 诊断信息
        # --------------------------------------------------
        if self.not_found_a1411:
            uq = list(dict.fromkeys(self.not_found_a1411))
            self.messages.append(f"A1411 unmatched({len(uq)}): {', '.join(uq[:5])}")
        if self.not_found_a2411:
            uq = list(dict.fromkeys(self.not_found_a2411))
            self.messages.append(f"A2411 unmatched({len(uq)}): {', '.join(uq[:5])}")
        filled = sum(1 for rd in rows_data if rd['value'] is not None)
        self.messages.append(f"Done: {filled} rows computed.")
        return self.messages
