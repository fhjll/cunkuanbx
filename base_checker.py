# -*- coding: utf-8 -*-
"""
=============================================================================
保费基数核对表生成引擎（合并模板）
=============================================================================

功能说明：
  按《保费基数核对表（模板）》把三块数据填进同一张表：
    1. 存款归属核对（C~F 列）—— 取自 A1411 / A2411 月报表（万元）
    2. 会计科目核对（G~M 列）—— 取自科目余额表（元）
    3. 被保险存款申报表（N~P 列）—— 取自 IRP001 申报表（元）
  再按模板中的公式对子项数据逐级计算合并（个人/单位/财政性/非存款类/
  高级管理人员/不予保险/被保险存款合计），并把结果固化为静态数值。

与旧版两个生成器的关系：
  旧版「存款归属表」和「科目余额核对表」是两张独立的表，
  现合并为一张《保费基数核对表》，本模块取代：
    - deposit_calculator.DepositCalculator 的模板填充逻辑（F 列）
    - account_checker.AccountChecker 的模板填充逻辑（I / L 列）

计算顺序（与模板公式一致）：
  Step 1 读模板：解析每行的子项代码（C 列）、科目代码（G 列本金 / J 列利息）
         以及 F/I/L/M/N/O/P 各列的公式
  Step 2 明细取数：按代码逐项取数（存款归属取 A1411/A2411，会计科目取科目余额表）
  Step 3 公式计算：按模板公式（含跨行引用、SUM 区域、IF/ROUND）逐条求值
  Step 4 固化输出：把计算值与校验页差额写入 .xlsx（覆盖公式，全部为静态数值）

调整值（可自定义的调整项列表）：
  每个调整项包含：名称、用途、对象、字段、方式、金额（元）
    用途：会计科目核对（写入表内"调整项"行，参与会计科目合计）
          存款归属核对（比较时按方式增减 IRP001 本金，写入 Q9/Q10）
    对象：个人存款 / 单位存款
    字段：本金 / 应付利息（用途=存款归属核对时只比较本金，固定为"本金"）
    方式：增加 / 扣减
  对象可选：
    · 个人存款 / 单位存款 / 个人↔单位（一增一减，个体工商户用）
    · 会计科目核对表其他合计项：财政性存款 / 非存款类金融机构存款 /
      高级管理人员 / 不予保险的其他存款（模板无专门调整行，
      调整额直接并入该合计项的公式，上层汇总同步生效）
  项目可自由增删（标准项除外，其 locked=True 不可删除）；
  标准口径（个体工商户 + 久悬户）为：
    · 个体工商户本金  会计科目核对 个人↔单位 本金     增加   （个人＋、单位－）
    · 个体工商户利息  会计科目核对 个人↔单位 应付利息 增加   （个人＋、单位－）
    · 个人久悬户      存款归属核对 个人存款 本金     扣减   → Q9
    · 单位久悬户      存款归属核对 单位存款 本金     扣减   → Q10
  Q7/Q8 保存"由单位调整至个人"的个人侧净额（本金/利息），供外部核对。

单位说明：
  - A1411/A2411：万元（与存款归属表口径一致，无需换算）
  - 科目余额表：元
  - IRP001：元
  - 附件：元

依赖模块：
  - openpyxl：读写 .xlsx 文件
  - common.py：公共工具（to_float, ExcelSheet, _is_xlsx_format）
  - deposit_calculator / account_checker：复用取数、附件与安全写入工具
=============================================================================
"""
import os
import re
import time

import openpyxl
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter, column_index_from_string

from common import to_float, excel_round, ExcelSheet
from deposit_calculator import _parse_sub_code, _read_attachment_total
from account_checker import (_ExcelSheets, _read_attachment_summary,
                             _resolve_merged, _safe_set_value)


# ============================================================
# 模板结构常量
# ============================================================

# 主表名称（模板内的 sheet 名）
MAIN_SHEET_NAME = '保费基数核对表'
CHECK_SHEET_NAME = '校验'

# 需要读取公式的列：C 列（兼容旧式"合计公式写在 C 列"的模板）、
# F 列（存款归属金额）、I/L 列（会计科目本金/利息）、
# M 列（会计科目合计）、N/O/P 列（IRP001 本金/利息/合计）
FORMULA_COLUMNS = ('C', 'F', 'I', 'L', 'M', 'N', 'O', 'P')

# 调整值写入单元格（模板 Q 列为「差额说明」，
# Q7/Q8=个人侧净调整额（会计科目），Q9/Q10=存款归属核对调整额）
ADJ_CELLS = {'c1': (7, 'Q'), 'c2': (8, 'Q'), 'c3': (9, 'Q'), 'c4': (10, 'Q')}

# ------------------------------------------------------------
# 调整项取值范围
# ------------------------------------------------------------
ADJ_SCOPE_ACCT = '会计科目核对'      # 参与会计科目合计（写入表内调整项行）
ADJ_SCOPE_DEP = '存款归属核对'       # 参与存款归属与 IRP001 的比较（写入 Q9/Q10）
ADJ_SCOPES = (ADJ_SCOPE_ACCT, ADJ_SCOPE_DEP)
ADJ_TARGETS = ('个人存款', '单位存款')
ADJ_TARGET_BOTH = '个人↔单位'        # 一增一减：个人按方式增减、单位同步反向
# 会计科目核对表的其他合计项（模板无专门调整行，调整额直接并入该合计公式）
ADJ_TOTALS = ('财政性存款', '非存款类金融机构存款', '高级管理人员',
              '不予保险的其他存款')
ADJ_TARGET_OPTIONS = ADJ_TARGETS + (ADJ_TARGET_BOTH,) + ADJ_TOTALS
# 存款归属核对（校验页 Q9/Q10）只支持个人/单位
ADJ_DEP_TARGETS = ADJ_TARGETS + (ADJ_TARGET_BOTH,)

# 合计项 → 会计科目名称关键词（用于定位模板中的合计行）
ADJ_TOTAL_KEYWORDS = {
    '个人存款': '（一）个人存款',
    '单位存款': '（二）单位存款',
    '财政性存款': '（三）财政性存款',
    '非存款类金融机构存款': '（四）非存款类',
    '高级管理人员': '高级管理人员',
    '不予保险的其他存款': '不予保险',
}
ADJ_FIELDS = ('本金', '应付利息')
ADJ_MODES = ('增加', '扣减')
ADJ_DEP_FIELD = '本金'               # 存款归属核对只比较本金

# 标准口径的调整项模板：(名称, 用途, 对象, 字段, 方式, 是否标准项)
# 个体工商户属于"由单位调整至个人"，个人增、单位减由一条记录表达
# （对象=个人↔单位）；标准项默认不可删除。
DEFAULT_ADJUSTMENT_ITEMS = [
    ('个体工商户本金', ADJ_SCOPE_ACCT, ADJ_TARGET_BOTH, '本金', '增加', True),
    ('个体工商户利息', ADJ_SCOPE_ACCT, ADJ_TARGET_BOTH, '应付利息', '增加', True),
    ('个人久悬户', ADJ_SCOPE_DEP, '个人存款', '本金', '扣减', False),
    ('单位久悬户', ADJ_SCOPE_DEP, '单位存款', '本金', '扣减', False),
]

# 说明：所有计算口径（合计公式、调整项所在单元格、取数代码）都从传入的模板
# 读取，程序内不内置任何计算公式——只需修改模板即可改变计算方式。若模板中的
# 公式引用了没有子项代码/科目代码的行，工具会在运行日志中提示，而不会自行补数。
#
# IRP001 各项目在申报表中的关键词（章节, 项目）——按会计科目分组的行名匹配
IRP001_ITEM_KEYS = [
    ('（一）个人存款', '一、境内存款', '（一）个人存款'),
    ('（二）单位存款', '一、境内存款', '（二）单位存款'),
    ('（三）财政性存款', None, '（三）财政性存款'),
    ('（四）非存款类', None, '（四）非存款类'),
    ('高级管理人员', None, '高级管理人员'),
    ('不予保险', None, '不予保险的其他存款'),
]

# 附件2/3 的其中项名称（存款归属计算行，仅作参考值记录）
ATT_NAME_HINTS = {'att2': '住房公积金', 'att3': '社保基金'}


# ============================================================
# 调整项工具函数
# ============================================================

def build_default_adjustments(c1=0, c2=0, c3=0, c4=0):
    """
    按标准口径生成调整项列表（个体工商户本金/利息 + 个人/单位久悬户）。

    参数:
        c1 (float): 个体工商户本金（个人增、单位减）
        c2 (float): 个体工商户应付利息（个人增、单位减）
        c3 (float): 个人久悬户（存款归属核对时从 IRP001 个人本金中扣减）
        c4 (float): 单位久悬户（存款归属核对时从 IRP001 单位本金中扣减）

    返回:
        list[dict]: 调整项列表（4 条：个体工商户本金/利息、个人/单位久悬户）
    """
    amounts = [c1, c2, c3, c4]
    return [
        {'name': name, 'scope': scope, 'target': target, 'field': field,
         'mode': mode, 'amount': to_float(amount), 'locked': locked}
        for (name, scope, target, field, mode, locked), amount
        in zip(DEFAULT_ADJUSTMENT_ITEMS, amounts)
    ]


def normalize_adjustments(items):
    """
    规范化调整项列表：统一取值范围、过滤空行、修正存款归属核对字段。

    参数:
        items (list[dict]): 调整项列表，每项含
            name / scope / target / field / mode / amount

    返回:
        list[dict]: 规范化后的调整项列表
    """
    out = []
    for it in (items or []):
        if not isinstance(it, dict):
            continue
        name = str(it.get('name', '') or '').strip()
        scope = it.get('scope', ADJ_SCOPE_ACCT)
        target = it.get('target', ADJ_TARGETS[0])
        field = it.get('field', ADJ_FIELDS[0])
        mode = it.get('mode', ADJ_MODES[0])
        amount = to_float(it.get('amount', 0.0))
        locked = bool(it.get('locked', False))
        if scope not in ADJ_SCOPES:
            scope = ADJ_SCOPE_ACCT
        if target not in ADJ_TARGET_OPTIONS:
            target = ADJ_TARGETS[0]
        if field not in ADJ_FIELDS:
            field = ADJ_FIELDS[0]
        if mode not in ADJ_MODES:
            mode = ADJ_MODES[0]
        if scope == ADJ_SCOPE_DEP:
            field = ADJ_DEP_FIELD      # 存款归属核对只比较本金
        if not name and amount == 0:
            continue                   # 空行忽略
        out.append({'name': name, 'scope': scope, 'target': target,
                    'field': field, 'mode': mode, 'amount': amount,
                    'locked': locked})
    return out


def merge_standard_pairs(items):
    """
    把"个体工商户"式的成对调整项合并为一条"个人↔单位"标准项。

    早期版本把个体工商户本金/利息拆成两条（个人·增加 + 单位·扣减）；
    现统一为一条 对象=个人↔单位 的标准项（locked=True，不可删除）。
    仅当两条记录 名称、字段、金额 相同且方式互为增减时合并，
    合并前后计算结果完全一致。

    参数:
        items (list[dict]): 调整项列表（可能是旧格式的成对记录）

    返回:
        list[dict]: 合并后的调整项列表
    """
    items = normalize_adjustments(items)
    out = []
    used = set()
    for i, it in enumerate(items):
        if i in used or it.get('locked'):
            if i not in used:
                out.append(it)
            continue
        for j in range(i + 1, len(items)):
            other = items[j]
            if j in used or other.get('locked'):
                continue
            same = (it['name'] == other['name'] and
                    it['scope'] == other['scope'] and
                    it['field'] == other['field'] and
                    abs(it['amount'] - other['amount']) <= 0.001 and
                    it['target'] in ADJ_TARGETS and
                    other['target'] in ADJ_TARGETS and
                    it['target'] != other['target'] and
                    it['mode'] != other['mode'])
            if not same:
                continue
            # 以"个人侧"的方式为准：个人增加 = 单位扣减 → 个人↔单位·增加
            person = it if it['target'] == '个人存款' else other
            out.append({'name': person['name'], 'scope': person['scope'],
                        'target': ADJ_TARGET_BOTH, 'field': person['field'],
                        'mode': person['mode'], 'amount': person['amount'],
                        'locked': True})
            used.add(i)
            used.add(j)
            break
        else:
            if i not in used:
                out.append(it)
    return out


def summarize_adjustments(items):
    """
    汇总调整项为两张"净额表"。

    参数:
        items (list[dict]): 调整项列表

    返回:
        tuple:
          acct (dict): {(对象, 字段): 净额} —— 会计科目核对（写入表内调整项行）
          dep  (dict): {对象: 应从 IRP001 本金中扣减的净额} —— 存款归属核对
                       （方式=扣减 记正数、方式=增加 记负数，供 Q9/Q10 使用）
    """
    acct = {}
    dep = {}
    for it in normalize_adjustments(items):
        sign = -1.0 if it['mode'] == '扣减' else 1.0
        amount = sign * it['amount']
        if it['scope'] == ADJ_SCOPE_DEP:
            if it['target'] not in ADJ_DEP_TARGETS:
                continue      # 存款归属核对只支持个人/单位（由 compute 记录提示）
            # Q 值是"从 IRP001 本金中扣减的金额"：方式=扣减记正数、方式=增加记负数
            if it['target'] == ADJ_TARGET_BOTH:
                dep['个人存款'] = dep.get('个人存款', 0.0) - amount
                dep['单位存款'] = dep.get('单位存款', 0.0) + amount
            else:
                dep[it['target']] = dep.get(it['target'], 0.0) - amount
        elif it['target'] == ADJ_TARGET_BOTH:
            # 一增一减：个人按方式增减，单位同步反向（如个体工商户存款）
            acct[('个人存款', it['field'])] = \
                acct.get(('个人存款', it['field']), 0.0) + amount
            acct[('单位存款', it['field'])] = \
                acct.get(('单位存款', it['field']), 0.0) - amount
        else:
            key = (it['target'], it['field'])
            acct[key] = acct.get(key, 0.0) + amount
    return acct, dep


def adjustment_row_values(items):
    """
    把调整项换算为表内"调整项"行的填写值（模板行口径）。

    模板合计公式的方向不同：
      · 个人行（＋）：个人合计为 SUM 累加（SUM(I14:I24)）
        → 该行填"应加进个人存款的金额"
      · 单位行（－）：单位合计对其做减法（I8 = SUM(I25:I37)-I38-I39）
        → 该行填"应从单位存款扣减的金额"（即单位净增减额取反）
    因此个体工商户（个人增、单位减）在两行中都填正数，
    与银行填报口径一致。

    参数:
        items (list[dict]): 调整项列表

    返回:
        dict: {(对象, 字段): 填写值}
    """
    acct, _ = summarize_adjustments(items)
    return {(target, field): (net if target == '个人存款' else -net)
            for (target, field), net in acct.items()}


def describe_adjustments(items, scope=None, target=None, field=None):
    """
    生成调整项明细文本（用于单元格批注与运行日志）。

    参数:
        items (list[dict]): 调整项列表
        scope (str, optional): 只保留该用途的项目
        target (str, optional): 只保留该对象的项目（"个人↔单位"项目两侧都算）
        field (str, optional): 只保留该字段（本金/应付利息）的项目

    返回:
        str: 形如 "个体工商户本金 增加 2,074,479.68" 的多行文本（无内容返回 ''）
    """
    lines = []
    for it in normalize_adjustments(items):
        if scope and it['scope'] != scope:
            continue
        if target:
            # "个人↔单位"项目同时属于个人侧与单位侧；其他合计项只匹配自身
            if it['target'] != target and not (
                    it['target'] == ADJ_TARGET_BOTH and target in ADJ_TARGETS):
                continue
        if field and it['field'] != field:
            continue
        lines.append(f"{it['name'] or '调整项'}（{it['target']}·{it['field']}"
                     f"·{it['mode']}）{it['amount']:,.2f}")
    return '\n'.join(lines)


# ============================================================
# 公式求值（支持模板中用到的公式形态）
# ============================================================
_SUM_RANGE_RE = re.compile(
    r'SUM\(\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*\)',
    re.IGNORECASE)
_RANGE_RE = re.compile(
    r'\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)')
_CELL_RE = re.compile(r'\$?([A-Za-z]{1,3})\$?(\d+)')
_SHEET_PREFIX_RE = re.compile(r"[^\s()+\-*/,<>=!]+!")


def _range_cols(c1, c2):
    """展开列区间 F:H → ['F','G','H']"""
    i1 = column_index_from_string(c1.upper())
    i2 = column_index_from_string(c2.upper())
    return [get_column_letter(i) for i in range(min(i1, i2), max(i1, i2) + 1)]


def _sum_range(values, c1, r1, c2, r2):
    """按模板 SUM 语义求和（支持单列区间与单行区间）"""
    total = 0.0
    for col in _range_cols(c1, c2):
        for row in range(min(int(r1), int(r2)), max(int(r1), int(r2)) + 1):
            total += values.get((col, row), 0.0)
    return total


def _formula_refs(formula):
    """
    提取公式引用的单元格坐标（用于依赖排序）。

    返回:
        set[tuple]: {(列字母, 行号), ...}
    """
    refs = set()
    body = _SHEET_PREFIX_RE.sub('', str(formula))
    for m in _SUM_RANGE_RE.finditer(body):
        for col in _range_cols(m.group(1), m.group(3)):
            for rr in range(int(m.group(2)), int(m.group(4)) + 1):
                refs.add((col, rr))
    body = _SUM_RANGE_RE.sub('', body)
    for m in _CELL_RE.finditer(_RANGE_RE.sub('', body)):
        refs.add((m.group(1).upper(), int(m.group(2))))
    return refs


def _eval_formula(formula, values):
    """
    求值模板公式。

    支持形态：单元格引用（可带工作表前缀）、SUM(区域)、ROUND(x,n)、
    IF(条件,a,b)、比较运算（= <> < <= > >=）与四则运算。
    未出现在 values 中的引用按 0 处理。

    参数:
        formula (str): 公式文本（可含前导 '='）
        values (dict): {(列字母, 行号): float}

    返回:
        float: 计算结果
    """
    expr = str(formula).strip()
    if expr.startswith('='):
        expr = expr[1:]
    expr = _SHEET_PREFIX_RE.sub('', expr)          # 去掉 '保费基数核对表!' 前缀

    # SUM(区域) → 数值
    def _sum_repl(m):
        return repr(_sum_range(values, m.group(1), m.group(2),
                               m.group(3), m.group(4)))
    expr = _SUM_RANGE_RE.sub(_sum_repl, expr)

    # 其余区域引用（如 F7:F7）→ 单格取值
    expr = _RANGE_RE.sub(
        lambda m: repr(values.get((m.group(1).upper(), int(m.group(2))), 0.0)),
        expr)

    # 函数名（避免与单元格引用冲突）
    expr = expr.replace('ROUND(', '_round(').replace('IF(', '_if(')

    # 单元格引用 → _v('F', 7)
    expr = _CELL_RE.sub(
        lambda m: "_v('%s',%d)" % (m.group(1).upper(), int(m.group(2))), expr)

    # 比较运算符
    expr = expr.replace('<>', '!=').replace('=', '==')

    env = {
        '__builtins__': {},
        '_v': lambda col, row: values.get((col, row), 0.0),
        '_round': excel_round,
        '_if': lambda cond, a, b: a if cond else b,
    }
    return float(eval(expr, env, {}))


# ============================================================
# 保费基数核对表生成器
# ============================================================

class BaseChecker:
    """
    保费基数核对表生成器。

    使用步骤:
      1. checker = BaseChecker()
      2. checker.load_a1411(path); checker.load_a2411(path)
      3. checker.load_account_balance(path)      # 科目余额（源数据）
      4. checker.load_irp001(path)               # 被保险存款申报表
      5. checker.load_attachments(att2, att3, att4)
      6. checker.set_adjustments(items)        # 调整项列表
      7. checker.compute(template_path, output_path)
    """

    def __init__(self):
        """初始化生成器，清空全部数据容器"""
        self.a1411_data = {}      # A1411 {科目代码: 余额（万元）}
        self.a2411_data = {}      # A2411 {科目代码: 余额（万元）}
        self.acct_data = {}       # 科目余额表 {科目代码: 余额（元）}
        self.irp001_p = {}        # IRP001 {(章节关键词, 项目关键词): 本金（元）}
        self.irp001_i = {}        # IRP001 {(章节关键词, 项目关键词): 应付利息（元）}
        self.att2_value = None    # 附件2（住房公积金）本金合计（元）——参考值
        self.att3_value = None    # 附件3（社保基金）本金合计（元）——参考值
        self.att4_p = None        # 附件4（高管人员）本金合计（元）
        self.att4_i = None        # 附件4（高管人员）应付利息合计（元）
        self.adjustments = []     # 调整项列表（见 set_adjustments）
        self.not_found_a1411 = []
        self.not_found_a2411 = []
        self.not_found_acct = []
        self.messages = []        # 运行日志

    # ----------------------------------------------------------
    # 数据加载
    # ----------------------------------------------------------

    def _load_monthly_report(self, filepath, target, not_found):
        """
        加载 A1411 / A2411 月报表。

        读取规则：数据从第4行开始，A列 = 指标编号，C列 = 期末余额（万元）。

        参数:
            filepath (str): 月报表路径
            target (dict): 目标字典
            not_found (list): 未匹配代码记录列表
        """
        ws = ExcelSheet(str(filepath))
        target.clear()
        del not_found[:]
        for r in range(3, ws.nrows):
            code = str(ws.cell_value(r, 0) or '').strip()
            if code:
                target[code] = to_float(ws.cell_value(r, 2))

    def load_a1411(self, filepath):
        """加载 A1411 月报表（人民币，万元）"""
        self._load_monthly_report(filepath, self.a1411_data, self.not_found_a1411)
        self.messages.append(f"A1411: {len(self.a1411_data)} 个指标")

    def load_a2411(self, filepath):
        """加载 A2411 月报表（外币，万元）"""
        self._load_monthly_report(filepath, self.a2411_data, self.not_found_a2411)
        self.messages.append(f"A2411: {len(self.a2411_data)} 个指标")

    def load_account_balance(self, filepath):
        """
        加载科目余额表（源数据）。

        支持多 sheet；数据从第4行开始，A列 = 科目代码，
        G列 = 期末借方余额，H列 = 期末贷方余额（取非零的一方）。

        参数:
            filepath (str): 科目余额表路径（如 科目余额查询报表.xls）
        """
        xls = _ExcelSheets(str(filepath))
        self.acct_data.clear()
        for sname in xls.sheet_names():
            ws = xls.sheet_by_name(sname)
            for r in range(3, ws.nrows):
                code = str(ws.cell_value(r, 0)).strip()
                if not code:
                    continue
                debit = to_float(ws.cell_value(r, 6))
                credit = to_float(ws.cell_value(r, 7))
                if code not in self.acct_data:
                    self.acct_data[code] = debit if debit != 0 else credit
        self.messages.append(
            f"科目余额表: {len(self.acct_data)} 个科目, "
            f"{len(xls.sheet_names())} 个 sheet")

    def load_irp001(self, filepath):
        """
        加载 IRP001 被保险存款申报表的本金/应付利息。

        B 列为行标签（"一、境内存款"、"（一）个人存款"等），
        C 列 = 本金，D 列 = 应付利息。

        参数:
            filepath (str): IRP001 申报表路径
        """
        ws = ExcelSheet(str(filepath))
        self.irp001_p.clear()
        self.irp001_i.clear()
        for _key, section, item in IRP001_ITEM_KEYS:
            row = self._find_irp_row(ws, section, item)
            if row is None:
                self.messages.append(f"  ⚠ IRP001 中未找到「{item}」行")
                continue
            self.irp001_p[(section, item)] = to_float(ws.cell_value(row, 2))
            self.irp001_i[(section, item)] = to_float(ws.cell_value(row, 3))
        self.messages.append(f"IRP001: 读取 {len(self.irp001_p)} 个项目")

    @staticmethod
    def _find_irp_row(ws, section, item):
        """在 IRP001 中按「章节 + 项目」定位行号（未找到返回 None）"""
        start = 0
        if section:
            for r in range(ws.nrows):
                if section in str(ws.cell_value(r, 1)):
                    start = r + 1
                    break
        for r in range(start, ws.nrows):
            if item in str(ws.cell_value(r, 1)):
                return r
        return None

    def load_attachments(self, attachment2=None, attachment3=None, attachment4=None):
        """
        加载附件数据。

        附件2（住房公积金）/附件3（社保基金）仅作为参考值记录到日志
        —— 表中公积金/社保金额按模板公式对子项求和，不直接取附件数。
        附件4（高管人员存款）直接填入「三、高级管理人员在本机构存款」行。

        参数:
            attachment2 (str, optional): 附件2 路径
            attachment3 (str, optional): 附件3 路径
            attachment4 (str, optional): 附件4 路径
        """
        if attachment2:
            self.att2_value = _read_attachment_total(str(attachment2))
            self.messages.append(f"附件2(住房公积金)本金合计: {self.att2_value:,.2f} 元（参考）")
        if attachment3:
            self.att3_value = _read_attachment_total(str(attachment3))
            self.messages.append(f"附件3(社保基金)本金合计: {self.att3_value:,.2f} 元（参考）")
        if attachment4:
            p, i = _read_attachment_summary(str(attachment4))
            self.att4_p, self.att4_i = p, i
            self.messages.append(f"附件4(高管人员): 本金={p:,.2f}, 利息={i:,.2f} 元")

    def set_adjustments(self, items):
        """
        设置调整项列表（可自由增删的项目）。

        参数:
            items (list[dict]): 调整项列表，每项含
                name    名称（自定义）
                scope   用途：ADJ_SCOPE_ACCT（会计科目核对）
                             / ADJ_SCOPE_DEP（存款归属核对）
                target  对象：个人存款 / 单位存款
                field   字段：本金 / 应付利息
                mode    方式：增加 / 扣减
                amount  金额（元）
            也可传入 build_default_adjustments() 生成的标准口径列表。
        """
        self.adjustments = normalize_adjustments(items)

    # ----------------------------------------------------------
    # 取数
    # ----------------------------------------------------------

    def lookup_monthly(self, code_str):
        """
        按子项代码从 A1411/A2411 取数（万元）。

        代码格式："12MI1/22MI1"（A1411代码 / A2411代码）、
        "22MMU"（仅 A2411）、"12MI1"（仅 A1411）。

        参数:
            code_str (str): 子项代码

        返回:
            float: A1411 与 A2411 取值之和（万元）
        """
        c1, c2 = _parse_sub_code(code_str)
        v1 = v2 = 0.0
        if c1:
            if c1 in self.a1411_data:
                v1 = self.a1411_data[c1]
            else:
                self.not_found_a1411.append(c1)
        if c2:
            if c2 in self.a2411_data:
                v2 = self.a2411_data[c2]
            else:
                self.not_found_a2411.append(c2)
        return v1 + v2

    def lookup_acct(self, code):
        """
        按科目代码从科目余额表取数（元）。

        参数:
            code (str): 科目代码

        返回:
            float: 科目余额（元），未匹配返回 0
        """
        code = _clean_code(code)
        if not code:
            return 0.0
        if code in self.acct_data:
            return self.acct_data[code]
        self.not_found_acct.append(code)
        return 0.0

    # ----------------------------------------------------------
    # 主流程
    # ----------------------------------------------------------

    def compute(self, template_path, output_path):
        """
        生成保费基数核对表。

        参数:
            template_path (str): 模板路径（templates/保费基数核对表（模板）.xlsx）
            output_path (str): 输出路径（.xlsx）

        返回:
            list[str]: 运行日志

        异常:
            ValueError: output_path 为空
        """
        if not output_path:
            raise ValueError("output_path 是必填参数")

        wb = openpyxl.load_workbook(str(template_path))
        ws_main = wb[MAIN_SHEET_NAME] if MAIN_SHEET_NAME in wb.sheetnames else wb.active
        ws_check = wb[CHECK_SHEET_NAME] if CHECK_SHEET_NAME in wb.sheetnames else None

        # ---------- Step 1: 解析模板结构 ----------
        rows_data = {}   # 行号 -> 该行结构信息
        formulas = {}    # (列字母, 行号) -> 公式文本
        for r in range(1, ws_main.max_row + 1):
            info = {
                'f_code': _clean_code(ws_main.cell(r, 3).value),   # C列 子项代码
                'f_name': _text(ws_main.cell(r, 4).value),         # D列 子项名称
                'sign': _text(ws_main.cell(r, 5).value),           # E列 符号
                'p_code': _clean_code(ws_main.cell(r, 7).value),   # G列 科目本金代码
                'p_name': _text(ws_main.cell(r, 8).value),         # H列 科目名称
                'i_code': _clean_code(ws_main.cell(r, 10).value),  # J列 利息代码
                'i_name': _text(ws_main.cell(r, 11).value),        # K列 科目名称
                'b_name': _text(ws_main.cell(r, 2).value),         # B列 行列名称
            }
            rows_data[r] = info
            for col in FORMULA_COLUMNS:
                v = ws_main.cell(r, column_index_from_string(col)).value
                if isinstance(v, str) and v.strip().startswith('='):
                    formulas[(col, r)] = v.strip()

        # ---------- Step 2: 明细行取数 ----------
        values = {}
        filled_f = filled_i = filled_l = 0

        # 存款归属（F 列）：按子项代码从 A1411/A2411 取数（万元）
        for r, info in rows_data.items():
            if _is_sub_code(info['f_code']):
                values[('F', r)] = round(self.lookup_monthly(info['f_code']), 2)
                filled_f += 1

        # 会计科目（I/L 列）：按科目代码从科目余额表取数（元）
        for r, info in rows_data.items():
            if info['p_code'] and info['p_code'].isdigit():
                values[('I', r)] = round(self.lookup_acct(info['p_code']), 2)
                filled_i += 1
            if info['i_code'] and info['i_code'].isdigit():
                values[('L', r)] = round(self.lookup_acct(info['i_code']), 2)
                filled_l += 1

        self.messages.append(
            f"明细取数: 存款归属(F)={filled_f}行, 科目本金(I)={filled_i}行, "
            f"科目利息(L)={filled_l}行")

        # ----- 调整项（可自定义的项目列表）-----
        # 会计科目核对类项目 → 汇总后叠加到对应合计项（见 Step 3b）
        # 存款归属核对类项目 → 汇总写入 Q9/Q10，供校验页比较使用
        # 说明：模板公式按"调整项行=0"计算基础值，调整额在公式计算之后再叠加，
        #       因此调整额不受模板调整项行是否可定位影响
        acct_adj, dep_adj = summarize_adjustments(self.adjustments)
        row_vals = adjustment_row_values(self.adjustments)
        adj_rows = self._find_adjustment_rows(rows_data, formulas)
        for target in ADJ_TARGETS:
            if target not in adj_rows:
                self.messages.append(
                    f"  ⚠ 未在模板中定位到「{target}」调整项行（仅影响表内展示）")
        self._adj_rows = adj_rows
        self.messages.append(
            f"调整项: 共{len(normalize_adjustments(self.adjustments))}项, "
            f"会计科目核对 {len(acct_adj)} 个口径, 存款归属核对 个人={dep_adj.get('个人存款', 0):,.2f} "
            f"单位={dep_adj.get('单位存款', 0):,.2f}")
        for text in self._adjustment_log_lines():
            self.messages.append(text)

        # 高级管理人员存款（附件4）：定位"高级管理人员"明细行填入 I/L。
        # 注意模板中既有汇总行（I11=I109，带公式）也有明细行（R109），
        # 只填明细行，避免覆盖汇总公式。
        if self.att4_p is not None:
            target = None
            for r, info in rows_data.items():
                if ('高级管理人员' in info['p_name'] or '高级管理人员' in info['i_name']) \
                        and ('I', r) not in formulas:
                    target = r
                    break
            if target is None:
                self.messages.append("  ⚠ 未找到「高级管理人员」明细行，附件4 未填入")
            else:
                values[('I', target)] = round(self.att4_p, 2)
                values[('L', target)] = round(self.att4_i or 0.0, 2)
                self.messages.append(f"高级管理人员存款 (R{target}): 按附件4填入")

        # IRP001（N/O 列）：按行标签匹配申报表项目
        for r, info in rows_data.items():
            name = info['p_name']
            if not name:
                continue
            for kw, section, item in IRP001_ITEM_KEYS:
                if kw in name:
                    values[('N', r)] = self.irp001_p.get((section, item), 0.0)
                    values[('O', r)] = self.irp001_i.get((section, item), 0.0)
                    break

        # 调整值写入 Q 列（差额说明列的前四格）：既是表格内容，
        # 也是校验页公式（存款归属 vs IRP001）的输入项
        #   Q7/Q8 = 个人侧净调整额（会计科目：本金 / 应付利息）
        #   Q9/Q10 = 存款归属核对时从 IRP001 本金中扣减的净额（个人 / 单位）
        q_values = {
            'c1': acct_adj.get(('个人存款', '本金'), 0.0),
            'c2': acct_adj.get(('个人存款', '应付利息'), 0.0),
            'c3': dep_adj.get('个人存款', 0.0),
            'c4': dep_adj.get('单位存款', 0.0),
        }
        # 写入位置：Q9/Q10 优先取校验页公式实际引用的 Q 单元格（模板驱动），
        # 未引用时回退默认位置；Q7/Q8 为附件5 口径的个人侧净调整额（信息项）
        adj_cells = dict(ADJ_CELLS)
        if ws_check is not None:
            q_ref_re = re.compile(r"Q\$?(\d+)")
            for r in range(1, ws_check.max_row + 1):
                label = str(ws_check.cell(r, 1).value or '')
                if '个人存款' in label:
                    key = 'c3'
                elif '单位存款' in label:
                    key = 'c4'
                else:
                    continue
                for c in range(2, ws_check.max_column + 1):
                    v = ws_check.cell(r, c).value
                    if isinstance(v, str) and v.strip().startswith('='):
                        m = q_ref_re.search(v)
                        if m:
                            adj_cells[key] = (int(m.group(1)), 'Q')
                            break
        for key, (row, col) in adj_cells.items():
            values[(col, row)] = round(q_values[key], 2)

        # ---------- Step 3: 按模板公式计算合并 ----------
        # 所有公式（存款归属列 F、会计科目列 I/L/M、IRP001 列 N/O/P 与
        # 校验页）均来自模板，程序不内置任何公式
        all_formulas = dict(formulas)

        # 存款归属核对不支持"合计项"调整，提示忽略
        for it in normalize_adjustments(self.adjustments):
            if it['scope'] == ADJ_SCOPE_DEP and it['target'] not in ADJ_DEP_TARGETS:
                self.messages.append(
                    f"  ⚠ 存款归属核对仅支持个人/单位存款调整项，"
                    f"已忽略「{it['name'] or it['target']}」({it['target']})")

        # ---- 3a) 先按模板公式计算基础结果（调整项行按 0 计）----
        self.messages.extend(self._check_formula_gaps(all_formulas, rows_data, formulas))
        self._eval_formula_pass(values, all_formulas)
        self.messages.append(f"公式计算: 主表 {len(all_formulas)} 个公式（基础结果）")

        # ---- 3b) 再按调整项对相应合计项进行调整，得到最终结果 ----
        total_rows = self._find_total_rows(rows_data, formulas, ws_check)
        adjusted_cells = self._apply_adjustments(values, all_formulas,
                                                 total_rows, acct_adj)
        if adjusted_cells:
            # 调整后重算受影响的派生公式（合计列 M、被保险存款合计等）
            remaining = {k: v for k, v in all_formulas.items()
                         if k not in adjusted_cells}
            self._eval_formula_pass(values, remaining)

        # ---- 3c) 写入"调整项"行展示值（等于已叠加的调整额，供与附件5 对照）----
        for target, r in adj_rows.items():
            values[('I', r)] = round(row_vals.get((target, '本金'), 0.0), 2)
            values[('L', r)] = round(row_vals.get((target, '应付利息'), 0.0), 2)

        # ---------- 校验页（差额）计算 ----------
        check_msgs = []
        if ws_check is not None:
            # 校验页比较用"写入文件后的数值"（两位小数），避免浮点尾差
            # 导致 L7=O7 这类精确比较误判为"不一致"
            rounded = {k: (round(v, 2) if isinstance(v, float) else v)
                       for k, v in values.items()}
            check_values = {}
            for r in range(1, ws_check.max_row + 1):
                for c in range(1, ws_check.max_column + 1):
                    v = ws_check.cell(r, c).value
                    letter = get_column_letter(c)
                    if isinstance(v, str) and v.strip().startswith('='):
                        check_values[(letter, r)] = _eval_formula(v, rounded)
            check_msgs = self._describe_checks(ws_check, check_values)

        # ---------- Step 4: 固化数值并保存 ----------
        # 数值固化为静态值（模板公式被覆盖），保留模板原有的字体/边框/对齐，
        # 仅在模板未设置数字格式时补一个千分位两位小数格式。
        # 写入统一走 _safe_set_value：模板存在合并单元格，需定位到左上角。
        for (col, r), value in values.items():
            if value is None:
                continue
            c_idx = column_index_from_string(col)
            _safe_set_value(ws_main, r, c_idx, round(float(value), 2),
                            number_format=_num_fmt(ws_main, r, c_idx))

        if ws_check is not None:
            for (letter, r), value in check_values.items():
                c_idx = column_index_from_string(letter)
                _safe_set_value(ws_check, r, c_idx, round(float(value), 2),
                                number_format=_num_fmt(ws_check, r, c_idx))

        self._annotate_adjustments(ws_main, adj_rows)
        self._annotate_total_adjustments(ws_main)
        self._clear_stale_notes(ws_main)
        # 生成信息（写在标题单元格批注里），便于辨别文件是否为最新一次生成
        _append_comment(ws_main, 2, 1,
                        f"本表由『存款保险申报表审核工具』于 "
                        f"{time.strftime('%Y-%m-%d %H:%M:%S')} 生成；"
                        f"调整项 {len(normalize_adjustments(self.adjustments))} 项"
                        f"（明细见相关合计单元格批注）")

        wb.save(str(output_path))
        wb.close()

        # ---------- 诊断信息 ----------
        self.messages.extend(check_msgs)
        for name, lst in (('A1411', self.not_found_a1411),
                          ('A2411', self.not_found_a2411),
                          ('科目余额表', self.not_found_acct)):
            if lst:
                uq = list(dict.fromkeys(lst))
                self.messages.append(
                    f"{name} 未匹配代码({len(uq)})（按0处理）: {', '.join(uq[:6])}")
        try:
            size = os.path.getsize(output_path)
            stamp = time.strftime('%Y-%m-%d %H:%M:%S',
                                  time.localtime(os.path.getmtime(output_path)))
            self.messages.append(f"输出: {output_path}（生成于 {stamp}，{size:,} 字节）")
            self.messages.append(
                "提示: 若表内数值与本次日志不符，说明打开的是以前生成的旧文件"
                "（旧文件不会自动更新，请重新生成或打开上面的文件）")
        except OSError:
            self.messages.append(f"输出: {output_path}")
        return self.messages

    # ----------------------------------------------------------
    # 调整项：行定位 / 日志 / 批注
    # ----------------------------------------------------------

    @staticmethod
    def _find_adjustment_rows(rows_data, formulas=None):
        """
        定位模板中的会计科目"调整项"行（个人侧 / 单位侧各一行）。

        定位策略（依次尝试）：
          1. 按科目名称：模板中两行名称含"调整项"，
             带"（＋）"的为个人侧、带"（－）"的为单位侧
          2. 按公式结构（模板名称被改掉/删掉时兜底）：
             · 个人侧 = 个人存款合计 SUM 区域的最后一行
             · 单位侧 = 单位存款合计公式中被减去的、且无科目代码的行

        参数:
            rows_data (dict): {行号: 行结构信息}
            formulas (dict, optional): {(列, 行): 公式}（结构兜底定位用）

        返回:
            dict: {'个人存款': 行号, '单位存款': 行号}
        """
        # ---- 1) 按名称（要求该行没有科目代码，避免覆盖科目行数值）----
        def has_code(row):
            info = rows_data.get(row) or {}
            for key in ('p_code', 'i_code'):
                code = info.get(key) or ''
                if code.isdigit():
                    return True
            return False

        cands = []
        for r, info in sorted(rows_data.items()):
            if '调整项' not in (info['p_name'] or info['i_name']):
                continue
            if has_code(r):
                continue
            cands.append(r)
        rows = {}
        for r in cands:
            label = rows_data[r]['p_name'] or rows_data[r]['i_name']
            if '＋' in label or '+' in label:
                rows.setdefault('个人存款', r)
            elif '－' in label or '-' in label:
                rows.setdefault('单位存款', r)
        rest = [r for r in cands if r not in rows.values()]
        for key in ADJ_TARGETS:
            if key not in rows and rest:
                rows[key] = rest.pop(0)
        return rows

    def _adjustment_log_lines(self):
        """按用途分组输出调整项明细，便于运行日志核对。"""
        lines = []
        items = normalize_adjustments(self.adjustments)
        for scope in ADJ_SCOPES:
            detail = describe_adjustments(items, scope=scope)
            if detail:
                lines.append(f"  [{scope}]")
                lines.extend('    ' + ln for ln in detail.split('\n'))
        return lines

    def _clear_stale_notes(self, ws_main):
        """
        清空模板中遗留的"差额说明"文字（模板示例来自其他机构，直接沿用会误导）。

        仅清空 Q 列第 7 行起的文字；Q7~Q10 的调整值不受影响。
        """
        cleared = []
        for row in range(7, ws_main.max_row + 1):
            if row in (7, 8, 9, 10):      # Q7~Q10 为调整值
                continue
            try:
                r, c = _resolve_merged(ws_main, row, 17)
            except Exception:
                continue
            cell = ws_main.cell(r, c)
            value = cell.value
            if isinstance(value, str) and value.strip():
                cell.value = None
                cleared.append(r)
        if cleared:
            self.messages.append(
                "已清空模板遗留的差额说明文字（第 " +
                "、".join(str(r) for r in cleared) +
                " 行 Q 列）——模板中的说明来自其他机构，请按本期实际差异填写")

    def _eval_formula_pass(self, values, formulas):
        """
        按依赖关系求值一组公式（就地写入 values）。

        参数:
            values (dict): {(列, 行): 数值}，就地更新
            formulas (dict): {(列, 行): 公式文本}
        """
        pending = dict(formulas)
        while pending:
            progressed = False
            for key in list(pending):
                refs = _formula_refs(pending[key])
                if all(ref in values or ref not in formulas for ref in refs):
                    values[key] = _eval_formula(pending[key], values)
                    del pending[key]
                    progressed = True
            if not progressed:
                for key, formula in pending.items():
                    values[key] = _eval_formula(formula, values)
                self.messages.append("WARNING: 部分公式引用无法解析，按已知值计算")
                break

    def _apply_adjustments(self, values, formulas, total_rows, acct_adj):
        """
        按调整项对"会计科目核对"各合计项进行调整（最终结果 = 模板公式结果 + 调整额）。

        参数:
            values (dict): {(列, 行): 数值}，就地叠加
            formulas (dict): 全部公式（用于确认合计行是公式单元格）
            total_rows (dict): {合计项: 行号}
            acct_adj (dict): {(对象, 字段): 净调整额}

        返回:
            set[tuple]: 被调整的单元格 {(列, 行)}
        """
        adjusted = {}
        for (target, field), net in sorted(acct_adj.items()):
            col = 'I' if field == '本金' else 'L'
            row = total_rows.get(target)
            if not net:
                continue
            if row is None or (col, row) not in formulas:
                self.messages.append(
                    f"  ⚠ 未找到「{target}」合计行/公式，{field}调整 "
                    f"{net:,.2f} 元未生效")
                continue
            base = values.get((col, row), 0.0)
            values[(col, row)] = base + net
            expr = str(formulas.get((col, row), '')).lstrip('=')
            adjusted[(col, row)] = (target, field, base, net, expr)
        self._adj_details = adjusted
        for (col, row), (target, field, base, net, expr) in sorted(
                adjusted.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            self.messages.append(
                f"合计项调整: {target}·{field} 模板公式 {col}{row}"
                + (f"={expr}" if expr else "")
                + f" → {base:,.2f}；{'加' if net > 0 else '扣减'}调整项 "
                f"{abs(net):,.2f} = 最终 {base + net:,.2f}（第{row}行{col}列）")
        return set(adjusted)

    @staticmethod
    def _find_total_rows(rows_data, formulas, ws_check=None):
        """
        定位各"合计项"所在的会计科目合计行。

        优先从**模板的校验页公式**推导（模板驱动）：校验页每行 A 列是项目
        名称，右侧公式形如 IF(核对表!I7=核对表!N7,0,1)，其中引用的主表
        I 列单元格即该项目的会计科目合计行；
        校验页缺失或结构变化时，回退按会计科目名称关键词匹配
        （ADJ_TOTAL_KEYWORDS）。

        参数:
            rows_data (dict): {行号: 行结构信息}
            formulas (dict): {(列, 行): 公式}
            ws_check: openpyxl 校验页工作表（可选）

        返回:
            dict: {合计项名称: 行号}
        """
        out = {}
        if ws_check is not None:
            ref_re = re.compile(r"!\$?([A-Za-z]{1,3})\$?(\d+)")
            for r in range(1, ws_check.max_row + 1):
                label = str(ws_check.cell(r, 1).value or '').strip()
                if not label:
                    continue
                target = None
                for kw in ADJ_TOTAL_KEYWORDS:
                    if kw in label:
                        target = kw
                        break
                if target is None or target in out:
                    continue
                for c in range(2, ws_check.max_column + 1):
                    v = ws_check.cell(r, c).value
                    if not (isinstance(v, str) and v.strip().startswith('=')):
                        continue
                    for m in ref_re.finditer(v):
                        if m.group(1).upper() != 'I':
                            continue
                        row = int(m.group(2))
                        if ('I', row) in formulas:
                            out.setdefault(target, row)
                            break

        for target, kw in ADJ_TOTAL_KEYWORDS.items():
            if target in out:
                continue
            for r, info in sorted(rows_data.items()):
                name = info['p_name'] or info['i_name']
                if kw in name and ('I', r) in formulas:
                    out[target] = r
                    break
        return out

    @staticmethod
    def _check_formula_gaps(all_formulas, rows_data, formulas):
        """
        检查存款归属公式（F 列）是否引用了"没有数据来源"的行。

        工具只按模板中的子项代码（C 列）取数；若公式引用了既无子项代码
        又无公式的行（例如财政性拆分行没写在公式引用的行上），该处取数为
        0，这里提示使用者修改模板，而不是由程序自行补数。
        会计科目列与 IRP001 列的空行属正常（历史值行、调整项行、
        高管/不予保险行、对方申报表数据行），不在此检查。

        参数:
            all_formulas (dict): 全部公式 {(列, 行): 公式}
            rows_data (dict): {行号: 行结构信息}
            formulas (dict): 模板原始公式（用于区分"公式格"与"空格"）

        返回:
            list[str]: 日志消息
        """
        msgs = []
        seen = set()
        for key, expr in all_formulas.items():
            if key[0] != 'F':      # 只检查存款归属列：会计科目/IRP001 列的空行属正常
                continue
            for ref in _formula_refs(expr):
                if ref in formulas or ref in seen:
                    continue
                col, row = ref
                if col != 'F':
                    continue
                info = rows_data.get(row) or {}
                if _is_sub_code(info.get('f_code') or ''):
                    continue
                seen.add(ref)
                msgs.append(
                    f"  ⚠ 模板公式 {key[0]}{key[1]} 引用了 {col}{row}，"
                    f"但该行在模板中没有子项代码（C 列），取数为 0"
                    f"——如需取数请在模板中给该行补上子项代码")
        return msgs

    def _annotate_total_adjustments(self, ws_main):
        """
        给被调整过的"合计项"单元格加批注，写明
        「模板公式结果 ± 调整项 = 最终值」及调整项明细。
        """
        for (col, row), (target, field, base, net, expr) in sorted(
                getattr(self, '_adj_details', {}).items(),
                key=lambda kv: (kv[0][1], kv[0][0])):
            detail = describe_adjustments(self.adjustments,
                                          scope=ADJ_SCOPE_ACCT,
                                          target=target, field=field)
            c_idx = column_index_from_string(col)
            _append_comment(
                ws_main, row, c_idx,
                f"取值＝模板公式 {col}{row}" + (f"={expr}" if expr else "")
                + f" 结果 {base:,.2f} "
                + (f"+ {net:,.2f}" if net > 0 else f"- {abs(net):,.2f}")
                + f"（{target}·{field} 调整项）= 最终 {base + net:,.2f}\n"
                f"调整项明细：\n{detail}")

    def _annotate_adjustments(self, ws_main, adj_rows):
        """
        在表内写批注：调整项明细。

        写入位置：会计科目"调整项"行（I/L 列）与存款归属调整值 Q9/Q10；
        模板原有批注（口径说明）保留，明细追加在其后。

        参数:
            ws_main: openpyxl Worksheet
            adj_rows (dict): {'个人存款': 行号, '单位存款': 行号}
        """
        items = normalize_adjustments(self.adjustments)
        for target, r in adj_rows.items():
            sign_note = ('（个人合计 SUM 累加该行，故填应加进个人的金额）'
                         if target == '个人存款'
                         else '（模板公式对该行做扣减，故填应从单位扣减的金额）')
            # 模板若缺少该行名称，则在输出中补上（便于阅读与后续校验定位）
            name_cols = ((8, '个体工商户调整项（本金）', 9),
                         (11, '个体工商户调整项（应付利息）', 12))
            for name_col, base_name, _ in name_cols:
                r_name, c_name = _resolve_merged(ws_main, r, name_col)
                if not str(ws_main.cell(r_name, c_name).value or '').strip():
                    sign = '（＋）' if target == '个人存款' else '（－）'
                    _safe_set_value(ws_main, r, name_col, base_name + sign)
                    self.messages.append(
                        f"  第{r}行模板缺少调整项名称，已在输出中补为"
                        f"「{base_name + sign}」")
            for col_idx, field, scope in ((9, '本金', ADJ_SCOPE_ACCT),
                                          (12, '应付利息', ADJ_SCOPE_ACCT)):
                detail = describe_adjustments(items, scope=scope, target=target,
                                              field=field)
                if detail:
                    _append_comment(ws_main, r, col_idx,
                                    f"取值＝{target}·{field} 调整项合计{sign_note}：\n{detail}")
        for target, row in (('个人存款', 9), ('单位存款', 10)):
            detail = describe_adjustments(items, scope=ADJ_SCOPE_DEP,
                                          target=target)
            if detail:
                _append_comment(ws_main, row, 17,
                                f"存款归属核对调整（从 IRP001 本金中按方式增减）：\n{detail}")

    # ----------------------------------------------------------
    # 校验页差额说明
    # ----------------------------------------------------------

    @staticmethod
    def _describe_checks(ws_check, check_values):
        """
        把校验页的计算结果整理成日志。

        校验页列定义（模板固定）：
          A=项目名称, B=存款归属本金(标志), C=存款归属差额(万元),
          D=会计科目本金(标志), E=会计科目利息(标志), F=会计科目合计(标志)
        标志：0=一致，1=不一致。
        最后一行为"被保险存款合计"行。
        """
        msgs = ["---------- 校验页（差额）----------"]
        for r in range(3, ws_check.max_row + 1):
            name = _text(ws_check.cell(r, 1).value)
            if not name:
                continue
            flag_b = check_values.get(('B', r))
            diff_c = check_values.get(('C', r))
            flags = [check_values.get(('D', r)), check_values.get(('E', r)),
                     check_values.get(('F', r))]
            if flag_b is None and diff_c is None:
                continue
            dep_txt = (f"存款归属差额={diff_c:,.2f}万元"
                       if diff_c is not None else "存款归属未校验")
            if flag_b is not None and abs(flag_b) > 0.001:
                dep_txt += " [不一致]"
            acct_txt = "会计科目" + ("一致" if all(
                (f is None or abs(f) <= 0.001) for f in flags) else "不一致")
            msgs.append(f"  {name}: {dep_txt}; {acct_txt}")
        return msgs


# ============================================================
# 工具函数
# ============================================================

def _append_comment(ws, row, col, text):
    """
    在单元格批注中追加文本（保留模板原有批注）。

    合并单元格会定位到区域左上角；批注写入失败时静默跳过，
    不影响主流程。

    参数:
        ws: openpyxl Worksheet
        row (int): 1-indexed 行号
        col (int): 1-indexed 列号
        text (str): 追加的批注内容
    """
    try:
        r, c = _resolve_merged(ws, row, col)
        cell = ws.cell(r, c)
        old = cell.comment.text.strip() if cell.comment and cell.comment.text else ''
        cell.comment = Comment((old + '\n——\n' + text) if old else text, '审核工具')
    except Exception:
        pass


def _text(value):
    """单元格值 → 去空白字符串（None → ''）"""
    return str(value).strip() if value is not None else ''


def _is_sub_code(text):
    """
    判断 C 列取值是否为存款归属子项代码（如 "12MI1/22MI1"、"22MMU"）。

    表头行的"子项代码""存款归属核对"等中文文本不算代码，
    避免被误当成代码去取数。
    """
    return bool(re.fullmatch(r'[0-9A-Za-z]+(/[0-9A-Za-z]+)?', text or ''))


def _num_fmt(ws, row, col):
    """
    写入数值时的数字格式：模板已设格式则保持不变（None 表示不改），
    否则补一个千分位两位小数格式。
    """
    r, c = _resolve_merged(ws, row, col)
    cur = ws.cell(r, c).number_format
    return None if cur not in (None, 'General') else '#,##0.00'


def _clean_code(value):
    """
    单元格值 → 科目代码字符串。

    兼容 Excel 把代码当数字读出的情况（"2003.0" → "2003"）；
    非数字文本（如"三、高级管理人员在本机构存款"）原样返回，
    由调用方用 isdigit() 判断是否为科目代码。
    """
    text = _text(value)
    if not text:
        return ''
    if text.endswith('.0') and text[:-2].isdigit():
        return text[:-2]
    if re.fullmatch(r'\d+\.0+', text):
        return text.split('.')[0]
    return text

def read_adjustments(filepath):
    """
    从《保费基数核对表》（如银行报送的附件5）读取调整值 c1~c4。

    调整值存放在主表 Q 列（差额说明列）的前四格：
      Q7=c1 个体工商户本金、Q8=c2 个体工商户利息、
      Q9=c3 个人久悬户、Q10=c4 单位久悬户。

    参数:
        filepath (str): 保费基数核对表路径（.xls / .xlsx 均可）

    返回:
        dict: {'c1','c2','c3','c4'}；读取失败或四格均为 0 时返回 None
    """
    try:
        xls = _ExcelSheets(str(filepath))
        names = xls.sheet_names()
        ws = xls.sheet_by_name(MAIN_SHEET_NAME if MAIN_SHEET_NAME in names
                               else names[0])
    except Exception:
        return None
    out = {}
    for key, row in (('c1', 7), ('c2', 8), ('c3', 9), ('c4', 10)):
        try:
            out[key] = to_float(ws.cell_value(row - 1, 16))   # Q 列 = 第17列
        except (IndexError, ValueError):
            out[key] = 0.0
    return out if any(out.values()) else None
