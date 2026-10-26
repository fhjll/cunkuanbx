# -*- coding: utf-8 -*-
"""
=============================================================================
保费基数核对表生成引擎（合并模板）
=============================================================================

功能说明：
  按《保费基数核对表（模板）》把三块数据填进同一张表：
    1. 存款归属核对 —— 取自 A1411 / A2411 月报表（万元）
    2. 会计科目核对 —— 取自科目余额表（元）
    3. 被保险存款申报表（IRP001）—— 取自 IRP001 申报表（元）
  各数据块所在列按模板表头自动识别（旧版模板：存款归属 C~F、会计科目 G~M、
  IRP001 N~P；新版附5模板去掉了"符号"列，依次为 C~E、F~L、M~O），
  再按模板中的公式对子项数据逐级计算合并（个人/单位/财政性/非存款类/
  高级管理人员/不予保险/被保险存款合计），并把结果固化为静态数值。

与旧版两个生成器的关系：
  旧版「存款归属表」和「科目余额核对表」是两张独立的表，
  现合并为一张《保费基数核对表》，本模块取代：
    - deposit_calculator.DepositCalculator 的模板填充逻辑（F 列）
    - account_checker.AccountChecker 的模板填充逻辑（I / L 列）

计算顺序（与模板公式一致）：
  Step 1 读模板：按表头识别各数据块列位，解析每行的子项代码、科目代码，
          以及主表中的全部公式
  Step 2 明细取数：按代码逐项取数（存款归属取 A1411/A2411，会计科目取科目余额表）
  Step 3 公式计算：按模板公式（含跨行引用、SUM 区域、IF/ROUND）逐条求值
  Step 4 固化输出：把计算值与校验页差额写入 .xlsx（覆盖公式，全部为静态数值）

调整值（可自定义的调整项列表）：
  每个调整项包含：名称、用途、对象、字段、方式、金额（元）
    用途：会计科目核对（写入表内"调整项"行，参与会计科目合计）
          存款归属核对（直接加到核对表"存款归属核对"对应项本金上）
    对象：个人存款 / 单位存款
    字段：本金 / 应付利息（用途=存款归属核对时只比较本金，固定为"本金"）
    方式：增加 / 扣减
  对象可选：
    · 个人存款 / 单位存款 / 个人↔单位（一增一减，个体工商户用）
    · 会计科目核对表其他合计项：财政性存款 / 非存款类金融机构存款 /
      高级管理人员 / 不予保险的其他存款（模板无专门调整行，
      调整额直接并入该合计项的公式，上层汇总同步生效）
  项目可自由增删（标准项除外，其 locked=True 不可删除）；
  标准口径（个体工商户 + 久悬户轧差值）为：
    · 个体工商户本金      会计科目核对 个人↔单位 本金     增加（个人＋、单位－）
    · 个体工商户利息      会计科目核对 个人↔单位 应付利息 增加（个人＋、单位－）
    · 久悬户轧差值        会计科目核对 个人存款 本金     增加
    · 久悬户轧差值        存款归属核对 个人存款 本金     增加
  差额说明列第7/8行保存"由单位调整至个人"的个人侧净额（本金/利息），供外部核对。

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

# ------------------------------------------------------------
# 模板列布局（按表头自动识别，兼容新旧两版附5模板）
# ------------------------------------------------------------
# 旧版合并模板：存款归属 C/D/E/F（F=本金，E=符号）；
#              会计科目 G/H/I（本金）与 J/K/L（利息），合计 M；
#              IRP001 N/O/P，差额说明 Q。
# 新版附5模板（2026年10月版 / 模板1019）：去掉"符号"列，其后各列左移一列：
#              存款归属 C/D/E（E=本金）；
#              会计科目 F/G/H（本金）与 I/J/K（利息），合计 L；
#              IRP001 M/N/O，差额说明 P。
# 实际列位由 detect_template_layout() 读表头文字确定，业务代码不再写死列号。
DEFAULT_LAYOUT = {
    'dep_code': 'C', 'dep_name': 'D', 'dep_amount': 'F',
    'acct_p_code': 'G', 'acct_p_name': 'H', 'acct_p_amount': 'I',
    'acct_i_code': 'J', 'acct_i_name': 'K', 'acct_i_amount': 'L',
    'acct_total': 'M', 'irp_p': 'N', 'irp_i': 'O', 'irp_total': 'P',
    'note': 'Q',
}

# 调整值所在行：7~10 行依次为 c1/c2（个人侧净调整额，会计科目）
# 与 c3/c4（存款归属核对调整额），列位见 layout['note']。
ADJ_ROWS = {'c1': 7, 'c2': 8, 'c3': 9, 'c4': 10}


def detect_template_layout(get_value):
    """
    按表头文字识别主表各数据块所在列（兼容旧版/新版附5模板）。

    识别依据（模板第 4~6 行的表头文字）：
      · 「存款归属核对」区块 → 子项代码列、本金（金额）列
      · 「会计科目核对」区块 → 本金/利息各一组 代码/科目名称/金额 列、合计列
      · 「被保险存款申报表」区块 → 本金/利息/合计列
      · 「差额说明」列 → 调整值 c1~c4 所在列

    参数:
        get_value (callable): get_value(row, col) → 单元格值（1-indexed）

    返回:
        dict: 各数据块的列字母（键见 DEFAULT_LAYOUT）；
              识别失败时回退为旧版模板列位（DEFAULT_LAYOUT）
    """
    layout = dict(DEFAULT_LAYOUT)

    def text(r, c):
        v = get_value(r, c)
        return str(v).strip() if v is not None else ''

    def find_in_row(row, keyword, start, end):
        for c in range(start, end + 1):
            if keyword in text(row, c):
                return c
        return None

    def find_exact(row, keyword, start, end):
        for c in range(start, end + 1):
            if text(row, c) == keyword:
                return c
        return None

    try:
        # ---- 1) 第 4 行：四个区块的起始列（兼容新旧表名） ----
        # 新表名：C4=A1411、2411合计表，F4=存款归属表；旧表名：存款归属核对/会计科目核对
        dep_start = (find_in_row(4, 'A1411', 1, 40)
                     or find_in_row(4, '存款归属', 1, 40))
        acct_start = (find_in_row(4, '会计科目', 1, 40)
                      or find_in_row(4, '存款归属表', 1, 40))
        irp_start = find_in_row(4, '被保险存款申报表', 1, 40)
        note_start = find_in_row(4, '差额说明', 1, 40)
        if not (dep_start and acct_start and irp_start and note_start):
            return layout
        if not (dep_start < acct_start < irp_start <= note_start):
            return layout

        # ---- 2) 存款归属：子项代码列 + 本金金额列 ----
        code_col = None
        for r in (5, 6):
            code_col = find_in_row(r, '子项代码', dep_start, acct_start - 1)
            if code_col:
                break
        amount_col = None
        for r in (6, 5):
            amount_col = find_exact(r, '本金', dep_start, acct_start - 1)
            if amount_col:
                break
        name_col = None
        for r in (5, 6):
            name_col = find_in_row(r, '子项名称', dep_start, acct_start - 1)
            if name_col:
                break
        if code_col:
            layout['dep_code'] = get_column_letter(code_col)
        if name_col or code_col:
            layout['dep_name'] = get_column_letter(
                name_col if name_col else code_col + 1)
        if amount_col:
            layout['dep_amount'] = get_column_letter(amount_col)

        # ---- 3) 会计科目核对：本金/利息两组 代码/科目名称/金额 ----
        codes, names, amounts = [], [], []
        for c in range(acct_start, irp_start):
            v = text(6, c)
            if '代码' in v:
                codes.append(c)
            elif '科目名称' in v:
                names.append(c)
            elif '金额' in v:
                amounts.append(c)
        if len(codes) >= 2 and len(names) >= 2 and len(amounts) >= 2:
            layout['acct_p_code'] = get_column_letter(codes[0])
            layout['acct_i_code'] = get_column_letter(codes[1])
            layout['acct_p_name'] = get_column_letter(names[0])
            layout['acct_i_name'] = get_column_letter(names[1])
            layout['acct_p_amount'] = get_column_letter(amounts[0])
            layout['acct_i_amount'] = get_column_letter(amounts[1])
        total_col = find_exact(5, '合计', acct_start, irp_start - 1)
        if total_col:
            layout['acct_total'] = get_column_letter(total_col)

        # ---- 4) IRP001：本金 / 利息 / 合计 ----
        irp_p = find_exact(5, '本金', irp_start, note_start - 1)
        irp_i = find_exact(5, '利息', irp_start, note_start - 1)
        irp_t = find_exact(5, '合计', irp_start, note_start - 1)
        if irp_p:
            layout['irp_p'] = get_column_letter(irp_p)
        if irp_i:
            layout['irp_i'] = get_column_letter(irp_i)
        if irp_t:
            layout['irp_total'] = get_column_letter(irp_t)

        # ---- 5) 差额说明列（调整值 c1~c4 所在列） ----
        layout['note'] = get_column_letter(note_start)
    except Exception:
        return dict(DEFAULT_LAYOUT)

    # 列序校验：识别结果必须保持 存款归属 < 会计科目 < IRP001 < 说明列
    order = [
        'dep_code', 'dep_amount', 'acct_p_code', 'acct_p_name',
        'acct_p_amount', 'acct_i_code', 'acct_i_name', 'acct_i_amount',
        'acct_total', 'irp_p', 'irp_i', 'irp_total', 'note',
    ]
    idx = [column_index_from_string(layout[k]) for k in order]
    return layout if idx == sorted(idx) else dict(DEFAULT_LAYOUT)

# ------------------------------------------------------------
# 调整项取值范围
# ------------------------------------------------------------
ADJ_SCOPE_ACCT = '存款归属表'          # 科目余额表口径（写入表内调整项行）
ADJ_SCOPE_DEP = 'A1411、2411合计表'     # A1411/A2411 口径（直接加到该表对应项本金）
ADJ_SCOPES = (ADJ_SCOPE_ACCT, ADJ_SCOPE_DEP)
# 旧版用途名称 → 新表名（兼容已保存的调整值配置.json）
ADJ_SCOPE_ALIASES = {
    '会计科目核对': ADJ_SCOPE_ACCT,
    '存款归属核对': ADJ_SCOPE_DEP,
}
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
    ('久悬户轧差值', ADJ_SCOPE_ACCT, '个人存款', '本金', '增加', True),
    ('久悬户轧差值', ADJ_SCOPE_DEP, '个人存款', '本金', '增加', True),
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
    按标准口径生成调整项列表。

    标准项：个体工商户本金/利息 + 久悬户轧差值（会计科目侧、存款归属侧各一条）。
    久悬户轧差值按"增加"处理：会计科目侧加到会计科目核对表本金，
    存款归属侧加到存款归属核对表本金（见 compute 的 Step 3b）。

    参数:
        c1 (float): 个体工商户本金（个人增、单位减）
        c2 (float): 个体工商户应付利息（个人增、单位减）
        c3 (float): 久悬户轧差值 —— 会计科目核对侧（元，方式=增加）
        c4 (float): 久悬户轧差值 —— 存款归属核对侧（元，方式=增加）

    返回:
        list[dict]: 4 条标准调整项
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
        scope = ADJ_SCOPE_ALIASES.get(it.get('scope'), it.get('scope', ADJ_SCOPE_ACCT))
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
          dep  (dict): {对象: 净额} —— 存款归属核对：直接加到核对表存款归属
                       对应项本金（元；方式=增加 记正数、方式=扣减 记负数）
    """
    acct = {}
    dep = {}
    for it in normalize_adjustments(items):
        sign = -1.0 if it['mode'] == '扣减' else 1.0
        amount = sign * it['amount']
        if it['scope'] == ADJ_SCOPE_DEP:
            if it['target'] not in ADJ_DEP_TARGETS:
                continue      # 存款归属核对只支持个人/单位（由 compute 记录提示）
            # 存款归属核对：调整值直接加到核对表"存款归属核对"对应项上
            if it['target'] == ADJ_TARGET_BOTH:
                dep['个人存款'] = dep.get('个人存款', 0.0) + amount
                dep['单位存款'] = dep.get('单位存款', 0.0) - amount
            else:
                dep[it['target']] = dep.get(it['target'], 0.0) + amount
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
    # 非 SUM 的区间（如 =...+H19:H20-H21）：展开为区间的每个单元格，
    # 否则依赖排序会漏掉这些单元格，导致取数顺序错误
    for m in _RANGE_RE.finditer(body):
        for col in _range_cols(m.group(1), m.group(3)):
            for rr in range(int(m.group(2)), int(m.group(4)) + 1):
                refs.add((col, rr))
    body = _RANGE_RE.sub('', body)
    for m in _CELL_RE.finditer(body):
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

    # 其余区域引用（如 H19:H20）→ 区间求和。
    # 说明：单元格公式里直接写多格区间在 Excel 中会返回 #VALUE!（数组公式），
    # 这里按"区间求和"容错，避免整格取数丢失；单格区间行为不变。
    def _range_repl(m):
        return repr(_sum_range(values, m.group(1), m.group(2),
                               m.group(3), m.group(4)))
    expr = _RANGE_RE.sub(_range_repl, expr)

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

        # ---------- 模板列布局识别（新旧附5模板自适应） ----------
        layout = detect_template_layout(
            lambda r, c: ws_main.cell(r, c).value)
        self.layout = layout
        col = {k: column_index_from_string(v) for k, v in layout.items()}
        self.messages.append(
            "模板布局: 存款归属{}/会计科目{}+{}/IRP001{}/说明列{}"
            .format(layout['dep_amount'], layout['acct_p_amount'],
                    layout['acct_i_amount'], layout['irp_p'], layout['note']))

        # ---------- Step 1: 解析模板结构 ----------
        rows_data = {}   # 行号 -> 该行结构信息
        formulas = {}    # (列字母, 行号) -> 公式文本
        for r in range(1, ws_main.max_row + 1):
            info = {
                'f_code': _clean_code(ws_main.cell(r, col['dep_code']).value),
                'f_name': _text(ws_main.cell(r, col['dep_name']).value),
                'sign': '',
                'p_code': _clean_code(ws_main.cell(r, col['acct_p_code']).value),
                'p_name': _text(ws_main.cell(r, col['acct_p_name']).value),
                'i_code': _clean_code(ws_main.cell(r, col['acct_i_code']).value),
                'i_name': _text(ws_main.cell(r, col['acct_i_name']).value),
                'b_name': _text(ws_main.cell(r, 2).value),   # B列 行列名称
            }
            rows_data[r] = info

        # 读取主表全部公式（列位随模板变化，不再按固定列清单读取）
        # 兼容普通公式与 openpyxl 的 ArrayFormula（数组公式）：
        # Excel 中误操作产生的数组公式不应让整格取数直接丢失。
        for r in range(1, ws_main.max_row + 1):
            for c in range(1, ws_main.max_column + 1):
                v = ws_main.cell(r, c).value
                if isinstance(v, str):
                    text = v.strip()
                else:
                    text = str(getattr(v, 'text', '') or '').strip()
                if text.startswith('='):
                    formulas[(get_column_letter(c), r)] = text

        # ---------- Step 2: 明细行取数 ----------
        values = {}
        filled_f = filled_i = filled_l = 0

        # 存款归属（本金列）：按子项代码从 A1411/A2411 取数（万元）
        for r, info in rows_data.items():
            if _is_sub_code(info['f_code']):
                values[(layout['dep_amount'], r)] = round(
                    self.lookup_monthly(info['f_code']), 2)
                filled_f += 1

        # 会计科目（本金/利息金额列）：按科目代码从科目余额表取数（元）
        for r, info in rows_data.items():
            if info['p_code'] and info['p_code'].isdigit():
                values[(layout['acct_p_amount'], r)] = round(
                    self.lookup_acct(info['p_code']), 2)
                filled_i += 1
            if info['i_code'] and info['i_code'].isdigit():
                values[(layout['acct_i_amount'], r)] = round(
                    self.lookup_acct(info['i_code']), 2)
                filled_l += 1

        self.messages.append(
            f"明细取数: 存款归属({layout['dep_amount']})={filled_f}行, "
            f"科目本金({layout['acct_p_amount']})={filled_i}行, "
            f"科目利息({layout['acct_i_amount']})={filled_l}行")

        # ----- 调整项（可自定义的项目列表）-----
        # 会计科目核对类项目 → 汇总后叠加到对应合计项（见 Step 3b）
        # 存款归属核对类项目 → 汇总后直接加到存款归属对应项本金（见 Step 3b）
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

        # 高级管理人员存款（附件4）：定位"高级管理人员"明细行填入会计科目金额列。
        # 注意模板中既有汇总行（带公式）也有明细行，只填明细行，避免覆盖汇总公式。
        if self.att4_p is not None:
            target = None
            for r, info in rows_data.items():
                if ('高级管理人员' in info['p_name'] or '高级管理人员' in info['i_name']) \
                        and (layout['acct_p_amount'], r) not in formulas:
                    target = r
                    break
            if target is None:
                self.messages.append("  ⚠ 未找到「高级管理人员」明细行，附件4 未填入")
            else:
                values[(layout['acct_p_amount'], target)] = round(self.att4_p, 2)
                values[(layout['acct_i_amount'], target)] = round(
                    self.att4_i or 0.0, 2)
                self.messages.append(f"高级管理人员存款 (R{target}): 按附件4填入")

        # IRP001（本金/利息列）：按行标签匹配申报表项目
        for r, info in rows_data.items():
            name = info['p_name']
            if not name:
                continue
            for kw, section, item in IRP001_ITEM_KEYS:
                if kw in name:
                    values[(layout['irp_p'], r)] = self.irp001_p.get(
                        (section, item), 0.0)
                    values[(layout['irp_i'], r)] = self.irp001_i.get(
                        (section, item), 0.0)
                    break

        # 调整值写入"差额说明"列前四格（旧模板 Q7~Q10，新模板 P7~P10）：
        #   c1/c2 = 个人侧净调整额（会计科目：本金 / 应付利息）
        #   c3/c4 = 固定为 0：久悬户调整现在直接写入存款归属表对应项，
        #           与 IRP001 的比较改为直接比较，不再从 IRP001 扣减 Q9/Q10
        q_values = {
            'c1': acct_adj.get(('个人存款', '本金'), 0.0),
            'c2': acct_adj.get(('个人存款', '应付利息'), 0.0),
            'c3': 0.0,
            'c4': 0.0,
        }
        # 写入位置：c3/c4 优先取校验页公式实际引用的说明列单元格（模板驱动），
        # 未引用时回退默认行；c1/c2 为附件5 口径的个人侧净调整额（信息项）
        adj_cells = {key: (row, layout['note']) for key, row in ADJ_ROWS.items()}
        if ws_check is not None:
            note_ref_re = re.compile(
                r"(?<![A-Za-z0-9])" + re.escape(layout['note'])
                + r"\$?(\d+)(?![0-9])")
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
                        m = note_ref_re.search(v)
                        if m:
                            adj_cells[key] = (int(m.group(1)), layout['note'])
                            break
        for key, (row, col_letter) in adj_cells.items():
            values[(col_letter, row)] = round(q_values[key], 2)

        # ---------- Step 3: 按模板公式计算合并 ----------
        # 所有公式（含存款归属/会计科目/IRP001 三块与校验页）均来自模板，
        # 程序不内置任何计算公式
        all_formulas = dict(formulas)

        # 存款归属核对不支持"合计项"调整，提示忽略
        for it in normalize_adjustments(self.adjustments):
            if it['scope'] == ADJ_SCOPE_DEP and it['target'] not in ADJ_DEP_TARGETS:
                self.messages.append(
                    f"  ⚠ A1411、2411合计表调整仅支持个人/单位存款调整项，"
                    f"已忽略「{it['name'] or it['target']}」({it['target']})")

        # ---- 3a) 先按模板公式计算基础结果（调整项行按 0 计）----
        self.messages.extend(self._check_formula_gaps(all_formulas, rows_data, formulas))
        self._eval_formula_pass(values, all_formulas)
        self.messages.append(f"公式计算: 主表 {len(all_formulas)} 个公式（基础结果）")

        # ---- 3b) 再按调整项对相应合计项进行调整，得到最终结果 ----
        #   · 会计科目核对类：调整会计科目合计（个人/单位/其他合计项）
        #   · 存款归属核对类：直接调整核对表"存款归属核对"列对应项本金
        total_rows = self._find_total_rows(rows_data, formulas, ws_check)
        adjusted_cells = self._apply_adjustments(values, all_formulas,
                                                 total_rows, acct_adj)
        dep_rows = self._find_deposit_total_rows(rows_data, formulas,
                                                 ws_check, layout['dep_amount'],
                                                 layout.get('dep_name'))
        adjusted_cells |= self._apply_deposit_adjustments(values, dep_rows,
                                                          dep_adj)
        if adjusted_cells:
            # 调整后重算受影响的派生公式（合计列、境内合计、校验页差额等）
            remaining = {k: v for k, v in all_formulas.items()
                         if k not in adjusted_cells}
            self._eval_formula_pass(values, remaining)

        # ---- 3c) 写入"调整项"行展示值（等于已叠加的调整额，供与附件5 对照）----
        for target, r in adj_rows.items():
            values[(layout['acct_p_amount'], r)] = round(
                row_vals.get((target, '本金'), 0.0), 2)
            values[(layout['acct_i_amount'], r)] = round(
                row_vals.get((target, '应付利息'), 0.0), 2)

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

        仅清空"差额说明"列第 7 行起的文字；第 7~10 行的调整值不受影响。
        """
        note_idx = column_index_from_string(self.layout['note'])
        adj_rows = set(ADJ_ROWS.values())
        cleared = []
        for row in range(7, ws_main.max_row + 1):
            if row in adj_rows:      # 第 7~10 行为调整值
                continue
            try:
                r, c = _resolve_merged(ws_main, row, note_idx)
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
                f" 行 {self.layout['note']} 列）——模板中的说明来自其他机构，"
                "请按本期实际差异填写")

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
        layout = self.layout
        for (target, field), net in sorted(acct_adj.items()):
            col = (layout['acct_p_amount'] if field == '本金'
                   else layout['acct_i_amount'])
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
    def _find_deposit_total_rows(rows_data, formulas, ws_check, dep_col,
                                 dep_name_col=None):
        """
        定位"A1411、2411合计表"中个人/单位存款的合计行。

        优先按"子项名称"列（模板 D 列，用户模板表头为"子项名称"）的名称
        精确匹配"（一）个人存款 / （二）单位存款"；这样即使合计行公式被改成
        数组公式、或公式解析失败，调整值仍能加到正确的行上。
        名称匹配不到时，再退回校验页公式引用定位。

        参数:
            rows_data (dict): {行号: 行结构信息}
            formulas (dict): {(列, 行): 公式}（保底校验用）
            ws_check: openpyxl 校验页工作表（可为 None）
            dep_col (str): 存款归属金额列字母
            dep_name_col (str, optional): 子项名称列字母（默认 D）

        返回:
            dict: {'个人存款': 行号, '单位存款': 行号}
        """
        out = {}
        want = {'个人存款': '（一）个人存款', '单位存款': '（二）单位存款'}
        # 1) 按子项名称列匹配（仅取境内汇总行，第一处即可）
        for target, label in want.items():
            for r, info in sorted(rows_data.items()):
                name = str(info.get('f_name') or '').replace(' ', '')
                if name == label:
                    out[target] = r
                    break
        if len(out) == 2:
            return out
        # 2) 退回校验页公式引用（个人/单位存款行的公式中引用的存款归属单元格）
        if ws_check is not None:
            ref_re = re.compile(r"!\$?([A-Za-z]{1,3})\$?(\d+)")
            for r in range(1, ws_check.max_row + 1):
                label = str(ws_check.cell(r, 1).value or '')
                if '个人存款' in label:
                    target = '个人存款'
                elif '单位存款' in label:
                    target = '单位存款'
                else:
                    continue
                if target in out:
                    continue
                for c in range(2, ws_check.max_column + 1):
                    v = ws_check.cell(r, c).value
                    text = v if isinstance(v, str) else str(
                        getattr(v, 'text', '') or '')
                    if not text.strip().startswith('='):
                        continue
                    for m in ref_re.finditer(text):
                        if m.group(1).upper() != dep_col:
                            continue
                        row = int(m.group(2))
                        if (dep_col, row) in formulas or (dep_col, row) in rows_data:
                            out.setdefault(target, row)
                            break
                    if target in out:
                        break
        return out

    def _apply_deposit_adjustments(self, values, dep_rows, dep_adj):
        """
        把"存款归属核对"类调整项直接加到核对表存款归属列对应项本金。

        说明：
          · 调整项金额单位为元，核对表存款归属金额单位为万元，按 /10000 换算；
          · 加完后四舍五入到整数万元（与存款归属表"万元"口径一致），
            避免校验页用 ROUND(IRP001/10000) 比较时因小数产生假差异；
          · 个人↔单位：个人加、单位减（summarize_adjustments 已折算成净额）。

        参数:
            values (dict): {(列, 行): 数值}，就地叠加
            dep_rows (dict): {'个人存款': 行号, '单位存款': 行号}
            dep_adj (dict): {对象: 净额（元）}

        返回:
            set[tuple]: 被调整的单元格 {(列, 行)}
        """
        adjusted = {}
        dep_col = self.layout['dep_amount']
        for target, amount in dep_adj.items():
            if not amount:
                continue
            row = dep_rows.get(target)
            if row is None:
                self.messages.append(
                    f"  ⚠ 未找到「{target}」A1411、2411合计表合计行，"
                    f"调整 {amount:,.2f} 元未生效")
                continue
            base = values.get((dep_col, row), 0.0)
            new_val = excel_round(base + amount / 10000.0, 0)
            values[(dep_col, row)] = new_val
            adjusted[(dep_col, row)] = (target, base, amount)
        self._dep_adjust_details = adjusted
        for (col, row), (target, base, amount) in sorted(
                adjusted.items(), key=lambda kv: kv[0][1]):
            self.messages.append(
                f"A1411、2411合计表调整: {target}·本金 第{row}行{col}列 "
                f"{base:,.2f} 万元 {'+' if amount >= 0 else '-'} "
                f"{abs(amount) / 10000.0:,.4f} 万元 = "
                f"{values[(col, row)]:,.0f} 万元")
        return set(adjusted)

    def _find_total_rows(self, rows_data, formulas, ws_check=None):
        """
        定位各"合计项"所在的会计科目合计行。

        优先从**模板的校验页公式**推导（模板驱动）：校验页每行 A 列是项目
        名称，右侧公式形如 IF(核对表!H7=核对表!M7,0,1)，其中引用的主表
        会计科目本金金额列单元格即该项目的会计科目合计行；
        校验页缺失或结构变化时，回退按会计科目名称关键词匹配
        （ADJ_TOTAL_KEYWORDS）。

        参数:
            rows_data (dict): {行号: 行结构信息}
            formulas (dict): {(列, 行): 公式}
            ws_check: openpyxl 校验页工作表（可选）

        返回:
            dict: {合计项名称: 行号}
        """
        p_col = self.layout['acct_p_amount']
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
                        if m.group(1).upper() != p_col:
                            continue
                        row = int(m.group(2))
                        if (p_col, row) in formulas:
                            out.setdefault(target, row)
                            break

        for target, kw in ADJ_TOTAL_KEYWORDS.items():
            if target in out:
                continue
            for r, info in sorted(rows_data.items()):
                name = info['p_name'] or info['i_name']
                if kw in name and (p_col, r) in formulas:
                    out[target] = r
                    break
        return out

    def _check_formula_gaps(self, all_formulas, rows_data, formulas):
        """
        检查存款归属公式（本金金额列）是否引用了"没有数据来源"的行。

        工具只按模板中的子项代码（存款归属代码列）取数；若公式引用了既无
        子项代码又无公式的行（例如财政性拆分行没写在公式引用的行上），
        该处取数为 0，这里提示使用者修改模板，而不是由程序自行补数。
        会计科目列与 IRP001 列的空行属正常（历史值行、调整项行、
        高管/不予保险行、对方申报表数据行），不在此检查。

        参数:
            all_formulas (dict): 全部公式 {(列, 行): 公式}
            rows_data (dict): {行号: 行结构信息}
            formulas (dict): 模板原始公式（用于区分"公式格"与"空格"）

        返回:
            list[str]: 日志消息
        """
        dep_col = self.layout['dep_amount']
        msgs = []
        seen = set()
        for key, expr in all_formulas.items():
            if key[0] != dep_col:      # 只检查存款归属列
                continue
            for ref in _formula_refs(expr):
                if ref in formulas or ref in seen:
                    continue
                col, row = ref
                if col != dep_col:
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

        写入位置：会计科目"调整项"行（本金/利息金额列）与存款归属调整值
        （"差额说明"列第 9/10 行）；模板原有批注（口径说明）保留，明细追加其后。

        参数:
            ws_main: openpyxl Worksheet
            adj_rows (dict): {'个人存款': 行号, '单位存款': 行号}
        """
        layout = self.layout
        p_name_idx = column_index_from_string(layout['acct_p_name'])
        i_name_idx = column_index_from_string(layout['acct_i_name'])
        p_amt_idx = column_index_from_string(layout['acct_p_amount'])
        i_amt_idx = column_index_from_string(layout['acct_i_amount'])
        note_idx = column_index_from_string(layout['note'])
        items = normalize_adjustments(self.adjustments)
        for target, r in adj_rows.items():
            sign_note = ('（个人合计 SUM 累加该行，故填应加进个人的金额）'
                         if target == '个人存款'
                         else '（模板公式对该行做扣减，故填应从单位扣减的金额）')
            # 模板若缺少该行名称，则在输出中补上（便于阅读与后续校验定位）
            name_cols = ((p_name_idx, '个体工商户调整项（本金）'),
                         (i_name_idx, '个体工商户调整项（应付利息）'))
            for name_col, base_name in name_cols:
                r_name, c_name = _resolve_merged(ws_main, r, name_col)
                if not str(ws_main.cell(r_name, c_name).value or '').strip():
                    sign = '（＋）' if target == '个人存款' else '（－）'
                    _safe_set_value(ws_main, r, name_col, base_name + sign)
                    self.messages.append(
                        f"  第{r}行模板缺少调整项名称，已在输出中补为"
                        f"「{base_name + sign}」")
            for col_idx, field, scope in ((p_amt_idx, '本金', ADJ_SCOPE_ACCT),
                                          (i_amt_idx, '应付利息', ADJ_SCOPE_ACCT)):
                detail = describe_adjustments(items, scope=scope, target=target,
                                              field=field)
                if detail:
                    _append_comment(ws_main, r, col_idx,
                                    f"取值＝{target}·{field} 调整项合计{sign_note}：\n{detail}")
        for target, row in (('个人存款', ADJ_ROWS['c3']),
                            ('单位存款', ADJ_ROWS['c4'])):
            detail = describe_adjustments(items, scope=ADJ_SCOPE_DEP,
                                          target=target)
            if detail:
                _append_comment(ws_main, row, note_idx,
                                f"A1411、2411合计表调整（已直接加到该表对应项本金）：\n{detail}")

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

    调整值存放在主表"差额说明"列（旧模板 Q 列 / 新模板 P 列）的前四格：
      第7行=c1 个体工商户本金、第8行=c2 个体工商户利息、
      第9行=c3 个人久悬户、第10行=c4 单位久悬户。
    列位按表头文字自动识别，新旧模板均适用。

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
    layout = detect_template_layout(
        lambda r, c: ws.cell_value(r - 1, c - 1))
    note_idx = column_index_from_string(layout['note']) - 1   # 0-indexed
    out = {}
    for key, row in ADJ_ROWS.items():
        try:
            out[key] = to_float(ws.cell_value(row - 1, note_idx))
        except (IndexError, ValueError):
            out[key] = 0.0
    return out if any(out.values()) else None
