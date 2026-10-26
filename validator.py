# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险申报表审核工具 - 数据校验引擎
=============================================================================

本模块包含三个独立功能，用于验证不同数据源之间的数据一致性：

  1. validate_irp002()                — IRP002 数据校验
       - 内部校验：最高偿付限额检查、户均本息检查、分组 vs 合计检查
       - 交叉校验：IRP002 vs IRP001 各项目数据一致性检查

  2. validate_base_check_vs_irp001()  — 保费基数核对表 vs IRP001 校验
       - A1411、2411合计表（本金金额列，万元）与 IRP001 本金（精确至万元）比较
       - 存款归属表（本金/利息/合计，元）与 IRP001 本金/利息/合计比较
       （各列列位按核对表表头自动识别，新旧附5模板均支持）

  3. generate_validation_report()     — 生成全面校验报告（Excel）
       - 以 IRP001 为基准，逐项输出 A1411、2411合计表（B列）、
         存款归属表（C/D/E列）、IRP002（F/G/H列）的差额，不一致标红
       - 存款归属 vs IRP001（本金）：IRP001 本金先扣久悬户调整额再比较
         （个人=c3、单位=c4）；会计科目（被保险存款生成表）vs IRP001
         不做该调整，直接比较

校验逻辑（与《保费基数核对表（模板）》校验页一致）：
  - 存款归属核对：先对子项数据取数（A1411/A2411 按子项代码逐项取数），
    再按模板公式计算合并（个人/单位/财政性/非存款类各合计），
    最后与 IRP001 比较：
      个人存款：F7 = ROUND((IRP001个人本金 - c3) / 10000)
      单位存款：F8 = ROUND((IRP001单位本金 - c4) / 10000)
      财政性 / 非存款类：F11 = ROUND(IRP001本金 / 10000) 等
  - 会计科目核对：科目余额按科目代码逐项取数 → 按模板公式计算合并
    （个体工商户调整项已在表中通过"＋/－"行体现），
    再与 IRP001 的本金/应付利息/合计比较，绝对容差 0.001 元。
  - 每一组比较均采用"一致→0，不一致→差额"的报告口径。

调整项说明（可自定义增减，均以元为单位）：
  每个调整项含 名称 / 用途 / 对象 / 字段 / 方式 / 金额：
  - 用途=会计科目核对：生成核对表时写入表内"调整项"行
    （个人行按方式计入个人存款合计、单位行按方式计入单位存款合计），
    差额说明列第7/8行保存个人侧净调整额（本金/利息）供外部核对
  - 用途=存款归属核对：按方式汇总到差额说明列第9（个人）/10（单位）行，
    校验时从 IRP001 本金中扣减后再与存款归属（A1411/A2411）比较
  标准口径 = 个体工商户本金/利息（个人增、单位减）+ 个人/单位久悬户（扣减）

依赖模块：
  - openpyxl：读写 .xlsx 文件
  - xlrd：读取旧式 .xls 文件
  - common.py：公共工具（to_float, ExcelSheet）
=============================================================================
"""
import os
import xlrd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import column_index_from_string
from common import to_float, excel_round, ExcelSheet as _ExcelSheet
from base_checker import (ADJ_ROWS, adjustment_row_values,
                          detect_template_layout, normalize_adjustments)
from bank_config import resolve_template
from account_checker import _resolve_merged


# ============================================================
# 辅助函数
# ============================================================

def _get_cell_val(ws, r, c):
    """
    安全获取 _ExcelSheet 中指定单元格的值（0-indexed）。
    捕获 IndexError 和 ValueError 异常，异常时返回 0.0。

    参数:
        ws: _ExcelSheet 对象
        r (int): 行索引（0-indexed）
        c (int): 列索引（0-indexed）

    返回:
        float: 单元格的数值
    """
    try:
        v = ws.cell_value(r, c)
    except (IndexError, ValueError):
        return 0.0
    return to_float(v)


# ============================================================
# Excel 行号查找工具函数
# ============================================================

def _find_irp_row(ws, keyword):
    """
    在 _ExcelSheet（0-indexed）中按 B 列关键词查找行号。

    IRP 申报表的 B 列存放行标签，如"一、境内存款"、"（一）个人存款"等。

    参数:
        ws: _ExcelSheet 对象
        keyword (str): 要搜索的关键词

    返回:
        int or None: 匹配到的行号（0-indexed），未找到返回 None
    """
    for r in range(ws.nrows):
        b_val = str(ws.cell_value(r, 1)).strip()
        if keyword in b_val:
            return r
    return None


def _find_irp_cross_rows(ws, section_keyword, item_keyword):
    """
    在 IRP001 中按"章节 + 项目"关键词精确查找行。

    搜索策略：
      1. 先用 section_keyword 定位章节行
      2. 在章节行之后查找 item_keyword

    参数:
        ws: _ExcelSheet 对象
        section_keyword (str): 章节关键词，如 "一、境内存款"
        item_keyword (str): 项目关键词，如 "（一）个人存款"

    返回:
        int or None: 匹配到的行号（0-indexed）
    """
    start = 0
    if section_keyword:
        sec_row = _find_irp_row(ws, section_keyword)
        if sec_row is not None:
            start = sec_row + 1
    for r in range(start, ws.nrows):
        b_val = str(ws.cell_value(r, 1)).strip()
        if item_keyword in b_val:
            return r
    return None



# ============================================================
# 1. IRP002 数据校验
# ============================================================

def validate_irp002(filepath, irp001_path=None):
    """
    IRP002 数据校验。

    校验内容：
      (1) 户均本息合计校验
          检查 ≤最高偿付限额 和 >最高偿付限额 两组的户均金额是否合理。
          ≤ 限额组户均 ≤ 限额；> 限额组户均 > 限额。

      (2) 分组合计 vs 总合计校验
          检查组内各行：≤限额 + >限额 的合计是否等于总合计。
          允许 ±0.1% 的容差。

      (3) IRP002 vs IRP001 交叉校验（当 irp001_path 提供时）
          按关键词精确匹配 IRP002 和 IRP001 中的对应项目，
          比较本金和利息数据的一致性。

    参数:
        filepath (str): IRP002 申报表文件路径
        irp001_path (str, optional): IRP001 申报表路径（用于交叉校验）

    返回:
        list[str]: 校验结果消息列表
    """
    msgs = []
    try:
        ws = _ExcelSheet(str(filepath))
    except Exception as e:
        return [f"ERROR: 无法读取 IRP002: {e}"]

    # --- 动态查找最高偿付限额行 ---
    max_comp_row = _find_irp_row(ws, '最高偿付限额')
    if max_comp_row is not None:
        max_comp = _get_cell_val(ws, max_comp_row, 2)
        msgs.append(f"最高偿付限额 (第{max_comp_row}行): {max_comp:,.0f} 元")
    else:
        max_comp = 0
        msgs.append("WARNING: 未找到「最高偿付限额」行，设为0")

    # --- 动态查找校验数据行 ---
    irp002_keywords = [
        ('境内存款', "1.境内存款"),
        ('个人存款', "  1.1 个人存款"),
        ('单位存款', "  1.2 单位存款"),
        ('住房公积金', "     1.2.1 住房公积金存款"),
        ('社保基金', "     1.2.2 社保基金存款"),
        ('财政性存款', "  1.3 财政性存款"),
        ('境外存款', "2.境外存款"),
        ('高级管理人员', "3.高级管理人员在本机构存款"),
        ('被保险存款', "5.被保险存款"),
    ]
    check_rows = {}  # {row_number: label}
    for kw, label in irp002_keywords:
        r = _find_irp_row(ws, kw)
        if r is not None:
            check_rows[r] = label
        else:
            msgs.append(f"  WARNING: 未找到「{kw}」行")

    # --- 空数据检查 ---
    if check_rows:
        first_row = next(iter(check_rows))
        sample_total = _get_cell_val(ws, first_row, 2)
        sample_principal = _get_cell_val(ws, first_row, 3)
        if max_comp == 0 and sample_total == 0 and sample_principal == 0:
            msgs.append("WARNING: 读取到的关键数据均为0，可能是文件为空或格式不匹配！")
        elif max_comp == 0:
            msgs.append("WARNING: 最高偿付限额为0，请确认文件数据是否正确")

    # --- (1) 户均本息合计校验 ---
    msgs.append("\n========== (1) 户均本息合计 校验 ==========")
    for r, name in check_rows.items():
        cnt_lte = _get_cell_val(ws, r, 6)    # ≤限额户数
        sum_lte = _get_cell_val(ws, r, 9)    # ≤限额本息合计
        cnt_gt = _get_cell_val(ws, r, 10)    # >限额户数
        sum_gt = _get_cell_val(ws, r, 13)    # >限额本息合计

        avg_lte = sum_lte / cnt_lte if cnt_lte > 0 else 0
        avg_gt = sum_gt / cnt_gt if cnt_gt > 0 else 0

        ok_lte = avg_lte <= max_comp if cnt_lte > 0 else True
        ok_gt = avg_gt > max_comp if cnt_gt > 0 else True

        msgs.append(f"  {name}:")
        if cnt_lte > 0:
            msgs.append(f"    <= {max_comp:,.0f}: 户均={avg_lte:,.2f} [{'OK' if ok_lte else 'ERROR'}]")
        else:
            msgs.append(f"    <= {max_comp:,.0f}: 无数据")
        if cnt_gt > 0:
            msgs.append(f"    > {max_comp:,.0f}: 户均={avg_gt:,.2f} [{'OK' if ok_gt else 'ERROR'}]")
        else:
            msgs.append(f"    > {max_comp:,.0f}: 无数据")

    # --- (2) 分组合计 vs 总合计校验 ---
    msgs.append("\n========== (2) 分组合计 vs 总合计 校验 ==========")
    all_ok = True
    for r, name in check_rows.items():
        col_defs = [
            (2, 6, 10, "户数"),      # 合计列, ≤限额列, >限额列, 标签
            (3, 7, 11, "本金"),
            (4, 8, 12, "应付利息"),
            (5, 9, 13, "本息合计"),
        ]
        row_ok = True
        for c_t, c_l, c_g, label in col_defs:
            total = _get_cell_val(ws, r, c_t)
            lte = _get_cell_val(ws, r, c_l)
            gt = _get_cell_val(ws, r, c_g)
            s = lte + gt
            ok = abs(s - total) <= max(1, abs(total * 0.001))
            if not ok:
                if row_ok:
                    msgs.append(f"  {name}: ERROR")
                    row_ok = False
                    all_ok = False
                msgs.append(f"    {label}: {lte:,.2f} + {gt:,.2f} = {s:,.2f} != {total:,.2f}")
    if all_ok:
        msgs.append("  所有行校验通过")
    else:
        msgs.append("\n>>> 内部校验: 存在差异，请检查")

    # --- (3) IRP002 vs IRP001 交叉校验 ---
    if irp001_path:
        msgs.append("\n========== (3) IRP002 vs IRP001 交叉校验 ==========")
        try:
            ws_irp = _ExcelSheet(str(irp001_path))
        except Exception as e:
            msgs.append(f"  ERROR: 无法读取 IRP001: {e}")
            msgs.append("\n>>> IRP002 校验结果: 内部校验" + ("全部通过" if all_ok else "存在差异"))
            return msgs

        # 交叉校验映射表
        # (irp002关键词, irp001关键词, irp001章节关键词, irp002列, irp001列, 标签后缀)
        cross_pairs = [
            ('境内存款',   '境内存款',            '',  3, 2, "本金"),
            ('境内存款',   '境内存款',            '',  4, 3, "应付利息"),
            ('个人存款',   '（一）个人存款',       '一、境内存款',  3, 2, "本金"),
            ('个人存款',   '（一）个人存款',       '一、境内存款',  4, 3, "应付利息"),
            ('单位存款',   '（二）单位存款',       '一、境内存款',  3, 2, "本金"),
            ('单位存款',   '（二）单位存款',       '一、境内存款',  4, 3, "应付利息"),
            ('被保险存款', '被保险存款余额',       '',  3, 2, "本金"),
            ('被保险存款', '被保险存款余额',       '',  4, 3, "应付利息"),
        ]
        cross_ok = True
        for kw2, kw1, sec1, c2, c1, suffix in cross_pairs:
            r2 = _find_irp_row(ws, kw2)
            if r2 is None:
                msgs.append(f"  WARNING: IRP002中未找到「{kw2}」")
                continue
            if sec1:
                r1 = _find_irp_cross_rows(ws_irp, sec1, kw1)
            else:
                r1 = _find_irp_row(ws_irp, kw1)
            if r1 is None:
                msgs.append(f"  WARNING: IRP001中未找到「{kw1}」")
                continue

            label_item = f"{kw2}-{suffix}"
            val_irp002 = _get_cell_val(ws, r2, c2)
            val_irp001 = _get_cell_val(ws_irp, r1, c1)
            diff = abs(val_irp002 - val_irp001)
            ok = diff <= 0.001  # 绝对容差 0.001 元（严格一致）
            status = "OK" if ok else "ERROR"
            if not ok:
                cross_ok = False
            msgs.append(
                f"  {label_item}: IRP002[{r2},{c2}]={val_irp002:,.2f}  "
                f"IRP001[{r1},{c1}]={val_irp001:,.2f}  差额={diff:,.2f}  [{status}]")

        if cross_ok:
            msgs.append("  IRP002与IRP001交叉校验全部通过")
        else:
            msgs.append("  IRP002与IRP001交叉校验存在差异，请检查")

    # --- 结论 ---
    if all_ok and (not irp001_path):
        msgs.append("\n>>> IRP002 校验结果: 全部通过")
    elif all_ok and cross_ok:
        msgs.append("\n>>> IRP002 校验结果: 全部通过")
    else:
        msgs.append("\n>>> IRP002 校验结果: 存在差异，请检查")
    return msgs


# ============================================================
# 2. 保费基数核对表 vs IRP001 校验
# ============================================================

# 《保费基数核对表》主表名称（与 base_checker.MAIN_SHEET_NAME 保持一致）
BASE_SHEET_NAME = '保费基数核对表'


def _read_base_check_map(filepath):
    """
    读取《保费基数核对表》(已生成) 的三大数据块。

    各数据块所在列按模板表头自动识别（兼容旧版/新版附5模板）：
      存款归属核对：子项代码 / 子项名称 / 本金金额（万元）
      会计科目核对：本金（代码/科目名称/金额）、利息（代码/科目名称/金额）、合计
      被保险存款申报表（IRP001）：本金 / 应付利息 / 合计
      差额说明列前四格（第7~10行）：c1/c2 个人侧净调整额，
                                    c3/c4 存款归属核对调整额

    参数:
        filepath (str): 保费基数核对表_已生成.xlsx 路径

    返回:
        dict: {
          'deposit': {项目标签: 万元},
          'acct':    {项目标签: {'principal','interest','total'}},
          'irp':     {项目标签: {'principal','interest','total'}},
          'adjust':  {'c1','c2','c3','c4'},
          'rows':    {项目标签: 行号},
          'adjust_rows': {调整项名称: (本金, 利息)},
          'layout':  识别出的模板列布局,
        }；读取失败返回 {}
    """
    try:
        wb = openpyxl.load_workbook(str(filepath), data_only=True)
    except Exception:
        return {}
    ws = wb[BASE_SHEET_NAME] if BASE_SHEET_NAME in wb.sheetnames else wb.active

    layout = detect_template_layout(lambda r, c: ws.cell(r, c).value)
    col = {k: column_index_from_string(v) for k, v in layout.items()}
    out = {'deposit': {}, 'acct': {}, 'irp': {}, 'adjust': {}, 'rows': {},
           'adjust_rows': {}, 'layout': layout}

    # 调整值（"差额说明"列第 7~10 行）：
    #   c1/c2=个人侧净调整额（会计科目），c3/c4=存款归属核对调整额
    for key, row in ADJ_ROWS.items():
        out['adjust'][key] = to_float(ws.cell(row, col['note']).value)

    section = '境内'
    acct_p = acct_i = 0.0    # 会计科目 境内存款累计（个人+单位+财政性+非存款类）
    def cell_num(row, col):
        """取值：空单元格返回 None（用于区分"无数据"与"数值 0"）"""
        v = ws.cell(row, col).value
        return float(v) if isinstance(v, (int, float)) else None

    for r in range(1, ws.max_row + 1):
        b_name = str(ws.cell(r, 2).value or '').strip()
        d_name = str(ws.cell(r, col['dep_name']).value or '').strip()
        h_name = str(ws.cell(r, col['acct_p_name']).value or '').strip()
        k_name = str(ws.cell(r, col['acct_i_name']).value or '').strip()

        if '境外存款' in b_name:
            section = '境外'
        elif '境内存款' in b_name:
            section = '境内'

        # 存款归属核对（本金金额列，万元）
        # 同一标签可能先出现合计行、后出现明细行（新模板"住房公积金存款"
        # 合计在 D9，明细在 D54~D62），已取到数值后不再被明细行覆盖
        label = _deposit_label(section, b_name, d_name)
        if label:
            f_val = ws.cell(r, col['dep_amount']).value
            val = float(f_val) if isinstance(f_val, (int, float)) else None
            if label not in out['deposit'] or (
                    out['deposit'][label] is None and val is not None):
                out['deposit'][label] = val
            out['rows'].setdefault(label, r)
        # ---- 表内"调整项"行（本金/利息金额列）----
        if '调整项' in h_name or '调整项' in k_name:
            vals = (to_float(ws.cell(r, col['acct_p_amount']).value),
                    to_float(ws.cell(r, col['acct_i_amount']).value))
            for key in (h_name, k_name):
                if key:
                    out['adjust_rows'][key] = vals

        # ---- 会计科目核对（本金/利息/合计金额列，元）----
        if h_name or k_name:
            a_label = _acct_label(h_name)
            if a_label and a_label not in out['acct']:
                p_val = to_float(ws.cell(r, col['acct_p_amount']).value)
                i_val = to_float(ws.cell(r, col['acct_i_amount']).value)
                out['acct'][a_label] = {
                    'principal': p_val,
                    'interest': i_val,
                    'total': to_float(ws.cell(r, col['acct_total']).value) or (p_val + i_val),
                }
                out['rows'].setdefault(a_label, r)
                if a_label in ('（一）个人存款', '（二）单位存款',
                               '（三）财政性存款', '（四）非存款类金融机构存款'):
                    acct_p += p_val
                    acct_i += i_val
            # 财政性存款明细行（国库定期/代理财政）标签在 D 列
            sub_label = _acct_sub_label(d_name)
            if sub_label and sub_label not in out['acct']:
                s_p = to_float(ws.cell(r, col['acct_p_amount']).value)
                s_i = to_float(ws.cell(r, col['acct_i_amount']).value)
                out['acct'][sub_label] = {
                    'principal': s_p,
                    'interest': s_i,
                    'total': to_float(ws.cell(r, col['acct_total']).value) or (s_p + s_i),
                }
                out['rows'].setdefault(sub_label, r)

            # ---- IRP001（本金/利息/合计列，元）----
            irp_label = _acct_label(h_name)
            if irp_label and irp_label not in out['irp']:
                out['irp'][irp_label] = {
                    'principal': to_float(ws.cell(r, col['irp_p']).value),
                    'interest': to_float(ws.cell(r, col['irp_i']).value),
                    'total': to_float(ws.cell(r, col['irp_total']).value),
                }

    # 会计科目"一、境内存款"= 个人 + 单位 + 财政性 + 非存款类
    if acct_p or acct_i:
        out['acct']['一、境内存款'] = {
            'principal': acct_p, 'interest': acct_i, 'total': acct_p + acct_i,
        }
    # 存款归属"二、境外存款"= 三个境外分项之和
    ovs = [out['deposit'].get(k, 0.0) for k in
           ('境外（一）个人存款', '境外（二）单位存款', '境外（三）境外同业存放')]
    if any(ovs):
        out['deposit']['二、境外存款'] = sum(ovs)

    wb.close()
    return out


def _deposit_label(section, b_name, d_name):
    """
    由行标签推断"存款归属核对"所属项目。

    表内行标签有两类：一是计算行/合计行（名称在 D 列，
    如"（一）个人存款""境内存款合计"），二是明细行（不在本函数返回）。
    """
    if not d_name:
        return None
    if b_name.startswith('二、') or section == '境外':
        # 境外明细分行：D 列名称与境内重名，需加前缀区分
        if d_name in ('（一）个人存款', '（二）单位存款', '（三）境外同业存放'):
            return '境外' + d_name
        return None
    # 新旧模板的名称写法略有差异（"1. 国库定期存款" / "国库定期存款"），
    # 统一按关键词判断；明细行名称（如"住房公积金活期存款"）不在此匹配
    if '住房公积金' in d_name and ('其中' in d_name or d_name == '住房公积金存款'):
        return '住房公积金存款'
    if d_name == '社保基金存款':
        return '社保基金存款'
    if '境内存款合计' in d_name:
        return '一、境内存款'
    if '代理财政' in d_name:
        return '2.代理财政性存款'
    if '国库定期' in d_name:
        return '1.国库定期存款'
    mapping = {
        '（一）个人存款': '（一）个人存款',
        '（二）单位存款': '（二）单位存款',
        '（三）财政性存款': '（三）财政性存款',
        '（四）非存款类金融机构存款': '（四）非存款类金融机构存款',
    }
    return mapping.get(d_name)


def _acct_label(name):
    """由"会计科目核对"科目名称推断项目标签"""
    if not name:
        return None
    for kw, label in (
            ('（一）个人存款', '（一）个人存款'),
            ('（二）单位存款', '（二）单位存款'),
            ('（三）财政性存款', '（三）财政性存款'),
            ('（四）非存款类', '（四）非存款类金融机构存款'),
            ('高级管理人员', '高级管理人员'),
            ('不予保险', '四、不予保险的其他存款'),
            ('被保险存款合计', '被保险存款余额')):
        if kw in name:
            return label
    return None


def _acct_sub_label(d_name):
    """财政性存款明细行（D 列标签，新旧模板写法不同）"""
    if not d_name:
        return None
    if '代理财政' in d_name:
        return '2.代理财政性存款'
    if '国库定期' in d_name:
        return '1.国库定期存款'
    return None


def validate_base_check_vs_irp001(base_check_path, irp001_path=None,
                                  adjustments=None):
    """
    校验《保费基数核对表》(已生成) 与 IRP001 的数据一致性。

    校验逻辑（沿用原存款归属表/科目余额表两项校验的口径，合并为一次）：
      1. 先对子项数据取数、再按模板公式计算合并（已在生成表内完成）
      2. 存款归属核对（F 列，万元）与 IRP001 本金比较：
           各分项均按 ROUND(IRP001本金 / 10000) 比较（四舍五入至万元），
           差额≠0 即不一致。
           久悬户调整值已在生成核对表时直接加到存款归属表对应项，
           因此这里直接比较，不再扣减 Q9/Q10。
      3. 会计科目核对（I/L/M 列，元）与 IRP001 本金/利息/合计比较：
           绝对容差 0.001 元（"用途=会计科目核对"的调整项已在表内
           "调整项"行体现，此处不再重复调整）
      4. 当提供 IRP001 文件时，同时核对表中 N/O 列取数是否与申报表一致
      5. 当提供调整项时，核对表内"调整项"行与配置是否一致
         （不一致提示需重新生成核对表）

    参数:
        base_check_path (str): 保费基数核对表_已生成.xlsx 路径
        irp001_path (str, optional): IRP001 申报表路径（用于核对表中取数）
        adjustments (list[dict], optional):
            调整项列表（与生成核对表时相同），每项含
            name/scope/target/field/mode/amount；
            不传时仅按表内数值校验

    返回:
        list[str]: 校验结果消息列表
    """
    msgs = []
    data = _read_base_check_map(base_check_path)
    if not data:
        return [f"ERROR: 无法读取保费基数核对表: {base_check_path}"]

    dep, acct, irp = data['deposit'], data['acct'], data['irp']
    adjust = data['adjust']

    items = normalize_adjustments(adjustments) if adjustments else []

    msgs.append("========== 保费基数核对表 vs IRP001 校验 ==========")
    msgs.append("久悬户调整值已直接计入核对表（A1411、2411合计表/存款归属表），"
                "本次按表内数值直接比较")

    # ---- (1) 存款归属核对 vs IRP001 ----
    msgs.append("")
    msgs.append("--- （1）A1411、2411合计表 vs IRP001（模板校验页 B/C 列口径，单位：万元）---")
    dep_items = [
        ('（一）个人存款', '（一）个人存款'),
        ('（二）单位存款', '（二）单位存款'),
        ('（三）财政性存款', '（三）财政性存款'),
        ('（四）非存款类金融机构存款', '（四）非存款类金融机构存款'),
    ]
    dep_ok = True
    for label, key in dep_items:
        dep_val = dep.get(key)
        irp_val = irp.get(key, {}).get('principal', 0.0)
        if dep_val is None:
            msgs.append(f"  {label}: 表中未找到，跳过")
            continue
        dep_wan = excel_round(dep_val, 0)
        expect = excel_round(irp_val / 10000.0, 0)
        diff = dep_wan - expect
        ok = abs(diff) <= 0.001
        dep_ok = dep_ok and ok
        msgs.append(
            f"  {label}: 存款归属={dep_wan:,.0f}万元  "
            f"IRP001={expect:,.0f}万元  "
            f"差额={diff:,.0f}万元  [{'一致' if ok else '不一致'}]")

    # ---- (2) 会计科目核对 vs IRP001 ----
    msgs.append("")
    msgs.append("--- （2）存款归属表 vs IRP001（元，容差 0.001 元）---")
    acct_items = [
        ('（一）个人存款', '（一）个人存款'),
        ('（二）单位存款', '（二）单位存款'),
        ('（三）财政性存款', '（三）财政性存款'),
        ('（四）非存款类金融机构存款', '（四）非存款类金融机构存款'),
        ('三、高级管理人员在本机构存款', '高级管理人员'),
        ('四、不予保险的其他存款', '四、不予保险的其他存款'),
        ('被保险存款合计', '被保险存款余额'),
    ]
    acct_ok = True
    for label, key in acct_items:
        a = acct.get(key)
        i = irp.get(key)
        if a is None:
            msgs.append(f"  {label}: 表中未找到，跳过")
            continue
        if i is None:
            msgs.append(f"  {label}: 表中无 IRP001 数据，跳过")
            continue
        row_results = []
        for field, fname in (('principal', '本金'), ('interest', '应付利息'),
                             ('total', '合计')):
            av = a.get(field, 0.0)
            iv = i.get(field, 0.0)
            ok = _diff_ok(av, iv)
            acct_ok = acct_ok and ok
            row_results.append(
                f"{fname}: 核对={av:,.2f} IRP001={iv:,.2f} "
                f"差额={av - iv:,.2f} [{'一致' if ok else '不一致'}]")
        msgs.append(f"  {label}：" + "；".join(row_results))

    # ---- (3) 表内 IRP001 取数 vs 申报表文件 ----
    if irp001_path:
        msgs.append("")
        msgs.append("--- （3）表中 IRP001 列取数核对（与本表申报表文件比对）---")
        try:
            ws_irp = _ExcelSheet(str(irp001_path))
            file_ok = True
            for key, (kw, section) in {
                    '（一）个人存款': ('（一）个人存款', '一、境内存款'),
                    '（二）单位存款': ('（二）单位存款', '一、境内存款'),
                    '（三）财政性存款': ('（三）财政性存款', None),
                    '（四）非存款类金融机构存款': ('（四）非存款类', None),
                    '高级管理人员': ('高级管理人员', None),
                    '四、不予保险的其他存款': ('不予保险', None)}.items():
                r = (_find_irp_cross_rows(ws_irp, section, kw) if section
                     else _find_irp_row(ws_irp, kw))
                if r is None:
                    msgs.append(f"  {key}: IRP001 文件中未找到，跳过")
                    continue
                f_p = _get_cell_val(ws_irp, r, 2)
                f_i = _get_cell_val(ws_irp, r, 3)
                t = irp.get(key, {})
                ok = _diff_ok(t.get('principal', 0.0), f_p) and \
                    _diff_ok(t.get('interest', 0.0), f_i)
                file_ok = file_ok and ok
                msgs.append(
                    f"  {key}: 表内={t.get('principal', 0):,.2f}/{t.get('interest', 0):,.2f}  "
                    f"申报表={f_p:,.2f}/{f_i:,.2f}  [{'一致' if ok else '不一致'}]")
        except Exception as e:
            msgs.append(f"  ERROR: 无法读取 IRP001: {e}")

    # ---- (4) 表内调整项核对 ----
    # "用途=会计科目核对"的调整项在生成核对表时写入表内"调整项"行，
    # 并已计入个人/单位存款合计；此处核对表内数值与调整项配置是否一致。
    if items:
        msgs.append("")
        msgs.append("--- （4）表内调整项核对（与调整项配置比对）---")
        # 表内"调整项"行的本金/利息分别存放在 I 列与 L 列
        # （单位行为模板公式的扣减项，故期望值按 adjustment_row_values 口径换算）
        row_vals = adjustment_row_values(items)
        adjust_rows = data.get('adjust_rows', {})
        for r_name, target, field, idx in (
                ('个体工商户调整项（本金）（＋）', '个人存款', '本金', 0),
                ('个体工商户调整项（本金）（－）', '单位存款', '本金', 0),
                ('个体工商户调整项（应付利息）（＋）', '个人存款', '应付利息', 1),
                ('个体工商户调整项（应付利息）（－）', '单位存款', '应付利息', 1)):
            hit = next((v for k, v in adjust_rows.items() if r_name in k), None)
            if hit is None:
                continue
            val = to_float(hit[idx])
            expect = row_vals.get((target, field), 0.0)
            ok = abs(val - expect) <= 0.001
            msgs.append(f"  {r_name}: 表内={val:,.2f} 配置={expect:,.2f} "
                        f"[{'一致' if ok else '不一致'}]")

    # ---- 结论 ----
    msgs.append("")
    if dep_ok and acct_ok:
        msgs.append(">>> 保费基数核对表校验结果: 全部通过")
    else:
        msgs.append(">>> 保费基数核对表校验结果: 存在差异，请检查")
    return msgs

# ============================================================
# 4. 全面校验报告生成
# ============================================================

# --- 样式定义（延迟初始化） ---
RED_FILL = None
GRAY_FILL = None
HEADER_FILL = None
TITLE_FONT = None
HEADER_FONT = None
NORMAL_FONT_V = None
RED_FONT_V = None
THIN_BORDER = None
NUM_FMT_V = '#,##0.00'


def _init_styles():
    """
    延迟初始化 openpyxl 样式常量。

    因为 PatternFill/Font/Border 等对象在模块加载时创建可能有问题，
    所以在第一次使用时通过 _init_styles() 统一初始化。
    """
    global RED_FILL, GRAY_FILL, HEADER_FILL, TITLE_FONT, HEADER_FONT
    global NORMAL_FONT_V, RED_FONT_V, THIN_BORDER
    if RED_FILL is None:
        RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        TITLE_FONT = Font(name='微软雅黑', size=14, bold=True)
        HEADER_FONT = Font(name='微软雅黑', size=10, bold=True, color="FFFFFF")
        NORMAL_FONT_V = Font(name='微软雅黑', size=10)
        RED_FONT_V = Font(name='微软雅黑', size=10, color="CC0000", bold=True)
        THIN_BORDER = Border(
            left=Side('thin'), right=Side('thin'),
            top=Side('thin'), bottom=Side('thin'))


# ============================================================
# 数据读取辅助函数（各数据源的数据提取）
# ============================================================

def _read_irp001_map(filepath):
    """
    读取 IRP001 全部项目数据。

    IRP001 各行位置是固定的（0-indexed 行号已预定义）。
    返回 {label: {principal, interest, total}} 字典。

    参数:
        filepath (str): IRP001 文件路径

    返回:
        dict: {项目标签: {principal, interest, total}}
    """
    ws = _ExcelSheet(str(filepath))
    pairs = [
        (12, '一、境内存款'),
        (13, '（一）个人存款'),
        (14, '（二）单位存款'),
        (15, '住房公积金存款'),
        (16, '社保基金存款'),
        (17, '（三）财政性存款'),
        (18, '1.国库定期存款'),
        (19, '2.代理财政性存款'),
        (20, '（四）非存款类金融机构存款'),
        (21, '二、境外存款'),
        (22, '境外个人存款'),
        (23, '境外单位存款'),
        (24, '（三）境外同业存放'),
        (25, '高级管理人员'),
        (27, '被保险存款余额'),
        (28, '受保存款'),
        (29, '被保险存款客户数'),
        (30, '最高偿付限额以内客户数'),
    ]
    out = {}
    for r, label in pairs:
        out[label] = {
            'principal': _get_cell_val(ws, r, 2),   # C列：本金
            'interest': _get_cell_val(ws, r, 3),    # D列：应付利息
            'total': _get_cell_val(ws, r, 4),       # E列：本息合计
        }
    return out


def _read_irp002_map(filepath):
    """
    读取 IRP002 各项目的合计列数据。

    通过关键词动态查找行号，返回 {label: {principal, interest, total, count}}。

    参数:
        filepath (str): IRP002 文件路径

    返回:
        dict: {项目标签: {principal, interest, total, count}}
    """
    ws = _ExcelSheet(str(filepath))
    mapping = [
        ('1.境内存款',          '一、境内存款'),
        ('1.1 个人存款',        '（一）个人存款'),
        ('1.2 单位存款',        '（二）单位存款'),
        ('1.2.1 住房公积金',    '住房公积金存款'),
        ('1.2.2 社保基金',      '社保基金存款'),
        ('1.3 财政性存款',      '（三）财政性存款'),
        ('1.4 非存款类',        '（四）非存款类金融机构存款'),
        ('2.境外存款',          '二、境外存款'),
        ('2.1 个人存款',        '境外个人存款'),
        ('2.2 单位存款',        '境外单位存款'),
        ('2.3 境外同业存放',    '（三）境外同业存放'),
        ('3.高级管理人员',      '高级管理人员'),
        ('5.被保险存款',        '被保险存款余额'),
    ]
    out = {}
    for kw, label in mapping:
        r = _find_irp_row(ws, kw)
        if r is not None:
            out[label] = {
                'principal': _get_cell_val(ws, r, 3),  # D列：本金合计
                'interest': _get_cell_val(ws, r, 4),   # E列：利息合计
                'total': _get_cell_val(ws, r, 5),      # F列：本息合计
                'count': _get_cell_val(ws, r, 2),      # C列：户数合计
            }
    # 受保存款
    rs = _find_irp_row(ws, '5.1受保存款')
    if rs is not None:
        v = _get_cell_val(ws, rs, 2)
        out['受保存款'] = {'principal': v, 'interest': 0, 'total': v, 'count': 0}
    # 客户数
    for kw, label in [('5.被保险存款', '被保险存款客户数'),
                       ('5.被保险存款', '最高偿付限额以内客户数')]:
        r = _find_irp_row(ws, kw)
        if r is not None:
            out[label] = out.get(label, {})
            out[label]['count'] = _get_cell_val(ws, r, 2)
    # 最高偿付限额以内客户数 → col 6 (≤限额户数)
    r5 = _find_irp_row(ws, '5.被保险存款')
    if r5 is not None:
        cnt_lte = _get_cell_val(ws, r5, 6)
        out['最高偿付限额以内客户数'] = out.get('最高偿付限额以内客户数', {})
        out['最高偿付限额以内客户数']['count'] = cnt_lte
    return out



# ============================================================
# 校验报告（基于《校验表（新）》模板，文本式结论）
# ============================================================

# 校验报告模板（相对路径；支持 工作目录/模板 → 程序目录 → 内置 三级查找）
REPORT_TEMPLATE_REL = 'templates/校验表（新）.xlsx'

# 报告填写约定
NO_DATA_TEXT = '无对应数据'      # 没有数据的单元格
MATCH_TEXT = '核对一致'          # 核对一致

# 「匹配一致」的容差
TOL_YUAN = 0.001                 # 元级比较（会计科目 vs IRP001、IRP002 vs IRP001）
TOL_WAN = 0.001                  # 万元级比较（存款归属 vs 会计科目/IRP001）

def _diff_ok(v1, v2):
    """
    判断两个数值是否在容差内一致（元级，绝对容差 0.001 元）。

    参数:
        v1 (float): 值1
        v2 (float): 值2

    返回:
        bool: True 表示一致
    """
    return abs(to_float(v1) - to_float(v2)) <= TOL_YUAN


# 模板行项目 → 各数据源标签（按模板 A 列关键词匹配；§境外 表示境外章节内同名行）
# (A列关键词, 存款归属标签, 会计科目标签, IRP001标签, IRP002标签, 是否境外行)
_REPORT_ITEMS = [
    ('一、境内存款',               '一、境内存款', '一、境内存款', '一、境内存款', '一、境内存款', False),
    ('（一）个人存款',             '（一）个人存款', '（一）个人存款', '（一）个人存款', '（一）个人存款', False),
    ('（二）单位存款',             '（二）单位存款', '（二）单位存款', '（二）单位存款', '（二）单位存款', False),
    ('住房公积金',                 '住房公积金存款', None, '住房公积金存款', '住房公积金存款', False),
    ('社保基金存款',               '社保基金存款', None, '社保基金存款', '社保基金存款', False),
    ('（三）财政性存款',           '（三）财政性存款', '（三）财政性存款', '（三）财政性存款', '（三）财政性存款', False),
    ('1.国库定期存款',             '1.国库定期存款', '1.国库定期存款', '1.国库定期存款', None, False),
    ('2.代理财政性存款',           '2.代理财政性存款', '2.代理财政性存款', '2.代理财政性存款', None, False),
    ('（四）非存款类金融机构存款', '（四）非存款类金融机构存款', '（四）非存款类金融机构存款',
                                   '（四）非存款类金融机构存款', '（四）非存款类金融机构存款', False),
    ('二、境外存款',               '二、境外存款', None, '二、境外存款', '二、境外存款', False),
    ('（一）个人存款',             '境外（一）个人存款', None, '境外个人存款', '境外个人存款', True),
    ('（二）单位存款',             '境外（二）单位存款', None, '境外单位存款', '境外单位存款', True),
    ('（三）境外同业存放',         '境外（三）境外同业存放', '（三）境外同业存放',
                                   '（三）境外同业存放', '（三）境外同业存放', False),
    ('三、高级管理人员',           None, '高级管理人员', '高级管理人员', '高级管理人员', False),
    ('被保险存款',                 None, '被保险存款余额', '被保险存款余额', '被保险存款余额', False),
    ('受保存款',                   None, None, '受保存款', '受保存款', False),
    ('被保险存款客户数',           None, None, '被保险存款客户数', '被保险存款客户数', False),
    ('最高偿付限额以内客户数',     None, None, '最高偿付限额以内客户数', '最高偿付限额以内客户数', False),
]


def _row_item_for_label(label, in_overseas):
    """
    按模板 A 列标签找到对应项目定义。

    参数:
        label (str): 模板 A 列的项目名称
        in_overseas (bool): 当前是否处于"二、境外存款"章节内

    返回:
        tuple or None: 项目定义，未匹配返回 None
    """
    if not label:
        return None
    hits = [it for it in _REPORT_ITEMS if it[0] in label]
    # 长关键词优先（如"被保险存款客户数"优先于"被保险存款"）
    for item in sorted(hits, key=lambda it: -len(it[0])):
        if item[5] and not in_overseas:
            continue          # 境外同名行只在境外章节内匹配
        if not item[5] and item[0] in ('（一）个人存款', '（二）单位存款') and in_overseas:
            continue          # 境内同名行不在境外章节内匹配
        return item
    return None


def _cmp_text(left_name, left_val, right_name, right_val, tol, unit='元'):
    """
    按报告约定的三种结论生成单元格文本。

    返回:
        tuple: (文本, 是否不一致)
    """
    if left_val is None or right_val is None:
        return NO_DATA_TEXT, False
    if abs(left_val - right_val) <= tol:
        return MATCH_TEXT, False
    return (f"{left_name} {left_val:,.2f} {unit} / "
            f"{right_name} {right_val:,.2f} {unit}"), True


def _resolve_report_template(base_check_path=None, template_path=None):
    """
    定位校验报告模板（支持工作目录/模板 覆盖，与核对表模板同一套查找规则）。

    参数:
        base_check_path (str, optional): 已生成的保费基数核对表路径（取其目录为工作目录）
        template_path (str, optional): 直接指定模板路径

    返回:
        str: 模板文件路径
    """
    if template_path and os.path.exists(str(template_path)):
        return str(template_path)
    work_dir = os.path.dirname(os.path.abspath(str(base_check_path))) if base_check_path else None
    path, _ = resolve_template(REPORT_TEMPLATE_REL, work_dir)
    return path


def generate_validation_report(irp001_path, output_path,
                               irp002_path=None, base_check_path=None,
                               adjustments=None, template_path=None):
    """
    按《校验表（新）》模板生成校验报告。

    表格结构（模板固定）：
      A 列  = 项目名称
      B 列  = 存款归属核对（本金，万元） 与 会计科目核对（本金，万元）的比较
      C~E 列= 存款归属核对 与 IRP001（本金 / 利息 / 合计）的比较
              久悬户调整值已在生成核对表时直接加到存款归属表对应项，
              因此这里直接按万元比较，不再扣减 Q9/Q10
      F~H 列= 会计科目核对（被保险存款生成表）与 IRP001 的比较
              同样直接比较（会计科目调整值已在生成核对表时计入了表内）
      （模板若另有 I~K 列（IRP002 块），则填 IRP002 与 IRP001 的比较）

    填写约定：
      · 没有数据的单元格        → “无对应数据”
      · 核对一致                → “核对一致”
      · 核对不一致              → 把两个值都写出来（如 “存款归属 992,503.00 万元 / IRP001 992,532.00 万元”）

    参数:
        irp001_path:     IRP001 申报表路径（基准数据源，必填）
        output_path:     输出 .xlsx 路径（必填）
        irp002_path:     IRP002 申报表路径（可选）
        base_check_path: 保费基数核对表_已生成.xlsx 路径（可选，提供存款归属/会计科目数据）
        adjustments:     调整项列表（保留参数，用于日志/兼容；比较不再使用）
        template_path:   校验报告模板路径（可选，默认模板内置/程序目录）

    返回:
        list[str]: 运行日志消息列表
    """
    _init_styles()
    msgs = []

    # ---- 读取各数据源 ----
    msgs.append("========== 数据源读取 ==========")
    irp001 = _read_irp001_map(irp001_path)
    msgs.append(f"  IRP001: {len(irp001)} 个项目")
    irp002 = {}
    if irp002_path and os.path.exists(str(irp002_path)):
        irp002 = _read_irp002_map(irp002_path)
        msgs.append(f"  IRP002: {len(irp002)} 个项目")
    dep, acct = {}, {}
    if base_check_path and os.path.exists(str(base_check_path)):
        base = _read_base_check_map(base_check_path)
        dep, acct = base.get('deposit', {}), base.get('acct', {})
        dep_valid = sum(1 for v in dep.values() if v is not None)
        msgs.append(f"  保费基数核对表: 存款归属 {len(dep)} 项（有数值 {dep_valid} 项）, "
                    f"会计科目 {len(acct)} 项")
        if not dep_valid:
            msgs.append("  ⚠ 核对表中未读到存款归属数值（列布局不匹配或单元格为空），"
                        "B/C 列将填“无对应数据”。若该核对表是旧版本工具生成的，"
                        "请先重新执行「1. 生成保费基数核对表」")
        if not acct:
            msgs.append("  ⚠ 核对表中未读到会计科目数据（列布局不匹配），"
                        "B 列及 F~H 列将填“无对应数据”。若该核对表是旧版本工具"
                        "生成的，请先重新执行「1. 生成保费基数核对表」")
    else:
        msgs.append("  保费基数核对表: 未提供（存款归属/会计科目列将填“无对应数据”）")

    # 说明：久悬户调整值已在生成核对表时直接写入存款归属表/会计科目表，
    #       因此本报告直接比较表内数值与 IRP001，不再扣减 Q9/Q10。
    msgs.append("  久悬户调整值已计入核对表（A1411、2411合计表/存款归属表），本次直接比较")

    # ---- 打开模板 ----
    tpl_path = _resolve_report_template(base_check_path, template_path)
    if not os.path.exists(tpl_path):
        return [f"ERROR: 未找到校验报告模板: {tpl_path}"]
    wb = openpyxl.load_workbook(tpl_path)
    ws = wb.active
    msgs.append(f"  模板: {tpl_path}")
    msgs.append("")
    msgs.append("========== 逐项校验 ==========")

    # ---- 逐行填写 ----
    overseas = False
    mismatch = 0
    filled = 0
    # 单元格级统计：有效比较数 / 一致 / 不一致 / 无对应数据
    cell_stat = {'compared': 0, 'matched': 0, 'mismatch': 0, 'no_data': 0}

    def tally(text):
        """按单元格结论累计统计（无对应数据不计入"比较"）。"""
        if text == NO_DATA_TEXT:
            cell_stat['no_data'] += 1
        elif text == MATCH_TEXT:
            cell_stat['compared'] += 1
            cell_stat['matched'] += 1
        else:
            cell_stat['compared'] += 1
            cell_stat['mismatch'] += 1
    # 是否含 IRP002 区块：以表头行 3 的 I/J/K 标题为准（避免模板中零散误输入）
    has_irp002_cols = all(str(ws.cell(3, c).value or '').strip()
                          for c in (9, 10, 11))

    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or '').strip()
        if not label:
            continue
        if '说明' in label and '：' in label:
            continue
        if label.startswith('二、'):
            overseas = True
        elif label.startswith('三、') or label.startswith('一、'):
            overseas = False

        item = _row_item_for_label(label, overseas)
        if item is None:
            continue
        _, dep_key, acct_key, irp_key, irp002_key, _is_ovs = item
        d_val = dep.get(dep_key) if dep_key else None
        a = acct.get(acct_key) if acct_key else None
        i1 = irp001.get(irp_key) if irp_key else None
        i2 = irp002.get(irp002_key) if irp002_key else None

        # B 列：存款归属（万元，已含存款归属调整）vs 会计科目（万元，已含会计科目调整）
        if d_val is None:
            b_text, b_err = NO_DATA_TEXT, False
        else:
            b_right = None
            if label == '一、境内存款':
                b_right = 0.0
                any_acct = False
                for k in ('（一）个人存款', '（二）单位存款',
                          '（三）财政性存款', '（四）非存款类金融机构存款'):
                    ai = acct.get(k)
                    if not ai:
                        continue
                    any_acct = True
                    b_right += excel_round(ai.get('principal', 0) / 10000.0, 0)
                if not any_acct:
                    b_right = None
            else:
                a_principal = (a or {}).get('principal') if a else None
                if a_principal is not None:
                    b_right = excel_round(a_principal / 10000.0, 0)
            b_left = excel_round(d_val, 0)
            b_text, b_err = _cmp_text('A1411、2411合计表', b_left, '存款归属表', b_right,
                                      TOL_WAN, '万元')
        _write_report_cell(ws, r, 2, b_text, b_err)
        tally(b_text)

        # C~E 列：存款归属（万元）vs IRP001 本金（直接比较，精确至万元）
        if d_val is not None and i1:
            dep_wan = excel_round(d_val, 0)
            if _is_ovs:
                irp_wan = excel_round((i1.get('principal', 0)) / 10000.0, 0)
            elif label == '一、境内存款':
                # 境内合计：按各分项四舍五入至万元后求和，与核对表境内合计口径一致
                irp_wan = 0.0
                for k in ('（一）个人存款', '（二）单位存款',
                          '（三）财政性存款', '（四）非存款类金融机构存款'):
                    irp_wan += excel_round(
                        irp001.get(k, {}).get('principal', 0) / 10000.0, 0)
            else:
                irp_wan = excel_round(i1.get('principal', 0) / 10000.0, 0)
            if abs(dep_wan - irp_wan) <= TOL_WAN:
                c_text, c_err = MATCH_TEXT, False
            else:
                c_text = (f"A1411、2411合计表 {dep_wan:,.0f} 万元 / "
                          f"IRP001 {irp_wan:,.0f} 万元")
                c_err = True
        else:
            c_text, c_err = NO_DATA_TEXT, False
        _write_report_cell(ws, r, 3, c_text, c_err)
        _write_report_cell(ws, r, 4, NO_DATA_TEXT, False)   # 存款归属口径无应付利息
        _write_report_cell(ws, r, 5, NO_DATA_TEXT, False)   # 存款归属口径无合计
        tally(c_text)
        tally(NO_DATA_TEXT)
        tally(NO_DATA_TEXT)

        # F~H 列：会计科目（元）vs IRP001（本金/利息/合计）
        for col, field, left_name in ((6, 'principal', '会计科目'),
                                      (7, 'interest', '会计科目'),
                                      (8, 'total', '会计科目')):
            l_val = (a or {}).get(field) if a else None
            r_val = (i1 or {}).get(field) if i1 else None
            text, err = _cmp_text(left_name, l_val, 'IRP001', r_val, TOL_YUAN, '元')
            _write_report_cell(ws, r, col, text, err)
            tally(text)

        # I~K 列（模板若存在 IRP002 块）：IRP002 vs IRP001
        if has_irp002_cols:
            is_count_row = '客户数' in label
            for col, field in ((9, 'principal'), (10, 'interest'), (11, 'total')):
                if is_count_row:
                    # 户数行：IRP002 户数合计 vs IRP001 户数（本金列存放的是户数）
                    if col == 9:
                        text, err = _cmp_text('IRP002户数', (i2 or {}).get('count'),
                                              'IRP001户数', (i1 or {}).get('principal'),
                                              TOL_YUAN, '户')
                    else:
                        text, err = NO_DATA_TEXT, False
                    _write_report_cell(ws, r, col, text, err)
                    tally(text)
                    continue
                l_val = (i2 or {}).get(field) if i2 else None
                r_val = (i1 or {}).get(field) if i1 else None
                if field == 'total' and i1 and r_val == 0:
                    r_val = i1.get('principal', 0)      # 受保存款等只有本金的情况
                text, err = _cmp_text('IRP002', l_val, 'IRP001', r_val, TOL_YUAN, '元')
                _write_report_cell(ws, r, col, text, err)
                tally(text)

        if b_err or c_err:
            mismatch += 1
        filled += 1
        msgs.append(f"  {label}: B列[{b_text}] C列[{c_text}]")
        if a and i1:
            for field, fname in (('principal', '本金'), ('interest', '利息'), ('total', '合计')):
                l_val, r_val = (a or {}).get(field), (i1 or {}).get(field)
                if l_val is not None and r_val is not None and not _diff_ok(l_val, r_val):
                    mismatch += 0
                    msgs.append(f"    F/G/H {fname}: 会计科目 {l_val:,.2f} / IRP001 {r_val:,.2f}")

    # ---- 说明行 ----
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or '')
        if v.startswith('说明'):
            cell = ws.cell(r, 1)
            cell.value = (
                "说明：B列＝A1411、2411合计表与存款归属表（本金，万元）比较；"
                "C~E列＝A1411、2411合计表与IRP001（本金/利息/合计）比较；"
                "F~H列＝存款归属表与IRP001（本金/利息/合计）比较"
                + ("；I~K列＝IRP002与IRP001比较" if has_irp002_cols else "") +
                "。填写约定：无对应数据＝该口径没有这项数据；核对一致＝在容差内相等"
                f"（元级 {TOL_YUAN:g} 元）；核对不一致＝列出双方数值（红色底）。")
            cell.font = Font(name='微软雅黑', size=9, italic=True, color="666666")
            cell.alignment = Alignment(horizontal='left', vertical='center')
            break

    ws.column_dimensions['A'].width = 26
    for col in 'BCDEFGHIJK':
        ws.column_dimensions[col].width = 30

    wb.save(str(output_path))
    wb.close()

    msgs.append("")
    msgs.append("========== 校验结果汇总 ==========")
    msgs.append(f"校验报告已生成: {output_path}")
    msgs.append(f"共填写 {filled} 个项目行，其中 {mismatch} 行存在不一致（已标红）")
    msgs.append(
        f"单元格统计：有效比较 {cell_stat['compared']} 个"
        f"（一致 {cell_stat['matched']} / 不一致 {cell_stat['mismatch']}），"
        f"无对应数据 {cell_stat['no_data']} 个")
    if cell_stat['compared'] == 0:
        msgs.append(">>> 未进行任何有效比较：报告中的“无对应数据”不是校验结论，"
                    "请检查保费基数核对表是否已生成/路径是否正确、"
                    "以及表中存款归属与会计科目数据是否为空")
    elif cell_stat['mismatch'] == 0:
        msgs.append(">>> 所有校验项一致 ✓")
    else:
        msgs.append(">>> 请检查标红单元格对应的两个数值")
    return msgs


def _write_report_cell(ws, row, col, text, is_error):
    """
    写入报告单元格：不一致标红，其余按模板原样式。

    模板存在合并单元格（如 F19:H19），合并区域内只有左上角可写，
    其余位置跳过（内容已由左上角单元格承载）。

    返回:
        str or None: 实际写入的文本；跳过时返回 None
    """
    r, c = _resolve_merged(ws, row, col)
    if (r, c) != (row, col):
        return None                     # 合并区域内的从属单元格
    cell = ws.cell(r, c)
    cell.value = text
    cell.font = RED_FONT_V if is_error else NORMAL_FONT_V
    cell.fill = RED_FILL if is_error else PatternFill(fill_type=None)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = THIN_BORDER
    return text
