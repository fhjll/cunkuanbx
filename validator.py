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
       - 存款归属核对（F列，万元）与 IRP001 本金（扣久悬户调整后精确至万元）比较
       - 会计科目核对（I/L/M列，元）与 IRP001 本金/利息/合计比较

  3. generate_validation_report()     — 生成全面校验报告（Excel）
       - 以 IRP001 为基准，逐项输出存款归属核对（B列）、
         会计科目核对（C/D/E列）、IRP002（F/G/H列）的差额，不一致标红

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
    Q7/Q8 保存个人侧净调整额（本金/利息）供外部核对
  - 用途=存款归属核对：按方式汇总到 Q9（个人）/Q10（单位），
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
from common import to_float, excel_round, ExcelSheet as _ExcelSheet
from base_checker import (ADJ_SCOPE_ACCT, ADJ_SCOPE_DEP,
                          adjustment_row_values, describe_adjustments,
                          normalize_adjustments, summarize_adjustments)


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

    表格结构（模板固定列）：
      C~F 列：存款归属核对（子项代码 / 子项名称 / 符号 / 金额万元）
      G~M 列：会计科目核对（本金代码、科目名称、金额；
                           利息代码、科目名称、金额；合计）
      N~P 列：被保险存款申报表（IRP001）本金 / 应付利息 / 合计
      Q 列  ：差额说明（Q7/Q8 个人侧净调整额；Q9/Q10 存款归属核对调整额）

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
        }；读取失败返回 {}
    """
    try:
        wb = openpyxl.load_workbook(str(filepath), data_only=True)
    except Exception:
        return {}
    ws = wb[BASE_SHEET_NAME] if BASE_SHEET_NAME in wb.sheetnames else wb.active

    out = {'deposit': {}, 'acct': {}, 'irp': {}, 'adjust': {}, 'rows': {},
           'adjust_rows': {}}

    # 调整值（Q 列前四格：c1/c2=个人侧净调整额，c3/c4=存款归属核对调整额）
    for key, row in (('c1', 7), ('c2', 8), ('c3', 9), ('c4', 10)):
        out['adjust'][key] = to_float(ws.cell(row, 17).value)

    section = '境内'
    acct_p = acct_i = 0.0    # 会计科目 境内存款累计（个人+单位+财政性+非存款类）
    for r in range(1, ws.max_row + 1):
        b_name = str(ws.cell(r, 2).value or '').strip()
        d_name = str(ws.cell(r, 4).value or '').strip()
        h_name = str(ws.cell(r, 8).value or '').strip()
        k_name = str(ws.cell(r, 11).value or '').strip()

        if '境外存款' in b_name:
            section = '境外'
        elif '境内存款' in b_name:
            section = '境内'

        # ---- 存款归属核对（F 列，万元）----
        label = _deposit_label(section, b_name, d_name)
        if label:
            out['deposit'][label] = to_float(ws.cell(r, 6).value)
            out['rows'].setdefault(label, r)

        # ---- 表内"调整项"行（本金在 I 列、利息在 L 列）----
        if '调整项' in h_name or '调整项' in k_name:
            vals = (to_float(ws.cell(r, 9).value), to_float(ws.cell(r, 12).value))
            for key in (h_name, k_name):
                if key:
                    out['adjust_rows'][key] = vals

        # ---- 会计科目核对（I / L / M 列，元）----
        if h_name or k_name:
            a_label = _acct_label(h_name)
            if a_label:
                p_val = to_float(ws.cell(r, 9).value)
                i_val = to_float(ws.cell(r, 12).value)
                out['acct'][a_label] = {
                    'principal': p_val,
                    'interest': i_val,
                    'total': to_float(ws.cell(r, 13).value) or (p_val + i_val),
                }
                out['rows'].setdefault(a_label, r)
                if a_label in ('（一）个人存款', '（二）单位存款',
                               '（三）财政性存款', '（四）非存款类金融机构存款'):
                    acct_p += p_val
                    acct_i += i_val
            # 财政性存款明细行（国库定期/代理财政）标签在 D 列
            sub_label = _acct_sub_label(d_name)
            if sub_label:
                s_p = to_float(ws.cell(r, 9).value)
                s_i = to_float(ws.cell(r, 12).value)
                out['acct'][sub_label] = {
                    'principal': s_p,
                    'interest': s_i,
                    'total': to_float(ws.cell(r, 13).value) or (s_p + s_i),
                }
                out['rows'].setdefault(sub_label, r)

            # ---- IRP001（N / O / P 列，元）----
            if _acct_label(h_name):
                out['irp'][_acct_label(h_name)] = {
                    'principal': to_float(ws.cell(r, 14).value),
                    'interest': to_float(ws.cell(r, 15).value),
                    'total': to_float(ws.cell(r, 16).value),
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
    mapping = {
        '（一）个人存款': '（一）个人存款',
        '（二）单位存款': '（二）单位存款',
        '其中：住房公积金存款': '住房公积金存款',
        '社保基金存款': '社保基金存款',
        '（三）财政性存款': '（三）财政性存款',
        '（四）非存款类金融机构存款': '（四）非存款类金融机构存款',
        '境内存款合计': '一、境内存款',
        '1. 国库定期存款': '1.国库定期存款',
        '2. 代理财政性存款': '2.代理财政性存款',
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
    """财政性存款明细行（D 列标签：1. 国库定期存款 / 2. 代理财政性存款）"""
    mapping = {'1. 国库定期存款': '1.国库定期存款',
               '2. 代理财政性存款': '2.代理财政性存款'}
    return mapping.get(d_name)


def validate_base_check_vs_irp001(base_check_path, irp001_path=None,
                                  adjustments=None):
    """
    校验《保费基数核对表》(已生成) 与 IRP001 的数据一致性。

    校验逻辑（沿用原存款归属表/科目余额表两项校验的口径，合并为一次）：
      1. 先对子项数据取数、再按模板公式计算合并（已在生成表内完成）
      2. 存款归属核对（F 列，万元）与 IRP001 本金比较：
           个人存款：ROUND((IRP001个人本金 - 个人调整额) / 10000)
           单位存款：ROUND((IRP001单位本金 - 单位调整额) / 10000)
           财政性 / 非存款类：ROUND(IRP001本金 / 10000)
         （四舍五入至万元后应完全相等，差额≠0 即不一致）
         调整额 = 调整项中"用途=存款归属核对"的净额（方式=扣减记正数）；
         未提供调整项时取表内 Q9/Q10。
      3. 会计科目核对（I/L/M 列，元）与 IRP001 本金/利息/合计比较：
           绝对容差 0.001 元（"用途=会计科目核对"的调整项已在表内
           "调整项"行体现，此处不再重复调整）
      4. 当提供 IRP001 文件时，同时核对表中 N/O 列取数是否与申报表一致
      5. 当提供调整项时，核对表内"调整项"行与 Q9/Q10 是否与配置一致
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

    # 存款归属核对调整额：优先用调整项配置，未提供时取表内 Q9/Q10
    items = normalize_adjustments(adjustments) if adjustments else []
    _, dep_adj = summarize_adjustments(items) if items else ({}, {})
    c3_val = dep_adj.get('个人存款', 0.0)
    c4_val = dep_adj.get('单位存款', 0.0)
    if items:
        src3 = src4 = '调整项配置'
        q3, q4 = adjust.get('c3', 0.0), adjust.get('c4', 0.0)
        if abs(c3_val - q3) > 0.001:
            msgs.append(f"  ⚠ 表中 Q9({q3:,.2f}) 与调整项配置({c3_val:,.2f}) 不一致，"
                        f"请重新生成核对表")
        if abs(c4_val - q4) > 0.001:
            msgs.append(f"  ⚠ 表中 Q10({q4:,.2f}) 与调整项配置({c4_val:,.2f}) 不一致，"
                        f"请重新生成核对表")
    else:
        c3_val, c4_val = adjust.get('c3', 0.0), adjust.get('c4', 0.0)
        src3 = src4 = '表内Q9/Q10'

    msgs.append("========== 保费基数核对表 vs IRP001 校验 ==========")
    msgs.append(f"存款归属核对调整额: 个人={c3_val:,.2f} 元（{src3}）、"
                f"单位={c4_val:,.2f} 元（{src4}）")

    # ---- (1) 存款归属核对 vs IRP001 ----
    msgs.append("")
    msgs.append("--- （1）存款归属核对 vs IRP001（模板校验页 B/C 列口径，单位：万元）---")
    dep_items = [
        ('（一）个人存款', '（一）个人存款', c3_val),
        ('（二）单位存款', '（二）单位存款', c4_val),
        ('（三）财政性存款', '（三）财政性存款', 0.0),
        ('（四）非存款类金融机构存款', '（四）非存款类金融机构存款', 0.0),
    ]
    dep_ok = True
    for label, key, adj in dep_items:
        dep_val = dep.get(key)
        irp_val = irp.get(key, {}).get('principal', 0.0)
        if dep_val is None:
            msgs.append(f"  {label}: 表中未找到，跳过")
            continue
        expect = excel_round((irp_val - adj) / 10000.0, 0)
        diff = dep_val - expect
        ok = abs(diff) <= 0.001
        dep_ok = dep_ok and ok
        adj_note = f"（IRP001本金 - {adj:,.2f}）" if adj else ""
        msgs.append(
            f"  {label}: 存款归属={dep_val:,.2f}万元  "
            f"IRP001调整后={expect:,.0f}万元{adj_note}  "
            f"差额={diff:,.2f}万元  [{'一致' if ok else '不一致'}]")

    # ---- (2) 会计科目核对 vs IRP001 ----
    msgs.append("")
    msgs.append("--- （2）会计科目核对 vs IRP001（元，容差 0.001 元）---")
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
        for target, qkey, qcell in (('个人存款', 'c3', 'Q9'),
                                    ('单位存款', 'c4', 'Q10')):
            detail = describe_adjustments(items, scope=ADJ_SCOPE_DEP, target=target)
            if detail:
                msgs.append(f"  {target}·存款归属核对调整额 {adjust.get(qkey, 0):,.2f}"
                            f"（{qcell}）：" + detail.replace('\n', '；'))

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
# 校验表行定义
# ============================================================

# 每条记录格式：
# (excel_row_1based, item_label, irp001_key, deposit_key, acct_key, irp002_key, skip_cols)
# skip_cols: {'B', 'C', 'D', 'E', 'F', 'G', 'H'} — 标记 '/' 的列（不校验）
_VALIDATION_ITEMS = [
    # row  项目名称                        IRP001关键词          存款归属表关键词       科目核对表关键词         IRP002关键词          跳过列
    (4,  '一、境内存款',                  '一、境内存款',         '一、境内存款',       '一、境内存款',           '一、境内存款',            set()),
    (5,  '（一）个人存款',                '（一）个人存款',       '（一）个人存款',     '（一）个人存款',          '（一）个人存款',          set()),
    (6,  '（二）单位存款',                '（二）单位存款',       '（二）单位存款',     '（二）单位存款',          '（二）单位存款',          set()),
    (7,  '住房公积金存款',                '住房公积金存款',       '住房公积金存款',     '住房公积金存款',          '住房公积金存款',          set()),
    (8,  '社保基金存款',                  '社保基金存款',         '社保基金存款',       '社保基金存款',            '社保基金存款',            set()),
    (9,  '（三）财政性存款',              '（三）财政性存款',     '（三）财政性存款',    '（三）财政性存款',        '（三）财政性存款',        set()),
    (10, '1.国库定期存款',                '1.国库定期存款',       '1.国库定期存款',     '1.国库定期存款',          None,                      set()),
    (11, '2.代理财政性存款',              '2.代理财政性存款',     '2.代理财政性存款',    '2.代理财政性存款',        None,                      set()),
    (12, '（四）非存款类金融机构存款',    '（四）非存款类金融机构存款', '（四）非存款类金融机构存款', '（四）非存款类金融机构存款', '（四）非存款类金融机构存款', set()),
    (13, '二、境外存款',                  '二、境外存款',         None,                 '二、境外存款',            '二、境外存款',            set()),
    (14, '境外个人存款',                  '境外个人存款',         '境外（一）个人存款', None,                      '境外个人存款',            {'C','D','E'}),
    (15, '境外单位存款',                  '境外单位存款',         '境外（二）单位存款', None,                      '境外单位存款',            {'C','D','E'}),
    (16, '（三）境外同业存放',            '（三）境外同业存放',    '境外（三）境外同业存放', '（三）境外同业存放',    '（三）境外同业存放',      set()),
    (17, '高级管理人员',                  '高级管理人员',         None,                 '高级管理人员',            '高级管理人员',            {'B'}),
    (18, '被保险存款余额',                '被保险存款余额',       None,                 '被保险存款余额',           '被保险存款余额',          {'B'}),
    (19, '受保存款',                      '受保存款',             None,                 None,                      '受保存款',                {'B','C'}),
    (20, '被保险存款客户数',              '被保险存款客户数',     None,                 None,                      '被保险存款客户数',        {'B','C'}),
    (21, '最高偿付限额以内客户数',         '最高偿付限额以内客户数', None,                None,                      '最高偿付限额以内客户数',  {'B','C'}),
]


# ============================================================
# 辅助函数：单元格写入和比较
# ============================================================

def _col_letter(idx):
    """0-indexed 列索引 → Excel 列字母（A-H）"""
    return chr(ord('A') + idx)


def _skip_check(skip_set, col_letter_str):
    """检查某列是否被标记为跳过（不校验）"""
    return col_letter_str in skip_set


def _diff_ok(v1, v2):
    """
    判断两个值是否在容差范围内一致。

    容差规则：绝对容差 0.001 元（严格一致，仅消除浮点精度尾差）。

    参数:
        v1 (float): 值1
        v2 (float): 值2

    返回:
        bool: True 表示在容差范围内一致
    """
    return abs(v1 - v2) <= 0.001


def _write_cell(ws, row, col, value, is_error=False, number_fmt=None,
                font=None, fill=None, alignment=None, border=None):
    """
    向工作表写入单元格，is_error 时自动应用红色样式。

    is_error=True 时：
      - 字体：红色加粗（#CC0000）
      - 填充：红色底色（#FFC7CE）

    参数:
        ws: openpyxl Worksheet
        row (int): 行号（1-indexed）
        col (int): 列号（1-indexed）
        value: 单元格值
        is_error (bool): 是否为差异项（标红）
        number_fmt (str, optional): 数字格式
        font (Font, optional): 字体（is_error 时忽略）
        fill (PatternFill, optional): 填充（is_error 时忽略）
        alignment (Alignment, optional): 对齐方式
        border (Border, optional): 边框
    """
    _init_styles()
    cell = ws.cell(row, col)
    cell.value = value
    cell.number_format = number_fmt or NUM_FMT_V
    cell.font = RED_FONT_V if is_error else (font or NORMAL_FONT_V)
    cell.fill = RED_FILL if is_error else (fill or PatternFill(fill_type=None))
    cell.alignment = alignment or Alignment(horizontal='right', vertical='center')
    cell.border = border or THIN_BORDER


# ============================================================
# 核心函数：生成全面校验报告
# ============================================================

def generate_validation_report(irp001_path, output_path,
                                irp002_path=None, base_check_path=None,
                                adjustments=None, template_path=None):
    """
    基于校验表模板，生成全面校验报告（Excel 格式）。

    校验范围涵盖最多 3 个数据源的交叉比对：
      - IRP001（基准数据源，必填）
      - 保费基数核对表_已生成（可选，内含存款归属核对 + 会计科目核对）
      - IRP002（可选）

    报告格式：
      - 标题行：存保申报数据校验表
      - 分组表头：数据来源分组（存款归属核对、会计科目核对、IRP002）
      - 数据行：每行一个校验项，输出 0（一致）或差额（不一致）
      - 差异行自动标红（红色底色）

    参数:
        irp001_path:     IRP001 申报表路径（必填，作为基准数据源）
        output_path:     输出 .xlsx 路径（必填）
        irp002_path:     IRP002 申报表路径（可选）
        base_check_path: 保费基数核对表_已生成.xlsx 路径（可选）
        adjustments:     调整项列表（可选，与生成核对表时一致）；
                         "用途=会计科目核对"的调整项已含在核对表内，
                         "用途=存款归属核对"的调整额用于从 IRP001 本金中扣减；
                         不传时取核对表内 Q9/Q10
        template_path:   （保留参数，暂未使用）

    返回:
        list[str]: 运行日志消息列表
    """
    _init_styles()
    msgs = []

    # 调整项：会计科目侧已含在核对表内；存款归属侧用于比较时调整 IRP001 本金
    items = normalize_adjustments(adjustments) if adjustments else []
    acct_adj, dep_adj = summarize_adjustments(items) if items else ({}, {})

    # ---- 读取各数据源 ----
    msgs.append("========== 校验参数 ==========")
    if items:
        for scope in (ADJ_SCOPE_ACCT, ADJ_SCOPE_DEP):
            detail = describe_adjustments(items, scope=scope)
            if detail:
                msgs.append(f"  [{scope}]（{len(detail.splitlines())}项）")
                msgs.extend('    ' + ln for ln in detail.splitlines())
    else:
        msgs.append("  未提供调整项，按核对表内数值比较")
    msgs.append("")
    msgs.append("========== 数据源读取 ==========")

    irp001 = _read_irp001_map(irp001_path)
    msgs.append(f"  IRP001: 读取 {len(irp001)} 个项目（基准数据源）")

    irp002 = {}
    if irp002_path and os.path.exists(str(irp002_path)):
        irp002 = _read_irp002_map(irp002_path)
        msgs.append(f"  IRP002: 读取 {len(irp002)} 个项目")
    else:
        msgs.append("  IRP002: 未提供，跳过 IRP002 相关列（F/G/H）校验")

    deposit = {}
    deposit, acct = {}, {}
    if base_check_path and os.path.exists(str(base_check_path)):
        base = _read_base_check_map(base_check_path)
        deposit = base.get('deposit', {})
        acct = base.get('acct', {})
        if not base:
            msgs.append("  保费基数核对表: 读取失败，跳过 B/C/D/E 列校验")
        else:
            msgs.append(f"  保费基数核对表: 存款归属 {len(deposit)} 项, "
                        f"会计科目 {len(acct)} 项")
            if not items:   # 未提供调整项时取表内 Q9/Q10
                dep_adj = {'个人存款': base.get('adjust', {}).get('c3', 0.0),
                           '单位存款': base.get('adjust', {}).get('c4', 0.0)}
    else:
        msgs.append("  保费基数核对表: 未提供，跳过 B/C/D/E 列校验")

    adj_c3 = dep_adj.get('个人存款', 0.0)   # 存款归属核对：个人调整额（元）
    adj_c4 = dep_adj.get('单位存款', 0.0)   # 存款归属核对：单位调整额（元）
    msgs.append("")
    msgs.append("========== 逐项校验明细 ==========")

    # ---- 构建输出工作簿 ----
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "校验表"

    # ---- 第1行：标题 ----
    ws.merge_cells('A1:H1')
    _write_cell(ws, 1, 1, "存保申报数据校验表", number_fmt=None,
                font=TITLE_FONT, fill=PatternFill(fill_type=None),
                alignment=Alignment(horizontal='center', vertical='center'))
    ws.row_dimensions[1].height = 32

    # ---- 第2-3行：分组表头 ----
    # A2:A3 合并 = "项目名称"
    ws.merge_cells('A2:A3')
    _write_cell(ws, 2, 1, "项目名称", number_fmt=None,
                font=Font(name='微软雅黑', size=10, bold=True),
                alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    # B2 = "存款归属表与IRP001核对（0=一致/差额）"
    ws.merge_cells('B2:B2')
    _write_cell(ws, 2, 2, "存款归属表与\nIRP001核对\n（0=一致/差额）", number_fmt=None,
                font=Font(name='微软雅黑', size=8, bold=True),
                alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    # C2:E2 = "科目余额核对表与IRP001核对"
    ws.merge_cells('C2:E2')
    _write_cell(ws, 2, 3, "科目余额核对表与IRP001核对\n（0=一致/差额）", number_fmt=None,
                font=Font(name='微软雅黑', size=9, bold=True),
                alignment=Alignment(horizontal='center', vertical='center'))
    # F2:H2 = "IRP002与IRP001核对"
    ws.merge_cells('F2:H2')
    _write_cell(ws, 2, 6, "IRP002与IRP001核对", number_fmt=None,
                font=Font(name='微软雅黑', size=9, bold=True),
                alignment=Alignment(horizontal='center', vertical='center'))

    # 第3行子表头
    sub_headers = ['', '差额（元）', '差额（元）', '差额（元）', '差额（元）',
                   '本金（元）', '利息（元）', '合计（元）']
    for ci, h in enumerate(sub_headers):
        if ci == 0:
            continue
        _write_cell(ws, 3, ci + 1, h, number_fmt=None,
                    font=Font(name='微软雅黑', size=9, bold=True),
                    alignment=Alignment(horizontal='center', vertical='center'))
    ws.row_dimensions[2].height = 36

    # ---- 逐项填充数据 ----
    error_count = 0
    error_items = []   # 差异项明细，用于末尾汇总
    for item in _VALIDATION_ITEMS:
        excel_row, label, irp001_key, dep_key, acct_key, irp002_key, skip_cols = item
        irp001_data = irp001.get(irp001_key, {})

        msgs.append(f"--- 第{excel_row}行「{label}」---")
        if not irp001_data:
            msgs.append(f"  WARNING: IRP001 中未找到「{irp001_key}」对应数据，基准值按 0 处理")
        else:
            msgs.append(
                f"  IRP001基准: 本金={irp001_data.get('principal', 0):,.2f}  "
                f"利息={irp001_data.get('interest', 0):,.2f}  "
                f"本息合计={irp001_data.get('total', 0):,.2f}")

        # A列：项目名称
        _write_cell(ws, excel_row, 1, label, number_fmt=None,
                    font=Font(name='微软雅黑', size=10, bold=True),
                    alignment=Alignment(horizontal='left', vertical='center'))

        # ---- B列：存款归属核对 vs IRP001 ----
        if not _skip_check(skip_cols, 'B'):
            dep_val_wan = deposit.get(dep_key) if dep_key else None
            if dep_val_wan is not None:
                irp001_principal = irp001_data.get('principal', 0)
                # 应用久悬户余额调整（模板校验页口径：先扣久悬户，再精确至万元）
                adj_note = ""
                if label == '（一）个人存款':
                    irp001_principal -= adj_c3
                    adj_note = f"（已扣除个人久悬户 c3={adj_c3:,.2f}）"
                elif label == '（二）单位存款':
                    irp001_principal -= adj_c4
                    adj_note = f"（已扣除单位久悬户 c4={adj_c4:,.2f}）"
                elif label == '一、境内存款':
                    irp001_principal -= (adj_c3 + adj_c4)
                    adj_note = f"（已扣除久悬户 c3+c4={adj_c3 + adj_c4:,.2f}）"
                irp_in_wan = excel_round(irp001_principal / 10000.0, 0)
                diff_wan = dep_val_wan - irp_in_wan
                is_err = abs(diff_wan) * 10000 > 0.001  # 绝对容差 0.001 元，超出即标红
                if is_err:
                    error_count += 1
                    error_items.append(
                        f"第{excel_row}行「{label}」B列: 存款归属核对={dep_val_wan:,.2f}万元 vs "
                        f"IRP001={irp_in_wan:,.0f}万元，差额={diff_wan:,.2f}万元")
                result = 0 if not is_err else diff_wan * 10000
                _write_cell(ws, excel_row, 2, result, is_error=is_err)
                msgs.append(
                    f"  B列 存款归属核对 vs IRP001: 核对表={dep_val_wan:,.2f}万元  "
                    f"IRP001调整后={irp_in_wan:,.0f}万元{adj_note}  "
                    f"差额={diff_wan:,.2f}万元  [{'ERROR' if is_err else 'OK'}]")
            else:
                _write_cell(ws, excel_row, 2, None, number_fmt=None)
                if dep_key and deposit:
                    msgs.append(f"  B列: 保费基数核对表中未找到「{dep_key}」，留空")
        else:
            _write_cell(ws, excel_row, 2, '/', number_fmt=None,
                        fill=GRAY_FILL, font=NORMAL_FONT_V,
                        alignment=Alignment(horizontal='center', vertical='center'))

        # ---- C/D/E列：会计科目核对 vs IRP001 ----
        # "用途=会计科目核对"的调整项在生成核对表时已通过"调整项"行计入
        # 个人/单位存款，此处不再重复调整。
        acct_data = acct.get(acct_key) if acct_key else None
        if acct_key and acct_data is None and acct:
            msgs.append(f"  C/D/E列: 保费基数核对表中未找到「{acct_key}」，留空")
        for col_offset, field, col_letter in [
            (3, 'principal', 'C'), (4, 'interest', 'D'), (5, 'total', 'E')
        ]:
            if _skip_check(skip_cols, col_letter):
                _write_cell(ws, excel_row, col_offset, '/', number_fmt=None,
                            fill=GRAY_FILL, font=NORMAL_FONT_V,
                            alignment=Alignment(horizontal='center', vertical='center'))
                continue

            acct_val = acct_data.get(field) if acct_data else None
            irp001_val = irp001_data.get(field, 0) if irp001_data else 0
            if acct_val is not None:
                adjusted_val = acct_val
                is_err = not _diff_ok(adjusted_val, irp001_val)
                tolerance = 0.001
                if is_err:
                    error_count += 1
                    error_items.append(
                        f"第{excel_row}行「{label}」{col_letter}列: 核对表={adjusted_val:,.2f} vs "
                        f"IRP001={irp001_val:,.2f}，差额={adjusted_val - irp001_val:,.2f}元")
                result = 0 if not is_err else (adjusted_val - irp001_val)
                _write_cell(ws, excel_row, col_offset, result, is_error=is_err)
                field_name = {'principal': '本金', 'interest': '利息', 'total': '本息合计'}[field]
                msgs.append(
                    f"  {col_letter}列 核对表{field_name} vs IRP001: 核对表={acct_val:,.2f}  "
                    f"IRP001={irp001_val:,.2f}  差额={adjusted_val - irp001_val:,.2f}  "
                    f"容差=±{tolerance:g}元  [{'ERROR' if is_err else 'OK'}]")
            else:
                _write_cell(ws, excel_row, col_offset, None, number_fmt=None)

        # ---- F/G/H列：IRP002 vs IRP001 ----
        irp002_data = irp002.get(irp002_key) if irp002_key else None
        if irp002_key and irp002_data is None and irp002:
            msgs.append(f"  F/G/H列: IRP002 中未找到「{irp002_key}」，留空")
        for col_offset, field, col_letter in [
            (6, 'principal', 'F'), (7, 'interest', 'G'), (8, 'total', 'H')
        ]:
            if _skip_check(skip_cols, col_letter):
                _write_cell(ws, excel_row, col_offset, '/', number_fmt=None,
                            fill=GRAY_FILL, font=NORMAL_FONT_V,
                            alignment=Alignment(horizontal='center', vertical='center'))
                continue

            irp002_val = irp002_data.get(field) if irp002_data else None
            is_count_item = label in ('被保险存款客户数', '最高偿付限额以内客户数')
            if is_count_item and col_letter == 'F':
                irp002_val = irp002_data.get('count') if irp002_data else None
                irp001_val = irp001_data.get('principal', 0)
            elif is_count_item:
                irp002_val = None
                irp001_val = 0
            else:
                irp001_val = irp001_data.get(field, 0) if irp001_data else 0

            if irp002_val is not None:
                if field == 'total' and irp001_val == 0 and irp001_data:
                    irp001_val = irp001_data.get('principal', 0)
                is_err = not _diff_ok(irp002_val, irp001_val)
                tolerance = 0.001
                if is_err:
                    error_count += 1
                    error_items.append(
                        f"第{excel_row}行「{label}」{col_letter}列: IRP002={irp002_val:,.2f} vs "
                        f"IRP001={irp001_val:,.2f}，差额={irp002_val - irp001_val:,.2f}元")
                result = 0 if not is_err else (irp002_val - irp001_val)
                _write_cell(ws, excel_row, col_offset, result, is_error=is_err)
                field_name = {'principal': '本金', 'interest': '利息', 'total': '本息合计'}[field]
                msgs.append(
                    f"  {col_letter}列 IRP002{field_name} vs IRP001: IRP002={irp002_val:,.2f}  "
                    f"IRP001={irp001_val:,.2f}  差额={irp002_val - irp001_val:,.2f}  "
                    f"容差=±{tolerance:g}元  [{'ERROR' if is_err else 'OK'}]")
            else:
                _write_cell(ws, excel_row, col_offset, None, number_fmt=None)

    # ---- 说明行 ----
    note_row = 22
    ws.merge_cells(f'A{note_row}:H{note_row}')
    _write_cell(ws, note_row, 1,
                '说明：灰色(/)部分不校验。红色底色单元格表示核对有误。'
                'B~H列：0=一致，非零值为差额（元）。'
                '三组核对（存款归属核对/会计科目核对/IRP002 vs IRP001）'
                '均采用绝对容差 0.001 元，超出即标红。'
                '调整项（用途=会计科目核对）已计入核对表，不再重复调整。'
                '住房公积金存款/社保基金存款仅核对存款归属列（合并模板未设会计科目行）。',
                number_fmt=None,
                font=Font(name='微软雅黑', size=9, italic=True, color="666666"),
                alignment=Alignment(horizontal='left', vertical='center'))

    # ---- 列宽调整 ----
    ws.column_dimensions['A'].width = 24
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws.column_dimensions[col_letter].width = 18

    # ---- 打印设置 ----
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(
        fitToPage=True)
    ws.page_setup.orientation = 'landscape'

    # ---- 保存 ----
    wb.save(str(output_path))
    wb.close()

    msgs.append("")
    msgs.append("========== 校验结果汇总 ==========")
    msgs.append(f"校验报告已生成: {output_path}")
    msgs.append(f"共校验 {len(_VALIDATION_ITEMS)} 个项目，发现 {error_count} 处差异")
    if error_count > 0:
        msgs.append("差异项明细:")
        for ei in error_items:
            msgs.append(f"  ✗ {ei}")
        msgs.append(">>> 请检查标红单元格对应的数据来源")
    else:
        msgs.append(">>> 所有校验项一致 ✓")
    return msgs
