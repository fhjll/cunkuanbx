# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险申报表审核工具 - 数据校验引擎
=============================================================================

本模块包含四个独立的校验功能，用于验证不同数据源之间的数据一致性：

  1. validate_irp002()                — IRP002 数据校验
       - 内部校验：最高偿付限额检查、户均本息检查、分组 vs 合计检查
       - 交叉校验：IRP002 vs IRP001 各项目数据一致性检查

  2. validate_deposit_vs_irp001()     — 存款归属表 vs IRP001 校验
       - 校验存款归属表（已生成）中的个人存款和单位存款
       - IRP001 数据扣除久悬户余额后精确至万元，与存款归属表比较

  3. validate_acct_vs_irp001()        — 科目余额表 vs IRP001 校验
       - 按科目代码从核对表明细行求和
       - 应用个体工商户调整值（c1/c2）后与 IRP001 比较

  4. generate_validation_report()     — 生成全面校验报告（Excel）
       - 基于校验表模板，从 IRP001、IRP002、存款归属表、科目余额表
         读取数据，逐项比较，生成标红（不一致）的 Excel 报告

核心校验逻辑：
  - 每一组比较都采用"一致→输出0，不一致→输出差额"的格式
  - 存款归属表 vs IRP001：不考虑误差，差额不为 0 即为不一致（直接标红）
  - 其他比较允许 ±max(1, 0.1%) 的容差范围
  - 校验报告用红色底色高亮不一致的单元格

调整值说明：
  - c1/c2：个体工商户本金/利息调整值，从单位扣减加到个人
  - c3/c4：个人/单位久悬户余额，从 IRP001 本金中扣除后再比较

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
from common import to_float, ExcelSheet as _ExcelSheet


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


def _find_xlsx_row(ws, keyword):
    """
    在 openpyxl Worksheet（1-indexed）中按 B 列关键词查找行号。

    参数:
        ws: openpyxl Worksheet
        keyword (str): 搜索关键词

    返回:
        int or None: 行号（1-indexed），未找到返回 None
    """
    for r in range(1, ws.max_row + 1):
        b_val = str(ws.cell(r, 2).value or '').strip()
        if keyword in b_val:
            return r
    return None


def _find_xlsx_rows_by_name(ws, keyword):
    """
    在 openpyxl Worksheet 中按 B 列关键词查找所有匹配行号。

    参数:
        ws: openpyxl Worksheet
        keyword (str): 搜索关键词

    返回:
        list[int]: 所有匹配行号列表
    """
    rows = []
    for r in range(1, ws.max_row + 1):
        b_val = str(ws.cell(r, 2).value or '').strip()
        if keyword in b_val:
            rows.append(r)
    return rows


def _find_xlsx_row_by_a(ws, code):
    """
    在 openpyxl Worksheet 中按 A 列精确匹配科目代码。

    自动处理 Excel 导出的 .0 后缀问题（如 "2003.0" → "2003"）。

    参数:
        ws: openpyxl Worksheet
        code (str): 科目代码

    返回:
        int or None: 行号（1-indexed）
    """
    for r in range(1, ws.max_row + 1):
        a_val = str(ws.cell(r, 1).value or '').strip()
        if '.' in a_val:
            try:
                a_val = str(int(float(a_val)))
            except:
                pass
        if a_val == code:
            return r
    return None


def _find_xlsx_rows_by_a_prefix(ws, prefix):
    """
    在 openpyxl Worksheet 中按 A 列代码前缀查找所有匹配行。

    参数:
        ws: openpyxl Worksheet
        prefix (str): 科目代码前缀

    返回:
        list[int]: 所有匹配行号列表
    """
    rows = []
    for r in range(1, ws.max_row + 1):
        a_val = str(ws.cell(r, 1).value or '').strip()
        if '.' in a_val:
            try:
                a_val = str(int(float(a_val)))
            except:
                pass
        if a_val == prefix or a_val.startswith(prefix):
            rows.append(r)
    return rows


def _find_calc_xlsx_row(ws, keyword):
    """
    在存款归属表中查找"计算行"行号。

    筛选条件：
      1. G列（备注）包含"计算行"
      2. D列（科目名称）包含目标关键词

    参数:
        ws: openpyxl Worksheet
        keyword (str): 科目名称关键词

    返回:
        int or None: 行号（1-indexed）
    """
    for r in range(1, ws.max_row + 1):
        remark = str(ws.cell(r, 7).value or '').strip()
        if '计算行' not in remark:
            continue
        d_val = str(ws.cell(r, 4).value or '').strip()
        if keyword in d_val:
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
            ok = diff <= max(1, abs(val_irp001 * 0.001))
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
# 2. 存款归属表 vs IRP001 校验
# ============================================================

def validate_deposit_vs_irp001(deposit_path, irp001_path, c3=0, c4=0):
    """
    校验存款归属表(已生成) vs IRP001 数据一致性。

    校验逻辑：
      1. 从存款归属表读取"计算行"的个人存款和单位存款金额（万元）
      2. 从 IRP001 读取个人存款本金和单位存款本金（元）
      3. IRP001 数据扣除久悬户余额 → 除以10000 → 精确至万元
      4. 与存款归属表的金额比较，绝对容差 0.01 元
         （仅消除浮点精度尾差，确保数据准确无误）

    参数:
        deposit_path (str): 存款归属表_已生成.xlsx 路径
        irp001_path (str): IRP001 申报表路径
        c3 (float): 个人久悬户余额（元）
        c4 (float): 单位久悬户余额（元）

    返回:
        list[str]: 校验结果消息列表
    """
    msgs = []
    try:
        wb_dep = openpyxl.load_workbook(str(deposit_path), data_only=True)
        ws_dep = wb_dep.active
    except Exception as e:
        return [f"ERROR: 无法读取存款归属表: {e}"]
    try:
        ws_irp = _ExcelSheet(str(irp001_path))
    except Exception as e:
        return [f"ERROR: 无法读取 IRP001: {e}"]

    c3_val = to_float(c3)
    c4_val = to_float(c4)

    # 查找存款归属表中"计算行"的个人存款和单位存款
    gr_row = _find_calc_xlsx_row(ws_dep, '个人存款')
    dw_row = _find_calc_xlsx_row(ws_dep, '单位存款')
    if gr_row is None:
        gr_row = _find_xlsx_row(ws_dep, '个人存款')
    if dw_row is None:
        dw_row = _find_xlsx_row(ws_dep, '单位存款')

    if gr_row is None:
        return ["ERROR: 在存款归属表中未找到「个人存款」行，请检查模板格式"]
    if dw_row is None:
        return ["ERROR: 在存款归属表中未找到「单位存款」行，请检查模板格式"]

    gr_dep_wan = to_float(ws_dep.cell(gr_row, 6).value)   # F列：金额（万元）
    dw_dep_wan = to_float(ws_dep.cell(dw_row, 6).value)

    # 查找 IRP001 中的个人/单位存款行
    irp_gr_row = _find_irp_cross_rows(ws_irp, '一、境内存款', '（一）个人存款')
    irp_dw_row = _find_irp_cross_rows(ws_irp, '一、境内存款', '（二）单位存款')
    if irp_gr_row is None:
        irp_gr_row = _find_irp_row(ws_irp, '个人存款')
    if irp_dw_row is None:
        irp_dw_row = _find_irp_row(ws_irp, '单位存款')

    irp_gr_yuan = _get_cell_val(ws_irp, irp_gr_row, 2) if irp_gr_row is not None else 0
    irp_dw_yuan = _get_cell_val(ws_irp, irp_dw_row, 2) if irp_dw_row is not None else 0

    # 扣除久悬户余额后，精确至万元
    irp_gr_adj_yuan = irp_gr_yuan - c3_val
    irp_dw_adj_yuan = irp_dw_yuan - c4_val
    irp_gr_adj_wan = round(irp_gr_adj_yuan / 10000)
    irp_dw_adj_wan = round(irp_dw_adj_yuan / 10000)

    # 绝对容差 0.01 元（仅消除浮点精度尾差，实际要求严格一致）
    tol_yuan = 0.01

    msgs.append("========== 存款归属表 vs IRP001 校验 ==========")
    msgs.append("校验逻辑: (IRP001本金 - 久悬户余额) / 10000 → 精确至万元 → 与存款归属表比较（绝对容差 0.01 元）")

    msgs.append("")
    msgs.append("--- 个人存款 ---")
    msgs.append(f"  IRP001 个人存款本金:                     {irp_gr_yuan:>16,.2f} 元")
    msgs.append(f"  - 个人久悬户余额 (c3):                   {c3_val:>16,.2f} 元")
    msgs.append(f"  调整后 IRP001 个人存款本金:               {irp_gr_adj_yuan:>16,.2f} 元")
    msgs.append(f"  精确至万元:                               {irp_gr_adj_wan:>16,.0f} 万元")
    msgs.append(f"  存款归属表 个人存款 (第{gr_row}行):           {gr_dep_wan:>16,.2f} 万元")
    gr_diff = abs(irp_gr_adj_wan - gr_dep_wan)
    gr_ok = gr_diff * 10000 <= tol_yuan  # 万元差额换算为元，绝对容差 0.01 元

    msgs.append("")
    msgs.append("--- 单位存款 ---")
    msgs.append(f"  IRP001 单位存款本金:                     {irp_dw_yuan:>16,.2f} 元")
    msgs.append(f"  - 单位久悬户余额 (c4):                   {c4_val:>16,.2f} 元")
    msgs.append(f"  调整后 IRP001 单位存款本金:               {irp_dw_adj_yuan:>16,.2f} 元")
    msgs.append(f"  精确至万元:                               {irp_dw_adj_wan:>16,.0f} 万元")
    msgs.append(f"  存款归属表 单位存款 (第{dw_row}行):           {dw_dep_wan:>16,.2f} 万元")
    dw_diff = abs(irp_dw_adj_wan - dw_dep_wan)
    dw_ok = dw_diff * 10000 <= tol_yuan  # 万元差额换算为元，绝对容差 0.01 元

    msgs.append("")
    msgs.append(f"  个人存款: {'✓ 一致' if gr_ok else '✗ 不一致'} (差额 {gr_diff * 10000:,.2f} 元，容差 {tol_yuan} 元)")
    msgs.append(f"  单位存款: {'✓ 一致' if dw_ok else '✗ 不一致'} (差额 {dw_diff * 10000:,.2f} 元，容差 {tol_yuan} 元)")
    if gr_ok and dw_ok:
        msgs.append("\n>>> 存款归属表校验结果: 全部通过")
    else:
        msgs.append("\n>>> 存款归属表校验结果: 存在差异，请检查")
    return msgs


# ============================================================
# 3. 科目余额表 vs IRP001 校验
# ============================================================

def _sum_codes_a(ws, codes, col=3):
    """
    在科目余额核对表中按 A 列代码精确查找并求和。

    参数:
        ws: openpyxl Worksheet
        codes (dict): {代码: 名称} 映射表
        col (int): 取值列（3=C列本金, 6=F列利息）

    返回:
        tuple: (total_sum, [(row, name, code, value), ...])
    """
    total = 0.0
    details = []
    for r in range(4, ws.max_row + 1):
        a_val = str(ws.cell(r, 1).value or '').strip()
        if '.' in a_val:
            try:
                a_val = str(int(float(a_val)))
            except:
                pass
        if a_val in codes:
            name = codes[a_val]
            val = to_float(ws.cell(r, col).value)
            total += val
            details.append((r, name, a_val, val))
    return total, details


def _sum_codes_d(ws, codes, col=6, a_filter=None):
    """
    在科目余额核对表中按 D 列代码精确查找并求和（用于利息代码）。

    参数:
        ws: openpyxl Worksheet
        codes (dict): {代码: 名称} 映射表
        col (int): 取值列（6=F列利息）
        a_filter (callable, optional): 可选的 A 列过滤函数，用于区分个人/单位

    返回:
        tuple: (total_sum, [(row, name, code, value), ...])
    """
    total = 0.0
    details = []
    for r in range(4, ws.max_row + 1):
        d_val = str(ws.cell(r, 4).value or '').strip()
        if '.' in d_val:
            try:
                d_val = str(int(float(d_val)))
            except:
                pass
        if d_val in codes:
            if a_filter is not None:
                a_val = str(ws.cell(r, 1).value or '').strip()
                if '.' in a_val:
                    try:
                        a_val = str(int(float(a_val)))
                    except:
                        pass
                if not a_filter(a_val):
                    continue
            name = codes[d_val]
            val = to_float(ws.cell(r, col).value)
            total += val
            details.append((r, name, d_val, val))
    return total, details


def validate_acct_vs_irp001(checktable_path, irp001_path, c1=0, c2=0):
    """
    校验科目余额核对表(已生成)与 IRP001 之间的数据一致性。

    校验逻辑：
      1. 从科目余额核对表的明细行，按科目代码（A列）和利息代码（D列）逐项求和
      2. 个体工商户调整值应用：
          个人存款 = 明细合计 + c1（本金）+ c2（利息）
          单位存款 = 明细合计 - c1（本金）- c2（利息）
      3. 与 IRP001 中对应项目的本金和利息比较
      4. 容差范围：≤1元

    参数:
        checktable_path (str): 科目余额核对表_已生成.xlsx 路径
        irp001_path (str): IRP001 申报表路径
        c1 (float): 个体工商户本金调整值（元）
        c2 (float): 个体工商户利息调整值（元）

    返回:
        list[str]: 校验结果消息列表
    """
    msgs = []
    adj_p = to_float(c1)
    adj_i = to_float(c2)

    try:
        wb = openpyxl.load_workbook(str(checktable_path), data_only=True)
        ws = wb.active
    except Exception as e:
        return [f"ERROR: 无法读取科目余额核对表: {e}"]

    try:
        ws_irp = _ExcelSheet(str(irp001_path))
    except Exception as e:
        return [f"ERROR: 无法读取 IRP001: {e}"]

    # ---- 从明细行按科目代码求和 ----
    # 个人存款本金（含久悬户 22410401）
    gr_p_total, gr_p_details = _sum_codes_a(ws, {
        '2003': '个人活期存款', '2004': '个人定期存款',
        '200503': '个人银行卡活期存款', '200504': '个人银行卡定期存款',
        '22410401': '久悬未取款项',
    }, col=3)
    # 保证金(个人部分) — 前缀 2014 + 尾缀 2
    gr_bzj_p = 0.0
    gr_bzj_p_details = []
    for r in range(4, ws.max_row + 1):
        a_val = str(ws.cell(r, 1).value or '').strip()
        if '.' in a_val:
            try:
                a_val = str(int(float(a_val)))
            except:
                pass
        if a_val.startswith('2014') and a_val.endswith('2'):
            val = to_float(ws.cell(r, 3).value)
            gr_bzj_p += val
            name = str(ws.cell(r, 2).value or '').strip() or f'保证金(个人)({a_val})'
            gr_bzj_p_details.append((r, name, a_val, val))
    gr_p_total += gr_bzj_p
    gr_p_details.extend(gr_bzj_p_details)

    # 单位存款本金（含久悬户 22410402/630108）
    dw_p_total, dw_p_details = _sum_codes_a(ws, {
        '2001': '单位活期存款', '2002': '单位定期存款',
        '200501': '单位银行卡活期存款', '200502': '单位银行卡定期存款',
        '200601': '财政性存款', '2011': '应解汇款(单位部分)',
        '22410402': '久悬不动户(对公)', '630108': '久悬未取款项收入',
    }, col=3)
    # 保证金(单位部分) — 前缀 2014 + 尾缀 1
    dw_bzj_p = 0.0
    dw_bzj_p_details = []
    for r in range(4, ws.max_row + 1):
        a_val = str(ws.cell(r, 1).value or '').strip()
        if '.' in a_val:
            try:
                a_val = str(int(float(a_val)))
            except:
                pass
        if a_val.startswith('2014') and a_val.endswith('1'):
            val = to_float(ws.cell(r, 3).value)
            dw_bzj_p += val
            name = str(ws.cell(r, 2).value or '').strip() or f'保证金(单位)({a_val})'
            dw_bzj_p_details.append((r, name, a_val, val))
    dw_p_total += dw_bzj_p
    dw_p_details.extend(dw_bzj_p_details)

    # 个人存款利息
    gr_i_codes = {
        '223103': '个人活期存款应付利息', '223104': '个人定期存款应付利息',
        '22310503': '个人银行卡活期存款应付利息', '22310504': '个人银行卡定期存款应付利息',
    }
    gr_i_total, gr_i_details = _sum_codes_d(ws, gr_i_codes, col=6)
    gr_bzj_i_total, gr_bzj_i_details = _sum_codes_d(
        ws, {'223108': '保证金应付利息(个人部分)'},
        col=6, a_filter=lambda a: a.endswith('2') if a else True)
    gr_i_total += gr_bzj_i_total
    gr_i_details.extend(gr_bzj_i_details)

    # 单位存款利息
    dw_i_codes = {
        '223101': '单位活期存款应付利息', '223102': '单位定期存款应付利息',
        '22310501': '单位银行卡活期存款应付利息', '22310502': '单位银行卡定期存款应付利息',
        '22310601': '财政性存款应付利息',
    }
    dw_i_total, dw_i_details = _sum_codes_d(ws, dw_i_codes, col=6)
    dw_bzj_i_total, dw_bzj_i_details = _sum_codes_d(
        ws, {'223108': '保证金应付利息(单位部分)'},
        col=6, a_filter=lambda a: a.endswith('1') if a else True)
    dw_i_total += dw_bzj_i_total
    dw_i_details.extend(dw_bzj_i_details)

    # ---- 应用个体工商户调整值 ----
    gr_adjusted_p = gr_p_total + adj_p
    dw_adjusted_p = dw_p_total - adj_p
    gr_adjusted_i = gr_i_total + adj_i
    dw_adjusted_i = dw_i_total - adj_i

    # ---- 查找 IRP001 目标行 ----
    irp_gr_row = _find_irp_cross_rows(ws_irp, '一、境内存款', '（一）个人存款')
    irp_dw_row = _find_irp_cross_rows(ws_irp, '一、境内存款', '（二）单位存款')
    if irp_gr_row is None:
        irp_gr_row = _find_irp_row(ws_irp, '个人存款')
    if irp_dw_row is None:
        irp_dw_row = _find_irp_row(ws_irp, '单位存款')

    irp_gr_p = _get_cell_val(ws_irp, irp_gr_row, 2) if irp_gr_row is not None else 0
    irp_dw_p = _get_cell_val(ws_irp, irp_dw_row, 2) if irp_dw_row is not None else 0
    irp_gr_i = _get_cell_val(ws_irp, irp_gr_row, 3) if irp_gr_row is not None else 0
    irp_dw_i = _get_cell_val(ws_irp, irp_dw_row, 3) if irp_dw_row is not None else 0

    msgs.append("========== 科目余额核对表 vs IRP001 校验 ==========")
    msgs.append(f"个体工商户本金调整值: {adj_p:,.2f} 元")
    msgs.append(f"个体工商户利息调整值: {adj_i:,.2f} 元")
    msgs.append("数据来源: 明细行按科目代码求和")
    msgs.append("")

    # --- 本金校验 ---
    msgs.append("--- 本金校验 ---")
    msgs.append("个人存款明细(科目余额核对表):")
    for r, name, code, val in gr_p_details:
        code_str = str(code) if code else ""
        msgs.append(f"  R{r:2d} {name:28s}({code_str:>8s}): {val:>16,.2f}")
    msgs.append(f"  {'':34s}{'合计':>8s}: {gr_p_total:>16,.2f}")
    msgs.append(f"  + 个体工商户本金调整值:       {adj_p:>16,.2f}")
    msgs.append(f"  调整后:                       {gr_adjusted_p:>16,.2f}")
    irp_row_info = f"(IRP001第{irp_gr_row}行)" if irp_gr_row else ""
    msgs.append(f"  IRP001 个人存款本金 {irp_row_info}: {irp_gr_p:>16,.2f}")
    gr_p_diff = abs(gr_adjusted_p - irp_gr_p)
    gr_p_ok = gr_p_diff <= 1
    msgs.append(f"  校验: {'一致 (差额 %.2f)' % gr_p_diff if gr_p_ok else '不一致 (差额 %.2f 元)' % gr_p_diff}")

    msgs.append("")
    msgs.append("单位存款明细(科目余额核对表):")
    for r, name, code, val in dw_p_details:
        code_str = str(code) if code else ""
        msgs.append(f"  R{r:2d} {name:28s}({code_str:>8s}): {val:>16,.2f}")
    msgs.append(f"  {'':34s}{'合计':>8s}: {dw_p_total:>16,.2f}")
    msgs.append(f"  - 个体工商户本金调整值:       {adj_p:>16,.2f}")
    msgs.append(f"  调整后:                       {dw_adjusted_p:>16,.2f}")
    irp_row_info = f"(IRP001第{irp_dw_row}行)" if irp_dw_row else ""
    msgs.append(f"  IRP001 单位存款本金 {irp_row_info}: {irp_dw_p:>16,.2f}")
    dw_p_diff = abs(dw_adjusted_p - irp_dw_p)
    dw_p_ok = dw_p_diff <= 1
    msgs.append(f"  校验: {'一致 (差额 %.2f)' % dw_p_diff if dw_p_ok else '不一致 (差额 %.2f 元)' % dw_p_diff}")

    # --- 利息校验 ---
    msgs.append("")
    msgs.append("--- 应付利息校验 ---")
    msgs.append("个人存款利息明细(科目余额核对表):")
    for r, name, code, val in gr_i_details:
        code_str = str(code) if code else ""
        msgs.append(f"  R{r:2d} {name:28s}({code_str:>8s}): {val:>16,.2f}")
    msgs.append(f"  {'':34s}{'合计':>8s}: {gr_i_total:>16,.2f}")
    msgs.append(f"  + 个体工商户利息调整值:       {adj_i:>16,.2f}")
    msgs.append(f"  调整后:                       {gr_adjusted_i:>16,.2f}")
    msgs.append(f"  IRP001 个人存款应付利息:      {irp_gr_i:>16,.2f}")
    gr_i_diff = abs(gr_adjusted_i - irp_gr_i)
    gr_i_ok = gr_i_diff <= 1
    msgs.append(f"  校验: {'一致 (差额 %.2f)' % gr_i_diff if gr_i_ok else '不一致 (差额 %.2f 元)' % gr_i_diff}")

    msgs.append("")
    msgs.append("单位存款利息明细(科目余额核对表):")
    for r, name, code, val in dw_i_details:
        code_str = str(code) if code else ""
        msgs.append(f"  R{r:2d} {name:28s}({code_str:>8s}): {val:>16,.2f}")
    msgs.append(f"  {'':34s}{'合计':>8s}: {dw_i_total:>16,.2f}")
    msgs.append(f"  - 个体工商户利息调整值:       {adj_i:>16,.2f}")
    msgs.append(f"  调整后:                       {dw_adjusted_i:>16,.2f}")
    msgs.append(f"  IRP001 单位存款应付利息:      {irp_dw_i:>16,.2f}")
    dw_i_diff = abs(dw_adjusted_i - irp_dw_i)
    dw_i_ok = dw_i_diff <= 1
    msgs.append(f"  校验: {'一致 (差额 %.2f)' % dw_i_diff if dw_i_ok else '不一致 (差额 %.2f 元)' % dw_i_diff}")

    all_ok = gr_p_ok and dw_p_ok and gr_i_ok and dw_i_ok
    if all_ok:
        msgs.append("\n>>> 校验结果: 全部通过")
    else:
        msgs.append("\n>>> 校验结果: 存在差异，请检查")
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


def _read_deposit_map(filepath):
    """
    读取存款归属表的"计算行"数据（万元）。

    通过扫描 B 列（科目名称）和 G 列（备注）提取所有"计算行"的值。
    支持分境内/境外两部分。

    参数:
        filepath (str): 存款归属表文件路径

    返回:
        dict: {项目标签: 万元值}
    """
    try:
        wb = openpyxl.load_workbook(str(filepath), data_only=True)
        ws = wb.active
    except Exception:
        return {}

    out = {}
    current_section = '境内'
    for r in range(5, ws.max_row + 1):
        col_b = str(ws.cell(r, 2).value or '').strip()
        if '境外' in col_b:
            current_section = '境外'
        elif '境内' in col_b:
            current_section = '境内'

        remark = str(ws.cell(r, 7).value or '').strip()
        if '计算行' not in remark:
            continue
        sub_name = str(ws.cell(r, 4).value or '').strip()
        val = to_float(ws.cell(r, 6).value)

        sub_name_clean = sub_name.replace('  ', ' ').replace('. ', '.').strip()
        for prefix in ['其中：', '其中:', '     ', '    ', '   ', '  ']:
            if sub_name_clean.startswith(prefix):
                sub_name_clean = sub_name_clean[len(prefix):]

        if current_section == '境外':
            label = '境外' + sub_name_clean
        else:
            label = sub_name_clean

        if label not in out:
            out[label] = val

    # 计算境内/境外存款合计
    # 境内存款 = 个人 + 单位 + 财政性 + 非存款类（其中项住房公积金/
    # 社保基金存款已包含在单位存款中，不重复计入）
    dom_sum = sum(out.get(k, 0) for k in [
        '（一）个人存款', '（二）单位存款',
        '（三）财政性存款', '（四）非存款类金融机构存款'])
    if '一、境内存款' not in out:
        out['一、境内存款'] = dom_sum

    ovs_sum = sum(out.get(k, 0) for k in [
        '境外（一）个人存款', '境外（二）单位存款', '境外（三）境外同业存放'])
    if '二、境外存款' not in out:
        out['二、境外存款'] = ovs_sum

    wb.close()
    return out


def _read_acct_check_map(filepath):
    """
    读取科目余额核对表的汇总行数据（元）。

    通过扫描 B 列（科目名称）关键词提取汇总行数据，
    返回 {label: {principal, interest, total}} 字典。

    参数:
        filepath (str): 科目余额核对表文件路径

    返回:
        dict: {项目标签: {principal, interest, total}}
    """
    try:
        wb = openpyxl.load_workbook(str(filepath), data_only=True)
        ws = wb.active
    except Exception:
        return {}

    out = {}

    # 按关键词匹配汇总行
    item_keywords = [
        ('（一）个人存款',        '（一）个人存款'),
        ('（二）单位存款',        '（二）单位存款'),
        ('住房公积金存款',        '住房公积金存款'),
        ('社保基金存款',          '社保基金存款'),
        ('（三）财政性存款',      '（三）财政性存款'),
        ('（四）非存款类金融机构存款', '（四）非存款类金融机构存款'),
        ('境外同业存放',          '（三）境外同业存放'),
        ('高级管理人员',          '高级管理人员'),
        ('被保险存款余额',        '被保险存款余额'),
    ]
    for b_kw, label in item_keywords:
        for r in range(4, ws.max_row + 1):
            b_val = str(ws.cell(r, 2).value or ws.cell(r, 1).value or '').strip()
            if b_kw in b_val:
                principal = to_float(ws.cell(r, 3).value)
                interest = to_float(ws.cell(r, 6).value)
                out[label] = {
                    'principal': principal,
                    'interest': interest,
                    'total': principal + interest,
                }
                break

    # 国库定期存款 & 代理财政性存款 — 从科目代码读取
    fiscal_map = {
        '1.国库定期存款':   ('200602', '22310602'),
        '2.代理财政性存款': ('2007',   '223107'),
    }
    for label, (code_p, code_i) in fiscal_map.items():
        p_val = _sum_codes_a(ws, {code_p: code_p}, col=3)[0]
        i_val = _sum_codes_d(ws, {code_i: code_i}, col=6)[0]
        out[label] = {'principal': p_val, 'interest': i_val, 'total': p_val + i_val}

    # 财政性存款汇总 — 优先用 B 列匹配值；若为 0 则从子项明细求和
    fiscal_summary = out.get('（三）财政性存款')
    if fiscal_summary is None or fiscal_summary.get('principal', 0) == 0:
        sub_p = out.get('1.国库定期存款', {}).get('principal', 0)
        sub_i = out.get('1.国库定期存款', {}).get('interest', 0)
        sub_p += out.get('2.代理财政性存款', {}).get('principal', 0)
        sub_i += out.get('2.代理财政性存款', {}).get('interest', 0)
        if sub_p != 0 or sub_i != 0:
            out['（三）财政性存款'] = {
                'principal': sub_p, 'interest': sub_i, 'total': sub_p + sub_i,
            }

    # 境内存款合计行
    for r in range(4, ws.max_row + 1):
        cell_val = str(ws.cell(r, 1).value or ws.cell(r, 2).value or '').strip()
        if '境内存款' in cell_val and '（一）' not in cell_val and '（二）' not in cell_val:
            p_val = to_float(ws.cell(r, 3).value)
            i_val = to_float(ws.cell(r, 6).value)
            if p_val == 0 and i_val == 0:
                # 兜底求和：个人+单位+财政性+非存款类（其中项住房公积金/
                # 社保基金存款已含在单位存款中，不重复计入）
                sub_p = sum(out.get(k, {}).get('principal', 0) for k in [
                    '（一）个人存款', '（二）单位存款',
                    '（三）财政性存款', '（四）非存款类金融机构存款'])
                sub_i = sum(out.get(k, {}).get('interest', 0) for k in [
                    '（一）个人存款', '（二）单位存款',
                    '（三）财政性存款', '（四）非存款类金融机构存款'])
                if sub_p != 0 or sub_i != 0:
                    p_val, i_val = sub_p, sub_i
            out['一、境内存款'] = {
                'principal': p_val, 'interest': i_val, 'total': p_val + i_val,
            }

    # 境外存款合计行
    for r in range(4, ws.max_row + 1):
        cell_val = str(ws.cell(r, 1).value or ws.cell(r, 2).value or '').strip()
        if '境外存款' in cell_val and '（一）' not in cell_val and \
           '（三）' not in cell_val and '利息' not in cell_val:
            out['二、境外存款'] = {
                'principal': to_float(ws.cell(r, 3).value),
                'interest': to_float(ws.cell(r, 6).value),
                'total': to_float(ws.cell(r, 3).value) + to_float(ws.cell(r, 6).value),
            }

    wb.close()
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

    容差规则：允许差额 ≤ max(1, abs(v2 * 0.001))
    即至少允许 1 元的绝对误差，同时允许 0.1% 的相对误差。

    参数:
        v1 (float): 值1
        v2 (float): 值2（作为基准值）

    返回:
        bool: True 表示在容差范围内一致
    """
    return abs(v1 - v2) <= max(1, abs(v2 * 0.001))


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
                                irp002_path=None, deposit_path=None,
                                acct_check_path=None,
                                c1=0, c2=0, c3=0, c4=0,
                                template_path=None):
    """
    基于校验表模板，生成全面校验报告（Excel 格式）。

    校验范围涵盖最多 4 个数据源的交叉比对：
      - IRP001（基准数据源，必填）
      - IRP002（可选）
      - 存款归属表_已生成（可选）
      - 科目余额核对表_已生成（可选）

    报告格式：
      - 标题行：存保申报数据校验表
      - 分组表头：数据来源分组（存款归属表、科目余额核对表、IRP002）
      - 数据行：每行一个校验项，输出 0（一致）或差额（不一致）
      - 差异行自动标红（红色底色）

    参数:
        irp001_path:     IRP001 申报表路径（必填，作为基准数据源）
        output_path:     输出 .xlsx 路径（必填）
        irp002_path:     IRP002 申报表路径（可选）
        deposit_path:    存款归属表_已生成.xlsx 路径（可选）
        acct_check_path: 科目余额核对表_已生成.xlsx 路径（可选）
        c1, c2:          个体工商户调整值（本金/利息，元）
                         个人 += c1/c2，单位 -= c1/c2
        c3, c4:          久悬户余额（个人/单位，元）
                         从 IRP001 本金中扣除后精确至万元再比较
        template_path:   校验表（表样）.xls 模板（可选，预留给未来使用）

    返回:
        list[str]: 运行日志消息列表
    """
    _init_styles()
    msgs = []

    adj_p = to_float(c1)
    adj_i = to_float(c2)
    adj_c3 = to_float(c3)   # 个人久悬户余额（元）
    adj_c4 = to_float(c4)   # 单位久悬户余额（元）

    # ---- 读取各数据源 ----
    msgs.append("========== 校验参数 ==========")
    msgs.append(f"  c1 个体工商户本金调整值: {adj_p:,.2f} 元")
    msgs.append(f"  c2 个体工商户利息调整值: {adj_i:,.2f} 元")
    msgs.append(f"  c3 个人久悬户余额:       {adj_c3:,.2f} 元")
    msgs.append(f"  c4 单位久悬户余额:       {adj_c4:,.2f} 元")
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
    if deposit_path and os.path.exists(str(deposit_path)):
        deposit = _read_deposit_map(deposit_path)
        msgs.append(f"  存款归属表: 读取 {len(deposit)} 个计算行")
    else:
        msgs.append("  存款归属表: 未提供，跳过 B 列校验")

    acct = {}
    if acct_check_path and os.path.exists(str(acct_check_path)):
        acct = _read_acct_check_map(acct_check_path)
        msgs.append(f"  科目余额核对表: 读取 {len(acct)} 个汇总项")
    else:
        msgs.append("  科目余额核对表: 未提供，跳过 C/D/E 列校验")
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

        # ---- B列：存款归属表 vs IRP001 ----
        if not _skip_check(skip_cols, 'B'):
            dep_val_wan = deposit.get(dep_key) if dep_key else None
            if dep_val_wan is not None:
                irp001_principal = irp001_data.get('principal', 0)
                # 应用久悬户余额调整
                adj_note = ""
                if label == '（一）个人存款':
                    irp001_principal -= adj_c3
                    adj_note = f"（已扣除个人久悬户 c3={adj_c3:,.2f}）"
                elif label == '（二）单位存款':
                    irp001_principal -= adj_c4
                    adj_note = f"（已扣除单位久悬户 c4={adj_c4:,.2f}）"
                irp_in_wan = round(irp001_principal / 10000)
                diff_wan = dep_val_wan - irp_in_wan
                is_err = diff_wan != 0  # 无容差，差额不为0即标红
                if is_err:
                    error_count += 1
                    error_items.append(
                        f"第{excel_row}行「{label}」B列: 存款归属表={dep_val_wan:,.2f}万元 vs "
                        f"IRP001={irp_in_wan:,.0f}万元，差额={diff_wan:,.2f}万元")
                result = 0 if not is_err else diff_wan * 10000
                _write_cell(ws, excel_row, 2, result, is_error=is_err)
                msgs.append(
                    f"  B列 存款归属表 vs IRP001: 归属表={dep_val_wan:,.2f}万元  "
                    f"IRP001调整后={irp_in_wan:,.0f}万元{adj_note}  "
                    f"差额={diff_wan:,.2f}万元  [{'ERROR' if is_err else 'OK'}]")
            else:
                _write_cell(ws, excel_row, 2, None, number_fmt=None)
                if dep_key and deposit:
                    msgs.append(f"  B列: 存款归属表中未找到「{dep_key}」，留空")
        else:
            _write_cell(ws, excel_row, 2, '/', number_fmt=None,
                        fill=GRAY_FILL, font=NORMAL_FONT_V,
                        alignment=Alignment(horizontal='center', vertical='center'))

        # ---- C/D/E列：科目余额核对表 vs IRP001 ----
        acct_data = acct.get(acct_key) if acct_key else None
        if acct_key and acct_data is None and acct:
            msgs.append(f"  C/D/E列: 科目余额核对表中未找到「{acct_key}」，留空")
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
                adj_note = ""
                if label == '（一）个人存款':
                    if field == 'principal':
                        adjusted_val = acct_val + adj_p
                        adj_note = f"（+c1 {adj_p:,.2f}）"
                    elif field == 'interest':
                        adjusted_val = acct_val + adj_i
                        adj_note = f"（+c2 {adj_i:,.2f}）"
                    elif field == 'total':
                        adjusted_val = acct_val + adj_p + adj_i
                        adj_note = f"（+c1+c2 {adj_p + adj_i:,.2f}）"
                elif label == '（二）单位存款':
                    if field == 'principal':
                        adjusted_val = acct_val - adj_p
                        adj_note = f"（-c1 {adj_p:,.2f}）"
                    elif field == 'interest':
                        adjusted_val = acct_val - adj_i
                        adj_note = f"（-c2 {adj_i:,.2f}）"
                    elif field == 'total':
                        adjusted_val = acct_val - adj_p - adj_i
                        adj_note = f"（-c1-c2 {adj_p + adj_i:,.2f}）"
                is_err = not _diff_ok(adjusted_val, irp001_val)
                tolerance = max(1, abs(irp001_val * 0.001))
                if is_err:
                    error_count += 1
                    error_items.append(
                        f"第{excel_row}行「{label}」{col_letter}列: 核对表调整后={adjusted_val:,.2f} vs "
                        f"IRP001={irp001_val:,.2f}，差额={adjusted_val - irp001_val:,.2f}元")
                result = 0 if not is_err else (adjusted_val - irp001_val)
                _write_cell(ws, excel_row, col_offset, result, is_error=is_err)
                field_name = {'principal': '本金', 'interest': '利息', 'total': '本息合计'}[field]
                msgs.append(
                    f"  {col_letter}列 核对表{field_name} vs IRP001: 核对表={acct_val:,.2f}{adj_note}  "
                    f"IRP001={irp001_val:,.2f}  差额={adjusted_val - irp001_val:,.2f}  "
                    f"容差=±{tolerance:,.2f}  [{'ERROR' if is_err else 'OK'}]")
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
                tolerance = max(1, abs(irp001_val * 0.001))
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
                    f"容差=±{tolerance:,.2f}  [{'ERROR' if is_err else 'OK'}]")
            else:
                _write_cell(ws, excel_row, col_offset, None, number_fmt=None)

    # ---- 说明行 ----
    note_row = 22
    ws.merge_cells(f'A{note_row}:H{note_row}')
    _write_cell(ws, note_row, 1,
                '说明：灰色(/)部分不校验。红色底色单元格表示核对有误。'
                'B~E列：0=一致，非零值为差额（元）。'
                'B列（存款归属表核对）无容差，差额不为0即标红；'
                'C~E列差额容差 ±max(1, 0.1%)。',
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
