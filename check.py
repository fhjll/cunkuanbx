# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险数据校验脚本（旧版引擎，保留兼容）
=============================================================================

【说明】本文件是旧版校验引擎，通过 pandas 直接操作 Excel 数据。
        新校验功能已迁移到 validator.py 中（功能更完善、逻辑更清晰）。
        本文件保留以兼容旧项目引用。

功能：
  从工作目录读取各申报文件和附件，逐项校验数据一致性。

校验项包括：
  (1) 个人存款本金（科目余额表 VS IRP001 VS A1411 资产负债表）
  (2) 单位存款本金（科目余额表 VS IRP001 VS A1411 资产负债表）
  (3) 个人存款利息（科目余额表 VS IRP001）
  (4) 单位存款利息（科目余额表 VS IRP001）
  (5) 利息和（IRP001 VS A1411 资产负债表）
  (6) 住房公积金存款（附件2明细 VS IRP001 VS A1411 资产负债表）
  (7) 社保基金存款（附件3明细 VS IRP001 VS A1411 资产负债表）
  (8) 财政性存款（科目余额表 VS IRP001）
  (9) 非存款类金融机构存款（科目余额表 VS IRP001）
  (10) 境外存款（是否为0）
  (11) 高管存款（附件5明细 VS IRP001）

科目代码说明：
  2001=单位活期存款 2002=单位定期存款
  2003=个人活期存款 2004=个人定期存款
  200501/200502=单位银行卡 200503/200504=个人银行卡
  200601/200602=财政性存款 2007=代理财政性存款
  2011=应解汇款 2014xxx1/2=保证金(单位/个人)
  223101~223108=各类应付利息
  22410401/2=久悬未取款项(个人/对公)
  630108=久悬未取款项收入

依赖：
  - pandas：读取和操作 Excel 数据
  - openpyxl/xlrd：pandas 底层引擎
=============================================================================
"""
import pandas as pd
import os


# ============================================================
# 工具函数
# ============================================================

def _find_excel(path_prefix):
    """
    自动查找 .xls 或 .xlsx 文件，返回存在的路径。

    参数:
        path_prefix (str): 不带扩展名的文件路径

    返回:
        str: 存在的文件路径
    """
    for ext in ['.xls', '.xlsx']:
        p = path_prefix + ext
        if os.path.exists(p):
            return p
    if os.path.exists(path_prefix):
        return path_prefix
    return path_prefix + '.xls'


def quzhi(data, ind, col="Unnamed: 7"):
    """
    从 DataFrame 中按索引取值。

    参数:
        data (pd.DataFrame): 科目余额表 DataFrame
        ind: 索引值（科目代码）
        col (str): 列名，默认 "Unnamed: 7"

    返回:
        float: 对应的值，未找到返回 0
    """
    if ind in data.index:
        return data.loc[ind, col]
    else:
        return 0


# ============================================================
# 各校验函数
# ============================================================

# ---- 个人存款本金 ----

def check_grck(data, kmye, ce, col="Unnamed: 7"):
    """
    校验个人存款本金（IRP001 VS 科目余额表）。

    个人存款 = 活期(2003) + 定期(2004) + 银行卡活期(200503)
            + 银行卡定期(200504) + 保证金个人(2014xxx2)
            + 久悬户(22410401)
    """
    grhq = quzhi(kmye, "2003", col) + float(ce)
    grdq = quzhi(kmye, "2004", col)
    grkhq = quzhi(kmye, "200503", col)
    grkdq = quzhi(kmye, "200504", col)
    grbzj = (quzhi(kmye, "20140102", col) + quzhi(kmye, "20140202", col)
             + quzhi(kmye, "20140302", col) + quzhi(kmye, "20140402", col)
             + quzhi(kmye, "201405302", col) + quzhi(kmye, "20140602", col))
    grjxh = quzhi(kmye, "22410401", col)
    print(grjxh)
    print(round(data.iloc[12, 2], 2)
          - round(grhq + grdq + grkhq + grkdq + grbzj + grjxh, 2))
    if round(data.iloc[12, 2], 2) == round(grhq + grdq + grkhq + grkdq + grbzj + grjxh, 2):
        return "个人存款校验成功\n"
    else:
        return "个人存款数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_grckA(data1, ce, data2):
    """
    校验个人存款与 A1411 资产负债表的一致性。
    """
    gjxh = float(ce)
    print(gjxh)
    left = (data1.at["12MI1", "本期余额"] - data1.at["12MI3", "本期余额"]
            - data1.at["12MI5", "本期余额"] - data1.at["12MI7", "本期余额"]
            - data1.at["12MI9", "本期余额"] - data1.at["12MIB", "本期余额"]
            - data1.at["12MID", "本期余额"] - data1.at["12MIF", "本期余额"]
            - data1.at["12MIH", "本期余额"] - data1.at["12N1A", "本期余额"]
            + data1.at["12MW5", "本期余额"])
    right = round((data2.iloc[12, -3] - gjxh) / 10000, 0)
    print(left - right)
    if left == right:
        return "个人存款与A1411校验成功\n"
    else:
        return "个人存款与A1411校验错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 单位存款本金 ----

def check_dwck(data, kmye, ce, col="Unnamed: 7"):
    """
    校验单位存款本金（IRP001 VS 科目余额表）。

    单位存款 = 活期(2001) + 定期(2002) + 银行卡活期(200501)
            + 银行卡定期(200502) + 财政性(200601) + 应解汇款(2011)
            + 保证金单位(2014xxx1) + 久悬(22410402+630108)
    """
    dwhq = quzhi(kmye, "2001", col) - float(ce)
    dwdq = quzhi(kmye, "2002", col)
    dwkhq = quzhi(kmye, "200501", col)
    dwkdq = quzhi(kmye, "200502", col)
    dwczck = quzhi(kmye, "200601", col)
    dwyjk = quzhi(kmye, "2011", col)
    dwbzj = (quzhi(kmye, "20140101", col) + quzhi(kmye, "20140201", col)
             + quzhi(kmye, "20140301", col) + quzhi(kmye, "20140401", col)
             + quzhi(kmye, "20140501", col) + quzhi(kmye, "20140601", col))
    dwjxh = (quzhi(kmye, "22410402", col) + quzhi(kmye, "63010801", col)
             + quzhi(kmye, "63010802", col))
    print(dwjxh)
    print(round(data.iloc[13, 2], 2)
          - round(dwhq + dwdq + dwkhq + dwkdq + dwbzj + dwyjk + dwczck + dwjxh, 2))
    if round(data.iloc[13, 2], 2) == round(
            dwhq + dwdq + dwkhq + dwkdq + dwbzj + dwyjk + dwczck + dwjxh, 2):
        return "单位存款校验成功\n"
    else:
        return "单位存款数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_dwckA(data1, ce, data2):
    """
    校验单位存款与 A1411 资产负债表的一致性。
    """
    djxh = float(ce)
    print(djxh)
    left = (data1.at["12MG3", "本期余额"] - data1.at["12MGA", "本期余额"]
            - data1.at["12MGH", "本期余额"] - data1.at["12MGQ", "本期余额"]
            - data1.at["12MH3", "本期余额"] - data1.at["12MHA", "本期余额"]
            - data1.at["12MHH", "本期余额"] - data1.at["12MHQ", "本期余额"]
            - data1.at["12N18", "本期余额"] + data1.at["12MIV", "本期余额"]
            - data1.at["12MIX", "本期余额"] + data1.at["12MW6", "本期余额"]
            + data1.at["12MW7", "本期余额"] + data1.at["12MW8", "本期余额"]
            + data1.at["12MW9", "本期余额"] + data1.at["12MWA", "本期余额"])
    right = round((data2.iloc[13, -3] - float(ce)) / 10000, 0)
    print(left - right)
    if left == right:
        return "单位存款与A1411校验成功\n"
    else:
        return "单位存款与A1411校验错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 个人存款利息 ----

def check_grcklx(data, kmye, ce, col="Unnamed: 7"):
    """
    校验个人存款应付利息（IRP001 VS 科目余额表）。
    """
    grhqlx = quzhi(kmye, "223103", col)
    grdqlx = quzhi(kmye, "223104", col)
    grhqlx_ka = quzhi(kmye, "22310503", col)
    grdqlx_ka = quzhi(kmye, "22310504", col)
    grbzjlx = float(ce)
    if round(data.iloc[12, 3], 2) == round(
            grhqlx + grdqlx + grhqlx_ka + grdqlx_ka + grbzjlx, 2):
        return "个人存款利息校验成功\n"
    else:
        return "个人存款利息数据错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 单位存款利息 ----

def check_dwcklx(data, kmye, ce, col="Unnamed: 7"):
    """
    校验单位存款应付利息（IRP001 VS 科目余额表）。
    """
    dwhqlx = quzhi(kmye, "223101", col)
    dwdqlx = quzhi(kmye, "223102", col)
    dwhqlx_ka = quzhi(kmye, "22310501", col)
    dwdqlx_ka = quzhi(kmye, "22310502", col)
    dwczcklx = quzhi(kmye, "22310601", col)
    dwbzjlx = quzhi(kmye, "223108", col) - float(ce)
    if round(data.iloc[13, 3], 2) == round(
            dwhqlx + dwdqlx + dwhqlx_ka + dwdqlx_ka + dwczcklx + dwbzjlx, 2):
        return "单位存款利息校验成功\n"
    else:
        return "单位存款利息数据错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 利息和（IRP001 VS A1411） ----

def check_lxhA(data1, data2):
    """校验利息和与 A1411 数据的一致性"""
    if data1.at["12MR8", "本期余额"] == round(
            data2.iloc[12, -2] / 10000 + data2.iloc[13, -2] / 10000, 0):
        return "利息和与A1411校验成功\n"
    else:
        return "利息和与A1411校验错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 住房公积金存款 ----

def check_gjj(data, data1):
    """校验住房公积金存款（附件2明细 VS IRP001）"""
    if data1["余额"].iloc[:-2].fillna(0).sum() == data1['余额'].iloc[-2]:
        if round(data1['余额'].iloc[-2], 2) == round(data.iloc[14, 2], 2):
            return "公积金数据校验成功\n"
        else:
            return "公积金数据核对错误!!!!!!!!!!!!!!!!!!！\n"
    else:
        return "公积金内部数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_gjjA(data1, data2):
    """校验住房公积金与 A1411 的一致性"""
    left = (data1.at["12MG9", "本期余额"] + data1.at["12MGG", "本期余额"]
            + data1.at["12MGP", "本期余额"] + data1.at["12MH2", "本期余额"]
            + data1.at["12MH9", "本期余额"] + data1.at["12MHG", "本期余额"]
            + data1.at["12MHP", "本期余额"] + data1.at["12N17", "本期余额"]
            + data1.at["12MWA", "本期余额"])
    print(left)
    print(round(data2.iloc[14, -3] / 10000, 0))
    if left == round(data2.iloc[14, -3] / 10000, 0):
        return "住房公积金存款与A1411校验成功\n"
    else:
        return "住房公积金存款与A1411校验错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 社保基金存款 ----

def check_sbjj(data, data1):
    """校验社保基金存款（附件3明细 VS IRP001）"""
    if round(data1["余额"].iloc[:-2].fillna(0).sum(), 2) == round(
            data1['余额'].iloc[-2], 2):
        if round(data1["余额"].iloc[:-2].sum(), 2) == round(data.iloc[15, 2], 2):
            return "社保基金数据校验成功\n"
        else:
            return "社保基金数据核对错误!!!!!!!!!!!!!!!!!!！\n"
    else:
        return "社保基金内部数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_sbjjA(data1, data2):
    """校验社保基金与 A1411 的一致性"""
    left = (data1.at["12MG7", "本期余额"] + data1.at["12MGE", "本期余额"]
            + data1.at["12MGM", "本期余额"] + data1.at["12MGU", "本期余额"]
            + data1.at["12MH7", "本期余额"] + data1.at["12MHE", "本期余额"]
            + data1.at["12MHM", "本期余额"] + data1.at["12N15", "本期余额"]
            + data1.at["12MW8", "本期余额"])
    print(left)
    print(round(data2.iloc[15, -3] / 10000, 0))
    if left == round(data2.iloc[15, -3] / 10000, 0):
        return "社保基金存款与A1411校验成功\n"
    else:
        return "社保基金存款与A1411校验错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 财政性存款 ----

def check_gkck(data, data1):
    """校验国库存款本金（科目余额表 VS IRP001）"""
    if data.iloc[17, 2] == quzhi(data1, "200602"):
        return "国库存款本金数据校验成功\n"
    else:
        return "国库存款本金数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_gkcklx(data, data1):
    """校验国库存款应付利息（科目余额表 VS IRP001）"""
    if data.iloc[17, 3] == quzhi(data1, "22310602"):
        return "国库存款应付利息数据校验成功\n"
    else:
        return "国库存款应付利息数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_dlczck(data, data1):
    """校验代理财政性存款（科目余额表 VS IRP001）"""
    if data.iloc[18, 2] == quzhi(data1, "2007"):
        return "代理财政性存款数据校验成功\n"
    else:
        return "代理财政性存款数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_dlczcklx(data, data1):
    """校验代理财政性存款应付利息（科目余额表 VS IRP001）"""
    if data.iloc[18, 3] == quzhi(data1, "223107"):
        return "代理财政性存款应付利息数据校验成功\n"
    else:
        return "代理财政性存款应付利息数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_czxckA(data1, data2):
    """校验财政性存款与 A1411 的一致性"""
    if data1.at["12MIK", "本期余额"] == round(data2.iloc[16, -3] / 10000, 0):
        return "财政性存款与A1411校验成功\n"
    else:
        return "财政性存款与A1411校验错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 非存款类金融机构存款 ----

def check_jjck(data, data1):
    """校验非存款类金融机构存款（科目余额表 VS IRP001）"""
    if data.iloc[19, 2] == round(quzhi(data1, "201705") + quzhi(data1, "201799"), 2):
        return "非存款类金融机构存款数据校验成功\n"
    else:
        return "非存款类金融机构存款数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_jjcklx(data, data1):
    """校验非存款类金融机构存款利息（科目余额表 VS IRP001）"""
    if data.iloc[19, 3] == round(quzhi(data1, "22311005") + quzhi(data1, "22311099"), 2):
        return "非存款类金融机构存款利息数据校验成功\n"
    else:
        return "非存款类金融机构存款利息数据错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 境外存款 ----

def check_jwck(data):
    """校验境外存款本金是否为0"""
    if data.iloc[20, 2] == 0.:
        return "该机构境外存款本金数据校验成功\n"
    else:
        return "该机构境外存款本金数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_jwcklx(data):
    """校验境外存款利息是否为0"""
    if data.iloc[20, 3] == 0.:
        return "该机构境外存款利息数据校验成功\n"
    else:
        return "该机构境外存款利息数据错误!!!!!!!!!!!!!!!!!!！\n"


# ---- 高管存款 ----

def check_ggck(data, data1):
    """校验高管存款（附件5明细 VS IRP001）"""
    if data1["余额"].iloc[:-1].fillna(0).sum() == data1['余额'].iloc[-1]:
        if round(data1['余额'].iloc[-1], 2) == round(data.iloc[24, 2], 2):
            return "高管存款数据校验成功\n"
        else:
            return "高管存款数据核对错误!!!!!!!!!!!!!!!!!!！\n"
    else:
        return "高管存款内部数据错误!!!!!!!!!!!!!!!!!!！\n"


def check_ggckbx(data, data1):
    """校验高管存款本息（附件5明细 VS IRP001）"""
    if data1["余额"].iloc[:-1].fillna(0).sum() == data1['余额'].iloc[-1]:
        if round(data1["余额"].iloc[:-1].fillna(0).sum(), 2) == round(
                data.iloc[24, 4], 2):
            return "高管存款本息数据校验成功\n"
        else:
            return "高管存款本息数据核对错误!!!!!!!!!!!!!!!!!!！\n"
    else:
        return "高管存款内部数据错误!!!!!!!!!!!!!!!!!!！\n"


# ============================================================
# 主校验函数
# ============================================================

def check_nsh(path):
    """
    对指定工作目录下的所有申报文件执行全面校验。

    读取文件：
      - 附件5：保费基数核对表（获取调整值 ce1~ce4）
      - 科目余额表
      - IRP001 投保机构被保险存款申报表
      - 附件2：住房公积金明细表
      - 附件3：社保基金存款明细表
      - 附件4：高管人员存款统计表
      - A1411 资产负债表

    参数:
        path (str): 工作目录路径

    返回:
        str: 所有校验结果拼接的字符串
    """
    # 从附件5读取调整值
    ce1 = pd.read_excel(_find_excel(path + '/附件5：保费基数核对表'))[5:].iloc[0, -1]
    ce2 = pd.read_excel(_find_excel(path + '/附件5：保费基数核对表'))[5:].iloc[1, -1]
    ce3 = pd.read_excel(_find_excel(path + '/附件5：保费基数核对表'))[5:].iloc[2, -1]
    ce4 = pd.read_excel(_find_excel(path + '/附件5：保费基数核对表'))[5:].iloc[3, -1]

    # 科目余额表
    kmye = pd.read_excel(_find_excel(path + "/科目余额表"), index_col=0)
    # IRP001 申报表
    data = pd.read_excel(_find_excel(path + "/IRP001投保机构被保险存款申报表"))
    # 附件2：住房公积金
    gjj = pd.read_excel(_find_excel(path + "/附件2：住房公积金明细表"))[3:]
    gjj.columns = ['1', '2', '3', '4', '5', '余额', '7', '8']
    # 附件3：社保基金
    sbjj = pd.read_excel(_find_excel(path + "/附件3：社保基金存款明细表"))[3:]
    sbjj.columns = ['1', '2', '3', '4', '5', '余额', '7', '8']
    # 附件4：高管人员存款
    ggck = pd.read_excel(_find_excel(path + "/附件4：高管人员存款统计表"))[3:]
    ggck.columns = ['1', '2', '3', '4', '5', '6', '余额', '8', '9']
    # A1411 资产负债表
    zcfz = pd.read_excel(_find_excel(path + "/A1411"), index_col=0, header=2)
    # 附件5：保费基数核对表（再读一次用于调整值）
    ce = pd.read_excel(_find_excel(path + '/附件5：保费基数核对表'))[3:]

    massage = (
        check_grck(data, kmye, ce1)
        + check_grcklx(data, kmye, ce2)
        + check_dwck(data, kmye, ce1)
        + check_dwcklx(data, kmye, ce2)
        + check_gjj(data, gjj)
        + check_sbjj(data, sbjj)
        + check_gkck(data, kmye)
        + check_gkcklx(data, kmye)
        + check_dlczck(data, kmye)
        + check_dlczcklx(data, kmye)
        + check_jjck(data, kmye)
        + check_jjcklx(data, kmye)
        + check_jwck(data)
        + check_jwcklx(data)
        + check_ggck(data, ggck)
        + check_ggckbx(data, ggck)
        + check_grckA(zcfz, ce3, data)
        + check_dwckA(zcfz, ce4, data)
        + check_gjjA(zcfz, data)
        + check_sbjjA(zcfz, data)
        + check_czxckA(zcfz, data)
        + check_lxhA(zcfz, data)
    )
    return massage
