# -*- coding: utf-8 -*-
"""
=============================================================================
银行机构类型配置模块
=============================================================================

本模块管理不同银行机构类型的模板文件配置。
当前使用合并模板：
  1. 保费基数核对表模板（base_check_template）—— 存款归属核对、
     会计科目核对、被保险存款申报表（IRP001）三块数据合在一张表内
  2. 旧版模板（deposit_template / acct_check_template）—— 仅保留兼容

【扩展说明】新增银行类型只需在 BANK_TYPES 字典中添加条目：
    "银行名称": {
        "base_check_template": "templates/保费基数核对表_XXX（模板）.xlsx",
    },

然后：
  1. 将对应的模板文件放入 templates/ 目录
  2. 打包脚本按目录整体打包 templates/，无需逐个登记
  3. 重新打包即可

注意：模板文件打包后嵌入 exe，无需随程序分发
=============================================================================
"""

import os
import shutil
import sys

from common import get_resource_path


# ============================================================
# 银行类型配置字典
# ============================================================
# 键：银行类型名称（字符串）
# 值：字典，包含模板文件相对路径
# ============================================================

BANK_TYPES = {
    "农村商业银行": {
        "base_check_template": "templates/保费基数核对表（模板）.xlsx",
        # —— 以下为旧版模板（存款归属表 / 科目余额核对表），仅保留兼容 ——
        "deposit_template": "templates/deposit_template.xlsx",
        "acct_check_template": "templates/acct_check_template.xlsx",
    },
    # === 未来扩展示例（取消注释并放入对应模板文件即可） ===
    # "城市商业银行": {
    #     "base_check_template": "templates/保费基数核对表_城市（模板）.xlsx",
    # },
}

# 默认选中的银行类型（GUI 启动时默认选中）
DEFAULT_BANK = "农村商业银行"


def get_bank_list():
    """
    返回所有已配置的银行类型名称列表。

    返回值用于 GUI 下拉框（QComboBox）填充选项。

    返回:
        list[str]: 银行类型名称列表
    """
    return list(BANK_TYPES.keys())


def get_template(bank_type, template_key):
    """
    获取指定银行类型的模板文件相对路径。

    参数:
        bank_type (str): 银行类型名称，如 "农村商业银行"
        template_key (str): 模板类型键名，如 "base_check_template"

    返回:
        str: 模板文件相对路径；若未找到则返回空字符串

    注意:
        如果指定的 bank_type 不存在，会自动回退使用 DEFAULT_BANK 的配置
    """
    bank = BANK_TYPES.get(bank_type)
    if not bank:
        bank = BANK_TYPES[DEFAULT_BANK]
    return bank.get(template_key, "")


# ============================================================
# 模板定位（支持用户自行修改模板）
# ============================================================
# 模板按以下优先级查找，前者优先：
#   1. <工作目录>/模板/<模板文件名>      用户改过的模板（随机构工作目录）
#   2. <程序目录>/templates/<模板文件名> 开发模式或 exe 同级目录
#   3. 内置模板（打包进 exe，运行时解压到临时目录，只读、不可长期修改）
# 通过界面「模板 → 修改模板」把模板复制到工作目录后编辑，即自动生效。
# ============================================================

WORK_TEMPLATE_DIR = '模板'      # 工作目录下存放自定义模板的子目录


def _program_dir():
    """
    返回程序所在目录。

    打包后为 exe 所在目录；开发模式为脚本所在目录（即项目根目录）。
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def template_candidates(rel_path, work_dir=None):
    """
    列出模板的候选路径（按优先级排序）。

    参数:
        rel_path (str): 模板相对路径，如 "templates/保费基数核对表（模板）.xlsx"
        work_dir (str, optional): 工作目录（机构数据目录）

    返回:
        list[tuple]: [(来源标签, 绝对路径), ...]（已去重）
    """
    name = os.path.basename(str(rel_path))
    cands = []
    if work_dir:
        cands.append(('工作目录模板', os.path.join(work_dir, WORK_TEMPLATE_DIR, name)))
    cands.append(('程序目录模板', os.path.join(_program_dir(), 'templates', name)))
    cands.append(('内置模板', get_resource_path(str(rel_path))))

    seen, uniq = set(), []
    for label, path in cands:
        key = os.path.abspath(path)
        if key in seen:
            continue
        seen.add(key)
        uniq.append((label, key))
    return uniq


def resolve_template(rel_path, work_dir=None):
    """
    按优先级返回实际使用的模板路径与来源标签。

    参数:
        rel_path (str): 模板相对路径
        work_dir (str, optional): 工作目录

    返回:
        tuple: (模板绝对路径, 来源标签)；候选都不存在时返回内置模板路径
    """
    cands = template_candidates(rel_path, work_dir)
    for label, path in cands:
        if os.path.exists(path):
            return path, label
    return cands[-1][1], cands[-1][0]


def work_template_path(rel_path, work_dir):
    """
    返回"工作目录模板"的目标路径（用于把模板复制出来编辑）。

    参数:
        rel_path (str): 模板相对路径
        work_dir (str): 工作目录

    返回:
        str: <工作目录>/模板/<模板文件名>
    """
    return os.path.join(str(work_dir), WORK_TEMPLATE_DIR,
                        os.path.basename(str(rel_path)))


def install_template_copy(rel_path, work_dir, source_path=None):
    """
    把模板复制到工作目录（供用户修改），返回复制后的路径。

    参数:
        rel_path (str): 模板相对路径
        work_dir (str): 工作目录
        source_path (str, optional): 源模板路径；默认取当前实际使用的模板

    返回:
        tuple: (目标路径, 源路径, 是否新复制)

    异常:
        OSError: 复制失败（由调用方处理并提示）
    """
    if not source_path:
        source_path, _ = resolve_template(rel_path, work_dir)
    target = work_template_path(rel_path, work_dir)
    if os.path.exists(target) and os.path.abspath(target) == os.path.abspath(source_path):
        return target, source_path, False
    os.makedirs(os.path.dirname(target), exist_ok=True)
    shutil.copy2(source_path, target)
    return target, source_path, True
