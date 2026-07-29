# -*- coding: utf-8 -*-
"""
=============================================================================
银行机构类型配置模块
=============================================================================

本模块管理不同银行机构类型的模板文件配置。
每种银行类型对应两套模板文件：
  1. 存款归属表模板（deposit_template）
  2. 科目余额核对表模板（acct_check_template）

【扩展说明】新增银行类型只需在 BANK_TYPES 字典中添加条目：
    "银行名称": {
        "deposit_template": "templates/存款归属表_XXX（模板）.xlsx",
        "acct_check_template": "templates/科目余额核对表_XXX(模板).xlsx",
    },

然后：
  1. 将对应的模板文件放入 templates/ 目录
  2. 在 build.spec 的 datas 中添加模板文件路径
  3. 重新打包即可

注意：模板文件打包后嵌入 exe，无需随程序分发
=============================================================================
"""

# ============================================================
# 银行类型配置字典
# ============================================================
# 键：银行类型名称（字符串）
# 值：字典，包含两个模板文件相对路径
# ============================================================

BANK_TYPES = {
    "农村商业银行": {
        "deposit_template": "templates/deposit_template.xlsx",
        "acct_check_template": "templates/acct_check_template.xlsx",
    },
    # === 未来扩展示例（取消注释并放入对应模板文件即可） ===
    # "城市商业银行": {
    #     "deposit_template": "templates/deposit_template_city.xlsx",
    #     "acct_check_template": "templates/acct_check_template_city.xlsx",
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
        template_key (str): 模板类型键名，如 "deposit_template" 或 "acct_check_template"

    返回:
        str: 模板文件相对路径；若未找到则返回空字符串

    注意:
        如果指定的 bank_type 不存在，会自动回退使用 DEFAULT_BANK 的配置
    """
    bank = BANK_TYPES.get(bank_type)
    if not bank:
        bank = BANK_TYPES[DEFAULT_BANK]
    return bank.get(template_key, "")
