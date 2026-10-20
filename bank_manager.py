# -*- coding: utf-8 -*-
"""
=============================================================================
银行类型 — 模板绑定管理模块
=============================================================================

设计目标：
  · 每种银行类型（农村商业银行 / 城市商业银行 / ...）绑定一个
    《保费基数核对表》模板文件；
  · 绑定关系持久化保存，可随时把某个银行类型改绑为其他模板文件；
  · 支持新增银行类型、导入（新增）模板文件；
  · 与界面解耦：gui.py / bank_selector.py 只调用本模块，不直接读写配置。

配置存放位置（按优先级）：
  1. 环境变量 CUNBAO_CONFIG_DIR 指定的目录（便于测试/多套配置）
  2. 程序目录（便携模式：exe 或源码所在目录，可写时）
  3. 用户配置目录：
       macOS   ~/Library/Application Support/存保审核工具
       Windows %APPDATA%\\存保审核工具
       Linux   ~/.config/存保审核工具

配置文件：<配置目录>/bank_bindings.json
  {
    "version": 1,
    "last_bank": "农村商业银行",
    "banks": [
      {"name": "农村商业银行",
       "template": "templates/保费基数核对表（模板）.xlsx",
       "default_template": "templates/保费基数核对表（模板）.xlsx"}
    ],
    "templates": ["templates/保费基数核对表（模板）.xlsx"]
  }

  · template          = 当前绑定的模板（相对程序目录则存相对路径，否则存绝对路径）
  · default_template  = 该银行类型的内置默认模板（「恢复默认」用）
  · templates         = 已登记的模板库（含默认模板与导入的模板）

说明：
  - 绑定模板文件本身不会被程序修改，工具每次生成时读取并另存输出文件；
  - 导入的模板会复制到 <配置目录>/模板库/ 下，便于长期保存与随程序携带。
=============================================================================
"""
import json
import os
import shutil
import sys
import time

import openpyxl

from bank_config import DEFAULT_BANK, _program_dir, get_template, resolve_template

# 配置文件名与目录名
CONFIG_NAME = 'bank_bindings.json'
CONFIG_VERSION = 1
LIBRARY_DIR_NAME = '模板库'

# 模板主表 / 校验页工作表名（与 base_checker 保持一致）
MAIN_SHEET_NAME = '保费基数核对表'
CHECK_SHEET_NAME = '校验'

# 判断"是不是保费基数核对表模板"的表头关键词
_HEADER_KEYWORDS = ('存款归属', '会计科目', '被保险存款申报表', '差额说明')


# ============================================================
# 配置目录
# ============================================================

def user_config_dir():
    """返回当前平台的用户配置目录（<base>/存保审核工具）。"""
    if sys.platform.startswith('darwin'):
        base = os.path.join(os.path.expanduser('~'), 'Library',
                            'Application Support')
    elif os.name == 'nt':
        base = (os.environ.get('APPDATA')
                or os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming'))
    else:
        base = (os.environ.get('XDG_CONFIG_HOME')
                or os.path.join(os.path.expanduser('~'), '.config'))
    return os.path.join(base, '存保审核工具')


def _is_writable_dir(path):
    """目录是否可创建/可写（用于判断程序目录能否存配置）。"""
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, '.write_probe')
        with open(probe, 'w', encoding='utf-8') as f:
            f.write('')
        os.remove(probe)
        return True
    except OSError:
        return False


def config_dir():
    """
    返回绑定配置目录。

    优先级：CUNBAO_CONFIG_DIR → 程序目录（可写时，便携）→ 用户配置目录。
    """
    env = os.environ.get('CUNBAO_CONFIG_DIR')
    if env:
        return os.path.abspath(os.path.expanduser(env))
    prog = _program_dir()
    if _is_writable_dir(prog):
        return prog
    return user_config_dir()


# ============================================================
# 路径工具
# ============================================================

def _norm_abs(path):
    return os.path.abspath(os.path.expanduser(str(path)))


def _same_path(a, b):
    """比较两个路径是否指向同一文件（忽略大小写/分隔符差异）。"""
    if not a or not b:
        return False
    try:
        return os.path.normcase(_norm_abs(a)) == os.path.normcase(_norm_abs(b))
    except Exception:
        return False


def _resolve_stored(raw):
    """
    把配置里保存的模板路径解析为实际存在的绝对路径。

    · 绝对路径：原样返回；
    · 相对路径：先按程序目录解析（templates/...），
      再通过 bank_config.resolve_template 兼容打包后的内置模板。
    """
    if not raw:
        return ''
    raw = str(raw).strip()
    if not raw:
        return ''
    if os.path.isabs(raw):
        return _norm_abs(raw)
    cand = os.path.join(_program_dir(), raw)
    if os.path.isfile(cand):
        return cand
    try:
        path, _ = resolve_template(raw, None)
        return path
    except Exception:
        return cand


def _store_path(path):
    """
    把用户选择的模板路径转成配置里的保存形式：
    在程序目录内的文件存相对路径（便于整目录搬迁），否则存绝对路径。
    """
    path = _norm_abs(path)
    prog = _norm_abs(_program_dir())
    try:
        rel = os.path.relpath(path, prog)
    except ValueError:
        return path
    if rel and not rel.startswith('..') and not os.path.isabs(rel):
        return rel.replace(os.sep, '/')
    return path


def _sanitize_name(name):
    """姓名/文件名词干清洗：去掉路径分隔符等非法字符。"""
    name = str(name or '').strip()
    for ch in '\\/:*?"<>|\r\n\t':
        name = name.replace(ch, '_')
    return name.strip(' .')


# ============================================================
# 模板文件校验
# ============================================================

def validate_template(path):
    """
    校验模板文件是否可作为《保费基数核对表》模板使用。

    参数:
        path (str): 模板文件路径

    返回:
        tuple: (是否通过, 说明文字)
    """
    if not path:
        return False, '未选择模板文件'
    path = _norm_abs(path)
    if not os.path.isfile(path):
        return False, f'文件不存在：{path}'
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        return False, ('模板必须是 .xlsx / .xlsm 文件；.xls 请先用 Excel/WPS '
                       '另存为 .xlsx')
    try:
        wb = openpyxl.load_workbook(path, read_only=False)
    except Exception as e:
        return False, f'模板无法读取：{e}'

    try:
        names = list(wb.sheetnames)
        ws = wb[MAIN_SHEET_NAME] if MAIN_SHEET_NAME in names else wb.active
        hits = set()
        max_row = min(ws.max_row, 12)
        max_col = min(ws.max_column, 30)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    for kw in _HEADER_KEYWORDS:
                        if kw in v:
                            hits.add(kw)
        wb.close()
    except Exception as e:
        return False, f'模板结构检查失败：{e}'

    if MAIN_SHEET_NAME not in names:
        return False, (f'模板缺少主表工作表「{MAIN_SHEET_NAME}」'
                       f'（现有工作表：{"、".join(names)}）')
    if len(hits) < 3:
        return False, ('选中的文件不像《保费基数核对表》模板（未找到'
                       '“存款归属 / 会计科目 / 被保险存款申报表 / 差额说明”'
                       '表头），请确认文件是否正确')
    if CHECK_SHEET_NAME not in names:
        return True, ('模板可用，但缺少「校验」工作表——生成时不会输出'
                      '校验页差额（主表仍可正常生成）')
    return True, '模板可用'


# ============================================================
# 绑定管理器
# ============================================================

class BankManager:
    """
    银行类型与模板绑定关系管理器。

    典型用法::

        mgr = BankManager()
        mgr.bank_names()                    # 所有银行类型
        mgr.template_info('农村商业银行')   # 绑定信息
        mgr.bind_template('农村商业银行', '/path/to/模板.xlsx')
        mgr.add_bank('城市商业银行')
        mgr.import_template('/path/to/新模板.xlsx')
    """

    def __init__(self, config_dir_path=None):
        """
        参数:
            config_dir_path (str, optional): 自定义配置目录（默认见 config_dir()）
        """
        self.config_dir = _norm_abs(config_dir_path or config_dir())
        self.config_path = os.path.join(self.config_dir, CONFIG_NAME)
        self.library_dir = os.path.join(self.config_dir, LIBRARY_DIR_NAME)
        self._data = None
        self.load()

    # ---------------- 读写配置 ----------------

    def _default_data(self):
        """生成默认绑定配置：内置的「农村商业银行」+ 默认模板。"""
        rel = get_template(DEFAULT_BANK, 'base_check_template')
        return {
            'version': CONFIG_VERSION,
            'last_bank': DEFAULT_BANK,
            'banks': [{'name': DEFAULT_BANK,
                       'template': rel,
                       'default_template': rel}],
            'templates': [rel],
        }

    def load(self):
        """从磁盘载入绑定配置；文件不存在时创建默认配置。"""
        data = None
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, ValueError):
            data = None

        if not isinstance(data, dict):
            self._data = self._default_data()
            self.save()
            return

        # 结构归一化，兼容手工编辑过的配置
        data.setdefault('version', CONFIG_VERSION)
        data.setdefault('banks', [])
        data.setdefault('templates', [])
        if not isinstance(data['banks'], list):
            data['banks'] = []
        if not isinstance(data['templates'], list):
            data['templates'] = []

        default_rel = get_template(DEFAULT_BANK, 'base_check_template')
        cleaned = []
        seen = set()
        for b in data['banks']:
            if not isinstance(b, dict):
                continue
            name = str(b.get('name') or '').strip()
            if not name or name in seen:
                continue
            seen.add(name)
            default_tpl = str(b.get('default_template') or default_rel)
            tpl = str(b.get('template') or default_tpl)
            cleaned.append({'name': name,
                            'template': tpl,
                            'default_template': default_tpl})
        # 配置里一个银行都没有时（首次/配置损坏）补回内置默认银行；
        # 用户主动删除内置银行后不再自动加回
        if not cleaned:
            cleaned.insert(0, {'name': DEFAULT_BANK,
                               'template': default_rel,
                               'default_template': default_rel})
        data['banks'] = cleaned

        # 模板库去重（保留首次出现顺序）
        lib, lib_seen = [], set()
        for t in data['templates']:
            t = str(t or '').strip()
            if t and t not in lib_seen:
                lib_seen.add(t)
                lib.append(t)
        for b in cleaned:
            if b['default_template'] not in lib_seen:
                lib_seen.add(b['default_template'])
                lib.append(b['default_template'])
        data['templates'] = lib

        names = [b['name'] for b in cleaned]
        if data.get('last_bank') not in names:
            data['last_bank'] = (DEFAULT_BANK if DEFAULT_BANK in names
                                 else names[0])
        self._data = data
        self.save()

    def save(self):
        """把配置写回磁盘（先写临时文件再替换，避免写坏配置）。"""
        if self._data is None:
            return
        try:
            os.makedirs(self.config_dir, exist_ok=True)
            tmp = self.config_path + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.config_path)
        except OSError:
            # 配置写不进去时不阻塞主流程（本次修改仍在本进程内生效）
            pass

    # ---------------- 查询 ----------------

    def bank_names(self):
        """返回所有银行类型名称（按配置顺序）。"""
        return [b['name'] for b in self._data['banks']]

    def get_bank(self, name):
        """返回某个银行类型的绑定记录（dict）；不存在返回 None。"""
        for b in self._data['banks']:
            if b['name'] == name:
                return b
        return None

    @property
    def last_bank(self):
        """上次选择的银行类型（启动时默认选中）。"""
        return self._data.get('last_bank') or self.bank_names()[0]

    def set_last_bank(self, name):
        """记录当前选择的银行类型。"""
        if name in self.bank_names():
            self._data['last_bank'] = name
            self.save()

    def default_template_path(self, bank_name=None):
        """返回某银行类型的内置默认模板的实际路径（可能不存在）。"""
        bank = self.get_bank(bank_name or self.last_bank)
        rel = (bank or {}).get('default_template')
        if not rel:
            rel = get_template(DEFAULT_BANK, 'base_check_template')
        return _resolve_stored(rel)

    def template_info(self, bank_name=None):
        """
        返回某银行类型当前绑定模板的完整信息。

        返回:
            dict: {
              'bank': 银行类型,
              'path': 实际使用的模板绝对路径,
              'raw': 配置中保存的值,
              'exists': 实际可用的模板文件是否存在（缺失时可能已回退默认模板）,
              'bound_exists': 绑定的模板文件本身是否存在,
              'source': '默认模板' / '自定义绑定' / '默认模板（绑定缺失，已回退）',
              'is_default': 是否为内置默认模板,
              'bound_missing': 绑定的文件不存在时，记录原路径,
            }
        """
        name = bank_name or self.last_bank
        bank = self.get_bank(name) or {}
        raw = bank.get('template') or bank.get('default_template') or ''
        path = _resolve_stored(raw)
        bound_exists = bool(path and os.path.isfile(path))
        is_default = _same_path(raw, bank.get('default_template'))
        source = '默认模板' if is_default else '自定义绑定'
        info = {'bank': name, 'path': path, 'raw': raw,
                'exists': bound_exists, 'bound_exists': bound_exists,
                'source': source,
                'is_default': is_default, 'bound_missing': ''}
        if not bound_exists:
            fb = self.default_template_path(name)
            if fb and os.path.isfile(fb) and not _same_path(fb, path):
                info.update(path=fb, exists=True,
                            source='默认模板（绑定缺失，已回退）',
                            is_default=True, bound_missing=path)
        return info

    def template_library(self):
        """
        返回模板库列表（仅登记信息，不判断存在性）。

        返回:
            list[dict]: [{'path': 绝对路径, 'exists': bool, 'is_default': bool}]
        """
        out = []
        defaults = {b['default_template'] for b in self._data['banks']}
        for raw in self._data['templates']:
            path = _resolve_stored(raw)
            out.append({'path': path, 'exists': bool(path and
                                                    os.path.isfile(path)),
                        'is_default': raw in defaults})
        return out

    # ---------------- 绑定 / 银行管理 ----------------

    def _register_template(self, raw):
        """把模板登记到模板库（去重）。"""
        raw = str(raw or '').strip()
        if not raw:
            return
        if raw not in self._data['templates']:
            self._data['templates'].append(raw)

    def bind_template(self, bank_name, template_path):
        """
        把某银行类型改绑为指定模板文件。

        参数:
            bank_name (str): 银行类型名称
            template_path (str): 模板文件路径

        返回:
            tuple: (是否成功, 说明文字)
        """
        bank = self.get_bank(bank_name)
        if bank is None:
            return False, f'银行类型不存在：{bank_name}'
        ok, msg = validate_template(template_path)
        if not ok:
            return False, msg
        stored = _store_path(template_path)
        bank['template'] = stored
        self._register_template(stored)
        self.save()
        return True, f'「{bank_name}」已绑定模板：{_norm_abs(template_path)}'

    def reset_template(self, bank_name):
        """
        把某银行类型的模板恢复为内置默认模板。

        返回:
            tuple: (是否成功, 说明文字)
        """
        bank = self.get_bank(bank_name)
        if bank is None:
            return False, f'银行类型不存在：{bank_name}'
        bank['template'] = bank.get('default_template') or \
            get_template(DEFAULT_BANK, 'base_check_template')
        self.save()
        return True, f'「{bank_name}」已恢复默认模板'

    def add_bank(self, name, template_path=None):
        """
        新增银行类型。

        参数:
            name (str): 新银行类型名称
            template_path (str, optional): 绑定模板；为空时沿用内置默认模板

        返回:
            tuple: (是否成功, 说明文字)
        """
        name = _sanitize_name(name)
        if not name:
            return False, '银行类型名称不能为空'
        if self.get_bank(name) is not None:
            return False, f'银行类型已存在：{name}'
        default_rel = get_template(DEFAULT_BANK, 'base_check_template')
        stored = default_rel
        if template_path:
            ok, msg = validate_template(template_path)
            if not ok:
                return False, msg
            stored = _store_path(template_path)
        self._data['banks'].append({'name': name,
                                    'template': stored,
                                    'default_template': default_rel})
        self._register_template(stored)
        self.save()
        return True, f'已新增银行类型：{name}'

    def remove_bank(self, name):
        """
        删除银行类型（至少保留一个；内置默认银行允许删除）。

        返回:
            tuple: (是否成功, 说明文字)
        """
        if len(self._data['banks']) <= 1:
            return False, '至少要保留一个银行类型'
        bank = self.get_bank(name)
        if bank is None:
            return False, f'银行类型不存在：{name}'
        self._data['banks'] = [b for b in self._data['banks']
                               if b['name'] != name]
        if self._data.get('last_bank') == name:
            self._data['last_bank'] = self.bank_names()[0]
        self.save()
        return True, f'已删除银行类型：{name}'

    def rename_bank(self, old_name, new_name):
        """
        重命名银行类型。

        返回:
            tuple: (是否成功, 说明文字)
        """
        new_name = _sanitize_name(new_name)
        if not new_name:
            return False, '银行类型名称不能为空'
        bank = self.get_bank(old_name)
        if bank is None:
            return False, f'银行类型不存在：{old_name}'
        if new_name != old_name and self.get_bank(new_name) is not None:
            return False, f'银行类型已存在：{new_name}'
        bank['name'] = new_name
        if self._data.get('last_bank') == old_name:
            self._data['last_bank'] = new_name
        self.save()
        return True, f'已重命名：{old_name} → {new_name}'

    # ---------------- 模板导入 ----------------

    def import_template(self, src_path, name=None):
        """
        导入（新增）模板文件：校验后复制到 <配置目录>/模板库/ 并登记。

        参数:
            src_path (str): 源模板路径
            name (str, optional): 另存文件名（不含扩展名）；默认用源文件名

        返回:
            tuple: (是否成功, 说明文字, 新模板绝对路径)
        """
        ok, msg = validate_template(src_path)
        if not ok:
            return False, msg, ''
        stem = _sanitize_name(name or os.path.splitext(
            os.path.basename(src_path))[0]) or '模板'
        os.makedirs(self.library_dir, exist_ok=True)
        dest = os.path.join(self.library_dir, stem + '.xlsx')
        # 同名文件加序号，避免覆盖
        i = 2
        while os.path.exists(dest) and not _same_path(dest, src_path):
            dest = os.path.join(self.library_dir, f'{stem}_{i}.xlsx')
            i += 1
        try:
            if not _same_path(dest, src_path):
                shutil.copy2(src_path, dest)
        except OSError as e:
            return False, f'复制模板失败：{e}', ''
        stored = _store_path(dest)
        self._register_template(stored)
        self.save()
        return True, f'已导入模板：{dest}', dest
