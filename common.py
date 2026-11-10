# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险申报表审核工具 - 公共工具模块
=============================================================================

本模块提供所有其他模块共用的工具函数和类：
  1. get_resource_path()       — 获取资源绝对路径（兼容开发模式和 PyInstaller 打包）
  2. extract_embedded_template() — 解压嵌入式模板文件到临时目录
  3. to_float()                — 统一数值转换（支持千分位逗号/中文逗号）
  4. _is_xlsx_format()         — 检测文件是否为 OOXML (xlsx) 格式
  5. detect_excel_kind()       — 按文件内容识别真实格式
                                （xls/xlsx/网页 MHTML/HTML/XML 表格）
  6. ExcelSheet 类             — 统一 Excel 读取器（兼容 .xls/.xlsx 及
                                系统导出的网页格式"伪 Excel"文件）

注意事项：
  - 所有 Excel 文件读取都通过 ExcelSheet 类统一处理，不要在业务代码中直接调用 xlrd 或 openpyxl
  - 本模块不依赖其他自定义模块
=============================================================================
"""
import os
import sys
import tempfile
import shutil
import zipfile
import base64
import quopri
import email
import email.policy
from html.parser import HTMLParser
from xml.etree import ElementTree as ET
import xlrd
import openpyxl


# ============================================================
# 资源路径 — 支持 PyInstaller 打包和开发模式
# ============================================================

def get_resource_path(relative_path):
    """
    获取资源文件的绝对路径。

    在开发模式下，路径相对于本脚本所在目录；
    在 PyInstaller 打包后，路径相对于运行时解压的 _MEIPASS 临时目录。

    参数:
        relative_path (str): 相对路径，例如 "templates/deposit_template.xlsx"

    返回:
        str: 资源的绝对路径

    使用示例:
        >>> tpl = get_resource_path("templates/deposit_template.xlsx")
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包模式：文件被解压到 _MEIPASS 临时目录
        base = sys._MEIPASS
    else:
        # 开发模式：相对于脚本目录
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


def extract_embedded_template(relative_path):
    """
    将嵌入 exe 的模板文件提取到临时目录，返回可用路径。

    模板文件在打包时被嵌入 exe，运行时需解压到 %TEMP% 目录。
    在开发模式下，直接返回原始路径。

    参数:
        relative_path (str): 相对路径，例如 "templates/deposit_template.xlsx"

    返回:
        str: 模板文件的可用路径
    """
    src = get_resource_path(relative_path)
    if os.path.exists(src):
        return src  # 开发模式下直接用
    # 打包模式下文件可能在 _MEIPASS 中但路径格式有差异，兜底返回
    return src


# ============================================================
# 数值转换
# ============================================================

def to_float(val):
    """
    统一数值转换函数。

    将各种可能格式的输入值转换为 float 类型。
    支持：None/空字符串 → 0.0；带千分位逗号或中文逗号的字符串。

    参数:
        val: 输入值，可以是 None、str、int、float 等

    返回:
        float: 转换后的浮点数

    使用示例:
        >>> to_float(None)
        0.0
        >>> to_float("1,234.56")
        1234.56
        >>> to_float("1，234.56")  # 中文逗号
        1234.56
    """
    if val is None or val == '':
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        pass
    try:
        if isinstance(val, str):
            val = val.replace(",", "").replace("，", "").strip()
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def excel_round(value, digits=0):
    """
    Excel 的 ROUND 语义：四舍五入（0.5 远离零）。

    与 Python 内置 round() 的"银行家舍入"不同，模板中
    ROUND((IRP001本金 - 久悬户)/10000, 0) 这类公式依赖 Excel 语义，
    因此统一通过本函数计算。

    参数:
        value (float): 待舍入的数值
        digits (int): 保留小数位数

    返回:
        float: 舍入结果

    使用示例:
        >>> excel_round(2.5)      # Excel ROUND(2.5,0) = 3，Python round(2.5) = 2
        3.0
    """
    from decimal import Decimal, ROUND_HALF_UP
    q = Decimal(1).scaleb(-int(digits or 0))
    return float(Decimal(repr(float(value))).quantize(q, rounding=ROUND_HALF_UP))


def _is_xlsx_format(filepath):
    """
    检测文件是否实际上是 OOXML (xlsx) 格式，无论扩展名是什么。

    检测原理：OOXML 文件本质是 ZIP 压缩包，且内部包含标志文件 xl/workbook.xml。
    某些系统导出的文件虽然扩展名为 .xls，但实际是 OOXML 格式，
    这种情况需要用 openpyxl 而非 xlrd 读取。

    参数:
        filepath (str): 文件路径

    返回:
        bool: True 表示文件是 OOXML 格式，False 表示不是
    """
    if not zipfile.is_zipfile(filepath):
        return False
    try:
        with zipfile.ZipFile(filepath) as z:
            return any("xl/workbook.xml" in n for n in z.namelist())
    except Exception:
        return False


# ============================================================
# 真实文件格式识别 + 网页格式"伪 Excel"解析
# ============================================================
# 部分银行/业务系统导出的"Excel"文件，实际内容是 MHTML（邮件网页）、
# HTML 表格或 Excel 2003 XML（SpreadsheetML），只是扩展名写成了 .xls。
# xlrd/openpyxl 都读不了这种文件（报 "Expected BOF record; found b'MIME-Ver'"）。
# 这里统一按文件头识别真实格式，并把网页/XML 表格解析成可读取的表格数据。
_OLE_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'


def _read_head_bytes(filepath, n=8192):
    """读取文件头部若干字节（失败返回空字节串）。"""
    try:
        with open(str(filepath), 'rb') as f:
            return f.read(n)
    except Exception:
        return b''


def detect_excel_kind(filepath):
    """
    按文件内容识别真实格式（不依赖扩展名）。

    返回:
        str: 'xlsx' / 'xls' / 'mhtml' / 'html' / 'spreadsheetml'
             / 'zip'（普通压缩包）/ 'unknown'
    """
    head = _read_head_bytes(filepath)
    if not head:
        return 'unknown'
    if head[:4] == b'PK\x03\x04':
        return 'xlsx' if _is_xlsx_format(filepath) else 'zip'
    if head[:8] == _OLE_MAGIC:
        return 'xls'
    low = head.lstrip(b'\xef\xbb\xbf \t\r\n').lower()
    # SpreadsheetML(Excel 2003 XML) 要在 HTML 之前判断：
    # 其内容里同样含 <Table> 标签，先按 HTML 判断会误判
    if (b'urn:schemas-microsoft-com:office:spreadsheet' in low
            or (b'<?xml' in low and b'<workbook' in low)):
        return 'spreadsheetml'
    if (b'mime-version' in low[:2048]
            or b'content-type: multipart' in low[:2048]
            or b'content-type: text/html' in low[:2048]):
        return 'mhtml'
    if b'<html' in low or b'<table' in low or b'<body' in low:
        return 'html'
    ext = os.path.splitext(str(filepath))[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        return 'xlsx'
    if ext == '.xls':
        return 'xls'
    return 'unknown'


class _TableHTMLParser(HTMLParser):
    """
    从 HTML 中提取表格。

    支持嵌套表格：每个 <table> 单独收集（外层布局表格里再套数据表格时，
    内层数据表也能被解析到），最终由 _tables_to_sheets 选取数据量表。
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self._stack = []       # 每层 table 的状态

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == 'table':
            grid = []
            self.tables.append(grid)
            self._stack.append({'grid': grid, 'row': None,
                                'in_cell': False, 'buf': []})
        elif tag == 'tr':
            if self._stack:
                row = []
                self._stack[-1]['row'] = row
                self._stack[-1]['grid'].append(row)
        elif tag in ('td', 'th'):
            if self._stack:
                state = self._stack[-1]
                state['in_cell'] = True
                state['buf'] = []
        elif tag == 'br':
            if self._stack and self._stack[-1]['in_cell']:
                self._stack[-1]['buf'].append(' ')

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == 'table':
            if self._stack:
                self._stack.pop()
        elif tag in ('td', 'th'):
            if self._stack:
                state = self._stack[-1]
                if state['in_cell']:
                    text = ''.join(state['buf']).strip()
                    if state['row'] is not None:
                        state['row'].append(text)
                    state['in_cell'] = False
                    state['buf'] = []
        elif tag == 'tr':
            if self._stack:
                self._stack[-1]['row'] = None

    def handle_data(self, data):
        if self._stack and self._stack[-1]['in_cell']:
            self._stack[-1]['buf'].append(data)


class _HtmlSheet:
    """网页表格解析后的工作表（提供 xlrd 风格接口）。"""

    def __init__(self, rows, name='Sheet1'):
        self._rows = rows or []
        self.name = name

    @property
    def nrows(self):
        return len(self._rows)

    @property
    def ncols(self):
        return max((len(r) for r in self._rows), default=0)

    def cell_value(self, r, c):
        try:
            v = self._rows[r][c]
        except Exception:
            return ''
        return v if v is not None else ''


def _decode_html_bytes(raw):
    """字节内容 → 文本（优先 UTF-8，失败再按 GB18030/UTF-16 尝试）。"""
    for enc in ('utf-8', 'gb18030', 'utf-16'):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode('utf-8', errors='replace')


def _tables_to_sheets(tables):
    """
    解析出的表格 → 工作表列表。

    · 过滤 1 行/1 列以下的布局用空表；
    · 按表格数据量（行×列）从大到小排序——网页导出的 Excel 一般主数据表
      最大，这样单表读取器优先拿到真正的数据表，多表读取器也先读到主表。
    """
    candidates = []
    for idx, rows in enumerate(tables):
        if not rows:
            continue
        ncols = max((len(r) for r in rows), default=0)
        if len(rows) < 2 or ncols < 2:
            continue
        candidates.append((len(rows) * ncols, idx, rows))
    candidates.sort(key=lambda x: (-x[0], x[1]))
    return [_HtmlSheet(rows, f'表{idx + 1}')
            for _, idx, rows in candidates]


def _load_pseudo_excel_sheets(filepath, kind):
    """
    读取网页/XML 格式的"伪 Excel"文件，返回 [_HtmlSheet, ...]。

    参数:
        filepath (str): 文件路径
        kind (str): detect_excel_kind() 的返回值（'mhtml'/'html'/'spreadsheetml'）

    返回:
        list[_HtmlSheet]: 解析出的工作表列表（可能为空）
    """
    with open(str(filepath), 'rb') as f:
        raw = f.read()
    if kind == 'mhtml':
        msg = email.message_from_bytes(raw, policy=email.policy.default)
        html_text = ''
        candidate_types = ('text/html', 'application/vnd.ms-excel',
                           'application/msexcel', 'text/plain',
                           'application/xml', 'text/xml')
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                # 有些系统的 MIME 段类型写成 application/vnd.ms-excel，
                # 但内容其实是 HTML 表格：按内容再判断一次
                if ctype not in candidate_types and b'<table' not in payload.lower() \
                        and b'<html' not in payload.lower():
                    continue
                charset = part.get_content_charset() or 'utf-8'
                try:
                    html_text += payload.decode(charset, errors='replace') + '\n'
                except Exception:
                    html_text += _decode_html_bytes(payload) + '\n'
        else:
            payload = msg.get_payload(decode=True)
            decoded = _decode_html_bytes(payload) if payload else ''
            # 单段 MIME：正文可能没有 HTML 标签，从原文里再找表格
            html_text = decoded or _decode_html_bytes(raw)
        text = html_text if '<table' in html_text.lower() else _decode_html_bytes(raw)
        parser = _TableHTMLParser()
        parser.feed(text)
        parser.close()
        return _tables_to_sheets(parser.tables)
    if kind == 'html':
        parser = _TableHTMLParser()
        parser.feed(_decode_html_bytes(raw))
        parser.close()
        return _tables_to_sheets(parser.tables)
    if kind == 'spreadsheetml':
        root = ET.fromstring(raw)
        sheets = []
        for ws_elem in root.iter():
            if not str(ws_elem.tag).endswith('Worksheet'):
                continue
            rows = []
            # Worksheet → Table → Row，需用 iter() 递归找 Row
            for row_elem in ws_elem.iter():
                if not str(row_elem.tag).endswith('Row'):
                    continue
                row = []
                for cell_elem in row_elem:
                    if not str(cell_elem.tag).endswith('Cell'):
                        continue
                    idx = cell_elem.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                    if idx:
                        try:
                            want = int(idx) - 1
                            while len(row) < want:
                                row.append('')
                        except Exception:
                            pass
                    text = ''
                    for child in cell_elem.iter():
                        if str(child.tag).endswith('Data'):
                            text = (child.text or '').strip()
                            break
                    row.append(text)
                while row and row[-1] == '':
                    row.pop()
                rows.append(row)
            name = ws_elem.get('{urn:schemas-microsoft-com:office:spreadsheet}Name')
            sheets.append((name, rows))
        # 过滤空表；按数据量从大到小排序（与网页表格一致）
        out = []
        for name, rows in sheets:
            ncols = max((len(r) for r in rows), default=0)
            if len(rows) < 2 or ncols < 2:
                continue
            out.append((len(rows) * ncols, name, rows))
        out.sort(key=lambda x: -x[0])
        return [_HtmlSheet(rows, name or f'表{len(out)}')
                for _, name, rows in out]
    return []


def _pseudo_excel_error(filepath, detail):
    """生成"伪 Excel/格式不符"的友好错误信息。"""
    name = os.path.basename(str(filepath))
    return ValueError(
        f"无法读取文件「{name}」：{detail}\n\n"
        f"该文件可能不是真正的 Excel，而是系统/网页导出的网页格式"
        f"（MHTML/HTML/XML）或已损坏。\n"
        f"建议：用 Excel/WPS 打开该文件，另存为 .xls 或 .xlsx 后重试。")


# ============================================================
# ExcelSheet 统一读取器
# ============================================================

class ExcelSheet:
    """
    统一 Excel 读取器，兼容 .xls 和 .xlsx 格式。

    提供与 xlrd 兼容的接口（cell_value, nrows），
    自动根据文件扩展名选择底层引擎：
      - .xlsx/.xlsm  → openpyxl
      - .xls         → xlrd

    特殊处理：
      - 若文件扩展名为 .xls 但内部为 OOXML 格式，会自动复制为临时 .xlsx 再读取
      - cell_value(r, c) 统一使用 0-indexed 行列索引

    使用示例:
        >>> ws = ExcelSheet("data.xls")
        >>> print(ws.nrows)          # 总行数
        >>> val = ws.cell_value(3, 2)  # 第4行第3列的值
    """

    def __init__(self, filepath):
        """
        初始化 Excel 读取器。

        按文件内容（而非扩展名）识别真实格式：
          - OOXML(xlsx) → openpyxl
          - 旧式 xls   → xlrd
          - MHTML/HTML/SpreadsheetML（系统导出的"伪 Excel"）→ 网页表格解析器

        参数:
            filepath (str): Excel 文件路径

        异常:
            ValueError: 文件无法按任何已知格式读取（错误信息含文件名与处理建议）
        """
        filepath = str(filepath)
        self._kind = detect_excel_kind(filepath)
        self._temp_copy = None

        if self._kind in ('mhtml', 'html', 'spreadsheetml'):
            try:
                sheets = _load_pseudo_excel_sheets(filepath, self._kind)
            except Exception as e:
                raise _pseudo_excel_error(filepath, str(e)) from e
            if not sheets:
                raise _pseudo_excel_error(filepath, '未能从网页/XML 内容中解析出表格')
            self._ws = sheets[0]          # 读取第一个表格
            self._is_xlsx = False
            return

        try:
            if self._kind in ('xlsx', 'zip'):
                # OOXML：扩展名不是 .xlsx 时复制为临时文件再交给 openpyxl
                load_path = filepath
                if not filepath.lower().endswith((".xlsx", ".xlsm")):
                    fd, temp_copy = tempfile.mkstemp(suffix=".xlsx")
                    os.close(fd)
                    shutil.copy2(filepath, temp_copy)
                    self._temp_copy = temp_copy
                    load_path = temp_copy
                try:
                    wb = openpyxl.load_workbook(load_path, data_only=True)
                except Exception as e:
                    raise _pseudo_excel_error(filepath, str(e)) from e
                self._ws = wb.active           # 读取活动工作表
                self._is_xlsx = True
            else:
                # 旧式 .xls（或未知格式，再兜底试一次）
                try:
                    wb = xlrd.open_workbook(filepath)
                except Exception as e:
                    if _is_xlsx_format(filepath):
                        fd, temp_copy = tempfile.mkstemp(suffix=".xlsx")
                        os.close(fd)
                        shutil.copy2(filepath, temp_copy)
                        self._temp_copy = temp_copy
                        wb2 = openpyxl.load_workbook(temp_copy, data_only=True)
                        self._ws = wb2.active
                        self._is_xlsx = True
                        return
                    raise _pseudo_excel_error(filepath, str(e)) from e
                self._ws = wb.sheet_by_index(0)  # 读取第一个工作表
                self._is_xlsx = False
        finally:
            # 清理临时文件
            if self._temp_copy:
                try:
                    os.remove(self._temp_copy)
                except Exception:
                    pass

    @property
    def nrows(self):
        """返回工作表的总行数（0-indexed）"""
        return self._ws.max_row if self._is_xlsx else self._ws.nrows

    def cell_value(self, r, c):
        """
        获取指定单元格的值。

        参数:
            r (int): 行索引（0-indexed）
            c (int): 列索引（0-indexed）

        返回:
            单元格的值（类型取决于 Excel 中的实际数据类型）
        """
        if self._is_xlsx:
            # openpyxl 使用 1-indexed
            return self._ws.cell(r + 1, c + 1).value
        # xlrd 使用 0-indexed
        return self._ws.cell_value(r, c)
