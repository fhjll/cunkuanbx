# -*- coding: utf-8 -*-
"""
=============================================================================
.xls 模板 → .xlsx 模板 转换工具（开发用，不参与打包运行）
=============================================================================

用途:
  监管下发的表样（如《保费基数核对表（模板）.xls》）是旧式 BIFF 格式，
  而本工具链统一使用 openpyxl 读写 .xlsx。本脚本把 .xls 表样一次性
  转换为等价的 .xlsx 模板，供程序作为内嵌模板使用。

转换内容:
  - 单元格文本/数值
  - 单元格公式（含共享公式 SHRFMLA 展开为逐格显式公式）
  - 合并单元格
  - 列宽 / 行高
  - 字体（名称/字号/加粗/斜体/颜色）、对齐、边框、填充、数字格式
  - 批注（单元格注释，用于说明调整值口径）

用法:
  python tools/convert_xls_template.py 源模板.xls 目标模板.xlsx

说明:
  - 只支持 BIFF8（Excel 97-2003）格式的 .xls；若文件实为 OOXML，
    直接改扩展名为 .xlsx 即可，无需本脚本。
  - 批注文字为 BIFF 富文本，脚本按注释内容关键词映射到固定单元格
    （见 _NOTES_BY_KEYWORD），未命中的注释会被丢弃并打印提示。
=============================================================================
"""
import os
import re
import struct
import sys

import xlrd
from xlrd import compdoc
from xlrd.formula import decompile_formula, colname, FMLA_TYPE_CELL, FMLA_TYPE_SHARED
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# BIFF 记录常量
# ============================================================
REC_FORMULA = 0x0006
REC_BOF = 0x0809
REC_CONTINUE = 0x003C
REC_NOTE = 0x001C
REC_SHRFMLA = 0x04BC


# ============================================================
# 批注文字 → 单元格映射
# ============================================================
# BIFF 不直接给出「注释属于哪个单元格」，且 OBJ/TXO 记录顺序不保证与
# NOTE 记录一致。这里按批注内容关键词映射到模板中的固定单元格。
# 关键词命中即写入，未命中的批注会被忽略（打印提示）。
_NOTES_BY_KEYWORD = [
    ('个体工商户由单位调整至个人部分的本金', 'Q7'),
    ('个体工商户由单位调整至个人部分的应付利息', 'Q8'),
    ('久悬户归属于个人存款的部分', 'Q9'),
    ('久悬户归属于单位存款的部分', 'Q10'),
    ('每年度末重新调整', ['H23', 'H37']),
]


# ============================================================
# 边框 / 填充 映射
# ============================================================
_BORDER_STYLES = {
    0: None,
    1: 'thin',
    2: 'medium',
    3: 'dashed',
    4: 'dotted',
    5: 'thick',
    6: 'double',
    7: 'hair',
    8: 'mediumDashed',
    9: 'dashDot',
    10: 'mediumDashDot',
    11: 'dashDotDot',
    12: 'mediumDashDotDot',
    13: 'slantDashDot',
}

_FILL_PATTERNS = {
    0: None,
    1: 'solid',
    2: 'mediumGray',
    3: 'darkGray',
    4: 'lightGray',
    5: 'darkHorizontal',
    6: 'darkVertical',
    7: 'darkDown',
    8: 'darkUp',
    9: 'darkGrid',
    10: 'darkTrellis',
    11: 'lightHorizontal',
    12: 'lightVertical',
    13: 'lightDown',
    14: 'lightUp',
    15: 'lightGrid',
    16: 'lightTrellis',
    17: 'gray125',
    18: 'gray0625',
}

_H_ALIGN = {0: None, 1: 'left', 2: 'center', 3: 'right', 4: 'fill',
            5: 'justify', 6: 'centerContinuous', 7: 'distributed'}
_V_ALIGN = {0: 'top', 1: 'center', 2: 'bottom', 3: 'justify', 4: 'distributed'}

# xlrd 内置数字格式（format_key < 50 时 bk.format_map 可能没有）
_BUILTIN_FORMATS = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
}


# ============================================================
# 颜色 / 样式转换
# ============================================================

def _rgb_str(book, colour_index):
    """把 xlrd 调色板索引转成 openpyxl 的 6 位 RGB 字符串；无颜色返回 None。"""
    if colour_index in (None, 0x40, 0x41, 0x7FFF, 64, 65):
        return None  # 自动/默认色
    rgb = book.colour_map.get(colour_index)
    if not rgb:
        return None
    return '%02X%02X%02X' % rgb


def _convert_font(book, xf):
    """XF → openpyxl Font"""
    f = book.font_list[xf.font_index]
    kwargs = {
        'name': f.name,
        'size': round(f.height / 20.0, 2) if f.height else 11,
        'bold': bool(f.bold or f.weight > 400),
        'italic': bool(f.italic),
    }
    color = _rgb_str(book, f.colour_index)
    if color:
        kwargs['color'] = color
    return Font(**kwargs)


def _convert_alignment(xf):
    """XF → openpyxl Alignment"""
    a = xf.alignment
    return Alignment(
        horizontal=_H_ALIGN.get(a.hor_align),
        vertical=_V_ALIGN.get(a.vert_align),
        wrap_text=bool(a.text_wrapped),
    )


def _convert_border(book, xf):
    """XF → openpyxl Border"""
    def side(style_id, colour_index):
        style = _BORDER_STYLES.get(style_id)
        if not style:
            return Side()
        colour = _rgb_str(book, colour_index)
        return Side(style=style, color=colour) if colour else Side(style=style)

    b = xf.border
    return Border(left=side(b.left_line_style, b.left_colour_index),
                  right=side(b.right_line_style, b.right_colour_index),
                  top=side(b.top_line_style, b.top_colour_index),
                  bottom=side(b.bottom_line_style, b.bottom_colour_index))


def _convert_fill(book, xf):
    """XF → openpyxl PatternFill"""
    bg = xf.background
    pattern = _FILL_PATTERNS.get(bg.fill_pattern)
    if not pattern or pattern == 'solid':
        pattern = 'solid' if bg.fill_pattern else None
    if not pattern:
        return PatternFill(fill_type=None)
    colour = _rgb_str(book, bg.pattern_colour_index)
    if colour:
        return PatternFill(fill_type=pattern, start_color=colour, end_color=colour)
    return PatternFill(fill_type=pattern)


def _convert_number_format(book, xf):
    """XF → 数字格式字符串"""
    key = xf.format_key
    fmt = book.format_map.get(key)
    fmt = getattr(fmt, 'format_str', fmt)   # xlrd 存的是 Format 对象
    return fmt or _BUILTIN_FORMATS.get(key) or 'General'


# ============================================================
# BIFF 公式解析
# ============================================================

class _FormulaReader:
    """从 BIFF 流中读取普通公式与共享公式（SHRFMLA）。"""

    def __init__(self, book, stream):
        self.book = book
        self.stream = stream
        self.cell_formulas = {}    # {(sheet_idx, row0, col0): 公式文本}
        self.shared_formulas = []  # [(rwfirst0, rwlast0, colfirst0, collast0, text)]
        self.notes_text = []       # 批注文字（按流中出现顺序）
        self._shared_stubs = []    # [(sheet_idx, row0, col0)] 待解析的共享公式占位格
        self._parse()
        self._resolve_shared_stubs()

    # --------------------------------------------------
    def _parse(self):
        pos = 0
        sheet = -1
        while pos < len(self.stream) - 4:
            rec, ln = struct.unpack_from('<HH', self.stream, pos)
            body = self.stream[pos + 4:pos + 4 + ln]
            if rec == REC_BOF and ln >= 4:
                # 只统计工作表 BOF（dt=0x0010）——工作簿全局 BOF 不计入
                dt = struct.unpack_from('<H', body, 2)[0]
                if dt == 0x0010:
                    sheet += 1
            elif rec == REC_SHRFMLA:
                self._read_shrfmla(body)
            elif rec == REC_FORMULA:
                self._read_formula(sheet, body)
            elif rec == REC_CONTINUE and ln > 0:
                text = self._decode_note_text(body)
                if text:
                    self.notes_text.append(text)
            pos += 4 + ln

    # --------------------------------------------------
    def _read_shrfmla(self, body):
        """ShrFmla: ref(rwFirst2 rwLast2 colFirst1 colLast1) + 2字节保留 + cce2 + rgce"""
        rw_first, rw_last = struct.unpack_from('<HH', body, 0)
        col_first, col_last = body[4], body[5]
        cce = struct.unpack_from('<H', body, 8)[0]
        rgce = body[10:10 + cce]
        text = self._decompile(rgce, rw_first, col_first, FMLA_TYPE_SHARED)
        if not text:
            return
        self.shared_formulas.append((rw_first, rw_last, col_first, col_last, text))

    def _read_formula(self, sheet, body):
        """Formula: rw2 col2 ixfe2 num8 grbit2 chn4 cce2 rgce"""
        rw, col = struct.unpack_from('<HH', body, 0)
        cce = struct.unpack_from('<H', body, 20)[0]
        rgce = body[22:22 + cce]
        text = self._decompile(rgce, rw, col, FMLA_TYPE_CELL)
        if 'SHARED FMLA' in text:
            # 共享公式占位格：母公式需在 SHRFMLA 全部解析后再展开
            self._shared_stubs.append((sheet, rw, col))
            return
        if text:
            self.cell_formulas[(sheet, rw, col)] = text

    def _resolve_shared_stubs(self):
        """共享公式占位格：用所在共享区域的母公式按偏移量展开。"""
        for sheet, rw, col in self._shared_stubs:
            text = ''
            for rw_first, rw_last, col_first, col_last, master_text in self.shared_formulas:
                if rw_first <= rw <= rw_last and col_first <= col <= col_last:
                    text = _shift_formula(master_text, rw - rw_first, col - col_first)
                    break
            if text:
                self.cell_formulas[(sheet, rw, col)] = text

    def _decompile(self, rgce, rw, col, fmla_type):
        try:
            res = decompile_formula(self.book, rgce, len(rgce),
                                    fmla_type, rw, col, 0, 0, 0)
            text = res[0] if isinstance(res, tuple) else res
        except Exception:
            return ''
        return _normalize_formula(text)

    @staticmethod
    def _decode_note_text(body):
        """批注 CONTINUE 记录：首字节为长度，其后为 UTF-16LE 文本。"""
        if len(body) < 4:
            return ''
        for offset in (1, 0):
            try:
                text = body[offset:].decode('utf-16-le', errors='ignore')
            except Exception:
                continue
            text = ''.join(ch for ch in text if ch.isprintable() or ch == '\n')
            text = text.strip()
            # 只保留看起来是中文说明的文本
            if text and re.search(r'[\u4e00-\u9fff]', text) and len(text) > 4:
                return text
        return ''


def _normalize_formula(text):
    """去掉单格区域的 ':X9' 尾巴，如 'F7:F7' → 'F7'。"""
    if not text:
        return text
    return re.sub(r'(\$?[A-Z]{1,3}\$?\d+):\1', r'\1', text)


def _shift_formula(text, drow, dcol):
    """把公式中的相对引用平移 (drow, dcol)；绝对引用（带 $）保持不变。"""
    if not text or (drow == 0 and dcol == 0):
        return text

    def repl(m):
        dollars_row, col_letters, dollars_col, row_num = m.groups()
        if dollars_col:            # 列绝对 → 列不动
            new_col = col_letters
        else:
            idx = 0
            for ch in col_letters:
                idx = idx * 26 + (ord(ch) - 64)
            new_col = ''
            idx += dcol
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                new_col = chr(65 + rem) + new_col
        new_row = row_num if dollars_row else str(int(row_num) + drow)
        return f'{dollars_row}{new_col}{dollars_col}{new_row}'

    # 匹配 [sheet!] 之后的单元格引用
    return re.sub(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', repl, text)


# ============================================================
# 转换主流程
# ============================================================

def convert(xls_path, xlsx_path):
    """
    把 BIFF8 的 .xls 模板转换为等价的 .xlsx 模板。

    参数:
        xls_path (str): 源 .xls 模板路径
        xlsx_path (str): 目标 .xlsx 模板路径

    返回:
        list[str]: 转换日志
    """
    msgs = []
    book = xlrd.open_workbook(xls_path, formatting_info=True)
    stream = compdoc.CompDoc(open(xls_path, 'rb').read()).get_named_stream('Workbook')
    fr = _FormulaReader(book, stream)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet_idx, sh in enumerate(book.sheets()):
        ws = wb.create_sheet(sh.name)
        formula_count = 0

        # --- 单元格 ---
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                value = sh.cell_value(r, c)
                formula = fr.cell_formulas.get((sheet_idx, r, c))
                xf_index = sh.cell_xf_index(r, c)
                xf = book.xf_list[xf_index]

                if formula:
                    # 公式单元格：写公式（数值由程序计算后固化，此处保留公式便于核对）
                    cell = ws.cell(row=r + 1, column=c + 1)
                    cell.value = '=' + formula
                    formula_count += 1
                elif value not in ('', None):
                    ws.cell(row=r + 1, column=c + 1).value = value

                cell = ws.cell(row=r + 1, column=c + 1)
                cell.font = _convert_font(book, xf)
                cell.alignment = _convert_alignment(xf)
                cell.border = _convert_border(book, xf)
                cell.fill = _convert_fill(book, xf)
                cell.number_format = _convert_number_format(book, xf)

        # --- 合并单元格 ---
        for rlo, rhi, clo, chi in sh.merged_cells:
            ws.merge_cells(start_row=rlo + 1, start_column=clo + 1,
                           end_row=rhi, end_column=chi)
        # 合并区域内被 openpyxl 置空的样式：重写左上角以外的边框（保持简单，跳过）

        # --- 列宽 ---
        for c, ci in sh.colinfo_map.items():
            if ci.width:
                ws.column_dimensions[get_column_letter(c + 1)].width = round(ci.width / 256.0, 2)

        # --- 行高 ---
        for r, ri in sh.rowinfo_map.items():
            if ri.height:
                ws.row_dimensions[r + 1].height = round(ri.height / 20.0, 2)

        msgs.append(f"  {sh.name}: {sh.nrows}行 x {sh.ncols}列, "
                    f"公式{formula_count}个, 合并区域{len(sh.merged_cells)}个")

    # --- 批注（按内容关键词映射到固定单元格） ---
    note_cells = {}
    for text in fr.notes_text:
        for keyword, target in _NOTES_BY_KEYWORD:
            if keyword in text:
                targets = target if isinstance(target, list) else [target]
                for t in targets:
                    if t not in note_cells:   # 同一关键词只填一次（第二个"每年度末"另填）
                        note_cells[t] = text
                        break
    # "每年度末重新调整" 需要写到两个单元格
    for keyword, target in _NOTES_BY_KEYWORD:
        if isinstance(target, list):
            hit = next((t for t in fr.notes_text if keyword in t), None)
            if hit:
                for t in target:
                    note_cells[t] = hit
    for addr, text in note_cells.items():
        sheet_name = '保费基数核对表'
        if sheet_name in wb.sheetnames:
            wb[sheet_name][addr].comment = Comment(text, '模板')
    msgs.append(f"  批注: 识别{len(fr.notes_text)}条, 写入{len(note_cells)}个单元格")

    wb.save(xlsx_path)
    msgs.append(f"  已保存: {xlsx_path}")
    return msgs


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        print("用法: python tools/convert_xls_template.py 源模板.xls 目标模板.xlsx")
        return 1
    for m in convert(sys.argv[1], sys.argv[2]):
        print(m)
    return 0


if __name__ == '__main__':
    sys.exit(main())
