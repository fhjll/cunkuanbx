# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险申报表审核工具 - 图形界面 v11
=============================================================================

程序入口与主界面。

界面布局（从上到下，单页面可滚动）：
  第〇部分 — 机构类型选择（银行类型下拉框）
  第一部分 — 工作目录（文件夹路径 + 文件匹配状态表格）
  第二部分 — 调整值设置（可自由增删的调整项表格）
  第三部分 — 操作按钮（3个主要操作按钮 + 进度条）
  第四部分 — 运行日志（黑色背景的日志输出区）

3个主要操作（需按顺序执行）：
  1. 生成保费基数核对表 — 合并模板：存款归属核对（A1411/A2411）+
                           会计科目核对（科目余额表）+ IRP001 申报数据
  2. 执行数据校验       — IRP002 校验 + 保费基数核对表 vs IRP001 校验
  3. 生成校验报告       — 输出跨数据源交叉校验的 Excel 报告

工作流程：
  选取工作目录 → 自动扫描匹配文件 → 调整参数（可选）→ 按顺序执行 1→2→3

关键特性：
  - 智能文件匹配：根据文件名关键词自动匹配所需文件
  - 调整项可自由增删：每项配置用途（会计科目核对/存款归属核对）、
    对象（个人存款/单位存款）、字段（本金/应付利息）、方式（增加/扣减）
    与金额；配置保存在工作目录的调整值配置.json，也可从附件5一键读取
  - 支持手动选择文件路径（兜底方案）
  - 所有耗时操作都在后台线程中执行，UI 不卡顿
  - 日志实时输出到黑色背景的日志区域

依赖模块：
  - PyQt5：图形界面框架
  - base_checker.py：保费基数核对表生成引擎（合并模板）
  - validator.py：数据校验引擎
  - bank_config.py：银行类型配置
  - common.py：公共工具
=============================================================================
"""
import sys
import os
import re
import glob
import json
import shutil
import subprocess
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
    QGroupBox, QMessageBox, QProgressBar, QScrollArea, QFrame, QGridLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QAbstractItemView,
    QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from base_checker import (BaseChecker, read_adjustments,
                          build_default_adjustments, normalize_adjustments,
                          merge_standard_pairs,
                          ADJ_SCOPES, ADJ_TARGET_OPTIONS, ADJ_DEP_TARGETS,
                          ADJ_FIELDS, ADJ_MODES, ADJ_SCOPE_DEP, ADJ_DEP_FIELD)
from validator import (
    validate_irp002, validate_base_check_vs_irp001,
    generate_validation_report)
from bank_config import (get_bank_list, get_template, DEFAULT_BANK,
                         WORK_TEMPLATE_DIR, install_template_copy,
                         resolve_template, template_candidates,
                         work_template_path)


# 脚本所在目录（也是默认工作目录）
BASE_DIR = str(Path(__file__).parent)

# 调整项配置文件（保存在工作目录，随机构切换自动载入）
ADJ_CONFIG_NAME = '调整值配置.json'


# ============================================================
# 文件类型定义表
# ============================================================
# 每个条目定义一种待匹配文件：
#   key:      类型标识
#   label:    显示标签
#   patterns: 文件名的匹配关键词列表
#   exts:     允许的扩展名
#   required: 是否为必备文件
#   is_output: 是否为输出文件（工具生成）
#   exclude:  排除关键词（匹配到则跳过）
# ============================================================
FILE_DEFS = [
    {'key': 'a1411',      'label': 'A1411 月报表',
     'patterns': ['a1411', 'A1411'],
     'exts': ['.xls', '.xlsx'], 'required': True},
    {'key': 'a2411',      'label': 'A2411 月报表',
     'patterns': ['a2411', 'A2411'],
     'exts': ['.xls', '.xlsx'], 'required': True},
    {'key': 'base_src',   'label': '科目余额表（源数据）',
     'patterns': ['科目余额表', '科目余额查询报表', '科目余额'],
     'exts': ['.xls', '.xlsx'], 'required': False,
     'exclude': ['核对', '已生成', '保费基数']},
    {'key': 'base_out',   'label': '保费基数核对表（输出）',
     'patterns': ['保费基数核对表_已生成'],
     'exts': ['.xlsx'], 'required': False, 'is_output': True},
    {'key': 'att5',       'label': '附件5：保费基数核对表（银行报送）',
     'patterns': ['附件5', '保费基数核对表'],
     'exts': ['.xls', '.xlsx'], 'required': False,
     'exclude': ['已生成', '模板']},
    {'key': 'irp001',     'label': 'IRP001 被保险存款申报表',
     'patterns': ['irp001', 'IRP001', '被保险存款申报表'],
     'exts': ['.xls', '.xlsx'], 'required': True},
    {'key': 'irp002',     'label': 'IRP002 申报表',
     'patterns': ['irp002', 'IRP002'],
     'exts': ['.xls', '.xlsx'], 'required': False},
    {'key': 'att2',       'label': '附件2：住房公积金明细表',
     'patterns': ['附件2', '住房公积金明细'],
     'exts': ['.xls', '.xlsx'], 'required': False},
    {'key': 'att3',       'label': '附件3：社保基金存款明细表',
     'patterns': ['附件3', '社保基金存款明细'],
     'exts': ['.xls', '.xlsx'], 'required': False},
    {'key': 'att4',       'label': '附件4：高管人员存款统计表',
     'patterns': ['附件4', '高管人员存款统计'],
     'exts': ['.xls', '.xlsx'], 'required': False},
]


# ============================================================
# 文件匹配工具
# ============================================================

def _is_temp_file(filename):
    """
    判断是否为需要忽略的"非数据文件"（办公软件临时/锁文件、隐藏文件）。

    WPS 打开 .xlsx 时会生成 ".~文件名.xlsx"，Excel 生成 "~$文件名.xlsx"，
    LibreOffice 生成 ".~lock.文件名#"；这些文件不应被当作待处理文件，
    否则可能把结果写进锁文件（软件一关文件就没了）。

    参数:
        filename (str): 文件名或路径

    返回:
        bool: True 表示应忽略
    """
    name = os.path.basename(str(filename))
    if not name:
        return False
    return (name.startswith(('.', '~$', '~')) or
            name.endswith(('.tmp', '.temp')) or
            '.~lock.' in name)


def _is_office_lock_file(filename):
    """
    判断是否为办公软件产生的临时/锁文件（比 _is_temp_file 更严格，
    不包含 .gitignore/.claude 之类的普通隐藏文件，用于日志提示）。

    参数:
        filename (str): 文件名或路径

    返回:
        bool: True 表示是办公软件临时/锁文件
    """
    name = os.path.basename(str(filename))
    return (name.startswith(('.~', '~$')) or
            '.~lock.' in name or
            name.endswith(('.tmp', '.temp')))


def _match_file(filename, fd):
    """
    检查文件名是否匹配某个文件类型定义。

    匹配规则：
      1. 不是办公软件临时/锁文件（.~xxx、~$xxx 等）
      2. 扩展名必须在允许列表中
      3. 文件名不包含排除关键词
      4. 文件名包含至少一个匹配关键词

    参数:
        filename (str): 文件名（不含路径）
        fd (dict): 文件类型定义

    返回:
        bool: True 表示匹配
    """
    name_lower = filename.lower()
    name_no_ext = os.path.splitext(filename)[0]
    ext = os.path.splitext(filename)[1].lower()

    if _is_temp_file(filename):
        return False
    if ext not in fd['exts']:
        return False
    # 排除关键词检查
    if 'exclude' in fd:
        for exc in fd['exclude']:
            if exc in name_no_ext:
                return False
    # 匹配关键词检查
    for pat in fd['patterns']:
        if pat.lower() in name_lower:
            return True
    return False


def scan_folder(folder):
    """
    扫描文件夹，返回每个文件类型的匹配文件列表。

    排序规则：
      - 不含"已生成"的文件排在前面（优先使用原始文件）
      - 匹配多个时取第一个

    参数:
        folder (str): 文件夹路径

    返回:
        dict: {文件类型key: [匹配文件路径列表]}
    """
    if not folder or not os.path.isdir(folder):
        return {fd['key']: [] for fd in FILE_DEFS}

    files = os.listdir(folder)
    result = {}
    for fd in FILE_DEFS:
        key = fd['key']
        matches = []
        for f in files:
            full = os.path.join(folder, f)
            if os.path.isfile(full) and _match_file(f, fd):
                matches.append(full)
        # 排序：不含"已生成"的排前面
        matches.sort(key=lambda p: ('已生成' in os.path.basename(p), p))
        result[key] = matches
    return result


# ============================================================
# 打开文件 / 文件夹（系统默认程序）
# ============================================================

def _open_with_default_app(path):
    """
    用系统默认程序打开文件（如 Excel / WPS）。

    参数:
        path (str): 文件路径
    """
    path = str(path)
    if sys.platform.startswith('darwin'):
        subprocess.Popen(['open', path])
    elif os.name == 'nt':
        os.startfile(path)          # noqa: S606 (Windows 专用)
    else:
        subprocess.Popen(['xdg-open', path])


def _show_in_file_manager(path):
    """
    在系统文件管理器中定位文件（或打开目录）。

    参数:
        path (str): 文件或目录路径
    """
    path = str(path)
    if os.path.isfile(path):
        if sys.platform.startswith('darwin'):
            subprocess.Popen(['open', '-R', path])
        elif os.name == 'nt':
            subprocess.Popen(['explorer', '/select,' + path])
        else:
            subprocess.Popen(['xdg-open', os.path.dirname(path)])
    else:
        _open_with_default_app(path)


# ============================================================
# 模板修改窗口
# ============================================================

class TemplateDialog(QDialog):
    """
    模板修改窗口。

    显示当前机构类型模板的实际位置与候选路径，可：
      · 打开实际使用的模板（系统默认程序，如 Excel/WPS 编辑）
      · 复制模板到工作目录「模板」子目录后再打开（便于长期修改；
        打包运行时内置模板在临时目录，改那里不生效）
      · 用内置模板重置工作目录模板
      · 打开模板所在文件夹
    模板按 工作目录/模板 → 程序目录/templates → 内置模板 的顺序生效，
    改完保存后重新执行「生成保费基数核对表」即使用新模板。
    """

    def __init__(self, parent, bank_name, rel_path, work_dir, log_cb=None):
        """
        参数:
            parent: 父窗口
            bank_name (str): 当前机构类型名称
            rel_path (str): 模板相对路径（bank_config 中的配置值）
            work_dir (str): 当前工作目录
            log_cb (callable, optional): 写主窗口日志的回调
        """
        super().__init__(parent)
        self.setWindowTitle("模板修改")
        self.setMinimumSize(760, 420)
        self.bank_name = bank_name
        self.rel_path = rel_path
        self.work_dir = work_dir or ''
        self.log_cb = log_cb or (lambda msg: None)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        title = QLabel(f"机构类型：{bank_name}    模板文件：{os.path.basename(rel_path)}")
        title.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        layout.addWidget(title)

        tip = QLabel(
            "模板按以下顺序生效（前者优先）：①工作目录「模板」子目录 ②程序目录 templates "
            "③程序内置模板。\n若要长期修改，请用「复制到工作目录并打开」——程序内置模板"
            "打包在程序内部（临时目录），改那里不会保留。\n修改并保存模板后，"
            "重新执行「生成保费基数核对表」即生效；模板里的行结构、科目代码与公式"
            "决定取数与计算口径。")
        tip.setWordWrap(True)
        tip.setFont(QFont("Microsoft YaHei", 8))
        tip.setStyleSheet("color:#555;")
        layout.addWidget(tip)

        self.tb = QTableWidget(0, 3)
        self.tb.setHorizontalHeaderLabels(['来源', '路径', '状态'])
        self.tb.setFont(QFont("Microsoft YaHei", 9))
        self.tb.verticalHeader().setDefaultSectionSize(24)
        self.tb.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tb.setEditTriggers(QAbstractItemView.NoEditTriggers)
        header = self.tb.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tb.doubleClicked.connect(lambda *_: self._open_selected())
        layout.addWidget(self.tb, 1)

        btns = QHBoxLayout()
        btns.setSpacing(8)
        for text, slot, tip_text in (
                ("打开当前模板", self._open_current,
                 "用系统默认程序打开实际生效的模板"),
                ("复制到工作目录并打开", self._copy_and_open,
                 "把模板复制到 <工作目录>/模板/ 后打开编辑（推荐）"),
                ("用内置模板重置", self._reset_from_builtin,
                 "用程序内置模板覆盖工作目录模板（恢复到初始版本）"),
                ("打开所在文件夹", self._open_folder, "在文件管理器中定位模板"),
                ("刷新", self._reload, "重新检查模板位置")):
            btn = QPushButton(text)
            btn.setFont(QFont("Microsoft YaHei", 9))
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip(tip_text)
            btn.clicked.connect(slot)
            btns.addWidget(btn)
        btns.addStretch()
        layout.addLayout(btns)

        box = QDialogButtonBox(QDialogButtonBox.Close)
        box.rejected.connect(self.reject)
        box.accepted.connect(self.accept)
        layout.addWidget(box)

        self._reload()

    # --------------------------------------------------
    def _candidates(self):
        """候选模板列表（含存在状态）"""
        return [(label, path, os.path.exists(path))
                for label, path in template_candidates(self.rel_path, self.work_dir)]

    def _effective(self):
        """当前实际生效的模板 (path, label)"""
        return resolve_template(self.rel_path, self.work_dir)

    def _reload(self):
        """刷新候选列表"""
        eff_path, eff_label = self._effective()
        rows = self._candidates()
        self.tb.setRowCount(0)
        for label, path, exists in rows:
            r = self.tb.rowCount()
            self.tb.insertRow(r)
            self.tb.setItem(r, 0, QTableWidgetItem(label))
            self.tb.setItem(r, 1, QTableWidgetItem(path))
            if os.path.abspath(path) == os.path.abspath(eff_path):
                status = "✓ 使用中" if exists else "（使用中，但文件缺失）"
            else:
                status = "存在" if exists else "不存在"
            item = QTableWidgetItem(status)
            self.tb.setItem(r, 2, item)
        self.tb.resizeRowsToContents()

    def _open_selected(self):
        """双击行：打开该行的模板（存在才打开）"""
        rows = {i.row() for i in self.tb.selectedIndexes()}
        if not rows:
            return
        r = sorted(rows)[0]
        path_item = self.tb.item(r, 1)
        if path_item and os.path.exists(path_item.text()):
            _open_with_default_app(path_item.text())
            self.log_cb(f"[模板] 已打开: {path_item.text()}")
        else:
            QMessageBox.information(self, "提示", "该位置没有模板文件。")

    def _open_current(self):
        """打开实际生效的模板"""
        path, label = self._effective()
        if not os.path.exists(path):
            QMessageBox.warning(self, "模板缺失", f"未找到模板文件：\n{path}")
            return
        try:
            _open_with_default_app(path)
        except Exception as e:
            QMessageBox.critical(self, "打开失败", str(e))
            return
        self.log_cb(f"[模板] 已打开（{label}）: {path}")
        if label == '内置模板':
            QMessageBox.information(
                self, "提示",
                "当前打开的是程序内置模板（位于程序内部临时目录）。\n"
                "在此修改不会保留，请改用「复制到工作目录并打开」。")

    def _copy_and_open(self):
        """复制模板到工作目录并打开"""
        if not self.work_dir or not os.path.isdir(self.work_dir):
            QMessageBox.warning(self, "缺少工作目录",
                                "请先选择有效的工作目录，再复制模板到工作目录。")
            return
        target = work_template_path(self.rel_path, self.work_dir)
        try:
            if os.path.exists(target):
                _, src, copied = install_template_copy(self.rel_path, self.work_dir)
                self.log_cb(f"[模板] 工作目录模板已存在，直接打开: {target}")
            else:
                target, src, copied = install_template_copy(self.rel_path, self.work_dir)
                self.log_cb(f"[模板] 已复制到工作目录: {target}\n"
                            f"       来源: {src}")
            _open_with_default_app(target)
        except Exception as e:
            QMessageBox.critical(self, "复制/打开失败", str(e))
            return
        self._reload()

    def _reset_from_builtin(self):
        """用内置模板覆盖工作目录模板"""
        if not self.work_dir or not os.path.isdir(self.work_dir):
            QMessageBox.warning(self, "缺少工作目录", "请先选择有效的工作目录。")
            return
        target = work_template_path(self.rel_path, self.work_dir)
        if QMessageBox.question(
                self, "确认重置",
                f"将用程序内置模板覆盖：\n{target}\n\n"
                f"（工作目录模板中的修改会丢失）是否继续？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        try:
            builtin, _ = resolve_template(self.rel_path, None)   # 不含工作目录时即程序/内置模板
            if os.path.abspath(builtin) == os.path.abspath(target):
                QMessageBox.information(self, "提示", "当前生效的就是内置模板，无需重置。")
                return
            os.makedirs(os.path.dirname(target), exist_ok=True)
            shutil.copy2(builtin, target)
            self.log_cb(f"[模板] 已用内置模板重置工作目录模板: {target}")
        except Exception as e:
            QMessageBox.critical(self, "重置失败", str(e))
            return
        self._reload()

    def _open_folder(self):
        """打开模板所在文件夹"""
        path, _ = self._effective()
        try:
            _show_in_file_manager(path)
            self.log_cb(f"[模板] 已在文件管理器中定位: {path}")
        except Exception as e:
            QMessageBox.critical(self, "打开失败", str(e))


# ============================================================
# 后台工作线程
# ============================================================
# 说明：耗时操作在 QThread 后台执行，避免 UI 卡顿
# 通过 pyqtSignal 与主线程通信（进度更新、完成通知、错误通知）
# ============================================================


class BaseCheckThread(QThread):
    """后台线程：生成保费基数核对表（合并模板）"""
    progress = pyqtSignal(str)    # 进度消息
    finished = pyqtSignal(list)   # 完成时返回消息列表
    error = pyqtSignal(str)       # 错误消息

    def __init__(self, a1411, a2411, acct_balance, irp001, template, output,
                 attachment2=None, attachment3=None, attachment4=None,
                 adjustments=None):
        super().__init__()
        self.a1411, self.a2411 = a1411, a2411
        self.acct_balance, self.irp001 = acct_balance, irp001
        self.template, self.output = template, output
        self.attachment2, self.attachment3, self.attachment4 = (
            attachment2, attachment3, attachment4)
        self.adjustments = adjustments or []

    def run(self):
        try:
            checker = BaseChecker()
            self.progress.emit("正在加载 A1411...")
            checker.load_a1411(self.a1411)
            self.progress.emit("正在加载 A2411...")
            checker.load_a2411(self.a2411)
            self.progress.emit("正在加载科目余额表...")
            checker.load_account_balance(self.acct_balance)
            self.progress.emit("正在加载 IRP001 申报表...")
            checker.load_irp001(self.irp001)
            checker.load_attachments(self.attachment2, self.attachment3,
                                     self.attachment4)
            checker.set_adjustments(self.adjustments)
            self.progress.emit("正在计算合并保费基数核对表...")
            msgs = checker.compute(self.template, self.output)
            self.finished.emit(msgs)
        except Exception as e:
            self.error.emit(str(e))


class ReportGenThread(QThread):
    """后台线程：生成全面校验报告"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, irp001_path, output_path,
                 irp002_path=None, base_check_path=None,
                 adjustments=None):
        super().__init__()
        self.irp001_path = irp001_path
        self.output_path = output_path
        self.irp002_path = irp002_path
        self.base_check_path = base_check_path
        self.adjustments = adjustments or []

    def run(self):
        try:
            self.progress.emit("正在生成全面校验报告...")
            msgs = generate_validation_report(
                irp001_path=self.irp001_path,
                output_path=self.output_path,
                irp002_path=self.irp002_path,
                base_check_path=self.base_check_path,
                adjustments=self.adjustments)
            self.finished.emit(msgs)
        except Exception as e:
            import traceback
            self.error.emit(traceback.format_exc())


class ValidationSeqThread(QThread):
    """
    后台线程：顺序执行校验。

    使用单线程顺序执行而非多线程并发，
    避免多线程同时写日志导致输出混乱的问题。
    """
    progress = pyqtSignal(str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, irp002_path=None, irp001_path=None,
                 base_check_path=None, adjustments=None):
        super().__init__()
        self.irp002_path = irp002_path
        self.irp001_path = irp001_path
        self.base_check_path = base_check_path
        self.adjustments = adjustments or []

    def run(self):
        try:
            all_msgs = []

            # 校验1: IRP002 校验
            if self.irp002_path:
                self.progress.emit("正在校验 IRP002...")
                msgs = validate_irp002(
                    self.irp002_path,
                    irp001_path=self.irp001_path)
                all_msgs += msgs
            else:
                all_msgs.append("[跳过] IRP002 文件未找到")

            # 校验2: 保费基数核对表 vs IRP001（存款归属核对 + 会计科目核对）
            if self.base_check_path and self.irp001_path:
                self.progress.emit("正在校验 保费基数核对表 vs IRP001...")
                msgs = validate_base_check_vs_irp001(
                    self.base_check_path, self.irp001_path,
                    adjustments=self.adjustments)
                all_msgs += msgs
            else:
                all_msgs.append("[跳过] 保费基数核对表校验")

            self.finished.emit(all_msgs)
        except Exception as e:
            self.error.emit(str(e))


# ============================================================
# 主窗口
# ============================================================

class MainWindow(QMainWindow):
    """
    主窗口类。

    构建完整的图形界面，包含文件选择、参数设置、操作按钮和日志输出。
    所有耗时操作通过 QThread 在后台执行。
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("存款保险申报表审核工具")
        self.setMinimumSize(960, 720)
        self.resize(1000, 780)

        self._current_bank = DEFAULT_BANK
        self._update_template_paths()
        self._apply_style()
        self._setup_ui()

    # ----------------------------------------------------------
    # 模板路径管理
    # ----------------------------------------------------------

    def _update_template_paths(self):
        """
        根据当前银行类型记录模板相对路径，并解析实际生效的模板文件。

        生效顺序：工作目录「模板」子目录 → 程序目录 templates → 程序内置模板。
        """
        self.base_tpl_rel = get_template(self._current_bank, 'base_check_template')
        self.base_tpl_path, self.base_tpl_source = resolve_template(
            self.base_tpl_rel, self._get_folder() if hasattr(self, 'ed_folder') else None)

    def _resolve_template(self, log=False):
        """
        解析当前实际生效的模板（工作目录可能已变化）。

        参数:
            log (bool): 是否把结果写入运行日志

        返回:
            tuple: (模板路径, 来源标签)
        """
        try:
            folder = self._get_folder()
        except Exception:
            folder = None
        path, source = resolve_template(self.base_tpl_rel, folder)
        self.base_tpl_path, self.base_tpl_source = path, source
        if log and (path != getattr(self, '_logged_tpl_path', None)):
            self._logged_tpl_path = path
            self._log(f"[模板] 使用{source}: {path}")
        return path, source

    def _open_template_dialog(self):
        """打开模板修改窗口"""
        self._resolve_template()
        dlg = TemplateDialog(self, self._current_bank, self.base_tpl_rel,
                             self._get_folder(), log_cb=self._log)
        dlg.exec_()
        self._resolve_template(log=True)

    # ----------------------------------------------------------
    # UI 搭建
    # ----------------------------------------------------------

    def _setup_ui(self):
        """搭建完整界面"""
        central = QWidget()
        self.setCentralWidget(central)

        # 外层 ScrollArea 保证单页面可滚动
        outer_scroll = QScrollArea()
        outer_scroll.setWidgetResizable(True)
        outer_scroll.setFrameShape(QFrame.NoFrame)

        inner = QWidget()
        inner.setObjectName("innerWidget")
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(16, 12, 16, 16)
        layout.setSpacing(10)

        # ---- 标题 ----
        title = QLabel("存款保险申报表审核工具")
        title.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # ====== 第〇部分：银行类型选择 ======
        grp_bank = QGroupBox("机构类型")
        grp_bank.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        bly = QHBoxLayout(grp_bank)
        bly.setSpacing(10)
        bly.addWidget(QLabel("选择银行机构类型:"))
        from PyQt5.QtWidgets import QComboBox
        self.cb_bank = QComboBox()
        self.cb_bank.setFont(QFont("Microsoft YaHei", 10))
        self.cb_bank.setMinimumHeight(28)
        self.cb_bank.addItems(get_bank_list())
        self.cb_bank.setCurrentText(self._current_bank)
        self.cb_bank.currentTextChanged.connect(self._on_bank_changed)
        bly.addWidget(self.cb_bank)

        btn_tpl = QPushButton("修改模板…")
        btn_tpl.setFont(QFont("Microsoft YaHei", 9))
        btn_tpl.setCursor(Qt.PointingHandCursor)
        btn_tpl.setToolTip("查看模板位置、打开模板编辑，或复制到工作目录后修改")
        btn_tpl.clicked.connect(self._open_template_dialog)
        bly.addWidget(btn_tpl)

        self.lbl_tpl = QLabel("")
        self.lbl_tpl.setFont(QFont("Microsoft YaHei", 8))
        self.lbl_tpl.setStyleSheet("color:#555;")
        bly.addWidget(self.lbl_tpl)
        bly.addStretch()
        layout.addWidget(grp_bank)

        # ====== 第一部分：文件夹选取 ======
        grp_folder = QGroupBox("工作目录")
        grp_folder.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        fly = QVBoxLayout(grp_folder)
        fly.setSpacing(6)

        # 文件夹选择行
        row_folder = QHBoxLayout()
        row_folder.setSpacing(6)
        row_folder.addWidget(QLabel("选择文件夹:"))
        self.ed_folder = QLineEdit(BASE_DIR)
        self.ed_folder.setFont(QFont("Microsoft YaHei", 9))
        self.ed_folder.setMinimumHeight(28)
        self.ed_folder.textChanged.connect(self._on_folder_changed)
        row_folder.addWidget(self.ed_folder, 1)
        btn_folder = QPushButton("浏览...")
        btn_folder.setFont(QFont("Microsoft YaHei", 9))
        btn_folder.setFixedHeight(28)
        btn_folder.setCursor(Qt.PointingHandCursor)
        btn_folder.clicked.connect(self._browse_folder)
        row_folder.addWidget(btn_folder)
        fly.addLayout(row_folder)

        # 文件匹配状态表格
        self.file_grid = QGridLayout()
        self.file_grid.setSpacing(4)
        self.file_widgets = {}  # key -> {'path_edit', 'status', 'btn'}
        self._build_file_grid()
        fly.addLayout(self.file_grid)
        layout.addWidget(grp_folder)

        # ====== 第二部分：调整值设置（可自由增删的调整项） ======
        grp_adj = QGroupBox("调整值设置（可自由增删调整项）")
        grp_adj.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        adj_layout = QVBoxLayout(grp_adj)
        adj_layout.setSpacing(6)

        tip = QLabel(
            "用途「会计科目核对」：调整会计科目合计（写入表内调整项行）；"
            "用途「存款归属核对」：比较时按方式增减 IRP001 本金（写入 Q9/Q10）。\n"
            "对象选「个人↔单位」表示一增一减（个人增、单位减），"
            "个体工商户即用此项；对象也可选财政性存款等其他合计项，"
            "调整额会并入该合计项。名称可自由填写，标准项不可删除。"
            f"配置自动保存在工作目录的 {ADJ_CONFIG_NAME}。")
        tip.setWordWrap(True)
        tip.setStyleSheet("color:#555;")
        tip.setFont(QFont("Microsoft YaHei", 8))
        adj_layout.addWidget(tip)

        self.tb_adj = QTableWidget(0, 6)
        self.tb_adj.setHorizontalHeaderLabels(
            ['名称', '用途', '对象', '字段', '方式', '金额（元）'])
        self.tb_adj.setFont(QFont("Microsoft YaHei", 9))
        self.tb_adj.verticalHeader().setDefaultSectionSize(26)
        self.tb_adj.setMinimumHeight(150)
        self.tb_adj.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = self.tb_adj.horizontalHeader()
        for idx in (1, 2, 3, 4):
            header.setSectionResizeMode(idx, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        adj_layout.addWidget(self.tb_adj)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_add = QPushButton("＋ 添加项目")
        btn_add.setFont(QFont("Microsoft YaHei", 9))
        btn_add.setCursor(Qt.PointingHandCursor)
        btn_add.clicked.connect(lambda: self._add_adj_row())
        btn_row.addWidget(btn_add)

        btn_del = QPushButton("－ 删除选中")
        btn_del.setFont(QFont("Microsoft YaHei", 9))
        btn_del.setCursor(Qt.PointingHandCursor)
        btn_del.clicked.connect(self._del_adj_row)
        btn_row.addWidget(btn_del)

        btn_seed = QPushButton("从附件5读取")
        btn_seed.setFont(QFont("Microsoft YaHei", 9))
        btn_seed.setCursor(Qt.PointingHandCursor)
        btn_seed.setToolTip("按银行报送的附件5（保费基数核对表）Q7~Q10 "
                            "重建标准调整项（个体工商户 + 久悬户）")
        btn_seed.clicked.connect(self._seed_adjustments_from_att5_clicked)
        btn_row.addWidget(btn_seed)

        btn_reset = QPushButton("恢复标准项")
        btn_reset.setFont(QFont("Microsoft YaHei", 9))
        btn_reset.setCursor(Qt.PointingHandCursor)
        btn_reset.setToolTip("恢复六条标准调整项（金额清零）")
        btn_reset.clicked.connect(lambda: self._set_adjustments(
            build_default_adjustments()))
        btn_row.addWidget(btn_reset)

        btn_row.addStretch()
        adj_layout.addLayout(btn_row)
        layout.addWidget(grp_adj)

        # ====== 第三部分：操作按钮 ======
        grp_actions = QGroupBox("操作")
        grp_actions.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        aly = QHBoxLayout(grp_actions)
        aly.setSpacing(16)

        self.btn_base = self._make_action_btn("1. 生成保费基数核对表", self._run_base_check)
        aly.addWidget(self.btn_base)

        self.btn_val = self._make_action_btn("2. 执行数据校验", self._run_validation)
        aly.addWidget(self.btn_val)

        self.btn_report = self._make_action_btn("3. 生成校验报告", self._run_report)
        aly.addWidget(self.btn_report)

        aly.addStretch()
        layout.addWidget(grp_actions)

        # 进度条
        self.prog_bar = QProgressBar()
        self.prog_bar.setMaximumHeight(4)
        self.prog_bar.setVisible(False)
        layout.addWidget(self.prog_bar)

        # ====== 第四部分：日志输出 ======
        grp_log = QGroupBox("运行日志")
        grp_log.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        lly = QVBoxLayout(grp_log)
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 9))
        self.log_view.setMinimumHeight(220)
        self.log_view.setStyleSheet(
            "background:#1e1e1e;color:#d4d4d4;border:1px solid #333;"
            "border-radius:3px;padding:6px;"
        )
        lly.addWidget(self.log_view)
        layout.addWidget(grp_log)

        # 组装
        outer_scroll.setWidget(inner)
        outer_wrapper = QVBoxLayout(central)
        outer_wrapper.setContentsMargins(0, 0, 0, 0)
        outer_wrapper.addWidget(outer_scroll)

        # 初始扫描
        self._on_folder_changed(BASE_DIR)
        self._update_tpl_label()

    # ----------------------------------------------------------
    # UI 辅助方法
    # ----------------------------------------------------------

    def _make_action_btn(self, text, callback):
        """创建操作按钮（统一样式）"""
        btn = QPushButton(text)
        btn.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        btn.setMinimumHeight(38)
        btn.setMinimumWidth(180)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(callback)
        return btn

    def _build_file_grid(self):
        """
        构建文件匹配状态表格。

        表格列：文件类型 | 匹配文件路径 | 状态 | 手动选择按钮
        每行对应一种文件类型定义（FILE_DEFS）
        """
        # 清空旧组件
        for widgets in self.file_widgets.values():
            for w in widgets.values():
                w.deleteLater()
        self.file_widgets.clear()
        while self.file_grid.count():
            item = self.file_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 表头
        headers = ["文件类型", "匹配文件路径", "状态", "手动选择"]
        for ci, h in enumerate(headers):
            lbl = QLabel(h)
            lbl.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
            lbl.setStyleSheet("color:#555;")
            self.file_grid.addWidget(lbl, 0, ci)

        # 每类文件一行
        for ri, fd in enumerate(FILE_DEFS):
            row = ri + 1
            key = fd['key']

            # 文件类型标签
            lbl = QLabel(f"  {fd['label']}")
            lbl.setFont(QFont("Microsoft YaHei", 9))
            self.file_grid.addWidget(lbl, row, 0)

            # 路径编辑框
            path_edit = QLineEdit()
            path_edit.setFont(QFont("Microsoft YaHei", 8))
            path_edit.setMinimumHeight(26)
            path_edit.setReadOnly(False)
            path_edit.setPlaceholderText("未找到匹配文件...")
            path_edit.setToolTip(fd['label'])
            self.file_grid.addWidget(path_edit, row, 1)

            # 状态标签
            status_lbl = QLabel("✗ 未找到")
            status_lbl.setFont(QFont("Microsoft YaHei", 9))
            status_lbl.setStyleSheet("color:#e74c3c;")
            self.file_grid.addWidget(status_lbl, row, 2)

            # 手动选择按钮
            btn = QPushButton("...")
            btn.setFixedSize(32, 26)
            btn.setFont(QFont("Microsoft YaHei", 9))
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, k=key, e=path_edit: self._browse_file(k, e))
            self.file_grid.addWidget(btn, row, 3)

            self.file_widgets[key] = {
                'path_edit': path_edit,
                'status': status_lbl,
                'btn': btn,
            }

        # 列拉伸比例
        self.file_grid.setColumnStretch(0, 0)
        self.file_grid.setColumnStretch(1, 1)  # 路径列可伸缩
        self.file_grid.setColumnStretch(2, 0)
        self.file_grid.setColumnStretch(3, 0)

    # ----------------------------------------------------------
    # 文件夹扫描与文件匹配
    # ----------------------------------------------------------

    def _on_folder_changed(self, text):
        """
        文件夹路径变化时重新扫描文件。

        自动匹配工作目录下的所有文件，更新文件状态表格。
        输出文件（is_output=True）如果没有匹配到，会自动生成默认路径。
        找到银行报送的附件5（保费基数核对表）时，自动读取其中的
        调整值 c1~c4（Q列前四格）填入界面（仅在界面值为空/0 时）。
        """
        folder = text.strip()
        matches = scan_folder(folder)

        for key, widgets in self.file_widgets.items():
            matched_list = matches.get(key, [])
            if matched_list:
                # 取最佳匹配（排序后第一个）
                best = matched_list[0]
                widgets['path_edit'].setText(best)
                widgets['status'].setText("✓ 已找到")
                widgets['status'].setStyleSheet("color:#27ae60;")
            else:
                # 输出文件类型：自动生成默认路径
                fd = next((f for f in FILE_DEFS if f['key'] == key), None)
                if fd and fd.get('is_output'):
                    default_name = "保费基数核对表_已生成.xlsx"
                    if default_name and folder:
                        default_path = os.path.join(folder, default_name)
                        widgets['path_edit'].setText(default_path)
                        widgets['status'].setText("(自动命名)")
                        widgets['status'].setStyleSheet("color:#f39c12;")
                    else:
                        widgets['path_edit'].setText("")
                        widgets['status'].setText("✗ 未找到")
                        widgets['status'].setStyleSheet("color:#e74c3c;")
                else:
                    widgets['path_edit'].setText("")
                    widgets['status'].setText("✗ 未找到")
                    widgets['status'].setStyleSheet("color:#e74c3c;")

        # 临时/锁文件提示（WPS/Excel 打开文件时产生）
        temp_files = [f for f in (os.listdir(folder) if folder and os.path.isdir(folder) else [])
                      if _is_office_lock_file(f)]
        if temp_files:
            self._log(f"[提示] 目录中存在办公软件临时/锁文件（已忽略）: "
                      f"{', '.join(temp_files[:3])}"
                      + ("..." if len(temp_files) > 3 else ""))
        # 调整项：优先载入本目录配置，其次附件5，最后标准项
        self._refresh_adjustments(matches.get('att5', []))

        # 模板来源（工作目录/模板 子目录存在时会切换来源）
        self._update_tpl_label()

    # ----------------------------------------------------------
    # 调整项表格（可自由增删）
    # ----------------------------------------------------------

    def _add_adj_row(self, item=None):
        """
        在调整项表格末尾添加一行。

        参数:
            item (dict, optional): 预填内容（name/scope/target/field/mode/amount）
        """
        item = item or {}
        locked = bool(item.get('locked', False))
        row = self.tb_adj.rowCount()
        self.tb_adj.insertRow(row)

        # 名称（可自由填写）；locked=True 为标准项（不可删除）
        name_item = QTableWidgetItem(str(item.get('name', '')))
        name_item.setData(Qt.UserRole, {'locked': locked})
        if locked:
            name_item.setToolTip("标准项：对象为「个人↔单位」的一增一减调整，不可删除")
        self.tb_adj.setItem(row, 0, name_item)
        if locked:
            for col in range(self.tb_adj.columnCount()):
                self._mark_locked_cell(row, col)

        # 用途 / 对象 / 字段 / 方式 —— 下拉选择
        combos = [
            (1, ADJ_SCOPES, item.get('scope', ADJ_SCOPES[0])),
            (2, ADJ_TARGET_OPTIONS, item.get('target', ADJ_TARGET_OPTIONS[0])),
            (3, ADJ_FIELDS, item.get('field', ADJ_FIELDS[0])),
            (4, ADJ_MODES, item.get('mode', ADJ_MODES[0])),
        ]
        for col, options, current in combos:
            combo = QComboBox()
            combo.setFont(QFont("Microsoft YaHei", 9))
            combo.addItems(options)
            combo.setCurrentText(current if current in options else options[0])
            if col == 2:
                combo.setToolTip("个人↔单位：个人按方式增减、单位同步反向（一增一减）")
            self.tb_adj.setCellWidget(row, col, combo)
        # 用途切换时联动"字段"（存款归属核对只比较本金）。
        # 注意：回调按控件定位行号（行号会随增删变化），不能用建行时的行号。
        scope_combo = self.tb_adj.cellWidget(row, 1)
        scope_combo.currentTextChanged.connect(
            lambda text, combo=scope_combo: self._on_adj_scope_changed_by(combo))

        # 金额（元）
        amount = item.get('amount', 0.0)
        amount_item = QTableWidgetItem(self._fmt_amount(amount))
        amount_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if locked:
            amount_item.setBackground(self._locked_brush())
        self.tb_adj.setItem(row, 5, amount_item)

        self._on_adj_scope_changed(row)

    def _on_adj_scope_changed_by(self, scope_combo):
        """按用途下拉框定位所在行，再联动"字段"列。"""
        row = next((r for r in range(self.tb_adj.rowCount())
                    if self.tb_adj.cellWidget(r, 1) is scope_combo), None)
        if row is not None:
            self._on_adj_scope_changed(row)

    def _on_adj_scope_changed(self, row):
        """
        用途选择联动：

        · 存款归属核对 → 字段固定为本金并禁用；对象只保留个人/单位
          （校验页只对这两项做久悬户调整）
        · 会计科目核对 → 字段可选；对象含各合计项
        """
        scope_combo = self.tb_adj.cellWidget(row, 1)
        field_combo = self.tb_adj.cellWidget(row, 3)
        target_combo = self.tb_adj.cellWidget(row, 2)
        if scope_combo is None or field_combo is None:
            return
        is_dep = scope_combo.currentText() == ADJ_SCOPE_DEP
        if is_dep:
            field_combo.setCurrentText(ADJ_DEP_FIELD)
            field_combo.setEnabled(False)
        else:
            field_combo.setEnabled(True)
        if target_combo is not None:
            options = ADJ_DEP_TARGETS if is_dep else ADJ_TARGET_OPTIONS
            current = target_combo.currentText()
            if [target_combo.itemText(i) for i in range(target_combo.count())] != list(options):
                target_combo.blockSignals(True)
                target_combo.clear()
                target_combo.addItems(options)
                target_combo.setCurrentText(current if current in options
                                            else options[0])
                target_combo.blockSignals(False)

    @staticmethod
    def _locked_brush():
        """标准项（不可删除）的背景色"""
        from PyQt5.QtGui import QColor, QBrush
        return QBrush(QColor(240, 240, 240))

    def _mark_locked_cell(self, row, col):
        """把标准项所在行标记为灰色"""
        item = self.tb_adj.item(row, col)
        if item is None:
            item = QTableWidgetItem('')
            self.tb_adj.setItem(row, col, item)
        item.setBackground(self._locked_brush())

    def _is_locked_row(self, row):
        """判断某行是否为标准项（不可删除）"""
        item = self.tb_adj.item(row, 0)
        if item is None:
            return False
        data = item.data(Qt.UserRole)
        return bool(isinstance(data, dict) and data.get('locked'))

    def _del_adj_row(self):
        """
        删除选中的调整项行；未选中行时删除最后一行。

        标准项（个体工商户本金/利息，locked=True）不可删除，会被跳过。
        """
        rows = sorted({idx.row() for idx in self.tb_adj.selectedIndexes()},
                      reverse=True)
        if not rows and self.tb_adj.rowCount():
            rows = [self.tb_adj.rowCount() - 1]
        skipped = [r for r in rows if self._is_locked_row(r)]
        for r in rows:
            if self._is_locked_row(r):
                continue
            self.tb_adj.removeRow(r)
        if skipped:
            QMessageBox.information(
                self, "标准项不可删除",
                "「个体工商户本金」「个体工商户利息」为固定标准项"
                "（默认个人增、单位减的一增一减调整），不能删除；\n"
                "如本行不适用，把金额改为 0 即可。")

    def _get_adjustments(self):
        """
        读取界面表格 → 调整项列表。

        返回:
            list[dict]: 每项含 name/scope/target/field/mode/amount
        """
        items = []
        for row in range(self.tb_adj.rowCount()):
            name_widget = self.tb_adj.item(row, 0)
            amount_widget = self.tb_adj.item(row, 5)
            combo = lambda col: (self.tb_adj.cellWidget(row, col).currentText()
                                 if self.tb_adj.cellWidget(row, col) else '')
            items.append({
                'name': name_widget.text().strip() if name_widget else '',
                'scope': combo(1),
                'target': combo(2),
                'field': combo(3),
                'mode': combo(4),
                'amount': self._parse_amount(amount_widget.text()
                                             if amount_widget else ''),
                'locked': self._is_locked_row(row),
            })
        return normalize_adjustments(items)

    def _set_adjustments(self, items):
        """
        用给定的调整项列表重建表格。

        参数:
            items (list[dict]): 调整项列表
        """
        self.tb_adj.setRowCount(0)
        for item in normalize_adjustments(items):
            self._add_adj_row(item)

    @staticmethod
    def _fmt_amount(value):
        """金额展示：整数不带小数，非整数保留两位。"""
        value = float(value or 0)
        return f"{value:,.0f}" if abs(value - round(value)) < 0.005 else f"{value:,.2f}"

    @staticmethod
    def _parse_amount(text):
        """解析金额文本（容忍千分位逗号、中文逗号、空格）。"""
        from common import to_float
        return to_float(text)

    # ----------------------------------------------------------
    # 调整项配置：保存 / 载入 / 从附件5读取
    # ----------------------------------------------------------

    def _adj_config_path(self):
        """调整项配置文件的完整路径（工作目录下）"""
        folder = self._get_folder()
        return os.path.join(folder, ADJ_CONFIG_NAME) if folder else ''

    def _load_adjust_config(self, folder):
        """
        载入工作目录下的调整值配置。

        参数:
            folder (str): 工作目录

        返回:
            list[dict] or None: 调整项列表；文件不存在或读取失败返回 None
        """
        path = os.path.join(folder, ADJ_CONFIG_NAME) if folder else ''
        if not path or not os.path.exists(path):
            return None
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            items = data.get('items') if isinstance(data, dict) else data
            items = normalize_adjustments(items)
            # 兼容旧版配置：个体工商户"个人增 + 单位减"两条合并为标准项
            merged = merge_standard_pairs(items)
            if len(merged) < len(items):
                self._log(f"[调整值] 已合并个体工商户成对项目: "
                          f"{len(items)} → {len(merged)} 项")
            return merged or None
        except Exception as e:
            self._log(f"[调整值] 配置文件读取失败: {e}")
            return None

    def _save_adjust_config(self):
        """
        保存当前调整项到工作目录的配置文件。

        无调整项或工作目录无效时跳过；写入失败只记录日志，不中断操作。
        """
        path = self._adj_config_path()
        if not path:
            return
        try:
            items = self._get_adjustments()
            with open(path, 'w', encoding='utf-8') as f:
                json.dump({'items': items}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._log(f"[调整值] 配置保存失败: {e}")

    def _seed_adjustments_from_att5_clicked(self):
        """按钮：从附件5重建标准调整项（金额取自 Q7~Q10）。"""
        att5 = self._get_path('att5')
        paths = [att5] if att5 else []
        if not paths:
            QMessageBox.information(self, "提示",
                "未找到附件5（保费基数核对表）。\n"
                "请将银行报送的附件5 放入工作目录，或手工填写调整项。")
            return
        self._seed_adjustments_from_att5(paths, force=True)

    def _seed_adjustments_from_att5(self, att5_paths, force=False):
        """
        按标准口径（个体工商户 + 久悬户）重建调整项，金额取自附件5。

        附件5 的 Q7/Q8 = 个人侧净调整额（本金/利息）、Q9/Q10 = 久悬户
        （个 人/单位），对应默认的六条调整项。

        参数:
            att5_paths (list[str]): 附件5 候选路径
            force (bool): True 时无条件覆盖当前表格（按钮触发）

        返回:
            bool: 是否成功重建
        """
        if not att5_paths:
            return
        if not force and any(it['amount'] for it in self._get_adjustments()):
            return   # 已手工填写金额，不覆盖
        adj = read_adjustments(att5_paths[0])
        if not adj:
            return False
        self._set_adjustments(build_default_adjustments(
            adj.get('c1', 0.0), adj.get('c2', 0.0),
            adj.get('c3', 0.0), adj.get('c4', 0.0)))
        self._log(f"[调整值] 已从附件5重建标准调整项: "
                  f"{os.path.basename(att5_paths[0])}  "
                  f"个体工商户本金={adj.get('c1', 0):,.2f} 利息={adj.get('c2', 0):,.2f} "
                  f"个人久悬户={adj.get('c3', 0):,.2f} 单位久悬户={adj.get('c4', 0):,.2f}")
        return True

    def _refresh_adjustments(self, att5_paths):
        """
        工作目录变化时刷新调整项：优先载入本地配置，其次附件5，
        再次保留当前内容，最后用标准项初始化。

        参数:
            att5_paths (list[str]): 该目录下的附件5 候选路径
        """
        items = self._load_adjust_config(self._get_folder())
        if items:
            self._set_adjustments(items)
            self._log(f"[调整值] 已载入配置: {ADJ_CONFIG_NAME}")
            return
        # 已手工填写金额时保留当前内容；否则用附件5（若有）重建标准调整项
        if any(it['amount'] for it in self._get_adjustments()):
            return
        if att5_paths and self._seed_adjustments_from_att5(att5_paths):
            return
        if not self._get_adjustments():
            self._set_adjustments(build_default_adjustments())

    def _on_bank_changed(self, bank_name):
        """银行类型切换时更新模板路径"""
        self._current_bank = bank_name
        self._update_template_paths()
        self._logged_tpl_path = None
        path, source = self._resolve_template(log=True)
        self._log(f"[切换银行类型] {bank_name}")
        self._update_tpl_label()

    def _update_tpl_label(self):
        """刷新机构类型区旁的模板来源提示"""
        if not hasattr(self, 'lbl_tpl'):
            return
        path, source = self._resolve_template()
        self.lbl_tpl.setText(f"模板来源：{source}")
        self.lbl_tpl.setToolTip(path)

    def _browse_folder(self):
        """弹出文件夹选择对话框"""
        folder = QFileDialog.getExistingDirectory(
            self, "选择工作目录", self.ed_folder.text().strip())
        if folder:
            self.ed_folder.setText(folder)

    def _browse_file(self, key, path_edit):
        """
        手动选择文件（兜底方案）。

        当自动匹配不到文件时，用户可以通过此按钮手动选择。
        """
        fd = next((f for f in FILE_DEFS if f['key'] == key), None)
        if fd is None:
            return
        exts = " ".join([f"*{e}" for e in fd['exts']])
        filter_str = f"Excel ({exts});;All (*)"
        start_dir = self.ed_folder.text().strip() or BASE_DIR
        path, _ = QFileDialog.getOpenFileName(
            self, f"选择 {fd['label']}", start_dir, filter_str)
        if path:
            path_edit.setText(path)

    # ----------------------------------------------------------
    # 辅助方法
    # ----------------------------------------------------------

    def _get_path(self, key):
        """获取某类文件的路径"""
        w = self.file_widgets.get(key)
        if w:
            return w['path_edit'].text().strip()
        return ""

    def _get_folder(self):
        """获取工作目录路径"""
        return self.ed_folder.text().strip()

    def _check_required_files(self, keys):
        """
        检查必要文件是否存在。

        参数:
            keys (list): 需要检查的文件类型 key 列表

        返回:
            list: 错误消息列表（空列表表示全部通过）
        """
        errs = []
        for k in keys:
            p = self._get_path(k)
            if not p:
                fd = next((f for f in FILE_DEFS if f['key'] == k), None)
                label = fd['label'] if fd else k
                errs.append(f"未找到: {label}")
            elif not os.path.exists(p):
                fd = next((f for f in FILE_DEFS if f['key'] == k), None)
                label = fd['label'] if fd else k
                errs.append(f"文件不存在: {p}")
        return errs

    # ----------------------------------------------------------
    # 日志输出
    # ----------------------------------------------------------

    def _log(self, msg):
        """向日志区域追加一行文本"""
        self.log_view.append(msg)

    def _log_sep(self):
        """输出分隔线"""
        self.log_view.append("-" * 60)

    # ----------------------------------------------------------
    # 操作 1: 生成保费基数核对表（合并模板）
    # ----------------------------------------------------------

    def _run_base_check(self):
        """
        执行操作1：生成保费基数核对表。

        数据来源：
          - 存款归属核对（C~F列）← A1411 / A2411 月报表（万元）
          - 会计科目核对（G~M列）← 科目余额表（元）
          - 被保险存款申报表（N~P列）← IRP001 申报表（元）
        再按模板公式对子项数据计算合并，并写入校验页差额。
        """
        errs = self._check_required_files(['a1411', 'a2411'])
        if errs:
            QMessageBox.warning(self, "缺少必要文件", "\n".join(errs))
            return

        bal = self._get_path('base_src')
        irp001 = self._get_path('irp001')
        missing = []
        if not bal or not os.path.exists(bal):
            missing.append("科目余额表（源数据）")
        if not irp001 or not os.path.exists(irp001):
            missing.append("IRP001 被保险存款申报表")
        if missing:
            QMessageBox.warning(self, "缺少必要文件",
                "生成保费基数核对表还需要以下文件：\n\n- " + "\n- ".join(missing) +
                "\n\n提示：将文件放入工作目录，工具会自动识别。")
            return

        a1411 = self._get_path('a1411')
        a2411 = self._get_path('a2411')
        tpl, tpl_source = self._resolve_template()
        out = self._get_path('base_out')
        att2 = self._get_path('att2')
        att3 = self._get_path('att3')
        att4 = self._get_path('att4')

        # 输出路径兜底：为空或误选成临时/锁文件时，一律回到标准文件名
        default_out = os.path.join(self._get_folder(), "保费基数核对表_已生成.xlsx")
        if not out or _is_temp_file(out):
            if out:
                self._log(f"[提示] 输出路径指向临时文件，已改为: "
                          f"{os.path.basename(default_out)}")
            out = default_out

        # 若目标文件正被 WPS/Excel 打开（存在锁文件），提示先关闭
        lock_path = os.path.join(os.path.dirname(out), '.' + os.path.basename(out))
        if os.path.exists(lock_path):
            self._log(f"[提示] 检测到 {os.path.basename(out)} 正被 WPS/Excel 打开"
                      f"（存在锁文件 {os.path.basename(lock_path)}），"
                      f"建议关闭该文件后重新生成，避免保存时互相覆盖")

        adjustments = self._get_adjustments()
        self._save_adjust_config()

        self.btn_base.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始生成保费基数核对表（存款归属核对 + 会计科目核对）")
        self._log(f"A1411: {a1411}")
        self._log(f"A2411: {a2411}")
        self._log(f"科目余额表: {bal}")
        self._log(f"IRP001: {irp001}")
        self._log(f"模板（{tpl_source}）: {tpl}")
        self._log(f"调整项: {len(adjustments)} 项（明细见下方运行日志）")
        self._log(f"输出: {out}")
        self._log_sep()

        self.thr_base = BaseCheckThread(
            a1411, a2411, bal, irp001, tpl, out,
            attachment2=att2 if att2 and os.path.exists(att2) else None,
            attachment3=att3 if att3 and os.path.exists(att3) else None,
            attachment4=att4 if att4 and os.path.exists(att4) else None,
            adjustments=adjustments)
        self.thr_base.progress.connect(self._log)
        self.thr_base.finished.connect(self._base_done)
        self.thr_base.error.connect(self._base_err)
        self.thr_base.start()

    def _base_done(self, msgs):
        """保费基数核对表生成完成回调"""
        self.prog_bar.setVisible(False)
        self.btn_base.setEnabled(True)
        self._log_sep()
        for m in msgs:
            self._log(m)
        QMessageBox.information(self, "完成", "\n".join(msgs[-12:]))
        self._on_folder_changed(self._get_folder())

    def _base_err(self, err):
        """保费基数核对表生成错误回调"""
        self.prog_bar.setVisible(False)
        self.btn_base.setEnabled(True)
        self._log(f"ERROR: {err}")
        QMessageBox.critical(self, "错误", err)

    # ----------------------------------------------------------
    # 操作 3: 执行全部数据校验
    # ----------------------------------------------------------

    def _run_validation(self):
        """
        执行操作2：数据校验。

        一次性顺序执行两项校验：
          1. IRP002 内部校验 + 与 IRP001 交叉校验
          2. 保费基数核对表 vs IRP001 校验
             （存款归属核对 F列 + 会计科目核对 I/L/M列）
        """
        folder = self._get_folder()

        adjustments = self._get_adjustments()
        self._save_adjust_config()

        self.btn_val.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始执行数据校验")
        self._log(f"调整项: {len(adjustments)} 项（会计科目核对类已含在核对表内，"
                  f"存款归属核对类用于比较时调整 IRP001 本金）")
        self._log_sep()

        irp001_path = self._get_path('irp001')
        irp002_path = self._get_path('irp002')
        base_path = self._get_path('base_out')

        self.thr_all = ValidationSeqThread(
            irp002_path=(irp002_path if irp002_path and
                         os.path.exists(irp002_path) else None),
            irp001_path=(irp001_path if irp001_path and
                         os.path.exists(irp001_path) else None),
            base_check_path=(base_path if base_path and
                             os.path.exists(base_path) else None),
            adjustments=adjustments)
        self.thr_all.progress.connect(self._log)
        self.thr_all.finished.connect(self._val_all_done)
        self.thr_all.error.connect(self._val_all_err)
        self.thr_all.start()

    def _val_all_done(self, msgs):
        """全部校验完成回调"""
        self.prog_bar.setVisible(False)
        self.btn_val.setEnabled(True)
        self._log_sep()
        for m in msgs:
            self._log(m)
        self._log(">>> 全部校验完成")
        summary_lines = [m for m in msgs if ">>>" in m]
        summary = "\n".join(summary_lines) if summary_lines else msgs[-1]
        QMessageBox.information(self, "校验完成", summary)

    def _val_all_err(self, err):
        """校验过程出错回调"""
        self.prog_bar.setVisible(False)
        self.btn_val.setEnabled(True)
        self._log(f"ERROR: {err}")
        QMessageBox.critical(self, "校验错误", err)

    # ----------------------------------------------------------
    # 操作 3: 生成校验报告
    # ----------------------------------------------------------

    def _run_report(self):
        """
        执行操作3：生成全面校验报告。

        基于 IRP001、保费基数核对表、IRP002 三个数据源，
        生成跨数据源交叉校验的 Excel 报告，不一致项标红显示。
        """
        folder = self._get_folder()
        irp001_path = self._get_path('irp001')
        if not irp001_path or not os.path.exists(irp001_path):
            QMessageBox.warning(self, "缺少必要文件",
                "未找到 IRP001 被保险存款申报表，无法生成校验报告。")
            return

        irp002_path = self._get_path('irp002')
        base_path = self._get_path('base_out')

        # 说明：报告模板若含 IRP002 列才会用到 IRP002，故不再强制要求该文件
        if not irp002_path or not os.path.exists(irp002_path):
            self._log("[提示] 未找到 IRP002 申报表，报告中的 IRP002 列将填“无对应数据”")

        output_path = os.path.join(folder, "校验报告.xlsx")
        adjustments = self._get_adjustments()
        self._save_adjust_config()

        self.btn_report.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始生成全面校验报告（《校验表（新）》格式）")
        self._log(f"IRP001: {irp001_path}")
        self._log(f"IRP002: {irp002_path or '（未提供）'}")
        self._log(f"保费基数核对表: {base_path}")
        self._log(f"调整项: {len(adjustments)} 项")
        self._log(f"输出: {output_path}")
        self._log_sep()

        self.thr_report = ReportGenThread(
            irp001_path=irp001_path,
            output_path=output_path,
            irp002_path=(irp002_path if os.path.exists(irp002_path) else None),
            base_check_path=(base_path if base_path and
                             os.path.exists(base_path) else None),
            adjustments=adjustments)
        self.thr_report.progress.connect(self._log)
        self.thr_report.finished.connect(self._report_done)
        self.thr_report.error.connect(self._report_err)
        self.thr_report.start()

    def _report_done(self, msgs):
        """校验报告生成完成回调"""
        self.prog_bar.setVisible(False)
        self.btn_report.setEnabled(True)
        self._log_sep()
        for m in msgs:
            self._log(m)
        self._log(">>> 校验报告生成完成")
        summary = [m for m in msgs if ">>>" in m or "共校验" in m or "发现" in m]
        QMessageBox.information(
            self, "报告已生成",
            "\n".join(summary) if summary else "校验报告已生成")

    def _report_err(self, err):
        """校验报告生成错误回调"""
        self.prog_bar.setVisible(False)
        self.btn_report.setEnabled(True)
        self._log(f"ERROR: {err}")
        QMessageBox.critical(self, "生成报告错误", err)

    # ----------------------------------------------------------
    # 关闭窗口
    # ----------------------------------------------------------

    def closeEvent(self, event):
        """关闭前保存调整项配置，避免手工填写的内容丢失。"""
        self._save_adjust_config()
        super().closeEvent(event)

    # ----------------------------------------------------------
    # 样式
    # ----------------------------------------------------------

    def _apply_style(self):
        """应用界面 CSS 样式"""
        self.setStyleSheet("""
            QMainWindow { background-color: #f5f6fa; }
            QGroupBox {
                border: 1px solid #dcdde1; border-radius: 5px;
                margin-top: 10px; padding-top: 14px; font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 10px; padding: 0 4px;
                color: #2d3436;
            }
            QLineEdit {
                border: 1px solid #dcdde1; border-radius: 3px;
                padding: 3px 6px; background: white;
            }
            QLineEdit:focus { border-color: #0984e3; }
            QPushButton {
                background-color: #0984e3; color: white; border: none;
                border-radius: 3px; padding: 5px 14px;
            }
            QPushButton:hover { background-color: #0773c5; }
            QPushButton:pressed { background-color: #065fa0; }
            QPushButton:disabled { background-color: #b2bec3; }
            QProgressBar {
                border: none; background: #dfe6e9; border-radius: 2px;
            }
            QProgressBar::chunk { background: #0984e3; border-radius: 2px; }
            QScrollArea { border: none; background: transparent; }
            QWidget#innerWidget { background: transparent; }
        """)


# ============================================================
# Qt 运行环境兼容处理
# ============================================================

def _setup_qt_plugin_paths():
    """
    修正 Qt 插件搜索路径（必须在创建 QApplication 之前调用）。

    部分 PyQt5 安装方式（典型如 conda 的 PyQt5 + qt-main）把 Qt 插件
    放在环境根目录下（<环境>/plugins、<环境>/Library/plugins），
    而 Qt 默认只在 PyQt5 包内查找，导致启动时报：
        Could not find the Qt platform plugin "cocoa" in ""
        This application failed to start because no Qt platform plugin
        could be initialized.
    这里把确实存在 platforms 子目录的位置补进 Qt 的库搜索路径。

    返回:
        list[str]: 补充的插件目录列表（为空表示无需补充）
    """
    from PyQt5.QtCore import QCoreApplication, QLibraryInfo

    candidates = []
    # 1) PyQt5 包内（pip 安装的 wheel 自带；PyInstaller 打包后亦在此列）
    try:
        import PyQt5
        pkg_dir = os.path.dirname(os.path.abspath(PyQt5.__file__))
        candidates += [os.path.join(pkg_dir, 'Qt5', 'plugins'),
                       os.path.join(pkg_dir, 'Qt', 'plugins')]
    except Exception:
        pass
    # 2) 运行环境根目录（conda：Linux/macOS 为 <环境>/plugins，
    #    Windows 为 <环境>\Library\plugins）
    prefix = sys.prefix
    candidates += [os.path.join(prefix, 'plugins'),
                   os.path.join(prefix, 'Library', 'plugins'),
                   os.path.join(prefix, 'lib', 'plugins')]
    # 3) Linux 发行版自带的 Qt5 插件目录（麒麟/统信/Debian/RHEL 系：
    #    /usr/lib/<triplet>/qt5/plugins、/usr/lib64/qt5/plugins 等）
    if not sys.platform.startswith(('darwin', 'win')):
        candidates.append('/usr/lib/qt5/plugins')
        candidates.append('/usr/lib64/qt5/plugins')
        candidates.extend(sorted(glob.glob('/usr/lib/*/qt5/plugins')))
        candidates.extend(sorted(glob.glob('/usr/lib/*/qt6/plugins')))

    existing = set(QCoreApplication.libraryPaths())
    added = []
    for path in candidates:
        path = os.path.abspath(path)
        if path in existing or not os.path.isdir(os.path.join(path, 'platforms')):
            continue
        QCoreApplication.addLibraryPath(path)
        added.append(path)
    if added and sys.stdout is not None:
        # 打包为无控制台的 exe 时 sys.stdout 可能为 None，故加此判断
        print(f"[Qt] 已补充插件搜索路径: {', '.join(added)}")
        print(f"[Qt] 原路径(不存在): {QLibraryInfo.location(QLibraryInfo.PluginsPath)}")
    return added


# ============================================================
# 程序入口
# ============================================================

def main():
    """程序入口函数"""
    _setup_qt_plugin_paths()
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName("存款保险申报表审核工具")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
