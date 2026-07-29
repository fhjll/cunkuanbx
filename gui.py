# -*- coding: utf-8 -*-
"""
=============================================================================
存款保险申报表审核工具 - 图形界面 v9
=============================================================================

程序入口与主界面。

界面布局（从上到下，单页面可滚动）：
  第〇部分 — 机构类型选择（银行类型下拉框）
  第一部分 — 工作目录（文件夹路径 + 文件匹配状态表格）
  第二部分 — 调整值设置（c1/c2/c3/c4 四个调整参数）
  第三部分 — 操作按钮（4个主要操作按钮 + 进度条）
  第四部分 — 运行日志（黑色背景的日志输出区域）

4个主要操作（需按顺序执行）：
  1. 生成存款归属表   — 基于 A1411/A2411 计算存款归属表
  2. 生成科目余额核对表 — 基于科目余额表生成核对表
  3. 执行数据校验     — 一次执行三项校验（IRP002校验、存款归属表校验、科目余额校验）
  4. 生成校验报告     — 输出跨数据源交叉校验的 Excel 报告

工作流程：
  选取工作目录 → 自动扫描匹配文件 → 调整参数（可选）→ 按顺序执行 1→2→3→4

关键特性：
  - 智能文件匹配：根据文件名关键词自动匹配所需文件
  - 支持手动选择文件路径（兜底方案）
  - 所有耗时操作都在后台线程中执行，UI 不卡顿
  - 日志实时输出到黑色背景的日志区域

依赖模块：
  - PyQt5：图形界面框架
  - deposit_calculator.py：存款归属表计算引擎
  - account_checker.py：科目余额核对表生成引擎
  - validator.py：数据校验引擎
  - bank_config.py：银行类型配置
  - common.py：公共工具
=============================================================================
"""
import sys
import os
import re
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
    QGroupBox, QMessageBox, QProgressBar, QScrollArea, QFrame, QGridLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from deposit_calculator import DepositCalculator
from account_checker import AccountChecker
from validator import (
    validate_irp002, validate_deposit_vs_irp001, validate_acct_vs_irp001,
    generate_validation_report)
from bank_config import get_bank_list, get_template, DEFAULT_BANK
from common import get_resource_path


# 脚本所在目录（也是默认工作目录）
BASE_DIR = str(Path(__file__).parent)


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
    {'key': 'dep_out',    'label': '存款归属表（输出）',
     'patterns': ['存款归属表_已生成'],
     'exts': ['.xlsx'], 'required': False, 'is_output': True},
    {'key': 'acct_src',   'label': '科目余额表（源数据）',
     'patterns': ['科目余额表', '科目余额'],
     'exts': ['.xls', '.xlsx'], 'required': False,
     'exclude': ['核对', '已生成']},
    {'key': 'acct_out',   'label': '科目余额核对表（输出）',
     'patterns': ['科目余额核对表_已生成'],
     'exts': ['.xlsx'], 'required': False, 'is_output': True},
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

def _match_file(filename, fd):
    """
    检查文件名是否匹配某个文件类型定义。

    匹配规则：
      1. 扩展名必须在允许列表中
      2. 文件名不包含排除关键词
      3. 文件名包含至少一个匹配关键词

    参数:
        filename (str): 文件名（不含路径）
        fd (dict): 文件类型定义

    返回:
        bool: True 表示匹配
    """
    name_lower = filename.lower()
    name_no_ext = os.path.splitext(filename)[0]
    ext = os.path.splitext(filename)[1].lower()

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
# 后台工作线程
# ============================================================
# 说明：耗时操作在 QThread 后台执行，避免 UI 卡顿
# 通过 pyqtSignal 与主线程通信（进度更新、完成通知、错误通知）
# ============================================================


class DepositComputeThread(QThread):
    """后台线程：计算存款归属表"""
    progress = pyqtSignal(str)    # 进度消息
    finished = pyqtSignal(list)   # 完成时返回消息列表
    error = pyqtSignal(str)       # 错误消息

    def __init__(self, a1411, a2411, template, output,
                 attachment2=None, attachment3=None):
        super().__init__()
        self.a1411, self.a2411 = a1411, a2411
        self.template, self.output = template, output
        self.attachment2, self.attachment3 = attachment2, attachment3

    def run(self):
        try:
            calc = DepositCalculator()
            self.progress.emit("正在加载 A1411...")
            calc.load_a1411(self.a1411)
            self.progress.emit("正在加载 A2411...")
            calc.load_a2411(self.a2411)
            self.progress.emit("正在计算存款归属表...")
            msgs = calc.compute(self.template, self.output,
                                attachment2=self.attachment2,
                                attachment3=self.attachment3)
            self.finished.emit(msgs)
        except Exception as e:
            self.error.emit(str(e))


class AcctCheckThread(QThread):
    """后台线程：生成科目余额核对表"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, acct_balance, template, output,
                 attachment2=None, attachment3=None, attachment4=None):
        super().__init__()
        self.acct_balance = acct_balance
        self.template, self.output = template, output
        self.attachment2, self.attachment3, self.attachment4 = (
            attachment2, attachment3, attachment4)

    def run(self):
        try:
            checker = AccountChecker()
            self.progress.emit("正在加载科目余额表...")
            checker.load_account_balance(self.acct_balance)
            self.progress.emit("正在填入核对表...")
            msgs = checker.compute(self.template, self.output,
                                   attachment2=self.attachment2,
                                   attachment3=self.attachment3,
                                   attachment4=self.attachment4)
            self.finished.emit(msgs)
        except Exception as e:
            self.error.emit(str(e))


class ReportGenThread(QThread):
    """后台线程：生成全面校验报告"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, irp001_path, output_path,
                 irp002_path=None, deposit_path=None, acct_check_path=None,
                 c1=0, c2=0, c3=0, c4=0):
        super().__init__()
        self.irp001_path = irp001_path
        self.output_path = output_path
        self.irp002_path = irp002_path
        self.deposit_path = deposit_path
        self.acct_check_path = acct_check_path
        self.c1, self.c2, self.c3, self.c4 = c1, c2, c3, c4

    def run(self):
        try:
            self.progress.emit("正在生成全面校验报告...")
            msgs = generate_validation_report(
                irp001_path=self.irp001_path,
                output_path=self.output_path,
                irp002_path=self.irp002_path,
                deposit_path=self.deposit_path,
                acct_check_path=self.acct_check_path,
                c1=self.c1, c2=self.c2, c3=self.c3, c4=self.c4)
            self.finished.emit(msgs)
        except Exception as e:
            import traceback
            self.error.emit(traceback.format_exc())


class ValidationSeqThread(QThread):
    """
    后台线程：顺序执行全部三项校验。

    使用单线程顺序执行而非多线程并发，
    避免多线程同时写日志导致输出混乱的问题。
    """
    progress = pyqtSignal(str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, irp002_path=None, irp001_path=None,
                 dep_path=None, acct_path=None,
                 c1=0, c2=0, c3=0, c4=0):
        super().__init__()
        self.irp002_path = irp002_path
        self.irp001_path = irp001_path
        self.dep_path = dep_path
        self.acct_path = acct_path
        self.c1, self.c2, self.c3, self.c4 = c1, c2, c3, c4

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

            # 校验2: 存款归属表 vs IRP001
            if self.dep_path and self.irp001_path:
                self.progress.emit("正在校验 存款归属表 vs IRP001...")
                msgs = validate_deposit_vs_irp001(
                    self.dep_path, self.irp001_path,
                    self.c3, self.c4)
                all_msgs += msgs
            else:
                all_msgs.append("[跳过] 存款归属表校验")

            # 校验3: 科目余额表 vs IRP001
            if self.acct_path and self.irp001_path:
                self.progress.emit("正在校验 科目余额表 vs IRP001...")
                msgs = validate_acct_vs_irp001(
                    self.acct_path, self.irp001_path,
                    self.c1, self.c2)
                all_msgs += msgs
            else:
                all_msgs.append("[跳过] 科目余额表校验")

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
        """根据当前银行类型更新内嵌模板路径"""
        self.dep_tpl_path = get_resource_path(
            get_template(self._current_bank, 'deposit_template'))
        self.acct_tpl_path = get_resource_path(
            get_template(self._current_bank, 'acct_check_template'))

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
        bly.addWidget(QLabel("（模板已内置于程序，无需外置）"))
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

        # ====== 第二部分：调整值设置 ======
        grp_adj = QGroupBox("调整值设置（用于数据校验）")
        grp_adj.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        adj_layout = QGridLayout(grp_adj)
        adj_layout.setSpacing(8)

        # (第0行) c1: 个体工商户本金调整值
        adj_layout.addWidget(QLabel("c1 个体工商户本金调整值:"), 0, 0)
        self.ed_c1 = QLineEdit("0")
        self.ed_c1.setMinimumHeight(28)
        self.ed_c1.setFixedWidth(180)
        self.ed_c1.setToolTip("校验时从单位存款本金扣除，加到个人存款本金")
        adj_layout.addWidget(self.ed_c1, 0, 1)
        adj_layout.addWidget(QLabel("元"), 0, 2)

        # (第0行) c2: 个体工商户利息调整值
        adj_layout.addWidget(QLabel("c2 个体工商户利息调整值:"), 0, 3)
        self.ed_c2 = QLineEdit("0")
        self.ed_c2.setMinimumHeight(28)
        self.ed_c2.setFixedWidth(180)
        self.ed_c2.setToolTip("校验时从单位存款利息扣除，加到个人存款利息")
        adj_layout.addWidget(self.ed_c2, 0, 4)
        adj_layout.addWidget(QLabel("元"), 0, 5)

        # (第1行) c3: 个人久悬户余额
        adj_layout.addWidget(QLabel("c3 个人久悬户余额:"), 1, 0)
        self.ed_c3 = QLineEdit("0")
        self.ed_c3.setMinimumHeight(28)
        self.ed_c3.setFixedWidth(180)
        self.ed_c3.setToolTip("校验时从 IRP001 个人本金中扣除后，精确至万元再比较")
        adj_layout.addWidget(self.ed_c3, 1, 1)
        adj_layout.addWidget(QLabel("元"), 1, 2)

        # (第1行) c4: 单位久悬户余额
        adj_layout.addWidget(QLabel("c4 单位久悬户余额:"), 1, 3)
        self.ed_c4 = QLineEdit("0")
        self.ed_c4.setMinimumHeight(28)
        self.ed_c4.setFixedWidth(180)
        self.ed_c4.setToolTip("校验时从 IRP001 单位本金中扣除后，精确至万元再比较")
        adj_layout.addWidget(self.ed_c4, 1, 4)
        adj_layout.addWidget(QLabel("元"), 1, 5)

        layout.addWidget(grp_adj)

        # ====== 第三部分：操作按钮 ======
        grp_actions = QGroupBox("操作")
        grp_actions.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        aly = QHBoxLayout(grp_actions)
        aly.setSpacing(16)

        self.btn_dep = self._make_action_btn("1. 生成存款归属表", self._run_deposit)
        aly.addWidget(self.btn_dep)

        self.btn_acct = self._make_action_btn("2. 生成科目余额核对表", self._run_acct_check)
        aly.addWidget(self.btn_acct)

        self.btn_val = self._make_action_btn("3. 执行数据校验", self._run_validation)
        aly.addWidget(self.btn_val)

        self.btn_report = self._make_action_btn("4. 生成校验报告", self._run_report)
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
                    if key == 'dep_out':
                        default_name = "存款归属表_已生成.xlsx"
                    elif key == 'acct_out':
                        default_name = "科目余额核对表_已生成.xlsx"
                    else:
                        default_name = ""
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

    def _on_bank_changed(self, bank_name):
        """银行类型切换时更新内嵌模板路径"""
        self._current_bank = bank_name
        self._update_template_paths()
        self._log(f"[切换银行类型] {bank_name}")
        self._log(f"  存款归属表模板: {os.path.basename(self.dep_tpl_path)}")
        self._log(f"  科目余额核对表模板: {os.path.basename(self.acct_tpl_path)}")

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

    def _get_adj_value(self, edit):
        """安全解析调整值输入框的数字"""
        try:
            return float(edit.text().strip())
        except ValueError:
            return 0.0

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
    # 操作 1: 生成存款归属表
    # ----------------------------------------------------------

    def _run_deposit(self):
        """
        执行操作1：生成存款归属表。

        从 A1411 和 A2411 中提取数据，按模板结构计算并填充存款归属表。
        """
        errs = self._check_required_files(['a1411', 'a2411'])
        if errs:
            QMessageBox.warning(self, "缺少必要文件", "\n".join(errs))
            return

        a1411 = self._get_path('a1411')
        a2411 = self._get_path('a2411')
        tpl = self.dep_tpl_path
        out = self._get_path('dep_out')
        att2 = self._get_path('att2')
        att3 = self._get_path('att3')

        if not out:
            out = os.path.join(self._get_folder(), "存款归属表_已生成.xlsx")

        self.btn_dep.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始生成存款归属表")
        self._log(f"A1411: {a1411}")
        self._log(f"A2411: {a2411}")
        self._log(f"输出: {out}")
        self._log_sep()

        self.thr_dep = DepositComputeThread(
            a1411, a2411, tpl, out,
            attachment2=att2 if att2 and os.path.exists(att2) else None,
            attachment3=att3 if att3 and os.path.exists(att3) else None)
        self.thr_dep.progress.connect(self._log)
        self.thr_dep.finished.connect(self._dep_done)
        self.thr_dep.error.connect(self._dep_err)
        self.thr_dep.start()

    def _dep_done(self, msgs):
        """存款归属表生成完成回调"""
        self.prog_bar.setVisible(False)
        self.btn_dep.setEnabled(True)
        self._log_sep()
        for m in msgs:
            self._log(m)
        QMessageBox.information(self, "完成", "\n".join(msgs))
        self._on_folder_changed(self._get_folder())

    def _dep_err(self, err):
        """存款归属表生成错误回调"""
        self.prog_bar.setVisible(False)
        self.btn_dep.setEnabled(True)
        self._log(f"ERROR: {err}")
        QMessageBox.critical(self, "错误", err)

    # ----------------------------------------------------------
    # 操作 2: 生成科目余额核对表
    # ----------------------------------------------------------

    def _run_acct_check(self):
        """
        执行操作2：生成科目余额核对表。

        从科目余额表中提取各科目余额，按模板结构填入核对表。
        同时从附件2/3/4中提取住房公积金、社保基金、高管存款数据。
        """
        bal = self._get_path('acct_src')
        tpl = self.acct_tpl_path
        out = self._get_path('acct_out')
        att2 = self._get_path('att2')
        att3 = self._get_path('att3')
        att4 = self._get_path('att4')

        if not bal or not os.path.exists(bal):
            QMessageBox.warning(self, "缺少必要文件",
                "未找到科目余额表（源数据）或科目余额核对表（已生成）。\n\n"
                "提示：将科目余额表.xls 放入工作目录，工具会自动识别。")
            return

        if not out:
            out = os.path.join(self._get_folder(), "科目余额核对表_已生成.xlsx")

        self.btn_acct.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始生成科目余额核对表")
        self._log(f"数据源: {bal}")
        self._log(f"模板: {tpl}")
        self._log(f"输出: {out}")
        self._log("(c1/c2/c3/c4 调整值将在点击「校验」时应用)")
        self._log_sep()

        self.thr_acct = AcctCheckThread(
            bal, tpl, out,
            attachment2=att2 if att2 and os.path.exists(att2) else None,
            attachment3=att3 if att3 and os.path.exists(att3) else None,
            attachment4=att4 if att4 and os.path.exists(att4) else None)
        self.thr_acct.progress.connect(self._log)
        self.thr_acct.finished.connect(self._acct_done)
        self.thr_acct.error.connect(self._acct_err)
        self.thr_acct.start()

    def _acct_done(self, msgs):
        """科目余额核对表生成完成回调"""
        self.prog_bar.setVisible(False)
        self.btn_acct.setEnabled(True)
        self._log_sep()
        for m in msgs:
            self._log(m)
        QMessageBox.information(self, "完成", "\n".join(msgs))
        self._on_folder_changed(self._get_folder())

    def _acct_err(self, err):
        """科目余额核对表生成错误回调"""
        self.prog_bar.setVisible(False)
        self.btn_acct.setEnabled(True)
        self._log(f"ERROR: {err}")
        QMessageBox.critical(self, "错误", err)

    # ----------------------------------------------------------
    # 操作 3: 执行全部数据校验
    # ----------------------------------------------------------

    def _run_validation(self):
        """
        执行操作3：数据校验。

        一次性顺序执行三项校验：
          1. IRP002 内部校验 + 与 IRP001 交叉校验
          2. 存款归属表 vs IRP001 校验
          3. 科目余额表 vs IRP001 校验
        """
        folder = self._get_folder()

        c1 = self._get_adj_value(self.ed_c1)
        c2 = self._get_adj_value(self.ed_c2)
        c3 = self._get_adj_value(self.ed_c3)
        c4 = self._get_adj_value(self.ed_c4)

        self.btn_val.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始执行全部数据校验")
        self._log(f"c1 (个体工商户本金调整值): {c1:,.2f} 元 "
                  f"(科目余额表: 个人+c1, 单位-c1)")
        self._log(f"c2 (个体工商户利息调整值): {c2:,.2f} 元 "
                  f"(科目余额表: 个人+c2, 单位-c2)")
        self._log(f"c3 (个人久悬户余额):       {c3:,.2f} 元 "
                  f"(存款归属表: 个人-c3)")
        self._log(f"c4 (单位久悬户余额):       {c4:,.2f} 元 "
                  f"(存款归属表: 单位-c4)")
        self._log_sep()

        irp001_path = self._get_path('irp001')
        irp002_path = self._get_path('irp002')
        dep_path = self._get_path('dep_out')
        acct_path = self._get_path('acct_out')

        self.thr_all = ValidationSeqThread(
            irp002_path=(irp002_path if irp002_path and
                         os.path.exists(irp002_path) else None),
            irp001_path=(irp001_path if irp001_path and
                         os.path.exists(irp001_path) else None),
            dep_path=(dep_path if dep_path and
                      os.path.exists(dep_path) else None),
            acct_path=(acct_path if acct_path and
                       os.path.exists(acct_path) else None),
            c1=c1, c2=c2, c3=c3, c4=c4)
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
    # 操作 4: 生成校验报告
    # ----------------------------------------------------------

    def _run_report(self):
        """
        执行操作4：生成全面校验报告。

        基于 IRP001、IRP002、存款归属表、科目余额核对表四个数据源，
        生成跨数据源交叉校验的 Excel 报告，不一致项标红显示。
        """
        folder = self._get_folder()
        irp001_path = self._get_path('irp001')
        if not irp001_path or not os.path.exists(irp001_path):
            QMessageBox.warning(self, "缺少必要文件",
                "未找到 IRP001 被保险存款申报表，无法生成校验报告。")
            return

        irp002_path = self._get_path('irp002')
        dep_path = self._get_path('dep_out')
        acct_path = self._get_path('acct_out')

        if not irp002_path or not os.path.exists(irp002_path):
            QMessageBox.warning(self, "缺少必要文件",
                "未找到 IRP002 申报表，无法进行 IRP002 交叉校验。\n"
                "请将 IRP002 文件放入工作目录。")
            return

        output_path = os.path.join(folder, "校验报告.xlsx")
        c1 = self._get_adj_value(self.ed_c1)
        c2 = self._get_adj_value(self.ed_c2)
        c3 = self._get_adj_value(self.ed_c3)
        c4 = self._get_adj_value(self.ed_c4)

        self.btn_report.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.log_view.clear()
        self._log(">>> 开始生成全面校验报告")
        self._log(f"IRP001: {irp001_path}")
        self._log(f"IRP002: {irp002_path}")
        self._log(f"存款归属表: {dep_path}")
        self._log(f"科目余额核对表: {acct_path}")
        self._log(f"c1={c1:,.2f} c2={c2:,.2f} c3={c3:,.2f} c4={c4:,.2f}")
        self._log(f"输出: {output_path}")
        self._log_sep()

        self.thr_report = ReportGenThread(
            irp001_path=irp001_path,
            output_path=output_path,
            irp002_path=(irp002_path if os.path.exists(irp002_path) else None),
            deposit_path=(dep_path if dep_path and
                          os.path.exists(dep_path) else None),
            acct_check_path=(acct_path if acct_path and
                             os.path.exists(acct_path) else None),
            c1=c1, c2=c2, c3=c3, c4=c4)
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
# 程序入口
# ============================================================

def main():
    """程序入口函数"""
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName("存款保险申报表审核工具")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
