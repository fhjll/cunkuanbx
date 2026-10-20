# -*- coding: utf-8 -*-
"""
=============================================================================
银行类型选择 / 模板绑定界面组件
=============================================================================

本模块把"银行选择"从主窗口里独立出来，包含：

  1. BankSelector     — 主界面上的银行类型选择条（下拉框 + 模板信息 +
                         「管理银行/模板…」「更换绑定模板…」「打开模板」）
  2. BankManageDialog — 银行类型与模板绑定管理对话框：
                         · 新增 / 重命名 / 删除银行类型
                         · 为银行类型更换绑定模板（选择任意 .xlsx 文件）
                         · 恢复内置默认模板
                         · 导入（新增）模板到模板库
                         · 打开模板 / 打开模板库文件夹

绑定关系的持久化、校验与文件操作全部由 bank_manager.BankManager 负责，
本模块只做界面与交互。
=============================================================================
"""
import os
import subprocess
import sys

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAbstractItemView, QComboBox, QDialog, QFileDialog, QGridLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QMessageBox,
    QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
)


def open_path(path):
    """用系统默认程序打开文件/文件夹。"""
    path = str(path)
    if sys.platform.startswith('darwin'):
        subprocess.Popen(['open', path])
    elif os.name == 'nt':
        try:
            os.startfile(path)          # noqa: S606 (Windows 专用)
        except AttributeError:
            subprocess.Popen(['explorer', path])
    else:
        subprocess.Popen(['xdg-open', path])


def _small_btn(text, tooltip=''):
    """统一样式的次要按钮。"""
    btn = QPushButton(text)
    btn.setFont(QFont('Microsoft YaHei', 9))
    btn.setCursor(Qt.PointingHandCursor)
    btn.setMinimumHeight(28)
    if tooltip:
        btn.setToolTip(tooltip)
    return btn


# ============================================================
# 主界面上的银行选择条
# ============================================================

class BankSelector(QGroupBox):
    """
    银行类型选择条。

    信号:
        bank_changed(str): 切换了银行类型（参数为银行名称）
        template_changed(str): 当前银行绑定的模板发生变化（参数为模板路径）
    """

    bank_changed = pyqtSignal(str)
    template_changed = pyqtSignal(str)

    def __init__(self, manager, parent=None):
        """
        参数:
            manager (BankManager): 绑定管理器
            parent: 父窗口
        """
        super().__init__('银行类型 / 模板绑定', parent)
        self.mgr = manager
        self.setFont(QFont('Microsoft YaHei', 10, QFont.Bold))
        self._build_ui()
        self.refresh()

    # ---------------- UI ----------------

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        row = QHBoxLayout()
        row.setSpacing(10)
        row.addWidget(QLabel('选择银行机构类型:'))
        self.cb_bank = QComboBox()
        self.cb_bank.setFont(QFont('Microsoft YaHei', 10))
        self.cb_bank.setMinimumHeight(28)
        self.cb_bank.setMinimumWidth(200)
        row.addWidget(self.cb_bank)

        self.btn_manage = _small_btn('管理银行/模板…',
                                     '新增银行类型、导入模板，或把某类银行的'
                                     '模板更换为其他文件')
        self.btn_bind = _small_btn('更换绑定模板…',
                                   '为当前银行类型选择其他模板文件')
        self.btn_open = _small_btn('打开模板',
                                   '用 Excel/WPS 打开当前绑定的模板')
        row.addWidget(self.btn_manage)
        row.addWidget(self.btn_bind)
        row.addWidget(self.btn_open)
        row.addStretch()
        layout.addLayout(row)

        self.lbl_tpl = QLabel('')
        self.lbl_tpl.setFont(QFont('Microsoft YaHei', 8))
        self.lbl_tpl.setStyleSheet('color:#555;')
        self.lbl_tpl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.lbl_tpl)

        self.cb_bank.currentTextChanged.connect(self._on_bank_selected)
        self.btn_manage.clicked.connect(self._open_manager)
        self.btn_bind.clicked.connect(self._bind_template)
        self.btn_open.clicked.connect(self._open_template)

    # ---------------- 数据刷新 ----------------

    def refresh(self, bank=None):
        """按管理器里的银行列表刷新下拉框与模板提示。"""
        names = self.mgr.bank_names()
        cur = bank or self.mgr.last_bank or (names[0] if names else '')
        self.cb_bank.blockSignals(True)
        self.cb_bank.clear()
        self.cb_bank.addItems(names)
        if cur in names:
            self.cb_bank.setCurrentText(cur)
        self.cb_bank.blockSignals(False)
        if cur not in names and names:
            self.mgr.set_last_bank(names[0])
        self._update_label()

    def current_bank(self):
        """当前选中的银行类型名称。"""
        return self.cb_bank.currentText()

    def _on_bank_selected(self, name):
        if not name:
            return
        self.mgr.set_last_bank(name)
        self._update_label()
        self.bank_changed.emit(name)

    def _update_label(self):
        info = self.mgr.template_info(self.current_bank())
        path = info.get('path') or ''
        missing = info.get('bound_missing')
        if missing:
            state = '✗ 绑定缺失（已回退默认模板）'
        else:
            state = '✓' if info.get('exists') else '✗ 缺失'
        name = os.path.basename(path) if path else '（未绑定）'
        text = f"当前模板（{info.get('source')}）：{state}  {name}"
        if missing:
            text += f"　原绑定路径：{missing}"
        self.lbl_tpl.setText(text)
        self.lbl_tpl.setToolTip(path)
        color = '#c0392b' if missing or not info.get('exists') else '#2d7d46'
        self.lbl_tpl.setStyleSheet(f'color:{color};')
        self.btn_open.setEnabled(bool(info.get('exists')))

    # ---------------- 操作 ----------------

    def _emit_template_changed(self):
        info = self.mgr.template_info(self.current_bank())
        self.template_changed.emit(info.get('path') or '')

    def _open_manager(self):
        dlg = BankManageDialog(self, self.mgr, self.current_bank())
        dlg.exec_()
        current = self.current_bank()
        self.refresh(current)
        self.bank_changed.emit(self.current_bank())
        self._emit_template_changed()

    def _bind_template(self):
        bank = self.current_bank()
        info = self.mgr.template_info(bank)
        start = os.path.dirname(info.get('path') or '') or os.path.expanduser('~')
        path, _ = QFileDialog.getOpenFileName(
            self, f'为「{bank}」选择模板文件', start,
            'Excel 模板 (*.xlsx *.xlsm);;所有文件 (*)')
        if not path:
            return
        ok, msg = self.mgr.bind_template(bank, path)
        if not ok:
            QMessageBox.warning(self, '模板不可用', msg)
            return
        self._update_label()
        self._emit_template_changed()
        QMessageBox.information(self, '已绑定模板', msg)

    def _open_template(self):
        info = self.mgr.template_info(self.current_bank())
        path = info.get('path') or ''
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, '模板缺失',
                                f'当前绑定模板的文件不存在：\n{path}\n\n'
                                f'请用「更换绑定模板…」重新选择文件。')
            return
        open_path(path)


# ============================================================
# 银行 / 模板管理对话框
# ============================================================

class BankManageDialog(QDialog):
    """银行类型与模板绑定管理对话框。"""

    def __init__(self, parent, manager, current_bank=None):
        super().__init__(parent)
        self.mgr = manager
        self._preferred_bank = current_bank or manager.last_bank
        self.setWindowTitle('银行类型与模板绑定管理')
        self.setMinimumSize(900, 480)
        self._build_ui()
        self.refresh()

    # ---------------- UI ----------------

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        tip = QLabel(
            '每种银行类型绑定一个《保费基数核对表》模板文件。可以把绑定更换为'
            '任意 .xlsx 模板；也可以新增银行类型、导入（新增）模板文件。\n'
            '绑定关系保存在：' + self.mgr.config_path)
        tip.setWordWrap(True)
        tip.setFont(QFont('Microsoft YaHei', 9))
        tip.setStyleSheet('color:#555;')
        layout.addWidget(tip)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ['银行类型', '绑定模板', '状态', '来源'])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeToContents)
        self.table.doubleClicked.connect(lambda *_: self._open_template())
        self.table.itemSelectionChanged.connect(self._update_buttons)
        layout.addWidget(self.table, 1)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        self.btn_add = _small_btn('新增银行…', '新增一个银行类型')
        self.btn_rename = _small_btn('重命名…')
        self.btn_remove = _small_btn('删除银行', '删除选中的银行类型')
        self.btn_bind = _small_btn('更换绑定模板…',
                                   '为选中的银行类型选择其他模板文件')
        self.btn_reset = _small_btn('恢复默认模板',
                                    '改回该银行类型的内置默认模板')
        self.btn_open = _small_btn('打开模板',
                                   '用 Excel/WPS 打开选中的模板')
        self.btn_import = _small_btn('导入模板（新增模板）…',
                                     '把其他 .xlsx 模板复制到模板库并登记')
        self.btn_open_lib = _small_btn('打开模板库文件夹')
        self.btn_close = _small_btn('关闭')

        grid.addWidget(self.btn_add, 0, 0)
        grid.addWidget(self.btn_rename, 0, 1)
        grid.addWidget(self.btn_remove, 0, 2)
        grid.addWidget(self.btn_bind, 1, 0)
        grid.addWidget(self.btn_reset, 1, 1)
        grid.addWidget(self.btn_open, 1, 2)
        grid.addWidget(self.btn_import, 2, 0)
        grid.addWidget(self.btn_open_lib, 2, 1)
        grid.addWidget(self.btn_close, 2, 3)
        layout.addLayout(grid)

        self.btn_add.clicked.connect(self._add_bank)
        self.btn_rename.clicked.connect(self._rename_bank)
        self.btn_remove.clicked.connect(self._remove_bank)
        self.btn_bind.clicked.connect(self._bind_template)
        self.btn_reset.clicked.connect(self._reset_template)
        self.btn_open.clicked.connect(self._open_template)
        self.btn_import.clicked.connect(self._import_template)
        self.btn_open_lib.clicked.connect(self._open_library)
        self.btn_close.clicked.connect(self.accept)

    # ---------------- 数据刷新 ----------------

    def refresh(self):
        """重新读取管理器数据并刷新表格。"""
        names = self.mgr.bank_names()
        self.table.setRowCount(0)
        select_row = 0
        for name in names:
            info = self.mgr.template_info(name)
            row = self.table.rowCount()
            self.table.insertRow(row)
            exists = info.get('exists')
            missing = info.get('bound_missing')
            if missing:
                status_text = '✗ 绑定缺失（已回退默认）'
            else:
                status_text = '✓ 存在' if exists else '✗ 缺失'
            values = [
                name,
                info.get('path') or '（未绑定）',
                status_text,
                info.get('source') or '',
            ]
            for col, text in enumerate(values):
                item = QTableWidgetItem(str(text))
                if col == 2:
                    item.setForeground(
                        Qt.red if missing or not exists else Qt.darkGreen)
                if col >= 1:
                    item.setToolTip(str(text))
                self.table.setItem(row, col, item)
            if name == self._preferred_bank:
                select_row = row
        if self.table.rowCount():
            self.table.selectRow(select_row)
        self._update_buttons()

    def _selected_bank(self):
        row = self.table.currentRow()
        if row < 0 or row >= self.table.rowCount():
            return None
        item = self.table.item(row, 0)
        return item.text() if item else None

    def _update_buttons(self):
        bank = self._selected_bank()
        has = bank is not None
        for btn in (self.btn_rename, self.btn_remove, self.btn_bind,
                    self.btn_reset, self.btn_open):
            btn.setEnabled(has)
        if has:
            info = self.mgr.template_info(bank)
            self.btn_open.setEnabled(bool(info.get('exists')))

    # ---------------- 操作 ----------------

    def _warn(self, msg):
        QMessageBox.warning(self, '提示', msg)

    def _info(self, msg):
        QMessageBox.information(self, '提示', msg)

    def _add_bank(self):
        name, ok = QInputDialog.getText(
            self, '新增银行类型', '银行类型名称（如：城市商业银行）：')
        if not ok or not name.strip():
            return
        cur = self._selected_bank()
        template = None
        if cur:
            info = self.mgr.template_info(cur)
            if info.get('exists'):
                ans = QMessageBox.question(
                    self, '新增银行类型',
                    f'是否让新银行类型「{name.strip()}」沿用「{cur}」的模板？\n'
                    f'模板：{info["path"]}\n\n'
                    f'选择「No」则使用内置默认模板（稍后可再更换）。',
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if ans == QMessageBox.Yes:
                    template = info['path']
        ok, msg = self.mgr.add_bank(name, template)
        if not ok:
            self._warn(msg)
            return
        self._preferred_bank = name.strip()
        self.refresh()

    def _rename_bank(self):
        old = self._selected_bank()
        if not old:
            return
        new, ok = QInputDialog.getText(self, '重命名银行类型',
                                       f'把「{old}」重命名为：', text=old)
        if not ok or not new.strip() or new.strip() == old:
            return
        ok, msg = self.mgr.rename_bank(old, new.strip())
        if not ok:
            self._warn(msg)
            return
        self._preferred_bank = new.strip()
        self.refresh()

    def _remove_bank(self):
        bank = self._selected_bank()
        if not bank:
            return
        ans = QMessageBox.question(
            self, '删除银行类型',
            f'确定删除银行类型「{bank}」吗？\n（只删除绑定关系，不会删除模板文件。）',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans != QMessageBox.Yes:
            return
        ok, msg = self.mgr.remove_bank(bank)
        if not ok:
            self._warn(msg)
            return
        self._preferred_bank = self.mgr.last_bank
        self.refresh()

    def _bind_template(self):
        bank = self._selected_bank()
        if not bank:
            return
        info = self.mgr.template_info(bank)
        start = os.path.dirname(info.get('path') or '') or os.path.expanduser('~')
        path, _ = QFileDialog.getOpenFileName(
            self, f'为「{bank}」选择模板文件', start,
            'Excel 模板 (*.xlsx *.xlsm);;所有文件 (*)')
        if not path:
            return
        ok, msg = self.mgr.bind_template(bank, path)
        if not ok:
            self._warn(msg)
            return
        self.refresh()
        self._info(msg)

    def _reset_template(self):
        bank = self._selected_bank()
        if not bank:
            return
        default_path = self.mgr.default_template_path(bank)
        ans = QMessageBox.question(
            self, '恢复默认模板',
            f'将「{bank}」的模板恢复为内置默认模板：\n{default_path}\n\n是否继续？',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if ans != QMessageBox.Yes:
            return
        ok, msg = self.mgr.reset_template(bank)
        if not ok:
            self._warn(msg)
            return
        self.refresh()

    def _open_template(self):
        bank = self._selected_bank()
        if not bank:
            return
        info = self.mgr.template_info(bank)
        path = info.get('path') or ''
        if not path or not os.path.isfile(path):
            self._warn(f'模板文件不存在：\n{path}')
            return
        open_path(path)

    def _import_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '选择要导入（新增）的模板文件', os.path.expanduser('~'),
            'Excel 模板 (*.xlsx *.xlsm);;所有文件 (*)')
        if not path:
            return
        ok, msg, dest = self.mgr.import_template(path)
        if not ok:
            self._warn(msg)
            return
        bank = self._selected_bank()
        if bank:
            ans = QMessageBox.question(
                self, '导入模板',
                f'已导入模板：\n{dest}\n\n是否把它绑定给「{bank}」？',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if ans == QMessageBox.Yes:
                self.mgr.bind_template(bank, dest)
        self.refresh()
        self._info(msg)

    def _open_library(self):
        try:
            os.makedirs(self.mgr.library_dir, exist_ok=True)
        except OSError:
            pass
        open_path(self.mgr.library_dir)
