"""分步向导界面：选择文件夹 → 表头映射 → 汇总设置 → 运行。

所有耗时操作（扫描、运行）放在后台线程，界面不卡死。
界面只负责编辑 TaskConfig，真正的处理全部交给 core.pipeline，保证与命令行结果一致。
"""
from __future__ import annotations

import copy
import subprocess
import sys
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from PySide6.QtCore import QObject, Qt, QThread, QUrl, Signal, Slot
from PySide6.QtGui import QAction, QBrush, QColor, QDesktopServices, QIcon
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox,
                               QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem,
                               QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
                               QPushButton, QStackedWidget, QTableWidget, QTableWidgetItem,
                               QToolBar, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

from config.task_config import AggSpec, TaskConfig
from core import __version__, aggregator, pipeline, reader
from core.header_mapper import (STATUS_AUTO, STATUS_IGNORED, STATUS_MANUAL, STATUS_PENDING,
                                STATUS_UNMATCHED, AliasStore, app_data_dir, normalize_header)
from core.validator import ERROR, INFO, WARNING, IssueCollector

APP_TITLE = "SheetMerger 多表汇总工具"
APP_ICON = Path(__file__).with_name("app_icon.png")
STEP_TITLES = ["① 选择文件夹", "② 表头映射", "③ 汇总设置", "④ 运行"]

KEEP_NAME = "（保留原表头）"
IGNORE = "（忽略此列）"

STATUS_COLORS = {
    STATUS_AUTO: "#E2EFDA",       # 绿：自动
    STATUS_PENDING: "#FFF2CC",    # 黄：待确认
    STATUS_UNMATCHED: "#F8CBAD",  # 红：未匹配
    STATUS_MANUAL: "#DDEBF7",     # 蓝：已手动确认
    STATUS_IGNORED: "#EDEDED",    # 灰：忽略
}
SEVERITY_COLORS = {ERROR: "#C00000", WARNING: "#C65911", INFO: "#595959"}
DEDUP_LABELS = [("只标记重复（保留全部记录）", aggregator.DEDUP_MARK),
                ("删除重复（保留第一次出现的）", aggregator.DEDUP_DROP),
                ("不检查重复", aggregator.DEDUP_OFF)]
AGG_FUNC_NAMES = list(aggregator.AGG_FUNCS)
ISSUE_PREVIEW_LIMIT = 1000


# ====================================================================== 后台任务
class _Worker(QObject):
    progress = Signal(int, str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, fn: Callable, cancel: threading.Event) -> None:
        super().__init__()
        self.fn, self.cancel = fn, cancel

    def run(self) -> None:
        try:
            self.done.emit(self.fn(lambda p, m: self.progress.emit(p, m), self.cancel))
        except pipeline.Cancelled:
            self.failed.emit("已取消")
        except Exception:  # noqa: BLE001 - 后台异常转成界面提示，不让程序崩溃
            self.failed.emit(traceback.format_exc())


@dataclass
class ScanResult:
    files: list[str]                                  # 扫描到的全部文件（相对路径）
    refs: list[reader.SheetRef]
    plans: list[pipeline.TablePlan]
    issues: IssueCollector
    key: tuple = ()
    plan_by_key: dict[str, pipeline.TablePlan] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.plan_by_key = {p.table.key: p for p in self.plans}


def _scan_key(cfg: TaskConfig) -> tuple:
    return (cfg.input_folder, cfg.recursive, tuple(sorted(cfg.header_rows.items())))


def _colored_item(text: Any, color: str | None = None, editable: bool = False) -> QTableWidgetItem:
    item = QTableWidgetItem("" if text is None else str(text))
    if not editable:
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
    if color:
        item.setBackground(QBrush(QColor(color)))
    return item


# ====================================================================== 第 1 步
class FolderPage(QWidget):
    scan_requested = Signal()

    COL_NAME, COL_HEADER, COL_ROWS, COL_STATUS = range(4)

    def __init__(self) -> None:
        super().__init__()
        self.scan: ScanResult | None = None
        self._loading = False

        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("选择包含表格文件的文件夹")
        browse = QPushButton("浏览…")
        browse.clicked.connect(self._browse)
        self.recursive_box = QCheckBox("包含子文件夹")
        self.scan_btn = QPushButton("扫描")
        self.scan_btn.clicked.connect(self.scan_requested)

        row = QHBoxLayout()
        row.addWidget(QLabel("文件夹："))
        row.addWidget(self.folder_edit, 1)
        row.addWidget(browse)
        row.addWidget(self.recursive_box)
        row.addWidget(self.scan_btn)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["文件 / Sheet", "表头行", "数据行数", "状态"])
        self.tree.header().setSectionResizeMode(self.COL_NAME, QHeaderView.ResizeToContents)
        self.tree.header().setStretchLastSection(True)
        self.tree.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.tree.itemChanged.connect(self._item_changed)
        self.summary = QLabel("请选择文件夹后点击“扫描”。勾选要参与汇总的 Sheet；双击“表头行”可手动修改。")
        self.summary.setWordWrap(True)

        lay = QVBoxLayout(self)
        lay.addLayout(row)
        lay.addWidget(self.tree, 1)
        lay.addWidget(self.summary)

    def _browse(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", self.folder_edit.text())
        if folder:
            self.folder_edit.setText(folder)
            self.scan_requested.emit()

    # ---------------------------------------------------------------- 配置读写
    def load_config(self, cfg: TaskConfig) -> None:
        self.folder_edit.setText(cfg.input_folder)
        self.recursive_box.setChecked(cfg.recursive)

    def save_basic(self, cfg: TaskConfig) -> None:
        cfg.input_folder = self.folder_edit.text().strip()
        cfg.recursive = self.recursive_box.isChecked()

    def save_config(self, cfg: TaskConfig) -> None:
        """把勾选状态和表头行写回配置（需已扫描）。"""
        self.save_basic(cfg)
        if self.scan is None:
            return
        excluded: list[str] = []
        header_rows = {k: v for k, v in cfg.header_rows.items()}
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            rel = top.data(0, Qt.UserRole)
            children = [top.child(j) for j in range(top.childCount())]
            if not children:
                continue
            unchecked = [c for c in children if c.checkState(0) != Qt.Checked]
            if len(unchecked) == len(children):
                excluded.append(f"{rel}|*")
            else:
                excluded += [c.data(0, Qt.UserRole) for c in unchecked]
            for c in children:
                key = c.data(0, Qt.UserRole)
                plan = self.scan.plan_by_key.get(key)
                text = c.text(self.COL_HEADER).strip()
                if not text.isdigit() or plan is None:
                    continue
                value = int(text)
                if value == plan.table.auto_header_row:
                    header_rows.pop(key, None)
                else:
                    header_rows[key] = value
        cfg.excluded = excluded
        cfg.header_rows = header_rows

    # ---------------------------------------------------------------- 显示扫描结果
    def show_scan(self, scan: ScanResult, cfg: TaskConfig) -> None:
        self.scan = scan
        self._loading = True
        self.tree.clear()
        file_issues: dict[str, str] = {}
        sheet_issues: dict[str, str] = {}
        for i in scan.issues.sorted():
            if i.severity == INFO:
                continue
            if i.sheet:
                sheet_issues.setdefault(reader.sheet_key(i.file, i.sheet), i.message)
            elif i.file:
                file_issues.setdefault(i.file, i.message)
        refs_by_file: dict[str, list[reader.SheetRef]] = {}
        for r in scan.refs:
            refs_by_file.setdefault(r.rel, []).append(r)

        n_sheets = 0
        for rel in scan.files:
            top = QTreeWidgetItem([rel])
            top.setData(0, Qt.UserRole, rel)
            refs = refs_by_file.get(rel, [])
            if not refs:
                top.setText(self.COL_STATUS, file_issues.get(rel, "无可读取的 Sheet"))
                top.setForeground(self.COL_STATUS, QBrush(QColor(SEVERITY_COLORS[ERROR])))
                top.setFlags(top.flags() & ~Qt.ItemIsUserCheckable)
                self.tree.addTopLevelItem(top)
                continue
            top.setFlags(top.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsAutoTristate)
            for ref in refs:
                n_sheets += 1
                child = QTreeWidgetItem([ref.sheet])
                child.setData(0, Qt.UserRole, ref.key)
                child.setFlags(child.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEditable)
                child.setCheckState(0, Qt.Unchecked if cfg.is_excluded(ref.rel, ref.sheet)
                                    else Qt.Checked)
                plan = scan.plan_by_key.get(ref.key)
                if plan is not None:
                    child.setText(self.COL_HEADER, str(plan.table.header_row))
                    child.setText(self.COL_ROWS, str(len(plan.table.rows)))
                    status = sheet_issues.get(ref.key, "正常")
                    if ref.key in cfg.header_rows:
                        status = f"表头行已手动指定（自动识别为第 {plan.table.auto_header_row} 行）"
                    child.setText(self.COL_STATUS, status)
                else:
                    child.setText(self.COL_STATUS, sheet_issues.get(ref.key, "无法读取"))
                    child.setForeground(self.COL_STATUS, QBrush(QColor(SEVERITY_COLORS[WARNING])))
                top.addChild(child)
            self.tree.addTopLevelItem(top)
            top.setExpanded(True)
        self._loading = False
        bad = sum(1 for rel in scan.files if rel not in refs_by_file)
        self.summary.setText(f"共 {len(scan.files)} 个文件、{n_sheets} 个 Sheet"
                             + (f"，其中 {bad} 个文件无法读取（运行后会列入问题清单）" if bad else "")
                             + "。勾选要参与汇总的 Sheet；双击“表头行”可手动修改。")

    def _item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if self._loading or column != self.COL_HEADER:
            return
        text = item.text(self.COL_HEADER).strip()
        if not text.isdigit() or int(text) < 1:
            QMessageBox.warning(self, APP_TITLE, "表头行请填写从 1 开始的行号")
            plan = self.scan.plan_by_key.get(item.data(0, Qt.UserRole)) if self.scan else None
            self._loading = True
            item.setText(self.COL_HEADER, str(plan.table.header_row) if plan else "")
            self._loading = False

    def sheet_items(self) -> list[QTreeWidgetItem]:
        out = []
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            out += [top.child(j) for j in range(top.childCount())]
        return out


# ====================================================================== 第 2 步
@dataclass
class MappingRow:
    source: str                 # 原表头（第一次出现的写法）
    count: int                  # 出现在多少个 Sheet
    target: str | None
    status: str
    score: float
    candidates: list[tuple[str, float]]
    confirmed: bool = False


class MappingPage(QWidget):
    COL_SRC, COL_COUNT, COL_TARGET, COL_STATUS, COL_SCORE, COL_CAND = range(6)

    def __init__(self, store: AliasStore) -> None:
        super().__init__()
        self.store = store
        self.rows: list[MappingRow] = []
        self.combos: list[QComboBox] = []
        self._loading = False

        legend = QHBoxLayout()
        legend.addWidget(QLabel("颜色说明："))
        for status in (STATUS_AUTO, STATUS_PENDING, STATUS_UNMATCHED, STATUS_MANUAL, STATUS_IGNORED):
            lab = QLabel(f"  {status}  ")
            lab.setStyleSheet(f"background:{STATUS_COLORS[status]}; border:1px solid #BFBFBF;")
            legend.addWidget(lab)
        legend.addStretch(1)
        self.confirm_all_btn = QPushButton("全部“待确认”按建议确认")
        self.confirm_all_btn.clicked.connect(self.confirm_all_pending)
        legend.addWidget(self.confirm_all_btn)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["原表头", "出现次数", "映射到", "状态", "相似度", "其他候选"])
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(self.COL_TARGET, QHeaderView.Interactive)
        header.setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)

        self.accept_pending_box = QCheckBox("未确认的“待确认”项也按建议映射（会在问题清单中提示）")
        self.keep_unmatched_box = QCheckBox("未匹配的列按原表头保留到明细")
        self.learn_box = QCheckBox("把确认/修改过的映射写入别名字典，下次自动识别")
        self.learn_box.setChecked(True)

        tip = QLabel("在“映射到”中选择或直接输入列名即视为确认；选择“（忽略此列）”则不导出该列。")
        tip.setWordWrap(True)
        lay = QVBoxLayout(self)
        lay.addLayout(legend)
        lay.addWidget(tip)
        lay.addWidget(self.table, 1)
        lay.addWidget(self.accept_pending_box)
        lay.addWidget(self.keep_unmatched_box)
        lay.addWidget(self.learn_box)

    def populate(self, plans: list[pipeline.TablePlan], cfg: TaskConfig) -> None:
        self.accept_pending_box.setChecked(cfg.accept_pending)
        self.keep_unmatched_box.setChecked(cfg.keep_unmatched)
        manual = {normalize_header(k) for k in cfg.column_mapping}
        merged: dict[str, MappingRow] = {}
        for p in plans:
            for s in p.suggestions:
                key = normalize_header(s.source)
                if key in merged:
                    merged[key].count += 1
                    continue
                merged[key] = MappingRow(source=s.source, count=1, target=s.target,
                                         status=s.status, score=s.score,
                                         candidates=s.candidates, confirmed=key in manual)
        order = {STATUS_UNMATCHED: 0, STATUS_PENDING: 1, STATUS_MANUAL: 2, STATUS_IGNORED: 3,
                 STATUS_AUTO: 4}
        self.rows = sorted(merged.values(), key=lambda r: (order.get(r.status, 9), -r.count))
        self._render()

    def _render(self) -> None:
        self._loading = True
        self.table.setRowCount(len(self.rows))
        self.combos = []
        names = self.store.standard_names
        for i, r in enumerate(self.rows):
            color = STATUS_COLORS.get(r.status)
            self.table.setItem(i, self.COL_SRC, _colored_item(r.source, color))
            self.table.setItem(i, self.COL_COUNT, _colored_item(r.count, color))
            combo = QComboBox()
            combo.setEditable(True)
            extra = [r.target] if r.target and r.target not in names and r.target != r.source else []
            combo.addItems(names + extra + [KEEP_NAME, IGNORE])
            if r.status == STATUS_IGNORED:
                combo.setCurrentText(IGNORE)
            elif r.target and (r.target != r.source or r.target in names):
                combo.setCurrentText(r.target)
            else:
                combo.setCurrentText(KEEP_NAME)
            combo.currentTextChanged.connect(lambda _t, row=i: self._changed(row))
            # 重新选择同一个建议值也算确认
            combo.activated.connect(lambda _i, row=i: self._changed(row))
            self.table.setCellWidget(i, self.COL_TARGET, combo)
            self.combos.append(combo)
            self.table.setItem(i, self.COL_STATUS, _colored_item(r.status, color))
            self.table.setItem(i, self.COL_SCORE, _colored_item(f"{r.score:.0f}", color))
            cands = "、".join(f"{n}({s:.0f})" for n, s in r.candidates if n != r.target)
            self.table.setItem(i, self.COL_CAND, _colored_item(cands, color))
        self.table.setColumnWidth(self.COL_TARGET, 180)
        self._loading = False

    def _set_status(self, i: int, status: str) -> None:
        r = self.rows[i]
        r.status = status
        color = STATUS_COLORS.get(status)
        for col in (self.COL_SRC, self.COL_COUNT, self.COL_STATUS, self.COL_SCORE, self.COL_CAND):
            item = self.table.item(i, col)
            if item is not None:
                item.setBackground(QBrush(QColor(color)))
        self.table.item(i, self.COL_STATUS).setText(status)

    def _changed(self, i: int) -> None:
        if self._loading:
            return
        self.rows[i].confirmed = True
        text = self.combos[i].currentText().strip()
        self._set_status(i, STATUS_IGNORED if text == IGNORE else STATUS_MANUAL)

    def confirm_all_pending(self) -> None:
        for i, r in enumerate(self.rows):
            if r.status == STATUS_PENDING:
                r.confirmed = True
                self._set_status(i, STATUS_MANUAL)

    def save_config(self, cfg: TaskConfig) -> list[tuple[str, str]]:
        """写回配置。返回需要学习的 (原表头, 目标列) 列表。"""
        cfg.accept_pending = self.accept_pending_box.isChecked()
        cfg.keep_unmatched = self.keep_unmatched_box.isChecked()
        if not self.rows:
            return []
        # 当前数据中没出现的表头，保留原有的手动映射
        shown = {normalize_header(r.source) for r in self.rows}
        mapping = {k: v for k, v in cfg.column_mapping.items() if normalize_header(k) not in shown}
        learn = []
        for r, combo in zip(self.rows, self.combos):
            if not r.confirmed:
                continue
            text = combo.currentText().strip()
            if text == IGNORE:
                mapping[r.source] = ""
            elif text in (KEEP_NAME, "") or text == r.source:
                mapping[r.source] = r.source
            else:
                mapping[r.source] = text
                learn.append((r.source, text))
        cfg.column_mapping = mapping
        return learn if self.learn_box.isChecked() else []


# ====================================================================== 第 3 步
class SettingsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.columns: list[str] = []

        self.group_list = QListWidget()
        self.group_list.setDragDropMode(QAbstractItemView.InternalMove)
        group_box = QGroupBox("分组列（可多选，可拖动调整顺序）")
        self.group_input = QComboBox()
        self.group_input.setEditable(True)
        self.group_input.setInsertPolicy(QComboBox.NoInsert)
        self.group_input.lineEdit().setPlaceholderText("输入或选择列名，如 部门、客户、项目…")
        self.group_input.lineEdit().returnPressed.connect(lambda: self.add_group())
        self.group_add_btn = QPushButton("添加分组列")
        self.group_add_btn.clicked.connect(lambda: self.add_group())
        add_row = QHBoxLayout()
        add_row.addWidget(self.group_input, 1)
        add_row.addWidget(self.group_add_btn)
        group_lay = QVBoxLayout(group_box)
        group_lay.addWidget(self.group_list)
        group_lay.addLayout(add_row)

        self.agg_table = QTableWidget(0, 2)
        self.agg_table.setHorizontalHeaderLabels(["汇总列", "汇总方式"])
        self.agg_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.agg_table.verticalHeader().setVisible(False)
        add_btn, del_btn = QPushButton("添加"), QPushButton("删除选中")
        add_btn.clicked.connect(lambda: self.add_agg())
        del_btn.clicked.connect(self._del_agg)
        agg_box = QGroupBox("汇总列")
        agg_lay = QVBoxLayout(agg_box)
        agg_lay.addWidget(self.agg_table)
        self.count_box = QCheckBox("附加“记录数”列（每组有多少条记录）")
        btns = QHBoxLayout()
        btns.addWidget(add_btn)
        btns.addWidget(del_btn)
        btns.addStretch(1)
        agg_lay.addLayout(btns)
        agg_lay.addWidget(self.count_box)

        self.dedup_combo = QComboBox()
        for label, value in DEDUP_LABELS:
            self.dedup_combo.addItem(label, value)
        self.dedup_list = QListWidget()
        self.dedup_list.setMaximumHeight(120)
        dedup_box = QGroupBox("重复记录")
        dl = QFormLayout(dedup_box)
        dl.addRow("处理方式：", self.dedup_combo)
        dl.addRow("判断依据（不勾选 = 全部列）：", self.dedup_list)

        self.out_dir = QLineEdit()
        self.out_dir.setPlaceholderText("留空 = 与输入文件夹相同")
        out_browse = QPushButton("浏览…")
        out_browse.clicked.connect(self._browse_out)
        self.out_name = QLineEdit()
        self.out_name.setPlaceholderText("留空 = 汇总结果_年月日_时分秒.xlsx")
        out_box = QGroupBox("输出")
        ol = QFormLayout(out_box)
        d_row = QHBoxLayout()
        d_row.addWidget(self.out_dir, 1)
        d_row.addWidget(out_browse)
        ol.addRow("保存到：", d_row)
        ol.addRow("文件名：", self.out_name)

        top = QHBoxLayout()
        top.addWidget(group_box, 1)
        top.addWidget(agg_box, 1)
        lay = QVBoxLayout(self)
        lay.addLayout(top, 1)
        lay.addWidget(dedup_box)
        lay.addWidget(out_box)

    def _browse_out(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择保存位置", self.out_dir.text())
        if folder:
            self.out_dir.setText(folder)

    def add_agg(self, column: str | None = None, func: str = "求和") -> None:
        row = self.agg_table.rowCount()
        self.agg_table.insertRow(row)
        col_combo, func_combo = QComboBox(), QComboBox()
        col_combo.addItems(self.columns)
        if column and column not in self.columns:
            col_combo.addItem(column)
        if column:
            col_combo.setCurrentText(column)
        func_combo.addItems(AGG_FUNC_NAMES)
        func_combo.setCurrentText(aggregator.normalize_agg_func(func) or "求和")
        self.agg_table.setCellWidget(row, 0, col_combo)
        self.agg_table.setCellWidget(row, 1, func_combo)

    def _del_agg(self) -> None:
        rows = sorted({i.row() for i in self.agg_table.selectedIndexes()}, reverse=True)
        if not rows and self.agg_table.rowCount():
            rows = [self.agg_table.rowCount() - 1]
        for r in rows:
            self.agg_table.removeRow(r)

    @staticmethod
    def _make_item(name: str, checked: bool, exists: bool) -> QListWidgetItem:
        item = QListWidgetItem(name if exists else f"{name}")
        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        item.setData(Qt.UserRole, name)
        if not exists:
            # 当前数据中没有这一列（自定义添加或来自旧配置）：灰色显示并说明
            item.setForeground(QBrush(QColor("#9E9E9E")))
            item.setToolTip("当前数据中没有这一列；如果运行时仍然没有，将被忽略并在问题清单中提示")
        return item

    @classmethod
    def _fill_checklist(cls, widget: QListWidget, columns: list[str], checked: list[str]) -> None:
        widget.clear()
        ordered = [c for c in checked if c in columns] + [c for c in columns if c not in checked]
        ordered += [c for c in checked if c not in columns]   # 配置里有、当前数据里没有的也保留
        for c in ordered:
            # 还没有扫描结果（如刚加载配置）时无法判断列是否存在，不标记为缺失
            widget.addItem(cls._make_item(c, c in checked, not columns or c in columns))

    def add_group(self, name: str | None = None) -> QListWidgetItem | None:
        """添加一个自定义分组列（已在列表中则直接勾选）。"""
        name = (name if name is not None else self.group_input.currentText()).strip()
        if not name:
            return None
        for i in range(self.group_list.count()):
            item = self.group_list.item(i)
            if item.data(Qt.UserRole) == name:
                item.setCheckState(Qt.Checked)
                self.group_list.setCurrentItem(item)
                self.group_input.setEditText("")
                return item
        item = self._make_item(name, True, name in self.columns)
        self.group_list.addItem(item)
        self.group_list.setCurrentItem(item)
        self.group_input.setEditText("")
        return item

    def missing_group_columns(self) -> list[str]:
        """已勾选、但当前数据中不存在的分组列。"""
        return [c for c in self._checked(self.group_list) if self.columns and c not in self.columns]

    @staticmethod
    def _checked(widget: QListWidget) -> list[str]:
        return [widget.item(i).data(Qt.UserRole) or widget.item(i).text()
                for i in range(widget.count()) if widget.item(i).checkState() == Qt.Checked]

    def populate(self, columns: list[str], cfg: TaskConfig) -> None:
        self.columns = columns
        self._fill_checklist(self.group_list, columns, cfg.group_by)
        self.group_input.clear()
        self.group_input.addItems(columns)
        self.group_input.setEditText("")
        self.agg_table.setRowCount(0)
        aggs = cfg.aggregations or ([AggSpec("金额", "求和")] if "金额" in columns else [])
        for a in aggs:
            self.add_agg(a.column, a.func)
        self.count_box.setChecked(cfg.add_count_column)
        idx = max(0, self.dedup_combo.findData(cfg.dedup_mode))
        self.dedup_combo.setCurrentIndex(idx)
        self._fill_checklist(self.dedup_list, columns, cfg.dedup_columns)
        self.out_dir.setText(cfg.output_dir)
        self.out_name.setText(cfg.output_name)

    def save_config(self, cfg: TaskConfig) -> None:
        if not self.columns and self.group_list.count() == 0:
            return          # 还没进入过这一步，保持配置不变
        cfg.group_by = self._checked(self.group_list)
        aggs = []
        for r in range(self.agg_table.rowCount()):
            col = self.agg_table.cellWidget(r, 0).currentText().strip()
            func = self.agg_table.cellWidget(r, 1).currentText()
            if col:
                aggs.append(AggSpec(col, func))
        cfg.aggregations = aggs
        cfg.add_count_column = self.count_box.isChecked()
        cfg.dedup_mode = self.dedup_combo.currentData()
        cfg.dedup_columns = self._checked(self.dedup_list)
        cfg.output_dir = self.out_dir.text().strip()
        cfg.output_name = self.out_name.text().strip()


# ====================================================================== 第 4 步
class RunPage(QWidget):
    run_requested = Signal()
    cancel_requested = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.output_path: Path | None = None
        self.overview = QPlainTextEdit()
        self.overview.setReadOnly(True)
        self.overview.setMaximumHeight(150)

        self.run_btn = QPushButton("开始运行")
        self.run_btn.setMinimumHeight(36)
        self.run_btn.clicked.connect(self.run_requested)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_requested)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.status = QLabel("")
        self.stats = QLabel("")
        self.stats.setWordWrap(True)
        self.open_file_btn = QPushButton("打开结果文件")
        self.open_dir_btn = QPushButton("打开所在文件夹")
        for b in (self.open_file_btn, self.open_dir_btn):
            b.setEnabled(False)
        self.open_file_btn.clicked.connect(self._open_file)
        self.open_dir_btn.clicked.connect(self._open_dir)

        self.issue_table = QTableWidget(0, 8)
        self.issue_table.setHorizontalHeaderLabels(["严重程度", "文件", "Sheet", "行号", "列名",
                                                    "问题类型", "原始值", "说明"])
        self.issue_table.verticalHeader().setVisible(False)
        self.issue_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.issue_table.horizontalHeader().setStretchLastSection(True)
        self.issue_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        btns = QHBoxLayout()
        btns.addWidget(self.run_btn)
        btns.addWidget(self.cancel_btn)
        btns.addStretch(1)
        btns.addWidget(self.open_file_btn)
        btns.addWidget(self.open_dir_btn)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("任务概要："))
        lay.addWidget(self.overview)
        lay.addLayout(btns)
        lay.addWidget(self.progress)
        lay.addWidget(self.status)
        lay.addWidget(self.stats)
        lay.addWidget(QLabel("问题清单预览（完整内容见结果文件的“问题清单”Sheet）："))
        lay.addWidget(self.issue_table, 1)

    def show_overview(self, cfg: TaskConfig) -> None:
        aggs = "、".join(f"{a.column}({aggregator.normalize_agg_func(a.func) or a.func})"
                        for a in cfg.aggregations) or "（无，仅统计记录数）"
        dedup = dict((v, k) for k, v in DEDUP_LABELS).get(cfg.dedup_mode, cfg.dedup_mode)
        lines = [
            f"输入文件夹：{cfg.input_folder}{'（含子文件夹）' if cfg.recursive else ''}",
            f"排除：{len(cfg.excluded)} 项；手动表头行：{len(cfg.header_rows)} 项；"
            f"手动映射：{len(cfg.column_mapping)} 项",
            f"分组列：{'、'.join(cfg.group_by) or '（不分组，只出总计）'}",
            f"汇总列：{aggs}" + ("，附加记录数" if cfg.add_count_column else ""),
            f"重复记录：{dedup}"
            + (f"（按 {'、'.join(cfg.dedup_columns)}）" if cfg.dedup_columns else ""),
            f"输出文件：{cfg.output_dir or cfg.input_folder}/"
            f"{cfg.output_name or '汇总结果_年月日_时分秒.xlsx'}",
        ]
        self.overview.setPlainText("\n".join(lines))

    def set_running(self, running: bool) -> None:
        self.run_btn.setEnabled(not running)
        self.cancel_btn.setEnabled(running)
        if running:
            self.progress.setValue(0)
            self.stats.setText("")
            self.issue_table.setRowCount(0)
            for b in (self.open_file_btn, self.open_dir_btn):
                b.setEnabled(False)

    def on_progress(self, pct: int, msg: str) -> None:
        self.progress.setValue(pct)
        self.status.setText(msg)

    def show_result(self, result: pipeline.RunResult) -> None:
        c = result.issue_counts
        self.output_path = result.output_path
        text = (f"处理文件 {result.files_ok}/{result.files_total} 个，Sheet {result.sheets_read} 个，"
                f"明细 {result.rows_detail} 行，用时 {result.elapsed:.1f} 秒。\n"
                f"问题：错误 {c[ERROR]} 条，警告 {c[WARNING]} 条，提示 {c[INFO]} 条。")
        if result.output_path:
            text += f"\n结果文件：{result.output_path}"
            self.status.setText("完成")
        else:
            self.status.setText("未能生成结果文件，请查看下方问题清单")
        self.stats.setText(text)
        self.open_file_btn.setEnabled(result.output_path is not None)
        self.open_dir_btn.setEnabled(result.output_path is not None)
        issues = result.issues.sorted()[:ISSUE_PREVIEW_LIMIT]
        self.issue_table.setRowCount(len(issues))
        for r, i in enumerate(issues):
            values = i.to_row()
            for col, v in enumerate(values):
                item = _colored_item(v)
                if col == 0:
                    item.setForeground(QBrush(QColor(SEVERITY_COLORS.get(i.severity, "#000000"))))
                self.issue_table.setItem(r, col, item)

    def _open_file(self) -> None:
        if self.output_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_path)))

    def _open_dir(self) -> None:
        if not self.output_path:
            return
        if sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", str(self.output_path)])
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_path.parent)))


# ====================================================================== 主窗口
class MainWindow(QMainWindow):
    def __init__(self, store: AliasStore | None = None, sync: bool = False) -> None:
        """sync=True 时任务在当前线程同步执行（用于自动化测试）。"""
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} v{__version__}")
        self.resize(1100, 720)
        self.store = store or AliasStore()
        self.sync = sync
        self.config = TaskConfig(dedup_mode=aggregator.DEDUP_MARK)
        self.scan: ScanResult | None = None
        self.last_result: pipeline.RunResult | None = None
        self._thread: QThread | None = None
        self._worker: _Worker | None = None
        self._cancel = threading.Event()
        self._handlers: tuple = ()

        self.page_folder = FolderPage()
        self.page_mapping = MappingPage(self.store)
        self.page_settings = SettingsPage()
        self.page_run = RunPage()
        self.stack = QStackedWidget()
        for p in (self.page_folder, self.page_mapping, self.page_settings, self.page_run):
            self.stack.addWidget(p)

        self.step_labels = [QLabel(t) for t in STEP_TITLES]
        steps = QHBoxLayout()
        for i, lab in enumerate(self.step_labels):
            if i:
                steps.addWidget(QLabel("  ›  "))
            steps.addWidget(lab)
        steps.addStretch(1)

        self.prev_btn, self.next_btn = QPushButton("上一步"), QPushButton("下一步")
        self.prev_btn.clicked.connect(self.go_prev)
        self.next_btn.clicked.connect(self.go_next)
        nav = QHBoxLayout()
        self.busy_label = QLabel("")
        nav.addWidget(self.busy_label, 1)
        nav.addWidget(self.prev_btn)
        nav.addWidget(self.next_btn)

        central = QWidget()
        lay = QVBoxLayout(central)
        lay.addLayout(steps)
        lay.addWidget(self.stack, 1)
        lay.addLayout(nav)
        self.setCentralWidget(central)

        tb = QToolBar("操作")
        tb.setMovable(False)
        self.addToolBar(tb)
        self.act_load = QAction("加载配置", self)
        self.act_save = QAction("保存配置", self)
        self.act_run = QAction("一键运行", self)
        self.act_load.triggered.connect(self.load_config_dialog)
        self.act_save.triggered.connect(self.save_config_dialog)
        self.act_run.triggered.connect(self.one_click_run)
        for a in (self.act_load, self.act_save, self.act_run):
            tb.addAction(a)

        self.page_folder.scan_requested.connect(self.start_scan)
        self.page_run.run_requested.connect(self.start_run)
        self.page_run.cancel_requested.connect(self._cancel.set)
        self._update_nav()

    # ---------------------------------------------------------------- 导航
    def _update_nav(self) -> None:
        idx = self.stack.currentIndex()
        for i, lab in enumerate(self.step_labels):
            lab.setStyleSheet("font-weight:bold; color:#1F4E79;" if i == idx else "color:#808080;")
        busy = self.is_busy()
        self.prev_btn.setEnabled(idx > 0 and not busy)
        self.next_btn.setEnabled(idx < self.stack.count() - 1 and not busy)
        for a in (self.act_load, self.act_save, self.act_run):
            a.setEnabled(not busy)

    def is_busy(self) -> bool:
        return self._thread is not None

    def _goto(self, idx: int) -> None:
        self.stack.setCurrentIndex(idx)
        if idx == 3:
            self.page_run.show_overview(self.config)
        self._update_nav()

    def go_prev(self) -> None:
        self.collect_config()
        self._goto(max(0, self.stack.currentIndex() - 1))

    def go_next(self) -> None:
        idx = self.stack.currentIndex()
        if idx == 0:
            self.page_folder.save_config(self.config)
            if not Path(self.config.input_folder or "").is_dir():
                QMessageBox.warning(self, APP_TITLE, "请先选择一个存在的文件夹")
                return
            if self.scan is None or self.scan.key != _scan_key(self.config):
                self.start_scan(then=self._enter_mapping)
                return
            self._enter_mapping()
        elif idx == 1:
            self._learn(self.page_mapping.save_config(self.config))
            self._enter_settings()
        elif idx == 2:
            missing = self.page_settings.missing_group_columns()
            if missing and QMessageBox.question(
                    self, APP_TITLE,
                    f"以下分组列在当前数据中不存在：{'、'.join(missing)}\n"
                    "运行时会被忽略并在问题清单中提示。仍要继续吗？") != QMessageBox.Yes:
                return
            self.page_settings.save_config(self.config)
            self._goto(3)

    def _included_plans(self) -> list[pipeline.TablePlan]:
        if self.scan is None:
            return []
        plans = [p for p in self.scan.plans
                 if not self.config.is_excluded(p.table.rel, p.table.sheet)]
        return pipeline.remap(plans, self.store, self.config.column_mapping)

    def _enter_mapping(self) -> None:
        self.page_mapping.populate(self._included_plans(), self.config)
        self._goto(1)

    def _enter_settings(self) -> None:
        _usable, columns = pipeline.select_usable(self.config, self.store, self._included_plans())
        self.page_settings.populate(columns, self.config)
        self._goto(2)

    def _learn(self, pairs: list[tuple[str, str]]) -> None:
        for src, tgt in pairs:
            try:
                self.store.learn(src, tgt)
            except OSError as e:
                QMessageBox.warning(self, APP_TITLE, f"别名字典保存失败：{e}")
                return

    def collect_config(self) -> TaskConfig:
        """把当前所在页面的修改写回配置。"""
        idx = self.stack.currentIndex()
        if idx == 0:
            self.page_folder.save_config(self.config)
        elif idx == 1:
            self.page_mapping.save_config(self.config)   # 只保存，不学习（学习在“下一步”时）
        elif idx == 2:
            self.page_settings.save_config(self.config)
        return self.config

    # ---------------------------------------------------------------- 后台任务
    def _start_task(self, fn: Callable, on_done: Callable[[Any], None],
                    on_progress: Callable[[int, str], None] | None = None,
                    on_fail: Callable[[str], None] | None = None) -> None:
        self._cancel.clear()
        on_progress = on_progress or (lambda p, m: self.busy_label.setText(f"{m}（{p}%）"))
        on_fail = on_fail or self._task_failed
        if self.sync:
            try:
                result = fn(on_progress, self._cancel)
            except pipeline.Cancelled:
                self.busy_label.setText("")
                on_fail("已取消")
                return
            except Exception:  # noqa: BLE001
                self.busy_label.setText("")
                on_fail(traceback.format_exc())
                return
            self.busy_label.setText("")
            on_done(result)
            return
        thread = QThread(self)
        worker = _Worker(fn, self._cancel)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        # 连接到主窗口的槽函数，保证回调在界面线程执行
        self._handlers = (on_done, on_fail, on_progress)
        worker.progress.connect(self._on_task_progress)
        worker.done.connect(self._on_task_done)
        worker.failed.connect(self._on_task_failed)
        self._thread, self._worker = thread, worker
        self._update_nav()
        thread.start()

    def _finish_task(self) -> None:
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait()
        self._thread = self._worker = None
        self.busy_label.setText("")
        self._update_nav()

    @Slot(int, str)
    def _on_task_progress(self, pct: int, msg: str) -> None:
        self._handlers[2](pct, msg)

    @Slot(object)
    def _on_task_done(self, result: object) -> None:
        self._finish_task()
        self._handlers[0](result)

    @Slot(str)
    def _on_task_failed(self, message: str) -> None:
        self._finish_task()
        self._handlers[1](message)

    def _task_failed(self, message: str) -> None:
        self.page_run.set_running(False)
        if message == "已取消":
            self.page_run.status.setText("已取消")
            return
        QMessageBox.critical(self, APP_TITLE, f"处理过程中出现意外错误：\n\n{message[-2000:]}")

    def start_scan(self, then: Callable[[], None] | None = None) -> None:
        self.page_folder.save_basic(self.config)
        if not Path(self.config.input_folder or "").is_dir():
            QMessageBox.warning(self, APP_TITLE, "请先选择一个存在的文件夹")
            return
        cfg = copy.deepcopy(self.config)
        cfg.excluded = []
        store = self.store

        def job(progress, cancel):
            issues = IssueCollector()
            refs = pipeline.discover(cfg, issues)
            plans, _ = pipeline.plan_tables(cfg, store, issues, refs, progress, cancel)
            files = [reader.rel_path(p, cfg.input_folder)
                     for p in reader.scan_folder(cfg.input_folder, cfg.recursive)
                     if not p.name.startswith(pipeline.OUTPUT_PREFIX)]
            return ScanResult(files=files, refs=refs, plans=plans, issues=issues,
                              key=_scan_key(cfg))

        def done(scan: ScanResult) -> None:
            self.scan = scan
            self.page_folder.show_scan(scan, self.config)
            if callable(then):
                then()

        self._start_task(job, done)

    def start_run(self) -> None:
        self.collect_config()
        cfg = copy.deepcopy(self.config)
        if not Path(cfg.input_folder or "").is_dir():
            QMessageBox.warning(self, APP_TITLE, "输入文件夹不存在，请回到第 1 步重新选择")
            return
        self.page_run.show_overview(cfg)
        self.page_run.set_running(True)
        store = self.store

        def job(progress, cancel):
            return pipeline.run(cfg, store, progress=progress, cancel=cancel)

        def done(result: pipeline.RunResult) -> None:
            self.last_result = result
            self.page_run.set_running(False)
            self.page_run.show_result(result)

        self._start_task(job, done, on_progress=self.page_run.on_progress)

    def one_click_run(self) -> None:
        self.collect_config()
        self._goto(3)
        self.start_run()

    # ---------------------------------------------------------------- 配置
    def apply_config(self, cfg: TaskConfig) -> None:
        """加载配置到界面。之后可直接“一键运行”，也可逐步检查。"""
        self.config = cfg
        self.scan = None
        self.page_folder.load_config(cfg)
        self.page_folder.tree.clear()
        self.page_mapping.rows = []
        self.page_mapping.table.setRowCount(0)
        self.page_settings.columns = []
        self.page_settings.populate([], cfg)
        self.page_settings.columns = []
        self._goto(0)
        if Path(cfg.input_folder or "").is_dir():
            self.start_scan()

    def load_config_dialog(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "加载配置", "", "任务配置 (*.json)")
        if not path:
            return
        try:
            cfg = TaskConfig.load(path)
        except (OSError, ValueError, TypeError) as e:
            QMessageBox.warning(self, APP_TITLE, f"配置文件无法读取：{e}")
            return
        self.apply_config(cfg)
        self.statusBar().showMessage(f"已加载配置：{path}。可点击“一键运行”直接生成结果。", 10000)

    def save_config_dialog(self) -> None:
        self.collect_config()
        path, _ = QFileDialog.getSaveFileName(self, "保存配置", "汇总任务.json", "任务配置 (*.json)")
        if not path:
            return
        try:
            self.config.save(path)
        except OSError as e:
            QMessageBox.warning(self, APP_TITLE, f"配置保存失败：{e}")
            return
        self.statusBar().showMessage(f"配置已保存：{path}", 10000)

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt 命名
        if self._thread is not None:
            if QMessageBox.question(self, APP_TITLE, "任务正在运行，确定要退出吗？") != QMessageBox.Yes:
                event.ignore()
                return
            self._cancel.set()
            self._thread.quit()
            self._thread.wait(5000)
        event.accept()


def _install_crash_handler() -> None:
    """未捕获的异常写入本地日志并弹窗提示，避免打包后的程序直接闪退。（只写本地文件，不联网）"""
    import datetime as dt

    def handler(exc_type, exc, tb) -> None:
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        log_path = None
        try:
            log_dir = app_data_dir() / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            log_path = log_dir / f"错误日志_{dt.datetime.now():%Y%m%d}.log"
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"==== {dt.datetime.now():%Y-%m-%d %H:%M:%S} v{__version__}\n{text}\n")
        except OSError:
            pass
        if QApplication.instance() is not None:
            QMessageBox.critical(None, APP_TITLE,
                                 "程序遇到意外错误，当前操作未完成。\n"
                                 + (f"错误详情已保存到：\n{log_path}\n" if log_path else "")
                                 + "\n" + text[-1500:])

    sys.excepthook = handler


def main() -> int:
    if sys.platform == "win32":
        # 让任务栏显示本程序的图标，而不是归到 Python 下
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("SheetMerger.App")
        except (AttributeError, OSError):
            pass
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("SheetMerger")
    if APP_ICON.exists():
        app.setWindowIcon(QIcon(str(APP_ICON)))
    _install_crash_handler()
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
