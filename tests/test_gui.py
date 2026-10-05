"""界面测试：用无界面模式（offscreen）把四步向导完整走一遍。"""
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
QtWidgets = pytest.importorskip("PySide6.QtWidgets", exc_type=ImportError)
from PySide6.QtCore import QCoreApplication, QEventLoop, Qt, QTimer  # noqa: E402

from config.task_config import AggSpec, TaskConfig  # noqa: E402
from core import pipeline  # noqa: E402
from core.header_mapper import STATUS_MANUAL, STATUS_PENDING, AliasStore, suggest  # noqa: E402
from gui.main_window import IGNORE, MainWindow  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


@pytest.fixture
def win(app):
    w = MainWindow(sync=True)
    yield w
    w.close()


def _sheet_item(win, rel, sheet):
    for item in win.page_folder.sheet_items():
        if item.data(0, Qt.UserRole) == f"{rel}|{sheet}":
            return item
    raise KeyError(rel)


def _mapping_row(win, source):
    for i, r in enumerate(win.page_mapping.rows):
        if r.source == source:
            return i
    raise KeyError(source)


def test_wizard_end_to_end(win, data_dir, tmp_path):
    # ① 选择文件夹并扫描
    win.page_folder.folder_edit.setText(str(data_dir))
    win.start_scan()
    assert win.scan is not None
    assert "16_损坏文件.xlsx" in win.scan.files
    _sheet_item(win, "12_多Sheet.xlsx", "说明").setCheckState(0, Qt.Unchecked)
    win.go_next()
    assert win.stack.currentIndex() == 1
    assert win.config.excluded == ["12_多Sheet.xlsx|说明"]

    # ② 表头映射：待确认项显示出来；确认一个，忽略一个
    i = _mapping_row(win, "单价/元")
    assert win.page_mapping.rows[i].status == STATUS_PENDING
    combo = win.page_mapping.combos[i]
    combo.activated.emit(combo.findText("单价"))     # 用户选择了建议值本身 = 确认
    assert win.page_mapping.rows[i].status == STATUS_MANUAL
    j = _mapping_row(win, "会员卡号")
    win.page_mapping.combos[j].setCurrentText(IGNORE)
    win.go_next()
    assert win.config.column_mapping == {"单价/元": "单价", "会员卡号": ""}
    # 确认的映射写回了别名字典
    assert suggest("单价/元", AliasStore()).status == "自动"

    # ③ 汇总设置
    page = win.page_settings
    assert "会员卡号" not in page.columns and "门店" in page.columns
    for k in range(page.group_list.count()):
        if page.group_list.item(k).text() == "门店":
            page.group_list.item(k).setCheckState(Qt.Checked)
    assert page.agg_table.rowCount() == 1            # 默认 金额 求和
    page.add_agg("数量", "平均")
    assert not page.count_box.isChecked()          # 默认不附加记录数
    page.count_box.setChecked(True)
    page.out_dir.setText(str(tmp_path))
    page.out_name.setText("gui.xlsx")
    win.go_next()
    assert win.stack.currentIndex() == 3
    assert win.config.group_by == ["门店"] and win.config.add_count_column is True
    assert [(a.column, a.func) for a in win.config.aggregations] == [("金额", "求和"), ("数量", "平均")]

    # ④ 运行
    win.start_run()
    result = win.last_result
    assert result.output_path == tmp_path / "gui.xlsx" and result.output_path.exists()
    assert win.page_run.open_file_btn.isEnabled()
    assert win.page_run.issue_table.rowCount() > 0
    assert win.page_run.issue_table.item(0, 0).text() == "错误"

    # 与命令行/核心流程用同一份配置运行，结果一致
    direct = pipeline.run(win.config, export=False)
    assert direct.summary.equals(result.summary)
    assert direct.detail.equals(result.detail)


def test_header_row_edit_triggers_rescan(win, data_dir):
    win.page_folder.folder_edit.setText(str(data_dir))
    win.start_scan()
    item = _sheet_item(win, "06_标题行和合计行.xlsx", "Sheet1")
    assert item.text(1) == "4"
    item.setText(1, "5")
    win.go_next()
    assert win.config.header_rows == {"06_标题行和合计行.xlsx|Sheet1": 5}
    assert win.scan.plan_by_key["06_标题行和合计行.xlsx|Sheet1"].table.header_row == 5


def test_save_load_one_click_same_result(win, data_dir, tmp_path, app):
    cfg = TaskConfig(input_folder=str(data_dir), group_by=["门店"],
                     aggregations=[AggSpec("金额", "求和")], column_mapping={"单价/元": "单价"},
                     excluded=["12_多Sheet.xlsx|说明"], output_dir=str(tmp_path),
                     output_name="a.xlsx")
    path = cfg.save(tmp_path / "task.json")

    win.apply_config(TaskConfig.load(path))
    assert win.page_folder.folder_edit.text() == str(data_dir)
    win.one_click_run()
    first = win.last_result
    assert first.output_path == tmp_path / "a.xlsx"
    assert win.config == cfg            # 一键运行不会改动加载的配置

    # 保存当前配置后再加载、再运行，结果一致
    win.config.save(tmp_path / "task2.json")
    w2 = MainWindow(sync=True)
    w2.apply_config(TaskConfig.load(tmp_path / "task2.json"))
    w2.one_click_run()
    second = w2.last_result
    w2.close()
    assert first.summary.equals(second.summary) and first.detail.equals(second.detail)
    assert first.rows_detail == 125


def test_background_thread_run(app, data_dir, tmp_path):
    """真正的后台线程模式：界面在运行期间可响应，结束后按钮恢复。"""
    w = MainWindow(sync=False)
    w.apply_config(TaskConfig(input_folder=str(data_dir), group_by=["门店"],
                              aggregations=[AggSpec("金额")], output_dir=str(tmp_path),
                              output_name="bg.xlsx"))
    loop = QEventLoop()
    QTimer.singleShot(20000, loop.quit)

    def wait_idle():
        while w.is_busy():
            QCoreApplication.processEvents(QEventLoop.AllEvents, 50)

    wait_idle()                          # 等加载配置后的自动扫描结束
    assert w.scan is not None
    w.one_click_run()
    assert w.is_busy() and not w.act_run.isEnabled()
    wait_idle()
    assert w.last_result is not None and (tmp_path / "bg.xlsx").exists()
    assert w.act_run.isEnabled() and w.page_run.run_btn.isEnabled()
    w.close()


def test_crash_handler_writes_local_log(app, isolated_home, monkeypatch):
    import sys

    from gui import main_window
    shown = []
    monkeypatch.setattr(main_window.QMessageBox, "critical", lambda *a: shown.append(a[2]))
    old = sys.excepthook
    try:
        main_window._install_crash_handler()
        try:
            raise ValueError("测试异常")
        except ValueError:
            sys.excepthook(*sys.exc_info())
    finally:
        sys.excepthook = old
    logs = list((isolated_home / "logs").glob("*.log"))
    assert len(logs) == 1 and "测试异常" in logs[0].read_text(encoding="utf-8")
    assert shown and "测试异常" in shown[0]


def test_custom_group_column(win, data_dir, monkeypatch):
    win.page_folder.folder_edit.setText(str(data_dir))
    win.start_scan()
    win.go_next()
    win.go_next()
    page = win.page_settings
    # 输入数据中已有的列：直接勾选，不重复添加
    count = page.group_list.count()
    page.group_input.setEditText("经手人")
    page.add_group()
    assert page.group_list.count() == count and page.missing_group_columns() == []
    # 自定义一个当前数据里没有的列：加入并勾选，标记为缺失
    page.add_group("项目")
    assert page.group_list.count() == count + 1
    assert page.missing_group_columns() == ["项目"]
    # 下一步时提示确认；选“否”停留在第 3 步
    monkeypatch.setattr(QtWidgets.QMessageBox, "question",
                        lambda *a, **k: QtWidgets.QMessageBox.No)
    win.go_next()
    assert win.stack.currentIndex() == 2
    monkeypatch.setattr(QtWidgets.QMessageBox, "question",
                        lambda *a, **k: QtWidgets.QMessageBox.Yes)
    win.go_next()
    assert win.stack.currentIndex() == 3
    assert win.config.group_by == ["经手人", "项目"]
    # 重新进入第 3 步，自定义列仍在
    win.go_prev()
    assert "项目" in [page.group_list.item(i).text() for i in range(page.group_list.count())]
