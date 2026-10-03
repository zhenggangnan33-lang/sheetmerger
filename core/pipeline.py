"""完整流程：扫描 → 读取 → 表头映射 → 清洗 → 合并去重 → 汇总 → 导出。

CLI 和 GUI 共用本模块，保证两边结果一致。
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

import pandas as pd

from . import aggregator, cleaner, exporter, reader
from .header_mapper import (STATUS_AUTO, STATUS_IGNORED, STATUS_MANUAL, STATUS_PENDING,
                            STATUS_UNMATCHED, AliasStore, MappingSuggestion,
                            suggest_for_columns)
from .validator import (ERROR, INFO, WARNING, IssueCollector, T_BAD_DATE, T_BAD_NUMBER,
                        T_MAP_CONFLICT, T_MAP_PENDING, T_MAP_UNMATCHED, T_MISSING_COLUMN,
                        T_NO_STD_COLUMN, T_READ_FAIL)

OUTPUT_PREFIX = "汇总结果_"     # 本工具生成的结果文件，扫描时跳过，避免重复汇总

# progress(百分比 0-100, 说明文字)
ProgressFn = Callable[[int, str], None]


class Cancelled(Exception):
    """用户取消运行。"""


@dataclass
class TablePlan:
    """一张表的读取结果 + 映射方案（界面第 2 步展示用）。"""
    table: reader.SheetTable
    suggestions: list[MappingSuggestion]
    conflicts: list[str] = field(default_factory=list)

    def resolved(self, accept_pending: bool, keep_unmatched: bool) -> list[str | None]:
        """每列最终写入明细的列名，None 表示丢弃该列。"""
        out: list[str | None] = []
        for s in self.suggestions:
            if s.status in (STATUS_AUTO, STATUS_MANUAL):
                out.append(s.target)
            elif s.status == STATUS_PENDING:
                out.append(s.target if accept_pending else (s.source if keep_unmatched else None))
            elif s.status == STATUS_UNMATCHED:
                out.append(s.source if keep_unmatched else None)
            else:
                out.append(None)
        return out

    @property
    def standard_hits(self) -> int:
        return sum(1 for s in self.suggestions
                   if s.target and s.status in (STATUS_AUTO, STATUS_MANUAL, STATUS_PENDING))


@dataclass
class RunResult:
    output_path: Path | None
    files_total: int = 0
    files_ok: int = 0
    sheets_read: int = 0
    rows_detail: int = 0
    summary: pd.DataFrame | None = None
    detail: pd.DataFrame | None = None
    issues: IssueCollector = field(default_factory=IssueCollector)
    elapsed: float = 0.0

    @property
    def issue_counts(self) -> dict[str, int]:
        return {s: self.issues.count(s) for s in (ERROR, WARNING, INFO)}


def _noop(_p: int, _m: str) -> None:
    pass


def _check(cancel: threading.Event | None) -> None:
    if cancel is not None and cancel.is_set():
        raise Cancelled()


def discover(config, issues: IssueCollector) -> list[reader.SheetRef]:
    """扫描文件夹，跳过本工具生成的结果文件。"""
    refs = reader.discover(config.input_folder, config.recursive, issues)
    kept = []
    for r in refs:
        if Path(r.rel).name.startswith(OUTPUT_PREFIX):
            continue
        kept.append(r)
    return kept


def count_files(config) -> int:
    """扫描到的表格文件数（含打不开的文件，不含本工具生成的结果文件）。"""
    return sum(1 for p in reader.scan_folder(config.input_folder, config.recursive)
               if not p.name.startswith(OUTPUT_PREFIX))


def plan_tables(config, store: AliasStore, issues: IssueCollector,
                refs: list[reader.SheetRef] | None = None,
                progress: ProgressFn = _noop, cancel: threading.Event | None = None,
                progress_span: tuple[int, int] = (0, 100)) -> tuple[list[TablePlan], int]:
    """读取所有未排除的 Sheet 并给出映射建议。返回 (方案列表, 涉及文件数)。"""
    if refs is None:
        refs = discover(config, issues)
    refs = [r for r in refs if not config.is_excluded(r.rel, r.sheet)]
    lo, hi = progress_span
    plans: list[TablePlan] = []
    for n, ref in enumerate(refs):
        _check(cancel)
        progress(lo + (hi - lo) * n // max(len(refs), 1), f"读取 {ref.rel} / {ref.sheet}")
        table = reader.load_sheet(ref, issues, config.header_rows.get(ref.key))
        if table is None:
            continue
        sugg, conflicts = suggest_for_columns(table.columns, store, config.column_mapping)
        plans.append(TablePlan(table=table, suggestions=sugg, conflicts=conflicts))
    return plans, len({r.rel for r in refs})


def _column_order(plans: list[TablePlan], store: AliasStore, resolved: list[list[str | None]]
                  ) -> list[str]:
    seen: list[str] = []
    for cols in resolved:
        for c in cols:
            if c and c not in seen:
                seen.append(c)
    standard = [c for c in store.standard_names if c in seen]
    return standard + [c for c in seen if c not in standard]


def _column_types(config, store: AliasStore, columns: list[str]) -> dict[str, str]:
    types: dict[str, str] = {}
    sum_like = {a.column for a in config.aggregations
                if aggregator.normalize_agg_func(a.func) in ("求和", "平均")}
    for c in columns:
        t = config.column_types.get(c) or store.column_type(c)
        if t is None and c in sum_like:
            t = cleaner.TYPE_NUMBER
        types[c] = t or cleaner.TYPE_AUTO
    return types


def remap(plans: list[TablePlan], store: AliasStore,
          column_mapping: dict[str, str] | None) -> list[TablePlan]:
    """表头映射修改后重新生成映射建议（不重新读文件）。"""
    out = []
    for p in plans:
        sugg, conflicts = suggest_for_columns(p.table.columns, store, column_mapping)
        out.append(TablePlan(table=p.table, suggestions=sugg, conflicts=conflicts))
    return out


def select_usable(config, store: AliasStore, plans: list[TablePlan],
                  issues: IssueCollector | None = None
                  ) -> tuple[list[tuple[TablePlan, list[str | None]]], list[str]]:
    """决定哪些表参与合并、每列写入明细的列名，并登记映射相关问题。

    返回 ([(方案, 每列目标列名)], 明细列顺序)。界面第 3 步也用它得到可选列。
    """
    issues = issues if issues is not None else IssueCollector()
    any_standard = any(p.standard_hits for p in plans)
    usable: list[tuple[TablePlan, list[str | None]]] = []
    for p in plans:
        t = p.table
        for msg in p.conflicts:
            issues.add(WARNING, T_MAP_CONFLICT, msg, file=t.rel, sheet=t.sheet, row=t.header_row)
        if any_standard and p.standard_hits == 0:
            issues.add(WARNING, T_NO_STD_COLUMN,
                       "该 Sheet 没有任何列能对应到标准列，已跳过；如需汇总请在表头映射中手动指定",
                       file=t.rel, sheet=t.sheet, row=t.header_row)
            continue
        for s in p.suggestions:
            if s.status == STATUS_PENDING:
                how = "已采用" if config.accept_pending else "未采用，按原表头处理"
                issues.add(WARNING, T_MAP_PENDING,
                           f"表头“{s.source}”疑似“{s.target}”（相似度 {s.score:.0f}），{how}，请确认",
                           file=t.rel, sheet=t.sheet, row=t.header_row, column=s.source)
            elif s.status == STATUS_UNMATCHED:
                how = "按原表头保留" if config.keep_unmatched else "已丢弃"
                issues.add(INFO, T_MAP_UNMATCHED, f"表头“{s.source}”未匹配到标准列，{how}",
                           file=t.rel, sheet=t.sheet, row=t.header_row, column=s.source)
        usable.append((p, p.resolved(config.accept_pending, config.keep_unmatched)))
    columns = _column_order([u[0] for u in usable], store, [u[1] for u in usable])
    return usable, columns


def run(config, store: AliasStore | None = None, progress: ProgressFn = _noop,
        cancel: threading.Event | None = None, export: bool = True) -> RunResult:
    """执行完整任务。除取消外不会抛出异常，所有问题记入 result.issues。"""
    t0 = time.perf_counter()
    store = store or AliasStore()
    issues = IssueCollector()
    result = RunResult(output_path=None, issues=issues)

    progress(0, "扫描文件夹")
    if not Path(config.input_folder or "").is_dir():
        issues.add(ERROR, T_READ_FAIL, f"输入文件夹不存在：{config.input_folder}")
        result.summary = aggregator.summarize(pd.DataFrame(columns=aggregator.SOURCE_COLUMNS),
                                              [], [], issues)
        result.detail = pd.DataFrame(columns=aggregator.SOURCE_COLUMNS)
        result.elapsed = time.perf_counter() - t0
        return result
    refs = discover(config, issues)
    plans, _ = plan_tables(config, store, issues, refs, progress, cancel, (2, 50))
    result.files_total = count_files(config)

    usable, columns = select_usable(config, store, plans, issues)
    types = _column_types(config, store, columns)
    # "缺少列"只针对多数 Sheet 都有的标准列提示，避免个别表多出的列在其他表里刷屏
    with_rows = [set(filter(None, cols)) for p, cols in usable if p.table.rows]
    common = [c for c in columns if c in store.specs
              and sum(c in s for s in with_rows) * 2 > len(with_rows)]

    # ---- 清洗并合并
    progress(50, "清洗数据")
    all_cols = columns + aggregator.SOURCE_COLUMNS
    col_pos = {c: i for i, c in enumerate(columns)}
    records: list[list[Any]] = []
    files_with_rows: set[str] = set()
    for n, (p, targets) in enumerate(usable):
        _check(cancel)
        t = p.table
        progress(50 + 30 * n // max(len(usable), 1), f"清洗 {t.rel} / {t.sheet}")
        present = {c for c in targets if c}
        missing = [c for c in common if c not in present]
        if missing:
            issues.add(INFO, T_MISSING_COLUMN, f"该 Sheet 缺少列：{'、'.join(missing)}（明细中留空）",
                       file=t.rel, sheet=t.sheet, row=t.header_row)
        plan = [(i, col_pos[c], types[c], src) for i, (c, src)
                in enumerate(zip(targets, t.columns)) if c]
        for excel_row, values in t.rows:
            rec: list[Any] = [None] * len(all_cols)
            for i, pos, ctype, src in plan:
                value, err = cleaner.clean_value(values[i], ctype)
                if err:
                    target = columns[pos]
                    label = src if src == target else f"{src}（→{target}）"
                    issues.add(ERROR, T_BAD_DATE if ctype == cleaner.TYPE_DATE else T_BAD_NUMBER,
                               f"{err}，该单元格在明细中留空", file=t.rel, sheet=t.sheet,
                               row=excel_row, column=label, value=values[i])
                rec[pos] = value
            rec[-3], rec[-2], rec[-1] = t.rel, t.sheet, excel_row
            records.append(rec)
        if t.rows:
            files_with_rows.add(t.rel)
        result.sheets_read += 1

    detail = aggregator.build_detail(records, all_cols)
    progress(82, "检查重复记录")
    detail = aggregator.deduplicate(detail, config.dedup_mode, config.dedup_columns, issues)
    progress(86, "分组汇总")
    summary = aggregator.summarize(detail, config.group_by,
                                   [(a.column, a.func) for a in config.aggregations], issues)
    result.detail, result.summary = detail, summary
    result.rows_detail = len(detail)
    result.files_ok = len(files_with_rows)

    if export:
        _check(cancel)
        progress(90, "写出结果文件")
        out = config.resolve_output_path()
        try:
            exporter.export_result(out, summary, detail, issues.sorted(),
                                   plain_columns=[aggregator.COL_ROW])
            result.output_path = out
        except PermissionError:
            issues.add(ERROR, T_READ_FAIL, f"无法写入结果文件，可能正被 Excel/WPS 打开：{out}")
        except OSError as e:
            issues.add(ERROR, T_READ_FAIL, f"无法写入结果文件：{out}（{e}）")
    result.elapsed = time.perf_counter() - t0
    progress(100, "完成")
    return result
