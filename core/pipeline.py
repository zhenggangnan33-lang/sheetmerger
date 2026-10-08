"""完整流程：扫描 → 读取 → 表头映射 → 清洗 → 合并去重 → 汇总 → 导出。

CLI 和 GUI 共用本模块，保证两边结果一致。
"""
from __future__ import annotations

import re
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

import pandas as pd

from . import aggregator, auto_summary, cleaner, exporter, reader
from .header_mapper import (STATUS_AUTO, STATUS_IGNORED, STATUS_MANUAL, STATUS_PENDING,
                            STATUS_UNMATCHED, AliasStore, MappingSuggestion,
                            suggest_for_columns)
from .validator import (ERROR, INFO, WARNING, IssueCollector, T_AGG, T_BAD_DATE,
                        T_BAD_NUMBER, T_MAP_CONFLICT, T_MAP_PENDING, T_MAP_UNMATCHED, T_MISSING_COLUMN,
                        T_NO_STD_COLUMN, T_READ_FAIL, T_EMPTY_VALUE, T_CODE_FIXED,
                        T_DATE_NO_YEAR, T_NAME_MISMATCH)

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


SPLIT_SHEET_LIMIT = 200        # 拆分成 Sheet 的上限，超过时提示改用“每个值一个文件”
SPLIT_SHEET, SPLIT_FILE = "sheet", "file"
_BAD_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                   *(f"LPT{i}" for i in range(1, 10))}


def safe_filename(name: str, used: set[str]) -> str:
    """Windows 文件名：去掉非法字符、避开保留名、不重名（不区分大小写）。"""
    base = _BAD_FILENAME.sub("_", name).strip().rstrip(".")[:100] or "未命名"
    if base.upper() in _RESERVED_NAMES:
        base = f"_{base}"
    title, n = base, 2
    while title.lower() in used:
        title = f"{base}_{n}"
        n += 1
    used.add(title.lower())
    return title


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
    split_dir: Path | None = None          # 按列拆分成文件时的文件夹
    split_count: int = 0                   # 拆分出的 Sheet / 文件数
    auto_plan: "auto_summary.AutoPlan | None" = None   # 没选分组列时的自动汇总方案

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


def _dominant_year(usable, types: dict[str, str]) -> int | None:
    """日期列中写全了年份的日期里，出现最多的年份（用于补全“9.30”这类没写年份的日期）。"""
    counts: dict[int, int] = {}
    for p, targets in usable:
        idx = [i for i, c in enumerate(targets) if c and types.get(c) == cleaner.TYPE_DATE]
        if not idx:
            continue
        for _row, values in p.table.rows:
            for i in idx:
                d, err = cleaner.to_date(values[i])
                if d is not None and not err:
                    counts[d.year] = counts.get(d.year, 0) + 1
    return max(counts, key=counts.get) if counts else None


def _pad_numeric_codes(records: list[list[Any]], numeric_codes: list[tuple[int, int, str, str]],
                       columns: list[str], issues: IssueCollector) -> None:
    """编码存成数字时前导零会丢（000123 → 123）。

    如果同一编码列里以文本保存的纯数字编码大多是 L 位、且带前导零，就把数字来源的短编码补齐到 L 位。
    """
    if not numeric_codes:
        return
    numeric_set = {(r, c) for r, c, _f, _s in numeric_codes}
    for pos in {c for _r, c, _f, _s in numeric_codes}:
        lengths: dict[int, int] = {}
        has_zero = False
        for r, rec in enumerate(records):
            v = rec[pos]
            if (r, pos) in numeric_set or not isinstance(v, str) or not v.isdigit():
                continue
            lengths[len(v)] = lengths.get(len(v), 0) + 1
            has_zero = has_zero or v.startswith("0")
        if not lengths or not has_zero:
            continue
        width = max(lengths, key=lengths.get)
        fixed: dict[tuple[str, str], int] = {}
        for r, c, f, sh in numeric_codes:
            v = records[r][c]
            if c == pos and isinstance(v, str) and v.isdigit() and len(v) < width:
                records[r][c] = v.zfill(width)
                fixed[(f, sh)] = fixed.get((f, sh), 0) + 1
        for (f, sh), cnt in fixed.items():
            issues.add(INFO, T_CODE_FIXED,
                       f"有 {cnt} 个编码存成了数字（前导零丢失），已按 {width} 位补齐，如 123 → {'123'.zfill(width)}",
                       file=f, sheet=sh, column=columns[pos])


NAME_COLUMNS = ("门店", "仓库", "部门", "客户", "供应商")


def _name_from_filename(records: list[list[Any]], columns: list[str],
                        issues: IssueCollector, fix: bool) -> None:
    """文件名是“北仑三号仓”，“仓库”列却全写着“北仑二号仓”，而文件夹里正好有“北仑二号仓”文件
    （套用模板没改）。fix=True 时按文件名改正，否则只提示。records 末三列是来源文件/Sheet/行号。"""
    for col in NAME_COLUMNS:
        if col not in columns:
            continue
        pos = columns.index(col)
        values_by_file: dict[str, set] = {}
        for rec in records:
            if rec[pos] is not None:
                values_by_file.setdefault(rec[-3], set()).add(rec[pos])
        stems = {Path(f).stem: f for f in values_by_file}
        for f, values in values_by_file.items():
            stem = Path(f).stem
            if len(values) != 1:
                continue
            (value,) = values
            if str(value) == stem or str(value) not in stems:
                continue
            if fix:
                for rec in records:
                    if rec[-3] == f and rec[pos] == value:
                        rec[pos] = stem
                msg = (f"文件名是“{stem}”，但“{col}”列全部写的是“{value}”（与文件“{stems[str(value)]}”"
                       f"相同，疑似套用模板没改），已按文件名改为“{stem}”；如以表内为准，请修改原表后重跑")
            else:
                msg = (f"文件名是“{stem}”，但“{col}”列全部写的是“{value}”，与文件“{stems[str(value)]}”"
                       f"相同；按“{col}”汇总时两个文件的数据会合在一起，请核对是否套用模板没改")
            issues.add(WARNING, T_NAME_MISMATCH, msg, file=f, column=col, value=value)


def _make_summary(config, detail: pd.DataFrame, aggs: list[tuple[str, str]],
                  issues: IssueCollector, auto: "auto_summary.AutoPlan | None" = None
                  ) -> pd.DataFrame:
    pivot = (getattr(config, "pivot_column", "") or "").strip()
    if auto is not None and not config.group_by and not pivot:
        return auto_summary.summarize(detail, auto)
    if pivot:
        return aggregator.pivot_summarize(detail, config.group_by, pivot, aggs, issues)
    return aggregator.summarize(detail, config.group_by, aggs, issues,
                                add_count=config.add_count_column)


def _export_split_files(config, out: Path, detail: pd.DataFrame, split_col: str, parts: list,
                        aggs: list[tuple[str, str]], result: RunResult, progress: ProgressFn,
                        cancel: threading.Event | None) -> None:
    """每个值一个文件：写到“结果文件名_按X拆分”文件夹，每个文件含该值的汇总和明细。"""
    folder = out.with_name(f"{out.stem}_按{safe_filename(split_col, set())}拆分")
    folder.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    for n, (value, positions) in enumerate(parts):
        _check(cancel)
        progress(92 + 7 * n // max(len(parts), 1), f"拆分输出 {aggregator.value_label(value)}")
        part = detail.iloc[positions].reset_index(drop=True)
        if (getattr(config, "pivot_column", "") or "").strip() == split_col:
            # 展开列就是拆分列时，单个文件里只有一个值，展开没有意义，改用普通汇总
            part_summary = aggregator.summarize(part, [g for g in config.group_by if g != split_col]
                                                or [split_col], aggs, IssueCollector(),
                                                add_count=config.add_count_column)
        else:
            part_summary = _make_summary(config, part, aggs, IssueCollector(), result.auto_plan)
        name = safe_filename(aggregator.value_label(value), used) + ".xlsx"
        exporter.export_result(folder / name, part_summary, part, None,
                               plain_columns=[aggregator.COL_ROW])
    result.split_dir = folder
    result.split_count = len(parts)


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
    sum_cols = {a.column for a in config.aggregations
                if aggregator.normalize_agg_func(a.func) in ("求和", "平均")}
    global_year = _dominant_year(usable, types) or time.localtime().tm_year
    numeric_codes: list[tuple[int, int, str, str]] = []   # (记录序号, 列位置, 文件, Sheet)
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
        year = _dominant_year([(p, targets)], types) or global_year
        no_year: dict[str, int] = {}
        empty_sum: dict[str, int] = {}
        for excel_row, values in t.rows:
            rec: list[Any] = [None] * len(all_cols)
            for i, pos, ctype, src in plan:
                raw = values[i]
                value, err = cleaner.clean_value(raw, ctype, year)
                if err:
                    target = columns[pos]
                    label = src if src == target else f"{src}（→{target}）"
                    issues.add(ERROR, T_BAD_DATE if ctype == cleaner.TYPE_DATE else T_BAD_NUMBER,
                               f"{err}，该单元格在明细中留空", file=t.rel, sheet=t.sheet,
                               row=excel_row, column=label, value=raw)
                elif ctype == cleaner.TYPE_DATE and value is not None and cleaner.is_month_day(raw):
                    no_year[src] = no_year.get(src, 0) + 1
                elif ctype == cleaner.TYPE_CODE and isinstance(raw, (int, float)) \
                        and not isinstance(raw, bool):
                    numeric_codes.append((len(records), pos, t.rel, t.sheet))
                if value is None and not err and columns[pos] in sum_cols:
                    empty_sum[src] = empty_sum.get(src, 0) + 1
                rec[pos] = value
            rec[-3], rec[-2], rec[-1] = t.rel, t.sheet, excel_row
            records.append(rec)
        for src, cnt in no_year.items():
            issues.add(INFO, T_DATE_NO_YEAR, f"有 {cnt} 个日期没写年份（如 9.30、9月30日），已按 {year} 年处理",
                       file=t.rel, sheet=t.sheet, row=t.header_row, column=src)
        for src, cnt in empty_sum.items():
            issues.add(INFO, T_EMPTY_VALUE, f"有 {cnt} 个空值（未填写，如未盘点），汇总时不按 0 计算",
                       file=t.rel, sheet=t.sheet, row=t.header_row, column=src)
        if t.rows:
            files_with_rows.add(t.rel)
        result.sheets_read += 1
    _pad_numeric_codes(records, numeric_codes, columns, issues)
    _name_from_filename(records, columns, issues, fix=config.name_from_file)

    detail = aggregator.build_detail(records, all_cols)
    progress(82, "检查重复记录")
    detail = aggregator.deduplicate(detail, config.dedup_mode, config.dedup_columns, issues)
    progress(86, "分组汇总")
    aggs = [(a.column, a.func) for a in config.aggregations]
    # 常见设置错误的提醒
    for col, func in aggs:
        if col in detail.columns and len(detail) and detail[col].isna().all():
            issues.add(WARNING, T_AGG, f"汇总列“{col}”在所有数据中都是空的，汇总结果为 0 或空；"
                       "请检查是否选错了汇总列（例如盘点表通常汇总“账面数量”“实盘数量”）", column=col)
    for col in config.group_by:
        if types.get(col) == cleaner.TYPE_NUMBER:
            issues.add(WARNING, T_AGG, f"分组列“{col}”是数值列，按它分组通常没有意义；"
                       "数值列一般放在“汇总列”里求和", column=col)
    # 没设汇总列时：数值列全部求和（单价等比率列除外），与自动汇总的习惯一致
    sources: dict[str, set] = {}
    for p, targets in usable:
        for target, src in zip(targets, p.table.columns):
            if target:
                sources.setdefault(target, set()).add(str(src))
    plan = auto_summary.make_plan(detail, store, types, aggs, sources)
    if not aggs:
        aggs = list(plan.measures)
    if not config.group_by and not (config.pivot_column or "").strip():
        result.auto_plan = plan
    summary = _make_summary(config, detail, aggs, issues, result.auto_plan)
    result.detail, result.summary = detail, summary

    # 按列拆分
    split_col = (config.split_by or "").strip()
    parts: list[tuple] = []
    if split_col:
        if split_col not in detail.columns:
            issues.add(ERROR, T_AGG, f"拆分列“{split_col}”在明细中不存在，未拆分", column=split_col)
        else:
            parts = aggregator.split_positions(detail, split_col)
            if config.split_mode != SPLIT_FILE and len(parts) > SPLIT_SHEET_LIMIT:
                issues.add(WARNING, T_AGG,
                           f"按“{split_col}”有 {len(parts)} 个不同的值，超过 {SPLIT_SHEET_LIMIT} 个，"
                           "未拆分成 Sheet；请改用“每个值一个文件”", column=split_col)
                parts = []
    result.rows_detail = len(detail)
    result.files_ok = len(files_with_rows)

    if export:
        _check(cancel)
        progress(90, "写出结果文件")
        out = config.resolve_output_path()
        extra = []
        if parts and config.split_mode != SPLIT_FILE:
            for value, positions in parts:
                part = detail.iloc[positions]
                rows = [list(r) for r in part.itertuples(index=False, name=None)]
                rows.append(aggregator.subtotal_row(part, split_col, aggs))
                extra.append((aggregator.value_label(value), list(detail.columns), rows))
            result.split_count = len(extra)
        try:
            exporter.export_result(out, summary, detail, issues.sorted(),
                                   plain_columns=[aggregator.COL_ROW], extra_sheets=extra)
            result.output_path = out
            if parts and config.split_mode == SPLIT_FILE:
                _export_split_files(config, out, detail, split_col, parts, aggs, result,
                                    progress, cancel)
        except PermissionError:
            issues.add(ERROR, T_READ_FAIL, f"无法写入结果文件，可能正被 Excel/WPS 打开：{out}")
        except OSError as e:
            issues.add(ERROR, T_READ_FAIL, f"无法写入结果文件：{out}（{e}）")
    result.elapsed = time.perf_counter() - t0
    progress(100, "完成")
    return result
