"""汇总：合并明细（附加来源列）、去重、分组汇总。"""
from __future__ import annotations

import datetime as dt
import re
from typing import Any

import pandas as pd

from .validator import (ERROR, WARNING, IssueCollector, T_AGG, T_DUPLICATE,
                        T_DUPLICATE_DROPPED)

COL_FILE = "来源文件"
COL_SHEET = "来源Sheet"
COL_ROW = "原始行号"
SOURCE_COLUMNS = [COL_FILE, COL_SHEET, COL_ROW]

COL_COUNT = "记录数"
TOTAL_LABEL = "总计"
EMPTY_GROUP_LABEL = "(空)"

# 汇总方式：中文名 -> pandas 方法；也接受英文写法
AGG_FUNCS = {"求和": "sum", "计数": "count", "平均": "mean", "最大": "max", "最小": "min"}
AGG_ALIASES = {"sum": "求和", "count": "计数", "mean": "平均", "avg": "平均", "average": "平均",
               "max": "最大", "min": "最小", "合计": "求和", "平均值": "平均",
               "最大值": "最大", "最小值": "最小"}

DEDUP_OFF = "off"      # 不检查重复
DEDUP_MARK = "mark"    # 只在问题清单中标记，保留全部记录
DEDUP_DROP = "drop"    # 删除重复记录（保留第一次出现的），并在问题清单中说明
DEDUP_MODES = (DEDUP_OFF, DEDUP_MARK, DEDUP_DROP)


def normalize_agg_func(name: str) -> str | None:
    name = (name or "").strip()
    if name in AGG_FUNCS:
        return name
    return AGG_ALIASES.get(name.lower())


def build_detail(records: list[list[Any]], columns: list[str]) -> pd.DataFrame:
    """records 中每行已按 columns（含来源列）对齐。统一用 object 类型保留日期和空值。"""
    return pd.DataFrame(records, columns=columns, dtype=object)


def deduplicate(df: pd.DataFrame, mode: str, key_columns: list[str] | None,
                issues: IssueCollector) -> pd.DataFrame:
    """按全部业务列（或指定列）查找重复记录。"""
    if mode == DEDUP_OFF or df.empty:
        return df
    business = [c for c in df.columns if c not in SOURCE_COLUMNS]
    keys = [c for c in (key_columns or []) if c in df.columns]
    missing = [c for c in (key_columns or []) if c not in df.columns]
    if missing:
        issues.add(WARNING, T_AGG, f"去重列不存在，已忽略：{'、'.join(missing)}")
    if not keys:
        keys = business
    if not keys:
        return df

    first_seen: dict[tuple, int] = {}
    dup_index: list[int] = []
    key_values = df[keys].itertuples(index=False, name=None)
    files, sheets, rows = df[COL_FILE].tolist(), df[COL_SHEET].tolist(), df[COL_ROW].tolist()
    key_desc = "全部列" if keys == business else "、".join(keys)
    for pos, key in enumerate(key_values):
        key = tuple(_hashable(v) for v in key)
        first = first_seen.get(key)
        if first is None:
            first_seen[key] = pos
            continue
        dup_index.append(pos)
        where = f"{files[first]} / {sheets[first]} / 第 {rows[first]} 行"
        if mode == DEDUP_DROP:
            issues.add(WARNING, T_DUPLICATE_DROPPED,
                       f"按{key_desc}与 {where} 重复，已从明细和汇总中删除",
                       file=files[pos], sheet=sheets[pos], row=rows[pos])
        else:
            issues.add(WARNING, T_DUPLICATE,
                       f"按{key_desc}与 {where} 重复（已保留，仍计入汇总）",
                       file=files[pos], sheet=sheets[pos], row=rows[pos])
    if mode == DEDUP_DROP and dup_index:
        df = df.drop(df.index[dup_index]).reset_index(drop=True)
    return df


def _hashable(v: Any) -> Any:
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


_CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
              "六": 6, "七": 7, "八": 8, "九": 9}
_CN_UNITS = {"十": 10, "百": 100, "千": 1000}
_NUM_RUN = re.compile(r"\d+|[零〇一二两三四五六七八九十百千]+")


def _cn_number(s: str) -> int | None:
    """中文数字转整数（支持到千位，如 十二、二十、一百零五）；不是数字返回 None。"""
    total, num = 0, 0
    for ch in s:
        if ch in _CN_DIGITS:
            num = _CN_DIGITS[ch]
        elif ch in _CN_UNITS:
            total += (num or 1) * _CN_UNITS[ch]
            num = 0
        else:
            return None
    return total + num


def _natural_key(s: str) -> tuple:
    """自然排序：文字中的阿拉伯数字和中文数字按数值比较（一店 < 二店 < 十二店，2号 < 10号）。"""
    parts: list = []
    pos = 0
    for m in _NUM_RUN.finditer(s):
        if m.start() > pos:
            parts.append((1, s[pos:m.start()], 0))
        run = m.group()
        n = int(run) if run.isdigit() else _cn_number(run)
        parts.append((0, "", n) if n is not None else (1, run, 0))
        pos = m.end()
    if pos < len(s):
        parts.append((1, s[pos:], 0))
    return tuple(parts)


def _sort_key(v: Any):
    """分组键排序：空值最后，数字和日期按值，文字按自然顺序（一店、二店……十二店）。"""
    if v is None or (isinstance(v, float) and v != v):
        return (2, ())
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (0, ((0, "", v),))
    if isinstance(v, (dt.date, dt.datetime)):
        return (0, ((1, v.isoformat(), 0),))
    return (1, _natural_key(str(v)))


def _numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _aggregate(values: pd.Series, func: str):
    if func == "计数":
        return int(values.notna().sum())
    if func in ("最大", "最小"):
        non_null = values.dropna()
        if non_null.empty:
            return None
        nums = _numeric(non_null)
        if nums.notna().all():
            return _clean_num(nums.max() if func == "最大" else nums.min())
        try:   # 日期等可比较类型
            return non_null.max() if func == "最大" else non_null.min()
        except TypeError:
            return None
    nums = _numeric(values).dropna()
    if nums.empty:
        return None if func == "平均" else 0
    return _clean_num(nums.sum() if func == "求和" else nums.mean())


def _clean_num(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    v = v.item() if hasattr(v, "item") else v
    if isinstance(v, float):
        v = round(v, 10)
        if v.is_integer():
            return int(v)
    return v


def agg_column_name(column: str, func: str) -> str:
    return f"{column}({func})"


def summarize(df: pd.DataFrame, group_by: list[str], aggs: list[tuple[str, str]],
              issues: IssueCollector, add_total: bool = True,
              add_count: bool = False) -> pd.DataFrame:
    """分组汇总。aggs 为 [(列名, 汇总方式中文名)]，末尾附加总计行。

    add_count=True 时附加"记录数"列；没有任何有效汇总列时也会附加，否则汇总表没有内容。
    """
    groups = [c for c in group_by if c in df.columns]
    for c in group_by:
        if c not in df.columns:
            issues.add(ERROR, T_AGG, f"分组列“{c}”在明细中不存在，已忽略", column=c)
    valid_aggs = _valid_aggs(df, aggs, issues)
    with_count = add_count or not valid_aggs
    out_cols = groups + [agg_column_name(c, f) for c, f in valid_aggs] + ([COL_COUNT] if with_count else [])

    rows: list[list[Any]] = []
    if groups:
        keyed: dict[tuple, list[int]] = {}
        for pos, key in enumerate(df[groups].itertuples(index=False, name=None)):
            keyed.setdefault(tuple(_hashable(v) for v in key), []).append(pos)
        for key in sorted(keyed, key=lambda k: tuple(_sort_key(v) for v in k)):
            part = df.iloc[keyed[key]]
            row = [EMPTY_GROUP_LABEL if v is None else v for v in key]
            row += [_aggregate(part[c], f) for c, f in valid_aggs]
            if with_count:
                row.append(len(part))
            rows.append(row)
    if add_total or not groups:
        total = ([TOTAL_LABEL] + [""] * (len(groups) - 1)) if groups else []
        total += [_aggregate(df[c], f) for c, f in valid_aggs] + ([len(df)] if with_count else [])
        rows.append(total)
    if not groups:
        out_cols = ["项目"] + out_cols
        rows[-1] = [TOTAL_LABEL] + rows[-1]
    return pd.DataFrame(rows, columns=out_cols, dtype=object)


# ---------------------------------------------------------------- 交叉表与拆分
PIVOT_TOTAL = "合计"
SUBTOTAL_LABEL = "小计"


def value_label(v: Any) -> str:
    """把分组 / 展开列的值转成显示文字（用于列名、Sheet 名、文件名）。"""
    v = _hashable(v)
    if v is None:
        return EMPTY_GROUP_LABEL
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S") if v.time() != dt.time(0) else v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


def _valid_aggs(df: pd.DataFrame, aggs: list[tuple[str, str]],
                issues: IssueCollector) -> list[tuple[str, str]]:
    valid = []
    for col, func in aggs:
        fn = normalize_agg_func(func)
        if col not in df.columns:
            issues.add(ERROR, T_AGG, f"汇总列“{col}”在明细中不存在，已忽略", column=col)
        elif fn is None:
            issues.add(ERROR, T_AGG, f"不支持的汇总方式“{func}”，已忽略", column=col)
        else:
            valid.append((col, fn))
    return valid


def split_positions(df: pd.DataFrame, column: str) -> list[tuple[Any, list[int]]]:
    """按某列的值分组，返回 [(值, 行位置列表)]，按值排序，空值在最后。"""
    keyed: dict[Any, list[int]] = {}
    for pos, v in enumerate(df[column].tolist()):
        keyed.setdefault(_hashable(v), []).append(pos)
    return [(k, keyed[k]) for k in sorted(keyed, key=_sort_key)]


def pivot_summarize(df: pd.DataFrame, group_by: list[str], pivot: str,
                    aggs: list[tuple[str, str]], issues: IssueCollector) -> pd.DataFrame:
    """交叉表：行为分组列，列为 pivot 列的每个值，右侧“合计”列，末尾“合计”行。

    多个汇总列时，列名为“值·金额(求和)”；没有汇总列时统计记录数。
    合计按明细重新计算（平均、最大等不是简单相加）。
    """
    if pivot not in df.columns:
        issues.add(ERROR, T_AGG, f"展开列“{pivot}”在明细中不存在，已按普通汇总输出", column=pivot)
        return summarize(df, group_by, aggs, issues)
    for c in group_by:
        if c not in df.columns:
            issues.add(ERROR, T_AGG, f"分组列“{c}”在明细中不存在，已忽略", column=c)
    groups = [c for c in group_by if c in df.columns and c != pivot]
    valid = _valid_aggs(df, aggs, issues)
    if valid:
        specs = [(agg_column_name(c, f), (lambda part, c=c, f=f: _aggregate(part[c], f)))
                 for c, f in valid]
    else:
        specs = [(COL_COUNT, lambda part: len(part))]
    multi = len(specs) > 1

    pivots = [v for v, _ in split_positions(df, pivot)]
    pv_index = {v: i for i, v in enumerate(pivots)}
    pv_values = [_hashable(v) for v in df[pivot].tolist()]
    if groups:
        gkeys = [tuple(_hashable(v) for v in k)
                 for k in df[groups].itertuples(index=False, name=None)]
    else:
        gkeys = [()] * len(df)
    cells: dict[tuple, dict[int, list[int]]] = {}
    for pos, (gk, pv) in enumerate(zip(gkeys, pv_values)):
        cells.setdefault(gk, {}).setdefault(pv_index[pv], []).append(pos)

    def header(pv_label: str, spec_label: str) -> str:
        return f"{pv_label}·{spec_label}" if multi else pv_label

    columns = list(groups) if groups else ["项目"]
    for pv in pivots:
        columns += [header(value_label(pv), lab) for lab, _ in specs]
    columns += [header(PIVOT_TOTAL, lab) for lab, _ in specs]

    rows: list[list[Any]] = []
    for gk in sorted(cells, key=lambda k: tuple(_sort_key(v) for v in k)):
        row = [EMPTY_GROUP_LABEL if v is None else v for v in gk] if groups else [TOTAL_LABEL]
        by_pv = cells[gk]
        for i in range(len(pivots)):
            pos = by_pv.get(i)
            part = df.iloc[pos] if pos else None
            row += [fn(part) if part is not None else None for _, fn in specs]
        all_pos = sorted(p for lst in by_pv.values() for p in lst)
        whole = df.iloc[all_pos]
        row += [fn(whole) for _, fn in specs]
        rows.append(row)
    if groups:
        total = [TOTAL_LABEL] + [""] * (len(groups) - 1)
        for pv, positions in split_positions(df, pivot):
            part = df.iloc[positions]
            total += [fn(part) for _, fn in specs]
        total += [fn(df) for _, fn in specs]
        rows.append(total)
    return pd.DataFrame(rows, columns=columns, dtype=object)


def subtotal_row(part: pd.DataFrame, label_column: str,
                 aggs: list[tuple[str, str]]) -> list[Any]:
    """拆分 Sheet 末尾的小计行：标签写在拆分列下，汇总值写在各汇总列下（同列多种方式取第一种）。"""
    cols = list(part.columns)
    row: list[Any] = [None] * len(cols)
    row[cols.index(label_column)] = SUBTOTAL_LABEL
    done = set()
    for col, func in aggs:
        fn = normalize_agg_func(func)
        if fn is None or col not in cols or col in done or col == label_column:
            continue
        row[cols.index(col)] = _aggregate(part[col], fn)
        done.add(col)
    return row
