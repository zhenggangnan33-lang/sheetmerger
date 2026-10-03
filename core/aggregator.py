"""汇总：合并明细（附加来源列）、去重、分组汇总。"""
from __future__ import annotations

import datetime as dt
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


def _sort_key(v: Any):
    """分组键排序：空值最后，同类型按值，不同类型按字符串。"""
    if v is None or (isinstance(v, float) and v != v):
        return (2, "")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (0, v, "")
    if isinstance(v, (dt.date, dt.datetime)):
        return (0, 0, v.isoformat())
    return (1, str(v))


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
              issues: IssueCollector, add_total: bool = True) -> pd.DataFrame:
    """分组汇总。aggs 为 [(列名, 汇总方式中文名)]。末尾附加"记录数"和总计行。"""
    groups = [c for c in group_by if c in df.columns]
    for c in group_by:
        if c not in df.columns:
            issues.add(ERROR, T_AGG, f"分组列“{c}”在明细中不存在，已忽略", column=c)
    valid_aggs = []
    for col, func in aggs:
        fn = normalize_agg_func(func)
        if col not in df.columns:
            issues.add(ERROR, T_AGG, f"汇总列“{col}”在明细中不存在，已忽略", column=col)
        elif fn is None:
            issues.add(ERROR, T_AGG, f"不支持的汇总方式“{func}”，已忽略", column=col)
        else:
            valid_aggs.append((col, fn))
    out_cols = groups + [agg_column_name(c, f) for c, f in valid_aggs] + [COL_COUNT]

    rows: list[list[Any]] = []
    if groups:
        keyed: dict[tuple, list[int]] = {}
        for pos, key in enumerate(df[groups].itertuples(index=False, name=None)):
            keyed.setdefault(tuple(_hashable(v) for v in key), []).append(pos)
        for key in sorted(keyed, key=lambda k: tuple(_sort_key(v) for v in k)):
            part = df.iloc[keyed[key]]
            row = [EMPTY_GROUP_LABEL if v is None else v for v in key]
            row += [_aggregate(part[c], f) for c, f in valid_aggs]
            row.append(len(part))
            rows.append(row)
    if add_total or not groups:
        total = ([TOTAL_LABEL] + [""] * (len(groups) - 1)) if groups else []
        total += [_aggregate(df[c], f) for c, f in valid_aggs] + [len(df)]
        rows.append(total)
    if not groups:
        out_cols = ["项目"] + out_cols
        rows[-1] = [TOTAL_LABEL] + rows[-1]
    return pd.DataFrame(rows, columns=out_cols).astype(object)
