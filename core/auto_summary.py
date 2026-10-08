"""自动汇总：没有选分组列时，按“编码 → 名称 → 出现在几个主体 → 数值合计”的方式输出汇总表。

做法来自客户认可的“按商品编码”汇总表，对任何表格都适用，不针对具体业务：
- 分组键：角色为“编码”的列（如 商品编码）；没有编码列时用“名称”列（如 商品），再没有就用“主体”列
- 名称列：跟在编码后面，取该编码下出现最多的写法（如 商品名称）
- 出现X数：每个编码出现在几个不同的主体里（如 出现仓库数、出现门店数）
- 数值列：全部求和；“单价”等比率类数值不求和（字典里标为 rate 的列，以及名称或原表头像
  单价 / 成本价 / 折扣率 / 占比 这类的列；总价、金额、价税合计等仍然求和）
各列角色写在别名字典里（key / name / entity / rate），用户可以改，程序据此“记住习惯”。
设置了汇总列时，只汇总设置的列（按设置的方式）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from . import aggregator, cleaner
from .header_mapper import ROLE_ENTITY, ROLE_KEY, ROLE_NAME, ROLE_RATE, AliasStore


@dataclass
class AutoPlan:
    """自动汇总用到的各列。"""
    key: str | None = None
    names: list[str] = field(default_factory=list)
    entities: list[str] = field(default_factory=list)
    measures: list[tuple[str, str]] = field(default_factory=list)   # (列名, 汇总方式)

    def describe(self) -> str:
        if not self.key:
            return "不分组，只出总计"
        parts = [f"按“{self.key}”汇总"]
        if self.names:
            parts.append(f"带出{'、'.join(self.names)}")
        if self.entities:
            parts.append("统计" + "、".join(f"出现{e}数" for e in self.entities))
        if self.measures:
            parts.append("合计" + "、".join(c for c, _f in self.measures))
        return "，".join(parts)


_RATE_WORDS = ("价", "率", "%", "％", "占比", "比例")
_AMOUNT_WORDS = ("金额", "总价", "价税", "价值", "价款", "合计", "总额")


def looks_like_rate(header: str) -> bool:
    """单价、成本价、折扣率、毛利率、占比 —— 这类数值按行相加没有意义。"""
    h = str(header)
    return any(w in h for w in _RATE_WORDS) and not any(w in h for w in _AMOUNT_WORDS)


def _is_numeric_column(values: list[Any]) -> bool:
    seen = False
    for v in values:
        if v is None:
            continue
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            return False
        seen = True
    return seen


def make_plan(detail: pd.DataFrame, store: AliasStore, types: dict[str, str],
              aggs: list[tuple[str, str]] = (), sources: dict[str, set] | None = None) -> AutoPlan:
    """sources：明细列 -> 各表中对应的原表头（用来识别“成本价(内部)”映射成“成本”这类情况）。"""
    cols = [c for c in detail.columns if c not in aggregator.SOURCE_COLUMNS]
    present = [c for c in cols if detail[c].notna().any()]
    role = {c: store.column_role(c) for c in present}
    plan = AutoPlan()
    by_role = {r: [c for c in present if role[c] == r] for r in (ROLE_KEY, ROLE_NAME, ROLE_ENTITY)}
    # 编码列也认“code”类型（用户自己指定为编码的列）
    keys = by_role[ROLE_KEY] or [c for c in present if types.get(c) == cleaner.TYPE_CODE]
    if keys:
        plan.key, plan.names = keys[0], by_role[ROLE_NAME]
    elif by_role[ROLE_NAME]:
        plan.key = by_role[ROLE_NAME][0]
    elif by_role[ROLE_ENTITY]:
        plan.key = by_role[ROLE_ENTITY][0]
    plan.entities = [c for c in by_role[ROLE_ENTITY] if c != plan.key]

    if aggs:
        plan.measures = [(c, aggregator.normalize_agg_func(f) or f) for c, f in aggs]
    else:
        for c in present:
            if c == plan.key or role[c] == ROLE_RATE or c in plan.names or c in plan.entities:
                continue
            if not role[c] and any(looks_like_rate(h) for h in {c, *(sources or {}).get(c, ())}):
                continue
            t = types.get(c)
            if t == cleaner.TYPE_NUMBER or (t in (None, cleaner.TYPE_AUTO)
                                            and _is_numeric_column(detail[c].tolist())):
                plan.measures.append((c, "求和"))
    return plan


def _measure_title(column: str, func: str) -> str:
    return column if func == "求和" else aggregator.agg_column_name(column, func)


def summarize(detail: pd.DataFrame, plan: AutoPlan) -> pd.DataFrame:
    measures = [(c, f) for c, f in plan.measures if c in detail.columns]
    titles = [_measure_title(c, f) for c, f in measures]
    if not plan.key:
        row = ["总计"] + [aggregator._aggregate(detail[c], f) for c, f in measures] + [len(detail)]
        return pd.DataFrame([row], columns=["项目"] + titles + [aggregator.COL_COUNT], dtype=object)

    keys = detail[plan.key].tolist()
    groups: dict[Any, list[int]] = {}
    for pos, k in enumerate(keys):
        groups.setdefault(aggregator._hashable(k), []).append(pos)
    header = [plan.key] + plan.names + [f"出现{e}数" for e in plan.entities] + titles
    rows = []
    for key in sorted(groups, key=aggregator._sort_key):
        part = detail.iloc[groups[key]]
        row: list[Any] = [aggregator.EMPTY_GROUP_LABEL if key is None else key]
        for n in plan.names:
            counts: dict[Any, int] = {}
            for v in part[n].tolist():
                if v is not None:
                    counts[v] = counts.get(v, 0) + 1
            row.append(max(counts, key=counts.get) if counts else None)
        for e in plan.entities:
            row.append(len({aggregator._hashable(v) for v in part[e].tolist() if v is not None}))
        row += [aggregator._aggregate(part[c], f) for c, f in measures]
        rows.append(row)
    total: list[Any] = [aggregator.TOTAL_LABEL] + [None] * len(plan.names)
    total += [len({aggregator._hashable(v) for v in detail[e].tolist() if v is not None})
              for e in plan.entities]
    total += [aggregator._aggregate(detail[c], f) for c, f in measures]
    rows.append(total)
    return pd.DataFrame(rows, columns=header, dtype=object)
