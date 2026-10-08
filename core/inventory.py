"""盘点报表：明细里同时有“账面数量”“实盘数量”时，按盘点口径整理并输出固定格式的报表。

口径（与客户提供的标准答案一致）：
- 同一仓库同一商品盘了多轮（“盘点轮次”列：初盘 / 复盘…），只保留最后一轮
- 文件名与“仓库”列不一致、且该值正好是另一个文件的名字（套用模板没改）时，以文件名为准
- 实盘金额 = 实盘数量 × 单价（原表已填写的不覆盖；单价或实盘数量为空时留空）
- 盘盈盘亏数量 = 实盘数量合计 − 账面数量合计；未盘点（实盘数量为空）的商品不按 0 补，在总览中单独列出

输出 Sheet：总览 / 按仓库 / 按商品编码，后面接明细和问题清单。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from . import aggregator
from .aggregator import COL_FILE, COL_ROW, COL_SHEET
from .validator import INFO, WARNING, IssueCollector, T_NAME_MISMATCH

COL_BOOK, COL_REAL = "账面数量", "实盘数量"
COL_PRICE, COL_AMOUNT = "单价", "实盘金额"
COL_CODE, COL_NAME = "商品编码", "商品"
COL_ROUND = "盘点轮次"
LOCATION_COLUMNS = ("仓库", "门店")      # 盘点地点，按顺序取第一个存在的列

REPORT_AUTO, REPORT_INVENTORY, REPORT_GENERIC = "auto", "inventory", "generic"
REPORT_MODES = (REPORT_AUTO, REPORT_INVENTORY, REPORT_GENERIC)

T_ROUND_DROPPED = "多轮盘点"

# 轮次先后：数字越大越晚；认不出的轮次不参与筛选
_ROUND_RANK = {"初盘": 1, "一盘": 1, "首盘": 1, "第一轮": 1, "复盘": 2, "二盘": 2, "第二轮": 2,
               "三盘": 3, "第三轮": 3, "终盘": 9}

SHEET_OVERVIEW = "总览"


@dataclass
class InventoryReport:
    """报表各 Sheet 的内容（DataFrame）。"""
    location: str                                   # 地点列名（仓库 / 门店 / 来源文件）
    overview: pd.DataFrame
    by_location: pd.DataFrame
    by_code: pd.DataFrame | None
    notes: list[str] = field(default_factory=list)

    @property
    def location_sheet(self) -> str:
        return f"按{self.location}"

    def sheets(self) -> list[tuple[str, pd.DataFrame]]:
        out = [(SHEET_OVERVIEW, self.overview), (self.location_sheet, self.by_location)]
        if self.by_code is not None:
            out.append((f"按{COL_CODE}", self.by_code))
        return out


def applies(detail_columns: list[str], mode: str) -> bool:
    """是否输出盘点报表。auto：明细里同时有账面数量和实盘数量。"""
    if mode == REPORT_GENERIC:
        return False
    has = COL_BOOK in detail_columns and COL_REAL in detail_columns
    return has if mode == REPORT_AUTO else has or mode == REPORT_INVENTORY


def location_column(columns: list[str]) -> str | None:
    return next((c for c in LOCATION_COLUMNS if c in columns), None)


# ---------------------------------------------------------------- 明细整理
def name_from_filename(records: list[list[Any]], columns: list[str],
                       issues: IssueCollector, fix: bool) -> None:
    """文件名是“北仑三号仓”，“仓库”列却全写着“北仑二号仓”，而文件夹里正好有“北仑二号仓”文件
    （套用模板没改）。fix=True 时按文件名改正，否则只提示。records 末三列是来源文件/Sheet/行号。"""
    for col in ("门店", "仓库", "部门", "客户", "供应商"):
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
                issues.add(WARNING, T_NAME_MISMATCH,
                           f"文件名是“{stem}”，但“{col}”列全部写的是“{value}”（与文件“{stems[str(value)]}”"
                           f"相同，疑似套用模板没改），已按文件名改为“{stem}”；如以表内为准，请修改原表后重跑",
                           file=f, column=col, value=value)
            else:
                issues.add(WARNING, T_NAME_MISMATCH,
                           f"文件名是“{stem}”，但“{col}”列全部写的是“{value}”，与文件“{stems[str(value)]}”"
                           f"相同；按“{col}”汇总时两个文件的数据会合在一起，请核对是否套用模板没改",
                           file=f, column=col, value=value)


def keep_last_round(detail: pd.DataFrame, issues: IssueCollector) -> pd.DataFrame:
    """同一地点同一商品盘了多轮时只保留最后一轮（如以复盘为准，剔除初盘）。"""
    if COL_ROUND not in detail.columns or COL_CODE not in detail.columns or detail.empty:
        return detail
    loc = location_column(list(detail.columns)) or COL_FILE
    ranks = [_ROUND_RANK.get(str(v).strip()) if v is not None else None
             for v in detail[COL_ROUND].tolist()]
    if not any(ranks):
        return detail
    keys = list(zip(detail[loc].tolist(), detail[COL_CODE].tolist()))
    latest: dict[tuple, int] = {}
    for k, r in zip(keys, ranks):
        if r is not None and k[1] is not None:
            latest[k] = max(latest.get(k, 0), r)
    drop = [i for i, (k, r) in enumerate(zip(keys, ranks))
            if r is not None and k in latest and r < latest[k]]
    if not drop:
        return detail
    rounds = detail[COL_ROUND].tolist()
    for i in drop:
        row = detail.iloc[i]
        later = next(name for name, rank in _ROUND_RANK.items() if rank == latest[keys[i]])
        issues.add(INFO, T_ROUND_DROPPED,
                   f"{loc}“{keys[i][0]}”商品 {keys[i][1]} 盘了多轮，以{later}为准，"
                   f"这一行（{rounds[i]}）不计入明细和汇总",
                   file=row[COL_FILE], sheet=row[COL_SHEET], row=row[COL_ROW], column=COL_ROUND,
                   value=rounds[i])
    return detail.drop(index=detail.index[drop]).reset_index(drop=True)


def _num(v: Any) -> float | None:
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if v != v else float(v)
    return None


def add_amount(detail: pd.DataFrame) -> pd.DataFrame:
    """补上实盘金额 = 实盘数量 × 单价（已填写的不覆盖）。"""
    if COL_PRICE not in detail.columns or COL_REAL not in detail.columns:
        return detail
    real, price = detail[COL_REAL].tolist(), detail[COL_PRICE].tolist()
    old = detail[COL_AMOUNT].tolist() if COL_AMOUNT in detail.columns else [None] * len(detail)
    amount = []
    for q, p, a in zip(real, price, old):
        if _num(a) is not None:
            amount.append(a)
            continue
        q, p = _num(q), _num(p)
        amount.append(round(q * p, 2) if q is not None and p is not None else None)
    detail = detail.copy()
    if COL_AMOUNT in detail.columns:
        detail[COL_AMOUNT] = pd.Series(amount, index=detail.index, dtype=object)
    else:
        cols = list(detail.columns)
        detail.insert(cols.index(COL_PRICE) + 1, COL_AMOUNT, pd.Series(amount, index=detail.index,
                                                                       dtype=object))
    return detail


# ---------------------------------------------------------------- 报表
def _sum(values) -> float:
    total = 0.0
    for v in values:
        n = _num(v)
        if n is not None:
            total += n
    return total


def _clean(v: float) -> int | float:
    v = round(v, 2)
    return int(v) if float(v).is_integer() else v


def _location_values(detail: pd.DataFrame, loc: str) -> list:
    if loc == COL_FILE:
        return [Path(str(f)).stem for f in detail[COL_FILE].tolist()]
    return detail[loc].tolist()


def build_report(detail: pd.DataFrame, issues: IssueCollector,
                 issue_counts: dict[str, int] | None = None) -> InventoryReport:
    cols = list(detail.columns)
    loc_col = location_column(cols)
    loc = loc_col or COL_FILE
    loc_label = loc_col or "文件"
    has_code = COL_CODE in cols
    has_amount = COL_AMOUNT in cols

    locs = _location_values(detail, loc)
    codes = detail[COL_CODE].tolist() if has_code else [None] * len(detail)
    names = detail[COL_NAME].tolist() if COL_NAME in cols else [None] * len(detail)
    book, real = detail[COL_BOOK].tolist(), detail[COL_REAL].tolist()
    amount = detail[COL_AMOUNT].tolist() if has_amount else [None] * len(detail)

    # ---- 按地点
    groups: dict[Any, dict] = {}
    for lv, code, b, r, a in zip(locs, codes, book, real, amount):
        key = lv if lv is not None else aggregator.EMPTY_GROUP_LABEL
        g = groups.setdefault(key, {"codes": set(), "rows": 0, "book": 0.0, "real": 0.0,
                                    "amount": 0.0, "uncounted": 0})
        g["rows"] += 1
        if code is not None:
            g["codes"].add(code)
        g["book"] += _num(b) or 0.0
        g["real"] += _num(r) or 0.0
        g["amount"] += _num(a) or 0.0
        if _num(b) is not None and _num(r) is None:
            g["uncounted"] += 1
    loc_rows = []
    for key in sorted(groups, key=aggregator._sort_key):
        g = groups[key]
        row = [key, len(g["codes"]) if has_code else g["rows"], _clean(g["book"]), _clean(g["real"])]
        if has_amount:
            row.append(_clean(g["amount"]))
        row.append(_clean(g["real"] - g["book"]))
        row.append(g["uncounted"] or None)
        loc_rows.append(row)
    loc_header = [loc_label, "商品数" if has_code else "记录数", "账面数量", "实盘数量"]
    if has_amount:
        loc_header.append("实盘金额")
    loc_header += ["盘盈盘亏数量", "未盘商品数"]
    by_location = pd.DataFrame(loc_rows, columns=loc_header, dtype=object)

    # ---- 按商品编码
    by_code = None
    if has_code:
        cg: dict[Any, dict] = {}
        for lv, code, name, b, r in zip(locs, codes, names, book, real):
            if code is None:
                continue
            g = cg.setdefault(code, {"names": {}, "locs": set(), "book": 0.0, "real": 0.0})
            if name is not None:
                g["names"][name] = g["names"].get(name, 0) + 1
            g["locs"].add(lv)
            g["book"] += _num(b) or 0.0
            g["real"] += _num(r) or 0.0
        code_rows = []
        for code in sorted(cg, key=aggregator._sort_key):
            g = cg[code]
            name = max(g["names"], key=g["names"].get) if g["names"] else None
            code_rows.append([code, name, len(g["locs"]), _clean(g["book"]), _clean(g["real"]),
                              _clean(g["real"] - g["book"])])
        by_code = pd.DataFrame(code_rows, columns=[COL_CODE, "商品名称", f"出现{loc_label}数",
                                                   "账面数量", "实盘数量", "盘盈盘亏数量"],
                               dtype=object)

    # ---- 总览
    pairs = {(lv, c) for lv, c in zip(locs, codes) if c is not None} if has_code else set()
    total_book, total_real = _sum(book), _sum(real)
    uncounted = sum(1 for b, r in zip(book, real) if _num(b) is not None and _num(r) is None)
    items: list[tuple[str, Any]] = [
        (f"{loc_label}数", len(groups)),
    ]
    if has_code:
        items += [("不同商品编码数", len(set(c for c in codes if c is not None))),
                  (f"{loc_label}×商品 条数", len(pairs))]
    items += [("明细行数", len(detail)),
              ("账面数量合计", _clean(total_book)),
              ("实盘数量合计", _clean(total_real)),
              ("盘盈盘亏数量", _clean(total_real - total_book))]
    if has_amount:
        items.append(("实盘金额合计（元）", _clean(_sum(amount))))
    items.append(("未盘商品数（实盘数量为空）", uncounted))
    if issue_counts:
        items.append(("问题清单", "，".join(f"{k} {v} 条" for k, v in issue_counts.items())))
    items.append(("统计口径", "多轮盘点以最后一轮为准；文件名与表内仓库名冲突时以文件名为准；"
                             "实盘金额 = 实盘数量 × 单价；未盘商品实盘数量按空值，不按 0"))
    overview = pd.DataFrame(items, columns=["项目", "值"], dtype=object)
    return InventoryReport(location=loc_label, overview=overview, by_location=by_location,
                           by_code=by_code)

