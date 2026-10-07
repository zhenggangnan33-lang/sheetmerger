import pytest
import datetime as dt

from core import aggregator as ag
from core.validator import ERROR, IssueCollector, T_DUPLICATE, T_DUPLICATE_DROPPED

COLS = ["门店", "金额", "日期"] + ag.SOURCE_COLUMNS


def _df():
    rows = [
        ["一店", 100, dt.date(2026, 9, 1), "a.xlsx", "S1", 2],
        ["一店", 100, dt.date(2026, 9, 1), "a.xlsx", "S1", 3],   # 与第 2 行重复
        ["二店", 50.5, dt.date(2026, 9, 3), "b.xlsx", "S1", 2],
        [None, 10, None, "b.xlsx", "S1", 3],
        ["二店", None, dt.date(2026, 9, 2), "b.xlsx", "S1", 4],
    ]
    return ag.build_detail(rows, COLS)


def test_dedup_mark_keeps_rows():
    issues = IssueCollector()
    out = ag.deduplicate(_df(), ag.DEDUP_MARK, None, issues)
    assert len(out) == 5
    (issue,) = issues.issues
    assert (issue.type, issue.file, issue.row) == (T_DUPLICATE, "a.xlsx", 3)
    assert "第 2 行" in issue.message


def test_dedup_drop_and_columns():
    issues = IssueCollector()
    out = ag.deduplicate(_df(), ag.DEDUP_DROP, None, issues)
    assert len(out) == 4 and issues.issues[0].type == T_DUPLICATE_DROPPED
    issues = IssueCollector()
    out = ag.deduplicate(_df(), ag.DEDUP_DROP, ["门店"], issues)
    assert len(out) == 3   # 一店×2、二店×2、空 → 去掉一店、二店各一条


def test_dedup_off():
    issues = IssueCollector()
    assert len(ag.deduplicate(_df(), ag.DEDUP_OFF, None, issues)) == 5 and not issues.issues


def test_summarize_all_funcs():
    issues = IssueCollector()
    aggs = [("金额", f) for f in ["求和", "计数", "平均", "最大", "最小"]] + [("日期", "最大")]
    s = ag.summarize(_df(), ["门店"], aggs, issues, add_count=True)
    assert list(s.columns) == ["门店", "金额(求和)", "金额(计数)", "金额(平均)", "金额(最大)",
                               "金额(最小)", "日期(最大)", "记录数"]
    rows = {r[0]: r for r in s.itertuples(index=False, name=None)}
    assert rows["一店"][1:] == (200, 2, 100, 100, 100, dt.date(2026, 9, 1), 2)
    assert rows["二店"][1:] == (50.5, 1, 50.5, 50.5, 50.5, dt.date(2026, 9, 3), 2)
    assert rows["(空)"][1] == 10
    assert rows["总计"][1] == 260.5 and rows["总计"][-1] == 5
    assert list(s["门店"])[-1] == "总计"   # 总计在最后
    assert not issues.issues


def test_summarize_no_group_and_bad_columns():
    issues = IssueCollector()
    s = ag.summarize(_df(), ["不存在"], [("金额", "求和"), ("x", "求和"), ("金额", "中位数")], issues,
                     add_count=True)
    assert list(s.columns) == ["项目", "金额(求和)", "记录数"]
    assert s.iloc[0].tolist() == ["总计", 260.5, 5]
    assert issues.count(ERROR) == 3


def test_normalize_agg_func():
    assert ag.normalize_agg_func("sum") == "求和"
    assert ag.normalize_agg_func("平均") == "平均"
    assert ag.normalize_agg_func("中位数") is None


def test_count_column_optional():
    issues = IssueCollector()
    s = ag.summarize(_df(), ["门店"], [("金额", "求和")], issues)
    assert list(s.columns) == ["门店", "金额(求和)"]           # 默认不附加记录数
    assert s.iloc[-1].tolist() == ["总计", 260.5]
    s = ag.summarize(_df(), ["门店"], [("金额", "求和")], issues, add_count=True)
    assert list(s.columns) == ["门店", "金额(求和)", "记录数"]
    s = ag.summarize(_df(), ["门店"], [], issues)               # 没有汇总列时仍保留记录数
    assert list(s.columns) == ["门店", "记录数"]


# ---------------------------------------------------------------- 交叉表与拆分
def _sales():
    rows = [
        ["一店", "苹果", 10, 1, "a", "S", 2],
        ["二店", "苹果", 20, 3, "a", "S", 3],
        ["一店", "香蕉", 5, 5, "a", "S", 4],
        ["十二店", "苹果", 7, 2, "a", "S", 5],
        ["二店", "苹果", 30, 5, "a", "S", 6],
    ]
    return ag.build_detail(rows, ["门店", "商品", "金额", "数量"] + ag.SOURCE_COLUMNS)


def test_natural_sort_order():
    names = ["十二店", "二店", "10号店", "一店", "2号店", "三店", "十店"]
    assert sorted(names, key=ag._sort_key) == ["一店", "2号店", "二店", "三店", "10号店", "十店", "十二店"]
    assert sorted([None, "b", 3], key=ag._sort_key) == [3, "b", None]


def test_pivot_single_agg():
    issues = IssueCollector()
    p = ag.pivot_summarize(_sales(), ["商品"], "门店", [("金额", "求和")], issues)
    assert list(p.columns) == ["商品", "一店", "二店", "十二店", "合计"]
    rows = {r[0]: r[1:] for r in p.itertuples(index=False, name=None)}
    assert rows["苹果"] == (10, 50, 7, 67)
    assert rows["香蕉"] == (5, None, None, 5)            # 没有数据的格子留空
    assert rows["总计"] == (15, 50, 7, 72)
    assert not issues.issues


def test_pivot_multi_agg_totals_recomputed():
    issues = IssueCollector()
    p = ag.pivot_summarize(_sales(), ["商品"], "门店", [("金额", "求和"), ("数量", "平均")], issues)
    assert list(p.columns)[:3] == ["商品", "一店·金额(求和)", "一店·数量(平均)"]
    assert list(p.columns)[-2:] == ["合计·金额(求和)", "合计·数量(平均)"]
    total = p.iloc[-1].tolist()
    assert total[0] == "总计" and total[-2] == 72
    assert total[-1] == pytest.approx(16 / 5)          # 平均按明细重新计算，不是各列平均再平均


def test_pivot_no_group_and_count_fallback():
    issues = IssueCollector()
    p = ag.pivot_summarize(_sales(), [], "门店", [], issues)
    assert list(p.columns) == ["项目", "一店", "二店", "十二店", "合计"]
    assert p.iloc[0].tolist() == ["总计", 2, 2, 1, 5]


def test_pivot_missing_column_falls_back():
    issues = IssueCollector()
    p = ag.pivot_summarize(_sales(), ["商品"], "不存在", [("金额", "求和")], issues)
    assert list(p.columns) == ["商品", "金额(求和)"] and issues.count(ERROR) == 1


def test_split_positions_and_subtotal():
    df = _sales()
    parts = ag.split_positions(df, "门店")
    assert [v for v, _ in parts] == ["一店", "二店", "十二店"]
    pos = dict(parts)["二店"]
    row = ag.subtotal_row(df.iloc[pos], "门店", [("金额", "求和"), ("金额", "平均"), ("数量", "最大")])
    assert row[:4] == ["小计", None, 50, 5]
