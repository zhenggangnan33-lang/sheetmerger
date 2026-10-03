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
    s = ag.summarize(_df(), ["门店"], aggs, issues)
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
    s = ag.summarize(_df(), ["不存在"], [("金额", "求和"), ("x", "求和"), ("金额", "中位数")], issues)
    assert list(s.columns) == ["项目", "金额(求和)", "记录数"]
    assert s.iloc[0].tolist() == ["总计", 260.5, 5]
    assert issues.count(ERROR) == 3


def test_normalize_agg_func():
    assert ag.normalize_agg_func("sum") == "求和"
    assert ag.normalize_agg_func("平均") == "平均"
    assert ag.normalize_agg_func("中位数") is None
