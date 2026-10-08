"""自动汇总：按列的角色（编码 / 名称 / 主体 / 单价）决定汇总表结构，适用于任何表格。"""
import datetime as dt

import pandas as pd

from config.task_config import TaskConfig
from core import auto_summary, pipeline
from core.header_mapper import AliasStore


def _detail(columns, rows):
    df = pd.DataFrame(rows, columns=columns, dtype=object)
    for c in ("来源文件", "来源Sheet", "原始行号"):
        df[c] = None
    return df


def test_code_name_entity_and_sums(tmp_path):
    store = AliasStore(user_path=tmp_path / "a.json")
    d = _detail(["门店", "商品编码", "商品", "数量", "单价", "金额"], [
        ("一店", "A01", "可乐", 2, 3.0, 6.0),
        ("二店", "A01", "可乐", 1, 3.0, 3.0),
        ("二店", "A01", "可口可乐", 1, 3.0, 3.0),     # 名称写法不一，取出现最多的
        ("一店", "B02", "雪碧", 5, 2.5, 12.5),
    ])
    types = {"门店": "text", "商品编码": "code", "商品": "text", "数量": "number",
             "单价": "number", "金额": "number"}
    plan = auto_summary.make_plan(d, store, types)
    assert (plan.key, plan.names, plan.entities) == ("商品编码", ["商品"], ["门店"])
    assert plan.measures == [("数量", "求和"), ("金额", "求和")]          # 单价不求和
    s = auto_summary.summarize(d, plan)
    assert list(s.columns) == ["商品编码", "商品", "出现门店数", "数量", "金额"]
    assert s.values.tolist() == [["A01", "可乐", 2, 4, 12],
                                 ["B02", "雪碧", 1, 5, 12.5],
                                 ["总计", None, 2, 9, 24.5]]


def test_without_code_uses_name_then_entity(tmp_path):
    store = AliasStore(user_path=tmp_path / "a.json")
    d = _detail(["门店", "商品", "金额"], [("一店", "可乐", 1), ("二店", "可乐", 2)])
    plan = auto_summary.make_plan(d, store, {"金额": "number"})
    assert (plan.key, plan.entities) == ("商品", ["门店"])
    d = _detail(["部门", "费用"], [("行政", 10), ("财务", 5)])
    plan = auto_summary.make_plan(d, store, {})
    assert plan.key == "部门" and plan.measures == [("费用", "求和")]    # 未识别的数值列也求和


def test_learned_roles_change_the_habit(tmp_path):
    """用户把“物料号”记为编码、“折扣率”记为不求和，下次自动按物料号汇总。"""
    store = AliasStore(user_path=tmp_path / "a.json")
    store.learn_role("物料号", "key")
    store.learn_role("折扣率", "rate")
    store = AliasStore(user_path=tmp_path / "a.json")              # 重新加载，确认已保存
    d = _detail(["物料号", "部门", "数量", "折扣率"], [("M1", "生产", 3, 0.9), ("M1", "仓储", 2, 0.8)])
    plan = auto_summary.make_plan(d, store, {})
    assert plan.key == "物料号" and plan.measures == [("数量", "求和")]
    assert auto_summary.summarize(d, plan).values.tolist()[0] == ["M1", 2, 5]


def test_no_roles_gives_total_only(tmp_path):
    store = AliasStore(user_path=tmp_path / "a.json")
    d = _detail(["日期", "说明", "金额"], [(dt.date(2026, 9, 1), "x", 1), (dt.date(2026, 9, 2), "y", 2)])
    plan = auto_summary.make_plan(d, store, {"日期": "date", "金额": "number"})
    s = auto_summary.summarize(d, plan)
    assert plan.key is None and s.values.tolist() == [["总计", 3, 2]]


def test_pipeline_default_on_sales_data(data_dir, tmp_path):
    """销售测试数据没有编码列：按商品汇总，统计出现门店数，数量和金额求和。"""
    r = pipeline.run(TaskConfig(input_folder=str(data_dir)), export=False)
    s = r.summary
    assert r.auto_plan.key == "商品"
    assert list(s.columns[:2]) == ["商品", "出现门店数"]
    assert {"数量", "金额"} <= set(s.columns) and "单价" not in s.columns


def test_explicit_aggregations_still_win(data_dir):
    from config.task_config import AggSpec
    r = pipeline.run(TaskConfig(input_folder=str(data_dir),
                                aggregations=[AggSpec("金额", "平均")]), export=False)
    assert list(r.summary.columns) == ["商品", "出现门店数", "金额(平均)"]


def test_price_like_columns_not_summed(tmp_path):
    store = AliasStore(user_path=tmp_path / "a.json")
    d = _detail(["商品编码", "成本", "折扣率", "总价", "数量"], [("A", 1.5, 0.9, 10, 2), ("A", 1.5, 0.8, 20, 3)])
    plan = auto_summary.make_plan(d, store, {"商品编码": "code"}, sources={"成本": {"成本价(内部)"}})
    assert plan.measures == [("总价", "求和"), ("数量", "求和")]
    assert auto_summary.looks_like_rate("毛利率%") and not auto_summary.looks_like_rate("价税合计")
