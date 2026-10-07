"""多仓库月末盘点场景：每个仓库文件埋一种常见的坑，汇总结果必须与标准答案一致。"""
import collections

import pytest

from config.task_config import AggSpec, TaskConfig
from core import pipeline
from generate_inventory_data import generate


@pytest.fixture(scope="module")
def inv(tmp_path_factory):
    folder = tmp_path_factory.mktemp("inv") / "盘点"
    return folder, generate(folder)


def _run(folder, **kw):
    cfg = TaskConfig(input_folder=str(folder), excluded=["仓库盘点模板.xlsx|*"], **kw)
    return pipeline.run(cfg, export=False)


def _by(result, kind):
    return collections.Counter(i.file for i in result.issues.issues if i.type == kind)


def test_totals_per_warehouse_match_answer(inv):
    folder, exp = inv
    r = _run(folder, group_by=["仓库"],
             aggregations=[AggSpec("账面数量"), AggSpec("实盘数量")])
    got = {row[0]: (row[1], row[2]) for row in r.summary.itertuples(index=False, name=None)}
    for wh, totals in exp["totals"].items():
        assert got[wh] == totals, wh
    assert "示例仓" not in got


def test_codes_merge_across_warehouses(inv):
    folder, exp = inv
    r = _run(folder, group_by=["商品编码"], aggregations=[AggSpec("实盘数量")])
    codes = set(r.summary["商品编码"]) - {"总计"}
    assert codes <= exp["catalog_codes"]                 # 补零、BX 编码统一后不会拆成两行
    assert "BX-001" in codes or "BX-002" in codes


def test_each_trap_is_reported(inv):
    folder, exp = inv
    r = _run(folder, group_by=["仓库"], aggregations=[AggSpec("实盘数量")])
    assert _by(r, "文件读取失败") == {"月湖仓.xlsx": 1}
    assert _by(r, "数值无法转换") == {"东钱湖仓.xlsx": 2}          # 单价是 #N/A、#DIV/0!
    assert _by(r, "编码已补齐")["北仑二号仓.xlsx"] == 1
    assert _by(r, "日期补全年份")["梅山仓.xlsx"] == 1
    empty = [i for i in r.issues.issues if i.type == "空值" and i.file == "宁海仓.xlsx"]
    assert empty and f"有 {exp['empty_real']} 个空值" in empty[0].message
    assert _by(r, "文件名与内容不一致") == {"北仑三号仓.xlsx": 1}
    assert any("另一张表" in i.message for i in r.issues.issues if i.file == "镇海仓.xlsx")
    assert not _by(r, "日期无法识别")                              # 梅山仓 5 种写法都认得


def test_rounds_flagged_by_warehouse_and_code(inv):
    folder, exp = inv
    r = _run(folder, group_by=["仓库"], aggregations=[AggSpec("实盘数量")],
             dedup_columns=["仓库", "商品编码"])
    assert _by(r, "重复记录")["象山仓.xlsx"] == exp["duplicate_rounds"]


def test_english_headers_mapped(inv):
    folder, _exp = inv
    r = _run(folder, group_by=["仓库"], aggregations=[AggSpec("实盘数量")])
    dx = r.detail[r.detail["来源文件"] == "大榭仓.xlsx"]
    assert dx["实盘数量"].notna().all() and dx["商品"].notna().all()
    assert not {"Count Qty", "Book Qty", "Item Name"} & set(r.detail.columns)
