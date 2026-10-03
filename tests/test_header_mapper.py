import json

from core.header_mapper import (STATUS_AUTO, STATUS_IGNORED, STATUS_MANUAL, STATUS_PENDING,
                                STATUS_UNMATCHED, AliasStore, normalize_header, similarity,
                                suggest, suggest_for_columns)


def test_normalize():
    assert normalize_header("  金额（元）\n") == "金额(元)"
    assert normalize_header("ＡＭＯＵＮＴ") == "amount"
    assert normalize_header("门 店：") == "门店"


def test_alias_and_unit_stripping():
    store = AliasStore()
    s = suggest("交易日期", store)
    assert (s.target, s.status) == ("日期", STATUS_AUTO)
    s = suggest("金额【元】", store)
    assert (s.target, s.status) == ("金额", STATUS_AUTO)
    assert s.score >= 90


def test_thresholds():
    store = AliasStore()
    s = suggest("单价/元", store)
    assert s.status == STATUS_PENDING and s.target == "单价"
    assert 60 <= s.score < 90
    s = suggest("会员卡号", store)
    assert s.status == STATUS_UNMATCHED and s.target is None
    # 一整句话里包含"门店"不应被当成门店列
    assert suggest("本文件由门店每日上报", store).status == STATUS_UNMATCHED


def test_similarity_bounds():
    assert similarity("金额", "金额") == 100
    assert similarity("金额(元)", "金额") == 95
    assert similarity("", "金额") == 0


def test_conflict_resolution():
    store = AliasStore()
    result, conflicts = suggest_for_columns(["金额", "销售额"], store)
    assert result[0].target == "金额"
    assert result[1].status == STATUS_UNMATCHED
    assert conflicts


def test_overrides_manual_and_ignore():
    store = AliasStore()
    result, _ = suggest_for_columns(["会员卡号", "单价/元"], store, {"会员卡号": "客户", "单价/元": ""})
    assert (result[0].target, result[0].status) == ("客户", STATUS_MANUAL)
    assert result[1].status == STATUS_IGNORED


def test_learn_persists(isolated_home):
    store = AliasStore()
    store.learn("单价/元", "单价")
    again = AliasStore()
    s = suggest("单价/元", again)
    assert (s.target, s.status, s.score) == ("单价", STATUS_AUTO, 100.0)
    data = json.loads((isolated_home / "aliases.json").read_text(encoding="utf-8"))
    assert "单价/元" in data["单价"]["aliases"]


def test_learn_new_standard_column(isolated_home):
    store = AliasStore()
    store.learn("卡号", "会员卡号", "text")
    assert AliasStore().column_type("会员卡号") == "text"
    assert suggest("卡号", AliasStore()).target == "会员卡号"
