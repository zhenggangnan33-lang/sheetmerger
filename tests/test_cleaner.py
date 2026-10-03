import datetime as dt

import pytest

from core.cleaner import clean_value, is_placeholder, to_date, to_number, to_text


@pytest.mark.parametrize("raw,expected", [
    ("1,234.50元", 1234.5), (" 88 ", 88), ("￥300.00", 300), ("2,000", 2000),
    ("１２", 12), ("(26.00)", -26), ("15%", 0.15), (12.0, 12), (3.5, 3.5), (7, 7),
    ("人民币 1,000", 1000),
])
def test_to_number_ok(raw, expected):
    value, err = to_number(raw)
    assert err is None and value == pytest.approx(expected)


@pytest.mark.parametrize("raw", ["约52元", "很多", "1.2.3", True, dt.date(2026, 1, 1)])
def test_to_number_errors(raw):
    value, err = to_number(raw)
    assert value is None and err


@pytest.mark.parametrize("raw", ["无", "/", "-", "—", "N/A", "n/a", "", "  ", None, "——"])
def test_placeholders(raw):
    assert is_placeholder(raw)
    assert to_number(raw) == (None, None)
    assert to_date(raw) == (None, None)
    assert to_text(raw) == (None, None)


@pytest.mark.parametrize("raw", [
    "2026-10-02", "2026/10/2", "20261002", 20261002, "2026年10月2日", "2026.10.02",
    46297, 46297.0, "46297", dt.datetime(2026, 10, 2), dt.date(2026, 10, 2), "２０２６-１０-０２",
])
def test_to_date_ok(raw):
    assert to_date(raw) == (dt.date(2026, 10, 2), None)


def test_to_date_with_time():
    assert to_date("2026/10/2 14:30")[0] == dt.datetime(2026, 10, 2, 14, 30)
    assert to_date(46297.5)[0] == dt.datetime(2026, 10, 2, 12, 0)


@pytest.mark.parametrize("raw", ["2026-13-45", "下周一", "20261340", -5, "2026年10月"])
def test_to_date_errors(raw):
    value, err = to_date(raw)
    assert value is None and err


def test_text_normalizes_numbers():
    assert to_text(101.0) == ("101", None)
    assert to_text(" 一店 ") == ("一店", None)
    assert to_text(dt.date(2026, 1, 2)) == ("2026-01-02", None)


def test_clean_value_auto_keeps_type():
    assert clean_value(5.0, "auto") == (5, None)
    assert clean_value(" x ", "auto") == ("x", None)
    assert clean_value("无", "auto") == (None, None)
