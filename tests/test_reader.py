from pathlib import Path

from core import reader
from core.validator import ERROR, INFO, WARNING, IssueCollector, T_READ_FAIL, T_SKIP_NOTE, T_SKIP_TOTAL


def _load(data_dir: Path, name: str, sheet: str | None = None, header_row=None):
    issues = IssueCollector()
    path = data_dir / name
    sheet = sheet or reader.list_sheets(path)[0]
    ref = reader.SheetRef(path=path, rel=name, sheet=sheet)
    return reader.load_sheet(ref, issues, header_row), issues


def test_scan_skips_temp_hidden_and_non_tables(data_dir):
    names = [p.name for p in reader.scan_folder(data_dir)]
    assert "01_标准格式.xlsx" in names
    assert "~$01_标准格式.xlsx" not in names
    assert ".隐藏文件.xlsx" not in names
    assert "说明.txt" not in names
    assert "20_子文件夹数据.xlsx" not in names


def test_scan_recursive(data_dir):
    rels = [p.relative_to(data_dir).as_posix() for p in reader.scan_folder(data_dir, recursive=True)]
    assert "子文件夹/20_子文件夹数据.xlsx" in rels


def test_header_detection_skips_title_total_and_note_rows(data_dir):
    table, issues = _load(data_dir, "06_标题行和合计行.xlsx")
    assert table.header_row == 4
    assert table.columns == ["日期", "门店", "商品", "数量", "单价", "金额"]
    assert [n for n, _ in table.rows] == list(range(5, 13))
    skipped = {(i.type, i.row) for i in issues.issues}
    assert (T_SKIP_TOTAL, 13) in skipped
    assert (T_SKIP_NOTE, 14) in skipped


def test_header_row_override(data_dir):
    table, _ = _load(data_dir, "06_标题行和合计行.xlsx", header_row=4)
    assert table.header_row == 4 and table.auto_header_row == 4


def test_merged_cells_filled_down(data_dir):
    table, _ = _load(data_dir, "07_合并单元格.xlsx")
    stores = [r[1] for _, r in table.rows]
    assert stores == ["一店"] * 4 + ["二店"] * 4
    assert all(r[0] is not None for _, r in table.rows)


def test_xls_with_merge(data_dir):
    table, issues = _load(data_dir, "15_老格式.xls")
    assert len(table.rows) == 6
    assert {r[1] for _, r in table.rows} == {"三店"}


def test_gbk_csv(data_dir):
    table, _ = _load(data_dir, "13_GBK编码.csv")
    assert table.columns[:3] == ["交易日期", "店铺", "商品名称"]
    assert len(table.rows) == 8


def test_utf8_bom_csv(data_dir):
    table, _ = _load(data_dir, "14_UTF8_BOM.csv")
    assert table.columns[0] == "日期"          # BOM 不能混进表头
    assert any("," in r[5] for _, r in table.rows)   # 带引号的千分位金额作为整体读入


def test_decode_csv_bytes():
    assert reader.decode_csv_bytes("门店,金额".encode("gbk"))[0] == "门店,金额"
    assert reader.decode_csv_bytes("门店,金额".encode("utf-8-sig"))[0] == "门店,金额"


def test_multi_sheet(data_dir):
    assert reader.list_sheets(data_dir / "12_多Sheet.xlsx") == ["一店9月", "二店9月", "说明"]


def test_discover_reports_corrupt_and_empty(data_dir):
    issues = IssueCollector()
    refs = reader.discover(data_dir, False, issues)
    rels = {r.rel for r in refs}
    assert "16_损坏文件.xlsx" not in rels
    by_file = {i.file: i for i in issues.issues}
    assert by_file["16_损坏文件.xlsx"].severity == ERROR
    assert by_file["16_损坏文件.xlsx"].type == T_READ_FAIL
    assert by_file["18_空文件.csv"].severity == WARNING


def test_empty_sheet_and_header_only(data_dir):
    table, issues = _load(data_dir, "17_空文件.xlsx")
    assert table is None and issues.count(WARNING) == 1
    table, issues = _load(data_dir, "19_只有表头.xlsx")
    assert table is not None and table.rows == [] and issues.count(WARNING) == 1


def test_detect_header_prefers_distinct_text():
    grid = [["报表标题"] * 5, [None] * 5, ["日期", "门店", "商品", "数量", "金额"], [1, 2, 3, 4, 5]]
    assert reader.detect_header_row(grid) == 2


def test_fill_merged_horizontal_and_vertical():
    grid = [["a", None, None], [None, None, None]]
    reader.fill_merged(grid, [(0, 0, 1, 2)])
    assert grid == [["a"] * 3, ["a"] * 3]


def test_load_never_raises_on_garbage(tmp_path):
    bad = tmp_path / "x.xls"
    bad.write_bytes(b"not an excel file")
    issues = IssueCollector()
    assert reader.load_sheet(reader.SheetRef(bad, "x.xls", "Sheet1"), issues) is None
    assert issues.count(ERROR) == 1
    assert INFO  # 保持导入
