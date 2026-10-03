"""生成故意弄脏的测试数据样本。

用法：python tests/generate_test_data.py [输出文件夹]（默认 tests/test_data）

所有数据均为虚构的门店销售流水。生成结果固定（随机种子固定），测试用例依赖其中的行号。
"""
from __future__ import annotations

import datetime as dt
import random
import shutil
import sys
from pathlib import Path

import openpyxl

STORES = ["一店", "二店", "三店", "四店"]
PRODUCTS = [("苹果", 6.5), ("香蕉", 3.2), ("橙子", 5.0), ("葡萄", 12.8), ("西瓜", 2.5)]
STD_HEADER = ["日期", "门店", "商品", "数量", "单价", "金额"]
EXCEL_EPOCH = dt.date(1899, 12, 30)


def excel_serial(d: dt.date) -> int:
    return (d - EXCEL_EPOCH).days


def make_records(rng: random.Random, n: int, store: str | None = None,
                 start: dt.date = dt.date(2026, 9, 1)) -> list[dict]:
    rows = []
    for i in range(n):
        name, price = rng.choice(PRODUCTS)
        qty = rng.randint(1, 50)
        rows.append({
            "日期": start + dt.timedelta(days=i % 28),
            "门店": store or rng.choice(STORES),
            "商品": name,
            "数量": qty,
            "单价": price,
            "金额": round(qty * price, 2),
        })
    return rows


def write_xlsx(path: Path, sheets: dict[str, list[list]], merges: dict[str, list[str]] | None = None,
               date_cols: dict[str, list[int]] | None = None) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
        for rng_ in (merges or {}).get(name, []):
            ws.merge_cells(rng_)
        for col in (date_cols or {}).get(name, []):
            for row in ws.iter_rows(min_col=col, max_col=col):
                for c in row:
                    if isinstance(c.value, (dt.date, dt.datetime)):
                        c.number_format = "yyyy-mm-dd"
    wb.save(path)


def table(records: list[dict], header: list[str], keys: list[str] | None = None) -> list[list]:
    keys = keys or header
    return [header] + [[r.get(k) for k in keys] for r in records]


def generate(out: Path) -> Path:
    rng = random.Random(20261002)
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)

    # 01 标准格式
    base = make_records(rng, 12)
    write_xlsx(out / "01_标准格式.xlsx", {"销售明细": table(base, STD_HEADER)},
               date_cols={"销售明细": [1]})

    # 02 表头同义不同名（含全角括号、换行、首尾空格）
    recs = make_records(rng, 10)
    write_xlsx(out / "02_表头同义.xlsx", {"Sheet1": table(
        recs, ["业务日期", " 门店名称 ", "品名", "销量", "单价/元", "金额（元）"], STD_HEADER)})

    # 03 列顺序不同
    recs = make_records(rng, 10)
    order = ["金额", "商品", "门店", "日期", "单价", "数量"]
    write_xlsx(out / "03_列顺序不同.xlsx", {"Sheet1": table(recs, order)})

    # 04 多一列（备注、经手人、一个字典里没有的列）
    recs = make_records(rng, 8)
    for r in recs:
        r["备注"] = rng.choice(["", "促销", "会员价"]) or None
        r["经手人"] = rng.choice(["张三", "李四"])
        r["会员卡号"] = f"VIP{rng.randint(1000, 9999)}"
    write_xlsx(out / "04_多一列.xlsx", {"Sheet1": table(
        recs, STD_HEADER + ["备注", "经手人", "会员卡号"])})

    # 05 少一列（没有单价）
    recs = make_records(rng, 8)
    write_xlsx(out / "05_少一列.xlsx", {"Sheet1": table(recs, ["日期", "门店", "商品", "数量", "金额"])})

    # 06 表头上方有标题行，下方有合计行和说明行
    #    第1行标题(合并A1:F1) 第2行 单位：元 第3行空 第4行表头 第5-12行数据 第13行合计 第14行制表人
    recs = make_records(rng, 8)
    rows = [["2026年9月门店销售报表"], ["单位：元"], []] + table(recs, STD_HEADER)
    rows.append(["合计", None, None, sum(r["数量"] for r in recs), None,
                 round(sum(r["金额"] for r in recs), 2)])
    rows.append(["制表人：王五"])
    write_xlsx(out / "06_标题行和合计行.xlsx", {"Sheet1": rows}, merges={"Sheet1": ["A1:F1"]})

    # 07 合并单元格：门店列按店合并（向下填充），日期也合并
    rows = [STD_HEADER]
    merges = []
    r_idx = 2
    for store in ["一店", "二店"]:
        recs = make_records(rng, 4, store=store, start=dt.date(2026, 9, 10))
        for i, r in enumerate(recs):
            rows.append([dt.date(2026, 9, 10) if i == 0 else None,
                         store if i == 0 else None, r["商品"], r["数量"], r["单价"], r["金额"]])
        merges += [f"A{r_idx}:A{r_idx + 3}", f"B{r_idx}:B{r_idx + 3}"]
        r_idx += 4
    write_xlsx(out / "07_合并单元格.xlsx", {"Sheet1": rows}, merges={"Sheet1": merges})

    # 08 文本型数字：千分位、"元"后缀、货币符号、前后空格；第 7 行有一个无法转换的值
    rows = [STD_HEADER,
            ["2026-09-01", "一店", "苹果", "10", "6.5", "65元"],
            ["2026-09-02", "二店", "葡萄", " 100 ", "12.8", "1,280.00"],
            ["2026-09-03", "三店", "西瓜", "1,000", "2.5", "2,500元"],
            ["2026-09-04", "四店", "香蕉", "20", "3.2", "￥64.00"],
            ["2026-09-05", "一店", "橙子", "１２", "5", " 60 "],     # 全角数字
            ["2026-09-06", "二店", "苹果", "8", "6.5", "约52元"],      # 第 7 行：无法转换
            ["2026-09-07", "三店", "苹果", "4", "6.5", "(26.00)"]]     # 会计格式负数（退货）
    write_xlsx(out / "08_文本数字.xlsx", {"Sheet1": rows})

    # 09 多种日期格式混用；第 10、11 行为无法识别的日期
    d = dt.date(2026, 9, 2)
    dates = ["2026-09-02", "2026/9/3", 20260904, "2026年9月5日", excel_serial(dt.date(2026, 9, 6)),
             "2026.09.07", "20260908", dt.datetime(2026, 9, 9, 0, 0), "2026-13-45", "下周一"]
    rows = [STD_HEADER]
    for i, dv in enumerate(dates):
        name, price = PRODUCTS[i % len(PRODUCTS)]
        rows.append([dv, STORES[i % 4], name, 2, price, round(2 * price, 2)])
    write_xlsx(out / "09_日期混用.xlsx", {"Sheet1": rows})
    del d

    # 10 占位值："无" "/" "-" "—" "N/A"；第 7 行数量为非法文本
    rows = [STD_HEADER,
            [dt.date(2026, 9, 1), "一店", "苹果", "无", 6.5, "-"],
            [dt.date(2026, 9, 2), "二店", "香蕉", 5, "/", 16],
            [dt.date(2026, 9, 3), "三店", "N/A", 3, 5.0, 15],
            ["—", "四店", "葡萄", 1, 12.8, 12.8],
            [dt.date(2026, 9, 5), "一店", "西瓜", "—", "—", "无"],
            [dt.date(2026, 9, 6), "二店", "苹果", "很多", 6.5, 13]]
    write_xlsx(out / "10_占位值.xlsx", {"Sheet1": rows}, date_cols={"Sheet1": [1]})

    # 11 重复记录：第 3 行重复第 2 行；第 5、6 行与 01 文件的前两条完全相同
    recs = make_records(rng, 3)
    rows = table(recs, STD_HEADER)
    rows.insert(2, list(rows[1]))
    rows += [[base[0][k] for k in STD_HEADER], [base[1][k] for k in STD_HEADER]]
    write_xlsx(out / "11_重复记录.xlsx", {"Sheet1": rows})

    # 12 多 Sheet：两个数据 Sheet + 一个说明 Sheet
    write_xlsx(out / "12_多Sheet.xlsx", {
        "一店9月": table(make_records(rng, 6, store="一店"), STD_HEADER),
        "二店9月": table(make_records(rng, 6, store="二店"), STD_HEADER),
        "说明": [["本文件由门店每日上报"], ["如有疑问请联系财务部"]],
    })

    # 13 GBK 编码 CSV（表头同义）
    recs = make_records(rng, 8)
    lines = ["交易日期,店铺,商品名称,数量,单价,销售额"]
    lines += [f"{r['日期']:%Y/%m/%d},{r['门店']},{r['商品']},{r['数量']},{r['单价']},{r['金额']}"
              for r in recs]
    (out / "13_GBK编码.csv").write_bytes(("\r\n".join(lines) + "\r\n").encode("gbk"))

    # 14 UTF-8-BOM CSV，金额带千分位（加引号）
    recs = make_records(rng, 6)
    lines = [",".join(STD_HEADER)]
    for r in recs:
        amount = f'"{r["金额"] * 10:,.2f}"'
        lines.append(f"{r['日期']:%Y-%m-%d},{r['门店']},{r['商品']},{r['数量'] * 10},{r['单价']},{amount}")
    (out / "14_UTF8_BOM.csv").write_bytes(("\n".join(lines) + "\n").encode("utf-8-sig"))

    # 15 老格式 .xls（含合并单元格）
    import xlwt
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("老系统导出")
    date_style = xlwt.easyxf(num_format_str="YYYY-MM-DD")
    for c, h in enumerate(STD_HEADER):
        ws.write(0, c, h)
    recs = make_records(rng, 6, store="三店")
    for i, r in enumerate(recs, start=1):
        ws.write(i, 0, r["日期"], date_style)
        if i == 1:
            ws.write_merge(1, 6, 1, 1, "三店")
        for c, k in enumerate(["商品", "数量", "单价", "金额"], start=2):
            ws.write(i, c, r[k])
    wb.save(str(out / "15_老格式.xls"))

    # 16 损坏的文件
    (out / "16_损坏文件.xlsx").write_bytes(b"PK\x03\x04" + bytes(rng.randrange(256) for _ in range(500)))

    # 17 空工作簿、18 空 CSV（0 字节）
    write_xlsx(out / "17_空文件.xlsx", {"Sheet1": []})
    (out / "18_空文件.csv").write_bytes(b"")

    # 19 只有表头没有数据
    write_xlsx(out / "19_只有表头.xlsx", {"Sheet1": [STD_HEADER]})

    # 20 子文件夹中的文件（只有勾选"包含子文件夹"时才读取）
    (out / "子文件夹").mkdir()
    write_xlsx(out / "子文件夹" / "20_子文件夹数据.xlsx",
               {"Sheet1": table(make_records(rng, 5, store="五店"), STD_HEADER)})

    # 应被跳过的文件：Excel 临时文件、隐藏文件、非表格文件
    (out / "~$01_标准格式.xlsx").write_bytes(b"lock")
    write_xlsx(out / ".隐藏文件.xlsx", {"Sheet1": table(make_records(rng, 2), STD_HEADER)})
    (out / "说明.txt").write_text("这不是表格文件", encoding="utf-8")
    return out


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("test_data")
    print(f"测试数据已生成：{generate(target)}")
