"""性能测试：生成 N 个文件 × M 行的脏数据并完整运行一次（不在 pytest 中自动运行）。

用法：python tests/benchmark.py [文件数=50] [每个文件行数=10000] [工作目录]
目标：50 个文件 × 1 万行在 60 秒内完成。
"""
from __future__ import annotations

import datetime as dt
import random
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from openpyxl import Workbook  # noqa: E402

from config.task_config import AggSpec, TaskConfig  # noqa: E402
from core import pipeline  # noqa: E402
from generate_test_data import PRODUCTS, STORES  # noqa: E402

HEADERS = [
    ["日期", "门店", "商品", "数量", "单价", "金额"],
    ["业务日期", "门店名称", "品名", "销量", "单价", "金额(元)"],
    ["金额", "商品", "门店", "交易日期", "单价", "数量"],
]


def _dirty_date(rng: random.Random, d: dt.date):
    kind = rng.random()
    if kind < 0.6:
        return d
    if kind < 0.75:
        return d.strftime("%Y/%m/%d")
    if kind < 0.85:
        return int(d.strftime("%Y%m%d"))
    if kind < 0.95:
        return f"{d.year}年{d.month}月{d.day}日"
    return "无"


def _dirty_amount(rng: random.Random, v: float):
    kind = rng.random()
    if kind < 0.7:
        return v
    if kind < 0.9:
        return f"{v:,.2f}元"
    if kind < 0.999:
        return f" {v} "
    return "待定"


def make_files(folder: Path, n_files: int, n_rows: int) -> None:
    rng = random.Random(7)
    folder.mkdir(parents=True, exist_ok=True)
    for f in range(n_files):
        header = HEADERS[f % len(HEADERS)]
        keys = ["日期", "门店", "商品", "数量", "单价", "金额"]
        order = [k for k in ["金额", "商品", "门店", "日期", "单价", "数量"]] if f % 3 == 2 else keys
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("数据")
        ws.append([f"第 {f + 1} 号门店上报"])
        ws.append(header)
        total = 0.0
        for i in range(n_rows):
            name, price = rng.choice(PRODUCTS)
            qty = rng.randint(1, 60)
            amount = round(qty * price, 2)
            total += amount
            rec = {"日期": _dirty_date(rng, dt.date(2026, 1, 1) + dt.timedelta(days=i % 270)),
                   "门店": rng.choice(STORES), "商品": name, "数量": qty, "单价": price,
                   "金额": _dirty_amount(rng, amount)}
            ws.append([rec[k] for k in order])
        ws.append(["合计", None, None, None, None, round(total, 2)])
        wb.save(folder / f"门店上报_{f + 1:02d}.xlsx")


def main() -> None:
    n_files = int(sys.argv[1]) if len(sys.argv) > 1 else 50
    n_rows = int(sys.argv[2]) if len(sys.argv) > 2 else 10_000
    work = Path(sys.argv[3]) if len(sys.argv) > 3 else Path(tempfile.mkdtemp(prefix="sm_bench_"))
    data = work / "data"
    if not data.exists():
        t = time.perf_counter()
        make_files(data, n_files, n_rows)
        print(f"生成 {n_files} 个文件 × {n_rows} 行：{time.perf_counter() - t:.1f} 秒")
    cfg = TaskConfig(input_folder=str(data), group_by=["门店", "商品"],
                     aggregations=[AggSpec("金额", "求和"), AggSpec("数量", "平均")],
                     output_dir=str(work), output_name="bench.xlsx")
    marks: list[tuple[float, str]] = []
    t0 = time.perf_counter()

    def progress(_p: int, msg: str) -> None:
        stage = msg.split(" ")[0]
        if not marks or marks[-1][1] != stage:
            marks.append((time.perf_counter() - t0, stage))

    result = pipeline.run(cfg, progress=progress)
    total = time.perf_counter() - t0
    for (t, stage), nxt in zip(marks, marks[1:] + [(total, "")]):
        print(f"  {stage:<8} {nxt[0] - t:6.1f} 秒")
    c = result.issue_counts
    print(f"明细 {result.rows_detail} 行，问题 错误{c['错误']}/警告{c['警告']}/提示{c['提示']}，"
          f"总用时 {total:.1f} 秒，结果 {result.output_path} "
          f"({result.output_path.stat().st_size / 1e6:.1f} MB)")
    print("达标" if total < 60 else "未达标（目标 60 秒）")


if __name__ == "__main__":
    main()
