"""生成程序图标：packaging/app.ico（多尺寸，打包 exe 用）和 gui/app_icon.png（窗口图标）。

图案：左侧多张小表格 → 箭头 → 右侧一张汇总大表（绿色表头）。
16–32 像素使用简化图案，保证任务栏里也清晰。
用法：python packaging/make_icon.py
"""
from __future__ import annotations

import struct
import sys
from pathlib import Path

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QLinearGradient, QPainter, QPainterPath, QPen

ROOT = Path(__file__).resolve().parents[1]
SIZES = [16, 24, 32, 48, 64, 128, 256]

BG_TOP, BG_BOTTOM = QColor("#3478E5"), QColor("#1C4FB0")
SHEET = QColor("#FFFFFF")
GRID = QColor("#C9D6EA")
HEADER_SMALL = QColor("#9CC0F5")
HEADER_BIG = QColor("#21A366")


def _sheet(p: QPainter, rect: QRectF, header: QColor, rows: int, cols: int, radius: float,
           line: float) -> None:
    p.setPen(Qt.NoPen)
    p.setBrush(SHEET)
    p.drawRoundedRect(rect, radius, radius)
    head = QRectF(rect.x(), rect.y(), rect.width(), rect.height() / (rows + 1))
    path = QPainterPath()
    path.addRoundedRect(head, radius, radius)
    path.addRect(QRectF(head.x(), head.center().y(), head.width(), head.height() / 2))
    p.setBrush(header)
    p.drawPath(path.simplified())
    if line <= 0:
        return
    p.setPen(QPen(GRID, line))
    for r in range(1, rows + 1):
        y = rect.y() + head.height() * (r + 1)
        if y < rect.bottom() - line:
            p.drawLine(QPointF(rect.x() + line, y), QPointF(rect.right() - line, y))
    for c in range(1, cols):
        x = rect.x() + rect.width() * c / cols
        p.drawLine(QPointF(x, head.bottom()), QPointF(x, rect.bottom() - line))


def draw(size: int) -> QImage:
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.scale(size / 256, size / 256)          # 统一按 256 坐标绘制

    grad = QLinearGradient(0, 0, 0, 256)
    grad.setColorAt(0, BG_TOP)
    grad.setColorAt(1, BG_BOTTOM)
    p.setPen(Qt.NoPen)
    p.setBrush(grad)
    margin = 8 if size >= 48 else 0
    p.drawRoundedRect(QRectF(margin, margin, 256 - 2 * margin, 256 - 2 * margin), 52, 52)

    simple = size <= 32
    arrow = QPen(QColor("#FFFFFF"), 22 if simple else 14, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
    if simple:
        # 简化：一张小表 → 一张大表
        _sheet(p, QRectF(26, 84, 72, 88), HEADER_SMALL, 2, 1, 12, 0)
        p.setPen(arrow)
        p.drawPolyline([QPointF(112, 100), QPointF(134, 128), QPointF(112, 156)])
        _sheet(p, QRectF(146, 46, 88, 164), HEADER_BIG, 3, 2, 14, 10)
    else:
        for i, y in enumerate((46, 106, 166)):
            _sheet(p, QRectF(30, y, 70, 46), HEADER_SMALL, 2, 2, 8, 4)
        p.setPen(arrow)
        for y in (69, 129, 189):
            p.drawLine(QPointF(108, y), QPointF(132, 128 + (y - 129) * 0.3))
        p.drawPolyline([QPointF(128, 108), QPointF(148, 128), QPointF(128, 148)])
        _sheet(p, QRectF(158, 50, 72, 156), HEADER_BIG, 5, 2, 10, 5)
    p.end()
    return img


def _png_bytes(img: QImage) -> bytes:
    data = QByteArray()
    buf = QBuffer(data)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    return bytes(data)


def write_ico(path: Path, images: list[QImage]) -> None:
    """写多尺寸 ICO（PNG 压缩条目，Windows Vista 及以上支持）。"""
    pngs = [_png_bytes(i) for i in images]
    header = struct.pack("<HHH", 0, 1, len(pngs))
    offset = len(header) + 16 * len(pngs)
    entries, blobs = b"", b""
    for img, png in zip(images, pngs):
        w = img.width() if img.width() < 256 else 0
        entries += struct.pack("<BBBBHHII", w, w, 0, 0, 1, 32, len(png), offset)
        blobs += png
        offset += len(png)
    path.write_bytes(header + entries + blobs)


def main() -> None:
    app = QGuiApplication.instance() or QGuiApplication(sys.argv)  # noqa: F841 - 绘图需要
    images = [draw(s) for s in SIZES]
    write_ico(ROOT / "packaging" / "app.ico", images)
    images[-1].save(str(ROOT / "gui" / "app_icon.png"))
    print("已生成 packaging/app.ico、gui/app_icon.png")


if __name__ == "__main__":
    main()
