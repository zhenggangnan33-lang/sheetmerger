"""自定义界面控件：顶部步骤条、统计卡片。"""
from __future__ import annotations

from PySide6.QtCore import Property, QPropertyAnimation, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QFrame, QLabel, QSizePolicy, QVBoxLayout, QWidget

from . import theme


class StepBar(QWidget):
    """横向步骤条：圆点 + 连接线。当前步骤为蓝色，已完成为绿色对勾，进度线随切换平滑填充。"""

    clicked = Signal(int)

    def __init__(self, titles: list[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.titles = titles
        self._current = 0
        self._progress = 0.0           # 进度线位置（0 ~ 步数-1），用于动画
        self.setMinimumHeight(70)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._anim = QPropertyAnimation(self, b"progress", self)
        self._anim.setEasingCurve(theme.EASE_IN_OUT)

    # 供动画使用的属性
    def _get_progress(self) -> float:
        return self._progress

    def _set_progress(self, v: float) -> None:
        self._progress = v
        self.update()

    progress = Property(float, _get_progress, _set_progress)

    @property
    def current(self) -> int:
        return self._current

    def set_current(self, idx: int) -> None:
        self._current = idx
        self._anim.stop()
        if theme.ANIMATIONS:
            self._anim.setStartValue(self._progress)
            self._anim.setEndValue(float(idx))
            self._anim.setDuration(450)
            self._anim.start()
        else:
            self._set_progress(float(idx))
        self.update()

    def _centers(self) -> list[float]:
        n = len(self.titles)
        margin = 70
        width = max(self.width() - 2 * margin, 1)
        return [margin + width * i / max(n - 1, 1) for i in range(n)]

    def paintEvent(self, _event) -> None:  # noqa: N802 - Qt 命名
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        xs = self._centers()
        y, r = 24.0, 13.0
        # 底线 + 已完成部分
        p.setPen(QPen(QColor(theme.BORDER), 3, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(int(xs[0]), int(y), int(xs[-1]), int(y))
        whole, frac = int(self._progress), self._progress - int(self._progress)
        end_x = xs[min(whole, len(xs) - 1)]
        if whole < len(xs) - 1:
            end_x += (xs[whole + 1] - xs[whole]) * frac
        p.setPen(QPen(QColor(theme.PRIMARY), 3, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(int(xs[0]), int(y), int(end_x), int(y))

        num_font = QFont(self.font())
        num_font.setBold(True)
        title_font = QFont(self.font())
        for i, (x, title) in enumerate(zip(xs, self.titles)):
            reached = self._progress >= i - 0.05
            done = i < self._current
            active = i == self._current
            if done:
                fill, border, fg = QColor(theme.SUCCESS), QColor(theme.SUCCESS), QColor("white")
            elif active and reached:
                fill, border, fg = QColor(theme.PRIMARY), QColor(theme.PRIMARY), QColor("white")
            else:
                fill, border, fg = QColor("white"), QColor("#C9D1DE"), QColor(theme.MUTED)
            if active and reached:
                halo = QColor(theme.PRIMARY)
                halo.setAlpha(40)
                p.setPen(Qt.NoPen)
                p.setBrush(halo)
                p.drawEllipse(QRectF(x - r - 5, y - r - 5, 2 * r + 10, 2 * r + 10))
            p.setPen(QPen(border, 2))
            p.setBrush(fill)
            p.drawEllipse(QRectF(x - r, y - r, 2 * r, 2 * r))
            p.setPen(fg)
            p.setFont(num_font)
            p.drawText(QRectF(x - r, y - r, 2 * r, 2 * r), Qt.AlignCenter,
                       "✓" if done else str(i + 1))
            title_font.setBold(active)
            p.setFont(title_font)
            p.setPen(QColor(theme.TEXT if (active or done) else theme.MUTED))
            p.drawText(QRectF(x - 80, y + r + 6, 160, 20), Qt.AlignHCenter | Qt.AlignTop, title)
        p.end()


class StatTile(QFrame):
    """统计卡片：小标题 + 大数字，数字可滚动显示。"""

    def __init__(self, title: str, color: str = theme.TEXT, suffix: str = "",
                 decimals: int = 0, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("card")
        self.suffix, self.decimals = suffix, decimals
        self.value_label = QLabel("–")
        f = QFont(self.font())
        f.setPointSize(18)
        f.setBold(True)
        self.value_label.setFont(f)
        self.value_label.setStyleSheet(f"color: {color}; border: none; background: transparent;")
        title_label = QLabel(title)
        title_label.setObjectName("muted")
        title_label.setStyleSheet("border: none; background: transparent;")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 8, 16, 8)
        lay.setSpacing(2)
        lay.addWidget(title_label)
        lay.addWidget(self.value_label)

    def set_value(self, v: float) -> None:
        text = f"{v:,.{self.decimals}f}" if self.decimals else f"{int(round(v)):,}"
        self.value_label.setText(text + self.suffix)

    def clear(self) -> None:
        self.value_label.setText("–")
