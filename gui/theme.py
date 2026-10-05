"""界面主题与动效。

配色沿用程序图标的蓝色。动效参考 GSAP 的常用做法，用 Qt 动画实现：
- 缓动：减速曲线为主（GSAP power2.out / power3.out ≈ Qt OutQuad / OutCubic）
- 时长：界面过渡 0.3–0.5 秒；多个元素依次出现时每个间隔约 0.08 秒（stagger）
- 只动画透明度和位移，不改变布局；动画结束后移除透明度效果（相当于 GSAP 的 clearProps）
- Windows 关闭了“动画效果”或设置了 SHEETMERGER_NO_ANIM=1 时，不播放动画（对应 prefers-reduced-motion）
"""
from __future__ import annotations

import os
import sys
from typing import Callable

from PySide6.QtCore import (QAbstractAnimation, QEasingCurve, QPoint, QPropertyAnimation,
                            QTimer, QVariantAnimation)
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import QApplication, QGraphicsOpacityEffect, QWidget

# ---------------------------------------------------------------- 配色
PRIMARY = "#2F6FDB"
PRIMARY_DARK = "#1C4FB0"
PRIMARY_LIGHT = "#EAF1FD"
SUCCESS = "#1E9E5A"
WARNING = "#D9822B"
DANGER = "#D64545"
BG = "#F3F5F9"
CARD = "#FFFFFF"
BORDER = "#E1E6EF"
TEXT = "#1F2733"
MUTED = "#6B7585"

# 缓动（与 GSAP 名称对应）
EASE_OUT = QEasingCurve.OutCubic        # power3.out：界面元素进入
EASE_SOFT = QEasingCurve.OutQuad        # power2.out：数值、进度变化
EASE_IN_OUT = QEasingCurve.InOutCubic   # power3.inOut：步骤条进度线

STYLE_SHEET = f"""
QMainWindow, QWidget#central {{ background: {BG}; }}
QWidget {{ color: {TEXT}; }}
QLabel#appTitle {{ font-size: 15pt; font-weight: 600; color: {TEXT}; }}
QLabel#appSubtitle {{ color: {MUTED}; }}
QLabel#muted {{ color: {MUTED}; }}
QLabel#hint {{ color: {MUTED}; padding: 2px 2px 6px 2px; }}
QWidget#header {{ background: {CARD}; border-bottom: 1px solid {BORDER}; }}
QFrame#card {{ background: {CARD}; border: 1px solid {BORDER}; border-radius: 10px; }}
QWidget#footer {{ background: transparent; }}

QGroupBox {{
    background: {CARD}; border: 1px solid {BORDER}; border-radius: 10px;
    margin-top: 14px; padding: 14px 10px 10px 10px; font-weight: 600;
}}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px; color: {TEXT}; }}

QPushButton {{
    background: {CARD}; border: 1px solid {BORDER}; border-radius: 6px;
    padding: 6px 14px; min-height: 20px;
}}
QPushButton:hover {{ border-color: {PRIMARY}; color: {PRIMARY}; }}
QPushButton:pressed {{ background: {PRIMARY_LIGHT}; }}
QPushButton:disabled {{ color: #A9B0BC; border-color: #E8EBF1; background: #F7F8FA; }}
QPushButton[role="primary"] {{
    background: {PRIMARY}; color: white; border: 1px solid {PRIMARY}; font-weight: 600;
}}
QPushButton[role="primary"]:hover {{ background: {PRIMARY_DARK}; border-color: {PRIMARY_DARK}; color: white; }}
QPushButton[role="primary"]:disabled {{ background: #A9C2EE; border-color: #A9C2EE; color: white; }}
QPushButton[role="ghost"] {{ background: transparent; border: 1px solid transparent; color: {MUTED}; }}
QPushButton[role="ghost"]:hover {{ background: {PRIMARY_LIGHT}; color: {PRIMARY}; }}

QLineEdit, QComboBox, QPlainTextEdit {{
    background: {CARD}; border: 1px solid {BORDER}; border-radius: 6px; padding: 5px 8px;
    selection-background-color: {PRIMARY};
}}
QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {{ border-color: {PRIMARY}; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{ border: 1px solid {BORDER}; selection-background-color: {PRIMARY_LIGHT};
    selection-color: {TEXT}; }}

QTreeWidget, QTableWidget, QListWidget {{
    background: {CARD}; border: 1px solid {BORDER}; border-radius: 8px;
    alternate-background-color: #F8FAFD; gridline-color: #EEF1F6;
    selection-background-color: {PRIMARY_LIGHT}; selection-color: {TEXT};
}}
QTreeWidget::item, QListWidget::item {{ padding: 3px 2px; }}
QHeaderView::section {{
    background: #F6F8FB; color: {MUTED}; font-weight: 600; border: none;
    border-bottom: 1px solid {BORDER}; border-right: 1px solid #EEF1F6; padding: 6px 8px;
}}

QProgressBar {{
    background: #E8EDF5; border: none; border-radius: 6px; height: 12px; text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{
    border-radius: 6px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {PRIMARY}, stop:1 #5B9BF0);
}}

QCheckBox {{ spacing: 6px; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: #CDD5E1; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: #AEB9C9; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: #CDD5E1; border-radius: 4px; min-width: 30px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QStatusBar {{ background: {CARD}; border-top: 1px solid {BORDER}; color: {MUTED}; }}
QToolTip {{ background: {TEXT}; color: white; border: none; padding: 4px 6px; }}
"""


def apply_theme(app: QApplication) -> None:
    """应用字体、调色板和样式表。"""
    app.setStyle("Fusion")
    font = QFont()
    font.setFamilies(["Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC",
                      "Noto Sans CJK SC", "Noto Sans SC", "sans-serif"])
    font.setPointSize(10)
    app.setFont(font)
    pal = app.palette()
    pal.setColor(QPalette.Window, QColor(BG))
    pal.setColor(QPalette.Base, QColor(CARD))
    pal.setColor(QPalette.AlternateBase, QColor("#F8FAFD"))
    pal.setColor(QPalette.Highlight, QColor(PRIMARY))
    pal.setColor(QPalette.Text, QColor(TEXT))
    pal.setColor(QPalette.WindowText, QColor(TEXT))
    app.setPalette(pal)
    app.setStyleSheet(STYLE_SHEET)


def set_role(button, role: str) -> None:
    """按钮角色：primary（主操作）/ ghost（次要）。"""
    button.setProperty("role", role)
    button.style().unpolish(button)
    button.style().polish(button)


# ---------------------------------------------------------------- 动效
def _system_animations_enabled() -> bool:
    if os.environ.get("SHEETMERGER_NO_ANIM") == "1":
        return False
    if sys.platform == "win32":
        # 设置 → 辅助功能 → 视觉效果 → “动画效果”关闭时，尊重用户选择
        try:
            import ctypes
            enabled = ctypes.c_bool(True)
            SPI_GETCLIENTAREAANIMATION = 0x1042
            if ctypes.windll.user32.SystemParametersInfoW(SPI_GETCLIENTAREAANIMATION, 0,
                                                          ctypes.byref(enabled), 0):
                return bool(enabled.value)
        except (AttributeError, OSError):
            pass
    return True


ANIMATIONS = _system_animations_enabled()


def _ms(seconds: float) -> int:
    return int(seconds * 1000) if ANIMATIONS else 0


def _keep(owner: QWidget, anim: QAbstractAnimation) -> QAbstractAnimation:
    """动画对象挂在控件上，避免被回收；同一控件同一用途只保留最新一个。"""
    anim.setParent(owner)
    return anim


def fade_slide_in(widget: QWidget, dx: int = 0, dy: int = 14, duration: float = 0.38,
                  delay: float = 0.0) -> None:
    """淡入并从偏移位置滑到原位（gsap.from({autoAlpha: 0, y: 14, ease: "power3.out"})）。"""
    if not ANIMATIONS or not widget.isVisible():
        return
    effect = QGraphicsOpacityEffect(widget)
    effect.setOpacity(0.0)
    widget.setGraphicsEffect(effect)
    end = widget.pos()
    start = end + QPoint(dx, dy)

    fade = _keep(widget, QPropertyAnimation(effect, b"opacity"))
    fade.setStartValue(0.0)
    fade.setEndValue(1.0)
    fade.setDuration(_ms(duration))
    fade.setEasingCurve(EASE_OUT)

    move = _keep(widget, QPropertyAnimation(widget, b"pos"))
    move.setStartValue(start)
    move.setEndValue(end)
    move.setDuration(_ms(duration))
    move.setEasingCurve(EASE_OUT)

    def finish() -> None:
        widget.setGraphicsEffect(None)   # 结束后移除效果（clearProps），避免影响表格等控件的绘制
        widget.move(end)

    fade.finished.connect(finish)

    def start_all() -> None:
        widget.move(start)
        fade.start(QAbstractAnimation.DeleteWhenStopped)
        move.start(QAbstractAnimation.DeleteWhenStopped)

    if delay > 0:
        QTimer.singleShot(_ms(delay), start_all)
    else:
        start_all()


def stagger_in(widgets: list[QWidget], each: float = 0.08, **kw) -> None:
    """依次淡入（GSAP stagger: 0.08）。"""
    for i, w in enumerate(widgets):
        fade_slide_in(w, delay=i * each, **kw)


def count_up(owner: QWidget, setter: Callable[[float], None], target: float,
             duration: float = 0.9, delay: float = 0.0) -> None:
    """数字从 0 滚动到目标值（gsap.to(obj, {val: target, ease: "power2.out"})）。"""
    if not ANIMATIONS:
        setter(target)
        return
    anim = _keep(owner, QVariantAnimation())
    anim.setStartValue(0.0)
    anim.setEndValue(float(target))
    anim.setDuration(_ms(duration))
    anim.setEasingCurve(EASE_SOFT)
    anim.valueChanged.connect(lambda v: setter(v))
    anim.finished.connect(lambda: setter(target))
    setter(0)
    if delay > 0:
        QTimer.singleShot(_ms(delay), lambda: anim.start(QAbstractAnimation.DeleteWhenStopped))
    else:
        anim.start(QAbstractAnimation.DeleteWhenStopped)


def smooth_value(widget, value: int, duration: float = 0.25) -> None:
    """进度条平滑移动到新值，而不是一格一格跳。"""
    if not ANIMATIONS or value < widget.value():
        widget.setValue(value)
        return
    old = getattr(widget, "_smooth_anim", None)
    if old is not None:
        old.stop()
    anim = QPropertyAnimation(widget, b"value", widget)
    anim.setStartValue(widget.value())
    anim.setEndValue(value)
    anim.setDuration(_ms(duration))
    anim.setEasingCurve(EASE_SOFT)
    widget._smooth_anim = anim
    anim.start()
