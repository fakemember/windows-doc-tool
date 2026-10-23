"""
TOC 样式缩略图绘制模块。

使用 QPainter 动态绘制 4 种视觉样式的缩略卡片，供用户直观选择。
"""

from __future__ import annotations

from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QFrame, QLabel, QRadioButton, QVBoxLayout

from wdt.toc.styles_templates import TocStyle


def draw_style_preview(style: TocStyle, width: int = 160, height: int = 100) -> QPixmap:
    """使用 QPainter 渲染指定 TOC 样式的矢量缩略图。"""
    pixmap = QPixmap(width, height)
    pixmap.fill(QColor("#FFFFFF"))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

    # 边框与纸张背景
    border_pen = QPen(QColor("#E2E8F0"), 1)
    painter.setPen(border_pen)
    painter.setBrush(QColor("#FAFAFA"))
    painter.drawRoundedRect(QRectF(1, 1, width - 2, height - 2), 4, 4)

    # 1. 样式专属参数配置
    if style == TocStyle.CLASSIC_DOT:
        title_text = "目  录"
        accent_color = QColor("#1E293B")
        text_color = QColor("#334155")
        dot_style = "dot"
        page_color = QColor("#334155")
        has_accent_bar = False
    elif style == TocStyle.MODERN_ACCENT:
        title_text = "Contents"
        accent_color = QColor("#0284C7")  # 现代天蓝
        text_color = QColor("#1E293B")
        dot_style = "sparse"
        page_color = QColor("#0284C7")
        has_accent_bar = True
    elif style == TocStyle.REGULATORY_LISTING:
        title_text = "Table of Contents"
        accent_color = QColor("#0F172A")
        text_color = QColor("#1E293B")
        dot_style = "dense"
        page_color = QColor("#0F172A")
        has_accent_bar = False
    else:  # MINIMAL_CLEAN
        title_text = "Contents"
        accent_color = QColor("#475569")
        text_color = QColor("#475569")
        dot_style = "none"
        page_color = QColor("#94A3B8")  # 弱化页码
        has_accent_bar = False

    # 2. 绘制标题
    title_font = QFont("sans-serif", 8, QFont.Weight.Bold)
    painter.setFont(title_font)
    painter.setPen(accent_color)
    painter.drawText(QRectF(10, 8, width - 20, 16), Qt.AlignmentFlag.AlignLeft, title_text)

    # 现代风绘制装饰色条
    if has_accent_bar:
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(accent_color)
        painter.drawRoundedRect(QRectF(10, 24, 24, 2), 1, 1)
    else:
        painter.setPen(QPen(QColor("#E2E8F0"), 1))
        painter.drawLine(10, 25, width - 10, 25)

    # 3. 绘制 3 条目录假项
    items = [
        (0, "1 研究概述", "1"),
        (1, "1.1 入组基线", "3"),
        (0, "2 安全性评价", "5"),
    ]

    base_y = 34
    row_height = 18

    for idx, (indent_level, label, page_str) in enumerate(items):
        y = base_y + idx * row_height
        is_level_1 = (indent_level == 0)

        # 字体
        weight = QFont.Weight.DemiBold if is_level_1 else QFont.Weight.Normal
        item_font = QFont("sans-serif", 7, weight)
        painter.setFont(item_font)
        painter.setPen(accent_color if (is_level_1 and style == TocStyle.MODERN_ACCENT) else text_color)

        indent_x = 10 + (indent_level * 10)
        label_rect = QRectF(indent_x, y, 70, 14)
        painter.drawText(label_rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, label)

        # 页码绘制在右侧
        page_rect = QRectF(width - 24, y, 16, 14)
        painter.setFont(QFont("sans-serif", 7, QFont.Weight.Normal))
        painter.setPen(page_color)
        painter.drawText(page_rect, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, page_str)

        # 导线绘制
        leader_start_x = indent_x + 55
        leader_end_x = width - 28
        if dot_style != "none" and leader_end_x > leader_start_x:
            leader_y = y + 9
            if dot_style == "dense":
                dot_pen = QPen(QColor("#64748B"), 1, Qt.PenStyle.DotLine)
            elif dot_style == "sparse":
                dot_pen = QPen(QColor("#CBD5E1"), 1, Qt.PenStyle.DashLine)
            else:  # dot
                dot_pen = QPen(QColor("#94A3B8"), 1, Qt.PenStyle.DotLine)
            painter.setPen(dot_pen)
            painter.drawLine(int(leader_start_x), int(leader_y), int(leader_end_x), int(leader_y))

    painter.end()
    return pixmap


class StyleCardWidget(QFrame):
    """带缩略图预览与高亮选中的 TOC 样式单选卡片控件。"""

    def __init__(self, style: TocStyle, title: str, parent=None) -> None:
        super().__init__(parent)
        self.style = style
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setObjectName("styleCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # 缩略图
        self.preview_label = QLabel(self)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setPixmap(draw_style_preview(style, width=140, height=85))
        layout.addWidget(self.preview_label)

        # 单选按钮
        self.radio = QRadioButton(title, self)
        self.radio.setStyleSheet("font-weight: bold; font-size: 11px;")
        layout.addWidget(self.radio, alignment=Qt.AlignmentFlag.AlignCenter)

        self._update_style(False)
        self.radio.toggled.connect(self._on_toggled)

    def mousePressEvent(self, event) -> None:
        self.radio.setChecked(True)
        super().mousePressEvent(event)

    def isChecked(self) -> bool:
        return self.radio.isChecked()

    def setChecked(self, checked: bool) -> None:
        self.radio.setChecked(checked)
        self._update_style(checked)

    def _on_toggled(self, checked: bool) -> None:
        self._update_style(checked)

    def _update_style(self, checked: bool) -> None:
        if checked:
            self.setStyleSheet(
                """
                QFrame#styleCard {
                    border: 2px solid #0284C7;
                    border-radius: 6px;
                    background-color: #F0F9FF;
                }
                """
            )
        else:
            self.setStyleSheet(
                """
                QFrame#styleCard {
                    border: 1px solid #E2E8F0;
                    border-radius: 6px;
                    background-color: #FFFFFF;
                }
                QFrame#styleCard:hover {
                    border: 1px solid #94A3B8;
                    background-color: #F8FAFC;
                }
                """
            )
