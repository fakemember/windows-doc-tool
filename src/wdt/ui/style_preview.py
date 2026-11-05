"""
TOC 样式缩略图与高清预览模块。

特性：
1. 采用 High-DPI 2x 超采样渲染矢量页面，确保在 Windows 缩放 (125%/150%/200%) 下极致细腻锐利；
2. 真实呈现 4 套样式排版特征（申报密集点线、现代海蓝饰条、经典学术双端、极简无导线）；
3. 提供 PreviewDialog，支持双击/点击放大镜弹出 1:1 A4 比例高保真大图审阅。
"""

from __future__ import annotations

from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QPixmap, QIcon
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from wdt.toc.styles_templates import TocStyle


def draw_style_preview(
    style: TocStyle,
    width: int = 240,
    height: int = 150,
    scale_factor: int = 2,
    is_large: bool = False,
) -> QPixmap:
    """使用 QPainter 渲染指定 TOC 样式的矢量缩略图，默认 2x 超采样消除毛边。"""
    render_w = width * scale_factor
    render_h = height * scale_factor
    pixmap = QPixmap(render_w, render_h)
    pixmap.fill(QColor("#FFFFFF"))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
    painter.scale(scale_factor, scale_factor)

    # 纸张底色与微投影边框
    border_pen = QPen(QColor("#CBD5E1"), 1.2)
    painter.setPen(border_pen)
    painter.setBrush(QColor("#FFFFFF"))
    painter.drawRoundedRect(QRectF(1, 1, width - 2, height - 2), 6, 6)

    # 顶部浅色装订边页眉条（模拟纸张顶边感）
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#F8FAFC"))
    painter.drawRoundedRect(QRectF(2, 2, width - 4, 18), 5, 5)

    # 样式参数映射
    if style == TocStyle.REGULATORY_LISTING:
        title_text = "目    录"
        sub_text = "TABLE OF CONTENTS (CLINICAL TFL)"
        accent_color = QColor("#000000")
        title_font_name = "SimSun"
        item_font_name = "Times New Roman"
        dot_style = "dense"
        page_color = QColor("#000000")
        has_bar = False
        align_center = True
    elif style == TocStyle.MODERN_ACCENT:
        title_text = "目录"
        sub_text = "CONTENTS & OVERVIEW"
        accent_color = QColor("#1F4E79")
        title_font_name = "Microsoft YaHei"
        item_font_name = "Calibri"
        dot_style = "sparse"
        page_color = QColor("#1F4E79")
        has_bar = True
        align_center = False
    elif style == TocStyle.CLASSIC_DOT:
        title_text = "目    录"
        sub_text = "TABLE OF CONTENTS"
        accent_color = QColor("#111827")
        title_font_name = "SimSun"
        item_font_name = "Times New Roman"
        dot_style = "dot"
        page_color = QColor("#111827")
        has_bar = False
        align_center = True
    else:  # MINIMAL_CLEAN
        title_text = "Contents"
        sub_text = "DOCUMENT INDEX"
        accent_color = QColor("#334155")
        title_font_name = "Microsoft YaHei"
        item_font_name = "Arial"
        dot_style = "none"
        page_color = QColor("#64748B")
        has_bar = False
        align_center = False

    # 1. 标题渲染
    font_sz = 11 if not is_large else 14
    title_font = QFont(title_font_name, font_sz, QFont.Weight.Bold)
    painter.setFont(title_font)
    painter.setPen(accent_color)
    
    t_rect = QRectF(14, 22, width - 28, 20)
    flag = Qt.AlignmentFlag.AlignCenter if align_center else Qt.AlignmentFlag.AlignLeft
    painter.drawText(t_rect, flag | Qt.AlignmentFlag.AlignVCenter, title_text)

    # 副标修饰
    sub_font = QFont("sans-serif", 5.5 if not is_large else 8, QFont.Weight.Medium)
    painter.setFont(sub_font)
    painter.setPen(QColor("#94A3B8"))
    sub_rect = QRectF(14, 40, width - 28, 12)
    painter.drawText(sub_rect, flag | Qt.AlignmentFlag.AlignVCenter, sub_text)

    # 现代装饰色条或分隔线
    if has_bar:
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(accent_color)
        painter.drawRoundedRect(QRectF(14, 53, 28, 2.5), 1, 1)
    else:
        painter.setPen(QPen(QColor("#E2E8F0"), 0.8))
        painter.drawLine(14, 54, width - 14, 54)

    # 2. 模拟目录项数据
    items = [
        (0, "1 研究概述与设计 (Study Overview)", "1"),
        (1, "1.1 受试者入组与人口学基线特征", "3"),
        (1, "1.2 给药合规性与暴露量统计", "12"),
        (0, "2 安全性汇总 (Safety Evaluation)", "25"),
        (1, "2.1 不良事件及严重不良反应列表", "28"),
    ]
    if is_large:
        items.extend([
            (2, "2.1.1 导致停药的不良事件", "35"),
            (0, "3 疗效终点分析 (Efficacy Endpoints)", "42"),
        ])

    base_y = 60
    row_h = 16 if not is_large else 22

    for idx, (indent_level, label, page_str) in enumerate(items):
        y = base_y + idx * row_h
        if y + row_h > height - 6:
            break
        is_level_1 = (indent_level == 0)

        # 字体
        weight = QFont.Weight.Bold if is_level_1 else QFont.Weight.Normal
        item_font = QFont(item_font_name, 7.5 if not is_large else 10, weight)
        painter.setFont(item_font)

        if is_level_1 and style == TocStyle.MODERN_ACCENT:
            item_color = accent_color
        elif is_level_1 and style == TocStyle.REGULATORY_LISTING:
            item_color = QColor("#000000")
        else:
            item_color = QColor("#1E293B")
        painter.setPen(item_color)

        indent_px = 14 + (indent_level * 12)
        label_w = width - indent_px - 45
        label_rect = QRectF(indent_px, y, label_w, row_h)
        painter.drawText(label_rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, label)

        # 页码绘制在最右侧
        page_font = QFont(item_font_name, 7.5 if not is_large else 10, QFont.Weight.Normal)
        painter.setFont(page_font)
        painter.setPen(page_color)
        page_rect = QRectF(width - 32, y, 20, row_h)
        painter.drawText(page_rect, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, page_str)

        # 导线绘制
        leader_start_x = indent_px + min(label_w - 10, len(label) * 6 + 10)
        leader_end_x = width - 36
        if dot_style != "none" and leader_end_x > leader_start_x + 10:
            leader_y = y + (row_h / 2.0)
            if dot_style == "dense":
                dot_pen = QPen(QColor("#475569"), 1.0, Qt.PenStyle.DotLine)
            elif dot_style == "sparse":
                dot_pen = QPen(QColor("#CBD5E1"), 1.0, Qt.PenStyle.DashLine)
            else:  # dot
                dot_pen = QPen(QColor("#94A3B8"), 1.0, Qt.PenStyle.DotLine)
            painter.setPen(dot_pen)
            painter.drawLine(QPointF(leader_start_x, leader_y), QPointF(leader_end_x, leader_y))

    painter.end()
    return pixmap


class StylePreviewDialog(QDialog):
    """点击放大后的高清 A4 比例大图预览视窗。"""

    def __init__(self, style: TocStyle, title: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"目录样式大图预览 - {title}")
        self.setMinimumSize(480, 560)
        self.setStyleSheet("background-color: #F1F5F9;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        header = QLabel(f"📄 {title} 真实排版效果示意 (A4 标准页)", self)
        header.setStyleSheet("font-size: 14px; font-weight: 700; color: #0F172A;")
        layout.addWidget(header)

        img_label = QLabel(self)
        img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        large_pixmap = draw_style_preview(style, width=420, height=460, scale_factor=2, is_large=True)
        img_label.setPixmap(large_pixmap.scaled(420, 460, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        layout.addWidget(img_label)

        btn_close = QPushButton("关闭预览", self)
        btn_close.setFixedWidth(100)
        btn_close.setStyleSheet(
            "background-color: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;"
        )
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignCenter)


class StyleCardWidget(QFrame):
    """带 2x 高清缩略图、放大镜查看与选中高亮的 TOC 样式单选卡片。"""

    def __init__(self, style: TocStyle, title: str, parent=None) -> None:
        super().__init__(parent)
        self.style = style
        self.title_str = title
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setObjectName("styleCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # 缩略图容器 (支持右上角放大镜)
        self.preview_label = QLabel(self)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setToolTip("点击卡片选中，点击右上角 🔍 可放大查看完整大图")
        
        # 尺寸增大至 150 x 95，并以 2x 采样缩放保证Retina清晰度
        self._thumb_w = 152
        self._thumb_h = 96
        self._refresh_thumb()
        layout.addWidget(self.preview_label)

        # 底部单选与放大镜按钮行
        bottom_row = QHBoxLayout()
        bottom_row.setContentsMargins(4, 2, 4, 2)
        bottom_row.setSpacing(4)

        self.radio = QRadioButton(title, self)
        self.radio.setStyleSheet("font-weight: 600; font-size: 11px; background: transparent; color: #1E293B;")
        bottom_row.addWidget(self.radio, 1)

        self.btn_zoom = QPushButton("🔍", self)
        self.btn_zoom.setFixedSize(22, 22)
        self.btn_zoom.setToolTip("放大查看高清样式细节")
        self.btn_zoom.setStyleSheet(
            "QPushButton { border: none; background: #E2E8F0; border-radius: 4px; font-size: 10px; }"
            "QPushButton:hover { background: #CBD5E1; }"
        )
        self.btn_zoom.clicked.connect(self._open_zoom_dialog)
        bottom_row.addWidget(self.btn_zoom)

        layout.addLayout(bottom_row)

        self._update_style(False)
        self.radio.toggled.connect(self._on_toggled)

    def _refresh_thumb(self) -> None:
        pix = draw_style_preview(self.style, width=self._thumb_w, height=self._thumb_h, scale_factor=2)
        scaled_pix = pix.scaled(
            self._thumb_w,
            self._thumb_h,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self.preview_label.setPixmap(scaled_pix)

    def _open_zoom_dialog(self) -> None:
        dlg = StylePreviewDialog(self.style, self.title_str, self)
        dlg.exec()

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
                    border: 2px solid #2563EB;
                    border-radius: 8px;
                    background-color: #EFF6FF;
                }
                """
            )
        else:
            self.setStyleSheet(
                """
                QFrame#styleCard {
                    border: 1px solid #E2E8F0;
                    border-radius: 8px;
                    background-color: #FFFFFF;
                }
                QFrame#styleCard:hover {
                    border: 1px solid #94A3B8;
                    background-color: #F8FAFC;
                }
                """
            )
