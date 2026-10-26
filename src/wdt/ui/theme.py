"""
全局浅色主题定义。

集中管理配色与 Qt 样式表，保证整个界面视觉一致。仅提供浅色主题。
"""

# 配色体系
COLOR_BG = "#F3F4F6"  # 窗口背景
COLOR_CARD = "#FFFFFF"  # 卡片背景
COLOR_BORDER = "#E5E7EB"  # 卡片/控件边框
COLOR_BORDER_STRONG = "#D1D5DB"  # 输入框边框
COLOR_TEXT = "#1F2937"  # 主文字
COLOR_TEXT_TITLE = "#111827"  # 标题文字
COLOR_TEXT_SECONDARY = "#6B7280"  # 次要文字
COLOR_TEXT_MUTED = "#9CA3AF"  # 弱化文字
COLOR_ACCENT = "#2563EB"  # 主强调色
COLOR_ACCENT_HOVER = "#1D4ED8"
COLOR_ACCENT_SOFT = "#EFF6FF"  # 强调色浅底
COLOR_SUCCESS = "#16A34A"
COLOR_DANGER = "#DC2626"
COLOR_WARNING = "#D97706"
COLOR_TABLE_STRIPE = "#F9FAFB"
COLOR_ROW_HOVER = "#F3F4F6"
COLOR_INPUT_BG = "#FFFFFF"
COLOR_ROW_SELECTED = "#EFF6FF"

STATUS_COLORS = {
    "queued": COLOR_TEXT_MUTED,
    "ok": COLOR_SUCCESS,
    "failed": COLOR_DANGER,
    "skipped": COLOR_WARNING,
    "cancelled": COLOR_TEXT_MUTED,
}

APP_STYLESHEET = f"""
QMainWindow {{
    background-color: {COLOR_BG};
}}
QWidget {{
    font-family: "Segoe UI", "Microsoft YaHei UI", "PingFang SC", sans-serif;
    font-size: 12px;
    color: {COLOR_TEXT};
}}
QToolTip {{
    background-color: {COLOR_TEXT_TITLE};
    color: #FFFFFF;
    border: none;
    padding: 4px 8px;
}}

/* ---------- 卡片容器 ---------- */
QFrame#headerBar {{
    background-color: {COLOR_CARD};
    border-bottom: 1px solid {COLOR_BORDER};
}}
QFrame#sidebarCard {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER};
    border-radius: 10px;
}}
QFrame#contentCard {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER};
    border-radius: 10px;
}}
QFrame#settingsCard {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER};
    border-radius: 10px;
}}
QFrame#actionBar {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER};
    border-radius: 10px;
}}

/* ---------- 侧边栏 ---------- */
QLabel#sidebarSection {{
    color: {COLOR_TEXT_MUTED};
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 1px;
}}
QLabel#sidebarFooter {{
    color: {COLOR_TEXT_MUTED};
    font-size: 11px;
}}

/* ---------- 拖拽投放区 ---------- */
QFrame#dropArea {{
    border: 1.5px dashed {COLOR_BORDER_STRONG};
    border-radius: 10px;
    background-color: {COLOR_CARD};
}}
QFrame#dropArea:hover {{
    border-color: {COLOR_ACCENT};
    background-color: {COLOR_ACCENT_SOFT};
}}

/* ---------- 表格 ---------- */
QTableWidget {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER};
    border-radius: 8px;
    gridline-color: {COLOR_BG};
    selection-background-color: {COLOR_ROW_SELECTED};
    selection-color: {COLOR_ACCENT_HOVER};
}}
QTableWidget::item {{
    padding: 5px 6px;
    border: none;
}}
QTableWidget::item:hover {{
    background-color: {COLOR_ROW_HOVER};
}}
QHeaderView::section {{
    background-color: {COLOR_TABLE_STRIPE};
    color: {COLOR_TEXT_SECONDARY};
    font-weight: 600;
    border: none;
    border-bottom: 1px solid {COLOR_BORDER};
    padding: 8px 6px;
}}
QTableCornerButton::section {{
    background-color: {COLOR_TABLE_STRIPE};
    border: none;
    border-bottom: 1px solid {COLOR_BORDER};
}}

/* ---------- 按钮 ---------- */
QPushButton {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER_STRONG};
    border-radius: 6px;
    padding: 6px 14px;
    color: {COLOR_TEXT};
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {COLOR_TABLE_STRIPE};
    border-color: {COLOR_TEXT_MUTED};
}}
QPushButton:pressed {{
    background-color: {COLOR_BORDER};
}}
QPushButton:disabled {{
    color: {COLOR_TEXT_MUTED};
    border-color: {COLOR_BORDER};
    background-color: {COLOR_TABLE_STRIPE};
}}
QPushButton#btnStart {{
    background-color: {COLOR_ACCENT};
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 8px 26px;
    font-size: 14px;
    font-weight: 600;
}}
QPushButton#btnStart:hover {{
    background-color: {COLOR_ACCENT_HOVER};
    border: none;
}}
QPushButton#btnStart:disabled {{
    background-color: #93C5FD;
    color: #FFFFFF;
    border: none;
}}
QPushButton#btnCancel {{
    background-color: {COLOR_DANGER};
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 8px 18px;
    font-weight: 600;
}}
QPushButton#btnCancel:hover {{
    background-color: #B91C1C;
    border: none;
}}
QPushButton#btnCancel:disabled {{
    background-color: {COLOR_BORDER};
    color: {COLOR_TEXT_MUTED};
    border: none;
}}

/* ---------- 输入控件 ---------- */
QLineEdit, QComboBox, QSpinBox {{
    background-color: {COLOR_INPUT_BG};
    border: 1px solid {COLOR_BORDER_STRONG};
    border-radius: 6px;
    padding: 6px 8px;
    selection-background-color: {COLOR_ACCENT};
    selection-color: #FFFFFF;
}}
QLineEdit:hover, QComboBox:hover, QSpinBox:hover {{
    border-color: {COLOR_TEXT_MUTED};
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{
    border: 1px solid {COLOR_ACCENT};
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox QAbstractItemView {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER};
    selection-background-color: {COLOR_ROW_SELECTED};
    selection-color: {COLOR_ACCENT_HOVER};
    outline: none;
}}
QCheckBox, QRadioButton {{
    spacing: 6px;
    color: {COLOR_TEXT};
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 15px;
    height: 15px;
}}

/* ---------- 分组卡内元素 ---------- */
QLabel#cardTitle {{
    color: {COLOR_TEXT_TITLE};
    font-size: 13px;
    font-weight: 600;
}}
QLabel#formLabel {{
    color: {COLOR_TEXT_SECONDARY};
    font-size: 12px;
}}
QLabel#hintLabel {{
    color: {COLOR_TEXT_MUTED};
    font-size: 11px;
}}

/* ---------- 滚动区与滚动条 ---------- */
QScrollArea#settingsScroll {{
    border: none;
    background-color: transparent;
}}
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px 2px 2px 0;
}}
QScrollBar::handle:vertical {{
    background: {COLOR_BORDER_STRONG};
    border-radius: 5px;
    min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{
    background: {COLOR_TEXT_MUTED};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 0 2px 2px 2px;
}}
QScrollBar::handle:horizontal {{
    background: {COLOR_BORDER_STRONG};
    border-radius: 5px;
    min-width: 30px;
}}
QScrollBar::handle:horizontal:hover {{
    background: {COLOR_TEXT_MUTED};
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

/* ---------- 进度条 ---------- */
QProgressBar {{
    border: 1px solid {COLOR_BORDER};
    border-radius: 6px;
    background-color: {COLOR_TABLE_STRIPE};
    text-align: center;
    color: {COLOR_TEXT_SECONDARY};
    min-height: 20px;
}}
QProgressBar::chunk {{
    background-color: {COLOR_ACCENT};
    border-radius: 5px;
}}

/* ---------- 状态栏 ---------- */
QStatusBar {{
    background-color: {COLOR_BG};
    color: {COLOR_TEXT_SECONDARY};
    border: none;
}}
"""
