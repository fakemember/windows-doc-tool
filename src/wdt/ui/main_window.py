"""
TFLOW 现代化主界面与交互中枢。
三阶控制台（交付目标、目录编排、输出配置）、专业场景预设卡片与沉浸式执行反馈。
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Sequence

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDropEvent, QFont, QIcon, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from wdt.core.engine_finder import WordEngine, find_word_engine
from wdt.pipeline import (
    PipelineEvent,
    PipelineOptions,
    Stage,
)
from wdt.toc.docx_stream_patcher import Position, TocConfig
from wdt.toc.page_calculator import PageMode
from wdt.toc.styles_templates import TocStyle
from wdt.ui.style_preview import StyleCardWidget
from wdt.ui.theme import APP_STYLESHEET, STATUS_COLORS
from wdt.ui.worker import PipelineWorker

SUPPORTED_EXTS = (".doc", ".rtf", ".docx")


class DropArea(QFrame):
    """宽适型文件与目录投放区。"""

    def __init__(self, main_win: MainWindow) -> None:
        super().__init__(main_win)
        self.main_win = main_win
        self.setAcceptDrops(True)
        self.setObjectName("dropArea")
        self.setMinimumHeight(68)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(14)

        icon_label = QLabel("📥", self)
        icon_label.setFont(QFont("Segoe UI Emoji", 20))
        icon_label.setFixedWidth(30)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        title_label = QLabel("拖拽待处理文档或目录到此处", self)
        title_label.setStyleSheet("font-size: 13px; font-weight: 700; color: #0F172A;")
        hint_label = QLabel("支持 DOC / RTF / DOCX 格式，自动递归展开目录并智能过滤临时锁文件 (~$*)", self)
        hint_label.setStyleSheet("font-size: 11px; color: #64748B;")
        text_layout.addWidget(title_label)
        text_layout.addWidget(hint_label)
        layout.addLayout(text_layout)

        layout.addStretch()

        btn_files = QPushButton("＋ 添加文件…", self)
        btn_files.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_files.clicked.connect(self.main_win.on_add_files_dialog)
        layout.addWidget(btn_files)

        btn_dir = QPushButton("＋ 添加目录…", self)
        btn_dir.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_dir.clicked.connect(self.main_win.on_add_dir_dialog)
        layout.addWidget(btn_dir)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                """
                QFrame#dropArea {
                    border: 2px dashed #2563EB;
                    background-color: #EFF6FF;
                }
                """
            )

    def dragLeaveEvent(self, event) -> None:
        self.setStyleSheet("")

    def dropEvent(self, event: QDropEvent) -> None:
        self.dragLeaveEvent(event)
        urls: list[QUrl] = event.mimeData().urls()
        paths: list[Path] = []
        for u in urls:
            local = u.toLocalFile()
            if local:
                paths.append(Path(local))
        if paths:
            self.main_win.add_paths(paths)


class MainWindow(QMainWindow):
    """TFLOW 临床 TFL 报告与文档批处理流水线主界面。"""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("TFLOW - 临床 TFL 报告与文档批处理流水线")
        self.resize(1200, 800)
        self.setMinimumSize(1000, 680)

        assets_dir = Path(__file__).parent / "assets"
        logo_ico = assets_dir / "tflow_logo.ico"
        logo_svg = assets_dir / "tflow_logo.svg"
        if logo_ico.is_file():
            self.setWindowIcon(QIcon(str(logo_ico)))
        elif logo_svg.is_file():
            self.setWindowIcon(QIcon(str(logo_svg)))

        self.files_list: list[Path] = []
        self.file_row_map: dict[Path, int] = {}
        self.worker: PipelineWorker | None = None
        self.word_engine: WordEngine | None = None

        self._init_engine()
        self._init_ui()
        self._update_engine_pill()
        self._apply_global_theme()

    def _init_engine(self) -> None:
        self.word_engine = find_word_engine()

    def _init_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(14, 10, 14, 12)
        root_layout.setSpacing(10)

        # 1. 顶部 Header
        root_layout.addWidget(self._build_header_bar())

        # 2. 拖拽投递栏
        self.drop_area = DropArea(self)
        root_layout.addWidget(self.drop_area)

        # 3. 中部核心区：左侧文档表格 + 右侧交付控制台
        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.setChildrenCollapsible(False)

        splitter.addWidget(self._build_table_card(splitter))
        splitter.addWidget(self._build_control_panel(splitter))

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([680, 440])
        root_layout.addWidget(splitter, 1)

        # 4. 底部执行中枢
        root_layout.addWidget(self._build_action_bar())

        # 底部状态栏
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪。请添加或拖拽文档开始处理。")

        # 兼顾向后兼容与测试绑定的虚拟模式 RadioButtons
        self.radio_mode_run = QRadioButton("run", self)
        self.radio_mode_conv = QRadioButton("convert", self)
        self.radio_mode_toc = QRadioButton("toc", self)
        self.radio_mode_run.setChecked(True)
        for w in (self.radio_mode_run, self.radio_mode_conv, self.radio_mode_toc):
            w.setVisible(False)

    def _build_header_bar(self) -> QFrame:
        header_bar = QFrame(self)
        header_bar.setObjectName("headerBar")
        header_layout = QHBoxLayout(header_bar)
        header_layout.setContentsMargins(16, 8, 16, 8)
        header_layout.setSpacing(12)

        brand_badge = QLabel(header_bar)
        brand_badge.setFixedSize(38, 38)
        brand_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_svg = Path(__file__).parent / "assets" / "tflow_logo.svg"
        if logo_svg.is_file():
            logo_pix = QIcon(str(logo_svg)).pixmap(38, 38)
            brand_badge.setPixmap(logo_pix)
            brand_badge.setScaledContents(True)
        else:
            brand_badge.setText("T")
            brand_badge.setStyleSheet(
                "background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2563EB, stop:1 #06B6D4); "
                "color: #FFFFFF; font-size: 19px; font-weight: 800; border-radius: 8px;"
            )
        header_layout.addWidget(brand_badge)

        brand_text = QVBoxLayout()
        brand_text.setSpacing(1)
        brand_title = QLabel("TFLOW", header_bar)
        brand_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0F172A; letter-spacing: 0.5px;")
        brand_sub = QLabel("Clinical TFL & Document Pipeline · 纯原生 Word COM 自动化引擎", header_bar)
        brand_sub.setStyleSheet("font-size: 11px; color: #64748B; font-weight: 500;")
        brand_text.addWidget(brand_title)
        brand_text.addWidget(brand_sub)
        header_layout.addLayout(brand_text)

        header_layout.addStretch()

        self.engine_pill = QLabel(header_bar)
        self.engine_pill.setCursor(Qt.CursorShape.PointingHandCursor)
        header_layout.addWidget(self.engine_pill)
        self.btn_recheck_engine = QPushButton("重新检测", header_bar)
        self.btn_recheck_engine.clicked.connect(self._on_recheck_engine)
        header_layout.addWidget(self.btn_recheck_engine)

        return header_bar

    def _build_table_card(self, parent: QWidget) -> QFrame:
        table_card = QFrame(parent)
        table_card.setObjectName("contentCard")
        layout = QVBoxLayout(table_card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        header_layout = QHBoxLayout()
        title = QLabel("待交付文档清单", table_card)
        title.setObjectName("cardTitle")
        header_layout.addWidget(title)

        self.label_file_count = QLabel("0 个文档", table_card)
        self.label_file_count.setObjectName("hintLabel")
        header_layout.addWidget(self.label_file_count)
        header_layout.addStretch()

        self.btn_remove = QPushButton("移除选中", table_card)
        self.btn_remove.clicked.connect(self.on_remove_selected)
        header_layout.addWidget(self.btn_remove)

        self.btn_clear = QPushButton("清空列表", table_card)
        self.btn_clear.clicked.connect(self.on_clear_files)
        header_layout.addWidget(self.btn_clear)
        layout.addLayout(header_layout)

        self.table = QTableWidget(table_card)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["文档名称", "大小", "处理状态", "耗时", "文件路径"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 220)
        self.table.setColumnWidth(2, 160)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        return table_card

    def _build_control_panel(self, parent: QWidget) -> QScrollArea:
        """右侧一体化交付控制台（三阶结构化卡片）。"""
        scroll = QScrollArea(parent)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(10)

        # 阶层 1: 交付目标
        layout.addWidget(self._build_deliverables_card(container))

        # 阶层 2: 目录排版与编排规范
        layout.addWidget(self._build_toc_settings_card(container))

        # 阶层 3: 输出目标设置
        layout.addWidget(self._build_output_card(container))

        # 隐藏的并发数控（向后兼容测试）
        self.spin_jobs = QSpinBox(container)
        self.spin_jobs.setRange(1, 1)
        self.spin_jobs.setValue(1)
        self.spin_jobs.setVisible(False)

        layout.addStretch()
        scroll.setWidget(container)
        return scroll

    def _build_deliverables_card(self, parent: QWidget) -> QFrame:
        card = QFrame(parent)
        card.setObjectName("settingsCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        title = QLabel("1. 目标交付物 (Deliverables)", card)
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        grid = QGridLayout()
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(8)

        self.chk_docx = QCheckBox("生成规范化 DOCX", card)
        self.chk_docx.setChecked(True)
        self.chk_docx.setStyleSheet("font-weight: 600; font-size: 13px;")

        self.chk_pdf = QCheckBox("导出最终发布 PDF", card)
        self.chk_pdf.setChecked(True)
        self.chk_pdf.setStyleSheet("font-weight: 600; font-size: 13px; color: #1D4ED8;")

        self.chk_gray = QCheckBox("灰度打印优化", card)
        self.chk_gray.setToolTip("以灰度色调进行排版导出")

        self.chk_force = QCheckBox("覆盖已有目标文件", card)
        self.chk_force.setChecked(True)

        grid.addWidget(self.chk_docx, 0, 0)
        grid.addWidget(self.chk_pdf, 0, 1)
        grid.addWidget(self.chk_gray, 1, 0)
        grid.addWidget(self.chk_force, 1, 1)
        layout.addLayout(grid)

        # PDF 模式选择区
        self.pdf_mode_box = QFrame(card)
        self.pdf_mode_box.setStyleSheet(
            "background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 4px;"
        )
        pdf_mode_layout = QVBoxLayout(self.pdf_mode_box)
        pdf_mode_layout.setContentsMargins(8, 6, 8, 6)
        pdf_mode_layout.setSpacing(4)

        mode_header = QLabel("PDF 导出与页码校准模式：", self.pdf_mode_box)
        mode_header.setStyleSheet("font-size: 11px; font-weight: 600; color: #475569;")
        pdf_mode_layout.addWidget(mode_header)

        mode_btn_layout = QHBoxLayout()
        mode_btn_layout.setSpacing(12)

        self.radio_pdf_fast = QRadioButton("⚡ 极速审阅模式 (直出秒级耗时·推荐草稿)", self.pdf_mode_box)
        self.radio_pdf_fast.setChecked(True)
        self.radio_pdf_fast.setToolTip("基于 DOCX 静态版面快速输出，跳过全量长表格重排，耗时缩短 60%+")

        self.radio_pdf_strict = QRadioButton("🛡️ 申报对齐模式 (Word逐页深度校准·提交级)", self.pdf_mode_box)
        self.radio_pdf_strict.setToolTip("Word 引擎逐页物理校准，计算自然跨页表格与精准页码，大文档需较长时间但页码 100% 对齐")

        self.pdf_mode_group = QButtonGroup(self)
        self.pdf_mode_group.addButton(self.radio_pdf_fast, 1)
        self.pdf_mode_group.addButton(self.radio_pdf_strict, 2)

        mode_btn_layout.addWidget(self.radio_pdf_fast)
        mode_btn_layout.addWidget(self.radio_pdf_strict)
        pdf_mode_layout.addLayout(mode_btn_layout)

        layout.addWidget(self.pdf_mode_box)
        self.chk_pdf.toggled.connect(self.pdf_mode_box.setVisible)

        return card

    def _build_toc_settings_card(self, parent: QWidget) -> QFrame:
        self.toc_card = QFrame(parent)
        self.toc_card.setObjectName("settingsCard")
        layout = QVBoxLayout(self.toc_card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        header_row = QHBoxLayout()
        title = QLabel("2. 目录排版规范 (TOC Enhancement)", self.toc_card)
        title.setObjectName("cardTitle")
        header_row.addWidget(title)
        header_row.addStretch()

        self.chk_enable_toc = QCheckBox("启用目录", self.toc_card)
        self.chk_enable_toc.setChecked(True)
        self.chk_enable_toc.toggled.connect(self._on_toc_toggle)
        header_row.addWidget(self.chk_enable_toc)
        layout.addLayout(header_row)

        self.toc_body_widget = QWidget(self.toc_card)
        body_layout = QVBoxLayout(self.toc_body_widget)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(10)

        # 4 款样式卡片
        cards_grid = QGridLayout()
        cards_grid.setContentsMargins(0, 0, 0, 0)
        cards_grid.setSpacing(8)

        self.card_classic = StyleCardWidget(TocStyle.CLASSIC_DOT, "经典点线 (CSR申报)", self.toc_body_widget)
        self.card_modern = StyleCardWidget(TocStyle.MODERN_ACCENT, "现代商务 (推荐)", self.toc_body_widget)
        self.card_regulatory = StyleCardWidget(TocStyle.REGULATORY_LISTING, "临床紧凑 (TFL汇编)", self.toc_body_widget)
        self.card_minimal = StyleCardWidget(TocStyle.MINIMAL_CLEAN, "极简纯净", self.toc_body_widget)

        self.style_group = QButtonGroup(self)
        self.style_group.addButton(self.card_classic.radio, 1)
        self.style_group.addButton(self.card_modern.radio, 2)
        self.style_group.addButton(self.card_regulatory.radio, 3)
        self.style_group.addButton(self.card_minimal.radio, 4)
        self.card_classic.setChecked(True)

        cards_grid.addWidget(self.card_classic, 0, 0)
        cards_grid.addWidget(self.card_modern, 0, 1)
        cards_grid.addWidget(self.card_regulatory, 1, 0)
        cards_grid.addWidget(self.card_minimal, 1, 1)
        body_layout.addLayout(cards_grid)

        # 参数网格
        params_grid = QGridLayout()
        params_grid.setHorizontalSpacing(10)
        params_grid.setVerticalSpacing(8)

        params_grid.addWidget(self._form_label("目录标题文案"), 0, 0)
        self.edit_toc_title = QLineEdit("目  录", self.toc_body_widget)
        params_grid.addWidget(self.edit_toc_title, 1, 0)

        params_grid.addWidget(self._form_label("标题收集层级"), 0, 1)
        self.combo_depth = QComboBox(self.toc_body_widget)
        self.combo_depth.addItems(["仅 1 级大标题", "1 ~ 2 级标题", "1 ~ 3 级全量标题"])
        self.combo_depth.setCurrentIndex(2)  # 默认 1-3 级
        params_grid.addWidget(self.combo_depth, 1, 1)

        params_grid.addWidget(self._form_label("插入位置"), 2, 0)
        self.combo_position = QComboBox(self.toc_body_widget)
        self.combo_position.addItems(["文档开头", "首个标题前", "原位替换占位符"])
        self.combo_position.setCurrentIndex(1)  # 默认首个标题前
        params_grid.addWidget(self.combo_position, 3, 0)

        params_grid.addWidget(self._form_label("页码重算策略"), 2, 1)
        self.combo_page_mode = QComboBox(self.toc_body_widget)
        self.combo_page_mode.addItems([
            "静态分节秒开 (首屏极速·正文从1重新编页·推荐)",
            "静态页码偏移",
            "静态原位替换",
            "实测页码",
            "Word 动态域 (打开自动更新字段·大文件卡顿)",
        ])
        self.combo_page_mode.setCurrentIndex(0)
        params_grid.addWidget(self.combo_page_mode, 3, 1)

        offset_row = QHBoxLayout()
        offset_row.setSpacing(8)
        self.spin_offset = QSpinBox(self.toc_body_widget)
        self.spin_offset.setRange(-999, 9999)
        self.spin_offset.setValue(0)
        offset_row.addWidget(self.spin_offset, 1)

        self.chk_backup = QCheckBox("保留原始备份 (.bak)", self.toc_body_widget)
        self.chk_backup.setChecked(False)
        offset_row.addWidget(self.chk_backup)

        params_grid.addWidget(self._form_label("起始页码偏移"), 4, 0)
        params_grid.addLayout(offset_row, 5, 0, 1, 2)

        body_layout.addLayout(params_grid)
        layout.addWidget(self.toc_body_widget)
        return self.toc_card

    def _build_output_card(self, parent: QWidget) -> QFrame:
        self.output_card = QFrame(parent)
        self.output_card.setObjectName("settingsCard")
        layout = QVBoxLayout(self.output_card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        title = QLabel("3. 输出归档路径 (Destination)", self.output_card)
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        outdir_layout = QHBoxLayout()
        outdir_layout.setSpacing(8)
        self.edit_outdir = QLineEdit(self.output_card)
        self.edit_outdir.setPlaceholderText("留空则保存在源文件同级目录")
        outdir_layout.addWidget(self.edit_outdir, 1)

        btn_browse_out = QPushButton("浏览…", self.output_card)
        btn_browse_out.clicked.connect(self._on_browse_outdir)
        outdir_layout.addWidget(btn_browse_out)
        layout.addLayout(outdir_layout)

        return self.output_card

    def _build_action_bar(self) -> QFrame:
        action_bar = QFrame(self)
        action_bar.setObjectName("actionBar")
        layout = QHBoxLayout(action_bar)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(14)

        progress_col = QVBoxLayout()
        progress_col.setSpacing(4)
        self.progress_bar = QProgressBar(action_bar)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(22)

        self.label_progress = QLabel("准备就绪。添加或拖拽文档后，点击右侧「开始处理」即可启动流水线。", action_bar)
        self.label_progress.setObjectName("hintLabel")
        progress_col.addWidget(self.progress_bar)
        progress_col.addWidget(self.label_progress)
        layout.addLayout(progress_col, 1)

        self.btn_cancel = QPushButton("⏹ 中止", action_bar)
        self.btn_cancel.setObjectName("btnCancel")
        self.btn_cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cancel.setFixedHeight(42)
        self.btn_cancel.setFixedWidth(100)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.on_cancel_pipeline)
        layout.addWidget(self.btn_cancel)

        self.btn_start = QPushButton("▶ 开始处理", action_bar)
        self.btn_start.setObjectName("btnStart")
        self.btn_start.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_start.setFixedHeight(42)
        self.btn_start.setFixedWidth(140)
        self.btn_start.clicked.connect(self.on_start_pipeline)
        layout.addWidget(self.btn_start)

        return action_bar

    def _form_label(self, text: str) -> QLabel:
        label = QLabel(text, self)
        label.setObjectName("formLabel")
        return label

    def _update_engine_pill(self) -> None:
        if self.word_engine is not None and self.word_engine.available:
            ver = self.word_engine.version or "已就绪"
            self.engine_pill.setText(f" ● Microsoft Word {ver} ")
            self.engine_pill.setToolTip(f"Word COM 引擎已就绪: {self.word_engine.app_path or '系统注册表调度'}")
            self.engine_pill.setStyleSheet(
                "background-color: #ECFDF5; color: #047857; "
                "border: 1px solid #A7F3D0; border-radius: 10px; padding: 4px 10px; font-size: 12px; font-weight: 600;"
            )
            self.btn_recheck_engine.setVisible(False)
        else:
            self.engine_pill.setText(" ⚠ 未检测到 Word COM 引擎 ")
            self.engine_pill.setToolTip(
                "未检测到 Microsoft Word 或 pywin32。\n"
                "请安装 Microsoft Word 后点击右侧「重新检测」。"
            )
            self.engine_pill.setStyleSheet(
                "background-color: #FEF2F2; color: #B91C1C; "
                "border: 1px solid #FECACA; border-radius: 10px; padding: 4px 10px; font-size: 12px; font-weight: 600;"
            )
            self.btn_recheck_engine.setVisible(True)

    def _on_recheck_engine(self) -> None:
        self._init_engine()
        self._update_engine_pill()

    def _on_toc_toggle(self, checked: bool) -> None:
        self.toc_body_widget.setVisible(checked)

    def _on_browse_outdir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择输出归档目录")
        if d:
            self.edit_outdir.setText(d)

    def on_add_files_dialog(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择待处理文档",
            "",
            "支持的文档 (*.doc *.rtf *.docx);;所有文件 (*.*)",
        )
        if files:
            self.add_paths([Path(f) for f in files])

    def on_add_dir_dialog(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择待扫描目录")
        if d:
            self.add_paths([Path(d)])

    def on_clear_files(self) -> None:
        self.files_list.clear()
        self.file_row_map.clear()
        self.table.setRowCount(0)
        self.progress_bar.setValue(0)
        self._refresh_file_count()
        self.label_progress.setText("已清空列表")
        self.status_bar.showMessage("已清空列表")

    def on_remove_selected(self) -> None:
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            if r < len(self.files_list):
                self.files_list.pop(r)
                self.table.removeRow(r)
        self.file_row_map.clear()
        for i, f in enumerate(self.files_list):
            self.file_row_map[f] = i
        self._refresh_file_count()

    def _refresh_file_count(self) -> None:
        self.label_file_count.setText(f"{len(self.files_list)} 个文档")

    def add_paths(self, paths: Sequence[Path | str]) -> None:
        new_files: list[Path] = []
        for p in paths:
            path_obj = Path(p).resolve()
            if not path_obj.exists():
                continue
            if path_obj.is_file():
                if (
                    path_obj.suffix.lower() in SUPPORTED_EXTS
                    and not path_obj.name.startswith("~$")
                    and not path_obj.name.startswith("._")
                ):
                    if path_obj not in self.files_list and path_obj not in new_files:
                        new_files.append(path_obj)
            elif path_obj.is_dir():
                for ext in SUPPORTED_EXTS:
                    for found in path_obj.rglob(f"*{ext}"):
                        if (
                            found.is_file()
                            and not found.name.startswith("~$")
                            and not found.name.startswith("._")
                        ):
                            res = found.resolve()
                            if res not in self.files_list and res not in new_files:
                                new_files.append(res)

        if not new_files:
            return

        for f in new_files:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.files_list.append(f)
            self.file_row_map[f] = row

            item_name = QTableWidgetItem(f.name)
            try:
                size_kb = f.stat().st_size / 1024
                size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
            except Exception:
                size_str = "未知"
            item_size = QTableWidgetItem(size_str)
            item_status = QTableWidgetItem("等待处理")
            item_elapsed = QTableWidgetItem("-")
            item_path = QTableWidgetItem(str(f.parent))

            item_status.setForeground(QColor(STATUS_COLORS["queued"]))

            self.table.setItem(row, 0, item_name)
            self.table.setItem(row, 1, item_size)
            self.table.setItem(row, 2, item_status)
            self.table.setItem(row, 3, item_elapsed)
            self.table.setItem(row, 4, item_path)

        self._refresh_file_count()
        self.status_bar.showMessage(f"已加载 {len(self.files_list)} 个待交付文档")

    def get_pipeline_options(self) -> PipelineOptions:
        """从当前 UI 状态提炼交付导向的 PipelineOptions 对象。"""
        outdir_text = self.edit_outdir.text().strip()
        out_dir = Path(outdir_text) if outdir_text else None

        style = TocStyle.CLASSIC_DOT
        checked_style_id = self.style_group.checkedId()
        if checked_style_id == 2:
            style = TocStyle.MODERN_ACCENT
        elif checked_style_id == 3:
            style = TocStyle.REGULATORY_LISTING
        elif checked_style_id == 4:
            style = TocStyle.MINIMAL_CLEAN

        depth = self.combo_depth.currentIndex() + 1
        toc_title = self.edit_toc_title.text().strip() or "目  录"

        pos_idx = self.combo_position.currentIndex()
        pos_map = {
            0: Position.DOCUMENT_START,
            1: Position.BEFORE_FIRST_HEADING,
            2: Position.REPLACE_PLACEHOLDER,
        }
        position = pos_map.get(pos_idx, Position.DOCUMENT_START)

        page_mode_idx = self.combo_page_mode.currentIndex()
        page_mode_map = {
            0: PageMode.STATIC_RESTART,
            1: PageMode.STATIC_OFFSET,
            2: PageMode.STATIC_IN_PLACE,
            3: PageMode.STATIC_MEASURED,
            4: PageMode.DYNAMIC,
        }
        page_mode = page_mode_map.get(page_mode_idx, PageMode.STATIC_RESTART)

        toc_config = TocConfig(
            style=style,
            depth=depth,
            title=toc_title,
            position=position,
            page_mode=page_mode,
            page_offset=self.spin_offset.value(),
            backup=self.chk_backup.isChecked(),
        )

        # 模式自适应判定：兼容单元测试以及纯交付目标驱动
        if self.radio_mode_conv.isChecked() or (not self.chk_enable_toc.isChecked()):
            mode = "convert"
        elif self.radio_mode_toc.isChecked():
            mode = "toc"
        else:
            mode = "run"

        pdf_mode: Literal["fast", "strict"] = (
            "strict" if self.radio_pdf_strict.isChecked() else "fast"
        )

        return PipelineOptions(
            mode=mode,  # type: ignore[arg-type]
            out_dir=out_dir,
            make_docx=self.chk_docx.isChecked(),
            make_pdf=self.chk_pdf.isChecked(),
            gray=self.chk_gray.isChecked(),
            toc=toc_config,
            jobs=self.spin_jobs.value(),
            force=self.chk_force.isChecked(),
            pdf_mode=pdf_mode,
        )

    def _set_running_state(self, running: bool) -> None:
        self.btn_start.setEnabled(not running)
        self.btn_cancel.setEnabled(running)
        self.btn_cancel.setText("⏹ 中止")
        self.drop_area.setEnabled(not running)
        self.btn_clear.setEnabled(not running)
        self.btn_remove.setEnabled(not running)
        self.output_card.setEnabled(not running)
        self.toc_card.setEnabled(not running)

    def on_start_pipeline(self) -> None:
        if not self.files_list:
            QMessageBox.information(self, "提示", "请先添加或拖入待处理的文档。")
            return

        try:
            options = self.get_pipeline_options()
        except Exception as e:
            QMessageBox.warning(self, "参数错误", str(e))
            return

        if self.word_engine is None or not self.word_engine.available:
            QMessageBox.critical(
                self,
                "缺少 Word COM 引擎",
                f"未检测到 Microsoft Word COM 自动化引擎 ({self.word_engine.error if self.word_engine else '未安装'})\n"
                "请安装 Microsoft Word 并确保 pywin32 正常后重试。",
            )
            return

        self._set_running_state(True)

        for row in range(self.table.rowCount()):
            item_st = self.table.item(row, 2)
            if item_st:
                item_st.setText("等待处理")
                item_st.setForeground(QColor(STATUS_COLORS["queued"]))
                item_st.setToolTip("")
            item_el = self.table.item(row, 3)
            if item_el:
                item_el.setText("-")

        events_per_file = 3 if options.make_pdf else 2
        total_steps = max(1, len(self.files_list) * events_per_file)
        self.progress_bar.setRange(0, total_steps)
        self.progress_bar.setValue(0)
        self.label_progress.setText(f"正在启动 Word COM 排版流水线 (0 / {total_steps})…")
        self.status_bar.showMessage(f"正在处理，共 {len(self.files_list)} 个文档")

        self.worker = PipelineWorker(
            files=self.files_list,
            options=options,
            parent=self,
        )
        self.worker.event_signal.connect(self._on_pipeline_event)
        self.worker.finished_signal.connect(self._on_pipeline_finished)
        self.worker.error_signal.connect(self._on_pipeline_error)
        self.worker.start()

    def on_cancel_pipeline(self) -> None:
        if self.worker and self.worker.isRunning():
            self.btn_cancel.setEnabled(False)
            self.btn_cancel.setText("正在中止…")
            self.status_bar.showMessage("正在中止，请稍候…")
            self.worker.cancel()

    def _on_pipeline_event(self, ev: PipelineEvent) -> None:
        row = self.file_row_map.get(ev.src)
        if row is None:
            for src, r in self.file_row_map.items():
                if src.resolve() == ev.src.resolve():
                    row = r
                    break

        if row is not None:
            stage_prefix = {
                Stage.CONVERT: "格式转换",
                Stage.TOC: "目录编排",
                Stage.PDF: "PDF导出",
            }.get(ev.stage, ev.stage.value)

            item_status = self.table.item(row, 2)
            item_elapsed = self.table.item(row, 3)

            if item_status:
                if ev.status == "running":
                    item_status.setText(f"⏳ {stage_prefix}中…")
                    item_status.setForeground(QColor(STATUS_COLORS["queued"]))
                elif ev.status == "ok":
                    item_status.setText(f"✓ {stage_prefix}完成")
                    item_status.setForeground(QColor(STATUS_COLORS["ok"]))
                elif ev.status == "failed":
                    item_status.setText(f"✗ {stage_prefix}失败")
                    item_status.setForeground(QColor(STATUS_COLORS["failed"]))
                    item_status.setToolTip(ev.message)
                elif ev.status == "skipped":
                    item_status.setText("已跳过")
                    item_status.setForeground(QColor(STATUS_COLORS["skipped"]))
                    item_status.setToolTip(ev.message)
                elif ev.status == "cancelled":
                    item_status.setText("已取消")
                    item_status.setForeground(QColor(STATUS_COLORS["cancelled"]))

            if item_elapsed and ev.status == "ok" and ev.message:
                item_elapsed.setText(ev.message.split()[-1].rstrip(")"))

            if ev.status == "running":
                return  # 中间状态：仅更新行内提示，不推进总进度

            val = self.progress_bar.value() + 1
            self.progress_bar.setValue(val)
            self.label_progress.setText(f"正在处理 {ev.src.name} [{stage_prefix}] ({val}/{self.progress_bar.maximum()})")

    def _on_pipeline_finished(self, events: list[PipelineEvent]) -> None:
        self._set_running_state(False)
        self.progress_bar.setValue(self.progress_bar.maximum())

        failed = [e for e in events if e.status == "failed"]
        ok = [e for e in events if e.status == "ok"]

        if failed:
            msg = f"流水线处理完成：{len(ok)} 成功，{len(failed)} 失败。请查看表格标记。"
            self.label_progress.setText(f"⚠ 处理完成但存在异常 ({len(failed)} 个失败)")
            self.status_bar.showMessage(msg)
            QMessageBox.warning(self, "执行告警", msg)
        else:
            msg = f"🎉 交付完成！全部 {len(self.files_list)} 份文档已成功生成并就绪。"
            self.label_progress.setText(msg)
            self.status_bar.showMessage(msg)
            QMessageBox.information(self, "交付完成", msg)

    def _on_pipeline_error(self, err_text: str) -> None:
        self._set_running_state(False)
        self.label_progress.setText(f"✗ 发生致命异常: {err_text}")
        self.status_bar.showMessage("执行异常终止")
        QMessageBox.critical(self, "执行异常", f"流水线异常中断:\n{err_text}")

    def _apply_global_theme(self) -> None:
        self.setStyleSheet(APP_STYLESHEET)


def launch_gui() -> int:
    import sys
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()
