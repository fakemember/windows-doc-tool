"""
Windows Doc Tool (wdt) PySide6 主窗口与桌面界面交互逻辑。

界面结构：
- 顶部标题栏：品牌标识 + 转换引擎状态徽标
- 左侧边栏：工作模式导航（一体化流水线 / 批量转换 / 目录注入）
- 中部：文件列表 与 模式相关设置面板 左右分栏
- 底部操作栏：进度 + 开始/中止
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Sequence

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont
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

from wdt.core.batch_runner import _default_jobs
from wdt.core.engine_finder import Engine, EngineNotFoundError, find_engine
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

MODE_PAGES = {
    1: ("一体化流水线", "依次完成格式转换、目录注入和 PDF 导出"),
    2: ("批量格式转换", "支持 DOC / RTF / DOCX 互转，可导出灰度 PDF，多进程并行加速"),
    3: ("目录注入", "无需 LibreOffice，直接写入目录与页码，大文档秒级完成"),
}


class DropArea(QFrame):
    """顶部拖拽投放与快捷添加面板。"""

    def __init__(self, main_win: MainWindow) -> None:
        super().__init__(main_win)
        self.main_win = main_win
        self.setAcceptDrops(True)
        self.setObjectName("dropArea")
        self.setMinimumHeight(64)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(14)

        icon_label = QLabel("↓", self)
        icon_label.setFont(QFont("sans-serif", 22, QFont.Weight.Bold))
        icon_label.setStyleSheet("color: #2563EB;")
        icon_label.setFixedWidth(26)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        title_label = QLabel("拖拽待处理文件或目录到此处", self)
        title_label.setStyleSheet("font-size: 13px; font-weight: 600; color: #111827;")
        hint_label = QLabel("支持 DOC / RTF / DOCX，可混合拖入多个文件或整个目录", self)
        hint_label.setStyleSheet("font-size: 11px; color: #6B7280;")
        text_layout.addWidget(title_label)
        text_layout.addWidget(hint_label)
        layout.addLayout(text_layout)

        layout.addStretch()

        btn_files = QPushButton("添加文件…", self)
        btn_files.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_files.clicked.connect(self.main_win.on_add_files_dialog)
        layout.addWidget(btn_files)

        btn_dir = QPushButton("添加目录…", self)
        btn_dir.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_dir.clicked.connect(self.main_win.on_add_dir_dialog)
        layout.addWidget(btn_dir)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                f"""
                QFrame#dropArea {{
                    border: 1.5px dashed #2563EB;
                    border-radius: 10px;
                    background-color: #EFF6FF;
                }}
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
    """Windows Doc Tool 主界面。"""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Windows Doc Tool (wdt) - 办公与临床文档批处理工具")
        self.resize(1180, 780)
        self.setMinimumSize(960, 640)

        self.files_list: list[Path] = []
        self.file_row_map: dict[Path, int] = {}
        self.worker: PipelineWorker | None = None
        self.engine: Engine | None = None

        self._init_engine()
        self._init_ui()
        self._update_engine_pill()
        self._apply_global_theme()
        self._on_mode_changed(1)

    # ------------------------------------------------------------------
    # 初始化
    # ------------------------------------------------------------------

    def _init_engine(self) -> None:
        try:
            self.engine = find_engine()
        except EngineNotFoundError:
            self.engine = None

    def _init_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        outer_layout = QHBoxLayout(central_widget)
        outer_layout.setContentsMargins(14, 12, 14, 12)
        outer_layout.setSpacing(14)

        # ---------- 左侧边栏 ----------
        self.sidebar = QFrame(self)
        self.sidebar.setObjectName("sidebarCard")
        self.sidebar.setFixedWidth(196)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(12, 14, 12, 12)
        sidebar_layout.setSpacing(4)

        section_label = QLabel("工作模式", self.sidebar)
        section_label.setObjectName("sidebarSection")
        sidebar_layout.addWidget(section_label)
        sidebar_layout.addSpacing(6)

        # 模式导航按钮（真实状态载体为隐藏的单选按钮，便于测试与逻辑复用）
        self.nav_buttons: dict[int, QPushButton] = {}
        nav_specs = [
            (1, "一体化流水线", "run"),
            (2, "批量格式转换", "convert"),
            (3, "目录注入", "toc"),
        ]
        for btn_id, text, attr in nav_specs:
            btn = QPushButton(text, self.sidebar)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setCheckable(True)
            btn.setFlat(True)
            btn.clicked.connect(lambda _checked=False, i=btn_id: self._on_nav_clicked(i))
            sidebar_layout.addWidget(btn)
            self.nav_buttons[btn_id] = btn

        self.radio_mode_run = QRadioButton("run", self)
        self.radio_mode_conv = QRadioButton("convert", self)
        self.radio_mode_toc = QRadioButton("toc", self)
        self.radio_mode_run.setChecked(True)
        # 隐藏真实单选按钮：不参与可视布局
        for w in (self.radio_mode_run, self.radio_mode_conv, self.radio_mode_toc):
            w.setVisible(False)

        self.mode_btn_group = QButtonGroup(self)
        self.mode_btn_group.addButton(self.radio_mode_run, 1)
        self.mode_btn_group.addButton(self.radio_mode_conv, 2)
        self.mode_btn_group.addButton(self.radio_mode_toc, 3)
        self.mode_btn_group.idToggled.connect(self._on_mode_changed)

        sidebar_layout.addStretch()

        footer = QLabel("Windows Doc Tool\n文档批量转换与目录注入", self.sidebar)
        footer.setObjectName("sidebarFooter")
        footer.setAlignment(Qt.AlignmentFlag.AlignLeft)
        sidebar_layout.addWidget(footer)

        outer_layout.addWidget(self.sidebar)

        # ---------- 主内容列 ----------
        main_column = QWidget(self)
        main_layout = QVBoxLayout(main_column)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(10)

        # 1. 顶部标题栏
        header_bar = QFrame(self)
        header_bar.setObjectName("headerBar")
        header_layout = QHBoxLayout(header_bar)
        header_layout.setContentsMargins(16, 10, 16, 10)
        header_layout.setSpacing(10)

        brand_badge = QLabel("W", header_bar)
        brand_badge.setFixedSize(34, 34)
        brand_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_badge.setStyleSheet(
            "background-color: #2563EB; color: #FFFFFF; "
            "font-size: 17px; font-weight: 700; border-radius: 8px;"
        )
        header_layout.addWidget(brand_badge)

        brand_text = QVBoxLayout()
        brand_text.setSpacing(0)
        brand_title = QLabel("Windows Doc Tool", header_bar)
        brand_title.setStyleSheet("font-size: 15px; font-weight: 700; color: #111827;")
        brand_sub = QLabel("办公与临床文档批处理工具", header_bar)
        brand_sub.setStyleSheet("font-size: 11px; color: #6B7280;")
        brand_text.addWidget(brand_title)
        brand_text.addWidget(brand_sub)
        header_layout.addLayout(brand_text)

        header_layout.addStretch()

        # 引擎状态徽标
        self.engine_pill = QLabel(header_bar)
        self.engine_pill.setCursor(Qt.CursorShape.PointingHandCursor)
        header_layout.addWidget(self.engine_pill)
        self.btn_recheck_engine = QPushButton("重新检测", header_bar)
        self.btn_recheck_engine.clicked.connect(self._on_recheck_engine)
        header_layout.addWidget(self.btn_recheck_engine)

        main_layout.addWidget(header_bar)

        # 2. 页面标题（随模式切换）
        page_header = QHBoxLayout()
        page_header.setSpacing(8)
        self.page_title = QLabel("", main_column)
        self.page_title.setStyleSheet("font-size: 16px; font-weight: 700; color: #111827;")
        self.page_desc = QLabel("", main_column)
        self.page_desc.setStyleSheet("font-size: 12px; color: #6B7280;")
        page_header.addWidget(self.page_title)
        page_header.addWidget(self.page_desc, 1)
        main_layout.addLayout(page_header)

        # 3. 拖拽投放面板
        self.drop_area = DropArea(self)
        main_layout.addWidget(self.drop_area)

        # 4. 中部：左侧任务表格 + 右侧设置面板
        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.setChildrenCollapsible(False)

        # --- 左侧文件列表卡 ---
        files_card = QFrame(splitter)
        files_card.setObjectName("contentCard")
        files_layout = QVBoxLayout(files_card)
        files_layout.setContentsMargins(12, 10, 12, 10)
        files_layout.setSpacing(8)

        table_header_layout = QHBoxLayout()
        table_title = QLabel("待处理文档", files_card)
        table_title.setObjectName("cardTitle")
        table_header_layout.addWidget(table_title)
        self.label_file_count = QLabel("0 个文件", files_card)
        self.label_file_count.setObjectName("hintLabel")
        table_header_layout.addWidget(self.label_file_count)
        table_header_layout.addStretch()

        self.btn_remove = QPushButton("移除选中", files_card)
        self.btn_remove.clicked.connect(self.on_remove_selected)
        table_header_layout.addWidget(self.btn_remove)

        self.btn_clear = QPushButton("清空列表", files_card)
        self.btn_clear.clicked.connect(self.on_clear_files)
        table_header_layout.addWidget(self.btn_clear)

        files_layout.addLayout(table_header_layout)

        self.table = QTableWidget(files_card)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["文件名", "大小", "状态", "耗时", "源路径"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 200)
        self.table.setColumnWidth(2, 150)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        files_layout.addWidget(self.table)

        splitter.addWidget(files_card)

        # --- 右侧设置面板 ---
        right_scroll = QScrollArea(splitter)
        right_scroll.setWidgetResizable(True)
        right_scroll.setMinimumWidth(320)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        right_scroll.setObjectName("settingsScroll")

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 2, 0)
        right_layout.setSpacing(10)

        right_layout.addWidget(self._build_output_card(right_widget))
        right_layout.addWidget(self._build_toc_card(right_widget))
        right_layout.addWidget(self._build_perf_card(right_widget))
        right_layout.addStretch()

        right_scroll.setWidget(right_widget)
        splitter.addWidget(right_scroll)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([640, 380])
        main_layout.addWidget(splitter, 1)

        # 5. 底部操作栏
        action_bar = QFrame(self)
        action_bar.setObjectName("actionBar")
        bottom_layout = QHBoxLayout(action_bar)
        bottom_layout.setContentsMargins(14, 10, 14, 10)
        bottom_layout.setSpacing(12)

        progress_col = QVBoxLayout()
        progress_col.setSpacing(4)
        self.progress_bar = QProgressBar(action_bar)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(20)
        self.label_progress = QLabel("就绪。添加或拖入文档后即可开始处理", action_bar)
        self.label_progress.setObjectName("hintLabel")
        progress_col.addWidget(self.progress_bar)
        progress_col.addWidget(self.label_progress)
        bottom_layout.addLayout(progress_col, 1)

        self.btn_start = QPushButton("开始处理", action_bar)
        self.btn_start.setObjectName("btnStart")
        self.btn_start.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_start.setFixedHeight(40)
        self.btn_start.clicked.connect(self.on_start_pipeline)
        bottom_layout.addWidget(self.btn_start)

        self.btn_cancel = QPushButton("中止", action_bar)
        self.btn_cancel.setObjectName("btnCancel")
        self.btn_cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cancel.setFixedHeight(40)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.on_cancel_pipeline)
        bottom_layout.addWidget(self.btn_cancel)

        main_layout.addWidget(action_bar)

        outer_layout.addWidget(main_column, 1)

        # 状态栏
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪。请添加或拖拽文档开始处理。")

    # ------------------------------------------------------------------
    # 设置卡片
    # ------------------------------------------------------------------

    def _build_output_card(self, parent: QWidget) -> QFrame:
        """输出设置卡（convert / run 模式）。"""
        self.output_card = QFrame(parent)
        self.output_card.setObjectName("settingsCard")
        output_layout = QVBoxLayout(self.output_card)
        output_layout.setContentsMargins(14, 12, 14, 12)
        output_layout.setSpacing(10)

        title = QLabel("输出设置", self.output_card)
        title.setObjectName("cardTitle")
        output_layout.addWidget(title)

        fmt_grid = QGridLayout()
        fmt_grid.setHorizontalSpacing(14)
        fmt_grid.setVerticalSpacing(6)
        self.chk_docx = QCheckBox("导出 DOCX", self.output_card)
        self.chk_docx.setChecked(True)
        self.chk_pdf = QCheckBox("导出 PDF", self.output_card)
        self.chk_gray = QCheckBox("灰度 PDF", self.output_card)
        self.chk_gray.setToolTip("需要 LibreOffice 7.4 及以上版本")
        self.chk_force = QCheckBox("覆盖已有文件", self.output_card)
        fmt_grid.addWidget(self.chk_docx, 0, 0)
        fmt_grid.addWidget(self.chk_pdf, 0, 1)
        fmt_grid.addWidget(self.chk_gray, 1, 0)
        fmt_grid.addWidget(self.chk_force, 1, 1)
        output_layout.addLayout(fmt_grid)

        outdir_layout = QHBoxLayout()
        outdir_layout.setSpacing(8)
        label = QLabel("输出目录", self.output_card)
        label.setObjectName("formLabel")
        outdir_layout.addWidget(label)
        self.edit_outdir = QLineEdit(self.output_card)
        self.edit_outdir.setPlaceholderText("留空则保存到源文件所在目录")
        outdir_layout.addWidget(self.edit_outdir, 1)
        btn_browse_out = QPushButton("浏览…", self.output_card)
        btn_browse_out.clicked.connect(self._on_browse_outdir)
        outdir_layout.addWidget(btn_browse_out)
        output_layout.addLayout(outdir_layout)

        return self.output_card

    def _build_toc_card(self, parent: QWidget) -> QFrame:
        """目录 (TOC) 配置卡（toc / run 模式）。"""
        self.toc_card = QFrame(parent)
        self.toc_card.setObjectName("settingsCard")
        toc_layout = QVBoxLayout(self.toc_card)
        toc_layout.setContentsMargins(14, 12, 14, 12)
        toc_layout.setSpacing(10)

        title = QLabel("目录 (TOC) 配置", self.toc_card)
        title.setObjectName("cardTitle")
        toc_layout.addWidget(title)

        self.chk_enable_toc = QCheckBox("启用目录注入", self.toc_card)
        self.chk_enable_toc.setChecked(True)
        self.chk_enable_toc.toggled.connect(self._on_toc_toggle)
        toc_layout.addWidget(self.chk_enable_toc)

        self.toc_body_widget = QWidget(self.toc_card)
        toc_body_layout = QVBoxLayout(self.toc_body_widget)
        toc_body_layout.setContentsMargins(0, 0, 0, 0)
        toc_body_layout.setSpacing(10)

        # 4 样式缩略图单选卡片网格
        self.style_cards_widget = QWidget(self.toc_body_widget)
        cards_grid = QGridLayout(self.style_cards_widget)
        cards_grid.setContentsMargins(0, 0, 0, 0)
        cards_grid.setSpacing(8)

        self.card_classic = StyleCardWidget(TocStyle.CLASSIC_DOT, "经典样式", self.style_cards_widget)
        self.card_modern = StyleCardWidget(TocStyle.MODERN_ACCENT, "现代（推荐）", self.style_cards_widget)
        self.card_regulatory = StyleCardWidget(TocStyle.REGULATORY_LISTING, "临床/监管", self.style_cards_widget)
        self.card_minimal = StyleCardWidget(TocStyle.MINIMAL_CLEAN, "极简", self.style_cards_widget)

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
        toc_body_layout.addWidget(self.style_cards_widget)

        # TOC 详细参数
        params_grid = QGridLayout()
        params_grid.setHorizontalSpacing(8)
        params_grid.setVerticalSpacing(8)

        params_grid.addWidget(self._form_label("标题深度"), 0, 0)
        self.combo_depth = QComboBox(self.toc_body_widget)
        self.combo_depth.addItems(["1 级标题", "1-2 级标题", "1-3 级标题"])
        self.combo_depth.setCurrentIndex(1)
        params_grid.addWidget(self.combo_depth, 1, 0)

        params_grid.addWidget(self._form_label("目录标题"), 0, 1)
        self.edit_toc_title = QLineEdit("目  录", self.toc_body_widget)
        params_grid.addWidget(self.edit_toc_title, 1, 1)

        params_grid.addWidget(self._form_label("插入位置"), 2, 0)
        self.combo_position = QComboBox(self.toc_body_widget)
        self.combo_position.addItems([
            "文档开头",
            "首个标题前",
            "原位替换占位符",
        ])
        self.combo_position.setToolTip(
            "document-start / before-first-heading / replace-placeholder"
        )
        params_grid.addWidget(self.combo_position, 3, 0)

        params_grid.addWidget(self._form_label("页码模式"), 2, 1)
        self.combo_page_mode = QComboBox(self.toc_body_widget)
        self.combo_page_mode.addItems([
            "重新编页（推荐）",
            "页码偏移",
            "保留原页码",
            "Word 动态域",
        ])
        self.combo_page_mode.setToolTip(
            "static-restart / static-offset / static-in-place / dynamic\n"
            "注意：dynamic 与 PDF 导出互斥"
        )
        params_grid.addWidget(self.combo_page_mode, 3, 1)

        params_grid.addWidget(self._form_label("页码偏移量"), 4, 0)
        offset_row = QHBoxLayout()
        offset_row.setSpacing(8)
        self.spin_offset = QSpinBox(self.toc_body_widget)
        self.spin_offset.setRange(-999, 9999)
        self.spin_offset.setValue(0)
        offset_row.addWidget(self.spin_offset, 1)
        self.chk_backup = QCheckBox("保留原始文件备份（.bak）", self.toc_body_widget)
        self.chk_backup.setChecked(True)
        offset_row.addWidget(self.chk_backup)
        params_grid.addLayout(offset_row, 5, 0)

        toc_body_layout.addLayout(params_grid)
        toc_layout.addWidget(self.toc_body_widget)

        return self.toc_card

    def _build_perf_card(self, parent: QWidget) -> QFrame:
        """性能与并发卡（convert / run 模式）。"""
        self.perf_card = QFrame(parent)
        self.perf_card.setObjectName("settingsCard")
        perf_layout = QVBoxLayout(self.perf_card)
        perf_layout.setContentsMargins(14, 12, 14, 12)
        perf_layout.setSpacing(8)

        title = QLabel("性能与并发", self.perf_card)
        title.setObjectName("cardTitle")
        perf_layout.addWidget(title)

        jobs_row = QHBoxLayout()
        jobs_row.setSpacing(10)
        label = QLabel("并发进程数", self.perf_card)
        label.setObjectName("formLabel")
        jobs_row.addWidget(label)
        self.spin_jobs = QSpinBox(self.perf_card)
        self.spin_jobs.setRange(1, 8)
        self.spin_jobs.setValue(_default_jobs())
        jobs_row.addWidget(self.spin_jobs)
        tip_label = QLabel("根据 CPU 与内存自动推荐", self.perf_card)
        tip_label.setObjectName("hintLabel")
        jobs_row.addWidget(tip_label, 1)
        perf_layout.addLayout(jobs_row)

        return self.perf_card

    def _form_label(self, text: str) -> QLabel:
        label = QLabel(text, self.toc_card)
        label.setObjectName("formLabel")
        return label

    # ------------------------------------------------------------------
    # 引擎状态
    # ------------------------------------------------------------------

    def _update_engine_pill(self) -> None:
        """刷新标题栏引擎状态徽标。"""
        if self.engine is not None:
            ver = f"{self.engine.version[0]}.{self.engine.version[1]}" if self.engine.version else "未知版本"
            source_map = {"bundled": "内置运行时", "path": "PATH", "system": "系统安装"}
            source = source_map.get(self.engine.source, self.engine.source)
            self.engine_pill.setText(f" ● LibreOffice {ver} · {source} ")
            self.engine_pill.setToolTip(f"转换引擎路径：{self.engine.soffice}")
            self.engine_pill.setStyleSheet(
                "background-color: #ECFDF5; color: #047857; "
                "border: 1px solid #A7F3D0; border-radius: 10px; padding: 4px 10px; font-size: 12px;"
            )
            self.btn_recheck_engine.setVisible(False)
        else:
            self.engine_pill.setText(" ⚠ 未检测到转换引擎 ")
            self.engine_pill.setToolTip(
                "未检测到 LibreOffice，仅目录注入模式可用。"
                "安装 LibreOffice 或放置便携运行时后，点击右侧「重新检测」。"
            )
            self.engine_pill.setStyleSheet(
                "background-color: #FFFBEB; color: #B45309; "
                "border: 1px solid #FDE68A; border-radius: 10px; padding: 4px 10px; font-size: 12px;"
            )
            self.btn_recheck_engine.setVisible(True)

    def _on_recheck_engine(self) -> None:
        self._init_engine()
        self._update_engine_pill()
        if self.engine is not None:
            QMessageBox.information(
                self,
                "检测成功",
                f"已检测到转换引擎：{self.engine.soffice}\n版本：{self.engine.version}",
            )
        else:
            QMessageBox.warning(
                self,
                "未检测到引擎",
                "仍未检测到 LibreOffice，请检查安装路径。",
            )

    # ------------------------------------------------------------------
    # 模式切换
    # ------------------------------------------------------------------

    def _on_nav_clicked(self, btn_id: int) -> None:
        """点击侧边栏导航：同步到对应单选按钮（触发模式切换逻辑）。"""
        radio = {1: self.radio_mode_run, 2: self.radio_mode_conv, 3: self.radio_mode_toc}[btn_id]
        radio.setChecked(True)

    def _on_mode_changed(self, btn_id: int) -> None:
        """模式单选组切换：更新导航高亮、页面标题与设置卡可见性。"""
        if not self.radio_mode_run.isChecked() and not self.radio_mode_conv.isChecked():
            active_id = 3
        elif self.radio_mode_run.isChecked():
            active_id = 1
        else:
            active_id = 2

        for nav_id, nav_btn in self.nav_buttons.items():
            nav_btn.setChecked(nav_id == active_id)
            if nav_id == active_id:
                nav_btn.setStyleSheet(
                    "QPushButton { background-color: #EFF6FF; color: #1D4ED8; "
                    "font-weight: 600; text-align: left; padding: 9px 12px; "
                    "border: none; border-radius: 8px; font-size: 13px; }"
                )
            else:
                nav_btn.setStyleSheet(
                    "QPushButton { background-color: transparent; color: #4B5563; "
                    "text-align: left; padding: 9px 12px; "
                    "border: none; border-radius: 8px; font-size: 13px; }"
                    "QPushButton:hover { background-color: #F3F4F6; }"
                )

        title, desc = MODE_PAGES[active_id]
        self.page_title.setText(title)
        self.page_desc.setText(desc)

        # 设置卡可见性：仅展示当前模式相关配置
        self.output_card.setVisible(active_id in (1, 2))
        self.perf_card.setVisible(active_id in (1, 2))
        self.toc_card.setVisible(active_id in (1, 3))

        if active_id == 3:  # toc 模式：目录恒启用
            self.chk_enable_toc.setChecked(True)
            self.chk_enable_toc.setVisible(False)
        else:
            self.chk_enable_toc.setVisible(True)
        self._on_toc_toggle(self.chk_enable_toc.isChecked())

    def _on_toc_toggle(self, checked: bool) -> None:
        self.toc_body_widget.setVisible(checked)

    def _on_browse_outdir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.edit_outdir.setText(d)

    # ------------------------------------------------------------------
    # 文件列表管理
    # ------------------------------------------------------------------

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
        self.status_bar.showMessage("已清空列表")

    def on_remove_selected(self) -> None:
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            if r < len(self.files_list):
                self.files_list.pop(r)
                self.table.removeRow(r)
        self._rebuild_row_map()
        self._refresh_file_count()
        self.status_bar.showMessage(f"当前剩余 {len(self.files_list)} 个文档")

    def _rebuild_row_map(self) -> None:
        self.file_row_map = {f: idx for idx, f in enumerate(self.files_list)}

    def _refresh_file_count(self) -> None:
        n = len(self.files_list)
        self.label_file_count.setText(f"{n} 个文件")

    def add_paths(self, paths: Sequence[Path]) -> None:
        """核心文件收集方法，供 UI 交互与单元测试复用。"""
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
        self.status_bar.showMessage(f"已加载 {len(self.files_list)} 个待处理文档")

    # ------------------------------------------------------------------
    # 流水线选项与执行
    # ------------------------------------------------------------------

    def get_pipeline_options(self) -> PipelineOptions:
        """从当前 UI 状态提炼 PipelineOptions 对象。"""
        # 1. 模式
        if self.radio_mode_run.isChecked():
            mode = "run"
        elif self.radio_mode_conv.isChecked():
            mode = "convert"
        else:
            mode = "toc"

        outdir_text = self.edit_outdir.text().strip()
        out_dir = Path(outdir_text) if outdir_text else None

        # 2. TOC 样式与深度
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
            3: PageMode.DYNAMIC,
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

        # 在 run 模式下若未勾选启用 TOC，则将其视作仅转换 (convert)
        if mode == "run" and not self.chk_enable_toc.isChecked():
            mode = "convert"

        return PipelineOptions(
            mode=mode,  # type: ignore[arg-type]
            out_dir=out_dir,
            make_docx=self.chk_docx.isChecked(),
            make_pdf=self.chk_pdf.isChecked(),
            gray=self.chk_gray.isChecked(),
            toc=toc_config,
            jobs=self.spin_jobs.value(),
            force=self.chk_force.isChecked(),
        )

    def _set_running_state(self, running: bool) -> None:
        """切换运行/空闲状态的工具栏可用性。"""
        self.btn_start.setEnabled(not running)
        self.btn_cancel.setEnabled(running)
        self.btn_cancel.setText("中止")
        self.drop_area.setEnabled(not running)
        self.btn_clear.setEnabled(not running)
        self.btn_remove.setEnabled(not running)
        self.sidebar.setEnabled(not running)
        self.output_card.setEnabled(not running)
        self.toc_card.setEnabled(not running)
        self.perf_card.setEnabled(not running)

    def on_start_pipeline(self) -> None:
        """开始执行流水线，包含启动前各项语义校验。"""
        if not self.files_list:
            QMessageBox.information(self, "提示", "请先添加或拖入待处理的文档。")
            return

        try:
            options = self.get_pipeline_options()
        except Exception as e:
            QMessageBox.warning(self, "参数错误", str(e))
            return

        # 校验 1: convert 模式下 DOCX 和 PDF 至少勾选一个
        if options.mode == "convert" and not options.make_docx and not options.make_pdf:
            QMessageBox.warning(
                self, "参数配置", "批量转换模式下，请至少勾选「导出 DOCX」或「导出 PDF」。"
            )
            return

        # 校验 2: dynamic + PDF 互斥拒绝
        if (
            options.mode == "run"
            and options.make_pdf
            and options.toc is not None
            and options.toc.page_mode == PageMode.DYNAMIC
        ):
            QMessageBox.critical(
                self,
                "模式冲突",
                "「Word 动态域」页码模式不能与 PDF 导出同时使用。\n"
                "转换为 PDF 时不会更新页码域，请改用「重新编页」或「页码偏移」。",
            )
            return

        # 校验 3: 如果需要调用 LibreOffice 进行格式转换，检查引擎
        if options.mode in ("run", "convert") and self.engine is None:
            QMessageBox.critical(
                self,
                "缺少转换引擎",
                "当前模式需要 LibreOffice 转换引擎，但系统中未检测到。\n"
                "请安装 LibreOffice 后点击右上角「重新检测」，或改用「目录注入」模式。",
            )
            return

        # 准备执行
        self._set_running_state(True)

        # 重置所有表格项状态
        for row in range(self.table.rowCount()):
            item_st = self.table.item(row, 2)
            if item_st:
                item_st.setText("等待处理")
                item_st.setForeground(QColor(STATUS_COLORS["queued"]))
                item_st.setToolTip("")
            item_el = self.table.item(row, 3)
            if item_el:
                item_el.setText("-")

        # 进度条总步数 = 文件数 × 单文件事件数（convert 每目标 1 个，run 为 转换+目录(+PDF) 最多 3 个）
        events_per_file = 1
        if options.mode == "convert":
            events_per_file = (1 if options.make_docx else 0) + (
                1 if options.make_pdf else 0
            )
        elif options.mode == "run":
            events_per_file = 3 if options.make_pdf else 2
        total_steps = max(1, len(self.files_list) * events_per_file)
        self.progress_bar.setRange(0, total_steps)
        self.progress_bar.setValue(0)
        self.label_progress.setText(f"正在处理：0 / {total_steps}")
        self.status_bar.showMessage(f"正在处理，共 {len(self.files_list)} 个文档")

        self.worker = PipelineWorker(
            files=self.files_list,
            options=options,
            engine=self.engine,
            parent=self,
        )
        self.worker.event_signal.connect(self._on_pipeline_event)
        self.worker.finished_signal.connect(self._on_pipeline_finished)
        self.worker.error_signal.connect(self._on_pipeline_error)
        self.worker.start()

    def on_cancel_pipeline(self) -> None:
        """用户点击中止按钮。"""
        if self.worker and self.worker.isRunning():
            self.btn_cancel.setEnabled(False)
            self.btn_cancel.setText("正在中止…")
            self.status_bar.showMessage("正在中止，请稍候…")
            self.worker.cancel()

    def _on_pipeline_event(self, ev: PipelineEvent) -> None:
        """接收 worker 线程推送的每个文件处理事件。"""
        row = self.file_row_map.get(ev.src)
        if row is None:
            # 兼容绝对路径与相对路径比对
            for src, r in self.file_row_map.items():
                if src.resolve() == ev.src.resolve():
                    row = r
                    break

        if row is not None:
            stage_prefix = {
                Stage.CONVERT: "转换",
                Stage.TOC: "加目录",
                Stage.PDF: "转PDF",
            }.get(ev.stage, ev.stage.value)

            item_status = self.table.item(row, 2)
            item_elapsed = self.table.item(row, 3)

            if item_status:
                if ev.status == "ok":
                    item_status.setText(f"✓ {stage_prefix}成功")
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

            if item_elapsed and ev.message and ("s" in ev.message or "条目" in ev.message):
                item_elapsed.setText(ev.message.split(",")[0].strip())

        # 进度条步进
        value = min(self.progress_bar.value() + 1, self.progress_bar.maximum())
        self.progress_bar.setValue(value)
        self.label_progress.setText(f"正在处理：{value} / {self.progress_bar.maximum()}")

    def _on_pipeline_finished(self, results: list[PipelineEvent]) -> None:
        """流水线执行完毕回调。"""
        self._set_running_state(False)

        ok_count = sum(1 for e in results if e.status == "ok")
        fail_count = sum(1 for e in results if e.status == "failed")
        cancel_count = sum(1 for e in results if e.status == "cancelled")

        msg = f"处理完成，成功 {ok_count} 个，失败 {fail_count} 个"
        if cancel_count:
            msg += f"，取消 {cancel_count} 个"

        self.status_bar.showMessage(msg)
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.label_progress.setText(msg)

    def _on_pipeline_error(self, err_msg: str) -> None:
        """后台严重异常回调。"""
        self._set_running_state(False)

        QMessageBox.critical(self, "处理出错", f"处理过程中发生错误：\n{err_msg}")
        self.status_bar.showMessage(f"处理中断：{err_msg}")
        self.label_progress.setText(f"处理中断：{err_msg}")

    def _apply_global_theme(self) -> None:
        """加载全局浅色主题样式表。"""
        self.setStyleSheet(APP_STYLESHEET)


def launch_gui() -> int:
    """启动 GUI 主函数。"""
    # 启用高分屏支持
    if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    window = MainWindow()
    window.show()
    return app.exec()
