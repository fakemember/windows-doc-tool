"""
Windows Doc Tool (wdt) PySide6 主窗口与桌面界面交互逻辑。
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Sequence

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
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
    resolve_destination,
)
from wdt.toc.docx_stream_patcher import Position, TocConfig
from wdt.toc.page_calculator import PageMode
from wdt.toc.styles_templates import TocStyle
from wdt.ui.style_preview import StyleCardWidget
from wdt.ui.worker import PipelineWorker

SUPPORTED_EXTS = (".doc", ".rtf", ".docx")


class DropArea(QFrame):
    """顶部拖拽投放与快捷添加面板。"""

    def __init__(self, main_win: MainWindow) -> None:
        super().__init__(main_win)
        self.main_win = main_win
        self.setAcceptDrops(True)
        self.setObjectName("dropArea")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(16)

        icon_label = QLabel("📥", self)
        icon_label.setFont(QFont("sans-serif", 24))
        layout.addWidget(icon_label)

        text_layout = QVBoxLayout()
        title_label = QLabel("拖拽待处理文件或目录到此处", self)
        title_label.setStyleSheet("font-size: 13px; font-weight: bold; color: #1E293B;")
        hint_label = QLabel("支持 DOC / RTF / DOCX 格式，可混合拖入多个文件或整个目录", self)
        hint_label.setStyleSheet("font-size: 11px; color: #64748B;")
        text_layout.addWidget(title_label)
        text_layout.addWidget(hint_label)
        layout.addLayout(text_layout)

        layout.addStretch()

        btn_files = QPushButton("📄 添加文件…", self)
        btn_files.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_files.clicked.connect(self.main_win.on_add_files_dialog)
        layout.addWidget(btn_files)

        btn_dir = QPushButton("📁 添加目录…", self)
        btn_dir.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_dir.clicked.connect(self.main_win.on_add_dir_dialog)
        layout.addWidget(btn_dir)

        self.setStyleSheet(
            """
            QFrame#dropArea {
                border: 2px dashed #CBD5E1;
                border-radius: 8px;
                background-color: #F8FAFC;
            }
            QFrame#dropArea:hover {
                border-color: #0284C7;
                background-color: #F0F9FF;
            }
            """
        )

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                """
                QFrame#dropArea {
                    border: 2px dashed #0284C7;
                    border-radius: 8px;
                    background-color: #E0F2FE;
                }
                """
            )

    def dragLeaveEvent(self, event) -> None:
        self.setStyleSheet(
            """
            QFrame#dropArea {
                border: 2px dashed #CBD5E1;
                border-radius: 8px;
                background-color: #F8FAFC;
            }
            """
        )

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
        self.resize(1100, 750)
        self.setMinimumSize(920, 620)

        self.files_list: list[Path] = []
        self.file_row_map: dict[Path, int] = {}
        self.worker: PipelineWorker | None = None
        self.engine: Engine | None = None

        self._init_engine()
        self._init_ui()
        self._update_engine_banner()
        self._apply_global_theme()

    def _init_engine(self) -> None:
        try:
            self.engine = find_engine()
        except EngineNotFoundError:
            self.engine = None

    def _init_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(14, 12, 14, 12)
        main_layout.setSpacing(10)

        # 1. 顶部引擎缺失警告横幅
        self.engine_banner = QFrame(self)
        self.engine_banner.setObjectName("engineBanner")
        banner_layout = QHBoxLayout(self.engine_banner)
        banner_layout.setContentsMargins(12, 8, 12, 8)
        self.banner_label = QLabel(self.engine_banner)
        self.banner_label.setWordWrap(True)
        self.banner_label.setText(
            "⚠️ <b>未检测到 LibreOffice 转换引擎！</b> 当前仅支持纯 Python「仅加目录 (TOC)」功能。"
            "若需执行 DOC/RTF/PDF 转换，请安装 LibreOffice 或将 portable 运行时放置于 ./runtime/libreoffice"
        )
        banner_layout.addWidget(self.banner_label)
        btn_recheck = QPushButton("重新检测", self.engine_banner)
        btn_recheck.clicked.connect(self._on_recheck_engine)
        banner_layout.addWidget(btn_recheck)
        main_layout.addWidget(self.engine_banner)

        # 2. 顶部拖拽投放面板
        self.drop_area = DropArea(self)
        main_layout.addWidget(self.drop_area)

        # 3. 中部：左侧任务表格 + 右侧参数面板 (QSplitter)
        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.setChildrenCollapsible(False)

        # 左侧表格容器
        left_container = QWidget(splitter)
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        table_header_layout = QHBoxLayout()
        table_title = QLabel("待处理文档列表", left_container)
        table_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        table_header_layout.addWidget(table_title)
        table_header_layout.addStretch()

        self.btn_clear = QPushButton("清空列表", left_container)
        self.btn_clear.clicked.connect(self.on_clear_files)
        table_header_layout.addWidget(self.btn_clear)

        self.btn_remove = QPushButton("移除选中", left_container)
        self.btn_remove.clicked.connect(self.on_remove_selected)
        table_header_layout.addWidget(self.btn_remove)
        left_layout.addLayout(table_header_layout)

        self.table = QTableWidget(left_container)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["文件名", "大小", "状态", "耗时", "源路径"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(2, 140)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        left_layout.addWidget(self.table)

        # 右侧配置面板 (外层放 QScrollArea 防溢出)
        right_scroll = QScrollArea(splitter)
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        right_scroll.setObjectName("rightScroll")

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 4, 10, 4)
        right_layout.setSpacing(12)

        # 模式选择单选组
        mode_group = QGroupBox("处理模式", right_widget)
        mode_layout = QVBoxLayout(mode_group)
        self.radio_mode_run = QRadioButton("转换 + 注入目录 (run 完整流水线)", mode_group)
        self.radio_mode_conv = QRadioButton("仅批量格式转换 (convert)", mode_group)
        self.radio_mode_toc = QRadioButton("仅 DOCX 目录注入 (toc 纯Python快速)", mode_group)
        self.radio_mode_run.setChecked(True)

        self.mode_btn_group = QButtonGroup(self)
        self.mode_btn_group.addButton(self.radio_mode_run, 1)
        self.mode_btn_group.addButton(self.radio_mode_conv, 2)
        self.mode_btn_group.addButton(self.radio_mode_toc, 3)
        self.mode_btn_group.idToggled.connect(self._on_mode_changed)

        mode_layout.addWidget(self.radio_mode_run)
        mode_layout.addWidget(self.radio_mode_conv)
        mode_layout.addWidget(self.radio_mode_toc)
        right_layout.addWidget(mode_group)

        # 输出格式与参数卡片
        self.output_group = QGroupBox("输出设置", right_widget)
        output_layout = QVBoxLayout(self.output_group)

        fmt_layout = QHBoxLayout()
        self.chk_docx = QCheckBox("导出 DOCX", self.output_group)
        self.chk_docx.setChecked(True)
        self.chk_pdf = QCheckBox("导出 PDF", self.output_group)
        self.chk_pdf.setChecked(False)
        self.chk_gray = QCheckBox("灰度 PDF (需 LO>=7.4)", self.output_group)
        self.chk_force = QCheckBox("覆盖已有文件 (-f)", self.output_group)
        fmt_layout.addWidget(self.chk_docx)
        fmt_layout.addWidget(self.chk_pdf)
        fmt_layout.addWidget(self.chk_gray)
        output_layout.addLayout(fmt_layout)
        output_layout.addWidget(self.chk_force)

        outdir_layout = QHBoxLayout()
        outdir_layout.addWidget(QLabel("输出目录:", self.output_group))
        self.edit_outdir = QLineEdit(self.output_group)
        self.edit_outdir.setPlaceholderText("默认与源文件同目录")
        outdir_layout.addWidget(self.edit_outdir)
        btn_browse_out = QPushButton("选择…", self.output_group)
        btn_browse_out.clicked.connect(self._on_browse_outdir)
        outdir_layout.addWidget(btn_browse_out)
        output_layout.addLayout(outdir_layout)
        right_layout.addWidget(self.output_group)

        # TOC 配置卡片
        self.toc_group = QGroupBox("目录 (TOC) 配置", right_widget)
        toc_layout = QVBoxLayout(self.toc_group)

        self.chk_enable_toc = QCheckBox("启用 TOC 目录注入", self.toc_group)
        self.chk_enable_toc.setChecked(True)
        self.chk_enable_toc.toggled.connect(self._on_toc_toggle)
        toc_layout.addWidget(self.chk_enable_toc)

        # 4 样式缩略图单选卡片网格
        self.style_cards_widget = QWidget(self.toc_group)
        cards_grid = QGridLayout(self.style_cards_widget)
        cards_grid.setContentsMargins(0, 4, 0, 4)
        cards_grid.setSpacing(8)

        self.card_classic = StyleCardWidget(TocStyle.CLASSIC_DOT, "经典经典点导线", self.style_cards_widget)
        self.card_modern = StyleCardWidget(TocStyle.MODERN_ACCENT, "现代科技蓝 (推荐)", self.style_cards_widget)
        self.card_regulatory = StyleCardWidget(TocStyle.REGULATORY_LISTING, "临床/监管密点", self.style_cards_widget)
        self.card_minimal = StyleCardWidget(TocStyle.MINIMAL_CLEAN, "极简无导线", self.style_cards_widget)

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
        toc_layout.addWidget(self.style_cards_widget)

        # TOC 详细参数行
        toc_params_grid = QGridLayout()
        toc_params_grid.addWidget(QLabel("标题深度:", self.toc_group), 0, 0)
        self.combo_depth = QComboBox(self.toc_group)
        self.combo_depth.addItems(["1 级标题", "1-2 级标题", "1-3 级标题"])
        self.combo_depth.setCurrentIndex(1)  # 默认 1-2 级
        toc_params_grid.addWidget(self.combo_depth, 0, 1)

        toc_params_grid.addWidget(QLabel("目录标题:", self.toc_group), 0, 2)
        self.edit_toc_title = QLineEdit("目  录", self.toc_group)
        toc_params_grid.addWidget(self.edit_toc_title, 0, 3)

        toc_params_grid.addWidget(QLabel("插入位置:", self.toc_group), 1, 0)
        self.combo_position = QComboBox(self.toc_group)
        self.combo_position.addItems([
            "文档开头 (document-start)",
            "首个标题前 (before-first-heading)",
            "原位替换占位符 (replace-placeholder)",
        ])
        toc_params_grid.addWidget(self.combo_position, 1, 1)

        toc_params_grid.addWidget(QLabel("页码模式:", self.toc_group), 1, 2)
        self.combo_page_mode = QComboBox(self.toc_group)
        self.combo_page_mode.addItems([
            "分节重启 (static-restart)",
            "偏移补偿 (static-offset)",
            "原位页码 (static-in-place)",
            "动态域代码 (dynamic)",
        ])
        toc_params_grid.addWidget(self.combo_page_mode, 1, 3)

        toc_params_grid.addWidget(QLabel("页码偏移量:", self.toc_group), 2, 0)
        self.spin_offset = QSpinBox(self.toc_group)
        self.spin_offset.setRange(-999, 9999)
        self.spin_offset.setValue(0)
        toc_params_grid.addWidget(self.spin_offset, 2, 1)

        self.chk_backup = QCheckBox("保留 .bak 原始备份", self.toc_group)
        self.chk_backup.setChecked(True)
        toc_params_grid.addWidget(self.chk_backup, 2, 2, 1, 2)

        toc_layout.addLayout(toc_params_grid)
        right_layout.addWidget(self.toc_group)

        # 并发控制组
        concurrency_group = QGroupBox("性能与并发", right_widget)
        conc_layout = QHBoxLayout(concurrency_group)
        conc_layout.addWidget(QLabel("并发进程数:", concurrency_group))
        self.spin_jobs = QSpinBox(concurrency_group)
        self.spin_jobs.setRange(1, 8)
        self.spin_jobs.setValue(_default_jobs())
        conc_layout.addWidget(self.spin_jobs)

        tip_label = QLabel("（系统推荐默认值，已结合 CPU 核心与物理内存钳制）", concurrency_group)
        tip_label.setStyleSheet("color: #64748B; font-size: 11px;")
        conc_layout.addWidget(tip_label)
        conc_layout.addStretch()
        right_layout.addWidget(concurrency_group)

        right_layout.addStretch()
        right_scroll.setWidget(right_widget)

        splitter.addWidget(left_container)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        main_layout.addWidget(splitter)

        # 4. 底部控制栏
        bottom_bar = QFrame(self)
        bottom_layout = QHBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(4, 6, 4, 6)
        bottom_layout.setSpacing(12)

        self.progress_bar = QProgressBar(bottom_bar)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        bottom_layout.addWidget(self.progress_bar, 1)

        self.btn_start = QPushButton("🚀 开始处理", bottom_bar)
        self.btn_start.setObjectName("btnStart")
        self.btn_start.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_start.setFixedHeight(36)
        self.btn_start.clicked.connect(self.on_start_pipeline)
        bottom_layout.addWidget(self.btn_start)

        self.btn_cancel = QPushButton("⏹ 中止", bottom_bar)
        self.btn_cancel.setObjectName("btnCancel")
        self.btn_cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cancel.setFixedHeight(36)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.on_cancel_pipeline)
        bottom_layout.addWidget(self.btn_cancel)

        main_layout.addWidget(bottom_bar)

        # 状态栏
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪。请添加或拖拽文档开始处理。")

    def _update_engine_banner(self) -> None:
        if self.engine is not None:
            self.engine_banner.setVisible(False)
        else:
            self.engine_banner.setVisible(True)

    def _on_recheck_engine(self) -> None:
        self._init_engine()
        self._update_engine_banner()
        if self.engine is not None:
            QMessageBox.information(
                self,
                "检测成功",
                f"已检测到转换引擎: {self.engine.soffice}\n版本: {self.engine.version}",
            )
        else:
            QMessageBox.warning(
                self,
                "未检测到引擎",
                "仍未找到可用的 LibreOffice / soffice 程序，请检查安装路径。",
            )

    def _on_mode_changed(self, btn_id: int) -> None:
        # 1 = run, 2 = convert, 3 = toc
        if btn_id == 1:  # run
            self.output_group.setEnabled(True)
            self.toc_group.setEnabled(True)
            self.chk_enable_toc.setEnabled(True)
        elif btn_id == 2:  # convert
            self.output_group.setEnabled(True)
            self.toc_group.setEnabled(False)
        elif btn_id == 3:  # toc
            self.output_group.setEnabled(False)
            self.toc_group.setEnabled(True)
            self.chk_enable_toc.setChecked(True)
            self.chk_enable_toc.setEnabled(False)

    def _on_toc_toggle(self, checked: bool) -> None:
        self.style_cards_widget.setEnabled(checked)
        self.combo_depth.setEnabled(checked)
        self.edit_toc_title.setEnabled(checked)
        self.combo_position.setEnabled(checked)
        self.combo_page_mode.setEnabled(checked)
        self.spin_offset.setEnabled(checked)
        self.chk_backup.setEnabled(checked)

    def _on_browse_outdir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
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
        self.status_bar.showMessage("已清空列表")

    def on_remove_selected(self) -> None:
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            if r < len(self.files_list):
                p = self.files_list.pop(r)
                self.table.removeRow(r)
        self._rebuild_row_map()
        self.status_bar.showMessage(f"当前剩余 {len(self.files_list)} 个文档")

    def _rebuild_row_map(self) -> None:
        self.file_row_map = {f: idx for idx, f in enumerate(self.files_list)}

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
            item_status = QTableWidgetItem("排队中")
            item_elapsed = QTableWidgetItem("-")
            item_path = QTableWidgetItem(str(f.parent))

            item_status.setForeground(QColor("#64748B"))

            self.table.setItem(row, 0, item_name)
            self.table.setItem(row, 1, item_size)
            self.table.setItem(row, 2, item_status)
            self.table.setItem(row, 3, item_elapsed)
            self.table.setItem(row, 4, item_path)

        self.status_bar.showMessage(f"已加载 {len(self.files_list)} 个待处理文档")

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

    def on_start_pipeline(self) -> None:
        """开始执行流水线，包含启动前各项语义校验。"""
        if not self.files_list:
            QMessageBox.information(self, "提示", "请先添加或拖入待处理的文档！")
            return

        try:
            options = self.get_pipeline_options()
        except Exception as e:
            QMessageBox.warning(self, "参数错误", str(e))
            return

        # 校验 1: convert 模式下 DOCX 和 PDF 至少勾选一个
        if options.mode == "convert" and not options.make_docx and not options.make_pdf:
            QMessageBox.warning(
                self, "参数配置", "仅转换模式下请至少勾选「导出 DOCX」或「导出 PDF」之一！"
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
                "模式互斥错误",
                "导出 PDF 不能与「动态域代码 (dynamic)」目录同时使用！\n"
                "无头 LibreOffice 在转换为 PDF 时不会刷新 Word 动态字段，请使用「分节重启」或「偏移补偿」静态页码模式。",
            )
            return

        # 校验 3: 如果需要调用 LibreOffice 进行格式转换，检查引擎
        if options.mode in ("run", "convert") and self.engine is None:
            QMessageBox.critical(
                self,
                "缺少转换引擎",
                "当前模式需要使用 LibreOffice 转换引擎，但系统中未找到！\n"
                "请安装 LibreOffice 后点击顶部「重新检测」，或改用「仅 DOCX 目录注入」模式。",
            )
            return

        # 准备执行
        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_cancel.setText("⏹ 中止")
        self.drop_area.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.btn_remove.setEnabled(False)

        # 重置所有表格项状态
        for row in range(self.table.rowCount()):
            item_st = self.table.item(row, 2)
            if item_st:
                item_st.setText("排队中")
                item_st.setForeground(QColor("#64748B"))
            item_el = self.table.item(row, 3)
            if item_el:
                item_el.setText("-")

        self.progress_bar.setRange(0, len(self.files_list))
        self.progress_bar.setValue(0)
        self.status_bar.showMessage(f"流水线运行中... 待处理: {len(self.files_list)} 个")

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
            self.status_bar.showMessage("正在中止未完成任务，请稍候...")
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
                    item_status.setForeground(QColor("#16A34A"))  # 绿色
                elif ev.status == "failed":
                    item_status.setText(f"✗ {stage_prefix}失败")
                    item_status.setForeground(QColor("#DC2626"))  # 红色
                    item_status.setToolTip(ev.message)
                elif ev.status == "skipped":
                    item_status.setText(f"⏩ 已跳过")
                    item_status.setForeground(QColor("#D97706"))  # 黄橙色
                    item_status.setToolTip(ev.message)
                elif ev.status == "cancelled":
                    item_status.setText("⏹ 已取消")
                    item_status.setForeground(QColor("#64748B"))  # 灰色

            if item_elapsed and ev.message and ("s" in ev.message or "条目" in ev.message):
                item_elapsed.setText(ev.message.split(",")[0].strip())

        # 进度条步进
        self.progress_bar.setValue(self.progress_bar.value() + 1)

    def _on_pipeline_finished(self, results: list[PipelineEvent]) -> None:
        """流水线执行完毕回调。"""
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.setText("⏹ 中止")
        self.drop_area.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.btn_remove.setEnabled(True)

        ok_count = sum(1 for e in results if e.status == "ok")
        fail_count = sum(1 for e in results if e.status == "failed")
        cancel_count = sum(1 for e in results if e.status == "cancelled")

        msg = f"处理完成！成功: {ok_count}，失败: {fail_count}"
        if cancel_count:
            msg += f"，取消: {cancel_count}"

        self.status_bar.showMessage(msg)
        self.progress_bar.setValue(self.progress_bar.maximum())

    def _on_pipeline_error(self, err_msg: str) -> None:
        """后台严重异常回调。"""
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.setText("⏹ 中止")
        self.drop_area.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.btn_remove.setEnabled(True)

        QMessageBox.critical(self, "流水线异常", f"执行遇到未捕获异常:\n{err_msg}")
        self.status_bar.showMessage(f"异常终止: {err_msg}")

    def _apply_global_theme(self) -> None:
        """设置整洁现代的扁平化样式表。"""
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #FFFFFF;
            }
            QWidget {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                font-size: 12px;
                color: #1E293B;
            }
            QFrame#engineBanner {
                background-color: #FEF3C7;
                border: 1px solid #FCD34D;
                border-radius: 6px;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 12px;
                background-color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                color: #0F172A;
            }
            QTableWidget {
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                background-color: #FFFFFF;
                gridline-color: #F1F5F9;
            }
            QTableWidget::item {
                padding: 4px;
            }
            QTableWidget::item:selected {
                background-color: #E0F2FE;
                color: #0369A1;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                font-weight: bold;
                border: none;
                border-bottom: 1px solid #E2E8F0;
                padding: 6px 4px;
            }
            QPushButton {
                background-color: #F1F5F9;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 5px 12px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #E2E8F0;
            }
            QPushButton:pressed {
                background-color: #CBD5E1;
            }
            QPushButton#btnStart {
                background-color: #0284C7;
                color: #FFFFFF;
                font-size: 13px;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 6px 20px;
            }
            QPushButton#btnStart:hover {
                background-color: #0369A1;
            }
            QPushButton#btnStart:disabled {
                background-color: #94A3B8;
            }
            QPushButton#btnCancel {
                background-color: #EF4444;
                color: #FFFFFF;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 6px 14px;
            }
            QPushButton#btnCancel:hover {
                background-color: #DC2626;
            }
            QPushButton#btnCancel:disabled {
                background-color: #CBD5E1;
                color: #94A3B8;
            }
            QProgressBar {
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                text-align: center;
                background-color: #F8FAFC;
                height: 18px;
            }
            QProgressBar::chunk {
                background-color: #0284C7;
                border-radius: 3px;
            }
            QLineEdit, QComboBox, QSpinBox {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 4px 6px;
                background-color: #FFFFFF;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
                border-color: #0284C7;
            }
            QScrollArea#rightScroll {
                border: none;
                background-color: transparent;
            }
            """
        )


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
