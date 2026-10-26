"""
PySide6 主窗口界面交互与参数映射单元测试。
"""

from __future__ import annotations

import os

# 必须在 import PySide6 之前设置 offscreen 平台插件
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pathlib import Path
from unittest.mock import MagicMock

import pytest
from PySide6.QtWidgets import QApplication, QMessageBox

from tests.fixtures_factory import build_docx
from wdt.pipeline import PipelineOptions, Stage
from wdt.toc.docx_stream_patcher import Position
from wdt.toc.page_calculator import PageMode
from wdt.toc.styles_templates import TocStyle
from wdt.ui.main_window import MainWindow


@pytest.fixture(scope="session")
def qapp():
    """管理全局 QApplication 单例。"""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def main_window(qapp):
    """创建新的 MainWindow 实例，并在测试结束后清理。"""
    win = MainWindow()
    yield win
    win.close()


def test_main_window_init(main_window):
    """测试主窗口与核心控件初始化。"""
    assert main_window.windowTitle().startswith("Windows Doc Tool")
    assert main_window.table.columnCount() == 5
    assert main_window.btn_start.isEnabled() is True
    assert main_window.btn_cancel.isEnabled() is False
    assert main_window.progress_bar.value() == 0


def test_add_paths_and_filtering(tmp_path: Path, main_window):
    """测试 add_paths 文件收集、目录展开与临时文件过滤。"""
    # 构造有效文件
    doc1 = tmp_path / "valid1.doc"
    rtf1 = tmp_path / "valid2.rtf"
    docx1 = tmp_path / "valid3.docx"
    doc1.write_bytes(b"dummy doc")
    rtf1.write_bytes(b"dummy rtf")
    docx1.write_bytes(b"dummy docx")

    # 构造应被过滤的文件
    tmp_file = tmp_path / "~$temp.docx"
    txt_file = tmp_path / "readme.txt"
    tmp_file.write_bytes(b"temp")
    txt_file.write_bytes(b"text")

    # 构造子目录及内部文件
    sub_dir = tmp_path / "nested"
    sub_dir.mkdir()
    nested_docx = sub_dir / "nested.docx"
    nested_docx.write_bytes(b"nested")

    # 执行添加
    main_window.add_paths([doc1, rtf1, docx1, tmp_file, txt_file, sub_dir])

    # 应只收录 doc1, rtf1, docx1, nested_docx (共 4 个)
    assert len(main_window.files_list) == 4
    assert main_window.table.rowCount() == 4

    names = [main_window.files_list[i].name for i in range(4)]
    assert "valid1.doc" in names
    assert "valid2.rtf" in names
    assert "valid3.docx" in names
    assert "nested.docx" in names
    assert "~$temp.docx" not in names
    assert "readme.txt" not in names


def test_ui_options_mapping_run_mode(main_window):
    """测试 UI 控件状态精准映射到 PipelineOptions (run 模式)。"""
    main_window.radio_mode_run.setChecked(True)
    main_window.chk_docx.setChecked(True)
    main_window.chk_pdf.setChecked(True)
    main_window.chk_gray.setChecked(True)
    main_window.chk_force.setChecked(True)

    # 选择现代科技蓝样式
    main_window.card_modern.setChecked(True)
    main_window.combo_depth.setCurrentIndex(2)  # 1-3 级
    main_window.edit_toc_title.setText("临床试验报告目录")
    main_window.combo_position.setCurrentIndex(1)  # before-first-heading
    main_window.combo_page_mode.setCurrentIndex(1)  # static-offset
    main_window.spin_offset.setValue(5)
    main_window.chk_backup.setChecked(False)
    main_window.spin_jobs.setValue(4)

    opts: PipelineOptions = main_window.get_pipeline_options()

    assert opts.mode == "run"
    assert opts.make_docx is True
    assert opts.make_pdf is True
    assert opts.gray is True
    assert opts.force is True
    assert opts.jobs == 4

    assert opts.toc is not None
    assert opts.toc.style == TocStyle.MODERN_ACCENT
    assert opts.toc.depth == 3
    assert opts.toc.title == "临床试验报告目录"
    assert opts.toc.position == Position.BEFORE_FIRST_HEADING
    assert opts.toc.page_mode == PageMode.STATIC_OFFSET
    assert opts.toc.page_offset == 5
    assert opts.toc.backup is False


def test_ui_options_mapping_convert_and_toc_modes(main_window):
    """测试 convert 与 toc 模式下的参数提取。"""
    # 1. 切换至仅转换模式
    main_window.radio_mode_conv.setChecked(True)
    main_window.chk_docx.setChecked(False)
    main_window.chk_pdf.setChecked(True)
    opts_conv = main_window.get_pipeline_options()
    assert opts_conv.mode == "convert"
    assert opts_conv.make_docx is False
    assert opts_conv.make_pdf is True

    # 2. 切换至仅加目录模式
    main_window.radio_mode_toc.setChecked(True)
    main_window.card_regulatory.setChecked(True)
    opts_toc = main_window.get_pipeline_options()
    assert opts_toc.mode == "toc"
    assert opts_toc.toc.style == TocStyle.REGULATORY_LISTING


def test_cancel_button_state(main_window):
    """测试中止操作时按钮状态切换。"""
    mock_worker = MagicMock()
    mock_worker.isRunning.return_value = True
    main_window.worker = mock_worker

    main_window.btn_cancel.setEnabled(True)
    main_window.btn_cancel.setText("⏹ 中止")

    main_window.on_cancel_pipeline()

    # 应当调用 worker.cancel()，并将文案切换为正在中止
    mock_worker.cancel.assert_called_once()
    assert main_window.btn_cancel.text() == "正在中止…"
    assert main_window.btn_cancel.isEnabled() is False


def test_clear_and_remove_items(tmp_path: Path, main_window):
    """测试清空列表与移除选中项。"""
    f1 = tmp_path / "1.doc"
    f2 = tmp_path / "2.docx"
    f1.write_bytes(b"1")
    f2.write_bytes(b"2")

    main_window.add_paths([f1, f2])
    assert main_window.table.rowCount() == 2

    # 选中第一行并移除
    main_window.table.selectRow(0)
    main_window.on_remove_selected()
    assert main_window.table.rowCount() == 1
    assert main_window.files_list == [f2.resolve()]

    # 清空列表
    main_window.on_clear_files()
    assert main_window.table.rowCount() == 0
    assert len(main_window.files_list) == 0


def test_dynamic_pdf_conflict_dialog_interception(tmp_path: Path, main_window, monkeypatch):
    """测试 dynamic + PDF 冲突时拦截并弹出警告。"""
    doc = tmp_path / "a.docx"
    doc.write_bytes(b"test")
    main_window.add_paths([doc])

    main_window.radio_mode_run.setChecked(True)
    main_window.chk_pdf.setChecked(True)
    main_window.combo_page_mode.setCurrentIndex(3)  # dynamic

    alert_called = False

    def mock_critical(parent, title, text):
        nonlocal alert_called
        alert_called = True
        assert "「Word 动态域」页码模式不能与 PDF 导出同时使用" in text

    monkeypatch.setattr(QMessageBox, "critical", mock_critical)
    main_window.on_start_pipeline()

    assert alert_called is True
    # 保证未启动 worker
    assert main_window.worker is None
