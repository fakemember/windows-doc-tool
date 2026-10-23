"""
tests/toc/test_zip_rewriter.py
验证 zip_rewriter 的双策略原子写回：zipfile 全量重建与 7z 增量更新、磁盘预检及异常回滚。
"""

import os
import shutil
import zipfile
from pathlib import Path
from unittest.mock import patch

import pytest

from tests.fixtures_factory import build_docx
from wdt.toc.zip_rewriter import rewrite_zip


def test_zipfile_rewrite_roundtrip(tmp_path: Path):
    """验证默认小文件 zipfile 路径的更新与新增条目能力"""
    docx_file = tmp_path / "sample.docx"
    build_docx(
        docx_file,
        [{"text": "原始段落"}],
        extra_files={"custom/keep.txt": b"preserve me"},
    )

    new_doc_xml = b"<w:document><w:body><w:p><w:t>Updated Content</w:t></w:p></w:body></w:document>"
    new_settings_xml = b"<w:settings><w:updateFields w:val=\"true\"/></w:settings>"

    rewrite_zip(
        docx_file,
        {
            "word/document.xml": new_doc_xml,
            "word/settings.xml": new_settings_xml,
        },
        big_threshold=100 * 1024 * 1024,
    )

    # 验证更新结果
    with zipfile.ZipFile(docx_file, "r") as z:
        assert z.read("word/document.xml") == new_doc_xml
        assert z.read("word/settings.xml") == new_settings_xml
        # 原有未修改条目完好
        assert z.read("custom/keep.txt") == b"preserve me"


def test_seven_zip_rewrite_roundtrip(tmp_path: Path):
    """调低阈值强制触发 7z 增量更新路径，验证更新正确性"""
    if not (shutil.which("7za") or shutil.which("7z")):
        pytest.skip("系统中未找到 7z/7za")

    docx_file = tmp_path / "sample_7z.docx"
    build_docx(
        docx_file,
        [{"text": "7z 模式段落"}],
        extra_files={"custom/keep.txt": b"7z keep"},
    )

    new_doc_xml = b"<w:document><w:body><w:p><w:t>7z Updated</w:t></w:p></w:body></w:document>"

    # 设置 big_threshold=0 强制走 7z
    rewrite_zip(
        docx_file,
        {"word/document.xml": new_doc_xml},
        big_threshold=0,
    )

    with zipfile.ZipFile(docx_file, "r") as z:
        assert z.read("word/document.xml") == new_doc_xml
        assert z.read("custom/keep.txt") == b"7z keep"


def test_disk_space_precheck_failure(tmp_path: Path):
    """验证磁盘空间不足时（< 1.3 倍文件大小）抛出 RuntimeError 且原文件未被破坏"""
    if not (shutil.which("7za") or shutil.which("7z")):
        pytest.skip("系统中未找到 7z/7za")

    docx_file = tmp_path / "sample_disk.docx"
    build_docx(docx_file, [{"text": "空间检测段落"}])
    orig_bytes = docx_file.read_bytes()

    # 模拟可用磁盘空间仅 10 字节
    with patch("shutil.disk_usage") as mock_usage:
        mock_usage.return_value = os.statvfs_result((4096, 4096, 1000, 1000, 10, 10, 10, 10, 0, 255))
        # 兼容 namedtuple: total, used, free
        from collections import namedtuple
        DiskUsage = namedtuple("DiskUsage", ["total", "used", "free"])
        mock_usage.return_value = DiskUsage(total=1000000, used=999990, free=10)

        with pytest.raises(RuntimeError, match="磁盘空间不足"):
            rewrite_zip(
                docx_file,
                {"word/document.xml": b"<dummy/>"},
                big_threshold=0,
            )

    # 验证原文件未损坏
    assert docx_file.read_bytes() == orig_bytes
