import shutil
import zipfile
from pathlib import Path

import pytest
from wdt.core.chunk_planner import Chunk
from wdt.core.converter_worker import (
    _prepare_profile_dir,
    _workload_mb,
    convert_chunk,
    default_timeout_sec,
)
from wdt.core.engine_finder import Engine, find_engine

_HAS_SOFFICE = shutil.which("soffice") is not None


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_convert_chunk_docx(tmp_path: Path):
    engine = find_engine()
    src_file = tmp_path / "doc1.rtf"
    src_file.write_bytes(b"{\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Courier;}}\\f0\\fs20 Line 1\\par}")

    dst_file = tmp_path / "out" / "doc1.docx"
    chunk = Chunk(index=0, files=[src_file], total_bytes=src_file.stat().st_size)
    dst_map = {src_file: dst_file}

    results = convert_chunk(engine, chunk, dst_map, target="docx")
    assert len(results) == 1
    res = results[0]
    assert res.ok, f"Convert failed: {res.error}"
    assert res.dst == dst_file
    assert dst_file.is_file()
    assert dst_file.stat().st_size > 0


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_convert_chunk_pdf_and_gray(tmp_path: Path):
    engine = find_engine()
    src_file = tmp_path / "doc2.rtf"
    src_file.write_bytes(b"{\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Courier;}}\\f0\\fs20 Hello PDF\\par}")

    # 1. 彩色 PDF
    dst_pdf = tmp_path / "out" / "doc2.pdf"
    chunk = Chunk(index=0, files=[src_file], total_bytes=src_file.stat().st_size)
    dst_map = {src_file: dst_pdf}
    results = convert_chunk(engine, chunk, dst_map, target="pdf", gray=False)
    assert len(results) == 1
    assert results[0].ok
    assert dst_pdf.is_file()
    assert dst_pdf.stat().st_size > 0

    # 2. 灰度 PDF
    dst_gray = tmp_path / "out" / "doc2_gray.pdf"
    dst_map_gray = {src_file: dst_gray}
    results_gray = convert_chunk(engine, chunk, dst_map_gray, target="pdf", gray=True)
    assert len(results_gray) == 1
    assert results_gray[0].ok
    assert dst_gray.is_file()
    assert dst_gray.stat().st_size > 0


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_convert_chunk_sniff_grouping(tmp_path: Path):
    engine = find_engine()
    # 一个标准 .rtf，一个伪装成 .doc 的真 rtf
    f1 = tmp_path / "real.rtf"
    f2 = tmp_path / "pseudo.doc"
    f1.write_bytes(b"{\\rtf1\\ansi File 1}")
    f2.write_bytes(b"{\\rtf1\\ansi File 2}")

    dst1 = tmp_path / "out" / "real.docx"
    dst2 = tmp_path / "out" / "pseudo.docx"

    chunk = Chunk(
        index=0,
        files=[f1, f2],
        total_bytes=f1.stat().st_size + f2.stat().st_size,
    )
    dst_map = {f1: dst1, f2: dst2}

    results = convert_chunk(engine, chunk, dst_map, target="docx")
    assert len(results) == 2
    assert results[0].ok
    assert results[1].ok
    assert dst1.is_file() and dst1.stat().st_size > 0
    assert dst2.is_file() and dst2.stat().st_size > 0


def test_convert_chunk_unsupported_gray(tmp_path: Path):
    fake_engine = Engine(soffice=Path("/bin/soffice"), source="system", version=(7, 2))
    f = tmp_path / "doc.rtf"
    f.write_bytes(b"{\\rtf1\\ansi test}")
    chunk = Chunk(index=0, files=[f], total_bytes=10)

    results = convert_chunk(fake_engine, chunk, {f: tmp_path / "doc.pdf"}, target="pdf", gray=True)
    assert len(results) == 1
    assert not results[0].ok
    assert "LibreOffice >= 7.4 required for gray" in results[0].error


def test_convert_chunk_invalid_target(tmp_path: Path):
    fake_engine = Engine(soffice=Path("/bin/soffice"), source="system", version=(26, 8))
    f = tmp_path / "doc.rtf"
    f.write_bytes(b"{\\rtf1\\ansi test}")
    chunk = Chunk(index=0, files=[f], total_bytes=10)

    with pytest.raises(ValueError, match="Unsupported target"):
        convert_chunk(fake_engine, chunk, {f: tmp_path / "doc.xyz"}, target="xyz")  # type: ignore


def _make_compressible_docx(path: Path, xml_size: int) -> Path:
    """构造高压缩比 docx：解压后 document.xml 为 xml_size，磁盘体积极小。"""
    payload = b"<w:p><w:t>" + b"x" * (xml_size - 20) + b"</w:t></w:p>"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", payload)
        z.writestr("[Content_Types].xml", b"<?xml version='1.0'?>")
    return path


def test_workload_mb_uses_uncompressed_docx_size(tmp_path: Path):
    docx = _make_compressible_docx(tmp_path / "big.docx", xml_size=8 * 1024 * 1024)
    assert docx.stat().st_size < 1024 * 1024  # 磁盘体积远小于解压后
    mb = _workload_mb(docx)
    assert 7 < mb < 9


def test_workload_mb_falls_back_to_disk_size(tmp_path: Path):
    rtf = tmp_path / "a.rtf"
    rtf.write_bytes(b"{\\rtf1\\ansi " + b"y" * (3 * 1024 * 1024) + b"}")
    mb = _workload_mb(rtf)
    assert 2 < mb < 4


def test_default_timeout_scales_with_uncompressed_size(tmp_path: Path):
    docx = _make_compressible_docx(tmp_path / "big.docx", xml_size=8 * 1024 * 1024)
    docx_timeout = default_timeout_sec([docx])
    # docx 源按解压后体积以 2 s/MiB 估算
    assert docx_timeout >= 300 + 2 * 7

    rtf = tmp_path / "a.rtf"
    rtf.write_bytes(b"{\\rtf1\\ansi " + b"y" * (1024 * 1024) + b"}")
    rtf_timeout = default_timeout_sec([rtf])
    # rtf/doc 源（高密表格排版）按 20 s/MiB 估算，高于同体积 docx
    assert 300 + 20 <= rtf_timeout
    assert rtf_timeout > docx_timeout


# --- SAS RTF 头部声明修复（\ansicpg936 与 {\stylesheet} 注入） --------------


def _make_gbk_rtf(
    with_ansicpg: bool = False,
    with_stylesheet: bool = False,
    body: bytes | None = None,
) -> bytes:
    cpg = b"\\ansicpg936" if with_ansicpg else b""
    ss = b"{\\stylesheet{\\s1 heading 1;}}" if with_stylesheet else b""
    payload = body if body is not None else "筛选情况（筛选受试者）".encode("gbk")
    return (
        b"{\\rtf1\\ansi"
        + cpg
        + ss
        + b"\\deff0 {\\fonttbl{\\f0 Arial;}}\\f0\\fs24\\s1 "
        + payload
        + b"\\par}"
    )


def test_rtf_missing_declarations_gbk_and_stylesheet(tmp_path: Path):
    """同时缺失 \\ansicpg 与 \\stylesheet，且正文包含 \\s1 引用与 GBK 内容"""
    from wdt.core.converter_worker import _rtf_missing_declarations, _rtf_needs_staging

    rtf = tmp_path / "sample.rtf"
    rtf.write_bytes(_make_gbk_rtf())

    needs_ansicpg, levels = _rtf_missing_declarations(rtf)
    assert needs_ansicpg is True
    assert levels == [1]
    assert _rtf_needs_staging(rtf) is True


def test_rtf_missing_declarations_already_declared(tmp_path: Path):
    """已声明 \\ansicpg 且包含 \\stylesheet → 无需暂存"""
    from wdt.core.converter_worker import _rtf_missing_declarations, _rtf_needs_staging

    rtf = tmp_path / "declared.rtf"
    rtf.write_bytes(_make_gbk_rtf(with_ansicpg=True, with_stylesheet=True))

    needs_ansicpg, levels = _rtf_missing_declarations(rtf)
    assert needs_ansicpg is False
    assert levels == []
    assert _rtf_needs_staging(rtf) is False


def test_rtf_missing_declarations_ascii_no_styles(tmp_path: Path):
    """纯 ASCII 且正文无 \\sN 引用 → 无需暂存"""
    from wdt.core.converter_worker import _rtf_missing_declarations, _rtf_needs_staging

    rtf = tmp_path / "ascii.rtf"
    rtf.write_bytes(b"{\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Arial;}}\\f0 Hello World\\par}")

    needs_ansicpg, levels = _rtf_missing_declarations(rtf)
    assert needs_ansicpg is False
    assert levels == []
    assert _rtf_needs_staging(rtf) is False


def test_stage_rtf_fixed_injects_both(tmp_path: Path):
    r"""暂存副本同时注入 \ansicpg936 与 {\stylesheet}，且跨 8KB 采样边界无损"""
    from wdt.core.converter_worker import _stage_rtf_fixed

    src = tmp_path / "big.doc"
    body_tail = b"x" * 100
    src.write_bytes(_make_gbk_rtf()[:-1].ljust(8192, b"y") + body_tail + b"}")
    stage_dir = tmp_path / "stage"

    staged = _stage_rtf_fixed(src, stage_dir)

    assert staged == stage_dir / src.name
    staged_bytes = staged.read_bytes()
    assert staged_bytes.startswith(b"{\\rtf1\\ansi\\ansicpg936{\\stylesheet")
    assert b"\\outlinelevel0 heading 1;" in staged_bytes

    orig = src.read_bytes()
    anchor = b"{\\rtf1\\ansi"
    insert_pos = orig.find(anchor) + len(anchor)
    # 头部注入后，后面的原始内容逐字节保留
    injected_header = staged_bytes[: staged_bytes.find(b"\\deff0")]
    assert b"\\ansicpg936" in injected_header
    assert b"{\\stylesheet" in injected_header
    assert orig[insert_pos:] == staged_bytes[len(injected_header) :]


def test_prepare_profile_dir_disables_updater_and_firstrun(tmp_path: Path):
    profile_dir = tmp_path / "profile"
    _prepare_profile_dir(profile_dir)
    xcu = profile_dir / "user" / "registrymodifications.xcu"
    assert xcu.is_file()
    content = xcu.read_text(encoding="utf-8")
    assert '/org.openoffice.Office.Update/Update' in content
    assert '<prop oor:name="Enabled" oor:op="fuse"><value>false</value></prop>' in content
    assert '<prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop>' in content
