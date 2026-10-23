import io
import zipfile
from pathlib import Path

from wdt.core.format_sniffer import sniff


def test_sniff_rtf(tmp_path: Path):
    rtf_file = tmp_path / "sample.rtf"
    rtf_file.write_bytes(b"{\\rtf1\\ansi\\deff0 Hello World}")
    assert sniff(rtf_file) == "rtf"

    # 伪装成 .doc 的真 rtf
    fake_doc = tmp_path / "fake.doc"
    fake_doc.write_bytes(b"{\\rtf1\\ansi Sample}")
    assert sniff(fake_doc) == "rtf"


def test_sniff_doc(tmp_path: Path):
    doc_file = tmp_path / "sample.doc"
    # OLE2 魔数
    doc_file.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 100)
    assert sniff(doc_file) == "doc"


def test_sniff_docx(tmp_path: Path):
    docx_file = tmp_path / "sample.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("word/document.xml", "<w:document/>")
    docx_file.write_bytes(buf.getvalue())
    assert sniff(docx_file) == "docx"


def test_sniff_other_zip(tmp_path: Path):
    zip_file = tmp_path / "sample.zip"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("test.txt", "hello")
    zip_file.write_bytes(buf.getvalue())
    assert sniff(zip_file) == "unknown"


def test_sniff_unknown_and_missing(tmp_path: Path):
    txt_file = tmp_path / "plain.txt"
    txt_file.write_text("Hello plain text")
    assert sniff(txt_file) == "unknown"

    missing = tmp_path / "non_existent.doc"
    assert sniff(missing) == "unknown"

    empty_file = tmp_path / "empty.doc"
    empty_file.write_bytes(b"")
    assert sniff(empty_file) == "unknown"
