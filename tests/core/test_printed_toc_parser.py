"""打印版目录解析单元测试。"""

from __future__ import annotations

import zlib
from pathlib import Path

from pypdf import PdfReader

from wdt.core.pdf_splitter import plan_volume_splits, resolve_outline, split_pdf_by_outline
from wdt.core.printed_toc_parser import extract_printed_toc


def _make_text_pdf(path: Path, page_texts: list[str], page_w: float = 595.0, page_h: float = 842.0) -> Path:
    """构造每页渲染给定纯文本行的简易 PDF（Helvetica，无需外部依赖）。"""
    objects: list[bytes] = []

    def add(obj: bytes) -> int:
        objects.append(obj)
        return len(objects)

    add(b"<< /Type /Catalog /Pages 2 0 R >>")
    add(b"__PAGES__")  # placeholder, index 2
    add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    font_obj = 3

    page_ids = []
    for text in page_texts:
        stream_lines = ["BT", "36 800 Td", "14 TL", f"/F1 10 Tf"]
        for line in text.splitlines():
            escaped = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            stream_lines.append(f"({escaped}) Tj T*")
        stream_lines.append("ET")
        stream = "\n".join(stream_lines).encode("latin-1", errors="replace")
        content_id = add(
            b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream"
        )
        page_id = add(
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %f %f] "
            b"/Resources << /Font << /F1 %d 0 R >> >> /Contents %d 0 R >>"
            % (page_w, page_h, font_obj, content_id)
        )
        page_ids.append(page_id)

    kids = b" ".join(b"%d 0 R" % pid for pid in page_ids)
    objects[1] = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, len(page_ids))

    out = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % i + obj + b"\nendobj\n"
    xref_pos = len(out)
    out += b"xref\n0 %d\n" % (len(objects) + 1)
    out += b"0000000000 65535 f \n"
    for off in offsets[1:]:
        out += b"%010d 00000 n \n" % off
    out += (
        b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF"
        % (len(objects) + 1, xref_pos)
    )
    path.write_bytes(bytes(out))
    return path


_TOC_PAGE = """05NOV2026:12:18:39  ClinStReport              Page 1 of 20
Table of Content
Listing 16.2.1.1 Screen Failure Subjects .................................................... 3
Listing 16.2.1.2 Discontinued Subjects (Safety/Full Analysis Set) .......................... 5
Listing 16.2.2.1 Protocol Deviations (Safety/Full Analysis Set) ............................ 8
Listing 16.2.7.9 Time of First TEAE (SS) ..................................................... 15
Listing 16.2.8.1 Hematology (SS) ............................................................. 17
"""

# 目录第二页（无 Table of Content 标题，但条目很多且紧邻目录首页）
_TOC_PAGE_2 = """05NOV2026:12:18:39  ClinStReport              Page 2 of 20
Listing 16.2.8.2 Chemistry (SS) ............................................................ 18
Listing 16.2.9.1 Vital Signs (SS) .......................................................... 20
Listing 16.2.9.2 ECG (SS) .................................................................. 21
"""

_BODY_PAGE = """05NOV2026:12:15:46  ClinStReport  Page 3 of 20

Listing 16.2.1.1 Screen Failure Subjects
Subject ID    Study Phase
SS1001        SAD
"""


def test_extract_printed_toc_basic(tmp_path: Path):
    pdf = _make_text_pdf(
        tmp_path / "toc.pdf",
        [_TOC_PAGE, _TOC_PAGE_2, _BODY_PAGE] + ["body"] * 19,
    )
    reader = PdfReader(str(pdf))
    nodes, toc_pages = extract_printed_toc(reader)

    # 两个目录页均被识别
    assert toc_pages == {0, 1}
    # 8 个条目全部解析
    assert len(nodes) == 8
    assert nodes[0].title == "Listing 16.2.1.1 Screen Failure Subjects"
    assert nodes[0].page_index == 2  # 打印页码 3 -> 物理页 2
    assert nodes[0].level == 1
    assert nodes[-1].title == "Listing 16.2.9.2 ECG (SS)"
    assert nodes[-1].page_index == 20
    # 层级已归一化：最小层级 (16.2.x) 为 1
    assert min(n.level for n in nodes) == 1
    # 页码严格递增
    pages = [n.page_index for n in nodes]
    assert pages == sorted(pages)


def test_extract_printed_toc_rejects_non_toc(tmp_path: Path):
    pdf = _make_text_pdf(tmp_path / "plain.pdf", [_BODY_PAGE] * 5)
    reader = PdfReader(str(pdf))
    nodes, toc_pages = extract_printed_toc(reader)
    assert nodes == []
    assert toc_pages == set()


def test_plan_volume_splits_uses_printed_toc(tmp_path: Path):
    """无大纲 PDF 应通过打印目录重建节点并按章节聚类分卷。"""
    pdf = _make_text_pdf(
        tmp_path / "plan.pdf",
        [_TOC_PAGE, _BODY_PAGE] + ["body"] * 19,
    )
    reader = PdfReader(str(pdf))
    vols = plan_volume_splits(reader, max_pages_per_volume=10)

    assert len(vols) >= 2
    # 目录页（物理页 0）被剔除，卷 1 从正文开始
    assert vols[0].page_start == 1
    assert "16.2.1.1" in " ".join(n.title for n in vols[0].inner_nodes)


def test_split_pdf_by_outline_with_printed_toc(tmp_path: Path):
    """端到端：无大纲 PDF 拆分后分卷应带真实目录条目。"""
    pdf = _make_text_pdf(
        tmp_path / "split.pdf",
        [_TOC_PAGE, _BODY_PAGE] + ["body"] * 19,
    )
    results = split_pdf_by_outline(pdf, out_dir=tmp_path / "out", with_visual_toc=True, max_pages=10)
    assert len(results) >= 2
    r1 = PdfReader(str(results[0].output_path))
    # 卷 1 大纲含 Table of Contents 书签与真实条目
    outline_titles = [n.title for n in resolve_outline(r1)[0]]
    assert any("Table of Contents" in t for t in outline_titles)
    assert any("16.2.1.1" in t for t in outline_titles)
    # 目录页未混入正文
    body_text = (r1.pages[1].extract_text() or "") + (r1.pages[2].extract_text() or "")
    assert "Table of Content" not in body_text


def test_measure_toc_row_baselines(tmp_path: Path):
    """实测行基线定位：折行/加行场景下仍逐条匹配页码。"""
    from wdt.core.pdf_splitter import _measure_toc_row_baselines

    page_text = (
        "目    录\n"
        "Volume 1 of 2: long volume title that may wrap\n"
        "Listing 16.2.1.1 Screen Failure Subjects .................................................... 1\n"
        "Listing 16.2.1.2 Discontinued Subjects (Safety Set) ....................................... 26\n"
        "Listing 16.2.2.1 Protocol Deviations ........................................................ 28\n"
        "Part 1 | footer\n"
    )
    pdf = _make_text_pdf(tmp_path / "measure.pdf", [page_text])
    reader = PdfReader(str(pdf))
    items = [
        ("Listing 16.2.1.1 Screen Failure Subjects", 1, 1),
        ("Listing 16.2.1.2 Discontinued Subjects (Safety Set)", 26, 1),
        ("Listing 16.2.2.1 Protocol Deviations", 28, 1),
    ]
    baselines = _measure_toc_row_baselines(reader.pages[0], items)
    assert baselines is not None
    assert len(baselines) == 3
    # 基线自上而下递减
    assert baselines[0] > baselines[1] > baselines[2]

    # 页码无法匹配时返回 None（调用方回退固定几何模型）
    bad_items = [("X", 999, 1)]
    assert _measure_toc_row_baselines(reader.pages[0], bad_items) is None
