"""
tests/core/test_docx_normalizer.py
验证 LibreOffice 转换产物规范化的三类修复：
1. 段落标记字号补齐为段落 run 最大字号
2. "空格+U+2006" 空白填充折叠
3. LO CJK 替代字体名换回 Windows 标准中文字体
"""

import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from wdt.core.docx_normalizer import normalize_docx

DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">项目 </w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>a</w:t></w:r><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>b</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>keep</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>
<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">\u2006 \u2006 填充</w:t></w:r></w:p>
</w:body>
</w:document>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Times New Roman" w:eastAsia="Noto Serif SC" w:hAnsi="Times New Roman"/>
<w:sz w:val="24"/><w:szCs w:val="24"/>
</w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>
"""

FONT_TABLE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:font w:name="Noto Serif SC"/>
</w:fonts>
"""


def _build_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/fontTable.xml", FONT_TABLE_XML)


def _read(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def test_normalize_paragraph_mark_sz(tmp_path: Path):
    """空段落标记补齐为 run 最大字号；已有字号的标记不受影响；XML 结构保持合法"""
    docx = tmp_path / "t.docx"
    _build_docx(docx)

    stats = normalize_docx(docx)
    xml = _read(docx, "word/document.xml")

    # 整体仍是合法 XML（防止开闭标签被破坏）
    ET.fromstring(xml)

    # 段落1：空标记 + run sz=20 -> 标记获得 sz=20，<w:p> 开标签保留
    assert '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>' in xml
    # 段落2：无 pPr，多 run 取最大 22，并新建 pPr
    assert '<w:p><w:pPr><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>' in xml
    # 段落3：标记已有 sz=24，保持原样
    assert xml.count('<w:sz w:val="24"/>') == 1
    # 空段落（无 run）不添加 pPr
    assert '<w:pStyle w:val="Normal"/></w:pPr></w:p>' in xml
    assert stats["paragraph_mark"] == 3


def test_normalize_thin_space_padding(tmp_path: Path):
    """U+2006 空白对折叠为普通空格"""
    docx = tmp_path / "t.docx"
    _build_docx(docx)

    stats = normalize_docx(docx)
    xml = _read(docx, "word/document.xml")

    assert "\u2006" not in xml
    assert "  填充" in xml  # 两对 -> 两个普通空格
    assert stats["thin_space"] > 0


def test_normalize_cjk_font_swap(tmp_path: Path):
    """LO 替代字体换回 Windows 标准中文字体（document/styles/fontTable）"""
    docx = tmp_path / "t.docx"
    _build_docx(docx)

    stats = normalize_docx(docx)

    assert "Noto Serif SC" not in _read(docx, "word/document.xml")
    assert 'w:eastAsia="宋体"' in _read(docx, "word/styles.xml")
    assert 'w:name="宋体"' in _read(docx, "word/fontTable.xml")
    assert stats["font_swap"] >= 2


def test_normalize_idempotent(tmp_path: Path):
    """二次规范化不产生变化"""
    docx = tmp_path / "t.docx"
    _build_docx(docx)
    normalize_docx(docx)
    snapshot = {
        n: _read(docx, n)
        for n in ("word/document.xml", "word/styles.xml", "word/fontTable.xml")
    }

    stats = normalize_docx(docx)
    assert stats["paragraph_mark"] == 0
    assert stats["thin_space"] == 0
    assert stats["font_swap"] == 0
    for n, text in snapshot.items():
        assert _read(docx, n) == text
