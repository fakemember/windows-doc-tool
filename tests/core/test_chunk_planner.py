from pathlib import Path

from wdt.core.chunk_planner import Chunk, plan_chunks


def test_plan_chunks_empty():
    assert plan_chunks([], jobs=4) == []


def test_plan_chunks_greedy_balancing(tmp_path: Path):
    # 创建 4 个不同大小的文件
    f1 = tmp_path / "f1.doc"
    f2 = tmp_path / "f2.doc"
    f3 = tmp_path / "f3.doc"
    f4 = tmp_path / "f4.doc"

    f1.write_bytes(b"x" * 1000)
    f2.write_bytes(b"x" * 800)
    f3.write_bytes(b"x" * 600)
    f4.write_bytes(b"x" * 400)

    # 2 个 job 桶：
    # 降序：f1(1000), f2(800), f3(600), f4(400)
    # f1 -> bucket0 (1000)
    # f2 -> bucket1 (800)
    # f3 -> bucket1 (800 + 600 = 1400)
    # f4 -> bucket0 (1000 + 400 = 1400)
    # 两桶均衡，大小各 1400
    chunks = plan_chunks([f1, f2, f3, f4], jobs=2)
    assert len(chunks) == 2
    assert chunks[0].total_bytes == 1400
    assert chunks[1].total_bytes == 1400
    assert chunks[0].index == 0
    assert chunks[1].index == 1

    all_files = set(chunks[0].files + chunks[1].files)
    assert all_files == {f1, f2, f3, f4}


def test_plan_chunks_exceeds_max_chunk_bytes(tmp_path: Path):
    f1 = tmp_path / "large1.doc"
    f2 = tmp_path / "large2.doc"
    f3 = tmp_path / "large3.doc"

    f1.write_bytes(b"a" * 80)
    f2.write_bytes(b"b" * 60)
    f3.write_bytes(b"c" * 80)

    # jobs=1，但 max_chunk_bytes=150
    # 降序：f1(80), f3(80), f2(60)
    # chunk 0: [f1(80), f2(60)] -> 140 <= 150
    # chunk 1: [f3(80)] -> 80 <= 150
    # 总计切分为 2 个 chunk
    chunks = plan_chunks([f1, f2, f3], jobs=1, max_chunk_bytes=150)
    assert len(chunks) == 2
    assert all(c.total_bytes <= 150 for c in chunks)
    assert sum(c.total_bytes for c in chunks) == 220


def test_plan_chunks_single_huge_file(tmp_path: Path):
    f = tmp_path / "huge.doc"
    f.write_bytes(b"x" * 500)

    # 单文件大于 max_chunk_bytes 时，必须独立成块且不无限循环
    chunks = plan_chunks([f], jobs=2, max_chunk_bytes=200)
    assert len(chunks) == 1
    assert chunks[0].files == [f]
    assert chunks[0].total_bytes == 500
