from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Sequence


@dataclass
class Chunk:
    index: int
    files: list[Path]
    total_bytes: int


def plan_chunks(
    files: Sequence[Path],
    jobs: int,
    max_chunk_bytes: int = 256 * 1024 * 1024,
) -> list[Chunk]:
    """按文件大小均衡规划分块。

    1. 文件按大小降序贪心分配到当前最小桶（LPT 算法）；
    2. 若桶总大小超过 max_chunk_bytes，则按大小切分为多个子块。
    """
    if not files:
        return []

    num_buckets = max(1, jobs)
    path_list = [Path(p) for p in files]

    # 获取文件大小并降序排列
    files_with_size: list[tuple[Path, int]] = []
    for p in path_list:
        try:
            sz = p.stat().st_size if p.is_file() else 0
        except OSError:
            sz = 0
        files_with_size.append((p, sz))

    files_with_size.sort(key=lambda item: item[1], reverse=True)

    # 贪心分配到当前总大小最小的桶
    buckets: list[list[tuple[Path, int]]] = [[] for _ in range(num_buckets)]
    bucket_sizes = [0] * num_buckets

    for p, sz in files_with_size:
        min_idx = bucket_sizes.index(min(bucket_sizes))
        buckets[min_idx].append((p, sz))
        bucket_sizes[min_idx] += sz

    # 对超限桶进行切分
    result_chunks: list[tuple[list[Path], int]] = []
    for bucket in buckets:
        if not bucket:
            continue
        total = sum(sz for _, sz in bucket)
        if total <= max_chunk_bytes:
            result_chunks.append(([p for p, _ in bucket], total))
        else:
            cur_files: list[Path] = []
            cur_bytes = 0
            for p, sz in bucket:
                if cur_files and (cur_bytes + sz > max_chunk_bytes):
                    result_chunks.append((cur_files, cur_bytes))
                    cur_files = [p]
                    cur_bytes = sz
                else:
                    cur_files.append(p)
                    cur_bytes += sz
            if cur_files:
                result_chunks.append((cur_files, cur_bytes))

    return [
        Chunk(index=i, files=f_list, total_bytes=b_count)
        for i, (f_list, b_count) in enumerate(result_chunks)
    ]
