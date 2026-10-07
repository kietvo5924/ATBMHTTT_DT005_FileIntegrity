"""
Parallel Hasher: Tối ưu hóa tính toán băm song song đa luồng cho tập dữ liệu lớn.
Tận dụng tối đa băng thông I/O và năng lực đa nhân CPU.
"""
import os
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Tuple, Callable, Optional


def hash_single_file_streaming(filepath: str, chunk_size: int = 65536) -> str:
    """Tính SHA-256 của tệp theo từng chunk đệm 64KB."""
    hasher = hashlib.sha256()
    with open(filepath, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            hasher.update(chunk)
    return hasher.hexdigest()


def _hash_worker(args: Tuple[str, str, int]) -> Tuple[str, str]:
    """Hàm worker xử lý 1 file: trả về (rel_path, hash_hex)."""
    full_path, rel_path, chunk_size = args
    h = hash_single_file_streaming(full_path, chunk_size)
    return rel_path, h


def parallel_hash_files(
    root_dir: str,
    max_workers: Optional[int] = None,
    chunk_size: int = 65536,
    ignore_files: Optional[List[str]] = None
) -> Dict[str, str]:
    """
    Tính mã băm song song cho tất cả các file trong thư mục root_dir.
    Trả về: Dict[rel_path, sha256_hex]
    """
    root_dir = os.path.abspath(root_dir)
    ignore = set(ignore_files or ["manifest_sha256.json", "manifest_hmac.json", "merkle_standard.json", "merkle_hd.json"])
    
    tasks = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename in ignore:
                continue
            full_path = os.path.join(dirpath, filename)
            rel_path = os.path.relpath(full_path, root_dir).replace("\\", "/")
            if rel_path in ignore:
                continue
            tasks.append((full_path, rel_path, chunk_size))
            
    results: Dict[str, str] = {}
    if not tasks:
        return results

    # Sử dụng số worker phù hợp với số nhân CPU
    workers = max_workers or min(32, (os.cpu_count() or 4) * 2)
    
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(_hash_worker, task): task[1] for task in tasks}
        for future in as_completed(futures):
            rel_path, file_hash = future.result()
            results[rel_path] = file_hash

    return dict(sorted(results.items()))
