"""
Baseline 1: Flat SHA-256 Manifest
Kiểm tra toàn vẹn bằng cách tính mã băm SHA-256 cho từng tệp và lưu vào danh sách phẳng (Manifest).
"""
import os
import json
import time
import hashlib
from typing import Dict, Tuple, List, Optional


def hash_file_sha256(filepath: str, chunk_size: int = 65536) -> str:
    """Tính mã băm SHA-256 của tệp tin theo từng khối nhỏ (streaming)."""
    hasher = hashlib.sha256()
    with open(filepath, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            hasher.update(chunk)
    return hasher.hexdigest()


class SHA256Manifest:
    """
    Hệ thống kiểm tra toàn vẹn dựa trên Flat SHA-256 Manifest.
    - Lưu trữ danh sách: {rel_path: sha256_hex}
    - Tính thêm Root Hash tổng bằng cách băm chuỗi nối các băm con theo thứ tự từ điển.
    """
    def __init__(self, root_dir: str):
        self.root_dir = os.path.abspath(root_dir)
        self.manifest: Dict[str, str] = {}
        self.root_hash: str = ""

    def build_manifest(self) -> Tuple[float, int]:
        """
        Quét toàn bộ thư mục và tạo manifest.
        Trả về: (thời gian thực thi (s), tổng số byte manifest)
        """
        start_time = time.perf_counter()
        manifest = {}
        
        for dirpath, _, filenames in os.walk(self.root_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, self.root_dir).replace("\\", "/")
                # Bỏ qua tệp manifest nếu nằm trong cùng thư mục
                if rel_path in ("manifest_sha256.json", "manifest_hmac.json"):
                    continue
                file_hash = hash_file_sha256(full_path)
                manifest[rel_path] = file_hash
                
        self.manifest = dict(sorted(manifest.items()))
        self._compute_aggregate_root()
        elapsed = time.perf_counter() - start_time
        return elapsed, self.get_storage_overhead()

    def _compute_aggregate_root(self) -> None:
        """Tính mã băm gốc tổng hợp (Aggregate Root Hash)."""
        aggregate = hashlib.sha256()
        for path in sorted(self.manifest.keys()):
            aggregate.update(path.encode("utf-8"))
            aggregate.update(bytes.fromhex(self.manifest[path]))
        self.root_hash = aggregate.hexdigest()

    def verify_all(self) -> Tuple[bool, List[str], float]:
        """
        Xác thực tính toàn vẹn của toàn bộ thư mục dựa trên manifest đã lưu.
        Trả về: (toàn vẹn hay không, danh sách các file bị vi phạm, thời gian xác thực)
        """
        start_time = time.perf_counter()
        violations = []
        
        current_files = set()
        for dirpath, _, filenames in os.walk(self.root_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, self.root_dir).replace("\\", "/")
                if rel_path in ("manifest_sha256.json", "manifest_hmac.json"):
                    continue
                current_files.add(rel_path)
                
                if rel_path not in self.manifest:
                    violations.append(f"UNREGISTERED_FILE: {rel_path}")
                else:
                    curr_hash = hash_file_sha256(full_path)
                    if curr_hash != self.manifest[rel_path]:
                        violations.append(f"MODIFIED_FILE: {rel_path}")

        # Kiểm tra file bị xóa
        for expected_path in self.manifest:
            if expected_path not in current_files:
                violations.append(f"MISSING_FILE: {expected_path}")

        elapsed = time.perf_counter() - start_time
        is_valid = len(violations) == 0
        return is_valid, violations, elapsed

    def verify_single_file(self, rel_path: str) -> Tuple[bool, float]:
        """Xác thực một file đơn lẻ khi đã có manifest tin cậy."""
        rel_path = rel_path.replace("\\", "/")
        if rel_path not in self.manifest:
            return False, 0.0
            
        full_path = os.path.join(self.root_dir, rel_path)
        if not os.path.exists(full_path):
            return False, 0.0
            
        start = time.perf_counter()
        curr_hash = hash_file_sha256(full_path)
        is_valid = (curr_hash == self.manifest[rel_path])
        elapsed = time.perf_counter() - start
        return is_valid, elapsed

    def update_file(self, rel_path: str) -> Tuple[float, int]:
        """
        Cập nhật khi 1 tệp bị thay đổi nội dung.
        Đo chi phí: tính lại hash file đó + cập nhật lại aggregate root hash.
        """
        rel_path = rel_path.replace("\\", "/")
        full_path = os.path.join(self.root_dir, rel_path)
        
        start_time = time.perf_counter()
        if os.path.exists(full_path):
            new_hash = hash_file_sha256(full_path)
            self.manifest[rel_path] = new_hash
        else:
            self.manifest.pop(rel_path, None)
            
        self._compute_aggregate_root()
        elapsed = time.perf_counter() - start_time
        return elapsed, self.get_storage_overhead()

    def get_storage_overhead(self) -> int:
        """Đo kích thước lưu trữ của manifest (bytes)."""
        data = {
            "root_hash": self.root_hash,
            "manifest": self.manifest
        }
        return len(json.dumps(data).encode("utf-8"))

    def save_to_file(self, out_path: Optional[str] = None) -> str:
        """Lưu manifest ra tệp JSON."""
        target = out_path or os.path.join(self.root_dir, "manifest_sha256.json")
        data = {
            "root_hash": self.root_hash,
            "manifest": self.manifest
        }
        with open(target, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return target
