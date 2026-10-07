"""
Baseline 2: HMAC-SHA256 Manifest
Kiểm tra tính toàn vẹn và xác thực nguồn gốc (Integrity + Authenticity)
sử dụng mã xác thực thông điệp dựa trên băm HMAC-SHA256 với khóa bí mật K.
"""
import os
import hmac
import json
import time
import hashlib
from typing import Dict, Tuple, List, Optional


def compute_file_hmac(filepath: str, secret_key: bytes, chunk_size: int = 65536) -> str:
    """Tính mã HMAC-SHA256 của tệp tin với khóa bí mật K."""
    mac = hmac.new(secret_key, digestmod=hashlib.sha256)
    with open(filepath, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            mac.update(chunk)
    return mac.hexdigest()


class HMACManifest:
    """
    Hệ thống kiểm tra toàn vẹn có xác thực dựa trên HMAC-SHA256.
    - Yêu cầu khóa bí mật K giữa bên tạo (Signer) và bên kiểm tra (Verifier).
    - Ngăn chặn kẻ tấn công tự sinh lại mã hash khi sửa đổi tệp và manifest.
    """
    def __init__(self, root_dir: str, secret_key: bytes = b"default_secret_key_atbmhttt_2026"):
        self.root_dir = os.path.abspath(root_dir)
        self.secret_key = secret_key
        self.manifest: Dict[str, str] = {}
        self.aggregate_hmac: str = ""

    def build_manifest(self) -> Tuple[float, int]:
        """
        Quét thư mục và tính toán HMAC cho từng tệp cùng HMAC tổng hợp.
        Trả về: (thời gian thực thi (s), dung lượng lưu trữ manifest)
        """
        start_time = time.perf_counter()
        manifest = {}
        
        for dirpath, _, filenames in os.walk(self.root_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, self.root_dir).replace("\\", "/")
                if rel_path in ("manifest_sha256.json", "manifest_hmac.json"):
                    continue
                file_mac = compute_file_hmac(full_path, self.secret_key)
                manifest[rel_path] = file_mac
                
        self.manifest = dict(sorted(manifest.items()))
        self._compute_aggregate_hmac()
        elapsed = time.perf_counter() - start_time
        return elapsed, self.get_storage_overhead()

    def _compute_aggregate_hmac(self) -> None:
        """Tính HMAC tổng hợp cho toàn bộ danh sách tệp."""
        mac = hmac.new(self.secret_key, digestmod=hashlib.sha256)
        for path in sorted(self.manifest.keys()):
            mac.update(path.encode("utf-8"))
            mac.update(bytes.fromhex(self.manifest[path]))
        self.aggregate_hmac = mac.hexdigest()

    def verify_all(self) -> Tuple[bool, List[str], float]:
        """
        Xác thực toàn bộ thư mục và kiểm tra tính hợp lệ của chữ ký HMAC.
        """
        start_time = time.perf_counter()
        violations = []
        
        # 1. Kiểm tra tính toàn vẹn của tệp và đối chiếu từng file HMAC
        current_files = set()
        for dirpath, _, filenames in os.walk(self.root_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, self.root_dir).replace("\\", "/")
                if rel_path in ("manifest_sha256.json", "manifest_hmac.json"):
                    continue
                current_files.add(rel_path)
                
                if rel_path not in self.manifest:
                    violations.append(f"UNAUTHORIZED_FILE: {rel_path}")
                else:
                    curr_mac = compute_file_hmac(full_path, self.secret_key)
                    if not hmac.compare_digest(curr_mac, self.manifest[rel_path]):
                        violations.append(f"AUTHENTICATION_FAILED: {rel_path}")

        for expected_path in self.manifest:
            if expected_path not in current_files:
                violations.append(f"MISSING_FILE: {expected_path}")

        # 2. Kiểm tra Aggregate HMAC
        computed_agg = hmac.new(self.secret_key, digestmod=hashlib.sha256)
        for path in sorted(self.manifest.keys()):
            computed_agg.update(path.encode("utf-8"))
            computed_agg.update(bytes.fromhex(self.manifest[path]))
        
        if not hmac.compare_digest(computed_agg.hexdigest(), self.aggregate_hmac):
            violations.append("AGGREGATE_HMAC_MISMATCH")

        elapsed = time.perf_counter() - start_time
        is_valid = len(violations) == 0
        return is_valid, violations, elapsed

    def verify_single_file(self, rel_path: str) -> Tuple[bool, float]:
        """Xác thực một file đơn lẻ với khóa bí mật."""
        rel_path = rel_path.replace("\\", "/")
        if rel_path not in self.manifest:
            return False, 0.0
            
        full_path = os.path.join(self.root_dir, rel_path)
        if not os.path.exists(full_path):
            return False, 0.0
            
        start = time.perf_counter()
        curr_mac = compute_file_hmac(full_path, self.secret_key)
        is_valid = hmac.compare_digest(curr_mac, self.manifest[rel_path])
        elapsed = time.perf_counter() - start
        return is_valid, elapsed

    def update_file(self, rel_path: str) -> Tuple[float, int]:
        """
        Cập nhật HMAC khi 1 file thay đổi.
        Phải tính lại HMAC của file đó + tính lại Aggregate HMAC toàn thư mục O(N).
        """
        rel_path = rel_path.replace("\\", "/")
        full_path = os.path.join(self.root_dir, rel_path)
        
        start_time = time.perf_counter()
        if os.path.exists(full_path):
            new_mac = compute_file_hmac(full_path, self.secret_key)
            self.manifest[rel_path] = new_mac
        else:
            self.manifest.pop(rel_path, None)
            
        self._compute_aggregate_hmac()
        elapsed = time.perf_counter() - start_time
        return elapsed, self.get_storage_overhead()

    def get_storage_overhead(self) -> int:
        """Kích thước siêu dữ liệu manifest (bytes)."""
        data = {
            "aggregate_hmac": self.aggregate_hmac,
            "manifest": self.manifest
        }
        return len(json.dumps(data).encode("utf-8"))

    def save_to_file(self, out_path: Optional[str] = None) -> str:
        """Lưu cấu hình HMAC manifest ra tệp JSON."""
        target = out_path or os.path.join(self.root_dir, "manifest_hmac.json")
        data = {
            "aggregate_hmac": self.aggregate_hmac,
            "manifest": self.manifest
        }
        with open(target, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return target
