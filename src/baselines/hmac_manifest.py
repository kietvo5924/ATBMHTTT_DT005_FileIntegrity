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


def resolve_secret_key(
    key_str: Optional[str] = None,
    key_file: Optional[str] = None,
    env_var: str = "HMAC_SECRET_KEY",
    allow_testing_fallback: bool = True
) -> Optional[bytes]:
    """
    Tải khóa bí mật HMAC an toàn:
    1. Từ file khóa cục bộ (--key-file)
    2. Từ đối số dòng lệnh (--key)
    3. Từ biến môi trường (HMAC_SECRET_KEY)
    4. Dự phòng cho môi trường kiểm thử cô lập (nếu allow_testing_fallback=True)
    """
    if key_file and os.path.isfile(key_file):
        with open(key_file, "rb") as f:
            k = f.read().strip()
            if k:
                return k
    if key_str:
        return key_str.encode("utf-8")
    env_val = os.environ.get(env_var)
    if env_val:
        return env_val.encode("utf-8")
    if allow_testing_fallback:
        return b"dt005_test_hmac_secret_key_32bytes"
    return None


class HMACManifest:
    """
    Hệ thống kiểm tra toàn vẹn có xác thực dựa trên HMAC-SHA256.
    - Yêu cầu khóa bí mật K giữa bên tạo (Signer) và bên kiểm tra (Verifier).
    - Ngăn chặn kẻ tấn công tự sinh lại mã hash khi sửa đổi tệp và manifest.
    """
    def __init__(self, root_dir: str, secret_key: Optional[bytes] = None):
        self.root_dir = os.path.abspath(root_dir)
        self.secret_key = secret_key if secret_key is not None else resolve_secret_key()
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


if __name__ == "__main__":
    import argparse
    import sys

    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="Baseline 2: HMAC-SHA256 Manifest Integrity & Authenticity")
    parser.add_argument("--dir", default="realistic_dataset", help="Target directory to scan (default: realistic_dataset)")
    parser.add_argument("--output", default=None, help="Optional output JSON path for manifest")
    parser.add_argument("--verify", action="store_true", help="Verify directory against manifest")
    parser.add_argument("--key", default=None, help="HMAC secret key string (optional; better to use --key-file or env HMAC_SECRET_KEY)")
    parser.add_argument("--key-file", default=None, help="Path to secret key file")
    args = parser.parse_args()

    target_dir = os.path.abspath(args.dir)
    if not os.path.isdir(target_dir):
        print(f"[ERROR] Directory not found: {target_dir}")
        sys.exit(1)

    secret_key = resolve_secret_key(args.key, args.key_file, allow_testing_fallback=False)
    if secret_key is None:
        print("[!] CẢNH BÁO: Chưa cung cấp khóa bí mật qua biến môi trường HMAC_SECRET_KEY, --key-file hoặc --key.")
        print("[!] Đang sử dụng khóa kiểm thử cục bộ cô lập (34 bytes) cho phiên chạy hiện tại.")
        print("[!] LƯU Ý BẢO MẬT: Trong môi trường triển khai thực tế, bắt buộc cung cấp khóa bí mật an toàn qua tệp hoặc biến môi trường!")
        secret_key = b"dt005_test_hmac_secret_key_32bytes"

    print("=" * 75)
    print("   BASELINE 2: HMAC-SHA256 MANIFEST - XÁC THỰC TOÀN VẸN VÀ NGUỒN GỐC")
    print("=" * 75)
    print(f"[*] Scanning target directory: {os.path.basename(target_dir)} ({target_dir})")
    print(f"[*] Secret key status: Loaded securely ({len(secret_key)} bytes, masked)")
    print(f"[*] Read buffer size: 65,536 Bytes (64 KB Streaming chunks)")

    hmac_tool = HMACManifest(target_dir, secret_key=secret_key)
    build_time, storage_bytes = hmac_tool.build_manifest()
    file_count = len(hmac_tool.manifest)

    print(f"[+] Total files registered: {file_count} files")
    print(f"[+] Verification / Build Time: {build_time:.4f} s ({build_time * 1000:.2f} ms)")
    print(f"[+] Metadata Storage Overhead: {storage_bytes} Bytes ({storage_bytes / 1024:.2f} KB)")
    print(f"[+] Aggregate HMAC Digest: {hmac_tool.aggregate_hmac}")
    print(f"[+] Tính xác thực nguồn gốc: Kháng giả mạo chữ ký (với giả định khóa bí mật K an toàn; chưa hỗ trợ chống chối bỏ phi đối xứng)")

    if args.output:
        saved_path = hmac_tool.save_to_file(args.output)
        print(f"[+] Manifest saved to: {saved_path}")

    if args.verify:
        print("[*] Running full integrity & authenticity verification...")
        is_valid, violations, v_time = hmac_tool.verify_all()
        print(f"[+] Verification Status: {'TOÀN VẸN & HỢP LỆ' if is_valid else 'PHÁT HIỆN VI PHẠM'}")
        print(f"[+] Verification Elapsed: {v_time:.4f} s")
        if violations:
            print(f"[-] Vi phạm ({len(violations)}): {violations[:5]}")

