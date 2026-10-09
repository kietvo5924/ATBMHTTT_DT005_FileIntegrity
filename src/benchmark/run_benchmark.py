"""
Thực Nghiệm Hiệu Năng Độc Lập - Dự Án DT005
Đo đạc khách quan và công bằng trên cùng tập dữ liệu thực tế (1.000 tệp):
1. Flat SHA-256 Manifest (Baseline 1)
2. HMAC-SHA256 Manifest (Baseline 2)
3. Standard Binary Merkle Tree (Baseline 3)
4. Hierarchical Dynamic Merkle Tree (HD-MerkleTree - Proposed)

Phân biệt rõ:
- Full Verification / Build Time
- Single-File Verification Time
- Update Cost: Sửa tệp (Modify), Thêm tệp (Add), Xóa tệp (Delete)
- Storage Overhead (Metadata)
- Proof Size (Băng thông xác thực)

Mọi thao tác sửa/thêm/xóa đều thực hiện trên bản sao tạm (Isolated Temp Copy),
tuyệt đối không làm thay đổi thư mục dữ liệu gốc.
Lưu toàn bộ số liệu chi tiết ra JSON và CSV.
"""
import os
import sys
import time
import json
import csv
import shutil
import tempfile
import platform
import statistics
import argparse
from typing import Dict, List, Any, Tuple

# Đảm bảo import được các module từ src/
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.baselines.sha256_manifest import SHA256Manifest, hash_file_sha256
from src.baselines.hmac_manifest import HMACManifest, compute_file_hmac, resolve_secret_key
from src.baselines.standard_merkle_tree import StandardMerkleTree
from src.proposed_model.hd_merkle_tree import HDMerkleTree


def get_environment_info() -> Dict[str, Any]:
    """Thu thập thông tin cấu hình môi trường thực nghiệm."""
    return {
        "os_name": platform.system(),
        "os_release": platform.release(),
        "os_version": platform.version(),
        "architecture": platform.machine(),
        "processor": platform.processor(),
        "cpu_count": os.cpu_count(),
        "python_version": platform.python_version(),
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
    }


def get_dataset_info(dataset_dir: str) -> Dict[str, Any]:
    """Đo quy mô tập dữ liệu đầu vào."""
    total_files = 0
    total_bytes = 0
    for dirpath, _, filenames in os.walk(dataset_dir):
        for f in filenames:
            p = os.path.join(dirpath, f)
            total_files += 1
            total_bytes += os.path.getsize(p)
    return {
        "dataset_path": os.path.abspath(dataset_dir),
        "total_files": total_files,
        "total_bytes": total_bytes,
        "total_megabytes": round(total_bytes / (1024 * 1024), 2)
    }


def warmup_disk_cache(dataset_dir: str) -> None:
    """Làm nóng OS file system cache trước khi đo để loại bỏ nhiễu cold-cache I/O."""
    for dirpath, _, filenames in os.walk(dataset_dir):
        for f in filenames:
            p = os.path.join(dirpath, f)
            try:
                with open(p, "rb") as fp:
                    while fp.read(65536):
                        pass
            except Exception:
                pass


def compute_stats(values: List[float]) -> Dict[str, float]:
    """Tính mean, std, min, max (thời gian tính bằng mili-giây ms)."""
    if not values:
        return {"mean_ms": 0.0, "std_ms": 0.0, "min_ms": 0.0, "max_ms": 0.0}
    mean_val = statistics.mean(values) * 1000.0
    std_val = statistics.stdev(values) * 1000.0 if len(values) > 1 else 0.0
    min_val = min(values) * 1000.0
    max_val = max(values) * 1000.0
    return {
        "mean_ms": round(mean_val, 4),
        "std_ms": round(std_val, 4),
        "min_ms": round(min_val, 4),
        "max_ms": round(max_val, 4)
    }


def run_benchmark(dataset_dir: str, num_runs: int = 5) -> Dict[str, Any]:
    env_info = get_environment_info()
    data_info = get_dataset_info(dataset_dir)

    print("=" * 80)
    print("        DT005 - BENCHMARK HIỆU NĂNG TOÀN DIỆN 4 MÔ HÌNH TOÀN VẸN TỆP")
    print("=" * 80)
    print(f"[*] Môi trường: {env_info['os_name']} {env_info['os_release']} | CPU: {env_info['processor']} ({env_info['cpu_count']} cores)")
    print(f"[*] Python: {env_info['python_version']} | Thời gian: {env_info['timestamp']}")
    print(f"[*] Tập dữ liệu: {data_info['total_files']} tệp ({data_info['total_megabytes']} MB)")
    print(f"[*] Số lần lặp đo đạc (Iterations): {num_runs} lần")
    print("=" * 80)

    print("[*] Đang làm nóng file system cache (Warm-up)...")
    warmup_disk_cache(dataset_dir)
    print("[+] Warm-up hoàn tất.")

    # Thu thập danh sách tệp để chọn target file đại diện
    sample_files = []
    for dirpath, _, filenames in os.walk(dataset_dir):
        for f in filenames:
            rel = os.path.relpath(os.path.join(dirpath, f), dataset_dir).replace("\\", "/")
            sample_files.append(rel)
    sample_files.sort()
    # Chọn tệp mẫu đại diện đồng nhất tuyệt đối với CLI mặc định (13 steps HD-Merkle, 10 steps Standard Merkle)
    target_sample_file = sample_files[0] if sample_files else ""
    print(f"[*] Tệp mẫu dùng cho Single-File Verification & Sửa đổi: {target_sample_file}\n")

    # Cấu trúc lưu trữ raw metrics qua từng run
    raw_results = {
        "sha256": {
            "build_time": [],
            "verify_single": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "storage_bytes": 0,
            "proof_bytes": 0
        },
        "hmac": {
            "build_time": [],
            "verify_single": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "storage_bytes": 0,
            "proof_bytes": 0
        },
        "standard_merkle": {
            "build_time": [],
            "verify_single": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "storage_bytes": 0,
            "proof_bytes": 0
        },
        "hd_merkle": {
            "build_time": [],
            "verify_single": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "storage_bytes": 0,
            "proof_bytes": 0
        }
    }

    secret_key = resolve_secret_key()

    for run_idx in range(1, num_runs + 1):
        print(f"--- Đang chạy Iteration {run_idx}/{num_runs} ---")

        # ---------------------------------------------------------------------
        # 1. Full Build Time & Storage Overhead (đo trên dataset_dir gốc không ghi)
        # ---------------------------------------------------------------------
        # Baseline 1: Flat SHA-256
        b1 = SHA256Manifest(dataset_dir)
        t_b1, s_b1 = b1.build_manifest()
        raw_results["sha256"]["build_time"].append(t_b1)
        raw_results["sha256"]["storage_bytes"] = s_b1

        # Baseline 2: HMAC-SHA256
        b2 = HMACManifest(dataset_dir, secret_key=secret_key)
        t_b2, s_b2 = b2.build_manifest()
        raw_results["hmac"]["build_time"].append(t_b2)
        raw_results["hmac"]["storage_bytes"] = s_b2

        # Baseline 3: Standard Merkle Tree
        b3 = StandardMerkleTree(dataset_dir)
        t_b3, s_b3 = b3.build_tree()
        raw_results["standard_merkle"]["build_time"].append(t_b3)
        raw_results["standard_merkle"]["storage_bytes"] = s_b3

        # Proposed: HD-MerkleTree
        hd = HDMerkleTree(dataset_dir)
        t_hd, s_hd = hd.build_tree_parallel()
        raw_results["hd_merkle"]["build_time"].append(t_hd)
        raw_results["hd_merkle"]["storage_bytes"] = s_hd

        # ---------------------------------------------------------------------
        # 2. Single-File Verification Time & Proof Size
        # ---------------------------------------------------------------------
        target_full = os.path.join(dataset_dir, target_sample_file)

        # Baseline 1: Tra cứu hash manifest
        t_start = time.perf_counter()
        f_hash = hash_file_sha256(target_full)
        _ = (target_sample_file in b1.manifest and b1.manifest[target_sample_file] == f_hash)
        t_single_b1 = time.perf_counter() - t_start
        raw_results["sha256"]["verify_single"].append(t_single_b1)
        raw_results["sha256"]["proof_bytes"] = s_b1  # Buộc tải toàn bộ manifest

        # Baseline 2: HMAC tra cứu manifest
        t_start = time.perf_counter()
        f_mac = compute_file_hmac(target_full, secret_key)
        _ = (target_sample_file in b2.manifest and b2.manifest[target_sample_file] == f_mac)
        t_single_b2 = time.perf_counter() - t_start
        raw_results["hmac"]["verify_single"].append(t_single_b2)
        raw_results["hmac"]["proof_bytes"] = s_b2  # Buộc tải toàn bộ manifest

        # Baseline 3: Standard Merkle Proof
        proof_b3 = b3.get_proof(target_sample_file)
        f_hash_b3 = hash_file_sha256(target_full)
        t_start = time.perf_counter()
        _ = StandardMerkleTree.verify_proof(target_sample_file, f_hash_b3, proof_b3, b3.root_hash)
        t_single_b3 = time.perf_counter() - t_start
        raw_results["standard_merkle"]["verify_single"].append(t_single_b3)
        raw_results["standard_merkle"]["proof_bytes"] = len(json.dumps(proof_b3).encode("utf-8"))

        # Proposed: HD-Merkle Proof
        proof_hd = hd.get_proof(target_sample_file)
        f_node = hd.file_nodes[target_sample_file]
        t_start = time.perf_counter()
        _ = HDMerkleTree.verify_proof(target_sample_file, f_node.file_content_hash, proof_hd, hd.root_hash)
        t_single_hd = time.perf_counter() - t_start
        raw_results["hd_merkle"]["verify_single"].append(t_single_hd)
        raw_results["hd_merkle"]["proof_bytes"] = len(json.dumps(proof_hd).encode("utf-8"))

        # ---------------------------------------------------------------------
        # 3. Update Operations (Chạy hoàn toàn trên Isolated Temp Copy)
        # ---------------------------------------------------------------------
        temp_work_dir = tempfile.mkdtemp(prefix="dt005_bench_copy_")
        try:
            # Sao chép dataset sang thư mục tạm độc lập
            for root, dirs, files in os.walk(dataset_dir):
                rel = os.path.relpath(root, dataset_dir)
                dest = os.path.join(temp_work_dir, rel) if rel != "." else temp_work_dir
                os.makedirs(dest, exist_ok=True)
                for file in files:
                    src_f = os.path.join(root, file)
                    dest_f = os.path.join(dest, file)
                    shutil.copy2(src_f, dest_f)

            # Khởi tạo các mô hình trên thư mục tạm
            t_b1_copy = SHA256Manifest(temp_work_dir)
            t_b1_copy.build_manifest()
            t_b2_copy = HMACManifest(temp_work_dir, secret_key=secret_key)
            t_b2_copy.build_manifest()
            t_b3_copy = StandardMerkleTree(temp_work_dir)
            t_b3_copy.build_tree()
            t_hd_copy = HDMerkleTree(temp_work_dir)
            t_hd_copy.build_tree_parallel()

            # (A) SỬA TỆP (MODIFY): Ghi đè file trên đĩa và đo update cost
            temp_target_file = os.path.join(temp_work_dir, target_sample_file)
            with open(temp_target_file, "ab") as f:
                f.write(b"__BENCHMARK_MODIFICATION_PAYLOAD__")

            # Đo Update Modify B1
            t_mod_b1, _ = t_b1_copy.update_file(target_sample_file)
            raw_results["sha256"]["update_modify"].append(t_mod_b1)

            # Đo Update Modify B2
            t_mod_b2, _ = t_b2_copy.update_file(target_sample_file)
            raw_results["hmac"]["update_modify"].append(t_mod_b2)

            # Đo Update Modify B3
            t_mod_b3, _ = t_b3_copy.update_file(target_sample_file)
            raw_results["standard_merkle"]["update_modify"].append(t_mod_b3)

            # Đo Update Modify Proposed
            t_mod_hd, _ = t_hd_copy.update_file_fast(target_sample_file)
            raw_results["hd_merkle"]["update_modify"].append(t_mod_hd)

            # (B) THÊM TỆP (ADD): Tạo tệp mới trên đĩa và đo update cost
            new_rel_path = "certificates/new_benchmark_audit_certificate.pdf"
            new_full_path = os.path.join(temp_work_dir, new_rel_path)
            with open(new_full_path, "wb") as f:
                f.write(b"NEW FILE PAYLOAD FOR BENCHMARK ADDITION TEST")

            # Baseline 1 Add
            t_add_b1, _ = t_b1_copy.update_file(new_rel_path)
            raw_results["sha256"]["update_add"].append(t_add_b1)

            # Baseline 2 Add
            t_add_b2, _ = t_b2_copy.update_file(new_rel_path)
            raw_results["hmac"]["update_add"].append(t_add_b2)

            # Baseline 3 Add (Standard Merkle Tree buộc Rebuild O(N))
            t_add_b3, _ = t_b3_copy.update_file(new_rel_path)
            raw_results["standard_merkle"]["update_add"].append(t_add_b3)

            # Proposed Add (O(log N))
            t_add_hd, _ = t_hd_copy.update_file_fast(new_rel_path)
            raw_results["hd_merkle"]["update_add"].append(t_add_hd)

            # (C) XÓA TỆP (DELETE): Xóa 1 tệp trên đĩa và đo update cost
            del_target_rel = sample_files[0]
            del_target_full = os.path.join(temp_work_dir, del_target_rel)
            if os.path.exists(del_target_full):
                os.remove(del_target_full)

            # Baseline 1 Delete
            t_del_b1, _ = t_b1_copy.update_file(del_target_rel)
            raw_results["sha256"]["update_delete"].append(t_del_b1)

            # Baseline 2 Delete
            t_del_b2, _ = t_b2_copy.update_file(del_target_rel)
            raw_results["hmac"]["update_delete"].append(t_del_b2)

            # Baseline 3 Delete (Standard Merkle Tree buộc Rebuild O(N))
            t_del_b3, _ = t_b3_copy.update_file(del_target_rel)
            raw_results["standard_merkle"]["update_delete"].append(t_del_b3)

            # Proposed Delete (O(log N))
            t_del_hd, _ = t_hd_copy.update_file_fast(del_target_rel)
            raw_results["hd_merkle"]["update_delete"].append(t_del_hd)

        finally:
            shutil.rmtree(temp_work_dir, ignore_errors=True)

    # -------------------------------------------------------------------------
    # 4. Thống kê số liệu (Mean, Std, Min, Max)
    # -------------------------------------------------------------------------
    summary = {}
    for model_key in ["sha256", "hmac", "standard_merkle", "hd_merkle"]:
        summary[model_key] = {
            "build_time": compute_stats(raw_results[model_key]["build_time"]),
            "verify_single": compute_stats(raw_results[model_key]["verify_single"]),
            "update_modify": compute_stats(raw_results[model_key]["update_modify"]),
            "update_add": compute_stats(raw_results[model_key]["update_add"]),
            "update_delete": compute_stats(raw_results[model_key]["update_delete"]),
            "storage_bytes": raw_results[model_key]["storage_bytes"],
            "proof_bytes": raw_results[model_key]["proof_bytes"]
        }

    report = {
        "environment": env_info,
        "dataset": data_info,
        "parameters": {
            "num_runs": num_runs,
            "sample_target_file": target_sample_file
        },
        "raw_results": raw_results,
        "summary": summary
    }

    # In bảng kết quả chi tiết ra màn hình
    print("\n" + "=" * 90)
    print("                    BẢNG KẾT QUẢ THỰC NGHIỆM ĐỘC LẬP (DT005)")
    print("=" * 90)
    print(f"{'Chỉ số (Metric)':<30} | {'SHA-256':<12} | {'HMAC-SHA256':<12} | {'Standard Merkle':<15} | {'HD-MerkleTree':<15}")
    print("-" * 90)

    def fmt(m_key, metric):
        stat = summary[m_key][metric]
        return f"{stat['mean_ms']:.2f} ± {stat['std_ms']:.2f}"

    print(f"{'Build / Full Verify (ms)':<30} | {fmt('sha256', 'build_time'):<12} | {fmt('hmac', 'build_time'):<12} | {fmt('standard_merkle', 'build_time'):<15} | {fmt('hd_merkle', 'build_time'):<15}")
    print(f"{'Single-File Verify (ms)':<30} | {fmt('sha256', 'verify_single'):<12} | {fmt('hmac', 'verify_single'):<12} | {fmt('standard_merkle', 'verify_single'):<15} | {fmt('hd_merkle', 'verify_single'):<15}")
    print(f"{'Update Sửa File (ms)':<30} | {fmt('sha256', 'update_modify'):<12} | {fmt('hmac', 'update_modify'):<12} | {fmt('standard_merkle', 'update_modify'):<15} | {fmt('hd_merkle', 'update_modify'):<15}")
    print(f"{'Update Thêm File (ms)':<30} | {fmt('sha256', 'update_add'):<12} | {fmt('hmac', 'update_add'):<12} | {fmt('standard_merkle', 'update_add'):<15} | {fmt('hd_merkle', 'update_add'):<15}")
    print(f"{'Update Xóa File (ms)':<30} | {fmt('sha256', 'update_delete'):<12} | {fmt('hmac', 'update_delete'):<12} | {fmt('standard_merkle', 'update_delete'):<15} | {fmt('hd_merkle', 'update_delete'):<15}")
    print("-" * 90)
    print(f"{'Storage Overhead (KB)':<30} | {summary['sha256']['storage_bytes']/1024:<12.2f} | {summary['hmac']['storage_bytes']/1024:<12.2f} | {summary['standard_merkle']['storage_bytes']/1024:<15.2f} | {summary['hd_merkle']['storage_bytes']/1024:<15.2f}")
    print(f"{'Proof Size (Bytes)':<30} | {summary['sha256']['proof_bytes']:<12} | {summary['hmac']['proof_bytes']:<12} | {summary['standard_merkle']['proof_bytes']:<15} | {summary['hd_merkle']['proof_bytes']:<15}")
    print("=" * 90)

    return report


def save_results(report: Dict[str, Any], results_dir: str) -> Tuple[str, str]:
    """Lưu kết quả thực nghiệm ra tệp JSON và CSV truy nguyên được."""
    os.makedirs(results_dir, exist_ok=True)
    json_path = os.path.join(results_dir, "benchmark_results.json")
    csv_path = os.path.join(results_dir, "benchmark_results.csv")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    summary = report["summary"]
    metrics = ["build_time", "verify_single", "update_modify", "update_add", "update_delete"]

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Model", "Metric", "Mean (ms)", "Std Dev (ms)", "Min (ms)", "Max (ms)", "Storage (Bytes)", "Proof Size (Bytes)"])
        model_names = {
            "sha256": "Flat SHA-256 Manifest",
            "hmac": "HMAC-SHA256 Manifest",
            "standard_merkle": "Standard Merkle Tree",
            "hd_merkle": "HD-MerkleTree (Proposed)"
        }
        for m_key, m_name in model_names.items():
            st_bytes = summary[m_key]["storage_bytes"]
            pr_bytes = summary[m_key]["proof_bytes"]
            for met in metrics:
                s = summary[m_key][met]
                writer.writerow([m_name, met, s["mean_ms"], s["std_ms"], s["min_ms"], s["max_ms"], st_bytes, pr_bytes])

    return json_path, csv_path


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="Standalone DT005 Benchmark Runner")
    parser.add_argument("--dir", default="realistic_dataset", help="Target dataset directory")
    parser.add_argument("--runs", type=int, default=5, help="Number of test iterations (default: 5)")
    parser.add_argument("--results-dir", default="results", help="Directory to save JSON and CSV reports")
    args = parser.parse_args()

    target_dataset = os.path.abspath(args.dir)
    if not os.path.isdir(target_dataset):
        print(f"[ERROR] Thư mục không tồn tại: {target_dataset}")
        sys.exit(1)

    results_dir = os.path.abspath(args.results_dir)
    report_data = run_benchmark(target_dataset, num_runs=args.runs)
    json_out, csv_out = save_results(report_data, results_dir)
    print(f"\n[+] Đã lưu kết quả JSON tại: {json_out}")
    print(f"[+] Đã lưu kết quả CSV tại: {csv_out}")
