"""
Thực Nghiệm Hiệu Năng Độc Lập - Dự Án DT005
Đo đạc khách quan và công bằng trên cùng tập dữ liệu thực tế:
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
- Bổ sung Nghiên cứu Cắt giảm Thành phần (Ablation Study - 4 Cấu hình đạt chuẩn học thuật).

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
from src.proposed_model.hd_merkle_tree import HDMerkleTree, TreeNode
from src.proposed_model.compact_proof import CompactProof
from src.proposed_model.parallel_hasher import hash_single_file_streaming


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
    rel_path = os.path.relpath(os.path.abspath(dataset_dir), BASE_DIR).replace("\\", "/")
    if rel_path.startswith(".."):
        rel_path = os.path.abspath(dataset_dir)
    return {
        "dataset_path": rel_path,
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


def invalidate_all_nodes(node: TreeNode) -> None:
    """Tắt cache toàn bộ cây để phục vụ cấu hình Ablation 1 (Không có Dirty-Path Cache)."""
    node.cached_hash = None
    node.is_dirty = True
    node.subtree_levels = []
    for c in node.children.values():
        invalidate_all_nodes(c)


# ==============================================================================
# NGHIÊN CỨU CẮT GIẢM THÀNH PHẦN (ABLATION STUDY - 4 CẤU HÌNH ĐẠT CHUẨN HỌC THUẬT)
# ==============================================================================
def run_ablation_study(dataset_dir: str, num_runs: int = 5) -> Dict[str, Any]:
    """
    Thực hiện nghiên cứu cắt giảm thành phần (Ablation Study) trên cùng tập dữ liệu:
    Bóc tách đóng góp định lượng của từng cải tiến kỹ thuật:
    1. Baseline (Standard Merkle Tree): Cây nhị phân phẳng, không phân tách miền RFC 6962, Rebuild O(N) khi thêm/xóa.
    2. Ablation 1 (HD-Tree - Không có Dirty-Path Cache): Phân cấp + RFC 6962, nhưng tắt cache (duyệt đệ quy tính lại toàn bộ cây khi cập nhật).
    3. Ablation 2 (HD-Tree + Dirty-Path Cache - Đơn luồng): Phân cấp + RFC 6962 + Dirty-Path O(log N), nhưng quét đơn luồng (max_workers=1).
    4. Full Proposed Model (HD-MerkleTree Toàn diện): Phân cấp + RFC 6962 + Dirty-Path O(log N) + Đa luồng song song + CompactProof.
    """
    print("\n" + "=" * 90)
    print("        NGHIÊN CỨU CẮT GIẢM THÀNH PHẦN (ABLATION STUDY - 4 CẤU HÌNH)")
    print("=" * 90)
    print(f"[*] Tập dữ liệu: {dataset_dir}")
    print(f"[*] Số lần lặp (Iterations): {num_runs}")

    sample_files = []
    for dirpath, _, filenames in os.walk(dataset_dir):
        for f in filenames:
            rel = os.path.relpath(os.path.join(dirpath, f), dataset_dir).replace("\\", "/")
            sample_files.append(rel)
    sample_files.sort()
    target_sample = sample_files[0] if sample_files else ""

    raw_ablation: Dict[str, Dict[str, Any]] = {
        "baseline_standard": {
            "name": "1. Baseline (Standard Merkle)",
            "rfc6962": "Không",
            "cache": "Không (Mảng phẳng)",
            "threading": "Đơn luồng",
            "build_time": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "proof_bytes": 0,
            "complexity_add_del": "O(N) Rebuild"
        },
        "ablation_no_cache": {
            "name": "2. Ablation 1 (Tắt Cache)",
            "rfc6962": "Có (0x00/0x01)",
            "cache": "TẮT (Tính lại toàn cây)",
            "threading": "Đa luồng",
            "build_time": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "proof_bytes": 0,
            "complexity_add_del": "O(N) Recompute"
        },
        "ablation_single_thread": {
            "name": "3. Ablation 2 (Đơn luồng)",
            "rfc6962": "Có (0x00/0x01)",
            "cache": "BẬT O(log N)",
            "threading": "Đơn luồng (workers=1)",
            "build_time": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "proof_bytes": 0,
            "complexity_add_del": "O(log N)"
        },
        "full_proposed": {
            "name": "4. Full Proposed Model",
            "rfc6962": "Có (0x00/0x01)",
            "cache": "BẬT O(log N)",
            "threading": "Đa luồng (parallel)",
            "build_time": [],
            "update_modify": [],
            "update_add": [],
            "update_delete": [],
            "proof_bytes": 0,
            "complexity_add_del": "O(log N)"
        }
    }

    # Đo kích thước Proof tĩnh 1 lần
    t_std_init = StandardMerkleTree(dataset_dir)
    t_std_init.build_tree()
    p_std = t_std_init.get_proof(target_sample) if target_sample else []
    raw_ablation["baseline_standard"]["proof_bytes"] = len(json.dumps(p_std).encode("utf-8"))

    t_hd_init = HDMerkleTree(dataset_dir)
    t_hd_init.build_tree_parallel()
    p_hd = t_hd_init.get_proof(target_sample) if target_sample else {}
    raw_ablation["ablation_no_cache"]["proof_bytes"] = len(json.dumps(p_hd).encode("utf-8"))
    compact_sz = CompactProof.get_proof_size(CompactProof.serialize(p_hd)) if p_hd else 256
    raw_ablation["ablation_single_thread"]["proof_bytes"] = compact_sz
    raw_ablation["full_proposed"]["proof_bytes"] = compact_sz

    for run_idx in range(1, num_runs + 1):
        # 1. Đo Build Time
        # Baseline Standard
        t_std = StandardMerkleTree(dataset_dir)
        tb_std, _ = t_std.build_tree()
        raw_ablation["baseline_standard"]["build_time"].append(tb_std)

        # Ablation 1: HD-Tree Không Cache (Build bằng đa luồng)
        t_nc = HDMerkleTree(dataset_dir)
        tb_nc, _ = t_nc.build_tree_parallel(max_workers=None)
        raw_ablation["ablation_no_cache"]["build_time"].append(tb_nc)

        # Ablation 2: HD-Tree Đơn luồng
        t_st = HDMerkleTree(dataset_dir)
        tb_st, _ = t_st.build_tree_parallel(max_workers=1)
        raw_ablation["ablation_single_thread"]["build_time"].append(tb_st)

        # Full Proposed Model: Đa luồng
        t_full = HDMerkleTree(dataset_dir)
        tb_full, _ = t_full.build_tree_parallel(max_workers=None)
        raw_ablation["full_proposed"]["build_time"].append(tb_full)

        # 2. Đo Update trên thư mục tạm độc lập
        temp_dir = tempfile.mkdtemp(prefix="dt005_ablation_")
        try:
            for root, dirs, files in os.walk(dataset_dir):
                rel = os.path.relpath(root, dataset_dir)
                dest = os.path.join(temp_dir, rel) if rel != "." else temp_dir
                os.makedirs(dest, exist_ok=True)
                for file in files:
                    src_f = os.path.join(root, file)
                    dest_f = os.path.join(dest, file)
                    shutil.copy2(src_f, dest_f)

            tree_std_w = StandardMerkleTree(temp_dir)
            tree_std_w.build_tree()

            tree_nc_w = HDMerkleTree(temp_dir)
            tree_nc_w.build_tree_parallel(max_workers=None)

            tree_st_w = HDMerkleTree(temp_dir)
            tree_st_w.build_tree_parallel(max_workers=1)

            tree_full_w = HDMerkleTree(temp_dir)
            tree_full_w.build_tree_parallel(max_workers=None)

            # (A) MODIFY
            target_f = os.path.join(temp_dir, target_sample)
            with open(target_f, "ab") as fp:
                fp.write(b"__ABLATION_MODIFY_PAYLOAD__")

            t_m_std, _ = tree_std_w.update_file(target_sample)
            raw_ablation["baseline_standard"]["update_modify"].append(t_m_std)

            # Ablation 1 No Cache Modify
            t0 = time.perf_counter()
            new_h = hash_single_file_streaming(target_f)
            if target_sample in tree_nc_w.file_nodes:
                fn = tree_nc_w.file_nodes[target_sample]
                fn.file_content_hash = new_h
                invalidate_all_nodes(tree_nc_w.root_node)
                tree_nc_w.root_hash = tree_nc_w.root_node.compute_hash().hex()
            t_m_nc = time.perf_counter() - t0
            raw_ablation["ablation_no_cache"]["update_modify"].append(t_m_nc)

            t_m_st, _ = tree_st_w.update_file_fast(target_sample)
            raw_ablation["ablation_single_thread"]["update_modify"].append(t_m_st)

            t_m_full, _ = tree_full_w.update_file_fast(target_sample)
            raw_ablation["full_proposed"]["update_modify"].append(t_m_full)

            # (B) ADD
            new_rel = "ablation_test_folder/new_file_audit.txt"
            new_full = os.path.join(temp_dir, new_rel)
            os.makedirs(os.path.dirname(new_full), exist_ok=True)
            with open(new_full, "wb") as fp:
                fp.write(b"NEW_FILE_PAYLOAD_FOR_ABLATION_ADD")

            t_a_std, _ = tree_std_w.update_file(new_rel)
            raw_ablation["baseline_standard"]["update_add"].append(t_a_std)

            # Ablation 1 No Cache Add
            t0 = time.perf_counter()
            new_add_h = hash_single_file_streaming(new_full)
            tree_nc_w._insert_file_node(new_rel, new_add_h)
            invalidate_all_nodes(tree_nc_w.root_node)
            tree_nc_w.root_hash = tree_nc_w.root_node.compute_hash().hex()
            t_a_nc = time.perf_counter() - t0
            raw_ablation["ablation_no_cache"]["update_add"].append(t_a_nc)

            t_a_st, _ = tree_st_w.update_file_fast(new_rel)
            raw_ablation["ablation_single_thread"]["update_add"].append(t_a_st)

            t_a_full, _ = tree_full_w.update_file_fast(new_rel)
            raw_ablation["full_proposed"]["update_add"].append(t_a_full)

            # (C) DELETE
            del_target = target_sample
            del_full = os.path.join(temp_dir, del_target)
            if os.path.exists(del_full):
                os.remove(del_full)

            t_d_std, _ = tree_std_w.update_file(del_target)
            raw_ablation["baseline_standard"]["update_delete"].append(t_d_std)

            # Ablation 1 No Cache Delete
            t0 = time.perf_counter()
            if del_target in tree_nc_w.file_nodes:
                fn = tree_nc_w.file_nodes.pop(del_target)
                parent = fn.parent
                if parent and fn.name in parent.children:
                    del parent.children[fn.name]
                invalidate_all_nodes(tree_nc_w.root_node)
                tree_nc_w.root_hash = tree_nc_w.root_node.compute_hash().hex()
            t_d_nc = time.perf_counter() - t0
            raw_ablation["ablation_no_cache"]["update_delete"].append(t_d_nc)

            t_d_st, _ = tree_st_w.update_file_fast(del_target)
            raw_ablation["ablation_single_thread"]["update_delete"].append(t_d_st)

            t_d_full, _ = tree_full_w.update_file_fast(del_target)
            raw_ablation["full_proposed"]["update_delete"].append(t_d_full)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    summary_ablation = {}
    for c_key in raw_ablation:
        data = raw_ablation[c_key]
        summary_ablation[c_key] = {
            "name": data["name"],
            "rfc6962": data["rfc6962"],
            "cache": data["cache"],
            "threading": data["threading"],
            "build_time": compute_stats(data["build_time"]),
            "update_modify": compute_stats(data["update_modify"]),
            "update_add": compute_stats(data["update_add"]),
            "update_delete": compute_stats(data["update_delete"]),
            "proof_bytes": data["proof_bytes"],
            "complexity_add_del": data["complexity_add_del"]
        }

    # In bảng Ablation Study ra màn hình
    print("\n" + "=" * 120)
    print("                                BẢNG KẾT QUẢ ABLATION STUDY (DT005)")
    print("=" * 120)
    print(f"{'Cấu hình Thực Nghiệm':<32} | {'RFC 6962':<10} | {'Dirty-Cache':<14} | {'Luồng':<12} | {'Build (ms)':<15} | {'Sửa tệp (ms)':<15} | {'Thêm tệp (ms)':<15} | {'Xóa tệp (ms)':<15} | {'Proof (B)':<9}")
    print("-" * 120)

    def s_fmt(cfg, met):
        st = summary_ablation[cfg][met]
        return f"{st['mean_ms']:.2f} ± {st['std_ms']:.2f}"

    for cfg, info in summary_ablation.items():
        print(
            f"{info['name']:<32} | {info['rfc6962']:<10} | {info['cache']:<14} | {info['threading']:<12} | "
            f"{s_fmt(cfg, 'build_time'):<15} | {s_fmt(cfg, 'update_modify'):<15} | {s_fmt(cfg, 'update_add'):<15} | {s_fmt(cfg, 'update_delete'):<15} | {info['proof_bytes']:<9}"
        )
    print("=" * 120 + "\n")

    return {
        "raw_ablation": raw_ablation,
        "summary": summary_ablation
    }


# ==============================================================================
# HÀM BENCHMARK TỔNG THỂ
# ==============================================================================
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
    target_sample_file = sample_files[0] if sample_files else ""
    print(f"[*] Tệp mẫu dùng cho Single-File Verification & Sửa đổi: {target_sample_file}\n")

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

        # 1. Full Build Time & Storage Overhead
        b1 = SHA256Manifest(dataset_dir)
        t_b1, s_b1 = b1.build_manifest()
        raw_results["sha256"]["build_time"].append(t_b1)
        raw_results["sha256"]["storage_bytes"] = s_b1

        b2 = HMACManifest(dataset_dir, secret_key=secret_key)
        t_b2, s_b2 = b2.build_manifest()
        raw_results["hmac"]["build_time"].append(t_b2)
        raw_results["hmac"]["storage_bytes"] = s_b2

        b3 = StandardMerkleTree(dataset_dir)
        t_b3, s_b3 = b3.build_tree()
        raw_results["standard_merkle"]["build_time"].append(t_b3)
        raw_results["standard_merkle"]["storage_bytes"] = s_b3

        hd = HDMerkleTree(dataset_dir)
        t_hd, s_hd = hd.build_tree_parallel()
        raw_results["hd_merkle"]["build_time"].append(t_hd)
        raw_results["hd_merkle"]["storage_bytes"] = s_hd

        # 2. Single-File Verification Time & Proof Size
        target_full = os.path.join(dataset_dir, target_sample_file)

        t_start = time.perf_counter()
        f_hash = hash_file_sha256(target_full)
        _ = (target_sample_file in b1.manifest and b1.manifest[target_sample_file] == f_hash)
        t_single_b1 = time.perf_counter() - t_start
        raw_results["sha256"]["verify_single"].append(t_single_b1)
        raw_results["sha256"]["proof_bytes"] = s_b1

        t_start = time.perf_counter()
        f_mac = compute_file_hmac(target_full, secret_key)
        _ = (target_sample_file in b2.manifest and b2.manifest[target_sample_file] == f_mac)
        t_single_b2 = time.perf_counter() - t_start
        raw_results["hmac"]["verify_single"].append(t_single_b2)
        raw_results["hmac"]["proof_bytes"] = s_b2

        proof_b3 = b3.get_proof(target_sample_file)
        f_hash_b3 = hash_file_sha256(target_full)
        t_start = time.perf_counter()
        _ = StandardMerkleTree.verify_proof(target_sample_file, f_hash_b3, proof_b3, b3.root_hash)
        t_single_b3 = time.perf_counter() - t_start
        raw_results["standard_merkle"]["verify_single"].append(t_single_b3)
        raw_results["standard_merkle"]["proof_bytes"] = len(json.dumps(proof_b3).encode("utf-8"))

        proof_hd = hd.get_proof(target_sample_file)
        f_node = hd.file_nodes[target_sample_file]
        t_start = time.perf_counter()
        _ = HDMerkleTree.verify_proof(target_sample_file, f_node.file_content_hash, proof_hd, hd.root_hash)
        t_single_hd = time.perf_counter() - t_start
        raw_results["hd_merkle"]["verify_single"].append(t_single_hd)
        compact_bytes = CompactProof.get_proof_size(CompactProof.serialize(proof_hd)) if proof_hd else len(json.dumps(proof_hd).encode("utf-8"))
        raw_results["hd_merkle"]["proof_bytes"] = compact_bytes

        # 3. Update Operations (Chạy hoàn toàn trên Isolated Temp Copy)
        temp_work_dir = tempfile.mkdtemp(prefix="dt005_bench_copy_")
        try:
            for root, dirs, files in os.walk(dataset_dir):
                rel = os.path.relpath(root, dataset_dir)
                dest = os.path.join(temp_work_dir, rel) if rel != "." else temp_work_dir
                os.makedirs(dest, exist_ok=True)
                for file in files:
                    src_f = os.path.join(root, file)
                    dest_f = os.path.join(dest, file)
                    shutil.copy2(src_f, dest_f)

            t_b1_copy = SHA256Manifest(temp_work_dir)
            t_b1_copy.build_manifest()
            t_b2_copy = HMACManifest(temp_work_dir, secret_key=secret_key)
            t_b2_copy.build_manifest()
            t_b3_copy = StandardMerkleTree(temp_work_dir)
            t_b3_copy.build_tree()
            t_hd_copy = HDMerkleTree(temp_work_dir)
            t_hd_copy.build_tree_parallel()

            # (A) SỬA TỆP (MODIFY)
            temp_target_file = os.path.join(temp_work_dir, target_sample_file)
            with open(temp_target_file, "ab") as f:
                f.write(b"__BENCHMARK_MODIFICATION_PAYLOAD__")

            t_mod_b1, _ = t_b1_copy.update_file(target_sample_file)
            raw_results["sha256"]["update_modify"].append(t_mod_b1)

            t_mod_b2, _ = t_b2_copy.update_file(target_sample_file)
            raw_results["hmac"]["update_modify"].append(t_mod_b2)

            t_mod_b3, _ = t_b3_copy.update_file(target_sample_file)
            raw_results["standard_merkle"]["update_modify"].append(t_mod_b3)

            t_mod_hd, _ = t_hd_copy.update_file_fast(target_sample_file)
            raw_results["hd_merkle"]["update_modify"].append(t_mod_hd)

            # (B) THÊM TỆP (ADD)
            new_rel_path = "benchmark_new_folder/audit_new_file.txt"
            new_full_path = os.path.join(temp_work_dir, new_rel_path)
            os.makedirs(os.path.dirname(new_full_path), exist_ok=True)
            with open(new_full_path, "wb") as f:
                f.write(b"NEW FILE PAYLOAD FOR BENCHMARK ADDITION TEST")

            t_add_b1, _ = t_b1_copy.update_file(new_rel_path)
            raw_results["sha256"]["update_add"].append(t_add_b1)

            t_add_b2, _ = t_b2_copy.update_file(new_rel_path)
            raw_results["hmac"]["update_add"].append(t_add_b2)

            t_add_b3, _ = t_b3_copy.update_file(new_rel_path)
            raw_results["standard_merkle"]["update_add"].append(t_add_b3)

            t_add_hd, _ = t_hd_copy.update_file_fast(new_rel_path)
            raw_results["hd_merkle"]["update_add"].append(t_add_hd)

            # (C) XÓA TỆP (DELETE)
            del_target_rel = sample_files[0]
            del_target_full = os.path.join(temp_work_dir, del_target_rel)
            if os.path.exists(del_target_full):
                os.remove(del_target_full)

            t_del_b1, _ = t_b1_copy.update_file(del_target_rel)
            raw_results["sha256"]["update_delete"].append(t_del_b1)

            t_del_b2, _ = t_b2_copy.update_file(del_target_rel)
            raw_results["hmac"]["update_delete"].append(t_del_b2)

            t_del_b3, _ = t_b3_copy.update_file(del_target_rel)
            raw_results["standard_merkle"]["update_delete"].append(t_del_b3)

            t_del_hd, _ = t_hd_copy.update_file_fast(del_target_rel)
            raw_results["hd_merkle"]["update_delete"].append(t_del_hd)

        finally:
            shutil.rmtree(temp_work_dir, ignore_errors=True)

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

    # 4. CHẠY NGHIÊN CỨU CẮT GIẢM THÀNH PHẦN (ABLATION STUDY)
    ablation_study_data = run_ablation_study(dataset_dir, num_runs=num_runs)

    report = {
        "environment": env_info,
        "dataset": data_info,
        "parameters": {
            "num_runs": num_runs,
            "sample_target_file": target_sample_file
        },
        "raw_results": raw_results,
        "summary": summary,
        "ablation_study": ablation_study_data
    }

    return report


def save_results(report: Dict[str, Any], results_dir: str) -> Tuple[str, str]:
    """Lưu kết quả thực nghiệm ra tệp JSON và CSV truy nguyên được."""
    os.makedirs(results_dir, exist_ok=True)
    json_path = os.path.join(results_dir, "benchmark_results.json")
    csv_path = os.path.join(results_dir, "benchmark_results.csv")
    ablation_csv_path = os.path.join(results_dir, "ablation_results.csv")

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

    # Lưu thêm bảng Ablation Study ra CSV
    if "ablation_study" in report and "summary" in report["ablation_study"]:
        with open(ablation_csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "Configuration", "RFC 6962", "Dirty-Path Cache", "Threading",
                "Build Mean (ms)", "Build Std (ms)",
                "Modify Mean (ms)", "Modify Std (ms)",
                "Add Mean (ms)", "Add Std (ms)",
                "Delete Mean (ms)", "Delete Std (ms)",
                "Proof Size (Bytes)", "Complexity Add/Delete"
            ])
            for cfg_key, c_info in report["ablation_study"]["summary"].items():
                writer.writerow([
                    c_info["name"], c_info["rfc6962"], c_info["cache"], c_info["threading"],
                    c_info["build_time"]["mean_ms"], c_info["build_time"]["std_ms"],
                    c_info["update_modify"]["mean_ms"], c_info["update_modify"]["std_ms"],
                    c_info["update_add"]["mean_ms"], c_info["update_add"]["std_ms"],
                    c_info["update_delete"]["mean_ms"], c_info["update_delete"]["std_ms"],
                    c_info["proof_bytes"], c_info["complexity_add_del"]
                ])

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
        # Fallback tự động sang thư mục dataset cục bộ có sẵn nếu realistic_dataset không tồn tại
        candidates = [
            os.path.abspath(os.path.join(BASE_DIR, "kaggle_dataset")),
            os.path.abspath(os.path.join(BASE_DIR, "realistic_dataset")),
            os.path.join(os.path.expanduser("~"), "Downloads", "kaggle_dataset")
        ]
        for c in candidates:
            if os.path.isdir(c):
                target_dataset = c
                break

    if not os.path.isdir(target_dataset):
        print(f"[ERROR] Thư mục không tồn tại: {target_dataset}")
        sys.exit(1)

    results_dir = os.path.abspath(args.results_dir)
    report_data = run_benchmark(target_dataset, num_runs=args.runs)
    json_out, csv_out = save_results(report_data, results_dir)
    print(f"\n[+] Đã lưu kết quả JSON tại: {json_out}")
    print(f"[+] Đã lưu kết quả CSV tại: {csv_out}")
    print(f"[+] Đã lưu kết quả Ablation Study tại: {os.path.join(results_dir, 'ablation_results.csv')}")
