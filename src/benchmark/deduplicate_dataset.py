"""
Kiểm Tra Trùng Lặp Dữ Liệu (Deduplication & Dataset Integrity Check) - Dự Án DT005
Tuân thủ tiêu chuẩn khoa học và yêu cầu đề tài DT005:
- Kiểm tra tính duy nhất (uniqueness) của các tệp tin trong tập dữ liệu.
- Phát hiện tệp tin trùng lặp nội dung (SHA-256 hash collision / duplicate content).
- Tính toán mã băm gốc toàn vẹn ban đầu (Initial Root Manifest Checksum).
- Xuất báo cáo khoa học ra định dạng JSON phục vụ kiểm định và tái lập kết quả.
"""
import os
import sys
import time
import json
import hashlib
import argparse
from typing import Dict, List, Any, Tuple

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def analyze_dataset_deduplication(dataset_dir: str) -> Dict[str, Any]:
    """
    Quét toàn bộ tập dữ liệu, tính SHA-256 từng tệp, kiểm tra trùng lặp
    và sinh Root Manifest Checksum.
    """
    target_dir = os.path.abspath(dataset_dir)
    if not os.path.isdir(target_dir):
        raise ValueError(f"Thư mục không tồn tại: {target_dir}")

    hash_to_files: Dict[str, List[str]] = {}
    file_manifest: List[Dict[str, Any]] = []
    total_files = 0
    total_bytes = 0

    t_start = time.perf_counter()

    for root, _, files in os.walk(target_dir):
        for filename in files:
            file_path = os.path.join(root, filename)
            rel_path = os.path.relpath(file_path, target_dir).replace("\\", "/")
            file_size = os.path.getsize(file_path)
            total_files += 1
            total_bytes += file_size

            # Tính SHA-256 streaming
            hasher = hashlib.sha256()
            try:
                with open(file_path, "rb") as fp:
                    while chunk := fp.read(65536):
                        hasher.update(chunk)
                digest = hasher.hexdigest()
            except Exception as e:
                digest = f"ERROR: {str(e)}"

            hash_to_files.setdefault(digest, []).append(rel_path)
            file_manifest.append({
                "rel_path": rel_path,
                "size_bytes": file_size,
                "sha256": digest
            })

    elapsed_sec = time.perf_counter() - t_start

    # Phân tích trùng lặp
    unique_hashes_count = len(hash_to_files)
    duplicates = {h: flist for h, flist in hash_to_files.items() if len(flist) > 1}
    duplicate_groups_count = len(duplicates)
    duplicate_files_count = sum(len(flist) for flist in duplicates.values())
    redundant_files_count = total_files - unique_hashes_count

    # Sắp xếp canonical theo đường dẫn để tính Root Manifest Checksum xác định (deterministic)
    file_manifest.sort(key=lambda x: x["rel_path"])
    manifest_string = "\n".join([f"{entry['rel_path']}:{entry['sha256']}" for entry in file_manifest])
    root_manifest_hash = hashlib.sha256(manifest_string.encode("utf-8")).hexdigest()

    rel_dataset_path = os.path.relpath(target_dir, BASE_DIR).replace("\\", "/")
    if rel_dataset_path.startswith(".."):
        rel_dataset_path = target_dir

    report = {
        "metadata": {
            "dataset_directory": rel_dataset_path,
            "scan_timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "scan_duration_seconds": round(elapsed_sec, 4),
            "hashing_algorithm": "SHA-256",
            "license": "CC0: Public Domain (Kaggle)",
            "source_url": "https://www.kaggle.com/datasets/snehaanbhawal/resume-dataset"
        },
        "summary": {
            "total_files": total_files,
            "total_bytes": total_bytes,
            "total_megabytes": round(total_bytes / (1024 * 1024), 2),
            "unique_content_hashes": unique_hashes_count,
            "duplicate_groups": duplicate_groups_count,
            "duplicate_files_total": duplicate_files_count,
            "redundant_files_to_remove": redundant_files_count,
            "deduplication_rate_percent": round((1.0 - (unique_hashes_count / total_files if total_files > 0 else 1.0)) * 100.0, 2),
            "root_manifest_checksum": root_manifest_hash
        },
        "duplicates": duplicates
    }

    return report


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="Dataset Deduplication & Integrity Scanner - DT005")
    parser.add_argument("--dir", default="kaggle_dataset", help="Dataset directory to analyze")
    parser.add_argument("--output", default="results/deduplication_report.json", help="Path to save output JSON report")
    args = parser.parse_args()

    target_dir = args.dir
    if not os.path.isdir(target_dir):
        # Fallback candidates
        candidates = [
            os.path.join(BASE_DIR, "kaggle_dataset"),
            os.path.join(BASE_DIR, "realistic_dataset"),
            r"C:\Users\Admin\Downloads\kaggle_dataset"
        ]
        for c in candidates:
            if os.path.isdir(c):
                target_dir = c
                break

    print(f"[*] Bắt đầu quét kiểm tra trùng lặp trên: {target_dir}")
    report = analyze_dataset_deduplication(target_dir)

    summary = report["summary"]
    meta = report["metadata"]

    print("\n" + "=" * 70)
    print("      KẾT QUẢ KIỂM TRA TRÙNG LẶP DỮ LIỆU (DEDUPLICATION REPORT) - DT005")
    print("=" * 70)
    print(f"• Thư mục kiểm tra        : {meta['dataset_directory']}")
    print(f"• Thời gian quét          : {meta['scan_duration_seconds']} giây")
    print(f"• Tổng số tệp tin quét    : {summary['total_files']:,} tệp")
    print(f"• Tổng dung lượng đĩa     : {summary['total_bytes']:,} Bytes ({summary['total_megabytes']} MB)")
    print(f"• Số mã băm duy nhất      : {summary['unique_content_hashes']:,} mã SHA-256")
    print(f"• Số nhóm tệp trùng hash  : {summary['duplicate_groups']} nhóm")
    print(f"• Số tệp dư thừa cần lọc  : {summary['redundant_files_to_remove']} tệp (Tỷ lệ trùng: {summary['deduplication_rate_percent']}%)")
    print(f"• Root Manifest Checksum  : {summary['root_manifest_checksum']}")
    print("=" * 70)

    # Lưu file JSON
    out_path = os.path.abspath(args.output)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    print(f"[+] Báo cáo chi tiết đã được ghi vào: {out_path}\n")


if __name__ == "__main__":
    main()
