# -*- coding: utf-8 -*-
"""
Red Team - Adversary Engine (Mô đun Tấn công Độc lập)
Dựa trên bài báo khoa học: "The Adversary: The Philosophy of Cryptography" (Harry Halpin, Oxford 2025).

Cung cấp các vector tấn công đối kháng trực tiếp vào thư mục dữ liệu mà Blue Team (Port 8080) đang giám sát:
1. Tamper File (Phá hoại / Sửa đổi nội dung tệp tin)
2. Inject Backdoor (Thả tệp lạ / Cửa sau vào thư mục)
3. Silent Delete (Xóa lén tệp dữ liệu ngầm)
4. CRC32 Linear Forge (Bẻ gãy mã kiểm tra tuyến tính CRC32 bằng thuật toán bù GF(2))
5. Dolev-Yao MitM Attack (Đánh tráo Merkle Proof giữa 2 tệp trên kênh truyền mạng)
6. Second-Preimage Exploit (Chế tạo nút lá giả mạo ghép nối nhị phân qua mặt Standard Merkle Tree)
7. Rollback & Restore (Khôi phục an toàn 100% về nguyên trạng ban đầu)
"""
import os
import sys
import time
import json
import zlib
import hmac
import shutil
import hashlib
import urllib.request
from typing import Dict, List, Tuple, Any, Optional

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from src.proposed_model.security_games import crc32_forge_patch
from src.proposed_model.hd_merkle_tree import HDMerkleTree, sha256_leaf, sha256_internal
from src.baselines.standard_merkle_tree import StandardMerkleTree, sha256_combine


class AdversaryEngine:
    """
    Động cơ Kẻ tấn công (The Adversary Engine):
    Quản lý trạng thái mục tiêu, sao lưu an toàn và phát động các đòn tấn công mạng / toán học.
    """

    def __init__(self, target_dir: Optional[str] = None, blue_team_url: str = "http://127.0.0.1:8080"):
        self.blue_team_url = blue_team_url
        self.target_dir = target_dir or self._discover_target_from_blue_team()
        self.backup_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".adversary_backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        self.state_file = os.path.join(self.backup_dir, "adversary_state.json")
        # Quản lý danh sách các tệp bị tác động để hoàn nguyên
        self.tampered_records: Dict[str, str] = {}  # rel_path -> backup_path
        self.injected_files: List[str] = []         # rel_paths
        self.deleted_records: Dict[str, str] = {}   # rel_path -> backup_path
        self.attack_logs: List[Dict[str, Any]] = []
        self.load_state()
        self.reconcile_state_with_disk()

    def save_state(self):
        """Lưu trạng thái tấn công ra tệp JSON để duy trì tính nhất quán khi máy chủ tải lại."""
        try:
            with open(self.state_file, "w", encoding="utf-8") as f:
                json.dump({
                    "target_dir": self.target_dir,
                    "tampered_records": self.tampered_records,
                    "injected_files": self.injected_files,
                    "deleted_records": self.deleted_records,
                    "attack_logs": self.attack_logs[:50]
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[RED TEAM] Cảnh báo lưu trạng thái: {e}")

    def load_state(self):
        """Khôi phục trạng thái tấn công từ tệp JSON."""
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        self.tampered_records = data.get("tampered_records", {})
                        self.injected_files = data.get("injected_files", [])
                        self.deleted_records = data.get("deleted_records", {})
                        self.attack_logs = data.get("attack_logs", [])
            except Exception as e:
                print(f"[RED TEAM] Cảnh báo tải trạng thái: {e}")

    def reconcile_state_with_disk(self):
        """Đồng bộ trạng thái tấn công thực tế trên đĩa (tránh mất vết khi khởi động lại)."""
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return

        # Rà soát vũ khí / artifact Red Team có thật trên đĩa
        known_artifacts = [
            ".adversary_attack_manifest.json",
            "forged_transaction_crc32.dat",
            "dolev_yao_swapped_proof.json",
            "second_preimage_exploit.bin",
            "adversary_metadata_recon.json"
        ]
        for art in known_artifacts:
            art_p = os.path.join(self.target_dir, art)
            if os.path.exists(art_p) and art not in self.injected_files:
                self.injected_files.append(art)

        # Rà soát mã độc backdoor có thật trên đĩa
        try:
            for fn in os.listdir(self.target_dir):
                if (fn.startswith("backdoor_") or fn.startswith("rootkit_")) and fn.endswith(".py"):
                    if fn not in self.injected_files:
                        self.injected_files.append(fn)
        except Exception:
            pass

        # Lọc bỏ tệp lạ đã bị xóa thủ công ngoài đĩa
        self.injected_files = [f for f in self.injected_files if os.path.exists(os.path.join(self.target_dir, f))]

    def _discover_target_from_blue_team(self) -> str:
        """Tự động truy vấn Blue Team trên Port 8080 để lấy thư mục đang giám sát."""
        try:
            req = urllib.request.Request(f"{self.blue_team_url}/api/status", headers={"User-Agent": "RedTeam-Adversary/1.0"})
            with urllib.request.urlopen(req, timeout=1.5) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                td = data.get("target_dir")
                if td and os.path.exists(td):
                    return td
        except Exception:
            pass

        # Fallback các ứng viên cục bộ
        candidates = [
            r"C:\Users\Admin\Downloads\kaggle_dataset",
            os.path.abspath(os.path.join(os.path.dirname(__file__), "../../realistic_dataset")),
            os.path.abspath(os.path.join(os.path.dirname(__file__), "../../kaggle_dataset"))
        ]
        for c in candidates:
            if os.path.isdir(c):
                return c
        return os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))

    def get_target_files(self) -> List[str]:
        """Lấy danh sách tệp tin trong thư mục mục tiêu (bỏ qua tệp hệ thống)."""
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return []
        flist = []
        for dp, _, fns in os.walk(self.target_dir):
            for fn in fns:
                rel = os.path.relpath(os.path.join(dp, fn), self.target_dir).replace("\\", "/")
                if rel.startswith(".git") or "__pycache__" in rel or rel.endswith(".bak"):
                    continue
                flist.append(rel)
        flist.sort()
        return flist

    def get_real_dataset_files(self) -> List[str]:
        """Lấy danh sách các tệp dữ liệu thật (loại trừ các file vũ khí/artifact của Red Team)."""
        files = self.get_target_files()
        ignore_names = {
            ".adversary_attack_manifest.json",
            "forged_transaction_crc32.dat",
            "dolev_yao_swapped_proof.json",
            "second_preimage_exploit.bin",
            "adversary_metadata_recon.json",
            "manifest_sha256.json",
            "manifest_hmac.json",
            "merkle_standard.json",
            "merkle_hd.json"
        }
        real_files = []
        for f in files:
            base_n = os.path.basename(f)
            if base_n in ignore_names or base_n.startswith(("backdoor_", "rootkit_")):
                continue
            if f in self.injected_files:
                continue
            real_files.append(f)
        return real_files

    def log_event(self, attack_name: str, status: str, details: Dict[str, Any]) -> Dict[str, Any]:
        event = {
            "timestamp": time.strftime("%H:%M:%S"),
            "attack": attack_name,
            "status": status,
            "details": details
        }
        self.attack_logs.insert(0, event)
        if len(self.attack_logs) > 50:
            self.attack_logs.pop()
        self.save_state()
        return event

    def set_target_dir(self, new_path: str) -> Dict[str, Any]:
        """Thiết lập thư mục mục tiêu tấn công linh hoạt (hỗ trợ tương đối, tuyệt đối, hoặc preset)."""
        new_path = (new_path or "").strip()
        if not new_path:
            return {"status": "error", "message": "Đường dẫn thư mục không được để trống!"}

        root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        downloads_kaggle = os.path.join(downloads_dir, "kaggle_dataset")
        downloads_kaggle_large = os.path.join(downloads_dir, "kaggle_dataset_large")
        realistic_dataset = os.path.join(root_dir, "realistic_dataset")
        kaggle_dataset = os.path.join(root_dir, "kaggle_dataset")

        # Phân giải đường dẫn
        if new_path == "realistic_dataset":
            abs_p = realistic_dataset if os.path.exists(realistic_dataset) else (downloads_kaggle if os.path.exists(downloads_kaggle) else realistic_dataset)
        elif new_path in ("kaggle_dataset_large", "Downloads/kaggle_dataset_large", "Downloads\\kaggle_dataset_large"):
            abs_p = downloads_kaggle_large
        elif new_path in ("kaggle_dataset", "Downloads/kaggle_dataset", "Downloads\\kaggle_dataset"):
            abs_p = downloads_kaggle if os.path.exists(downloads_kaggle) else kaggle_dataset
        else:
            check_downloads = os.path.join(downloads_dir, new_path)
            if not os.path.isabs(new_path) and os.path.exists(check_downloads) and os.path.isdir(check_downloads):
                abs_p = check_downloads
            else:
                abs_p = os.path.abspath(new_path if os.path.isabs(new_path) else os.path.join(root_dir, new_path))

        if not os.path.exists(abs_p) or not os.path.isdir(abs_p):
            return {"status": "error", "message": f"Thư mục không tồn tại: {new_path}"}

        self.target_dir = abs_p
        files = self.get_target_files()
        msg = f"Đã chuyển mục tiêu tấn công sang: {os.path.basename(abs_p)} ({len(files)} tệp)"
        print(f"\n[RED TEAM] [CHUYỂN MỤC TIÊU] {abs_p} - {len(files)} tệp tin", flush=True)
        self.log_event("Change Target", "CHUYỂN MỤC TIÊU THÀNH CÔNG", {
            "target_dir": abs_p,
            "target_file_count": len(files)
        })
        return {
            "status": "ok",
            "target_dir": abs_p,
            "file_count": len(files),
            "message": msg
        }

    def sync_target_from_blue_team(self) -> Dict[str, Any]:
        """Truy vấn Blue Team (Port 8080) và đồng bộ thư mục mục tiêu đang giám sát."""
        td = self._discover_target_from_blue_team()
        if td and os.path.isdir(td):
            self.target_dir = td
            files = self.get_target_files()
            msg = f"Đã đồng bộ từ Blue Team: {os.path.basename(td)} ({len(files)} tệp)"
            print(f"\n[RED TEAM] [ĐỒNG BỘ BLUE TEAM] {td} - {len(files)} tệp tin", flush=True)
            self.log_event("Sync Blue Team", "ĐỒNG BỘ THÀNH CÔNG", {
                "target_dir": td,
                "target_file_count": len(files)
            })
            return {
                "status": "ok",
                "target_dir": td,
                "file_count": len(files),
                "message": msg
            }
        return {
            "status": "error",
            "message": "Không thể kết nối hoặc Blue Team chưa chọn thư mục mục tiêu!"
        }

    def clear_target_dir(self) -> Dict[str, Any]:
        """Hủy chọn thư mục mục tiêu hiện tại."""
        self.target_dir = None
        print(f"\n[RED TEAM] [HỦY CHỌN] Đã hủy chọn thư mục mục tiêu.", flush=True)
        self.log_event("Clear Target", "HỦY CHỌN THƯ MỤC", {})
        return {
            "status": "ok",
            "target_dir": "",
            "file_count": 0,
            "message": "Đã hủy chọn thư mục mục tiêu thành công"
        }

    def open_target_in_explorer(self) -> Dict[str, Any]:
        """Mở thư mục trong File Explorer của hệ điều hành."""
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return {"status": "error", "message": "Chưa chọn thư mục mục tiêu!"}
        try:
            abs_p = os.path.abspath(self.target_dir)
            if sys.platform == "win32":
                import subprocess
                subprocess.Popen(["explorer.exe", os.path.normpath(abs_p)])
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", abs_p])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", abs_p])
            return {"status": "ok", "message": "Đã mở File Explorer"}
        except Exception as e:
            return {"status": "error", "message": str(e)}


    # ==========================================================================
    # ĐÒN TẤN CÔNG 1: SỬA ĐỔI NỘI DUNG TỆP (IN-PLACE FILE TAMPERING)
    # ==========================================================================
    def attack_tamper_file(self, target_rel_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Sửa đổi nội dung 1 tệp tin thật trên đĩa.
        Mục đích: Khiến hash của tệp thay đổi, kích hoạt cảnh báo Dirty-Path trên Blue Team.
        """
        real_files = self.get_real_dataset_files()
        if not real_files:
            real_files = self.get_target_files()
        if not real_files:
            return {"success": False, "error": "Không tìm thấy tệp mục tiêu"}

        # Ưu tiên tệp dữ liệu thật (Resume/Resume.csv hoặc tệp tiếp theo trong tập dữ liệu)
        # Lọc ra các tệp chưa bị sửa và chưa bị xóa để có thể tấn công chồng tiếp nhiều lần
        untampered = [f for f in real_files if f not in self.tampered_records and f not in self.deleted_records]
        if untampered:
            pref = [f for f in untampered if "Resume.csv" in f or "Resume" in f]
            default_file = pref[0] if pref else untampered[0]
        else:
            default_file = real_files[0] if real_files else ""
        rel_path = target_rel_path if (target_rel_path and target_rel_path in real_files) else default_file
        if not rel_path:
            return {"success": False, "error": "Không còn tệp khả dụng để sửa"}
        full_path = os.path.join(self.target_dir, rel_path)

        # Sao lưu trước khi sửa
        if rel_path not in self.tampered_records:
            bak_path = os.path.join(self.backup_dir, f"tamper_{os.path.basename(rel_path)}_{int(time.time()*1000)}.bak")
            shutil.copy2(full_path, bak_path)
            self.tampered_records[rel_path] = bak_path

        # Băm trước khi sửa
        with open(full_path, "rb") as f:
            orig_data = f.read()
        orig_hash = hashlib.sha256(orig_data).hexdigest()

        # Thực hiện tấn công chèn payload độc hại
        payload = b"\n/* [RED-TEAM ATTACK] MALICIOUS_INJECTION_DOLEV_YAO_PAYLOAD */\n"
        with open(full_path, "ab") as f:
            f.write(payload)

        # Băm sau khi sửa
        with open(full_path, "rb") as f:
            tampered_data = f.read()
        tampered_hash = hashlib.sha256(tampered_data).hexdigest()

        res = {
            "success": True,
            "file": rel_path,
            "original_sha256": orig_hash,
            "tampered_sha256": tampered_hash,
            "injected_payload_bytes": len(payload),
            "expected_blue_team_reaction": "Blue Team sẽ phát hiện sửa đổi tệp và định vị lỗi qua đường dẫn bẩn"
        }
        self.log_event("Tamper File", "THỰC THI THÀNH CÔNG", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 2: THẢ TỆP LẠ / CỬA SAU (BACKDOOR FILE INJECTION)
    # ==========================================================================
    def attack_inject_backdoor(self, filename: Optional[str] = None) -> Dict[str, Any]:
        """
        Tạo một tệp mới tinh vào thư mục giám sát.
        Mục đích: Phá vỡ tính toàn vẹn danh mục, vi phạm kiểm định Append-Only của Blue Team.
        """
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return {"success": False, "error": "Chưa chọn thư mục mục tiêu"}

        fname = filename or f"backdoor_exploit_{int(time.time()*1000)}.py"
        full_path = os.path.join(self.target_dir, fname)

        content = (
            f"# [RED TEAM] Malicious Injected Rootkit / Backdoor\n"
            f"# Created: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"import os; os.system('echo HACKED BY RED TEAM ADVERSARY')\n"
        ).encode("utf-8")

        with open(full_path, "wb") as f:
            f.write(content)

        rel_path = os.path.relpath(full_path, self.target_dir).replace("\\", "/")
        if rel_path not in self.injected_files:
            self.injected_files.append(rel_path)

        res = {
            "success": True,
            "injected_file": rel_path,
            "file_size": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
            "expected_blue_team_reaction": "Blue Team sẽ phát hiện tệp chưa đăng ký và xâm nhập lạ"
        }
        self.log_event("Inject Backdoor", "THỰC THI THÀNH CÔNG", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 3: XÓA LÉN TỆP DỮ LIỆU (SILENT FILE DELETION)
    # ==========================================================================
    def attack_delete_file(self, target_rel_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Xóa một tệp tin thật trong thư mục mục tiêu.
        Mục đích: Khiến Blue Team phát hiện MISSING_FILE khi quét toàn vẹn.
        """
        real_files = self.get_real_dataset_files()
        if not real_files:
            real_files = self.get_target_files()
        if not real_files:
            return {"success": False, "error": "Không tìm thấy tệp để xóa"}

        # Tránh xóa trùng tệp đang bị sửa hoặc đã bị xóa
        candidates = [f for f in real_files if f not in self.tampered_records and f not in self.deleted_records]
        if not candidates:
            candidates = [f for f in real_files if f not in self.deleted_records]
        if not candidates:
            return {"success": False, "error": "Không còn tệp khả dụng để xóa"}
        rel_path = target_rel_path if (target_rel_path and target_rel_path in candidates) else candidates[-1]
        full_path = os.path.join(self.target_dir, rel_path)

        # Sao lưu trước khi xóa
        bak_path = os.path.join(self.backup_dir, f"del_{os.path.basename(rel_path)}_{int(time.time()*1000)}.bak")
        shutil.copy2(full_path, bak_path)
        self.deleted_records[rel_path] = bak_path

        os.remove(full_path)

        res = {
            "success": True,
            "deleted_file": rel_path,
            "expected_blue_team_reaction": "Blue Team sẽ phát hiện thiếu tệp và tính lại gốc Merkle"
        }
        self.log_event("Delete File", "THỰC THI THÀNH CÔNG", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 4: BẺ GÃY MÃ TUYẾN TÍNH CRC32 (CRC32 FORGE COLLISION)
    # ==========================================================================
    def attack_crc32_forge(self, custom_msg: Optional[str] = None) -> Dict[str, Any]:
        """
        Hiện thực hóa Luận điểm Harry Halpin (Section 1 & 2):
        Sửa đổi thông điệp và tính 4 byte bù đại số trên GF(2) để qua mặt hoàn toàn mã CRC32.
        """
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return {"success": False, "error": "Chưa chọn thư mục mục tiêu"}
        orig_msg = b"Transaction: Transfer 1000 USD to Alice Account 10123"
        tampered_msg = (custom_msg or "Transaction: Transfer 999,999 USD to RED_TEAM_HACKER").encode("utf-8")

        # Bẻ gãy tuyến tính CRC32 và ghi tệp ra đĩa
        patch = crc32_forge_patch(orig_msg, tampered_msg)
        forged_payload = tampered_msg + patch

        orig_crc = zlib.crc32(orig_msg)
        forged_crc = zlib.crc32(forged_payload)

        # Đối chiếu với SHA-256
        orig_sha = hashlib.sha256(orig_msg).hexdigest()
        forged_sha = hashlib.sha256(forged_payload).hexdigest()

        crc_bypassed = (orig_crc == forged_crc)
        sha_detected = (orig_sha != forged_sha)

        # Ghi tệp giao dịch giả mạo thật ra đĩa
        disk_filename = "forged_transaction_crc32.dat"
        disk_path = os.path.join(self.target_dir, disk_filename)
        with open(disk_path, "wb") as f_out:
            f_out.write(forged_payload)
        if disk_filename not in self.injected_files:
            self.injected_files.append(disk_filename)

        res = {
            "success": True,
            "artifact_file": disk_filename,
            "original_message": orig_msg.decode("utf-8"),
            "tampered_message": tampered_msg.decode("utf-8"),
            "algebraic_patch_hex": patch.hex(),
            "crc32_result": {
                "original_crc": hex(orig_crc),
                "forged_crc": hex(forged_crc),
                "crc32_status": "BỊ ĐÁNH LỪA HOÀN TOÀN: Trùng khớp mã 100%" if crc_bypassed else "Thất bại"
            },
            "sha256_result": {
                "original_sha256": orig_sha,
                "forged_sha256": forged_sha,
                "sha256_status": "PHÁT HIỆN GIAN LẬN NGAY LẬP TỨC" if sha_detected else "Bị lừa"
            },
            "academic_conclusion": "CRC32 chỉ kháng được nhiễu ngẫu nhiên, sụp đổ trước kẻ thù toán học. SHA-256 an toàn tuyệt đối."
        }
        self.log_event("CRC32 Linear Forge", "BẺ GÃY CRC32 100%", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 5: TRÁO ĐỔI MERKLE PROOF DOLEV-YAO (CROSS-PROOF SWAP)
    # ==========================================================================
    def attack_dolev_yao_swap(self) -> Dict[str, Any]:
        """
        Kẻ thù Dolev-Yao chặn kênh truyền mạng và tráo đổi Proof của Tệp A gán cho Tệp B.
        Xuất bằng chứng giả mạo ra tệp thật trên đĩa để Blue Team kiểm định thực tế.
        """
        files = self.get_target_files()
        if len(files) < 2:
            return {"success": False, "error": "Cần ít nhất 2 tệp để thực hiện tráo đổi Proof"}

        file_a, file_b = files[0], files[1]
        tree = HDMerkleTree(self.target_dir)
        tree.build_tree_parallel()
        root = tree.root_hash

        proof_b = tree.get_proof(file_b)
        hash_a = tree.file_nodes[file_a].file_content_hash

        # Kiểm chứng khi gửi file_a, hash_a, proof_b, root
        accepted = HDMerkleTree.verify_proof(file_a, hash_a, proof_b, root)

        # Xuất vật chứng Dolev-Yao thật ra đĩa
        proof_filename = "dolev_yao_swapped_proof.json"
        proof_path = os.path.join(self.target_dir, proof_filename)
        with open(proof_path, "w", encoding="utf-8") as f_out:
            json.dump({
                "adversary_model": "Dolev-Yao Network MitM",
                "attack_type": "Cross-Proof Swap",
                "file_a": file_a,
                "hash_a": hash_a,
                "forged_proof_b": proof_b,
                "target_root_hash": root,
                "intercepted_at": time.strftime("%Y-%m-%d %H:%M:%S")
            }, f_out, ensure_ascii=False, indent=2)
        if proof_filename not in self.injected_files:
            self.injected_files.append(proof_filename)

        res = {
            "success": True,
            "artifact_file": proof_filename,
            "attacker_action": f"Đánh tráo Proof của '{file_b}' gán cho '{file_a}'",
            "proof_steps": len(proof_b),
            "target_root_hash": root,
            "blue_team_verification": "HỢP LỆ: Bị lừa" if accepted else "BỊ TỪ CHỐI 100%: Phát hiện gian lận",
            "safe": not accepted
        }
        self.log_event("Dolev-Yao Proof Swap", "CHẶN ĐỨNG BỞI HD-MERKLE", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 6: SECOND-PREIMAGE EXPLOIT (NODE-LEAF CONFUSION)
    # ==========================================================================
    def attack_second_preimage_exploit(self) -> Dict[str, Any]:
        """
        Chứng minh tấn công tiền ảnh cấp hai:
        Chế tạo nút lá giả mạo từ việc ghép nối 2 nút con và ghi payload ra tệp thật trên đĩa.
        """
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return {"success": False, "error": "Chưa chọn thư mục mục tiêu"}
        left = hashlib.sha256(b"red_team_leaf_node_1").digest()
        right = hashlib.sha256(b"red_team_leaf_node_2").digest()

        # 1. Standard Merkle không tiền tố:
        std_parent = sha256_combine(left, right)
        forged_payload = left + right
        std_forged_leaf_hash = hashlib.sha256(forged_payload).digest()
        std_bypassed = bool(std_forged_leaf_hash == std_parent)

        # 2. HD-MerkleTree chuẩn RFC 6962:
        hd_leaf_1 = sha256_leaf("f1", left.hex())
        hd_leaf_2 = sha256_leaf("f2", right.hex())
        hd_parent = sha256_internal(hd_leaf_1, hd_leaf_2)
        hd_fake_leaf = sha256_leaf("fake_leaf", (hd_leaf_1 + hd_leaf_2).hex())
        hd_blocked = bool(hd_fake_leaf != hd_parent)

        # Ghi payload giả mạo thật ra đĩa
        exploit_filename = "second_preimage_exploit.bin"
        exploit_path = os.path.join(self.target_dir, exploit_filename)
        with open(exploit_path, "wb") as f_out:
            f_out.write(forged_payload)
        if exploit_filename not in self.injected_files:
            self.injected_files.append(exploit_filename)

        res = {
            "success": True,
            "artifact_file": exploit_filename,
            "standard_merkle_tree": {
                "vulnerability": "Không dùng tiền tố phân tách miền",
                "attack_result": "HACK THÀNH CÔNG: Lá giả có mã băm trùng khớp nút cha" if std_bypassed else "Thất bại"
            },
            "hd_merkle_tree": {
                "defense": "Chuẩn phân tách miền RFC 6962: 0x00 lá, 0x01 nhánh",
                "attack_result": "CHẶN ĐỨNG 100%: 0x00 khác 0x01 triệt tiêu va chạm" if hd_blocked else "Bị hack"
            }
        }
        self.log_event("Second-Preimage Attack", "STANDARD BỊ HACK - HD-MERKLE AN TOÀN", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 7: TẤN CÔNG KÊNH PHỤ THỜI GIAN (KOCHER 1996 TIMING ATTACK)
    # ==========================================================================
    def attack_timing_side_channel(self) -> Dict[str, Any]:
        """
        Hiện thực hóa Luận điểm Harry Halpin Section 5 Ref 36:
        Tấn công kênh phụ thời gian nhắm vào so khớp mã băm / HMAC.
        """
        target_token = b"a8f7c9e2b1d40365" * 4  # 64 bytes token
        wrong_first_byte = b"x" + target_token[1:]
        wrong_last_byte = target_token[:-1] + b"x"

        def naive_compare(a: bytes, b: bytes) -> bool:
            if len(a) != len(b):
                return False
            for x, y in zip(a, b):
                if x != y:
                    return False
            return True

        # Đo thời gian non-constant time
        trials = 10000
        t0 = time.perf_counter()
        for _ in range(trials):
            naive_compare(target_token, wrong_first_byte)
        time_first_byte = (time.perf_counter() - t0) * 1e6 / trials

        t0 = time.perf_counter()
        for _ in range(trials):
            naive_compare(target_token, wrong_last_byte)
        time_last_byte = (time.perf_counter() - t0) * 1e6 / trials

        # Đo thời gian constant-time hmac.compare_digest
        t0 = time.perf_counter()
        for _ in range(trials):
            hmac.compare_digest(target_token, wrong_first_byte)
        time_const_first = (time.perf_counter() - t0) * 1e6 / trials

        t0 = time.perf_counter()
        for _ in range(trials):
            hmac.compare_digest(target_token, wrong_last_byte)
        time_const_last = (time.perf_counter() - t0) * 1e6 / trials

        timing_leak_ratio = (time_last_byte / time_first_byte) if time_first_byte > 0 else 1.0

        res = {
            "success": True,
            "paper_reference": "Harry Halpin 2025, Section 5: The Limits of Formalization, Kocher 1996",
            "naive_comparison": {
                "early_exit_mismatch_first_byte_us": round(time_first_byte, 4),
                "late_exit_mismatch_last_byte_us": round(time_last_byte, 4),
                "timing_leak_ratio": round(timing_leak_ratio, 2),
                "vulnerability": "RÒ RỈ THỜI GIAN NGHIÊM TRỌNG: Kẻ tấn công suy đoán được từng byte bí mật"
            },
            "constant_time_defense": {
                "first_byte_mismatch_us": round(time_const_first, 4),
                "last_byte_mismatch_us": round(time_const_last, 4),
                "status": "AN TOÀN TUYỆT ĐỐI: Thời gian thực thi đồng nhất, triệt tiêu kênh phụ"
            }
        }
        self.log_event("Timing Side-Channel Probe", "HOÀN TẤT ĐO ĐẠC KÊNH PHỤ", res)
        return res

    # ==========================================================================
    # ĐÒN TẤN CÔNG 8: THU THẬP SIÊU DỮ LIỆU & PHÂN TÍCH LƯU LƯỢNG (TRAFFIC ANALYSIS)
    # ==========================================================================
    def attack_metadata_recon(self) -> Dict[str, Any]:
        """
        Hiện thực hóa Luận điểm Harry Halpin Section 6 Ref 15, 100:
        Trích xuất toàn bộ cấu trúc tổ chức qua siêu dữ liệu và xuất file báo cáo ra đĩa.
        """
        files = self.get_target_files()
        if not files:
            return {"success": False, "error": "Chưa có tệp mục tiêu"}

        total_bytes = 0
        ext_distribution: Dict[str, int] = {}
        dir_distribution: Dict[str, int] = {}

        for rel in files:
            full = os.path.join(self.target_dir, rel)
            try:
                sz = os.path.getsize(full)
                total_bytes += sz
            except Exception:
                sz = 0
            ext = os.path.splitext(rel)[1].lower() or "no_ext"
            ext_distribution[ext] = ext_distribution.get(ext, 0) + 1
            top_dir = rel.split("/")[0] if "/" in rel else "root"
            dir_distribution[top_dir] = dir_distribution.get(top_dir, 0) + 1

        recon_filename = "adversary_metadata_recon.json"
        recon_path = os.path.join(self.target_dir, recon_filename)
        recon_data = {
            "adversary_intel": "Traffic and Metadata Reconnaissance",
            "paper_reference": "Harry Halpin 2025 Section 6",
            "total_monitored_files": len(files),
            "total_size_mb": round(total_bytes / (1024 * 1024), 2),
            "top_subdirectories": dict(sorted(dir_distribution.items(), key=lambda x: x[1], reverse=True)[:5]),
            "file_type_breakdown": dict(sorted(ext_distribution.items(), key=lambda x: x[1], reverse=True)[:6]),
            "scanned_at": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(recon_path, "w", encoding="utf-8") as f_out:
            json.dump(recon_data, f_out, ensure_ascii=False, indent=2)
        if recon_filename not in self.injected_files:
            self.injected_files.append(recon_filename)

        res = {
            "success": True,
            "artifact_file": recon_filename,
            "paper_reference": "Harry Halpin 2025, Section 6: The Adversary Revealed, Snowden Leaks",
            "traffic_metadata_intelligence": recon_data,
            "adversary_deduction": "Kẻ tấn công nắm rõ cấu trúc phân cấp, vị trí tệp quan trọng mà không cần phá vỡ mã băm SHA-256.",
            "blue_team_countermeasure": "Cần áp dụng kỹ thuật Padding độ dài cố định và ẩn danh hóa cấu trúc đường dẫn."
        }
        self.log_event("Traffic & Metadata Recon", "TRÍCH XUẤT SIÊU DỮ LIỆU HOÀN TẤT", res)
        return res

    # ==========================================================================
    # PHÁT ĐỘNG TOÀN DIỆN ĐỒNG LOẠT 7 ĐÒN TẤN CÔNG THỰC TẾ LÊN ĐĨA
    # ==========================================================================
    def attack_launch_all_7(self) -> Dict[str, Any]:
        """
        Thực thi đồng loạt 7 đòn tấn công đối kháng trực tiếp lên thư mục mục tiêu trên đĩa:
        1. Phá hoại / Sửa đổi tệp thật
        2. Chèn cửa sau / Tệp mã độc thật
        3. Xóa lén tệp dữ liệu ngầm thật
        4. Bẻ gãy đại số CRC32 và ghi tệp giao dịch giả mạo thật
        5. Đánh tráo bằng chứng Merkle Proof trên mạng Dolev-Yao thật
        6. Tấn công va chạm tiền ảnh cấp hai ghép nút lá thật
        7. Khai thác kênh phụ đo thời gian và trích xuất siêu dữ liệu thư mục
        """
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return {"success": False, "error": "Chưa chọn thư mục mục tiêu"}
        t0 = time.perf_counter()
        r1 = self.attack_tamper_file()
        r2 = self.attack_inject_backdoor()
        r3 = self.attack_delete_file()
        r4 = self.attack_crc32_forge()
        r5 = self.attack_dolev_yao_swap()
        r6 = self.attack_second_preimage_exploit()
        r7 = self.attack_timing_side_channel()
        r8 = self.attack_metadata_recon()
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        # Xuất file Manifest tấn công tổng hợp vào thư mục mục tiêu
        manifest_name = ".adversary_attack_manifest.json"
        manifest_path = os.path.join(self.target_dir, manifest_name)
        manifest_data = {
            "attacks_active": True,
            "launched_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "elapsed_ms": round(elapsed_ms, 2),
            "attack_1_tamper_file": r1.get("file"),
            "attack_2_injected_backdoor": r2.get("injected_file"),
            "attack_3_deleted_file": r3.get("deleted_file"),
            "attack_4_crc32_file": "forged_transaction_crc32.dat",
            "attack_5_dolev_yao_proof_file": "dolev_yao_swapped_proof.json",
            "attack_6_second_preimage_file": "second_preimage_exploit.bin",
            "attack_7_metadata_recon_file": "adversary_metadata_recon.json"
        }
        with open(manifest_path, "w", encoding="utf-8") as f_out:
            json.dump(manifest_data, f_out, ensure_ascii=False, indent=2)
        if manifest_name not in self.injected_files:
            self.injected_files.append(manifest_name)

        summary = {
            "success": True,
            "attacks_executed_count": 7,
            "elapsed_ms": round(elapsed_ms, 2),
            "manifest_file": manifest_name,
            "attacks": {
                "vector_1_tamper": r1,
                "vector_2_inject": r2,
                "vector_3_delete": r3,
                "vector_4_crc32": r4,
                "vector_5_dolev_yao": r5,
                "vector_6_second_preimage": r6,
                "vector_7_timing_and_metadata": {"timing": r7, "metadata": r8}
            },
            "status": "Đã phát động thành công toàn diện 7 đòn tấn công đối kháng trực tiếp lên đĩa"
        }
        self.log_event("Launch All 7 Attacks", "PHÁT ĐỘNG TOÀN DIỆN THÀNH CÔNG", summary)
        return summary


    # ==========================================================================
    # HOÀN NGUYÊN / KHÔI PHỤC AN TOÀN 100% (SAFETY ROLLBACK & RESTORE)
    # ==========================================================================
    def rollback_all(self) -> Dict[str, Any]:
        """
        Khôi phục tất cả các tệp bị sửa đổi, xóa bỏ tệp lạ và phục hồi tệp đã xóa.
        Đưa thư mục của Blue Team về trạng thái nguyên vẹn sạch sẽ 100%.
        """
        if not self.target_dir or not os.path.isdir(self.target_dir):
            return {
                "success": False,
                "error": "Chưa chọn thư mục mục tiêu",
                "restored_tampered_files": 0,
                "cleaned_injected_files": 0,
                "restored_deleted_files": 0,
                "status": "Chưa chọn thư mục mục tiêu để hoàn nguyên"
            }
        restored_tampered = 0
        cleaned_injected = 0
        restored_deleted = 0

        # 1. Khôi phục tệp bị sửa đổi
        for rel_path, bak_path in list(self.tampered_records.items()):
            full_path = os.path.join(self.target_dir, rel_path)
            if os.path.exists(bak_path):
                shutil.copy2(bak_path, full_path)
                try:
                    os.remove(bak_path)
                except Exception:
                    pass
                restored_tampered += 1
        self.tampered_records.clear()

        # 2. Xóa tệp độc hại đã inject
        for rel_path in list(self.injected_files):
            full_path = os.path.join(self.target_dir, rel_path)
            if os.path.exists(full_path):
                try:
                    os.remove(full_path)
                    cleaned_injected += 1
                except Exception:
                    pass
        self.injected_files.clear()

        # Xóa triệt để toàn bộ các tệp vật chứng đối kháng từ Red Team trên đĩa
        known_artifacts = [
            "forged_transaction_crc32.dat",
            "dolev_yao_swapped_proof.json",
            "second_preimage_exploit.bin",
            "adversary_metadata_recon.json",
            ".adversary_attack_manifest.json"
        ]
        for art in known_artifacts:
            art_path = os.path.join(self.target_dir, art)
            if os.path.exists(art_path):
                try:
                    os.remove(art_path)
                    cleaned_injected += 1
                except Exception:
                    pass

        # Xóa các tệp backdoor/rootkit còn sót lại
        if os.path.exists(self.target_dir):
            for fn in os.listdir(self.target_dir):
                if (fn.startswith("backdoor_") or fn.startswith("rootkit_")) and fn.endswith(".py"):
                    try:
                        os.remove(os.path.join(self.target_dir, fn))
                        cleaned_injected += 1
                    except Exception:
                        pass

        # 3. Phục hồi tệp bị xóa
        for rel_path, bak_path in list(self.deleted_records.items()):
            full_path = os.path.join(self.target_dir, rel_path)
            if os.path.exists(bak_path):
                os.makedirs(os.path.dirname(full_path), exist_ok=True)
                shutil.copy2(bak_path, full_path)
                try:
                    os.remove(bak_path)
                except Exception:
                    pass
                restored_deleted += 1
        self.deleted_records.clear()

        manifest_path = os.path.join(self.target_dir, ".adversary_attack_manifest.json")
        if os.path.exists(manifest_path):
            try:
                os.remove(manifest_path)
            except Exception:
                pass

        res = {
            "success": True,
            "restored_tampered_files": restored_tampered,
            "cleaned_injected_files": cleaned_injected,
            "restored_deleted_files": restored_deleted,
            "status": "Thư mục mục tiêu đã được hoàn nguyên về trạng thái nguyên vẹn 100%"
        }
        self.log_event("Rollback All", "HOÀN NGUYÊN HOÀN TẤT", res)
        return res

    def get_status(self) -> Dict[str, Any]:
        """Lấy trạng thái Red Team hiện tại."""
        files = self.get_target_files()
        return {
            "target_dir": self.target_dir,
            "target_file_count": len(files),
            "tampered_count": len(self.tampered_records),
            "tampered_files": list(self.tampered_records.keys()),
            "injected_count": len(self.injected_files),
            "injected_files": self.injected_files,
            "deleted_count": len(self.deleted_records),
            "deleted_files": list(self.deleted_records.keys()),
            "logs_count": len(self.attack_logs),
            "latest_logs": self.attack_logs[:10]
        }
