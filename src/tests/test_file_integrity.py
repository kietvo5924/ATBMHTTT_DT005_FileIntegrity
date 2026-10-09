"""
Bộ Kiểm Thử Tự Động Toàn Diện - Đề Tài DT005
Kiểm định tính đúng đắn của 4 mô hình:
1. Flat SHA-256 Manifest
2. HMAC-SHA256 Manifest (với xác thực khóa bí mật)
3. Standard Binary Merkle Tree
4. Hierarchical Dynamic Merkle Tree (HD-MerkleTree) & RFC 6962 Canonical Log

Mọi phép thử sửa đổi, thêm mới và xóa tệp đều được thực hiện trên thư mục tạm (tempfile),
tuyệt đối không can thiệp vào tập dữ liệu gốc.
"""
import os
import shutil
import tempfile
import unittest
import hashlib
import json

from src.baselines.sha256_manifest import SHA256Manifest, hash_file_sha256
from src.baselines.hmac_manifest import HMACManifest, compute_file_hmac, resolve_secret_key
from src.baselines.standard_merkle_tree import StandardMerkleTree
from src.proposed_model.hd_merkle_tree import (
    HDMerkleTree,
    RFC6962CanonicalLog,
    sha256_leaf,
    sha256_internal,
)


class TestFileIntegritySystem(unittest.TestCase):
    def setUp(self):
        """Khởi tạo môi trường kiểm thử với dữ liệu mẫu trong thư mục tạm lập riêng."""
        self.temp_dir = tempfile.mkdtemp(prefix="dt005_test_")
        
        # Tạo cấu trúc thư mục phân cấp mẫu
        os.makedirs(os.path.join(self.temp_dir, "sub_a", "nested"), exist_ok=True)
        os.makedirs(os.path.join(self.temp_dir, "sub_b"), exist_ok=True)
        os.makedirs(os.path.join(self.temp_dir, "empty_dir"), exist_ok=True)
        
        # 1. Tệp văn bản thường
        self.file1 = os.path.join(self.temp_dir, "file1.txt")
        with open(self.file1, "wb") as f:
            f.write(b"Hello World - ATBMHTTT DT005 Testing")
            
        # 2. Tệp trong thư mục con
        self.file2 = os.path.join(self.temp_dir, "sub_a", "file2.bin")
        with open(self.file2, "wb") as f:
            f.write(b"\x00\x01\x02\x03\xff\xfe\xfd\xfc" * 128)
            
        # 3. Tệp trong thư mục con lồng nhau
        self.file3 = os.path.join(self.temp_dir, "sub_a", "nested", "file3.log")
        with open(self.file3, "wb") as f:
            f.write(b"Log entry line 1\nLog entry line 2\n")
            
        # 4. Tệp trùng tên ở thư mục khác (test va chạm tên trong cấu trúc phân cấp)
        self.file4 = os.path.join(self.temp_dir, "sub_b", "file1.txt")
        with open(self.file4, "wb") as f:
            f.write(b"Different content for file1.txt in sub_b")
            
        # 5. Tệp rỗng (Zero-byte edge case)
        self.file_empty = os.path.join(self.temp_dir, "sub_b", "empty.dat")
        with open(self.file_empty, "wb") as f:
            pass

    def tearDown(self):
        """Dọn dẹp an toàn thư mục tạm sau mỗi bài test."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    # =========================================================================
    # PHẦN 1: KIỂM THỬ BASELINE 1 - FLAT SHA-256 MANIFEST
    # =========================================================================

    def test_sha256_streaming_hash_correctness(self):
        """Kiểm tra hàm băm SHA-256 luồng (streaming) cho kết quả chính xác tuyệt đối."""
        expected_hash = hashlib.sha256(b"Hello World - ATBMHTTT DT005 Testing").hexdigest()
        actual_hash = hash_file_sha256(self.file1)
        self.assertEqual(actual_hash, expected_hash)

    def test_sha256_empty_file(self):
        """Kiểm tra hàm băm với tệp rỗng 0 byte."""
        expected_empty_hash = hashlib.sha256(b"").hexdigest()
        actual_empty_hash = hash_file_sha256(self.file_empty)
        self.assertEqual(actual_empty_hash, expected_empty_hash)

    def test_sha256_manifest_build_and_verify(self):
        """Kiểm tra xây dựng manifest và xác thực toàn vẹn khi dữ liệu không đổi."""
        manifest = SHA256Manifest(self.temp_dir)
        elapsed, storage = manifest.build_manifest()
        
        self.assertEqual(len(manifest.manifest), 5)
        self.assertTrue(manifest.root_hash)
        
        is_valid, violations, v_time = manifest.verify_all()
        self.assertTrue(is_valid)
        self.assertEqual(len(violations), 0)

    def test_sha256_manifest_detect_tampering(self):
        """Kiểm tra SHA-256 manifest phát hiện ngay khi nội dung 1 tệp bị sửa đổi."""
        manifest = SHA256Manifest(self.temp_dir)
        manifest.build_manifest()
        
        # Sửa nội dung tệp trên đĩa
        with open(self.file1, "ab") as f:
            f.write(b" - TAMPERED EXTRA BYTE")
            
        is_valid, violations, _ = manifest.verify_all()
        self.assertFalse(is_valid)
        self.assertTrue(any("file1.txt" in v for v in violations))

    def test_sha256_manifest_detect_file_addition_and_deletion(self):
        """Kiểm tra phát hiện thêm tệp trái phép và xóa tệp."""
        manifest = SHA256Manifest(self.temp_dir)
        manifest.build_manifest()
        
        # 1. Thêm tệp lạ
        unauth_file = os.path.join(self.temp_dir, "unauthorized.exe")
        with open(unauth_file, "wb") as f:
            f.write(b"malicious payload")
        is_valid, violations, _ = manifest.verify_all()
        self.assertFalse(is_valid)
        self.assertTrue(any("UNREGISTERED_FILE" in v for v in violations))
        os.remove(unauth_file)
        
        # 2. Xóa tệp hiện hữu
        os.remove(self.file2)
        is_valid, violations, _ = manifest.verify_all()
        self.assertFalse(is_valid)
        self.assertTrue(any("MISSING_FILE" in v for v in violations))

    # =========================================================================
    # PHẦN 2: KIỂM THỬ BASELINE 2 - HMAC-SHA256 MANIFEST
    # =========================================================================

    def test_hmac_correct_key_verification(self):
        """HMAC xác thực thành công khi dùng đúng khóa bí mật."""
        key = b"super_secret_integrity_key_2026"
        hmac_tool = HMACManifest(self.temp_dir, secret_key=key)
        hmac_tool.build_manifest()
        
        is_valid, violations, _ = hmac_tool.verify_all()
        self.assertTrue(is_valid)
        self.assertEqual(len(violations), 0)

    def test_hmac_wrong_key_rejection(self):
        """HMAC từ chối xác thực khi đối tác kiểm tra dùng sai khóa bí mật."""
        key_signer = b"signer_correct_secret_key_123"
        key_attacker = b"attacker_wrong_guess_key_999"
        
        hmac_signer = HMACManifest(self.temp_dir, secret_key=key_signer)
        hmac_signer.build_manifest()
        
        # Verifier dùng sai khóa
        hmac_verifier = HMACManifest(self.temp_dir, secret_key=key_attacker)
        hmac_verifier.manifest = hmac_signer.manifest
        hmac_verifier.aggregate_hmac = hmac_signer.aggregate_hmac
        
        is_valid, violations, _ = hmac_verifier.verify_all()
        self.assertFalse(is_valid)
        self.assertTrue(len(violations) > 0)

    def test_hmac_key_resolution_env(self):
        """Kiểm tra cơ chế tải khóa an toàn từ biến môi trường."""
        test_env_key = "env_secret_key_dt005_test"
        os.environ["HMAC_SECRET_KEY"] = test_env_key
        try:
            resolved = resolve_secret_key()
            self.assertEqual(resolved, test_env_key.encode("utf-8"))
        finally:
            del os.environ["HMAC_SECRET_KEY"]

    # =========================================================================
    # PHẦN 3: KIỂM THỬ BASELINE 3 - STANDARD BINARY MERKLE TREE
    # =========================================================================

    def test_standard_merkle_build_and_proof(self):
        """Kiểm tra xây dựng Standard Merkle Tree và xác thực Merkle Proof O(log N)."""
        tree = StandardMerkleTree(self.temp_dir)
        tree.build_tree()
        
        self.assertTrue(tree.root_hash)
        self.assertEqual(len(tree.files_list), 5)
        
        # Lấy proof cho 1 file
        target_rel = "file1.txt"
        proof = tree.get_proof(target_rel)
        self.assertIsNotNone(proof)
        
        # Xác thực proof hợp lệ
        target_full = os.path.join(self.temp_dir, target_rel)
        f_hash = hash_file_sha256(target_full)
        is_valid = StandardMerkleTree.verify_proof(target_rel, f_hash, proof, tree.root_hash)
        self.assertTrue(is_valid)

    def test_standard_merkle_proof_reject_tampering(self):
        """Bằng chứng Merkle phải bị từ chối nếu nội dung tệp bị thay đổi."""
        tree = StandardMerkleTree(self.temp_dir)
        tree.build_tree()
        
        target_rel = "file1.txt"
        proof = tree.get_proof(target_rel)
        tampered_hash = hashlib.sha256(b"Tampered Content").hexdigest()
        
        is_valid = StandardMerkleTree.verify_proof(target_rel, tampered_hash, proof, tree.root_hash)
        self.assertFalse(is_valid)

    # =========================================================================
    # PHẦN 4: KIỂM THỬ PROPOSED MODEL - HD-MERKLETREE & THỰC NGHIỆM CẬP NHẬT TĂNG DẦN
    # =========================================================================

    def test_hd_merkle_domain_separation(self):
        """
        Kiểm tra chuẩn phân tách miền RFC 6962:
        0x00 cho nút lá, 0x01 cho nút nhánh nội bộ để chống Second-Preimage Attack.
        """
        leaf_h = sha256_leaf("test.txt", "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855")
        internal_h = sha256_internal(b"\x00" * 32, b"\x01" * 32)
        
        # Hai tiền tố khác nhau đảm bảo không bao giờ trùng hash
        self.assertNotEqual(leaf_h, internal_h)

    def test_hd_merkle_build_and_proof(self):
        """Kiểm tra xây dựng HD-MerkleTree và xác thực proof phân cấp."""
        tree = HDMerkleTree(self.temp_dir)
        tree.build_tree_parallel()
        
        self.assertTrue(tree.root_hash)
        self.assertEqual(len(tree.file_nodes), 5)
        
        # Xác thực proof cho tệp nằm sâu trong thư mục con lồng nhau
        target_rel = "sub_a/nested/file3.log"
        proof = tree.get_proof(target_rel)
        self.assertIsNotNone(proof)
        
        f_node = tree.file_nodes[target_rel]
        is_valid = HDMerkleTree.verify_proof(target_rel, f_node.file_content_hash, proof, tree.root_hash)
        self.assertTrue(is_valid)

    def test_hd_merkle_real_incremental_file_modification(self):
        """
        [YÊU CẦU 3 BẮT BUỘC]: Thực sự thay đổi nội dung tệp trên đĩa trong thư mục tạm,
        sau đó gọi update_file_fast(rel_path).
        So sánh root hash sau cập nhật với root hash khi dựng lại mới 100% từ đầu.
        """
        tree = HDMerkleTree(self.temp_dir)
        tree.build_tree_parallel()
        initial_root = tree.root_hash
        
        target_rel = "sub_a/nested/file3.log"
        target_full = os.path.join(self.temp_dir, "sub_a", "nested", "file3.log")
        
        # 1. Thực sự ghi đè tệp trên đĩa
        with open(target_full, "wb") as f:
            f.write(b"UPDATED REAL DISK CONTENT AT RUNTIME: 2026-10-09")
            
        # 2. Cập nhật tăng dần bằng hàm update_file_fast
        update_time, _ = tree.update_file_fast(target_rel)
        updated_root = tree.root_hash
        
        # Root hash phải đổi
        self.assertNotEqual(initial_root, updated_root)
        
        # 3. Dựng lại 1 cây mới toanh từ đĩa để đối chiếu
        fresh_tree = HDMerkleTree(self.temp_dir)
        fresh_tree.build_tree_parallel()
        fresh_root = fresh_tree.root_hash
        
        # ROOT HASH PHẢI KHỚP TUYỆT ĐỐI VỚI CÂY DỰNG MỚI
        self.assertEqual(
            updated_root,
            fresh_root,
            "Cập nhật tăng dần khi sửa tệp cho ra Root Hash KHÁC BIỆT với cây dựng mới từ đầu!"
        )

    def test_hd_merkle_real_incremental_file_addition(self):
        """
        [YÊU CẦU 3 BẮT BUỘC]: Thực sự tạo thêm 1 tệp mới trên đĩa trong thư mục tạm,
        sau đó gọi update_file_fast(rel_path).
        So sánh root hash sau cập nhật với root hash khi dựng lại từ đầu.
        """
        tree = HDMerkleTree(self.temp_dir)
        tree.build_tree_parallel()
        
        new_rel = "sub_a/nested/brand_new_file.txt"
        new_full = os.path.join(self.temp_dir, "sub_a", "nested", "brand_new_file.txt")
        
        # 1. Thực sự tạo tệp mới trên đĩa
        with open(new_full, "wb") as f:
            f.write(b"Brand new file content created during test")
            
        # 2. Cập nhật tăng dần
        tree.update_file_fast(new_rel)
        updated_root = tree.root_hash
        
        # 3. Dựng lại cây mới từ đầu trên cùng thư mục đĩa
        fresh_tree = HDMerkleTree(self.temp_dir)
        fresh_tree.build_tree_parallel()
        fresh_root = fresh_tree.root_hash
        
        # ROOT HASH PHẢI KHỚP TUYỆT ĐỐI
        self.assertEqual(
            updated_root,
            fresh_root,
            "Cập nhật tăng dần khi THÊM tệp cho ra Root Hash KHÁC BIỆT với cây dựng mới!"
        )

    def test_hd_merkle_real_incremental_file_deletion(self):
        """
        [YÊU CẦU 3 BẮT BUỘC]: Thực sự xóa 1 tệp trên đĩa trong thư mục tạm,
        sau đó gọi update_file_fast(rel_path).
        So sánh root hash sau cập nhật với root hash khi dựng lại từ đầu.
        """
        tree = HDMerkleTree(self.temp_dir)
        tree.build_tree_parallel()
        
        del_rel = "sub_b/file1.txt"
        del_full = os.path.join(self.temp_dir, "sub_b", "file1.txt")
        
        # 1. Thực sự xóa tệp khỏi đĩa
        os.remove(del_full)
        
        # 2. Cập nhật tăng dần
        tree.update_file_fast(del_rel)
        updated_root = tree.root_hash
        
        # 3. Dựng lại cây mới từ đầu
        fresh_tree = HDMerkleTree(self.temp_dir)
        fresh_tree.build_tree_parallel()
        fresh_root = fresh_tree.root_hash
        
        # ROOT HASH PHẢI KHỚP TUYỆT ĐỐI
        self.assertEqual(
            updated_root,
            fresh_root,
            "Cập nhật tăng dần khi XÓA tệp cho ra Root Hash KHÁC BIỆT với cây dựng mới!"
        )

    # =========================================================================
    # PHẦN 5: KIỂM THỬ RFC 6962 CANONICAL LOG & BẰNG CHỨNG NHẤT QUÁN
    # =========================================================================

    def test_rfc6962_canonical_log_consistency_proof(self):
        """
        Kiểm tra Bằng chứng Nhất quán Chỉ-thêm (Append-Only Consistency Proof)
        theo RFC 6962 Certificate Transparency Section 2.1.2.
        """
        # Tạo nhật ký 8 phần tử
        sample_hashes = [hashlib.sha256(f"leaf_{i}".encode()).digest() for i in range(8)]
        log_full = RFC6962CanonicalLog(sample_hashes)
        
        # Test proof từ kích thước quá khứ m=5 lên kích thước n=8
        m, n = 5, 8
        proof_nodes, old_root, new_root = log_full.get_consistency_proof(m)
        
        is_valid = RFC6962CanonicalLog.verify_consistency_proof(m, n, old_root, new_root, proof_nodes)
        self.assertTrue(is_valid, "Bằng chứng nhất quán RFC 6962 hợp lệ bị xác thực thất bại!")
        
        # Test phát hiện giả mạo old_root
        fake_old_root = "0" * 64
        is_fake_valid = RFC6962CanonicalLog.verify_consistency_proof(m, n, fake_old_root, new_root, proof_nodes)
        self.assertFalse(is_fake_valid, "Giả mạo old_root trong RFC 6962 không bị phát hiện!")

    def test_rfc6962_consistency_boundary_conditions(self):
        """Kiểm tra các trường hợp biên của Consistency Proof."""
        sample_hashes = [hashlib.sha256(f"leaf_{i}".encode()).digest() for i in range(4)]
        log = RFC6962CanonicalLog(sample_hashes)
        
        # 1. m == n (cây không đổi kích thước)
        proof, old_r, new_r = log.get_consistency_proof(4)
        self.assertEqual(len(proof), 0)
        self.assertTrue(RFC6962CanonicalLog.verify_consistency_proof(4, 4, old_r, new_r, proof))
        
        # 2. m <= 0 hoặc m > n phải raise ValueError
        with self.assertRaises(ValueError):
            log.get_consistency_proof(0)
        with self.assertRaises(ValueError):
            log.get_consistency_proof(5)


if __name__ == "__main__":
    unittest.main()
