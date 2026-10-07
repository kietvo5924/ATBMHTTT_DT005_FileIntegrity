"""
Baseline 3: Standard Binary Merkle Tree
Cây Merkle nhị phân truyền thống:
- Tập hợp tất cả các file làm nút lá (Leaves).
- Ghép đôi các nút cạnh nhau và băm cặp H(Left || Right) lên tới Root Hash.
- Hỗ trợ tạo Merkle Proof (Inclusion Proof) để xác thực từng tệp tin độc lập O(log N).
"""
import os
import json
import time
import hashlib
import hmac
from typing import Dict, List, Tuple, Optional, Any
from .sha256_manifest import hash_file_sha256


def sha256_combine(left: bytes, right: bytes) -> bytes:
    """Băm cặp nhị phân chuẩn H(Left || Right)."""
    return hashlib.sha256(left + right).digest()


class StandardMerkleTree:
    """
    Cây Merkle nhị phân tiêu chuẩn (Standard Binary Merkle Tree).
    """
    def __init__(self, root_dir: str):
        self.root_dir = os.path.abspath(root_dir)
        self.files_list: List[str] = [] # Danh sách sorted relative paths
        self.leaf_hashes: List[bytes] = [] # Danh sách các băm lá
        self.levels: List[List[bytes]] = [] # Các tầng của cây Merkle
        self.root_hash: str = ""

    def build_tree(self) -> Tuple[float, int]:
        """
        Duyệt thư mục, tính băm lá và xây dựng toàn bộ cây Merkle.
        Trả về: (thời gian xây dựng (s), dung lượng lưu trữ (bytes))
        """
        start_time = time.perf_counter()
        
        # 1. Thu thập danh sách tệp
        file_map: Dict[str, bytes] = {}
        for dirpath, _, filenames in os.walk(self.root_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, self.root_dir).replace("\\", "/")
                if rel_path in ("manifest_sha256.json", "manifest_hmac.json", "merkle_standard.json"):
                    continue
                # Băm nội dung file kết hợp với đường dẫn để chống va chạm tên
                h = hashlib.sha256()
                h.update(rel_path.encode("utf-8"))
                h.update(bytes.fromhex(hash_file_sha256(full_path)))
                file_map[rel_path] = h.digest()
                
        self.files_list = sorted(file_map.keys())
        self.leaf_hashes = [file_map[p] for p in self.files_list]
        
        if not self.leaf_hashes:
            self.root_hash = hashlib.sha256(b"EMPTY_TREE").hexdigest()
            self.levels = []
            elapsed = time.perf_counter() - start_time
            return elapsed, self.get_storage_overhead()

        # 2. Xây dựng các tầng cây từ dưới lên trên
        self.levels = [self.leaf_hashes]
        current_level = self.leaf_hashes
        
        while len(current_level) > 1:
            next_level: List[bytes] = []
            for i in range(0, len(current_level), 2):
                left = current_level[i]
                # Nếu số lượng nút lẻ, nhân bản nút cuối (theo quy chuẩn Merkle truyền thống)
                right = current_level[i + 1] if (i + 1 < len(current_level)) else left
                parent = sha256_combine(left, right)
                next_level.append(parent)
            self.levels.append(next_level)
            current_level = next_level

        self.root_hash = self.levels[-1][0].hex()
        elapsed = time.perf_counter() - start_time
        return elapsed, self.get_storage_overhead()

    def get_proof(self, rel_path: str) -> Optional[List[Dict[str, str]]]:
        """
        Tạo bằng chứng Merkle Proof cho một tệp tin cụ thể.
        Trả về danh sách các node anh em (sibling) và hướng (left/right).
        """
        rel_path = rel_path.replace("\\", "/")
        if rel_path not in self.files_list:
            return None
            
        index = self.files_list.index(rel_path)
        proof = []
        
        for level in self.levels[:-1]:
            is_right_child = (index % 2 == 1)
            sibling_index = index - 1 if is_right_child else index + 1
            
            if sibling_index < len(level):
                sibling_hash = level[sibling_index].hex()
            else:
                # Nút lẻ tự nhân bản chính nó
                sibling_hash = level[index].hex()
                
            proof.append({
                "position": "left" if is_right_child else "right",
                "hash": sibling_hash
            })
            index //= 2
            
        return proof

    @staticmethod
    def verify_proof(rel_path: str, file_hash_hex: str, proof: List[Dict[str, str]], expected_root: str) -> bool:
        """
        Xác thực tính toàn vẹn của tệp dựa trên Merkle Proof và Root Hash công khai.
        Độ phức tạp O(log N).
        """
        # Tính lại băm lá
        h = hashlib.sha256()
        h.update(rel_path.replace("\\", "/").encode("utf-8"))
        h.update(bytes.fromhex(file_hash_hex))
        current = h.digest()
        
        for step in proof:
            sibling = bytes.fromhex(step["hash"])
            if step["position"] == "left":
                current = sha256_combine(sibling, current)
            else:
                current = sha256_combine(current, sibling)
                
        return hmac.compare_digest(current.hex().lower(), expected_root.lower())

    def update_file(self, rel_path: str) -> Tuple[float, int]:
        """
        Cập nhật khi 1 tệp bị thay đổi trong Standard Merkle Tree.
        Trong mô hình tĩnh cơ sở, khi tệp thay đổi, thuật toán cập nhật lại nút lá
        và leo ngược lên gốc để tính lại các nút cha dọc nhánh, hoặc tái dựng toàn bộ.
        """
        rel_path = rel_path.replace("\\", "/")
        full_path = os.path.join(self.root_dir, rel_path)
        
        start_time = time.perf_counter()
        if rel_path in self.files_list and os.path.exists(full_path):
            index = self.files_list.index(rel_path)
            h = hashlib.sha256()
            h.update(rel_path.encode("utf-8"))
            h.update(bytes.fromhex(hash_file_sha256(full_path)))
            new_leaf = h.digest()
            self.leaf_hashes[index] = new_leaf
            self.levels[0][index] = new_leaf
            
            # Cập nhật dọc cây lên gốc
            curr_idx = index
            for l in range(len(self.levels) - 1):
                is_right = (curr_idx % 2 == 1)
                left_idx = curr_idx - 1 if is_right else curr_idx
                right_idx = curr_idx if is_right else curr_idx + 1
                
                left = self.levels[l][left_idx]
                right = self.levels[l][right_idx] if right_idx < len(self.levels[l]) else left
                parent = sha256_combine(left, right)
                
                parent_idx = curr_idx // 2
                self.levels[l + 1][parent_idx] = parent
                curr_idx = parent_idx
                
            self.root_hash = self.levels[-1][0].hex()
        else:
            # Nếu thêm mới hoặc xóa file, cấu trúc mảng bị xáo trộn -> Bắt buộc Rebuild toàn bộ
            self.build_tree()
            
        elapsed = time.perf_counter() - start_time
        return elapsed, self.get_storage_overhead()

    def get_storage_overhead(self) -> int:
        """Đo dung lượng cây Merkle (bytes)."""
        data = {
            "root_hash": self.root_hash,
            "file_count": len(self.files_list),
            "files": self.files_list,
            "levels_count": len(self.levels)
        }
        return len(json.dumps(data).encode("utf-8"))

    def save_to_file(self, out_path: Optional[str] = None) -> str:
        """Lưu cây Merkle ra tệp JSON."""
        target = out_path or os.path.join(self.root_dir, "merkle_standard.json")
        data = {
            "root_hash": self.root_hash,
            "files": self.files_list,
            "leaf_hashes": [h.hex() for h in self.leaf_hashes]
        }
        with open(target, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return target
