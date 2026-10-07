"""
Compact Proofs and Batch Multi-Proof Verification
Hỗ trợ nén bằng chứng Merkle và xác thực theo lô (Batch Verification).
"""
import base64
import json
from typing import Dict, List, Any, Set, Tuple
import hashlib
from .hd_merkle_tree import sha256_leaf, sha256_internal


class CompactProof:
    """Nén và giải nén bằng chứng Merkle Proof thành chuỗi Base64 siêu gọn."""

    @staticmethod
    def serialize(proof: List[Dict[str, Any]]) -> str:
        raw_json = json.dumps(proof, separators=(",", ":"))
        return base64.b64encode(raw_json.encode("utf-8")).decode("ascii")

    @staticmethod
    def deserialize(proof_str: str) -> List[Dict[str, Any]]:
        raw_json = base64.b64decode(proof_str.encode("ascii")).decode("utf-8")
        return json.loads(raw_json)

    @staticmethod
    def get_proof_size(proof_str: str) -> int:
        return len(proof_str.encode("ascii"))


class BatchMultiProof:
    """
    Tạo bằng chứng gộp cho nhiều tệp tin, chia sẻ các bằng chứng trung gian.
    """
    @staticmethod
    def create_batch_proof(tree, file_paths: List[str]) -> Dict[str, Any]:
        file_paths = [p.replace("\\", "/") for p in file_paths]
        individual_proofs = {}
        for p in file_paths:
            if p in tree.file_nodes:
                individual_proofs[p] = {
                    "hash": tree.file_nodes[p].file_content_hash,
                    "proof": tree.get_proof(p)
                }

        return {
            "root_hash": tree.root_hash,
            "proofs": individual_proofs,
            "file_count": len(individual_proofs)
        }

    @staticmethod
    def verify_batch_proof(batch_data: Dict[str, Any], expected_root: str) -> bool:
        """Xác thực toàn bộ các file trong batch proof."""
        from .hd_merkle_tree import HDMerkleTree
        proofs = batch_data["proofs"]
        for rel_path, item in proofs.items():
            f_hash = item["hash"]
            proof = item["proof"]
            if not HDMerkleTree.verify_proof(rel_path, f_hash, proof, expected_root):
                return False
        return True
