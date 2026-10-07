"""
Mô hình đề xuất cải tiến: HD-MerkleTree
(Hierarchical Dynamic Merkle Tree with RFC 6962 Domain Separation,
Parallel Multithreading, Incremental Hash Caching & Compact Proofs)
"""
from .hd_merkle_tree import HDMerkleTree, TreeNode, NodeType
from .parallel_hasher import parallel_hash_files
from .compact_proof import CompactProof, BatchMultiProof

__all__ = [
    "HDMerkleTree",
    "TreeNode",
    "NodeType",
    "parallel_hash_files",
    "CompactProof",
    "BatchMultiProof"
]
