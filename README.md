# ĐỀ TÀI DT005: HỆ THỐNG KIỂM TRA TOÀN VẸN TỆP BẰNG SHA-256, HMAC VÀ MERKLE TREE

**Môn học:** An Toàn và Bảo Mật Hệ Thống Thông Tin  
**Mã đề tài:** DT005  
**Bài báo tham chiếu lý thuyết:** *The adversary: the philosophy of cryptography*, Harry Halpin (*Journal of Cybersecurity*, Oxford Academic, DOI: 10.1093/cybsec/tyaf006, 2025).  
*Ghi chú khoa học:* Bài báo của Harry Halpin (2025) là công trình nghiên cứu triết học và mô hình hóa hình thức về Kẻ tấn công (The Adversary, Dolev-Yao, Random Oracle, Biba Integrity), không chứa mã nguồn, tập dữ liệu hay số liệu benchmark định lượng. Toàn bộ kiến trúc kỹ thuật, mã nguồn thực nghiệm, bộ kiểm thử tự động và số liệu benchmark trong đề tài DT005 là nghiên cứu thực nghiệm độc lập của nhóm sinh viên trên tập dữ liệu chuẩn hóa 1.000 tệp tin.

---

## 1. CÁC LỆNH CHẠY ĐỘC LẬP TỪ DÒNG LỆNH (CLI RUNNERS)

Mọi mô hình trong dự án đều hỗ trợ giao diện dòng lệnh độc lập (CLI), nhận tham số rõ ràng và trả về kết quả xác thực trực tiếp:

### 1.1. Baseline 1: Flat SHA-256 Manifest
```bash
python -m src.baselines.sha256_manifest --dir realistic_dataset --verify
```
* **Chức năng:** Quét toàn bộ thư mục, tính băm 64 KB streaming chunk, sinh manifest và kiểm tra toàn vẹn danh sách tệp.

### 1.2. Baseline 2: HMAC-SHA256 Manifest (Xác thực nguồn gốc & Chống sửa đổi)
```bash
python -m src.baselines.hmac_manifest --dir realistic_dataset --verify
```
* **Quản lý khóa:** Hỗ trợ nạp khóa bí mật qua biến môi trường `HMAC_SECRET_KEY`, tham số `--key-file <path>` hoặc `--key <str>`. Khóa bí mật được bảo vệ, che dấu (masked) trên log và terminal.

### 1.3. Baseline 3: Standard Binary Merkle Tree
```bash
python -m src.baselines.standard_merkle_tree --dir realistic_dataset
```
* **Chức năng:** Dựng cây Merkle nhị phân truyền thống, tính Root Hash, sinh Inclusion Proof $O(\log N)$ cho từng tệp và xác thực độc lập.

### 1.4. Proposed Model: Hierarchical Dynamic Merkle Tree (HD-MerkleTree)
```bash
python -m src.proposed_model.hd_merkle_tree --dir realistic_dataset --consistency 950
```
* **Chức năng:** Ánh xạ thư mục phân cấp, phân tách miền RFC 6962 (`0x00` lá, `0x01` nhánh), sinh proof phân cấp độc lập và kiểm tra Bằng chứng Nhất quán Chỉ-thêm (RFC 6962 Consistency Proof) từ quá khứ $m=950$ lên $n=1000$.

---

## 2. BỘ KIỂM THỬ TỰ ĐỘNG TOÀN DIỆN (AUTOMATED TEST SUITE)

Dự án trang bị bộ kiểm thử tự động gồm 17 bài test độc lập (`src/tests/test_file_integrity.py`), thực thi qua 1-click hoặc lệnh chuẩn:

```bash
run_tests.bat
# Hoặc:
python -m unittest discover -s src/tests -v
```

* **Phạm vi kiểm thử:**
  1. Tính đúng đắn của giải thuật băm luồng SHA-256 và HMAC-SHA256.
  2. Xác thực thành công khi dữ liệu toàn vẹn; từ chối khi sửa đổi tệp, thêm tệp lạ, xóa tệp hoặc dùng sai khóa bí mật HMAC.
  3. Xác thực tính hợp lệ của Merkle Inclusion Proof trên Standard Merkle Tree và HD-MerkleTree.
  4. Phân tách miền RFC 6962 chống va chạm Second-Preimage.
  5. **Kiểm thử cập nhật tăng dần thực tế:** Thực sự thay đổi, thêm và xóa tệp trên đĩa trong thư mục tạm độc lập, gọi `update_file_fast()`, và đối chiếu Root Hash sau cập nhật khớp chính xác 100% với cây dựng mới từ đầu.
  6. Kiểm thử Bằng chứng Nhất quán RFC 6962 Canonical Log (Consistency Proof) và các điều kiện biên ($m=n, m \le 0, m > n$).
  7. Xử lý tệp rỗng 0-byte, thư mục rỗng và tệp trùng tên ở các thư mục con khác nhau.

---

## 3. KẾT QUẢ BENCHMARK THỰC NGHIỆM ĐỘC LẬP (ĐO TRÊN 1.000 TỆP TIN)

Thực nghiệm được đo đạc tự động qua script benchmark độc lập `src/benchmark/run_benchmark.py` (5 iterations lặp độc lập, làm nóng cache trước khi đo, thao tác cập nhật thực hiện trên bản sao tạm độc lập, không làm biến đổi dữ liệu gốc):

```bash
run_benchmark.bat
# Hoặc:
python -m src.benchmark.run_benchmark --dir realistic_dataset --runs 5 --results-dir results
```

Kết quả được ghi nhận chi tiết, truy nguyên tại `results/benchmark_results.json` và `results/benchmark_results.csv`:

### Bảng Đối Sánh Hiệu Năng Thực Nghiệm (Mean ± Std Dev)

| Chỉ số (Metric) | Flat SHA-256 | HMAC-SHA256 | Standard Merkle | HD-MerkleTree (Đề xuất) | Đánh giá & Rút ra |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Full Build / Verify (ms)** | 62.58 ± 3.56 | 60.65 ± 2.17 | 60.72 ± 3.88 | 89.91 ± 4.52 | HD-Merkle dựng thêm cấu trúc cây thư mục phân cấp |
| **Single-File Verify (ms)** | 0.17 ± 0.01 | 0.07 ± 0.00 | 0.02 ± 0.00 | 0.02 ± 0.00 | Cây Merkle kiểm tra nhanh hơn nhờ $O(\log N)$ |
| **Update Sửa tệp (ms)** | 3.09 ± 3.15 | 0.57 ± 0.05 | 0.14 ± 0.03 | 0.14 ± 0.01 | Cả hai cây Merkle cập nhật lá và nhánh cực nhanh |
| **Update Thêm tệp (ms)** | 2.50 ± 2.03 | 0.67 ± 0.34 | **73.42 ± 5.35** | **0.16 ± 0.01** | **HD-Merkle nhanh hơn ~458 lần** (tránh Rebuild $O(N)$) |
| **Update Xóa tệp (ms)** | 0.32 ± 0.01 | 0.41 ± 0.01 | **70.81 ± 1.99** | **0.10 ± 0.01** | **HD-Merkle nhanh hơn ~708 lần** (tránh Rebuild $O(N)$) |
| **Metadata Storage (KB)** | 111.13 KB | 111.14 KB | 44.76 KB | 111.15 KB | Standard Merkle chỉ lưu mảng băm; các mô hình khác lưu đường dẫn |
| **Proof Size 1 tệp (Bytes)** | 113,802 B | 113,807 B | 990 B | 1,369 B | **Tiết kiệm 98.8% băng thông** so với Flat Manifest |

### Phân Tích Ý Nghĩa Kỹ Thuật & Giới Hạn:
1. **Phân biệt thuật toán và cài đặt (Implementation):** Trong mô hình cơ sở `standard_merkle_tree.py` dạng mảng tĩnh, khi thêm hoặc xóa file, cấu trúc mảng bị xáo trộn vị trí buộc implementation hiện tại phải Rebuild toàn bộ $O(N)$ mất ~71 – 73 ms. Mô hình đề xuất HD-MerkleTree duy trì cấu trúc thư mục động phân cấp: khi thêm/xóa chỉ tác động cục bộ đến nhánh thư mục cha tương ứng ($O(\log N)$ theo thiết kế thuật toán), thời gian đo thực tế đạt 0.10 – 0.16 ms, nhanh hơn **458 đến 708 lần** so với việc Rebuild toàn bộ. Khi sửa 1 file hiện có, cả hai cây Merkle đều đạt thời gian tương đương (~0.14 ms) do chỉ cần cập nhật dọc nhánh lá.
2. **Bản chất của việc tiết kiệm 98.8% băng thông:** Tỷ lệ tiết kiệm băng thông mạng 98.8% là kết quả so sánh giữa việc Light Client chỉ cần nhận Proof phân cấp 1.37 KB (1,369 Bytes) so với việc buộc phải tải toàn bộ Manifest Flat 111.13 KB (113,802 Bytes). So với Standard Merkle (990 Bytes), Proof của HD-Merkle lớn hơn ~38% để đổi lấy việc bảo toàn cấu trúc phân cấp thư mục tự nhiên.
3. **Phân tách rõ kiến trúc RFC 6962:** Cây HD-MerkleTree tính State Root Hash cho trạng thái cấu trúc thư mục hiện tại. Bên cạnh đó, cấu trúc phụ trợ độc lập `RFC6962CanonicalLog` cung cấp Bằng chứng Nhất quán Chỉ-thêm (Append-Only Consistency Proof) theo đúng chuẩn Certificate Transparency RFC 6962.

---

## 4. HƯỚNG DẪN TRẢI NGHIỆM GIAO DIỆN WEB & SLIDE

- **Khởi động Giao diện Web Trực quan:** Nhấp đúp vào `run_demo_web.bat` (hoặc chạy `python src/demo/web_demo.py --port 8080`). Mở trình duyệt tại `http://127.0.0.1:8080`.
- **Chạy Kiểm định 7 Security Games:** `python -m src.proposed_model.security_games`.
- **Tài liệu & Slide trong thư mục `docs/`:**
  - `docs/Bao_Cao_Giua_Ky_DT05.pptx`: Slide thuyết trình PowerPoint 16:9.
  - `docs/Bao_Cao_Tong_Hop_Giua_Ky_DT005.docx`: Báo cáo tổng hợp giữa kỳ.
  - `docs/Thiet_Ke_Kien_Truc_Va_Giai_Thuat.docx`: Đặc tả kiến trúc kỹ thuật và giải thuật.
  - `docs/Mo_Ta_Yeu_Cau_Va_Ke_Hoach_Thuc_Hien.docx`: Bản mô tả yêu cầu và kế hoạch.
  - `docs/Dich_Thuat_Va_Tai_Lap_Bai_Bao.docx`: Dịch thuật và phân tích bài báo của Harry Halpin (2025).
