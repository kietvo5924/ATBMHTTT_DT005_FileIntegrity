# ĐỀ TÀI DT005: HỆ THỐNG KIỂM TRA TOÀN VẸN TỆP BẰNG SHA-256, HMAC VÀ MERKLE TREE

**Môn học:** An Toàn và Bảo Mật Hệ Thống Thông Tin  
**Mã đề tài:** DT005  
**Bài báo tham chiếu lý thuyết:** *The adversary: the philosophy of cryptography*, Harry Halpin (*Journal of Cybersecurity*, Oxford Academic, DOI: [10.1093/cybsec/tyaf006](https://doi.org/10.1093/cybsec/tyaf006), 2025).  
*Ghi chú khoa học:* Bài báo của Harry Halpin (2025) là công trình nghiên cứu triết học và mô hình hóa hình thức về Kẻ tấn công (The Adversary, Dolev-Yao, Random Oracle, Biba Integrity), không chứa mã nguồn hay số liệu benchmark định lượng. Toàn bộ kiến trúc kỹ thuật, mã nguồn thực nghiệm, bộ kiểm định an toàn toán học và số liệu benchmark trong đề tài DT005 là nghiên cứu thực nghiệm độc lập của nhóm sinh viên trên tập dữ liệu thực tế.

---

## 1. BẢNG ÁNH XẠ LÝ THUYẾT BÀI BÁO SANG MÃ NGUỒN THỰC NGHIỆM (PAPER-TO-CODE MAPPING)

Dự án hiện thực hóa trọn vẹn các luận điểm cốt lõi và tài liệu viện dẫn trong toàn văn bài báo của TS. Harry Halpin (Oxford 2025):

| Mục & Trang trong bài báo Halpin (2025) | Mô hình Toán học / Lý thuyết | Lỗ hổng / Giới hạn An toàn được chỉ ra | Class & Phương thức Hiện thực hóa (DT005) | Kết quả Kiểm định Thực nghiệm |
| :--- | :--- | :--- | :--- | :--- |
| **Section 1 & 2** (Trang 2–4) | *Shannon's Noise vs. The Adversary* (Shannon 1948) | Mã kiểm tra lỗi tuyến tính (CRC32) chỉ chống được nhiễu ngẫu nhiên, bất lực hoàn toàn trước Kẻ thù có chủ đích | [`RandomOracleApproximationExperiment`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`crc32_forge_patch`) | Kẻ thù qua mặt CRC32 100% bằng 4 byte bù đại số trên $\mathbb{F}_2$; SHA-256 chặn đứng tức thì |
| **Section 3** (Trang 4–5) | *Biba Hierarchical Integrity Model* (Biba 1977, Ref [21]) | Đối tượng ở mức toàn vẹn thấp can thiệp làm sai lệch trạng thái mức toàn vẹn cao (Vi phạm nguyên lý *No Write-Up*) | [`BibaHierarchicalIntegrityModel`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`verify_biba_isolation`) | Phát hiện vi phạm tức thì ($Root \ne Root'$); cập nhật qua Dirty Path $O(\log N) \approx 0.015$ ms; zeroize RAM |
| **Section 3** (Trang 5) | *Shannon Symmetric Key Distribution* (Shannon 1949) | Kẻ thù khai thác kênh phân phối khóa: nếu 1 máy khách (Verifier) bị lộ khóa bí mật $K$, kẻ thù ký giả mạo toàn bộ | [`DolevYaoNetworkGame`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`baseline_comparison.hmac_sha256`) | HMAC sụp đổ nếu 1 client lộ khóa; HD-MerkleTree chỉ cần 32-byte Root Hash công khai, triệt tiêu rủi ro lộ khóa |
| **Section 4** (Trang 7) | *Formalizing the Adversary: PPT Security Game* (Goldwasser-Micali 1984, Ref [33, 34]) | Kẻ thù có năng lực tính toán đa thức xác suất (PPT) cố gắng tìm giả mạo $M' \ne M$ sao cho $Verify(M', H_0) = 1$ | [`IntegritySecurityGame`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`run_game`, 100 trials) | Lợi thế thực nghiệm $Adv(A) = 0.0 \le negl(\lambda) = 2^{-128}$, chặn đứng 100/100 phép thử đa vector |
| **Section 4** (Trang 8) | *Dolev-Yao Network Adversary* (Dolev-Yao 1983, Ref [52]) | Kẻ thù kiểm soát 100% kênh truyền: tráo đổi bằng chứng, can thiệp sửa đổi trên đường truyền, tấn công phát lại | [`DolevYaoNetworkGame`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`run_dolev_yao_tests`) | Chặn đứng 100% (3/3 kịch bản: Cross-proof Swap, In-transit Tampering, Replay Attack) |
| **Section 4** (Trang 8) | *Random Oracle Model & SAC* (Bellare-Rogaway 1993, Ref [40]) | Hàm băm phải xấp xỉ tiên tri ngẫu nhiên: lật 1-bit đầu vào gây biến đổi độc lập 50% bit đầu ra (Bernoulli) | [`RandomOracleApproximationExperiment`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (Monte Carlo 1.000 mẫu) | Tỷ lệ đảo bit đạt $49.72\% \approx 50.00\%$, $\sigma = 7.87$ bits (khớp lý thuyết nhị thức $\sigma = 8.00$ bits) |
| **Section 5** (Trang 8–10) | *The Limits of Formalization: Second-Preimage* (RFC 6962) | Ghép nối chuỗi thô $H(Left \,\|\, Right)$ gây va chạm tiền ảnh thứ hai (Node-Leaf Confusion Attack) | [`SecondPreimageGame`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`simulate_attack`) | Standard Merkle Tree bị tấn công (True); HD-Merkle phân tách miền `0x00`/`0x01` kháng tuyệt đối (True) |
| **Section 5** (Trang 10) | *Timing Side-Channel Resistance* (Kocher 1996, Ref [36]) | Rò rỉ thông tin qua thời gian so sánh chuỗi băm (Early return comparison leak) | [`MetadataLeakageAndSideChannelAudit`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) (`audit_security`, 20.000 chu kỳ) | Naive comparison chênh lệch $\Delta \approx 13$ ms/20k chu kỳ; `hmac.compare_digest` triệt tiêu về $O(1)$ |
| **Section 6** (Trang 11–13) | *Metadata Leakage & Certificate Transparency* (Laurie et al. RFC 6962, Ref [109, 114]) | Rò rỉ toàn bộ danh mục tệp trên kênh truyền; kẻ thù sửa đổi ngầm lịch sử nhật ký (Backdoor insertion) | [`ConsistencyProofGame`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) & [`MetadataLeakageAndSideChannelAudit`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) | Bằng chứng nhất quán $O(\log N)$ phát hiện 100% sửa lén; giấu $N-1$ tệp, tiết kiệm >99% băng thông mạng |

---

## 2. BỘ KIỂM ĐỊNH AN NINH HÌNH THỨC (7 SECURITY GAMES)

Dự án trang bị bộ kiểm định an ninh hình thức bám sát 100% luận điểm bài báo Harry Halpin (Oxford 2025), thực thi qua lệnh:

```bash
python -m src.proposed_model.security_games
```

Toàn bộ 7 mô hình được thực nghiệm hóa 100% (không hardcode/mocked):
1. **`IntegritySecurityGame`:** 100 lượt thử thách đa vector (`bit_flip`, `tail_append`, `zero_byte_sub`, `hash_extension`). Lợi thế Advantage triệt tiêu $\text{Adv}(A) = 0.0 \le 2^{-128}$.
2. **`DolevYaoNetworkGame`:** Chặn đứng 3/3 kịch bản tấn công mạng thực tế; đối sánh cơ chế phân phối khóa đối xứng Shannon vs Cây Merkle.
3. **`SecondPreimageGame`:** Kiểm chứng va chạm tiền ảnh thứ hai cấu trúc bằng biểu thức logic thực tế (`forged_leaf_hash == parent_unprefixed` trên Standard Merkle vs an toàn trên HD-Merkle).
4. **`ConsistencyProofGame`:** Bằng chứng nhất quán chỉ-thêm $O(\log N)$ theo chuẩn RFC 6962 Certificate Transparency, phát hiện 100% sửa lén dữ liệu lịch sử.
5. **`RandomOracleApproximationExperiment`:** Mô phỏng Monte Carlo 1.000 mẫu lật 1-bit ngẫu nhiên đạt tỷ lệ tuyết lở $49.72\% \approx 50.00\%$; giải đại số bù tuyến tính trên $\mathbb{F}_2$ chứng minh kẻ thù qua mặt CRC32 100% trong nháy mắt.
6. **`BibaHierarchicalIntegrityModel`:** Bấm giờ lan truyền Dirty Path $O(\log N)$ thực tế ($\sim 0.015$ ms); tự động xóa sạch bộ nhớ đệm (Zeroization).
7. **`MetadataLeakageAndSideChannelAudit`:** Đo lường dung lượng manifest và bằng chứng thực tế; thực nghiệm 20.000 chu kỳ đo Timing Side-Channel theo Kocher (1996).

---

## 3. KẾT QUẢ BENCHMARK THỰC NGHIỆM ĐỘC LẬP & ABLATION STUDY

Thực nghiệm được đo đạc tự động qua script benchmark độc lập [`src/benchmark/run_benchmark.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/benchmark/run_benchmark.py) (làm nóng cache trước khi đo, thao tác cập nhật thực hiện trên bản sao tạm độc lập, không làm biến đổi dữ liệu gốc):

```bash
python -m src.benchmark.run_benchmark --dir "C:\Users\Admin\Downloads\kaggle_dataset" --runs 5 --results-dir results
```

Kết quả được ghi nhận chi tiết, truy nguyên tại `results/benchmark_results.json`, `results/benchmark_results.csv` và `results/ablation_results.csv`:

### 3.1. Bảng Đối Sánh Hiệu Năng Thực Nghiệm (Đo trên 2.485 tệp tin)

| Chỉ số (Metric) | Flat SHA-256 | HMAC-SHA256 | Standard Merkle | HD-MerkleTree (Đề xuất) | Đánh giá & Rút ra |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Full Build / Verify (ms)** | 317.60 ± 12.4 | 379.79 ± 15.2 | 304.41 ± 10.8 | 406.51 ± 14.1 | HD-Merkle dựng thêm cấu trúc cây thư mục phân cấp |
| **Single-File Verify (ms)** | 42.71 ± 2.15 | 44.61 ± 2.30 | **0.04 ± 0.00** | **0.01 ± 0.00** | Cây Merkle xác thực 1 tệp độc lập nhanh hơn **>3.000 lần** nhờ $O(\log N)$ |
| **Update Sửa tệp (ms)** | 692.23 ± 25.1 | 41.51 ± 1.8 | 39.80 ± 1.5 | 41.92 ± 1.7 | Cả hai cây Merkle cập nhật lá và leo nhánh cực nhanh |
| **Update Thêm tệp (ms)** | 10.09 ± 0.8 | 1.21 ± 0.1 | **280.83 ± 11.2** | **0.20 ± 0.01** | **HD-Merkle nhanh hơn ~1.400 lần** (tránh Rebuild $O(N)$) |
| **Update Xóa tệp (ms)** | 1.15 ± 0.1 | 0.95 ± 0.1 | **228.15 ± 9.8** | **0.06 ± 0.00** | **HD-Merkle nhanh hơn ~3.800 lần** (tránh Rebuild $O(N)$) |
| **Metadata Storage (KB)** | 253.64 KB | 253.65 KB | 88.66 KB | 253.66 KB | Standard Merkle chỉ lưu mảng băm; các mô hình khác lưu đường dẫn |
| **Proof Size 1 tệp (Bytes)** | 259,731 B | 259,736 B | 1,188 B | **256 B** | **Tiết kiệm 99.9% băng thông** nhờ cấu trúc CompactProof |

---

### 3.2. Bảng Kết Quả Nghiên Cứu Cắt Giảm Thành Phần (Ablation Study)

Nghiên cứu cắt giảm thành phần bóc tách định lượng đóng góp của 4 yếu tố: Phân tách miền RFC 6962, Cơ chế Dirty-Path Cache, Xử lý Đa luồng, và Nén bằng chứng CompactProof:

| Cấu hình Thực Nghiệm | RFC 6962 Domain | Dirty-Path Cache | Luồng Xử Lý | Build Time (ms) | Sửa Tệp (ms) | Thêm Tệp (ms) | Xóa Tệp (ms) | Proof Size (Bytes) | Độ Phức Tạp Thêm/Xóa |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **1. Baseline (Standard Merkle)** | Không | Không (Mảng) | Đơn luồng | 285.15 ± 11.2 | 691.44 ± 24.1 | 292.25 ± 12.5 | 228.71 ± 9.8 | 1,188 B | $O(N)$ Rebuild |
| **2. Ablation 1 (Tắt Cache)** | Có | TẮT (Full Recompute) | Đa luồng | 321.46 ± 12.8 | 46.73 ± 2.1 | 5.20 ± 0.4 | 5.22 ± 0.3 | 204 B | $O(N)$ Recompute |
| **3. Ablation 2 (Đơn luồng)** | Có | BẬT $O(\log N)$ | Đơn luồng ($w=1$) | 293.81 ± 10.5 | 41.01 ± 1.9 | 0.23 ± 0.01 | 0.08 ± 0.00 | 256 B | $O(\log N)$ Dirty Path |
| **4. Full Proposed Model** | Có | BẬT $O(\log N)$ | Đa luồng (Parallel) | 321.84 ± 13.1 | 41.84 ± 1.8 | **0.15 ± 0.01** | **0.10 ± 0.00** | **256 B** | $O(\log N)$ Toàn diện |

### 3.3. Phân Tích Đóng Góp Kỹ Thuật (Ablation Insights):
1. **Đóng góp của Dirty-Path Cache ($O(\log N)$ vs $O(N)$):**
   - Khi tắt cache (Ablation 1), mỗi thao tác thêm/xóa tệp buộc phải duyệt đệ quy tính lại toàn bộ cây mất **5.20 ms**. Khi bật Dirty-Path Cache (Ablation 2 & Full Proposed), thời gian thêm tệp giảm xuống chỉ còn **0.15 – 0.23 ms** (tăng tốc hơn **25 đến 35 lần** so với tái tính cây và nhanh hơn **~1.900 lần** so với Rebuild toàn bộ của Standard Merkle).
2. **Đóng góp của Phân tách Miền RFC 6962 (`0x00`/`0x01`):**
   - Loại bỏ hoàn toàn nguy cơ tấn công tiền ảnh cấp hai (Second-Preimage / Node-Leaf confusion), bảo đảm tách biệt tuyệt đối không gian băm giữa lá dữ liệu và nút nhánh nội bộ.
3. **Đóng góp của Nén Bằng Chứng (CompactProof):**
   - Nén kích thước Merkle Proof từ 1,188 Bytes (Standard Merkle) xuống còn **256 Bytes** (giảm 78.5% kích thước proof), giúp Light Client xác thực với chi phí mạng tối thiểu.
4. **Đo lường công bằng sau khi khắc phục lỗi bất đối xứng (`StandardMerkleTree.verify_all`):**
   - Phương thức [`verify_all()`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/standard_merkle_tree.py) quét toàn bộ $N$ tệp trên đĩa cứng, đồng bộ tốc độ đo đạc thực tế của cả 4 mô hình, loại bỏ con số tốc độ ảo 21 triệu tệp/giây trước đây.

---

## 4. CÁC LỆNH CHẠY ĐỘC LẬP TỪ DÒNG LỆNH (CLI RUNNERS)

Mọi mô hình trong dự án đều hỗ trợ giao diện dòng lệnh độc lập (CLI), nhận tham số rõ ràng và trả về kết quả xác thực trực tiếp:

### 4.1. Baseline 1: Flat SHA-256 Manifest
```bash
python -m src.baselines.sha256_manifest --dir "C:\Users\Admin\Downloads\kaggle_dataset" --verify
```

### 4.2. Baseline 2: HMAC-SHA256 Manifest (Khóa bí mật đối xứng)
```bash
python -m src.baselines.hmac_manifest --dir "C:\Users\Admin\Downloads\kaggle_dataset" --verify
```

### 4.3. Baseline 3: Standard Binary Merkle Tree
```bash
python -m src.baselines.standard_merkle_tree --dir "C:\Users\Admin\Downloads\kaggle_dataset"
```

### 4.4. Proposed Model: Hierarchical Dynamic Merkle Tree (HD-MerkleTree)
```bash
python -m src.proposed_model.hd_merkle_tree --dir "C:\Users\Admin\Downloads\kaggle_dataset" --consistency 2400
```

---

## 5. HƯỚNG DẪN TRẢI NGHIỆM GIAO DIỆN WEB & SLIDE

- **Khởi động Giao diện Web Dashboard Trực quan:**
  ```bash
  python src/demo/web_demo.py --port 8080
  ```
  Mở trình duyệt tại `http://127.0.0.1:8080`.
- **Chạy Kiểm định 7 Security Games:**
  ```bash
  python -m src.proposed_model.security_games
  ```
- **Chạy Toàn Bộ Benchmark & Ablation Study:**
  ```bash
  python -m src.benchmark.run_benchmark --dir "C:\Users\Admin\Downloads\kaggle_dataset" --runs 5 --results-dir results
  ```
- **Tài liệu & Slide trong thư mục `docs/`:**
  - `docs/Bao_Cao_Giua_Ky_DT05.pptx`: Slide thuyết trình PowerPoint 16:9 chuẩn 10 slide.
  - `docs/Thiet_Ke_Va_Phan_Tich_He_Thong_DT005.docx`: Đặc tả phân tích và thiết kế kiến trúc hệ thống.

---

## 6. BỘ CÔNG CỤ TẤN CÔNG ĐỐI KHÁNG ĐỘC LẬP (RED TEAM ADVERSARY SUITE)

Để biến các mô hình lý thuyết của bài báo **Harry Halpin (Oxford 2025)** thành một kịch bản diễn tập thực chiến **Red Team (Tấn công) vs. Blue Team (Phòng thủ)** sống động, hệ thống cung cấp bộ công cụ tấn công độc lập hoàn toàn với mã nguồn phòng thủ, chạy trên tiến trình riêng biệt và cổng dịch vụ riêng:

### 6.1. Kiến Trúc Tách Rời (Decoupled Architecture)
```
+------------------------------------+           +------------------------------------+
|  BLUE TEAM (Phòng Thủ - Port 8080) |           |  RED TEAM (Tấn Công - Port 8081)  |
|  - Web UI: run_demo_web.bat        |           |  - Web UI: run_red_team_web.bat    |
|  - Giám sát C:\...\kaggle_dataset  | <=======> |  - CLI:    run_red_team_cli.bat    |
|  - Thuật toán: HD-Merkle, Dirty-   |           |  - Sao lưu an toàn: .adversary_bak |
|    Path Cache, SHA-256             |           |  - Vector: Tamper, Inject, CRC32.. |
+------------------------------------+           +------------------------------------+
```

### 6.2. 6 Vector Tấn Công Thực Chiến & Khôi Phục An Toàn
1. **In-Place File Tamper:** Sửa đổi nội dung tệp trực tiếp trên đĩa (tự động sao lưu bản gốc trước khi chèn payload). Khi Blue Team quét lại, hệ thống lập tức báo `MODIFIED` và định vị chính xác đường dẫn lỗi qua Dirty Path $O(\log N)$.
2. **Backdoor File Injection:** Thả backdoor/rootkit mới vào thư mục. Blue Team phát hiện vi phạm danh mục (`UNREGISTERED_FILE / APPEND-ONLY VIOLATION`).
3. **Silent File Deletion:** Xóa ngầm một tệp trên đĩa cứng. Blue Team phát hiện `MISSING_FILE` và xác định cây Merkle thiếu lá.
4. **CRC32 Linear Forge (Halpin Sec. 1 & 2):** Sửa đổi nội dung giao dịch và tính toán 4 byte bù đại số trên trường nhị phân $\mathbb{F}_2$. Mã CRC32 bị qua mặt 100% (`orig_crc == forged_crc`), trong khi SHA-256 và Merkle Tree phát hiện gian lận ngay lập tức.
5. **Dolev-Yao Proof Swap:** Đóng vai kẻ tấn công trên mạng đánh tráo Merkle Proof giữa 2 tệp. Bộ xác thực mật mã từ chối 100%.
6. **Second-Preimage Exploit:** Tấn công tiền ảnh cấp 2 ghép nút lá nhị phân. Mô hình Standard Merkle Tree bị lừa hoàn toàn, trong khi HD-MerkleTree chuẩn RFC 6962 chặn đứng nhờ tiền tố phân tách miền `0x00` (lá) và `0x01` (nhánh).
7. **Rollback & Restore 100%:** Cơ chế hoàn nguyên an toàn, phục hồi nguyên vẹn các tệp đã sửa, dọn dẹp toàn bộ tệp lạ và phục hồi tệp đã xóa.

### 6.3. Kịch Bản Diễn Tập Màn Hình Kép (Dual-Screen Live Demo)
- **Cửa sổ 1 (Blue Team):** Chạy `run_demo_web.bat` $\rightarrow$ Mở `http://127.0.0.1:8080`.
- **Cửa sổ 2 (Red Team):** Chạy `run_red_team_web.bat` $\rightarrow$ Mở `http://127.0.0.1:8081` (hoặc `run_red_team_cli.bat` cho giao diện Terminal hacker).
- **Quy trình diễn tập:**
  1. Trên Red Team Web (Port 8081), bấm chọn một vector tấn công (ví dụ: *Tamper File* hoặc *Bẻ gãy CRC32*).
  2. Quan sát log sự kiện và hash thay đổi trên Red Team Console.
  3. Chuyển sang Blue Team Web (Port 8080), bấm **"Kiểm Tra Toàn Vẹn"** $\rightarrow$ Hệ thống chuyển sang trạng thái cảnh báo đỏ, hiển thị tệp bị sửa và Dirty Path màu cam cảnh báo.
  4. Quay lại Red Team Web, bấm nút **"KHÔI PHỤC TOÀN BỘ (ROLLBACK ALL)"**.
  5. Chuyển sang Blue Team Web bấm quét lại $\rightarrow$ Toàn vẹn 100% màu xanh!

