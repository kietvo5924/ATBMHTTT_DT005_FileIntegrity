# BÁO CÁO KIỂM TOÁN MÃ NGUỒN VÀ MÔ HÌNH THỰC NGHIỆM GIỮA KỲ (DT005)
**Đề tài:** Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và Merkle Tree  
**Bài báo đối chiếu:** *The adversary: the philosophy of cryptography* — Harry Halpin, *Journal of Cybersecurity*, Oxford Academic, 2025 (DOI: [10.1093/cybsec/tyaf006](https://doi.org/10.1093/cybsec/tyaf006)).  
**Tài liệu:** Báo cáo kiểm toán kỹ thuật nội bộ (`MIDTERM_CODE_AUDIT.md`)

---

## 1. NGUYÊN TẮC HỌC THUẬT VÀ BẢN CHẤT CỦA BÀI BÁO GỐC

### 1.1. Bản chất công trình của Harry Halpin (2025)
- Bài báo của TS. Harry Halpin là một công trình nghiên cứu thuộc lĩnh vực **triết học của mật mã học (Philosophy of Cryptography)**, lịch sử phát triển các mô hình toán học về kẻ thù và phân tích giới hạn của sự trừu tượng hóa (limits of abstraction).
- **Hiện trạng thực tế:** Bài báo **HOÀN TOÀN KHÔNG CÓ** mã nguồn thực thi, không có tập dữ liệu (dataset) thực nghiệm, không có bất kỳ số liệu benchmark định lượng nào về thời gian thực thi (ms) hay dung lượng bằng chứng (bytes).
- **Phân định ranh giới học thuật:**
  1. *Luận điểm của Halpin:* Sự chuyển dịch từ mô hình nhiễu Shannon (1948) sang mô hình đối thủ có chủ đích; các mô hình hình thức (Goldwasser-Micali PPT 1984, Dolev-Yao 1983, Bellare-Rogaway Random Oracle Model 1993); phê phán giới hạn hình thức hóa (tấn công kênh kề Kocher 1996, rò rỉ siêu dữ liệu, đối thủ thể chế NSA/Snowden 2013); và 4 câu hỏi nghiên cứu mở về mạng xã hội-kỹ thuật phức hợp.
  2. *Bộ mã nguồn đề tài DT005:* Toàn bộ các lớp mã nguồn (`security_games.py`, `hd_merkle_tree.py`, `sha256_manifest.py`, `hmac_manifest.py`, `standard_merkle_tree.py`, `adversary_engine.py`) là **thí nghiệm minh họa độc lập do sinh viên tự xây dựng** nhằm mô phỏng và minh họa các luận điểm của Halpin.

---

## 2. KẾT QUẢ KIỂM TOÁN VÀ XỬ LÝ 8 VẤN ĐỀ TRỌNG YẾU (D1 – D8)

### D1. `IntegritySecurityGame` — Tính toán Advantage và Phân biệt Giới hạn Thực nghiệm vs Lý thuyết
- **File & Hàm liên quan:** [`src/proposed_model/security_games.py`](src/proposed_model/security_games.py) (`IntegritySecurityGame.run_game`).
- **Vấn đề phát hiện:**
  - Trước đây, mã nguồn tính `advantage = adversary_wins / num_trials` (với 100 phép thử), sau đó so sánh `advantage < 1e-6` và xuất ra thông báo: `theoretical_bound: Adv A <= 2^-128`.
  - Việc đồng nhất tỷ lệ thành công quan sát được trong 100 phép thử với giới hạn toán học $2^{-128}$ là sai lệch về phương pháp luận mật mã học. Thử nghiệm thực nghiệm trên mẫu hữu hạn chỉ đo khả năng phát hiện can thiệp thực tế, không thể chứng minh toán học cho toàn bộ không gian khóa $2^{256}$.
- **Đã xử lý & Sửa đổi:**
  - Đổi tên trường trả về: `empirical_success_rate = adversary_wins / num_trials` và `detection_rate_pct = (1.0 - empirical_success_rate) * 100.0`.
  - Tách bạch trường lý thuyết: `theoretical_security_note = "Giới hạn va chạm lý thuyết của SHA-256 là 2^-128; thực nghiệm 100 phép thử chỉ đo tỷ lệ phát hiện thực tế trong phạm vi mẫu, không phải chứng minh toán học"`.
  - Cập nhật docstring và verdict: Diễn giải đúng là "Thí nghiệm minh họa thách đấu kiểm tra toàn vẹn", ghi nhận tỷ lệ phát hiện thực tế trong 100 phép thử.

### D2. `SecondPreimageGame` — Va chạm Tiền ảnh cấp hai Cấu trúc & Phân tách miền RFC 6962
- **File & Hàm liên quan:** [`src/proposed_model/security_games.py`](src/proposed_model/security_games.py) (`SecondPreimageGame.simulate_attack`), [`src/proposed_model/hd_merkle_tree.py`](src/proposed_model/hd_merkle_tree.py).
- **Vấn đề phát hiện:**
  - Phép thử tạo `std_parent = sha256_combine(left, right)` và `std_forged_leaf_hash = hashlib.sha256(forged_payload).digest()` với `forged_payload = left + right`, nên theo định nghĩa hai chuỗi này bằng nhau.
  - Cần làm rõ đây là mô phỏng toán học của dạng tấn công **Node-Leaf Confusion** (sự nhầm lẫn giữa nút lá và nút nhánh nội bộ khi ghép nối nhị phân thô không có tiền tố phân tách miền).
  - Không được tuyên bố RFC 6962 "giải quyết mọi loại va chạm của SHA-256".
- **Đã xử lý & Sửa đổi:**
  - Cập nhật docstring giải thích rõ giả định: Trong cây nhị phân thuần túy ghép thô $H(L \,\|\, R)$, nếu một lá nhận nội dung 64 byte đúng bằng $L \,\|\, R$ thì mã băm lá trùng khớp với mã băm của nút cha.
  - Phân tách miền RFC 6962 sử dụng tiền tố `0x00` (Leaf) và `0x01` (Internal Node) giúp triệt tiêu sự nhầm lẫn cấu trúc tầng cây ($H(0x00 \,\|\, \text{leaf}) \ne H(0x01 \,\|\, L \,\|\, R)$).
  - Ghi rõ giới hạn: Cơ chế tiền tố phân tách miền chỉ giải quyết va chạm cấu trúc tầng cây, không loại bỏ bài toán tìm tiền ảnh thứ hai ngẫu nhiên $2^{256}$ của thuật toán băm SHA-256.

### D3. `ConsistencyProofGame` — Bằng chứng Nhất quán Toán học vs Dấu hiệu Ngoại vi
- **File & Hàm liên quan:** [`src/proposed_model/security_games.py`](src/proposed_model/security_games.py) (`ConsistencyProofGame.test_consistency`), [`src/proposed_model/hd_merkle_tree.py`](src/proposed_model/hd_merkle_tree.py).
- **Vấn đề phát hiện:**
  - Trước đây, code kiểm tra tệp chứa chuỗi "backdoor" hoặc đọc `.adversary_attack_manifest.json` trên đĩa và đưa vào verdict chung của mô hình toán học.
  - Việc tìm tên tệp hay metadata là cờ ngoại vi (heuristic), không phải bằng chứng mật mã (cryptographic proof).
- **Đã xử lý & Sửa đổi:**
  - Tách bạch hoàn toàn: Đưa cờ tệp vào nhóm `heuristic_disk_marker` (chỉ nhằm phục vụ quan sát trực quan trong kịch bản demo diễn tập).
  - Lõi kiểm chứng mật mã là hàm toán học `HDMerkleTree.verify_consistency_proof(past_size, current_size, old_root, new_root, proof_nodes)`.
  - Thực hiện kiểm thử phá hoại lịch sử thực sự: Tạo mốc `tampered_old_root = hashlib.sha256(b"corrupted_history_marker").hexdigest()` và xác nhận hàm toán học từ chối 100% (`tamper_detected == True`).

### D4. `BibaHierarchicalIntegrityModel` — Lan truyền Cập nhật Băm Cục bộ vs Mô hình Biba Đầy đủ
- **File & Hàm liên quan:** [`src/proposed_model/security_games.py`](src/proposed_model/security_games.py) (`BibaHierarchicalIntegrityModel.verify_biba_isolation`).
- **Vấn đề phát hiện:**
  - Code thực chất đo thời gian cập nhật lan truyền Dirty Path $O(\log N)$ trong RAM khi một tệp trên đĩa thay đổi và đối sánh với việc quét lại toàn bộ thư mục của Flat SHA-256.
  - Code không cài đặt một hệ thống kiểm soát truy cập Biba đầy đủ (vốn đòi hỏi mô hình toán học gồm tập chủ thể $S$, tập đối tượng $O$, hàm gán nhãn mức toàn vẹn $I(s), I(o)$, và bộ kiểm tra Reference Monitor cưỡng chế nguyên lý No Write-Up và No Read-Down).
- **Đã xử lý & Sửa đổi:**
  - Đổi tên và phân loại chính xác: "Thí nghiệm Minh họa Lan Truyền Cập Nhật Toàn Vẹn Phân Cấp: Tham chiếu Biba".
  - Ghi rõ trong docstring và báo cáo: Đây là thực nghiệm kỹ thuật độc lập minh họa tính chất toàn vẹn phân cấp (thay đổi ở lá làm thay đổi Root tức thì), không tuyên bố đã cài đặt toàn diện mô hình kiểm soát truy cập Biba.

### D5. `RandomOracleApproximationExperiment` — Hiệu ứng Tuyết Lở (SAC) vs Random Oracle Model
- **File & Hàm liên quan:** [`src/proposed_model/security_games.py`](src/proposed_model/security_games.py) (`RandomOracleApproximationExperiment.evaluate_rom_properties`).
- **Vấn đề phát hiện:**
  - Trước đây, verdict ghi "SHA-256 xấp xỉ hoàn hảo Random Oracle... an toàn tuyệt đối".
  - Việc đo hiệu ứng tuyết lở (Strict Avalanche Criterion - SAC) qua 1.000 mẫu lật 1-bit cho ra tỷ lệ ~49.7% chỉ là khảo sát thống kê về tính khuếch tán. Trong lý thuyết mật mã (Bellare-Rogaway 1993, Halpin Section 4), không có hàm băm cụ thể nào có thể chứng minh tương đương với một Random Oracle thực sự.
- **Đã xử lý & Sửa đổi:**
  - Sửa tiêu đề: "Thí nghiệm Khảo sát Tính chất Tuyết Lở (SAC) & Đối sánh Tính Phi Tuyến SHA-256 vs CRC32".
  - Bổ sung `scientific_note`: Khẳng định rõ Random Oracle là mô hình trừu tượng lý tưởng hóa; thực nghiệm SAC trên tập mẫu hữu hạn minh họa tính khuếch tán tốt, không phải chứng minh toán học cho toàn bộ hàm băm.
  - Giữ lại thí nghiệm bẻ gãy mã tuyến tính CRC32 bằng giải thuật khử Gauss trên trường $\mathbb{F}_2$: Minh họa chính xác luận điểm Halpin ở Section 1 & 2 rằng mã kiểm tra lỗi ngẫu nhiên tuyến tính bị kẻ thù có chủ đích qua mặt 100%.

### D6. `MetadataLeakageAndSideChannelAudit` — Khảo sát Kênh Kề Đo Thời Gian (Timing Side-Channel)
- **File & Hàm liên quan:** [`src/proposed_model/security_games.py`](src/proposed_model/security_games.py) (`MetadataLeakageAndSideChannelAudit.audit_security`).
- **Vấn đề phát hiện:**
  - Việc đo thời gian so sánh chuỗi bằng vòng lặp Python trên một máy đơn chịu tác động rất lớn bởi jitter hệ điều hành, lập lịch CPU và cơ chế dọn rác (GC).
  - Không thể kết luận một phép đo microsecond cục bộ chứng minh "triệt tiêu hoàn toàn" mọi nguy cơ kênh kề.
- **Đã xử lý & Sửa đổi:**
  - Cải tiến quy trình đo: Chạy 5 vòng độc lập, ngẫu nhiên hóa thứ tự thử nghiệm giữa các cặp chuỗi để triệt tiêu thiên lệch bộ nhớ đệm CPU.
  - Tính toán và báo cáo độ lệch trung vị (median delta) qua các vòng.
  - Bổ sung `measurement_limitation`: Nêu rõ hàm `hmac.compare_digest` triệt tiêu rò rỉ thời gian ở mức thư viện C (constant-time byte comparison), nhưng thực nghiệm đo đạc ở mức microsecond trên môi trường máy đơn mang tính minh họa nguyên lý, không phải chứng minh loại trừ 100% mọi dạng kênh kề phần cứng (cache timing, branch prediction, power analysis).

### D7. `HMACManifest` — Quản lý Khóa Bí Mật & Cảnh Báo Khóa Mặc Định
- **File & Hàm liên quan:** [`src/baselines/hmac_manifest.py`](src/baselines/hmac_manifest.py) (`resolve_secret_key`).
- **Vấn đề phát hiện:**
  - Hàm `resolve_secret_key()` có fallback trả về khóa bí mật hardcoded `b"dt005_test_hmac_secret_key_32bytes"` nếu người dùng không truyền khóa.
- **Đã xử lý & Sửa đổi:**
  - Bổ sung cảnh báo an ninh rõ ràng: Khi fallback được kích hoạt, hàm phát ra cảnh báo `warnings.warn("CẢNH BÁO AN NINH: Đang sử dụng khóa bí mật kiểm thử mặc định...")`.
  - Làm rõ trong docstring thứ tự ưu tiên lấy khóa: 1. `key_file` -> 2. `key_str` -> 3. `env_var` (`HMAC_SECRET_KEY`) -> 4. `allow_testing_fallback`.
  - Trong hướng dẫn và giao diện, ghi rõ khóa này chỉ dành cho chế độ kiểm thử tự động, bắt buộc cấu hình khóa riêng khi triển khai thực tế.

### D8. Benchmark và Dữ liệu — Tính Đồng nhất và Ranh giới Đo Lường
- **File & Thư mục liên quan:** [`src/benchmark/run_benchmark.py`](src/benchmark/run_benchmark.py), [`results/benchmark_results.json`](results/benchmark_results.json), [`results/benchmark_results.csv`](results/benchmark_results.csv).
- **Vấn đề kiểm tra:**
  - Kiểm tra xem số liệu benchmark có phải từ bài báo Halpin không: **Khẳng định 100% Halpin không có benchmark.** Benchmark này được nhóm thực hiện trên tập dữ liệu Kaggle 2.485 tệp tin (112.81 MB), máy Intel Core i7 12 nhân, Windows 11.
  - Kiểm tra ranh giới đo lường:
    - Trong phép đo xác thực 1 tệp (`verify_single`), Standard Merkle Tree và HD-Merkle chỉ cần đọc tệp mục tiêu và tính nhánh băm $O(\log N)$, trong khi Flat SHA-256 đọc lại tệp và tìm trong manifest.
    - Trong phép đo thêm/xóa tệp (`update_add`, `update_delete`), Standard Merkle Tree bắt buộc phải duyệt và tính lại toàn bộ cây $O(N)$ (mất ~283 ms), trong khi HD-Merkle chỉ cập nhật thư mục cục bộ và lan truyền dirty path $O(\log N)$ (mất ~0.24 ms).
- **Kết luận:** Số liệu benchmark là đóng góp kỹ thuật độc lập phục vụ đồ án môn học, **không đưa vào làm trọng tâm nội dung bài báo giữa kỳ**. Phần trình bày bài báo Halpin phải tập trung vào triết học mật mã học và phân tích đối thủ.

---

## 3. KẾT QUẢ THỰC THI KIỂM THỬ (VERIFICATION RESULTS)

Đã chạy kiểm thử toàn bộ mô-đun an toàn trong môi trường local:
```bash
python -m src.proposed_model.security_games
```
**Kết quả ghi nhận:**
- Thời gian thực thi: **2.885 ms**.
- Bộ dữ liệu: `kaggle_dataset` (2.483 tệp tin).
- Mô hình 1 (PPT Adversary): Tỷ lệ giả mạo thành công quan sát được = 0.0 (phát hiện 100/100 phép thử).
- Mô hình 2 (Dolev-Yao): Ngăn chặn 3/3 kịch bản can thiệp đường truyền (Swap, Tamper, Replay).
- Mô hình 3 (Second-Preimage): Minh họa thành công phân tách miền RFC 6962 loại bỏ va chạm cấu trúc tầng cây.
- Mô hình 4 (Consistency Proof): Xác thực toán học tính nhất quán chỉ-thêm; từ chối mốc lịch sử bị sửa đổi.
- Mô hình 5 (Avalanche & CRC32): SHA-256 đạt trung bình 127.28/256 bit biến đổi (~49.72%, $\sigma = 7.87$ bit); CRC32 bị bẻ gãy 100% bằng 4 byte bù đại số $\mathbb{F}_2$.
- Mô hình 6 (Hierarchical Propagation): Thay đổi ở lá phản ánh lập tức lên trạng thái Root ($O(\log N)$ cập nhật cục bộ).
- Mô hình 7 (Metadata & Timing): Khảo sát độ lệch so sánh chuỗi qua 5 vòng độc lập; `hmac.compare_digest` triệt tiêu chênh lệch thời gian ở mức C-level.
