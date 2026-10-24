# HƯỚNG DẪN THỰC NGHIỆM VÀ DEMO ĐỐI KHÁNG GIỮA KỲ (DT005)
**Đề tài:** Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và Merkle Tree  
**Bài báo đối chiếu:** *The adversary: the philosophy of cryptography* — Harry Halpin, *Journal of Cybersecurity*, Oxford Academic, 2025.  
**Tài liệu:** Hướng dẫn thực nghiệm nội bộ (`MIDTERM_DEMO.md`)

---

## 1. MỤC TIÊU VÀ PHẠM VI CỦA BẢN DEMO

### 1.1. Bản chất khoa học
- **Bài báo tham chiếu (Halpin 2025):** Là công trình triết học và mô hình hóa hình thức về kẻ thù, không có mã nguồn hay số liệu benchmark.
- **Bản demo DT005:** Là **thí nghiệm minh họa độc lập do sinh viên tự xây dựng** nhằm chuyển hóa các luận điểm trừu tượng của Halpin thành một kịch bản trực quan sống động:
  1. *Shannon vs The Adversary (Sec. 1 & 2):* Kẻ thù vượt qua mã tuyến tính CRC32 trong chớp mắt bằng 4 byte bù đại số trên trường $\mathbb{F}_2$; SHA-256 chặn đứng tức thì.
  2. *Dolev-Yao Network Adversary (Sec. 4):* Đối thủ can thiệp đường truyền mạng, tráo đổi bằng chứng băm (Proof Swap) và tấn công phát lại (Replay Attack).
  3. *Node-Leaf Confusion & RFC 6962 Domain Separation (Sec. 5):* Tấn công tiền ảnh cấp hai cấu trúc trên cây ghép thô và giải pháp phân tách miền `0x00`/`0x01`.
  4. *Append-Only Consistency Proof (Sec. 6):* Phát hiện hành vi chèn cửa sau hoặc sửa lén lịch sử theo chuẩn Certificate Transparency.
  5. *Timing Side-Channel (Sec. 5):* Khảo sát độ lệch so sánh chuỗi và giải pháp thời gian hằng số `hmac.compare_digest`.

### 1.2. Điều kiện tiền đề (Preconditions)
- Hệ điều hành: Windows 10/11 hoặc Linux / macOS.
- Môi trường: Python 3.9+ (đã kiểm thử hoàn hảo trên Python 3.14).
- Thư viện phụ thuộc: Chỉ sử dụng thư viện chuẩn của Python (`hashlib`, `hmac`, `zlib`, `json`, `http.server`, `urllib`). Không yêu cầu cài đặt gói mở rộng phức tạp.
- Thư mục dữ liệu: Thư mục `kaggle_dataset` (2.483 tệp tin) đã có sẵn trong workspace.

---

## 2. CÁC PHƯƠNG ÁN DEMO THỰC NGHIỆM

### PHƯƠNG ÁN 1: Chạy Tự Động Bộ Kiểm Định An Ninh Toán Học (Khuyên dùng khi báo cáo ngắn gọn)
Phương án này chạy trực tiếp toàn bộ 7 mô hình kiểm định toán học chỉ trong **~2.8 giây**, xuất báo cáo chi tiết trên terminal:

```bash
python -m src.proposed_model.security_games
```

**Các bước quan sát:**
1. Mở Terminal / PowerShell tại thư mục gốc dự án.
2. Chạy lệnh trên.
3. Quan sát log tiến trình:
   - Thư mục kiểm định được quét và băm mốc ban đầu $H_0$.
   - 7 thí nghiệm chạy tuần tự và in kết quả đo lường thực tế (tỷ lệ đảo bit SAC, giải mã CRC32 $\mathbb{F}_2$, bằng chứng nhất quán RFC 6962, đo thời gian so sánh chuỗi).
   - Tổng thời gian hoàn thành được in ra rõ ràng (~2.885 ms).

---

### PHƯƠNG ÁN 2: Diễn Tập Màn Hình Kép Red Team vs Blue Team (Khuyên dùng cho buổi bảo vệ trực tiếp)
Phương án này thể hiện trực quan kịch bản đối kháng hai phía trên giao diện Web:
- **Blue Team (Phòng thủ - Port 8080):** Giám sát tính toàn vẹn thư mục dữ liệu.
- **Red Team (Tấn công - Port 8081):** Đóng vai kẻ thù phát động 6 vector tấn công.

#### Bước 1: Khởi động Blue Team (Cửa sổ 1)
```bash
run_demo_web.bat
```
*(Hoặc chạy lệnh: `python src/demo/web_demo.py`)*  
Trình duyệt tự động mở: `http://127.0.0.1:8080`.  
- Bấm nút **"Bắt đầu Giám sát Toàn vẹn"**. Hệ thống quét 2.483 tệp và thiết lập mốc tin cậy $H_0$ (trạng thái xanh lá - Toàn vẹn 100%).

#### Bước 2: Khởi động Red Team (Cửa sổ 2)
```bash
run_red_team_web.bat
```
*(Hoặc chạy lệnh: `python src/adversary/red_team_web.py`)*  
Trình duyệt mở giao diện hacker Red Team tại: `http://127.0.0.1:8081`.

#### Bước 3: Diễn tập các đòn tấn công thực chiến
1. **Thử thách 1: Bẻ gãy mã tuyến tính CRC32 (Halpin Sec. 1 & 2):**
   - Trên Red Team (Port 8081), bấm **"Tấn công CRC32 Forge"**.
   - Kẻ tấn công sửa đổi giao dịch chuyển tiền và tự động tính 4 byte bù đại số trên trường $\mathbb{F}_2$.
   - Quan sát log: Mã CRC32 của tệp bị sửa trùng khớp 100% với tệp gốc (`orig_crc == forged_crc`).
   - Sang Blue Team (Port 8080), bấm **"Kiểm tra Toàn vẹn"**: SHA-256 và Merkle Tree lập tức bắt quả tang gian lận!
2. **Thử thách 2: Sửa đổi tệp trên đĩa (In-place Tamper):**
   - Trên Red Team, bấm **"Tamper File"**.
   - Sang Blue Team bấm kiểm tra: Hệ thống cảnh báo đỏ `FILE_MODIFIED`, hiển thị đường dẫn tệp bị sửa và Dirty Path $O(\log N)$ màu cam.
3. **Thử thách 3: Đánh tráo Proof Dolev-Yao (Network MitM):**
   - Trên Red Team, bấm **"Dolev-Yao Proof Swap"**.
   - Sang Blue Team: Bộ xác thực Merkle từ chối chữ ký bằng chứng vì hash nút lá không khớp với gốc $H_0$.

#### Bước 4: Hoàn nguyên an toàn 100% (Rollback All)
- Trên giao diện Red Team (Port 8081), bấm nút đỏ **"KHÔI PHỤC TOÀN BỘ (ROLLBACK ALL)"**.
- Động cơ Kẻ tấn công tự động phục hồi các tệp bị sửa từ thư mục sao lưu an toàn `.adversary_backups/`, xóa sạch tệp cửa sau và tệp tạm.
- Sang Blue Team bấm **"Kiểm tra Toàn vẹn"**: Hệ thống trở về trạng thái toàn vẹn 100% màu xanh!

---

## 3. CÁC HẠN CHẾ CẦN NÊU RÕ KHI THUYẾT TRÌNH (LIMITATIONS)

1. **Về phép đo thời gian kênh kề (Timing Side-Channel):**
   - Phép đo được thực hiện bằng Python trên hệ điều hành máy đơn. Độ lệch quan sát được mang tính minh họa nguyên lý rò rỉ thời gian giữa `early-return` và `constant-time`. Phép đo không chứng minh loại trừ 100% mọi dạng kênh kề phần cứng (cache, vi kiến trúc CPU).
2. **Về tính chất Random Oracle:**
   - Thực nghiệm SAC chỉ chứng minh hàm SHA-256 có đặc tính khuếch tán tuyết lở tốt trên tập mẫu thử, không phải chứng minh tương đương toán học với mô hình Tiên tri Ngẫu nhiên lý tưởng.
3. **Về mô hình Biba:**
   - Hệ thống minh họa sự lan truyền cập nhật băm từ lá lên gốc ($O(\log N)$ Dirty Path), không phải là hệ thống kiểm soát truy cập Biba hoàn chỉnh có bộ giám sát phân quyền Subject/Object.
