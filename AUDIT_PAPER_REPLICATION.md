# BÁO CÁO KIỂM TOÁN VÀ ĐỐI CHIẾU KHOA HỌC: BÀI BÁO GỐC VÀ SLIDE GIỮA KỲ DT005

**Đề tài:** Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và Cây Merkle  
**Môn học:** An toàn và bảo mật hệ thống thông tin (Mã đề tài: DT005)  
**Bài báo nghiên cứu:** *“The adversary: the philosophy of cryptography”* — Harry Halpin, *Journal of Cybersecurity*, Oxford Academic, Volume 11, Issue 1, 2025, tyaf006.  
**DOI:** [10.1093/cybsec/tyaf006](https://doi.org/10.1093/cybsec/tyaf006)  
**Tập tin PDF bài báo:** [`docs/tyaf006.pdf`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/docs/tyaf006.pdf) (17 trang)  
**Tập tin Slide hiện tại được kiểm toán:** [`docs/Bao_Cao_Giua_Ky_DT05.pptx`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/docs/Bao_Cao_Giua_Ky_DT05.pptx) (11 slide)  
**Sinh viên thực hiện:** Võ Anh Kiệt — MSSV: 1150080143 — Lớp: 11_CNPM2  
**Giảng viên hướng dẫn:** ThS. Phạm Trọng Huynh  

---

## MỤC LỤC
1. [Tóm tắt điều hành & Bối cảnh kiểm toán](#1-tóm-tắt-điều-hành--bối-cảnh-kiểm-toán)
2. [Giai đoạn A: Kiểm tra và tóm tắt toàn văn bài báo gốc (tyaf006.pdf)](#2-giai-đoạn-a-kiểm-tra-và-tóm-tắt-toàn-văn-bài-báo-gốc-tyaf006pdf)
3. [Giai đoạn B: Bảng đối chiếu slide hiện tại với bài báo (Slide-by-Slide Audit)](#3-giai-đoạn-b-bảng-đối-chiếu-slide-hiện-tại-với-bài-báo-slide-by-slide-audit)
4. [Giai đoạn C: Kiểm toán repository và xác minh việc triển khai (Paper-to-Code Mapping)](#4-giai-đoạn-c-kiểm-toán-repository-và-xác-minh-việc-triển-khai-paper-to-code-mapping)
5. [Giai đoạn D: Phân loại khả năng tái lập (Replicability Classification A, B, C, D)](#5-giai-đoạn-d-phân-loại-khả-năng-tái-lập-replicability-classification-a-b-c-d)
6. [Danh mục các sai lệch khoa học và kỹ thuật cần sửa chữa](#6-danh-mục-các-sai-lệch-khoa-học-và-kỹ-thuật-cần-sửa-chữa)
7. [Phương án tái cấu trúc slide giữa kỳ (12–15 slide chuẩn học thuật)](#7-phương-án-tái-cấu-trúc-slide-giữa-kỳ-1215-slide-chuẩn-học-thuật)

---

## 1. TÓM TẮT ĐIỀU HÀNH & BỐI CẢNH KIỂM TOÁN

### 1.1. Yêu cầu và Tiêu chuẩn đánh giá của Giảng viên cho Giai đoạn Giữa kỳ
- **Tiêu chuẩn đánh giá chính thức:**
  * **Khung điểm [6 – 8 Điểm]:** Cá nhân chọn đề tài, dịch bài báo, chạy lại kết quả trên tập dataset của bài báo.
  * **Khung điểm [9 – 10 Điểm]:** Cá nhân dịch bài báo, **TỰ XÂY DỰNG MODELS, THUẬT TOÁN VÀ CHẠY LẠI KẾT QUẢ TỪ TẬP DATASET CHO KẾT QUẢ TỐT HƠN!**
- **Quy định thể thức:**
  * Giữa kỳ **chỉ làm Slide báo cáo + Demo kết quả, tuyệt đối không cần làm file Word**.
  * Slide báo cáo giữa kỳ chuẩn: **12–15 slide** (hiện tại thiết kế 14 slide), định dạng 16:9, giao diện sáng sủa, dễ trình bày.
  * Trọng tâm chính (85% thời lượng) tập trung sâu sắc vào bài báo khoa học; phần mô hình/thuật toán tự đề xuất đạt điểm [9-10] được **gói gọn trong đúng 1 trang slide** để giữ trọn vẹn tính học thuật của bài báo.
  * Mã nguồn demo phải bám sát 100% các luận điểm và khái niệm về Kẻ đối đầu (*The Adversary*) của bài báo.

### 1.2. Phát hiện mấu chốt qua kiểm toán
1. **Bản chất bài báo gốc:** Bài báo của Harry Halpin (2025) là một công trình **triết học khoa học và lịch sử mật mã học** (*philosophy and history of cryptography*), tập trung phân tích tiến trình mô hình hóa khái niệm "kẻ đối đầu" (*The Adversary*) từ Shannon (1949), Diffie-Hellman (1976), Dolev-Yao (1983), Goldwasser-Micali (1984), Bellare-Rogaway (1993) đến sự sụp đổ của các mô hình hình thức trước đối thủ thực tế trong các tiết lộ của Edward Snowden về NSA (2013).
2. **Bài báo KHÔNG chứa:**
   - Bất kỳ mã nguồn phần mềm, repository GitHub, thuật toán cụ thể hay giả mã nào để tải về thực thi.
   - Bất kỳ bộ dữ liệu (*dataset*) nào (hoàn toàn không có Kaggle Resume Dataset hay bất kỳ tệp tin nào).
   - Bất kỳ bảng số liệu benchmark định lượng (thời gian mili-giây, dung lượng byte, tốc độ xử lý).
   - Thuật toán Cây Merkle phân cấp (HD-Merkle), kỹ thuật Dirty-Path Cache, cấu trúc nén CompactProof hay chuẩn RFC 6962 như là đóng góp của bài báo.
3. **Thực trạng slide cũ (`Bao_Cao_Giua_Ky_DT05.pptx`):**
   - Chỉ có 11 slide (thiếu so với khung 12–15 slide).
   - Trộn lẫn nghiêm trọng giữa nghiên cứu triết học của Halpin với phần code đồ án cuối kỳ của sinh viên (HD-Merkle, benchmark Kaggle, Ablation study, 1200x–2600x speedup).
   - Tuyên bố sai rằng Halpin hệ thống hóa "7 mô hình kẻ thù" và đề xuất các giải pháp SHA-256, HMAC, Merkle Tree.
   - Đưa ra các con số thực nghiệm của sinh viên (49.72%, Adv = 0.0, 0.01 ms) và tuyên bố "kháng tuyệt đối", "tái lập 100%" mà không làm rõ nguồn gốc.

---

## 2. GIAI ĐOẠN A: KIỂM TRA VÀ TÓM TẮT TOÀN VĂN BÀI BÁO GỐC (tyaf006.pdf)

### 2.1. Thông tin thư mục và xuất bản
- **Tên bài báo:** *The adversary: the philosophy of cryptography*
- **Tác giả:** Harry Halpin (Nym Technologies, Neuchâtel, Thụy Sĩ; liên kết nghiên cứu trước đó với Inria Pháp và MIT).
- **Tạp chí:** *Journal of Cybersecurity*, Volume 11, Issue 1, 2025, tyaf006.
- **Nhà xuất bản:** Oxford University Press (OUP).
- **DOI:** [10.1093/cybsec/tyaf006](https://doi.org/10.1093/cybsec/tyaf006)
- **Độ dài văn bản:** 17 trang PDF (gồm 14 trang nội dung chính và 3 trang tài liệu tham khảo với 134 trích dẫn).

### 2.2. Mục tiêu nghiên cứu và câu hỏi cốt lõi
- **Mục tiêu:** Xây dựng nền tảng triết học cho mật mã học dựa trên thuyết thông tin cổ điển của Fred Dretske (kế thừa Shannon) kết hợp với phê phán về sự trừu tượng hóa (*critique of abstraction*) của Brian Cantwell Smith; khảo cứu lịch sử quá trình chuyển dịch mật mã học từ một "nghệ thuật mang tính hoang tưởng" (*paranoid art*) thành một "khoa học toán học" (*mathematical science*) thông qua việc hình thức hóa đối thủ (*formalizing the adversary*).
- **Câu hỏi cốt lõi:**
  1. *Bản chất của bí mật là gì?* (Bí mật không phải là thuộc tính nội tại của thông tin, mà là một quan hệ tương đối theo thời gian giữa người giữ bí mật, kẻ đối đầu và thế giới bên ngoài).
  2. *Mô hình hóa hình thức về kẻ đối đầu có thực sự bảo đảm an ninh trong thế giới thực không?* (Không; việc trừu tượng hóa luôn bỏ rơi các chi tiết thực tế của tầng triển khai vật lý và bối cảnh chính trị - xã hội).
  3. *Hồ sơ Edward Snowden (2013) về NSA tiết lộ điều gì về giới hạn của mật mã học lý thuyết?* (Kẻ đối đầu cấp quốc gia không bẻ khóa toán học mà tấn công vào tiêu chuẩn, phần mềm, phần cứng, luồng siêu dữ liệu và con người).

### 2.3. Cấu trúc chi tiết theo từng Section và trang PDF

| Section | Tiêu đề | Trang PDF | Luận điểm chính & Bằng chứng học thuật |
| :--- | :--- | :---: | :--- |
| **Sec 1** | *Introduction* | Tr. 1–2 | Mật mã học là truyền thông an toàn khi có sự hiện diện của kẻ đối đầu (*secure communication in the presence of adversaries*). Nỗ lực chuyển mật mã học thành khoa học dựa trên việc hình thức hóa kẻ đối đầu, nhưng các mô hình này thất bại trước các đối thủ thực tế như NSA trong vụ rò rỉ Snowden. |
| **Sec 2** | *Philosophical preliminaries* | Tr. 2–3 | Dựa trên bản thể luận của Brian Cantwell Smith (1996) về dòng chảy siêu hình nguyên thủy (*primordial metaphysical flux*), phân biệt tương tác nhân quả gần (*proximal*) và xa (*distal*). Tiếp thu thuyết thông tin của Fred Dretske (1981): thông tin là tính quy luật của trạng thái sự vật. Tác nhân (*agent*) có tính chủ đích (*intentionality*). |
| **Sec 3** | *Foundations of modern cryptography* | Tr. 3–5 | Phân tích bản chất của bí mật: bí mật mang tính tương đối và hữu hạn theo thời gian (*no secrets are eternal*). Phân biệt *public knowledge* (tri thức công cộng) với *common knowledge* (tri thức phổ thông theo David Lewis). Khảo sát công trình nền tảng của Claude Shannon (1949): *A mathematical theory of communication* & *Communication theory of secrecy systems*. Shannon định nghĩa bí mật hoàn hảo (*perfect secrecy* / One-Time Pad) và bí mật thực tế (*practical secrecy*). Nguyên lý Kerckhoffs: kẻ thù biết toàn bộ thuật toán, chỉ có khóa là bí mật. Mô hình kiểm soát truy cập Bell-LaPadula (1973), Multics, và mô hình toàn vẹn Biba (1977) với nguyên lý *no write-up*. Bốn thuộc tính an toàn: Confidentiality, Integrity, Authentication, Non-repudiation. |
| **Sec 4** | *Formalizing the adversary* | Tr. 5–8 | Cuộc cách mạng khóa công khai: Diffie-Hellman (1976), Merkle (1978), RSA (1978) chuyển gánh nặng từ phân phối khóa vật lý sang giả định độ phức tạp tính toán (bài toán logarit rời rạc, phân tích thừa số nguyên tố). Kẻ đối đầu bị giới hạn trong thời gian đa thức xác suất (PPT Adversary). Goldwasser-Micali (1984): an ninh ngữ nghĩa (*semantic security*), các trò chơi an ninh (*security games*) và quy giảm toán học (*reductions*). Mô hình mạng Dolev-Yao (1983): kẻ thù tượng trưng kiểm soát toàn bộ đường truyền (*the adversary is the wire*), giả định mật mã học hoàn hảo (*perfect cryptography assumption*). Phương pháp kiểm chứng giao thức tự động: Logic BAN (1989), CSP/FDR, định lý Lowe (1996) bẻ gãy Needham-Schroeder sau 17 năm; chứng minh giao thức TLS bằng Isabelle (Paulson 1999). Bellare-Rogaway (1993): an ninh chứng minh được hướng thực hành (*practice-oriented provable security*) và mô hình tiên tri ngẫu nhiên (Random Oracle Model - ROM). |
| **Sec 5** | *The limits of formalization in cryptography* | Tr. 8–10 | Phê phán của Brian Cantwell Smith (1985) về *The limits of correctness in computers*: khoảng cách không thể san lấp giữa mô hình toán trừu tượng và thực tế vật lý. Koblitz & Menezes phê phán tính chân thực của provable security. Các đòn tấn công kênh kề (*side-channel attacks*): tấn công đo thời gian của Paul Kocher (1996), tấn công bộ đệm cache, rò rỉ điện năng. Giả định an ninh ngữ nghĩa cho phép độ dài bản mã công khai dẫn tới các cuộc tấn công thực tế như CRIME và BREACH trên TLS do cơ chế nén dữ liệu (Rizzo & Duong 2012, Gluck et al. 2013). Luận điểm nổi tiếng của Ross Anderson (1993) *Why Cryptosystems Fail*: hệ thống thất bại do lỗi lập trình và quản lý vận hành chứ không phải do toán học bị bẻ gãy. |
| **Sec 6** | *The adversary revealed* | Tr. 10–12 | Hồ sơ Edward Snowden (2013) phơi bày năng lực thực tế của NSA (kẻ đối đầu mạnh nhất lịch sử). NSA không tập trung giải các bài toán NP-hard mà: (1) Cài cửa sau vào tiêu chuẩn mật mã quốc gia (chương trình BULLRUN, thuật toán Dual_EC_DRBG với ngân sách 250 triệu USD); (2) Thu thập lưu lượng cáp quang diện rộng trước khi mã hóa (chương trình MUSCULAR, PRISM, XKeyscore); (3) Tấn công điểm cuối và can thiệp phần cứng chuỗi cung ứng; (4) Khai thác phân tích siêu dữ liệu (*metadata surveillance*). Mối đe dọa điện toán lượng tử (thuật toán Shor 1994) và sự hoài nghi đối với tiêu chuẩn Post-Quantum Cryptography (PQC) của NIST/NSA. |
| **Sec 7** | *Conclusion: a research agenda* | Tr. 12–14 | Đề xuất chương trình nghiên cứu mới vượt qua giới hạn của trừu tượng hóa: (1) Mở rộng mô hình truyền thông sang mạng xã hội - kỹ thuật không đồng nhất (*heterogeneous techno-social networks*); (2) Bảo vệ siêu dữ liệu và chống phân tích lưu lượng thông qua mạng ẩn danh mixnet (David Chaum 1981, Nym network - Diaz & Halpin); (3) Tích hợp dữ liệu thực nghiệm và các ngành khoa học xã hội/kinh tế an ninh (Ross Anderson 2001) thay vì chỉ giam mình trong chứng minh toán học thuần túy. |

---

## 3. GIAI ĐOẠN B: BẢNG ĐỐI CHIẾU SLIDE HIỆN TẠI VỚI BÀI BÁO (SLIDE-BY-SLIDE AUDIT)

Dưới đây là bảng kiểm toán chi tiết từng slide trong file [`docs/Bao_Cao_Giua_Ky_DT05.pptx`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/docs/Bao_Cao_Giua_Ky_DT05.pptx):

| Slide | Nội dung đang trình bày | Căn cứ trong bài báo (Section / Trang) | Mức độ chính xác | Nội dung bị nhầm với nghiên cứu riêng | Hành động khuyến nghị |
| :---: | :--- | :--- | :--- | :--- | :--- |
| **1** | Bìa: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động. GVHD: ThS. Phạm Trọng Huynh, SV: Võ Anh Kiệt. | Không có trong bài báo. | Đúng thông tin hành chính; sai tiêu đề hướng đề tài giữa kỳ. | Tiêu đề nghiêng 100% về giải pháp kỹ thuật riêng cuối kỳ thay vì bài báo được giao. | **SỬA:** Điều chỉnh tiêu đề thể hiện rõ: "Báo cáo giữa kỳ: Nghiên cứu bài báo khoa học Harry Halpin (2025) và hiện thực hóa thực nghiệm minh họa kiểm tra toàn vẹn tệp". Giữ nguyên thông tin GVHD và sinh viên. |
| **2** | Tóm tắt Halpin (2025): Liệt kê "7 mô hình kẻ thù" và "5 hạn chế của mô hình hóa hình thức". | Sec 1–6 (Trang 1–12). | **ĐÚNG MỘT PHẦN** (Nội dung khái niệm có xuất hiện nhưng cách đánh số 1–7 và 1–5 là do sinh viên tự tổng hợp). | Tự đóng khung thành danh sách "7 mô hình kẻ thù" và "5 hạn chế" khiến người nghe lầm tưởng Halpin lập bảng danh mục này. Đặc biệt mục 5 ("Kẻ thù va chạm tiền ảnh thứ hai") và mục 6 ("Biba") được gán ghép để khớp với code DT005. | **SỬA:** Trình bày đúng cấu trúc bài báo: Quá trình tiến hóa lịch sử của mô hình hóa đối thủ (Shannon $\to$ Diffie-Hellman $\to$ Dolev-Yao $\to$ Provable Security $\to$ NSA/Snowden). Ghi rõ "Tổng hợp phân loại bởi nhóm thực hiện". |
| **3** | Bộ dữ liệu Kaggle Resume Dataset (2.485 tệp, 118.29 MB, 24 thư mục nghề nghiệp). | **HOÀN TOÀN KHÔNG CÓ** trong bài báo (0 lần xuất hiện). | **SAI NGUỒN GỐC** nếu để trong phần nghiên cứu bài báo. | Toàn bộ tập dữ liệu Kaggle là do sinh viên tự tải về làm môi trường thực nghiệm kiểm thử toàn vẹn tệp tin. | **CHUYỂN:** Đưa sang phần Thiết kế thực nghiệm minh họa của sinh viên (Stage B/C). Nêu rõ: Bài báo không có dataset; sinh viên tự chọn Kaggle Resume Dataset để mô phỏng tải tệp thực tế. |
| **4** | Cơ chế SHA-256, HMAC-SHA256, Standard Merkle Tree đối phó Kerckhoffs, Shannon, Dolev-Yao. Trích dẫn trang 1, 3, 5, 8, 13. | Trang 3–5, 8 (nhắc Kerckhoffs, Shannon, Dolev-Yao). Merkle Tree không có trong bài báo. | **ĐÚNG MỘT PHẦN** (Các khái niệm lý thuyết có thật, nhưng việc dùng SHA-256/HMAC/Merkle là do sinh viên tự ánh xạ). | Khẳng định "Halpin trang 8 em ánh xạ sang cây Merkle" và "đáp ứng lời kêu gọi của Halpin" dễ gây ngộ nhận Halpin đề xuất kiến trúc này. | **SỬA:** Định nghĩa rõ: Đây là các "Mô hình cơ sở (Baselines) do sinh viên tự xây dựng nhằm minh họa thực nghiệm các luận điểm lý thuyết của Halpin". Không gọi là mô hình của Halpin. |
| **5** | 7 trò chơi an ninh hình thức đạt kết quả tuyệt đối: Adv = 0.0 $\le 2^{-128}$, RFC 6962, tuyết lở 49.72%, triệt tiêu Kocher. Trích dẫn trang 5–10. | Trang 6–10 (bàn về PPT, ROM, Kocher, CRIME/BREACH). Không có RFC 6962, không có 49.72%, không có Adv = 0.0. | **SAI BẢN CHẤT KHOA HỌC.** | Lấy kết quả chạy 100 lần thử trong code Python của sinh viên rồi tuyên bố "chứng minh thực nghiệm Adv $\le 2^{-128}$" và "kháng tuyệt đối". Halpin bài báo bàn về triết học phản bác sự an toàn tuyệt đối. | **SỬA:** Bỏ hoàn toàn các phát biểu "chứng minh an toàn tuyệt đối", "Adv = 0.0". Diễn đạt đúng: Đây là "Bộ kiểm thử tự động (Security Test Suite) do sinh viên thiết kế nhằm kiểm tra các vector tấn công đối kháng mô phỏng". |
| **6** | Đổi mới sáng tạo: Cây Merkle phân cấp động (HD-Merkle), cập nhật cục bộ $O(\log N)$, RFC 6962. | **KHÔNG CÓ** trong bài báo. | **SAI PHẠM VI GIỮA KỲ.** | Đây là giải pháp đề xuất riêng của sinh viên (phần cuối kỳ), không thuộc nội dung báo cáo tiến độ bài báo giữa kỳ. | **LOẠI BỎ KHỎI TRỌNG TÂM:** Không đưa cấu trúc HD-Merkle làm trọng tâm giữa kỳ. Chỉ nhắc ngắn gọn như một hướng phát triển thực nghiệm cho giai đoạn cuối kỳ. |
| **7** | Bằng chứng nén CompactProof 256 B (tiết kiệm 99.9% băng thông, 0.01 ms), chống sửa đổi lịch sử. | **KHÔNG CÓ** trong bài báo. | **SAI PHẠM VI GIỮA KỲ.** | Đóng góp riêng của sinh viên cho cuối kỳ. | **LOẠI BỎ KHỎI TRỌNG TÂM:** Loại bỏ các tuyên bố về CompactProof và tối ưu nén bằng chứng khỏi slide giữa kỳ. |
| **8** | Bảng đối sánh Big-O và Benchmark thực nghiệm trên Kaggle: 0.24 ms vs 283.9 ms (nhanh hơn 1.200–2.600 lần). | **KHÔNG CÓ** trong bài báo. | **SAI PHẠM VI GIỮA KỲ.** | Số liệu đo đạc riêng của sinh viên giữa thuật toán tự đề xuất và các baseline. | **LOẠI BỎ KHỎI TRỌNG TÂM:** Bỏ bảng benchmark 1.200x–2.600x. Thay bằng bảng so sánh đặc tính an ninh của các cơ chế toàn vẹn (Flat Hash vs HMAC vs Merkle) trước các lớp đối thủ trong bài báo. |
| **9** | Ablation Study 4 cấu hình đo trên 2.485 tệp tin; Phân tích đánh đổi (Trade-offs). | **KHÔNG CÓ** trong bài báo. | **SAI PHẠM VI GIỮA KỲ.** | Đóng góp kỹ thuật riêng của sinh viên cho báo cáo cuối kỳ. | **LOẠI BỎ KHỎI TRỌNG TÂM:** Bỏ toàn bộ slide Ablation Study. Giữa kỳ tập trung vào việc đối chiếu bài báo và kiểm chứng tính toàn vẹn. |
| **10** | Slide ghi vỏn vẹn chữ "DEMO". | Không có nội dung. | Thiếu thông tin thuyết trình. | Chưa nêu rõ kịch bản demo, mục tiêu minh họa và điều kiện kiểm thử. | **SỬA:** Thay bằng slide "Kế hoạch & Kịch bản Demo thực nghiệm", nêu rõ các bước đối kháng giữa Red Team (Adversary) và Blue Team (Integrity Verifier). |
| **11** | Lời cảm ơn. | Kết thúc. | Đạt yêu cầu thể thức. | Không có. | **GIỮ:** Giữ nguyên ở vị trí kết thúc bài thuyết trình (Slide 14/14). |

---

## 4. GIAI ĐOẠN C: KIỂM TOÁN REPOSITORY VÀ XÁC MINH VIỆC TRIỂN KHAI (PAPER-TO-CODE MAPPING)

### 4.1. Cấu trúc mã nguồn thực tế trong workspace
- [`src/baselines/sha256_manifest.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/sha256_manifest.py): Triển khai bảng băm phẳng SHA-256 (Baseline 1).
- [`src/baselines/hmac_manifest.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/hmac_manifest.py): Triển khai mã xác thực thông điệp HMAC-SHA256 có khóa đối xứng (Baseline 2).
- [`src/baselines/standard_merkle_tree.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/standard_merkle_tree.py): Triển khai cây Merkle nhị phân truyền thống ghép cặp $H(Left \,\|\, Right)$ (Baseline 3).
- [`src/proposed_model/security_games.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py): Bộ kịch bản kiểm định an ninh toán học mô phỏng 7 thách đấu đối kháng.
- [`src/proposed_model/hd_merkle_tree.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/hd_merkle_tree.py): Mô hình cây Merkle phân cấp động kết hợp phân tách miền RFC 6962.
- [`src/adversary/adversary_engine.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/adversary/adversary_engine.py): Động cơ mô phỏng các đòn tấn công của kẻ đối đầu (tamper, inject, delete, CRC32 forge, Dolev-Yao proof swap).
- [`src/adversary/red_team_cli.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/adversary/red_team_cli.py): Giao diện dòng lệnh độc lập cho Red Team.
- [`src/benchmark/run_benchmark.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/benchmark/run_benchmark.py): Script đo lường hiệu năng 4 mô hình trên tập dữ liệu Kaggle.

### 4.2. Bảng ánh xạ từ Luận điểm bài báo sang Hiện thực hóa Code (Paper-to-Code Mapping)

| # | Luận điểm / Khái niệm trong bài báo Halpin (2025) | Căn cứ cụ thể trong bài báo | Thành phần code tương ứng trong Repository | Thí nghiệm / Kiểm thử thực tế đã chạy | Kết quả quan sát được | Mức độ tương ứng | Kết luận trình bày khoa học trung thực trên slide |
| :-: | :--- | :--- | :--- | :--- | :--- | :---: | :--- |
| **1** | **Mã kiểm tra lỗi tuyến tính vs Kẻ đối đầu có chủ đích:** Mã kiểm tra lỗi như CRC/parity chỉ chống được nhiễu ngẫu nhiên của Shannon; kẻ đối đầu có chủ đích dễ dàng qua mặt. | Sec 1 & 2 (Trang 1–4) trích dẫn Shannon (1948, 1949). | Hàm `crc32_forge_patch` trong [`security_games.py:L60-L107`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py#L60-L107). | Giải hệ phương trình đại số nhị phân trên trường $\mathbb{F}_2$ để tìm 4 byte bù patch cho thông điệp giả mạo. | Kẻ tấn công tạo ra thông điệp giả có mã CRC32 trùng 100% với thông điệp gốc; SHA-256 phát hiện vi phạm ngay lập tức. | **Minh họa gần đúng** (Minh họa luận điểm lý thuyết bằng bài toán đại số thực tế). | Trình bày như một ví dụ thực nghiệm của sinh viên chứng minh sự cần thiết của hàm băm mật mã học so với mã kiểm tra lỗi cổ điển. |
| **2** | **Nguyên lý Kerckhoffs & Hàm băm công khai:** Độ an toàn của hệ thống không phụ thuộc vào việc giấu kín thuật toán, mà phụ thuộc vào tính bảo mật của khóa hoặc các giả định toán học công khai. | Sec 3 (Trang 3–4), dẫn chiếu Kerckhoffs (1883) & Shannon (1949). | [`SHA256Manifest`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/sha256_manifest.py). | Tính toán manifest SHA-256 công khai trên tập tệp tin và xác thực khi có tệp bị sửa đổi. | Phát hiện 100% tệp bị sửa đổi nếu manifest được lưu giữ an toàn; nhưng nếu kẻ tấn công ghi đè được cả manifest thì hệ thống bị qua mặt. | **Minh họa độc lập** (Sinh viên tự áp dụng để làm baseline kiểm tra toàn vẹn). | Trình bày đây là Baseline 1 minh họa cho việc kiểm tra toàn vẹn không khóa; chỉ an toàn khi danh mục băm được bảo vệ toàn vẹn tuyệt đối. |
| **3** | **Kẻ nghe lén Shannon & Điểm yếu phân phối khóa đối xứng:** Mật mã khóa đối xứng của Shannon yêu cầu chia sẻ khóa an toàn trước; rủi ro nằm ở kênh quản lý khóa. | Sec 3 (Trang 4–5), dẫn chiếu Shannon (1949). | [`HMACManifest`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/hmac_manifest.py). | Ký manifest bằng HMAC-SHA256 với khóa bí mật $K$; mô phỏng tình huống client bị lộ khóa trong `DolevYaoNetworkGame`. | Nếu máy khách giữ khóa bị lộ, kẻ đối đầu có thể tự tạo lại toàn bộ chữ ký HMAC cho dữ liệu độc hại. | **Minh họa độc lập**. | Trình bày đây là Baseline 2; giải thích rõ ưu điểm chống sửa đổi manifest của HMAC nhưng chỉ ra điểm yếu phụ thuộc vào bí mật khóa chia sẻ. |
| **4** | **Kẻ đối đầu mạng Dolev-Yao & Bằng chứng xác thực độc lập:** Kẻ đối đầu kiểm soát hoàn toàn kênh truyền, có thể nghe, chặn, sửa, tráo đổi gói tin. | Sec 4 (Trang 7–8), dẫn chiếu Dolev & Yao (1983). | [`StandardMerkleTree`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/baselines/standard_merkle_tree.py) & `DolevYaoNetworkGame`. | Mô phỏng 3 kịch bản tấn công đường truyền: tráo đổi bằng chứng giữa 2 file (Cross-proof swap), sửa payload, phát lại proof cũ. | Root Hash công khai 32 byte giúp phát hiện 100% bằng chứng bị tráo đổi mà không cần chia sẻ khóa bí mật cho từng client. | **Minh họa gần đúng**. | Trình bày cây Merkle là công cụ cấu trúc dữ liệu kinh điển giúp xác thực trong mô hình mạng thù địch Dolev-Yao với chi phí $O(\log N)$. |
| **5** | **Giới hạn của mô hình hình thức: Va chạm tiền ảnh thứ hai (Second-Preimage Attack):** Mô hình trừu tượng coi phép ghép chuỗi là hoàn hảo, nhưng trong thực tế ghép thô $H(Left \,\|\, Right)$ có thể bị nhầm lẫn nút lá và nút nhánh. | Sec 5 (Trang 8–10), Halpin phân tích về *limits of formalization* & abstraction. | `SecondPreimageGame` trong [`security_games.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py) so sánh Standard Merkle vs RFC 6962. | Tạo một nút giả mạo 64-byte đóng vai trò là cặp ghép của 2 nút con để kiểm tra xem cây có chấp nhận không. | Cây Merkle chuẩn không phân tách miền chấp nhận lá giả; cây có tiền tố `0x00`/`0x01` loại bỏ hoàn toàn khả năng này. | **Minh họa mở rộng**. | Nêu rõ đây là thí nghiệm minh họa cho luận điểm của Halpin về "sự thất bại của các giả định trừu tượng hóa khi hiện thực hóa mã nguồn". |
| **6** | **Tấn công kênh kề đo thời gian (Timing Side-Channel):** Mô hình toán học giả định phép so sánh chuỗi tốn thời gian cố định; trong thực tế so sánh chuỗi ngắt sớm (*early return*) làm rò rỉ thông tin khóa/băm. | Sec 5 (Trang 8, 10), dẫn chiếu Paul Kocher (1996). | `audit_timing_resistance` trong [`security_games.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py). | Đo thời gian 20.000 chu kỳ so sánh chuỗi bằng toán tử thông thường `==` so với hàm `hmac.compare_digest`. | Toán tử `==` lộ chênh lệch thời gian tỷ lệ thuận với số byte trùng khớp; `compare_digest` triệt tiêu chênh lệch này. | **Minh họa thực nghiệm**. | Minh họa xuất sắc cho luận điểm của Halpin về việc kẻ đối đầu thực tế khai thác các đặc tính vật lý nằm ngoài mô hình chứng minh toán học. |
| **7** | **Kẻ đối đầu thực tế (Snowden / NSA) & Giám sát siêu dữ liệu:** Kẻ đối đầu cấp quốc gia tập trung phân tích lưu lượng, siêu dữ liệu và chèn cửa sau lịch sử thay vì phá khóa. | Sec 6 (Trang 10–12), dẫn chiếu Snowden (2013). | `ConsistencyProofGame` & `MetadataLeakageAudit` trong [`security_games.py`](file:///d:/AnToanvaBaoMatHeThongThongTin/GK/src/proposed_model/security_games.py). | Đo lường rò rỉ dung lượng danh mục tệp và kiểm tra bằng chứng nhất quán chỉ thêm (Append-only Consistency). | Bằng chứng Merkle $O(\log N)$ chỉ làm lộ 1 đường dẫn xác minh thay vì gửi toàn bộ danh bạ tệp của hệ thống cho client. | **Minh họa mở rộng**. | Minh họa luận điểm của Halpin: bảo vệ tính toàn vẹn phải đi đôi với giảm thiểu rò rỉ siêu dữ liệu (*metadata minimization*). |

---

## 5. GIAI ĐOẠN D: PHÂN LOẠI KHẢ NĂNG TÁI LẬP (REPLICABILITY CLASSIFICATION A, B, C, D)

Căn cứ vào toàn văn bài báo và mã nguồn thực tế, mọi nội dung được phân loại nghiêm ngặt thành 4 nhóm:

```mermaid
pie title Phân loại khả năng tái lập đề tài DT005
    "Nhóm A: Tái lập trực tiếp từ bài báo (0%)" : 0
    "Nhóm B: Thí nghiệm minh họa luận điểm (50%)" : 50
    "Nhóm C: Đóng góp kỹ thuật riêng / Cuối kỳ (35%)" : 35
    "Nhóm D: Không thể tái lập do thiếu dữ liệu gốc (15%)" : 15
```

### Nhóm A: Có thể tái lập trực tiếp từ bài báo (Directly Replicable)
- **Tình trạng:** **KHÔNG CÓ (0%)**.
- **Giải thích khoa học:** Bài báo của Harry Halpin (2025) là công trình nghiên cứu triết học, bản thể luận và lịch sử mô hình hóa an ninh thông tin. Bài báo **hoàn toàn không cung cấp** bất kỳ thuật toán phần mềm, giả mã, giao thức cụ thể, bộ dữ liệu thực nghiệm hay số liệu định lượng nào để người đọc chạy lại nguyên văn. Do đó, tuyên bố "tái lập trực tiếp 100% kết quả của tác giả" là hoàn toàn sai sự thật về mặt khoa học.

### Nhóm B: Thí nghiệm do sinh viên thiết kế để kiểm chứng / minh họa luận điểm của bài báo (Student Verification Experiments)
- **Tình trạng:** **CÓ THỂ THỰC HIỆN VÀ ĐÃ HOÀN TẤT MINH HỌA TRUNG THỰC**.
- **Các thành phần thuộc nhóm này:**
  1. Thí nghiệm so sánh mã kiểm tra lỗi CRC32 (bị bẻ gãy bằng bù đại số GF(2)) với hàm băm SHA-256 $\to$ minh họa luận điểm Shannon vs Kẻ đối đầu có chủ đích (Sec 1 & 2).
  2. Baseline Flat SHA-256 Manifest $\to$ minh họa nguyên lý Kerckhoffs (Sec 3).
  3. Baseline HMAC-SHA256 Manifest $\to$ minh họa cơ chế xác thực có khóa và rủi ro kênh phân phối khóa (Sec 3).
  4. Baseline Standard Merkle Tree $\to$ minh họa cơ chế xác thực độc lập chống kẻ thù mạng Dolev-Yao (Sec 4).
  5. Thí nghiệm va chạm tiền ảnh thứ hai trên Merkle Tree không phân tách miền $\to$ minh họa luận điểm về giới hạn của sự trừu tượng hóa hình thức (Sec 5).
  6. Thí nghiệm đo thời gian so sánh chuỗi (Early Return vs Constant-Time) $\to$ minh họa tấn công kênh kề đo thời gian theo Paul Kocher (Sec 5).
- **Yêu cầu khi trình bày:** Phải nêu rõ đây là các thí nghiệm mô phỏng do sinh viên tự thiết kế để minh họa và làm sáng tỏ các luận điểm lý thuyết của Halpin, không phải số liệu do tác giả công bố.

### Nhóm C: Đóng góp Thuật toán tự xây dựng của Đề tài DT005 (Đạt Tiêu chuẩn [9-10 Điểm])
- **Tình trạng:** **ĐÃ HIỆN THỰC HÓA VÀ GÓI GỌN TRONG 1 TRANG SLIDE DUY NHẤT (SLIDE 12)**.
- **Căn cứ tiêu chí đánh giá [9-10 điểm]:** Giảng viên yêu cầu cá nhân sinh viên phải: *"dịch bài báo, tự xây dựng models, thuật toán và chạy lại kết quả từ tập dataset cho kết quả tốt hơn"*.
- **Các thành phần thuộc nhóm này:**
  1. Thuật toán cây Merkle phân cấp động (HD-MerkleTree) ánh xạ cấu trúc cây thư mục doanh nghiệp.
  2. Kỹ thuật cập nhật cục bộ Dirty-Path Cache đạt $O(\log N)$ ~0.24 ms so với Rebuild $O(N)$ ~283.9 ms của Standard Merkle Tree (tăng tốc ~1.200 lần).
  3. Phân tách miền RFC 6962 triệt tiêu 100% va chạm tiền ảnh thứ hai nút lá/nhánh.
  4. Cấu trúc nén bằng chứng xác minh CompactProof (256 Bytes).
- **Quyết định học thuật:** Trình bày ngắn gọn, súc tích trong đúng 1 slide (Slide 12) để chứng minh năng lực tự xây dựng thuật toán cho kết quả tốt hơn đạt điểm [9,10] mà không làm phân tán trọng tâm bài báo.

### Nhóm D: Không thể tái lập từ tài liệu hiện có (Irreproducible)
- **Tình trạng:**
  - Không thể tái lập các bằng chứng thực nghiệm về các chiến dịch nội bộ của NSA (BULLRUN, MUSCULAR, PRISM, QUANTUMINSERT) vì đây là các tài liệu tình báo mật được rò rỉ qua báo chí, không phải môi trường mở mã nguồn để chạy benchmark.
  - Không thể tái lập các kiểm chứng toán học bằng chứng minh tự động Isabelle của Paulson (1999) trên TLS hay phần mềm kiểm chứng CSP/FDR của Lowe (1996) do nằm ngoài phạm vi công cụ của đề tài.

---

## 6. DANH MỤC CÁC SAI LỆCH KHOA HỌC VÀ KỸ THUẬT CẦN SỬA CHỮA

Qua đối chiếu giữa Slide cũ (`docs/Bao_Cao_Giua_Ky_DT05.pptx`), bài báo gốc và source code, các lỗi sau đây bắt buộc phải sửa triệt để:

| STT | Vấn đề sai lệch phát hiện trong Slide cũ | Bản chất sai sót | Cách khắc phục chuẩn xác trong Slide mới |
| :-: | :--- | :--- | :--- |
| **1** | Ghi "Bảy mô hình kẻ thù được Halpin hệ thống hóa: 1... 2... 3... 4... 5... 6... 7..." | Halpin không hề đưa ra danh mục đánh số này. Sinh viên tự tổng hợp và gán ghép các khái niệm từ nhiều trang khác nhau để khớp với code của mình. | Trình bày theo tiến trình lịch sử phát triển của các mô hình đối thủ mà Halpin phân tích: Shannon (nhiễu ngẫu nhiên) $\to$ Access Control (Biba/Bell-LaPadula) $\to$ Bounded Computation (PPT/Diffie-Hellman) $\to$ Network Adversary (Dolev-Yao) $\to$ Side-Channels (Kocher) $\to$ Institutional Adversary (NSA). Nêu rõ "Phân loại tổng hợp bởi người báo cáo". |
| **2** | Ghi "5 Hạn chế của mô hình hóa hình thức trong bài báo cần giải quyết" | Bài báo bàn luận sâu rộng về phê phán tính trừu tượng của Cantwell Smith và các vụ tấn công thực tế; không đóng khung thành danh sách 5 gạch đầu dòng cứng nhắc. | Trình bày các luận điểm cốt lõi của Halpin về giới hạn của mô hình hóa: (1) Khoảng cách giữa trừu tượng hóa toán học và hiện thực vật lý; (2) Tấn công kênh kề và rò rỉ độ dài bản mã; (3) Tấn công vào chuỗi cung ứng và tiêu chuẩn hóa; (4) Khai thác lưu lượng và siêu dữ liệu. |
| **3** | Đưa thông tin bộ dữ liệu Kaggle Resume Dataset (2.485 tệp, 118.29 MB) vào Slide 3 như thể là dữ liệu của bài báo | Kaggle Resume Dataset là bộ dữ liệu sinh viên tự lấy trên mạng để thử nghiệm; bài báo của Halpin hoàn toàn không có dataset này. | Chuyển hẳn nội dung này sang phần "Môi trường thực nghiệm minh họa của sinh viên". Khẳng định trung thực: Bài báo gốc không có dataset; sinh viên tự chọn Kaggle Resume Dataset để mô phỏng khối lượng dữ liệu thực tế. |
| **4** | Tuyên bố: "Halpin trang 8 em ánh xạ sang cây Merkle; Halpin trang 1, 13 kêu gọi em hiện thực hóa hệ thống phần mềm" | Gán ghép chủ quan, ngụy biện thẩm quyền. Halpin không đề xuất cây Merkle và không kêu gọi sinh viên viết tool kiểm tra toàn vẹn tệp. | Phát biểu khách quan: "Kế thừa các luận điểm lý thuyết của Halpin về yêu cầu xác thực trong môi trường mạng thù địch và tinh thần kiểm nghiệm thực tế, đề tài DT005 xây dựng hệ thống kiểm tra toàn vẹn tệp bằng các cấu trúc kinh điển (Flat Hash, HMAC, Merkle Tree) làm mô hình minh họa". |
| **5** | Khẳng định: "Adv = 0.0 $\le 2^{-128}$", "Kháng tuyệt đối", "Triệt tiêu 100% nguy cơ", "Tái lập 100%" | Phản khoa học. Không có thử nghiệm hữu hạn nào chứng minh được lợi thế tấn công $\le 2^{-128}$. Bài báo của Halpin nhấn mạnh không có an toàn tuyệt đối ("no secret can be absolute"). | Loại bỏ hoàn toàn các từ "kháng tuyệt đối", "Adv = 0.0", "triệt tiêu 100%". Thay bằng: "Vượt qua toàn bộ các ca kiểm thử đối kháng mô phỏng trong tập thử nghiệm hữu hạn; tuân thủ các giả định an toàn chuẩn của SHA-256". |
| **6** | Trình bày HD-Merkle, Dirty-Path Cache, CompactProof, Benchmark 1.200x–2.600x, Ablation Study chiếm 4 slide chính (Slide 6, 7, 8, 9) | Lạc đề so với yêu cầu giữa kỳ của giảng viên ("Trình bày bài báo, chạy lại kết quả; chưa yêu cầu phát triển đóng góp mới hay tối ưu thuật toán"). | Loại bỏ các slide trên khỏi trọng tâm giữa kỳ. Dành không gian slide để phân tích sâu sắc bài báo của Halpin, phân tích khả năng tái lập và trình bày demo minh họa đối kháng trung thực. Đóng góp kỹ thuật HD-Merkle chỉ tóm tắt trong 1 slide định hướng cuối kỳ. |

---

## 7. PHƯƠNG ÁN TÁI CẤU TRÚC SLIDE GIỮA KỲ (12–15 SLIDE CHUẨN HỌC THUẬT)

Bản slide mới được thiết kế chuẩn mực với **14 trang (16:9)**, giao diện sáng (Light Theme), kiểu chữ hiện đại, bố cục dạng thẻ (cards) thoáng đãng, phân tách màu sắc rành mạch giữa **Lý thuyết bài báo (Navy/Blue)** và **Minh họa thực nghiệm DT005 (Teal/Emerald)**:

- **Slide 1: Trang bìa báo cáo giữa kỳ**
  - Tiêu đề đề tài DT005 & Bài báo tham chiếu Harry Halpin (2025).
  - Thông tin xuất bản: *Journal of Cybersecurity*, Oxford Academic, DOI: 10.1093/cybsec/tyaf006.
  - Thông tin SV thực hiện, MSSV, Lớp, GVHD, Học phần.
- **Slide 2: Bối cảnh nghiên cứu & Ý nghĩa của bài báo**
  - Sự chuyển dịch của mật mã học từ nghệ thuật sang khoa học thông qua việc hình thức hóa đối thủ.
  - Vấn đề cốt lõi: Tại sao các hệ thống mật mã đã được chứng minh an toàn toán học vẫn sụp đổ trong thế giới thực?
- **Slide 3: Mục tiêu & Câu hỏi nghiên cứu của tác giả**
  - Câu hỏi 1: Bản chất của bí mật và thông tin là gì?
  - Câu hỏi 2: Giới hạn của sự trừu tượng hóa hình thức nằm ở đâu?
  - Câu hỏi 3: Hồ sơ rò rỉ Snowden phơi bày những gì về kẻ đối đầu thực tế?
- **Slide 4: Cơ sở triết học: Dretske, Shannon & Bản chất của Bí mật**
  - Thuyết thông tin Fred Dretske kế thừa Claude Shannon.
  - Định nghĩa bí mật: Mối quan hệ tương đối theo thời gian, không có bí mật vĩnh cửu (*no secrets are eternal*).
  - Phân biệt Tri thức công cộng (*Public knowledge*) và Tri thức phổ thông (*Common knowledge*).
- **Slide 5: Tiến trình mô hình hóa kẻ đối đầu (Giai đoạn 1: 1949 – 1980s)**
  - Claude Shannon (1949): Kẻ nghe lén thụ động, bí mật hoàn hảo vs bí mật thực tế, Nguyên lý Kerckhoffs.
  - Kiểm soát truy cập: Mô hình Bell-LaPadula (Bảo mật) & Biba (Toàn vẹn - *No Write-Up*).
  - Khóa công khai: Diffie-Hellman & Merkle (1976–1978), bài toán độ phức tạp tính toán và kẻ thù thời gian đa thức (PPT Adversary).
- **Slide 6: Tiến trình mô hình hóa kẻ đối đầu (Giai đoạn 2: 1980s – 2000s)**
  - Trò chơi an ninh & Quy giảm toán học (Goldwasser-Micali 1984, an ninh ngữ nghĩa IND-CPA/CCA).
  - Kẻ đối đầu mạng Dolev-Yao (1983): Kẻ thù kiểm soát toàn bộ kênh truyền, kiểm chứng giao thức tự động (BAN, CSP, Isabelle).
  - An ninh chứng minh được hướng thực hành (Bellare-Rogaway 1993) & Mô hình Tiên tri Ngẫu nhiên (ROM).
- **Slide 7: Giới hạn của sự hình thức hóa (The Limits of Formalization)**
  - Phê phán của Brian Cantwell Smith (1985): Sự bất đối xứng giữa mô hình toán trừu tượng và thực tế vật lý.
  - Tấn công kênh kề: Tấn công đo thời gian của Paul Kocher (1996), rò rỉ bộ nhớ cache và năng lượng.
  - Tấn công rò rỉ độ dài bản mã: CRIME và BREACH trên giao thức TLS do cơ chế nén dữ liệu.
  - Ross Anderson (1993): Hệ thống thất bại ở tầng triển khai, quản lý và con người, không phải do bẻ gãy thuật toán.
- **Slide 8: Kẻ đối đầu thực tế lộ diện: Bài học từ Edward Snowden & NSA**
  - Kẻ đối đầu mạnh nhất lịch sử: NSA và các cơ quan tình báo quốc gia.
  - Chiến lược thực tế của NSA:
    * Can thiệp chuẩn hóa và cài cửa sau (BULLRUN, Dual_EC_DRBG, ngân sách 250 triệu USD).
    * Thu thập lưu lượng cáp quang diện rộng (MUSCULAR, PRISM, XKeyscore).
    * Can thiệp phần cứng chuỗi cung ứng và khai thác phân tích siêu dữ liệu (*metadata surveillance*).
- **Slide 9: Đóng góp & Chương trình nghiên cứu mới của Halpin**
  - Mở rộng mật mã học sang mạng xã hội - kỹ thuật không đồng nhất (*heterogeneous techno-social networks*).
  - Chuyển hướng bảo vệ sang siêu dữ liệu, kỹ thuật chống phân tích lưu lượng và mạng ẩn danh mixnet (Nym).
  - Kêu gọi xây dựng một nền khoa học an ninh thông tin thực nghiệm (*empirical science of security*).
- **Slide 10: Phân tích Khả năng Tái lập & Tiêu chuẩn Đánh giá Giữa kỳ**
  - Đối chiếu 2 khung điểm: Khung [6-8 Điểm] (Tái lập bài báo; nhưng Halpin 2025 không có code/dataset gốc) vs Khung [9-10 Điểm] (Tự xây dựng models, thuật toán cho kết quả tốt hơn trên tập dataset).
  - Phân loại rõ ràng: Nhóm A (0%), Nhóm B (100% minh họa bài báo), Nhóm C (Thuật toán đề xuất), Nhóm D (Tài liệu mật NSA).
- **Slide 11: Bảng đối chiếu Paper-to-Code của Đề tài DT005**
  - Ma trận đối chiếu chi tiết giữa 5 luận điểm triết học của Halpin và các module code minh họa do sinh viên xây dựng trong `src/baselines/` và `src/proposed_model/`.
- **Slide 12: Đóng góp Thuật toán Tự xây dựng Đạt Chuẩn [9-10 Điểm]: Cây Merkle Phân cấp Động (HD-Merkle)**
  - **Gói gọn trong đúng 1 trang slide duy nhất**: Khắc phục hạn chế của Merkle truyền thống (Rebuild $O(N)$ mất 283.9 ms khi thêm/xóa tệp).
  - Đột phá thuật toán: Cây phân cấp theo thư mục + Dirty-Path Cache $O(\log N)$ + Phân tách miền RFC 6962 + CompactProof 256 B.
  - Kết quả vượt trội trên Kaggle dataset: Thêm tệp mất **0.24 ms** (nhanh hơn ~1.200 lần), xóa tệp mất **0.09 ms** (nhanh hơn ~2.600 lần), tiết kiệm 99.9% băng thông.
- **Slide 13: Thiết kế Thực nghiệm Demo: Bám sát 100% Khái niệm Kẻ đối đầu của Bài báo**
  - Red Team (The Adversary Halpin 2025): Phát động các đòn tấn công bám sát bài báo: CRC32 Forge trên $\mathbb{F}_2$, Dolev-Yao Proof Swap, Tấn công kênh kề đo thời gian Kocher, và phá hoại tệp tin thực tế.
  - Blue Team (Hệ thống giám sát): 3 cơ chế toàn vẹn (Flat Hash Kerckhoffs, HMAC có khóa, Merkle Tree) và Rollback All hoàn nguyên 100%.
  - Kết quả quan sát: CRC32 bị qua mặt 100%; SHA-256 phát hiện gian lận ngay lập tức; Merkle Tree chặn đứng 100% tráo đổi Dolev-Yao.
- **Slide 14: Tổng kết Báo cáo Giữa kỳ & Sẵn sàng Thực hiện Demo Trực tiếp**
  - Hoàn thành toàn diện việc nghiên cứu bài báo gốc và kiểm toán khoa học trung thực.
  - Đáp ứng tiêu chuẩn đánh giá [9 - 10 Điểm] bằng việc tự xây dựng thuật toán cho kết quả vượt trội trên dataset mà không làm loãng bài báo.
  - Sẵn sàng demo tương tác qua Terminal CLI Red Team hoặc Web Dashboard (Port 8080).

---
*Báo cáo kiểm toán được lập ngày 24/10/2026 bởi Nhóm thực hiện Đề tài DT005.*
