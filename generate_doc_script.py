# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word (.docx) Kịch Bản Thuyết Trình & Hướng Dẫn Diễn Tập Demo:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn tinh giản:
- Giữ trọn vẹn toàn bộ nội dung học thuật, KHÔNG bỏ ý hay bỏ câu.
- Cắt gọt các từ ngữ rườm rà, từ đệm, từ lặp lại.
- Thay thế các từ ngữ quá cao siêu, trừu tượng hoặc mang tính cảm xúc/kịch tính bằng từ ngữ kỹ thuật rõ ràng, chuẩn mực và dễ hiểu.
- Hộp giải thích dễ hiểu (💡) tập trung đúng vào các khái niệm khó hiểu với câu văn ngắn gọn, trực quan.
"""

import sys
import os
import shutil
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def create_document():
    doc = docx.Document()

    # Thiết lập lề 1 inch chuẩn văn bản học thuật
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Đặt font chữ mặc định
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(11)
    font_normal.color.rgb = RGBColor(30, 41, 59) # Slate 800

    # TIÊU ĐỀ CHÍNH CỦA TÀI LIỆU
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_inst = p_title.add_run("TRƯỜNG ĐẠI HỌC TÀI NGUYÊN VÀ MÔI TRƯỜNG TP. HỒ CHÍ MINH\nKHOA CÔNG NGHỆ THÔNG TIN\n")
    run_inst.font.size = Pt(11)
    run_inst.font.bold = True
    run_inst.font.color.rgb = RGBColor(71, 85, 105)

    run_t = p_title.add_run("KỊCH BẢN THUYẾT TRÌNH & HƯỚNG DẪN DIỄN TẬP DEMO GIỮA KỲ\n")
    run_t.font.size = Pt(17)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(30, 64, 175) # Royal Blue

    run_sub = p_title.add_run("Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động\n")
    run_sub.font.size = Pt(12.5)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(15, 23, 42)

    run_paper = p_title.add_run('Khảo sát và Thẩm định học thuật bài báo: "The adversary: the philosophy of cryptography"\nTS. Harry Halpin — Journal of Cybersecurity, Oxford Academic, Vol. 11, No. 1, 2025 (DOI: 10.1093/cybsec/tyaf006)\n')
    run_paper.font.size = Pt(10)
    run_paper.font.italic = True
    run_paper.font.color.rgb = RGBColor(100, 116, 139)

    # THÔNG TIN SINH VIÊN & HỌC PHẦN
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.paragraph_format.space_after = Pt(16)
    run_meta = p_info.add_run("Học phần: An toàn và bảo mật hệ thống thông tin | Giảng viên hướng dẫn: ThS. Phạm Trọng Huynh\nSinh viên thực hiện: Võ Anh Kiệt — MSSV: 1150080143 — Lớp: 11_CNPM2")
    run_meta.font.size = Pt(10.5)
    run_meta.font.bold = True
    run_meta.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # =========================================================================
    # PHẦN 1: NỘI DUNG VÀ KỊCH BẢN THUYẾT MINH CHI TIẾT THEO 15 SLIDE
    # =========================================================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("PHẦN 1: NỘI DUNG TRÊN SLIDE & HƯỚNG DẪN GIẢI THÍCH DỄ HIỂU (15 SLIDE)")
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 64, 175)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(8)

    slides_data = [
        # SLIDE 1
        {
            "num": 1,
            "title": "BÌA BÁO CÁO GIỮA KỲ",
            "time": "0:30",
            "header": "Trường ĐH Tài nguyên và Môi trường TP.HCM — Khoa Công nghệ thông tin",
            "lead": "Báo cáo thuyết trình đồ án giữa kỳ — Đề tài DT005",
            "bullets": [
                ("Tên đề tài:", "Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động."),
                ("Bài báo khảo sát:", "The adversary: the philosophy of cryptography — TS. Harry Halpin (Journal of Cybersecurity, Oxford Academic, 2025)."),
                ("Thông tin thực hiện:", "Học phần: An toàn và bảo mật hệ thống thông tin | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143.")
            ],
            "explanation": (
                "Mở đầu báo cáo:\n"
                "Báo cáo gồm 2 phần: (1) Thẩm định bài báo của TS. Harry Halpin về mô hình đối thủ trong mật mã học; (2) Trình diễn hệ thống demo kiểm tra toàn vẹn tệp tin để minh họa thực tế cho các luận điểm của tác giả."
            ),
            "footer": "Bản trình chiếu chính thức phục vụ nghiệm thu đồ án giữa kỳ."
        },

        # SLIDE 2
        {
            "num": 2,
            "title": "BỐI CẢNH, MỤC TIÊU & BA ĐÓNG GÓP CỦA BÀI BÁO",
            "time": "1:00",
            "header": "SECTION 1 — GIỚI THIỆU & MỤC TIÊU BÀI BÁO (TRANG 1–2)",
            "lead": "Khảo cứu nhận thức luận về Kẻ thù trong Mật mã học và Chương trình nghiên cứu tương lai",
            "bullets": [
                ("1. Bối cảnh mật mã học:", "Khoa học về truyền thông bí mật khi có đối thủ có chủ đích trên môi trường mạng mở."),
                ("2. Nghịch lý nhận thức luận:", "Lý thuyết thông tin của Shannon (1948) được triết học nghiên cứu nhiều, nhưng mật mã học lại bị bỏ quên suốt nhiều thập kỷ."),
                ("3. Cái giá của sự trừu tượng hóa:", "Xem mật mã học chỉ là toán học trên giấy sẽ tạo ra điểm mù trước các đòn tấn công vật lý trong thế giới thực."),
                ("4. Ba đóng góp học thuật của Halpin:", "Hệ thống hóa lịch sử các mô hình đối thủ; Phê phán giới hạn của việc toán học hóa; Đề xuất 4 câu hỏi nghiên cứu mở kết hợp liên ngành.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Nghịch lý nhận thức luận: Triết học nghiên cứu nhiều về lý thuyết thông tin, nhưng lại bỏ quên mật mã học — nơi thực sự giải quyết bài toán truyền tin khi có đối thủ theo dõi. Halpin đặt vấn đề để lấp khoảng trống này.\n"
                "• Cái giá của sự trừu tượng hóa: Xem an toàn chỉ là công thức toán trên giấy sẽ bỏ sót thực tế. Kẻ tấn công ngoài đời không cần bẻ toán, mà đo thời gian xử lý hoặc mức tiêu thụ điện của thiết bị để tìm khóa."
            ),
            "footer": "Trích dẫn: Harry Halpin (2025), Section 1: Introduction, Journal of Cybersecurity, tr. 1–2."
        },

        # SLIDE 3
        {
            "num": 3,
            "title": "CƠ SỞ TRIẾT HỌC: THÔNG TIN, BÍ MẬT VÀ TÁC NHÂN",
            "time": "1:00",
            "header": "SECTION 2 & 3 — CƠ SỞ TRIẾT HỌC VỀ THÔNG TIN (TRANG 2–5)",
            "lead": "Khung nhận thức luận của Fred Dretske và Brian Cantwell Smith ứng dụng trong Mật mã học",
            "bullets": [
                ("1. Bản thể học thông tin (Brian Cantwell Smith):", "Thế giới khởi đầu từ dòng dữ liệu chưa có ranh giới; con người định hình thực tại bằng cách 'đăng ký đối tượng' (registration of objects) giữa các tác nhân ở xa nhau."),
                ("2. Thuyết thông tin khách quan (Fred Dretske):", "Thông tin tồn tại khách quan theo quy luật tự nhiên; bản rõ mang thông tin về nguồn, hệ mật đảm bảo bản mã không để lộ thông tin về nguồn."),
                ("3. Bản chất của đối thủ:", "Đối thủ là tác nhân có chủ đích (intentional agent), có mục tiêu rõ ràng, khác với nhiễu ngẫu nhiên vật lý."),
                ("4. Bản chất của Bí mật & Tri thức:", "Common Knowledge (David Lewis): Mọi người cùng biết và biết người khác cũng biết. Public Knowledge: Thông tin ai cũng có thể tiếp cận. Bí mật là 'tri thức không phổ biến', không đồng nhất với bản rõ hay bản mã.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Đăng ký đối tượng (Smith): Dữ liệu truyền đi vốn là dòng liên tục. Để bảo vệ, hệ thống phải phân định rõ từng phần: Bản rõ (M), Bản mã (C), Khóa (K). Có phân định rõ thì hệ thống mới kiểm soát được luồng bảo mật.\n"
                "• Thông tin khách quan (Dretske): Thông tin tồn tại độc lập với người nhận (như tiếng sấm báo trời mưa). Bản rõ mang thông tin gốc; khi mã hóa tốt, bản mã trở thành chuỗi ngẫu nhiên không để lộ thông tin về nguồn.\n"
                "• Common Knowledge vs Public Knowledge (Lewis 1969):\n"
                "  - Tri thức chung (Common Knowledge): Mọi người cùng biết một thông tin và biết người khác cũng biết (như luật đèn đỏ, ai cũng biết và tin người khác cũng tuân thủ).\n"
                "  - Tri thức công khai (Public Knowledge): Thông tin bất kỳ ai cũng có thể tiếp cận (như tài liệu bị rò rỉ lên mạng).\n"
                "  - Bí mật: Là 'tri thức không phổ biến', chỉ lưu hành trong nhóm nhất định, không để lộ ra ngoài."
            ),
            "footer": "Trích dẫn khoa học: Dretske (1981, Ref [1]); Cantwell Smith (1996, Ref [5]); Halpin (2025), tr. 2–5."
        },

        # SLIDE 4
        {
            "num": 4,
            "title": "TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ TRONG MẬT MÃ",
            "time": "1:00",
            "header": "SƠ ĐỒ 1: DÒNG THỜI GIAN TIẾN HÓA MÔ HÌNH ĐỐI THỦ",
            "lead": "Sự dịch chuyển từ Lý thuyết thông tin sang Giả định độ phức tạp tính toán và Phê phán hậu Snowden",
            "bullets": [
                ("Giai đoạn 1 (1948–1949) — Mô hình Shannon:", "Mô hình bảo mật điểm-điểm; định nghĩa Bí mật hoàn hảo H(K) >= H(M). Giới hạn: Trừu tượng hóa đối thủ trên kênh toán học, bỏ qua chi phí phân phối khóa."),
                ("Giai đoạn 2 (1976) — Cách mạng Khóa công khai:", "Diffie-Hellman và Merkle giải bài toán phân phối khóa bằng các bài toán độ phức tạp tính toán một chiều."),
                ("Giai đoạn 3 (1982–1984) — PPT Adversary & Security Games:", "Goldwasser-Micali mô hình hóa đối thủ thời gian đa thức xác suất (PPT) và kiểm chứng an toàn qua trò chơi an ninh."),
                ("Giai đoạn 4 (1981–1993) — Dolev-Yao & Random Oracle:", "Dolev-Yao mô hình hóa đối thủ kiểm soát kênh mạng; Bellare-Rogaway đề xuất Random Oracle Model (ROM) cho hàm băm."),
                ("Giai đoạn 5 (2013–2025) — Phê phán hậu Snowden:", "Snowden phơi bày năng lực thực tế của đối thủ thể chế: Kênh kề, cửa sau tiêu chuẩn, can thiệp chuỗi cung ứng và giám sát siêu dữ liệu.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Giới hạn của Shannon: Bí mật hoàn hảo đòi hỏi khóa phải dài bằng thông điệp (One-Time Pad). Gửi dữ liệu 10 GB mà phải chuyển trước khóa 10 GB tận tay thì tốn kém và không thực tế trên Internet.\n"
                "• PPT Adversary & Trò chơi an ninh (1984): Đưa ra giả định thực tế: Đối thủ dùng máy tính thời gian đa thức (năng lực tính toán hữu hạn). An toàn được chứng minh nếu đối thủ không phân biệt được hai bản mã với xác suất tốt hơn đoán ngẫu nhiên (50%)."
            ),
            "footer": "Sơ đồ tổng hợp dựa trên Section 3–6, Harry Halpin (2025), Journal of Cybersecurity, tr. 3–12."
        },

        # SLIDE 5
        {
            "num": 5,
            "title": "MÔ HÌNH TRUYỀN THÔNG CỦA SHANNON VÀ MẠNG THỰC TẾ",
            "time": "1:00",
            "header": "SƠ ĐỒ 2: ĐỐI SÁNH MÔ HÌNH TRUYỀN THÔNG ĐIỂM-ĐIỂM VÀ MẠNG XÃ HỘI-KỸ THUẬT",
            "lead": "Khoảng cách giữa mô hình điểm-điểm đóng kín của Shannon và mạng phức hợp không đồng nhất",
            "bullets": [
                ("Panel Trái — Mô hình điểm-điểm Shannon (1949):", "Alice gửi bản rõ M mã hóa bằng khóa K qua kênh truyền đến Bob; đối thủ Eve nghe lén thụ động. Bỏ qua chi phí quản lý khóa và bề mặt tấn công vật lý."),
                ("Panel Phải — Mạng xã hội - kỹ thuật (Halpin 2025):", "Thông tin luân chuyển qua hạ tầng mở, ISP, đám mây. Đối thủ thể chế (NSA) tấn công qua 4 hướng: Kênh kề, giám sát siêu dữ liệu, cửa sau phần cứng và phản bội nội bộ.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Shannon vs Thực tế: Shannon giả định đường truyền điểm-điểm đơn giản, chỉ có kẻ nghe lén trên dây dẫn. Thực tế Internet dữ liệu phải qua ISP và máy chủ đám mây. Đối thủ hiện đại không chỉ chặn nội dung mà còn thu thập Siêu dữ liệu (ai gửi cho ai, lúc nào, dung lượng bao nhiêu) để vẽ lại hành vi người dùng mà không cần giải mã."
            ),
            "footer": "Sơ đồ tổng hợp dựa trên Shannon (1949) và Harry Halpin (2025), Journal of Cybersecurity, tr. 3–5."
        },

        # SLIDE 6
        {
            "num": 6,
            "title": "BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI",
            "time": "1:00",
            "header": "SECTION 4 — HÌNH THỨC HÓA ĐỐI THỦ (TRANG 5–8)",
            "lead": "Tiến trình toán học hóa kẻ thù: Độ phức tạp tính toán đa thức, Logic kênh truyền và Tiên tri ngẫu nhiên",
            "bullets": [
                ("1. PPT Adversary (Goldwasser-Micali 1984):", "Đối thủ là thuật toán Turing xác suất chạy thời gian đa thức; an toàn định nghĩa qua trò chơi thách đấu; chứng minh bằng quy dẫn về bài toán khó."),
                ("2. Đối thủ mạng logic (Dolev-Yao 1983):", "Đối thủ kiểm soát toàn bộ kênh truyền (chặn, sửa, xóa, phát lại gói tin), nhưng xem mật mã là hộp đen an toàn tuyệt đối nếu không có khóa."),
                ("3. Tiên tri ngẫu nhiên (Bellare-Rogaway 1993):", "Giả định hàm băm hoạt động như hàm ngẫu nhiên lý tưởng (Random Oracle); thực tế không hàm băm nào đạt được tính ngẫu nhiên toán học này.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Điểm cốt lõi bài báo phê phán: Cả 3 mô hình đều lý tưởng hóa đối thủ trên giấy:\n"
                "  - Dolev-Yao cho đối thủ kiểm soát mạng nhưng lại giả định mật mã là hộp đen không thể lộ khóa.\n"
                "  - Random Oracle giả định hàm băm ngẫu nhiên hoàn hảo, trong khi hàm băm thực tế (như SHA-256) chỉ là thuật toán xác định.\n"
                "  - Thực tế, đối thủ không chỉ giải toán mà còn tấn công vào lỗi phần cứng, kênh kề và quy trình quản trị."
            ),
            "footer": "Trích dẫn: Goldwasser-Micali (1984); Dolev-Yao (1983); Bellare-Rogaway (1993); Halpin (2025), tr. 5–8."
        },

        # SLIDE 7
        {
            "num": 7,
            "title": "GIỚI HẠN CỦA SỰ HÌNH THỨC HÓA TRONG MẬT MÃ HỌC",
            "time": "1:00",
            "header": "SECTION 5 — GIỚI HẠN TRỪU TƯỢNG HÓA (TRANG 8–10)",
            "lead": "Khoảng cách giữa các giả định toán học lý tưởng và thế giới vật lý thực tế",
            "bullets": [
                ("Bước 1 — Giả định toán học:", "Hộp đen mật mã lý tưởng, thời gian thực thi vô hình, độ dài bản mã không làm lộ nội dung, kênh truyền độc lập."),
                ("Bước 2 — Yếu tố vật lý bị lược bỏ:", "Thời gian xử lý CPU, điện năng tiêu thụ, bức xạ điện từ, bộ nhớ đệm Cache, lỗi lập trình."),
                ("Bước 3 — Đối thủ thế giới thực khai thác:", "Tấn công kênh kề thời gian (Kocher 1996), phân tích lưu lượng và kích thước gói tin, lỗi quản lý khóa (Anderson 2001)."),
                ("Bước 4 — Kết luận của Halpin:", "Sự trừu tượng hóa quá mức làm mật mã học bỏ sót chính những năng lực thực tế của đối thủ ngoài đời thực.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Kênh kề thời gian (Kocher 1996): Khi so sánh mật khẩu, máy tính thường ngắt sớm ngay khi gặp ký tự sai đầu tiên. Đối thủ gửi thử các chuỗi và đo thời gian phản hồi: Thời gian xử lý lâu hơn nghĩa là đã đoán đúng thêm một ký tự. Nhờ đó, đối thủ dò ra khóa mà không cần phá thuật toán.\n"
                "• Rò rỉ kích thước: Bản mã có mã hóa kín nhưng độ dài gói tin vẫn lộ trên mạng, giúp đối thủ đoán được loại dữ liệu đang truyền."
            ),
            "footer": "Sơ đồ tổng hợp dựa trên Kocher (1996) và Halpin (2025), Section 5, tr. 8–10."
        },

        # SLIDE 8
        {
            "num": 8,
            "title": "SNOWDEN VÀ ĐỐI THỦ THỂ CHẾ (THE REAL ADVERSARY)",
            "time": "1:00",
            "header": "SECTION 6 — TIẾT LỘ CỦA SNOWDEN (TRANG 10–12)",
            "lead": "Khủng hoảng mô hình an ninh trước đối thủ giám sát hàng loạt quy mô quốc gia",
            "bullets": [
                ("1. Khủng hoảng thực tiễn (Snowden 2013):", "NSA không cần phá toán học của RSA/AES mà tấn công vào hạ tầng xung quanh (đánh cắp khóa, thâm nhập cáp quang biển)."),
                ("2. Năng lực đối thủ thể chế:", "Khai thác siêu dữ liệu để lập bản đồ liên lạc; cài cửa sau vào tiêu chuẩn NIST (Dual_EC_DRBG); gây áp lực pháp lý lên các tập đoàn công nghệ."),
                ("3. Tính tương đối của đối thủ:", "'Đối thủ' mang tính vị thế; mật mã học buộc phải nghiên cứu cơ chế phòng chống cửa sau (Cliptography).")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "• Cửa sau Dual_EC_DRBG: Bộ sinh số ngẫu nhiên là nền tảng tạo khóa. Thay vì phá khóa, NSA tác động đưa thuật toán có khiếm khuyết vào tiêu chuẩn NIST. Do nắm trước hằng số bí mật, NSA tính được khóa của mọi hệ thống dùng chuẩn này.\n"
                "• Cliptography: Ngành nghiên cứu thiết kế mật mã an toàn ngay cả khi thuật toán hoặc phần cứng bị cài sẵn cửa sau."
            ),
            "footer": "Trích dẫn: Harry Halpin (2025), Section 6: Snowden and the real adversary, tr. 10–12."
        },

        # SLIDE 9
        {
            "num": 9,
            "title": "CHƯƠNG TRÌNH NGHIÊN CỨU TƯƠNG LAI: BỐN CÂU HỎI MỞ",
            "time": "1:00",
            "header": "SECTION 7 — KẾT LUẬN & CHƯƠNG TRÌNH NGHIÊN CỨU (TRANG 12–14)",
            "lead": "Đề xuất hòa giải giữa trừu tượng hóa toán học và dữ liệu thực nghiệm thế giới thực",
            "bullets": [
                ("RQ1 — Mạng xã hội động:", "Nghiên cứu hành vi giữ bí mật thực tế của con người; đánh giá độ bền chính sách an ninh trước sự cố nội bộ Byzantine."),
                ("RQ2 — Mô hình hóa bằng thống kê:", "Dùng thống kê thực nghiệm và mô phỏng Monte Carlo để định lượng năng lực thực tế của đối thủ và liên minh giám sát."),
                ("RQ3 — Phương pháp liên ngành:", "Ứng dụng lý thuyết hệ thống phức hợp và động lực học mạng (như mô hình lây lan) vào phân tích sự rò rỉ thông tin."),
                ("RQ4 — Triết học chất vấn ngầm định:", "Xem xét lại các giả định nền tảng từ thời Shannon, điều chỉnh cho phù hợp với môi trường mạng hiện đại.")
            ],
            "explanation": (
                "Giải thích ngắn gọn:\n"
                "Ý nghĩa 4 câu hỏi: Mật mã học tương lai cần liên ngành: (1) Quan sát hành vi thực tế của người dùng; (2) Dùng dữ liệu thống kê lượng hóa đối thủ; (3) Áp dụng mô hình lây lan mạng để theo dõi rò rỉ tin; (4) Dùng triết học xem xét lại các giả định cũ từ thời Shannon."
            ),
            "footer": "Trích dẫn nguyên văn 4 câu hỏi nghiên cứu mở: Harry Halpin (2025), Section 7, tr. 12–14."
        },

        # SLIDE 10
        {
            "num": 10,
            "title": "KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG: RED TEAM VS BLUE TEAM",
            "time": "0:45",
            "header": "SƠ ĐỒ 4: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG (PRE-DEMO ARCHITECTURE)",
            "lead": "Mô hình đối kháng 2 màn hình trên tập dữ liệu kaggle_dataset (2.485 tệp tin) kiểm chứng trực quan luận điểm bài báo",
            "bullets": [
                ("1. Sơ đồ diễn tập 2 màn hình:", "Màn hình trái là Blue Team (Port 8080) giám sát toàn vẹn tệp; màn hình phải là Red Team (Port 8081) thực hiện can thiệp."),
                ("2. Dữ liệu thực nghiệm:", "Thư mục kaggle_dataset (2.485 tệp tin) đóng vai trò môi trường kiểm thử thực tế để đo kiểm tra toàn vẹn."),
                ("3. Cầu nối minh họa độc lập (Paper-to-Code):", "Bài báo Halpin là khảo luận triết học thuần túy (không có code). Hệ thống kiểm chứng trực quan 3 luận điểm: (1) Mã CRC32 bị qua mặt trên GF(2) vs SHA-256 phi tuyến; (2) Phân tách miền nút lá RFC 6962; (3) Kênh kề thời gian so sánh chuỗi."),
                ("4. Cơ chế hoàn nguyên an toàn:", "Thư mục .adversary_backups/ phục hồi nguyên trạng 100% dữ liệu với nút Rollback All.")
            ],
            "explanation": (
                "Hướng dẫn thuyết minh ngắn gọn:\n"
                "Đây là sơ đồ bố trí màn hình trước khi Demo: màn hình phải tấn công, màn hình trái phát hiện. Nhóm tự phát triển phần mềm này để minh họa trực quan cho các luận điểm của tác giả. Nói ngắn gọn trong 30 giây để chuyển ngay vào màn hình Demo trực tiếp."
            ),
            "footer": "Sơ đồ kiến trúc diễn tập đối kháng độc lập do sinh viên thiết kế minh họa các khái niệm an ninh."
        },

        # SLIDE 11
        {
            "num": 11,
            "title": "SLIDE CHUYỂN TIẾP: DEMO",
            "time": "0:15",
            "header": "GIAO DIỆN TỐI GIẢN NỀN SÁNG CHUẨN BỊ DIỄN TẬP TRỰC TIẾP",
            "lead": "Chữ hiển thị trên Slide: DEMO",
            "bullets": [
                ("Văn bản trung tâm:", "DEMO (Slide chuyển tiếp tối giản nền sáng, tạo nhịp nghỉ trước phần thực hành).")
            ],
            "explanation": (
                "Lời thoại ngắn gọn:\n"
                "Chuyển tiếp sang phần Demo thực tế 2 màn hình Red Team vs Blue Team để kiểm chứng các luận điểm đối kháng vừa trình bày."
            ),
            "footer": "Slide chuyển tiếp phần diễn tập thực hành đối kháng an ninh."
        },

        # SLIDE 12
        {
            "num": 12,
            "title": "BẢNG TỔNG KẾT BÁO CÁO: ĐÓNG GÓP HỌC THUẬT & BÀI HỌC AN NINH",
            "time": "1:00",
            "header": "SECTION 7 — TỔNG KẾT TOÀN VĂN BÀI BÁO (HARRY HALPIN 2025)",
            "lead": "Bảng đối sánh giữa các giả định hình thức truyền thống và nhận thức luận của Harry Halpin",
            "bullets": [
                ("1. Bản chất thông tin & Bí mật (Sec 1–2):", "Từ bỏ giả định nhiễu ngẫu nhiên của Shannon; đối thủ là tác nhân có chủ đích; mã tuyến tính CRC32 không thể thay thế hàm băm phi tuyến SHA-256."),
                ("2. Giới hạn trừu tượng hóa (Sec 5):", "Vượt qua hộp đen toán học lý tưởng; xây dựng hệ thống an toàn đa tầng; triệt tiêu kênh kề thời gian bằng hàm hằng số (hmac.compare_digest)."),
                ("3. Đối thủ thế giới thực (Sec 6):", "Snowden chứng minh đối thủ thể chế (NSA) tấn công chuỗi cung ứng và cài cửa sau; cần nghiên cứu Cliptography và phân tách miền minh bạch (RFC 6962)."),
                ("4. Chương trình nghiên cứu tương lai (Sec 7):", "Liên ngành hóa mật mã học: kết hợp toán học, thống kê lượng hóa đối thủ, hệ thống phức hợp và triết học nhận thức.")
            ],
            "explanation": (
                "Hướng dẫn thuyết minh Bảng tổng kết:\n"
                "Bảng này đúc kết 4 trụ cột cốt lõi của toàn văn bài báo Halpin:\n"
                "• Trụ cột 1: Đối thủ là tác nhân có chủ đích, không phải nhiễu ngẫu nhiên.\n"
                "• Trụ cột 2: Mật mã học trên giấy không bảo vệ được trước các đòn tấn công kênh kề vật lý.\n"
                "• Trụ cột 3: Đối thủ thể chế tấn công vào chuỗi cung ứng chứ không phá toán học.\n"
                "• Trụ cột 4: Mật mã học tương lai cần liên ngành với thống kê và triết học."
            ),
            "footer": "Tài liệu tham khảo chính: Halpin (2025), Shannon (1949), Goldwasser-Micali (1984), Dolev-Yao (1983), Kocher (1996), RFC 6962."
        },

        # SLIDE 13
        {
            "num": 13,
            "title": "CẢM ƠN & PHIÊN PHẢN BIỆN (Q&A)",
            "time": "0:15",
            "header": "Trường ĐH Tài nguyên và Môi trường TP.HCM — Khoa Công nghệ thông tin",
            "lead": "CẢM ƠN THẦY VÀ CÁC BẠN ĐÃ LẮNG NGHE!",
            "bullets": [
                ("Thông điệp kết thúc:", "Q & A — Xin đón nhận câu hỏi và góp ý từ Thầy và Hội đồng."),
                ("Thông tin đề tài:", "Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143.")
            ],
            "explanation": (
                "Lời kết thúc:\n"
                "Cảm ơn Thầy và các bạn đã theo dõi. Nhóm xin lắng nghe ý kiến đóng góp và trả lời các câu hỏi phản biện."
            ),
            "footer": "Phiên hỏi đáp học thuật và đóng góp ý kiến hoàn thiện đồ án."
        }
    ]

    for s in slides_data:
        # Tiêu đề Slide
        p_num = doc.add_paragraph()
        p_num.paragraph_format.space_before = Pt(12)
        p_num.paragraph_format.space_after = Pt(2)
        r_n = p_num.add_run(f"► SLIDE {s['num']}: {s['title']} ")
        r_n.font.size = Pt(12)
        r_n.font.bold = True
        r_n.font.color.rgb = RGBColor(30, 64, 175)

        r_time = p_num.add_run(f"[{s['time']}]")
        r_time.font.size = Pt(10)
        r_time.font.bold = True
        r_time.font.color.rgb = RGBColor(100, 116, 139)

        # Header / Lead trên slide
        p_hd = doc.add_paragraph()
        p_hd.paragraph_format.space_before = Pt(1)
        p_hd.paragraph_format.space_after = Pt(4)
        p_hd.paragraph_format.left_indent = Inches(0.2)
        r_hd = p_hd.add_run(f"• Vị trí Slide: {s['header']}\n  Tiêu đề phụ: {s['lead']}")
        r_hd.font.size = Pt(9.5)
        r_hd.font.italic = True
        r_hd.font.color.rgb = RGBColor(71, 85, 105)

        # Khối nội dung nguyên văn trên Slide
        p_lbl_sc = doc.add_paragraph()
        p_lbl_sc.paragraph_format.space_before = Pt(2)
        p_lbl_sc.paragraph_format.space_after = Pt(2)
        p_lbl_sc.paragraph_format.left_indent = Inches(0.2)
        r_lsc = p_lbl_sc.add_run("Nội dung hiển thị trên Slide:")
        r_lsc.font.size = Pt(10)
        r_lsc.font.bold = True
        r_lsc.font.color.rgb = RGBColor(15, 23, 42)

        for h, b in s['bullets']:
            p_it = doc.add_paragraph()
            p_it.paragraph_format.space_before = Pt(1)
            p_it.paragraph_format.space_after = Pt(2)
            p_it.paragraph_format.left_indent = Inches(0.35)
            r_h = p_it.add_run(f"– {h} ")
            r_h.font.size = Pt(10)
            r_h.font.bold = True
            r_h.font.color.rgb = RGBColor(30, 41, 59)

            r_b = p_it.add_run(b)
            r_b.font.size = Pt(10)
            r_b.font.color.rgb = RGBColor(51, 65, 85)

        # Hộp giải thích dễ hiểu (ngắn gọn, trực quan, không từ ngữ cao siêu)
        if s['explanation']:
            table = doc.add_table(rows=1, cols=1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            cell = table.cell(0, 0)
            cell.width = Inches(6.5)
            set_cell_background(cell, "F1F5F9") # Slate 100
            
            p_box = cell.paragraphs[0]
            p_box.paragraph_format.space_before = Pt(4)
            p_box.paragraph_format.space_after = Pt(4)
            p_box.paragraph_format.left_indent = Inches(0.1)
            p_box.paragraph_format.right_indent = Inches(0.1)
            
            r_bxt = p_box.add_run("💡 Hướng dẫn thuyết minh & Giải thích dễ hiểu:\n")
            r_bxt.font.size = Pt(10)
            r_bxt.font.bold = True
            r_bxt.font.color.rgb = RGBColor(30, 64, 175)
            
            r_bxb = p_box.add_run(s['explanation'])
            r_bxb.font.size = Pt(9.5)
            r_bxb.font.color.rgb = RGBColor(51, 65, 85)

        # Footer của slide
        p_ft = doc.add_paragraph()
        p_ft.paragraph_format.space_before = Pt(4)
        p_ft.paragraph_format.space_after = Pt(10)
        p_ft.paragraph_format.left_indent = Inches(0.2)
        r_ft = p_ft.add_run(f"• Ghi chú chân Slide: {s['footer']}")
        r_ft.font.size = Pt(9)
        r_ft.font.italic = True
        r_ft.font.color.rgb = RGBColor(148, 163, 184)

    doc.add_page_break()

    # =========================================================================
    # PHẦN 2: HƯỚNG DẪN THAO TÁC DIỄN TẬP DEMO THỰC TẾ
    # =========================================================================
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("PHẦN 2: HƯỚNG DẪN THAO TÁC DIỄN TẬP DEMO THỰC TẾ (TẬP TRUNG BÀI BÁO)")
    r2.font.size = Pt(13.5)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(30, 64, 175)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(8)

    p_intro_demo = doc.add_paragraph()
    p_intro_demo.add_run(
        "Mục tiêu diễn tập giữa kỳ: Tập trung vào 2 kịch bản đối kháng cốt lõi phản ánh bài báo của Harry Halpin (2025). "
        "Các tính năng nghiên cứu mở rộng phục vụ đồ án cuối kỳ được lược bỏ để đảm bảo thời lượng và làm nổi bật trọng tâm khoa học."
    )

    # PHƯƠNG ÁN 1: GIAO DIỆN WEB TRỰC QUAN
    p_opt1 = doc.add_paragraph()
    p_opt1.paragraph_format.space_before = Pt(8)
    r_op1 = p_opt1.add_run("Phương án 1: Diễn tập Đối kháng 2 Màn hình trên Trình duyệt Web (Khuyến nghị ưu tiên)")
    r_op1.font.bold = True
    r_op1.font.size = Pt(11.5)
    r_op1.font.color.rgb = RGBColor(15, 23, 42)

    demo_steps = [
        ("Bước 1: Khởi động 2 dịch vụ Web trong nền",
         "Mở terminal và chạy 2 dịch vụ:\n"
         "• Blue Team Service:  python -m src.demo.web_demo --port 8080\n"
         "• Red Team Service:   python -m src.demo.red_team_web --port 8081\n"
         "Mở 2 cửa sổ trình duyệt chia đôi màn hình: Cửa sổ trái vào http://localhost:8080 (Blue Team); cửa sổ phải vào http://localhost:8081 (Red Team)."),

        ("Bước 2: Thiết lập mốc tin cậy sạch tại Blue Team (Port 8080)",
         "Tại cổng 8080, chọn thư mục 'kaggle_dataset' (~112.8 MB).\n"
         "Bấm 'Nạp & Tính toán Toàn vẹn'. Hệ thống sinh Root Hash ban đầu H_0 sạch với trạng thái xanh: '100% NGUYÊN VẸN'.\n"
         "Điểm nói lý thuyết (Paper-to-Code): Minh họa việc thiết lập trạng thái tin cậy khách quan theo Dretske. Hệ thống áp dụng chuẩn RFC 6962 (tiền tố 0x00 cho lá, 0x01 cho nhánh) khi băm cây dữ liệu, giải quyết triệt để điểm mù nhập nhằng cấu trúc dữ liệu mà Halpin cảnh báo tại Mục 5."),

        ("Bước 3: Phát động Kịch bản 1 — Can thiệp tệp tin tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm 'Can thiệp tệp tin ngẫu nhiên' (In-place Tampering) để sửa đổi tệp Resume/Resume.csv trên đĩa.\n"
         "Quay lại Blue Team (Port 8080), bấm 'Kiểm tra toàn vẹn':\n"
         "• Hệ thống đổi sang cảnh báo đỏ: 'CẢNH BÁO: PHÁT HIỆN CAN THIỆP DỮ LIỆU!'.\n"
         "• Bảng truy vết chỉ rõ tệp Resume/Resume.csv bị thay đổi.\n"
         "Ý nghĩa: Chứng minh hàm băm phát hiện ngay lập tức hành vi sửa đổi của đối thủ."),

        ("Bước 4: Phát động Kịch bản 2 — Bù 4 byte GF(2) qua mặt CRC32 tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm 'Tấn công bù 4 byte GF(2)' (CRC32 Forge Attack). Động cơ Red Team giải hệ phương trình tuyến tính, làm sai lệch nội dung tệp nhưng bù 4 byte patch khiến mã CRC32 không đổi.\n"
         "Quay lại Blue Team (Port 8080), bấm 'Kiểm tra toàn vẹn':\n"
         "• Bảng đối chứng hiển thị: Thuật toán CRC32 bị qua mặt (báo khớp sai do tính chất tuyến tính).\n"
         "• Ngược lại, hàm băm SHA-256 phát hiện gian lận ngay lập tức nhờ tính phi tuyến.\n"
         "Điểm nói lý thuyết (Paper-to-Code): Minh họa trực tiếp luận điểm Halpin (Mục 1 & 2): Mã kiểm tra tuyến tính (CRC32) chỉ bắt được nhiễu ngẫu nhiên trong truyền dẫn, hoàn toàn bất lực trước kẻ thù có chủ đích. Ngược lại, hàm băm mật mã SHA-256 có tính phi tuyến và hiệu ứng tuyết lở (~50% bit đảo khi sửa 1 bit), ngăn chặn kẻ thù bẻ gãy bằng đại số."),

        ("Bước 5: Hoàn nguyên dữ liệu sạch (Rollback All)",
         "Tại giao diện Red Team (Port 8081), bấm 'ROLLBACK ALL' để khôi phục toàn bộ tệp từ thư mục sao lưu .adversary_backups/.\n"
         "Tại Blue Team bấm kiểm tra lại: Hệ thống trở về trạng thái an toàn màu xanh.")
    ]

    for s_title, s_desc in demo_steps:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(4)
        p_st.paragraph_format.space_after = Pt(2)
        p_st.paragraph_format.left_indent = Inches(0.2)
        r_st = p_st.add_run(f"{s_title}: ")
        r_st.font.bold = True
        r_st.font.size = Pt(10.5)
        r_st.font.color.rgb = RGBColor(30, 64, 175)

        p_sd = doc.add_paragraph()
        p_sd.paragraph_format.space_before = Pt(0)
        p_sd.paragraph_format.space_after = Pt(6)
        p_sd.paragraph_format.left_indent = Inches(0.4)
        r_sd = p_sd.add_run(s_desc)
        r_sd.font.size = Pt(10)
        r_sd.font.color.rgb = RGBColor(51, 65, 85)

    # PHƯƠNG ÁN 2: DÒNG LỆNH CLI (DỰ PHÒNG)
    p_opt2 = doc.add_paragraph()
    p_opt2.paragraph_format.space_before = Pt(10)
    r_op2 = p_opt2.add_run("Phương án 2: Diễn tập Dòng lệnh Terminal CLI (Dự phòng trường hợp không dùng UI)")
    r_op2.font.bold = True
    r_op2.font.size = Pt(11.5)
    r_op2.font.color.rgb = RGBColor(15, 23, 42)

    cli_text = (
        "1. Chạy Suite Kiểm định Toán học Đối kháng:\n"
        "   Lệnh: python -m src.proposed_model.security_games\n"
        "   Kết quả & Điểm nói lý thuyết: Đo đạc tính phi tuyến của SHA-256, hiệu ứng tuyết lở SAC ~50%, kịch bản bẻ gãy mã tuyến tính CRC32 trên GF(2), và kiểm chứng rò rỉ kênh kề thời gian (Kocher 1996 - Mục 5 bài báo) khi so sánh chuỗi ngắt sớm so với hàm thời gian hằng số hmac.compare_digest.\n\n"
        "2. Diễn tập 2 Terminal Red Team vs Blue Team:\n"
        "   • Terminal 1 (Red Team): Chạy 'python src/adversary/red_team_cli.py'. Chọn phím '1' can thiệp tệp, hoặc phím '4' bù mã CRC32.\n"
        "   • Terminal 2 (Blue Team): Chạy 'python -m src.baselines.sha256_manifest --dir kaggle_dataset --verify'. Blue Team phát hiện vi phạm: MODIFIED_FILE: Resume/Resume.csv.\n"
        "   • Terminal 1: Chọn phím '9' để Rollback All khôi phục đĩa cứng an toàn."
    )

    p_cli = doc.add_paragraph()
    p_cli.paragraph_format.left_indent = Inches(0.2)
    r_c = p_cli.add_run(cli_text)
    r_c.font.size = Pt(10)
    r_c.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_page_break()

    # =========================================================================
    # PHẦN 3: BỘ CÂU HỎI & LUẬN ĐIỂM PHẢN BIỆN CHUYÊN SÂU (Q&A DEFENSE GUIDE)
    # =========================================================================
    h3 = doc.add_heading(level=1)
    r3 = h3.add_run("PHẦN 3: BỘ CÂU HỎI & LUẬN ĐIỂM PHẢN BIỆN CHUYÊN SÂU (Q&A DEFENSE)")
    r3.font.size = Pt(13.5)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(30, 64, 175)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(8)

    qa_list = [
        ("Câu hỏi 1: Bài báo của Halpin là khảo luận triết học, tại sao đề tài lại chọn bài toán Kiểm tra toàn vẹn tệp tin?",
         "Luận điểm phản biện:\n"
         "1. Thuộc tính an ninh cốt lõi: Trong tam giác an ninh CIA (Bảo mật, Toàn vẹn, Khả dụng), Halpin phân tích rằng Tính toàn vẹn (Integrity) là điều kiện tiên quyết. Một thông điệp có thể không bí mật nhưng bắt buộc phải toàn vẹn thì người nhận mới có thể tin cậy sử dụng.\n"
         "2. Kiểm định thực nghiệm: Halpin phê phán việc toán học hóa trừu tượng đã làm mật mã học xa rời thực tế vật lý. Bài toán kiểm tra toàn vẹn tệp tin trên đĩa cứng là bài toán thực tế, cho phép đưa các khái niệm trừu tượng vào kiểm định vật lý rõ ràng.\n"
         "3. Mô hình hóa đối thủ: Hệ thống toàn vẹn tệp tin cho phép minh họa trực quan các dạng đối thủ được Halpin đề cập: Kẻ can thiệp Dolev-Yao, kẻ khai thác kênh kề Kocher và kẻ sửa đổi dữ liệu có chủ đích."),

        ("Câu hỏi 2: Bài báo có 134 tài liệu tham khảo, khi thẩm định và tái lập thì có bắt buộc phải làm theo toàn bộ không?",
         "Luận điểm phản biện:\n"
         "1. Bản chất bài khảo cứu: Bài báo của Halpin là bài tổng quan triết học (philosophical survey) dài 17 trang, 134 trích dẫn là kho tàng tài liệu lịch sử từ năm 1883 đến 2024 (bao gồm sách triết học, tài liệu lưu trữ chính phủ về NSA, v.v.). Không một công trình kỹ thuật nào có thể hay bắt buộc phải tái lập toàn bộ 134 tài liệu nền đó.\n"
         "2. Tinh lọc trọng tâm: Thẩm định bài báo lý thuyết là tinh lọc các trụ cột kỹ thuật trực tiếp liên quan đến bài toán an ninh. Nhóm đã tinh lọc và kết nối với các tài liệu kỹ thuật cốt lõi: Shannon (Mã tuyến tính vs Hàm băm), Goldwasser-Micali (Trò chơi an ninh PPT), Kocher (Kênh kề thời gian), Bellare-Rogaway (Random Oracle), Dolev-Yao (Mạng thù địch), Snowden (Đối thủ thể chế), và bổ sung chuẩn IETF RFC 6962 về phân tách miền."),

        ("Câu hỏi 3: Tập dữ liệu kaggle_dataset từ đâu ra? Có xuất hiện trong bài báo gốc không?",
         "Luận điểm phản biện:\n"
         "1. Khẳng định minh bạch: Bài báo gốc của Harry Halpin hoàn toàn KHÔNG CÓ tập dữ liệu, bảng số liệu thực nghiệm hay mã nguồn nào. Bài báo là công trình khảo luận nhận thức luận trên tạp chí lý thuyết.\n"
         "2. Mục đích đưa tập dữ liệu vào: Tập dữ liệu Kaggle Resume Dataset (2.485 tệp tin, ~112.81 MiB theo benchmark) do nhóm sinh viên tự chọn và tiền xử lý. Mục đích là cung cấp môi trường cây thư mục doanh nghiệp nhiều tầng thực tế, kích thước tệp biến thiên, phản ánh tải làm việc thực tế để đo tốc độ quét và kiểm tra khả năng bắt lỗi đối kháng, hoàn toàn không phải ngụy tạo số liệu tác giả."),

        ("Câu hỏi 4: Tại sao mã kiểm tra lỗi CRC32 bị qua mặt bằng GF(2), trong khi SHA-256 lại phát hiện được gian lận?",
         "Luận điểm phản biện:\n"
         "1. Tính tuyến tính của CRC32: CRC32 hoạt động dựa trên phép chia đa thức trên trường nhị phân GF(2). Do tính chất tuyến tính, giá trị CRC của thông điệp sửa đổi có thể tính trước được. Đối thủ giải hệ phương trình đại số tuyến tính để tìm ra 4 byte bù (patch) nối vào đuôi tệp, khiến CRC32 của tệp giả mạo trùng khớp với tệp gốc trong kịch bản kiểm thử này.\n"
         "2. Tính phi tuyến của SHA-256: SHA-256 xây dựng trên cấu trúc lặp Merkle-Damgard với các hàm phi tuyến. Hàm băm đạt hiệu ứng tuyết lở (SAC ~50%): Chỉ đổi 1 bit đầu vào thì trung bình 50% số bit đầu ra bị đảo ngẫu nhiên. Bài toán tìm tiền ảnh thứ hai trên SHA-256 có độ phức tạp 2^256, vượt quá năng lực tính toán của đối thủ thời gian đa thức (PPT)."),

        ("Câu hỏi 5: Tấn công Kênh kề thời gian (Timing Side-Channel) trong so sánh chuỗi hoạt động thế nào và cách giải quyết?",
         "Luận điểm phản biện:\n"
         "1. Cơ chế rò rỉ: Hàm so sánh chuỗi thông thường (toán tử ==) so sánh tuần tự từng ký tự từ trái sang phải và ngắt sớm (early return) khi thấy ký tự sai đầu tiên. Đối thủ gửi các chuỗi thử và đo thời gian phản hồi: Thời gian xử lý lâu hơn nghĩa là đã đoán đúng thêm ký tự. Nhờ đó, đối thủ dò ra khóa bí mật mà không cần giải mã.\n"
         "2. Giải pháp kỹ thuật: Sử dụng hàm so sánh thời gian hằng số (Constant-time comparison) `hmac.compare_digest`. Hàm này luôn duyệt hết độ dài chuỗi bằng phép XOR tích lũy, đảm bảo thời gian xử lý độc lập với vị trí sai khác, triệt tiêu kênh rò rỉ thời gian ở mức phần mềm."),

        ("Câu hỏi 6: Cây Merkle phân cấp và phân tách miền RFC 6962 đóng vai trò gì trong đề tài?",
         "Luận điểm phản biện:\n"
         "1. Đóng góp độc lập: Halpin không đề xuất Cây Merkle hay chuẩn RFC 6962. Đây là phần mở rộng độc lập của nhóm nhằm giải quyết bài toán toàn vẹn dữ liệu tệp tin.\n"
         "2. Phân tách miền RFC 6962: Phép ghép băm thông thường H(L || R) có nguy cơ dẫn đến lỗi nhập nhằng nút lá (Node-leaf ambiguity) khi một nút lá bị giả mạo thành cặp băm của nút nhánh nội bộ. Nhóm áp dụng chuẩn RFC 6962: Thêm tiền tố 0x00 cho lá và 0x01 cho nhánh, triệt tiêu hoàn toàn sự nhập nhằng cấu trúc ở cấp độ cây.\n"
         "3. Định hướng cuối kỳ: Nhóm tiếp tục tối ưu hóa cơ chế bộ đệm cập nhật cục bộ và đo đạc thực nghiệm trên các quy mô dữ liệu lớn hơn.")
    ]

    for q_title, q_ans in qa_list:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(8)
        p_q.paragraph_format.space_after = Pt(2)
        r_q = p_q.add_run(q_title)
        r_q.font.bold = True
        r_q.font.size = Pt(11)
        r_q.font.color.rgb = RGBColor(15, 23, 42)

        p_a = doc.add_paragraph()
        p_a.paragraph_format.space_before = Pt(2)
        p_a.paragraph_format.space_after = Pt(8)
        p_a.paragraph_format.left_indent = Inches(0.2)
        r_a = p_a.add_run(q_ans)
        r_a.font.size = Pt(10)
        r_a.font.color.rgb = RGBColor(51, 65, 85)

    # Lưu file Word tại docs/ và sao chép ra thư mục gốc
    target_docs = "docs/Kich_Ban_Thuyet_Trinh_Va_Demo_DT005.docx"
    target_root = "Kich_Ban_Thuyet_Trinh_Va_Demo_DT005.docx"
    doc.save(target_docs)
    shutil.copyfile(target_docs, target_root)
    print(f"[THÀNH CÔNG] Đã tạo tài liệu Word hoàn chỉnh tại:\n - {target_docs}\n - {target_root}")

if __name__ == "__main__":
    create_document()
