# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word (.docx) Kịch Bản Thuyết Trình & Hướng Dẫn Diễn Tập Demo (Bản Chuẩn Dễ Hiểu & Tự Nhiên):
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn biên tập:
1. Giúp sinh viên chưa chuyên sâu mật mã học dễ hiểu, dễ nhớ và tự tin nói thành lời trước lớp.
2. Bổ sung ví dụ đời thường, trực quan (dấu vân tay số cho SHA-256, con dấu niêm phong đối chiếu cho HMAC, cây gia phả/hóa đơn tổng cho Merkle Tree, đo thời gian tiếng click ổ khóa cho Timing Attack, đếm số chữ cho CRC32).
3. Câu văn tự nhiên, có chủ ngữ, có mạch nối ý rõ ràng; không ép câu cụt ngủn máy móc.
4. Giữ đúng các thuật ngữ chuyên ngành bắt buộc (SHA-256, HMAC, Merkle Tree, PPT, Dolev-Yao, Random Oracle), có giải thích ngắn khi xuất hiện lần đầu.
5. Giữ nguyên cấu trúc 3 phần, 13 slide, số liệu 2.485 tệp, trích dẫn học thuật và chuẩn định dạng Word.
6. Chuẩn hóa các khẳng định tuyệt đối (thay "triệt tiêu hoàn toàn", "an toàn tuyệt đối", "ngay lập tức" bằng các giới hạn kỹ thuật chính xác).
7. Minh bạch rạch ròi giữa nội dung lý thuyết của bài báo gốc và phần thực nghiệm độc lập do sinh viên tự thiết kế.
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
    # PHẦN 1: NỘI DUNG TRÊN SLIDE & KỊCH BẢN THUYẾT MINH CHI TIẾT (13 SLIDE)
    # =========================================================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("PHẦN 1: NỘI DUNG TRÊN SLIDE & KỊCH BẢN THUYẾT MINH CHI TIẾT (13 SLIDE)")
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
                ("Thông tin thực hiện:", "Học phần: An toàn và bảo mật hệ thống thông tin | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143.",
                 "Lời thoại mở đầu: 'Kính chào Thầy và các bạn. Hôm nay em xin trình bày đồ án giữa kỳ DT005. Đồ án của em xuất phát từ việc khảo cứu bài báo khoa học của TS. Harry Halpin xuất bản năm 2025 trên tạp chí Journal of Cybersecurity, từ đó nhóm xây dựng một hệ thống thực nghiệm kiểm tra toàn vẹn dữ liệu để đối chứng các luận điểm của tác giả.'")
            ],
            "footer": "Bản trình chiếu chính thức phục vụ nghiệm thu đồ án giữa kỳ."
        },

        # SLIDE 2
        {
            "num": 2,
            "title": "BỐI CẢNH, MỤC TIÊU & BA ĐÓNG GÓP CỦA BÀI BÁO",
            "time": "1:00",
            "header": "SECTION 1 — GIỚI THIỆU & MỤC TIÊU BÀI BÁO (TRANG 1–2)",
            "lead": "Góc nhìn triết học về Kẻ tấn công trong Mật mã học và Chương trình nghiên cứu tương lai",
            "bullets": [
                ("1. Bối cảnh mật mã học:", "Mật mã học là ngành khoa học bảo vệ dữ liệu bí mật khi truyền qua môi trường mạng mở có kẻ tấn công chủ đích rình rập."),
                ("2. Nghịch lý nghiên cứu:", "Lý thuyết thông tin của Shannon (1948) được nghiên cứu triết học rất kỹ, nhưng mật mã học lại bị lãng quên suốt nhiều thập kỷ.",
                 "Lời giải thích dễ nhớ: Halpin chỉ ra rằng lý thuyết thông tin chỉ xử lý nhiễu ngẫu nhiên của đường truyền; còn mật mã học mới là nơi trực tiếp đối đầu với kẻ thù có mưu mô và chủ đích phá hoại."),
                ("3. Hạn chế khi chỉ nhìn trên lý thuyết:", "Nếu chỉ xem mật mã là các phép toán thuần túy trên giấy, chúng ta sẽ có điểm mù trước các đòn tấn công vật lý ngoài đời thực.",
                 "Ví dụ gần gũi: Một cánh cửa có ổ khóa toán học cực kỳ vững chắc, nhưng kẻ trộm không bẻ khóa mà dùng xà beng cạy bản lề hoặc đo độ trễ khi chìa xoay. Trong máy tính, kẻ tấn công đo thời gian CPU xử lý và bộ nhớ đệm (Cache) để mò ra khóa."),
                ("4. Ba đóng góp học thuật của Halpin:", "Tác giả hệ thống hóa lịch sử các mô hình kẻ tấn công; Phê phán giới hạn của việc toán học hóa quá mức; và Đề xuất 4 câu hỏi nghiên cứu liên ngành cho tương lai.")
            ],
            "footer": "Trích dẫn: Harry Halpin (2025), Section 1: Introduction, Journal of Cybersecurity, tr. 1–2."
        },

        # SLIDE 3
        {
            "num": 3,
            "title": "CƠ SỞ TRIẾT HỌC: THÔNG TIN, BÍ MẬT VÀ TÁC NHÂN",
            "time": "1:00",
            "header": "SECTION 2 & 3 — CƠ SỞ TRIẾT HỌC VỀ THÔNG TIN (TRANG 2–5)",
            "lead": "Khung lý thuyết của Fred Dretske và Brian Cantwell Smith ứng dụng trong Mật mã học",
            "bullets": [
                ("1. Bản thể học thông tin (Brian Cantwell Smith):", "Dữ liệu ban đầu vốn là dòng chảy vô định; con người tạo ra sự an toàn bằng cách phân định rõ ranh giới các đối tượng ('đăng ký đối tượng' — registration of objects).",
                 "Lời giải thích dễ nhớ: Trong hệ thống mật mã, ta phải tách bạch rạch ròi 3 thực thể: Bản rõ cần bảo vệ (M), Bản mã gửi đi (C) và Chìa khóa bí mật (K)."),
                ("2. Thuyết thông tin khách quan (Fred Dretske):", "Thông tin tồn tại khách quan theo quy luật tự nhiên, độc lập với người nhận; bản rõ mang thông tin gốc, còn hệ mật phải bảo đảm bản mã không để lộ bất kỳ thông tin nào về nguồn.",
                 "Ví dụ gần gũi: Giống như việc xóa sạch dấu vết hiện trường để kẻ lạ nhìn vào bản mã chỉ thấy một chuỗi ngẫu nhiên hỗn loạn, không thể đoán được nội dung gốc là gì."),
                ("3. Bản chất của kẻ tấn công:", "Đối thủ là tác nhân có chủ đích (intentional agent), có mục tiêu phá hoại rõ ràng, hoàn toàn khác với sự cố chập chờn hay nhiễu đường truyền vật lý."),
                ("4. Bản chất của Bí mật & Tri thức (David Lewis):", "Phân biệt giữa Tri thức chung (Common Knowledge — tri thức ai cũng biết và biết người khác cùng biết), Tri thức công khai (Public Knowledge) và Bí mật (thông tin nội bộ chỉ số ít người nắm giữ).",
                 "Lời giải thích dễ nhớ: Mật mã học sinh ra để biến thông điệp thành bí mật: chỉ người nào có chìa khóa hợp lệ mới giải mã được thông tin.")
            ],
            "footer": "Trích dẫn khoa học: Dretske (1981, Ref [1]); Cantwell Smith (1996, Ref [5]); Halpin (2025), tr. 2–5."
        },

        # SLIDE 4
        {
            "num": 4,
            "title": "TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ TRONG MẬT MÃ",
            "time": "1:00",
            "header": "SƠ ĐỒ 1: DÒNG THỜI GIAN TIẾN HÓA MÔ HÌNH ĐỐI THỦ",
            "lead": "Từ mô hình thông tin của Shannon đến độ phức tạp tính toán và bài học thực tế sau vụ Snowden",
            "bullets": [
                ("1. Giai đoạn 1 (1948–1949) — Mô hình Shannon:", "Mô hình bảo mật điểm-điểm; định nghĩa Bí mật hoàn hảo H(K) >= H(M). Giới hạn: Trừu tượng hóa đối thủ trên kênh toán học, bỏ qua chi phí phân phối khóa.",
                 "Lời giải thích dễ nhớ: Shannon chứng minh mã hóa dùng khóa một lần (One-Time Pad) đạt độ an toàn lý thuyết tuyệt đối, nhưng không khả thi vì khóa phải dài bằng văn bản và không bao giờ được dùng lại."),
                ("2. Giai đoạn 2 (1976) — Cách mạng Khóa công khai:", "Diffie-Hellman và Merkle giải bài toán phân phối khóa bằng các bài toán toán học một chiều khó đảo ngược (như tính lũy thừa thì nhanh nhưng tìm logarit rời rạc thì cực khó)."),
                ("3. Giai đoạn 3 (1982–1984) — Kẻ thù PPT & Trò chơi an ninh:", "Goldwasser-Micali mô hình hóa đối thủ có năng lực tính toán đa thức xác suất (PPT — Probabilistic Polynomial-Time, tức đối thủ có máy tính mạnh nhưng tài nguyên và thời gian hữu hạn). Độ an toàn được kiểm chứng qua các kịch bản thử thách (Security Games).",
                 "Lời giải thích dễ nhớ: Chuyển từ an toàn tuyệt đối trên lý thuyết sang an toàn thực tế: kẻ tấn công dù có siêu máy tính cũng không thể giải nổi bài toán trong thời gian cho phép."),
                ("4. Giai đoạn 4 (1981–1993) — Dolev-Yao & Hộp đen ngẫu nhiên (ROM):", "Dolev-Yao mô hình hóa kẻ tấn công kiểm soát toàn bộ kênh truyền mạng; Bellare-Rogaway đề xuất mô hình Random Oracle coi hàm băm như một cỗ máy ngẫu nhiên lý tưởng."),
                ("5. Giai đoạn 5 (2013–2025) — Bài học thực tế từ Snowden:", "Vụ lộ lọt của Snowden phơi bày sức mạnh thực tế của các cơ quan tình báo: Tấn công kênh kề vật lý, cài cửa sau vào tiêu chuẩn quốc tế và thu thập siêu dữ liệu người dùng.")
            ],
            "footer": "Sơ đồ tổng hợp dựa trên Section 3–6, Harry Halpin (2025), Journal of Cybersecurity, tr. 3–12."
        },

        # SLIDE 5
        {
            "num": 5,
            "title": "MÔ HÌNH TRUYỀN THÔNG CỦA SHANNON VÀ MẠNG THỰC TẾ",
            "time": "1:00",
            "header": "SƠ ĐỒ 2: ĐỐI SÁNH MÔ HÌNH TRUYỀN THÔNG ĐIỂM-ĐIỂM VÀ MẠNG XÃ HỘI-KỸ THUẬT",
            "lead": "Khoảng cách giữa mô hình điểm-điểm khép kín của Shannon và môi trường mạng Internet thực tế",
            "bullets": [
                ("Panel Trái — Mô hình lý thuyết của Shannon (1949):", "Alice gửi bản rõ M đã mã hóa bằng khóa K đến Bob; kẻ nghe lén Eve chỉ thụ động đón bắt đường truyền. Mô hình này bỏ qua chi phí quản lý khóa và bề mặt tấn công vật lý."),
                ("Panel Phải — Mạng Internet thực tế ngày nay (Halpin 2025):", "Dữ liệu luân chuyển qua nhà mạng, máy chủ đám mây. Đối thủ quy mô quốc gia (như NSA) tấn công vào 4 hướng: Kênh kề vật lý, phân tích siêu dữ liệu, cài cửa sau phần cứng và nội gián.",
                 "Ví dụ gần gũi về siêu dữ liệu: Kẻ tấn công không cần bẻ khóa đọc thư, mà chỉ cần nhìn vào người gửi, người nhận, thời gian và kích thước gói tin là đã biết rõ bạn đang liên lạc với ai và làm gì.")
            ],
            "footer": "Sơ đồ tổng hợp dựa trên Shannon (1949) và Harry Halpin (2025), Journal of Cybersecurity, tr. 3–5."
        },

        # SLIDE 6
        {
            "num": 6,
            "title": "BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI",
            "time": "1:00",
            "header": "SECTION 4 — HÌNH THỨC HÓA ĐỐI THỦ (TRANG 5–8)",
            "lead": "Ba cách toán học hóa kẻ thù: Độ phức tạp tính toán đa thức, Logic kênh truyền và Tiên tri ngẫu nhiên",
            "bullets": [
                ("1. Kẻ tấn công bị giới hạn tính toán (PPT — Goldwasser-Micali 1984):", "Mô hình hóa đối thủ như một thuật toán chạy trong thời gian đa thức xác suất (PPT). Độ an toàn được đo bằng xác suất kẻ địch thắng trong trò chơi thách đấu."),
                ("2. Kẻ tấn công mạng logic (Dolev-Yao 1983):", "Giả định đối thủ kiểm soát toàn bộ đường truyền (bắt, sửa, xóa, gửi lại gói tin), nhưng lại xem mật mã là chiếc hộp đen an toàn tuyệt đối nếu không có khóa.",
                 "Điểm mù của Dolev-Yao: Mô hình này ngây thơ coi mật mã là hộp đen hoàn hảo trên giấy, bỏ qua các lỗ hổng rò rỉ vật lý khi phần mềm chạy trên máy tính thật."),
                ("3. Tiên tri ngẫu nhiên (Random Oracle — Bellare-Rogaway 1993):", "Giả định hàm băm hoạt động như một cỗ máy ngẫu nhiên lý tưởng; nhưng thực tế không hàm băm máy tính nào đạt được sự ngẫu nhiên toán học tuyệt đối này.",
                 "Lời giải thích dễ nhớ: SHA-256 là thuật toán xác định chạy trên CPU máy thật, không phải cỗ máy ngẫu nhiên lý tưởng, nên vẫn chịu tác động từ môi trường thực thi phần cứng.")
            ],
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
                ("Bước 1 — Giả định trên lý thuyết:", "Coi mật mã là hộp đen lý tưởng, thời gian chạy không bị đo lường, độ dài bản mã không làm lộ thông tin, đường truyền độc lập."),
                ("Bước 2 — Thực tế vật lý bị bỏ quên:", "Thời gian xử lý CPU, điện năng tiêu thụ, bức xạ điện từ, bộ nhớ đệm (Cache) và lỗi phần mềm."),
                ("Bước 3 — Đòn đánh của kẻ tấn công thực tế:", "Đo chênh lệch thời gian chạy để mò khóa (tấn công kênh kề — Kocher 1996), phân tích lưu lượng gói tin và khai thác sơ hở quản lý khóa.",
                 "Ví dụ gần gũi: Giống như kẻ trộm lắng nghe tiếng click của ổ khóa số để dò từng nấc mã; kẻ tấn công đo sự chênh lệch vài phần triệu giây khi CPU so sánh chuỗi để mò ra từng ký tự của mật khẩu."),
                ("Bước 4 — Kết luận của Halpin:", "Quá trừu tượng hóa trên lý thuyết khiến mật mã học bỏ sót chính những vũ khí nguy hiểm nhất của kẻ tấn công ngoài đời thực.")
            ],
            "footer": "Sơ đồ tổng hợp dựa trên Kocher (1996) và Halpin (2025), Section 5, tr. 8–10."
        },

        # SLIDE 8
        {
            "num": 8,
            "title": "SNOWDEN VÀ ĐỐI THỦ THỂ CHẾ (THE REAL ADVERSARY)",
            "time": "1:00",
            "header": "SECTION 6 — TIẾT LỘ CỦA SNOWDEN (TRANG 10–12)",
            "lead": "Thực tế an ninh trước đối thủ giám sát quy mô quốc gia",
            "bullets": [
                ("1. Bài học thực tiễn từ Snowden (2013):", "Kẻ tấn công không mất công bẻ khóa toán học RSA hay AES, mà tấn công thẳng vào hạ tầng xung quanh (đánh cắp khóa tại máy chủ, thâm nhập cáp quang biển)."),
                ("2. Năng lực của đối thủ thể chế:", "Khai thác siêu dữ liệu để vẽ bản đồ liên lạc; cài cửa sau vào tiêu chuẩn bảo mật quốc tế NIST (vụ bê bối Dual_EC_DRBG); gây sức ép pháp lý lên các tập đoàn công nghệ.",
                 "Lời giải thích dễ nhớ: Vụ bê bối Dual_EC_DRBG chứng minh cơ quan tình báo có thể bí mật cài điểm yếu toán học vào bộ sinh số ngẫu nhiên chuẩn quốc tế để đọc trộm thông tin toàn cầu."),
                ("3. Khái niệm 'Đối thủ' mang tính vị thế:", "Kẻ thù tùy thuộc vào góc nhìn; ngành mật mã buộc phải nghiên cứu cơ chế phòng chống cửa sau (Cliptography).",
                 "Lời giải thích dễ nhớ: Cliptography là hướng nghiên cứu thiết kế thuật toán tự bảo vệ an toàn ngay cả khi phần mềm hay thiết bị đã bị cài sẵn cửa sau từ trước.")
            ],
            "footer": "Trích dẫn: Harry Halpin (2025), Section 6: Snowden and the real adversary, tr. 10–12."
        },

        # SLIDE 9
        {
            "num": 9,
            "title": "CHƯƠNG TRÌNH NGHIÊN CỨU TƯƠNG LAI: BỐN CÂU HỎI MỞ",
            "time": "1:00",
            "header": "SECTION 7 — KẾT LUẬN & CHƯƠNG TRÌNH NGHIÊN CỨU (TRANG 12–14)",
            "lead": "Đề xuất dung hòa giữa trừu tượng hóa toán học và dữ liệu thực nghiệm thực tế",
            "bullets": [
                ("RQ1 — Mạng xã hội động:", "Nghiên cứu hành vi giữ bí mật thực tế của con người; đánh giá độ bền của chính sách an ninh trước sự cố nội bộ hoặc kẻ phản bội."),
                ("RQ2 — Đo lường bằng thống kê thực nghiệm:", "Dùng thống kê thực nghiệm và mô phỏng xác suất (Monte Carlo) để định lượng chính xác sức mạnh thật sự của đối thủ và liên minh giám sát."),
                ("RQ3 — Phương pháp tiếp cận liên ngành:", "Ứng dụng lý thuyết hệ thống phức hợp và mô hình lây lan dịch bệnh để phân tích đường đi của thông tin rò rỉ.",
                 "Ví dụ gần gũi: Xem sự rò rỉ thông tin lây lan giống như một đợt dịch bệnh giúp các chuyên gia an ninh dự báo trước đường đi và khoanh vùng dập tắt nguy cơ nhanh hơn."),
                ("RQ4 — Đánh giá lại các giả định nền tảng:", "Soi xét lại các tiên đề lý thuyết từ thời Shannon để điều chỉnh cho phù hợp với môi trường an ninh mạng hiện đại.")
            ],
            "footer": "Trích dẫn nguyên văn 4 câu hỏi nghiên cứu mở: Harry Halpin (2025), Section 7, tr. 12–14."
        },

        # SLIDE 10
        {
            "num": 10,
            "title": "KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG: RED TEAM VS BLUE TEAM",
            "time": "0:45",
            "header": "SƠ ĐỒ 4: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG (PRE-DEMO ARCHITECTURE)",
            "lead": "Mô hình đối kháng 2 màn hình do sinh viên tự thiết kế trên tập dữ liệu kaggle_dataset (2.485 tệp tin) để kiểm chứng luận điểm bài báo",
            "bullets": [
                ("Khối 1 (Trái) — Đội Phòng thủ Blue Team (Port 8080):", "Thiết lập mốc tin cậy sạch ban đầu H_0; tích hợp phân tách miền chuẩn RFC 6962 (tiền tố 0x00 cho lá, 0x01 cho nhánh).",
                 "Lời giải thích dễ nhớ: Phân tách miền 0x00/0x01 giúp phòng chống hiệu quả nguy cơ kẻ tấn công đánh tráo một nút lá thành cặp băm của nút nhánh trên cây Merkle (Node-Leaf Confusion)."),
                ("Khối 2 (Giữa) — Kho dữ liệu kiểm thử (kaggle_dataset):", "Bộ dữ liệu thực tế gồm 2.485 tệp tin (~112,8 MB) do nhóm sinh viên tự chọn để mô phỏng cây thư mục doanh nghiệp trên đĩa cứng."),
                ("Khối 3 (Phải) — Đội Tấn công Red Team (Port 8081):", "Mô phỏng đối thủ chủ đích: Kịch bản 1 (Sửa tệp tin trực tiếp) và Kịch bản 2 (Bù 4 byte GF(2) qua mặt mã kiểm tra CRC32).",
                 "Lời giải thích dễ nhớ: Kịch bản 2 chứng minh kẻ tấn công có chủ đích dễ dàng đánh lừa mã tuyến tính CRC32, nhưng hoàn toàn bất lực trước hàm băm phi tuyến SHA-256."),
                ("Khối 4 (Dưới) — Cơ chế khôi phục an toàn:", "Thư mục sao lưu an toàn .adversary_backups/ và nút Rollback All giúp phục hồi dữ liệu về nguyên trạng ban đầu.")
            ],
            "footer": "Sơ đồ kiến trúc diễn tập đối kháng do sinh viên tự xây dựng nhằm minh họa thực nghiệm các luận điểm trong bài báo."
        },

        # SLIDE 11
        {
            "num": 11,
            "title": "SLIDE CHUYỂN TIẾP: DEMO",
            "time": "0:15",
            "header": "GIAO DIỆN TỐI GIẢN NỀN SÁNG CHUẨN BỊ DIỄN TẬP TRỰC TIẾP",
            "lead": "Chữ hiển thị trên Slide: DEMO",
            "bullets": [
                ("Giao diện chuyển tiếp:", "DEMO — Chuyển sang màn hình thực nghiệm diễn tập đối kháng giữa Blue Team và Red Team.")
            ],
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
                ("1. Bản chất Thông tin & Kẻ thù (Mục 1–2):", "Mô hình cũ: Coi lỗi chỉ là do nhiễu ngẫu nhiên đường truyền ↔ Halpin: Thông tin mang tính khách quan (Dretske), đối thủ là kẻ có chủ đích phá hoại.",
                 "Bài học kỹ thuật: Vì kẻ tấn công có chủ đích tính toán trước nên ta bắt buộc phải dùng hàm băm mật mã SHA-256 (như dấu vân tay số độc nhất) thay vì dùng mã CRC32 đơn giản."),
                ("2. Giới hạn hình thức hóa (Mục 5):", "Mô hình cũ: Mật mã là chiếc hộp đen toán học hoàn hảo ↔ Halpin: Kẻ địch đo thời gian xử lý CPU và bộ nhớ đệm Cache ngoài đời thực để lấy cắp khóa.",
                 "Bài học kỹ thuật: Lập trình phòng thủ đòi hỏi phải dùng phép so sánh thời gian hằng số (hmac.compare_digest) để triệt tiêu nguy cơ lộ lọt qua kênh kề thời gian."),
                ("3. Đối thủ thế giới thực (Mục 6):", "Mô hình cũ: Giới hạn đối thủ trong năng lực tính toán thông thường (PPT) ↔ Halpin: Đối thủ thể chế (NSA) cài cửa sau vào chuẩn bảo mật và khai thác siêu dữ liệu diện rộng."),
                ("4. Chương trình nghiên cứu mở (Mục 7):", "Mô hình cũ: Mật mã thuần túy là toán học rời rạc trên giấy ↔ Halpin: 4 câu hỏi mở kết hợp liên ngành với thống kê thực nghiệm và triết học.")
            ],
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
                ("Thông điệp kết thúc:", "Q & A — Em xin cảm ơn Thầy và các bạn đã chú ý theo dõi. Rất mong nhận được câu hỏi phản biện và góp ý từ Thầy để hoàn thiện đề tài."),
                ("Thông tin đề tài:", "Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143.")
            ],
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

        # Khối nội dung trên Slide
        p_lbl_sc = doc.add_paragraph()
        p_lbl_sc.paragraph_format.space_before = Pt(2)
        p_lbl_sc.paragraph_format.space_after = Pt(2)
        p_lbl_sc.paragraph_format.left_indent = Inches(0.2)
        r_lsc = p_lbl_sc.add_run("Nội dung Slide & Mở rộng:")
        r_lsc.font.size = Pt(10)
        r_lsc.font.bold = True
        r_lsc.font.color.rgb = RGBColor(15, 23, 42)

        for item in s['bullets']:
            h = item[0]
            b = item[1]
            exp = item[2] if len(item) > 2 else ""

            # Dòng nội dung hiển thị trên Slide
            p_it = doc.add_paragraph()
            p_it.paragraph_format.space_before = Pt(3)
            p_it.paragraph_format.space_after = Pt(1)
            p_it.paragraph_format.left_indent = Inches(0.25)
            r_h = p_it.add_run(f"– {h} ")
            r_h.font.size = Pt(10)
            r_h.font.bold = True
            r_h.font.color.rgb = RGBColor(15, 23, 42)

            r_b = p_it.add_run(b)
            r_b.font.size = Pt(10)
            r_b.font.color.rgb = RGBColor(51, 65, 85)

            # Dòng mở rộng (giải thích tự nhiên cho khái niệm quan trọng)
            if exp:
                p_exp = doc.add_paragraph()
                p_exp.paragraph_format.space_before = Pt(0)
                p_exp.paragraph_format.space_after = Pt(3)
                p_exp.paragraph_format.left_indent = Inches(0.45)
                
                r_arr = p_exp.add_run("↳ Mở rộng: ")
                r_arr.font.size = Pt(9.5)
                r_arr.font.bold = True
                r_arr.font.color.rgb = RGBColor(30, 64, 175)
                
                r_exp = p_exp.add_run(exp)
                r_exp.font.size = Pt(9.5)
                r_exp.font.italic = True
                r_exp.font.color.rgb = RGBColor(71, 85, 105)

        # Footer của slide
        p_ft = doc.add_paragraph()
        p_ft.paragraph_format.space_before = Pt(3)
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
        "Mục tiêu diễn tập giữa kỳ: Minh họa trực quan 2 kịch bản đối kháng then chốt để kiểm chứng luận điểm của bài báo Harry Halpin (2025). "
        "Bài báo gốc là công trình triết học lý thuyết, không có mã nguồn; toàn bộ hệ thống demo đối kháng 2 màn hình này do nhóm sinh viên tự thiết kế "
        "nhằm đưa các khái niệm trừu tượng vào kiểm chứng thực nghiệm trên máy tính thật."
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
         "• Blue Team Service (Phía phòng thủ):  python -m src.demo.web_demo --port 8080\n"
         "• Red Team Service (Phía tấn công):   python -m src.demo.red_team_web --port 8081\n"
         "Chia đôi màn hình trình duyệt: Nửa trái mở http://localhost:8080 (Blue Team); nửa phải mở http://localhost:8081 (Red Team)."),

        ("Bước 2: Thiết lập mốc tin cậy sạch tại Blue Team (Port 8080)",
         "Tại cổng 8080, chọn thư mục 'kaggle_dataset' (~112.8 MB gồm 2.485 tệp tin).\n"
         "Bấm 'Nạp & Tính toán Toàn vẹn'. Hệ thống tính ra Root Hash ban đầu H_0 sạch với trạng thái xanh: '100% NGUYÊN VẸN'.\n"
         "Ý tưởng trực quan & Điểm nói lý thuyết: Giống như việc chụp lại dấu vân tay số của toàn bộ thư viện tệp tin. Cây Merkle hoạt động như một hóa đơn tổng: "
         "nó băm từ các tệp lá lên gốc, giúp quản lý hàng nghìn tệp cực nhanh. Hệ thống áp dụng chuẩn RFC 6962 (thêm tiền tố byte 0x00 cho lá, 0x01 cho nhánh), "
         "giúp phòng chống hiệu quả nguy cơ kẻ gian làm giả cấu trúc cây (Node-Leaf Confusion) — đúng theo tinh thần cảnh báo của Halpin tại Mục 5."),

        ("Bước 3: Phát động Kịch bản 1 — Can thiệp tệp tin tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm 'Can thiệp tệp tin ngẫu nhiên' (In-place Tampering) để cố tình sửa tệp Resume/Resume.csv trên đĩa cứng.\n"
         "Quay lại Blue Team (Port 8080), bấm 'Kiểm tra toàn vẹn':\n"
         "• Hệ thống đổi sang cảnh báo đỏ: 'CẢNH BÁO: PHÁT HIỆN CAN THIỆP DỮ LIỆU!'.\n"
         "• Bảng đối soát chỉ rõ chính xác tệp Resume/Resume.csv đã bị thay đổi.\n"
         "Ý nghĩa khi nói: SHA-256 hoạt động như một dấu vân tay số cực kỳ nhạy. Dù kẻ tấn công chỉ thay đổi đúng một dấu phẩy trong tệp tin, "
         "mã băm đầu ra cũng biến đổi hoàn toàn và hệ thống phát hiện chính xác ngay khi đối chiếu."),

        ("Bước 4: Phát động Kịch bản 2 — Bù 4 byte GF(2) qua mặt CRC32 tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm 'Tấn công bù 4 byte GF(2)' (CRC32 Forge Attack). Động cơ Red Team giải hệ phương trình tuyến tính, làm sai lệch nội dung tệp nhưng chèn thêm 4 byte bù khiến mã CRC32 không hề đổi.\n"
         "Quay lại Blue Team (Port 8080), bấm 'Kiểm tra toàn vẹn':\n"
         "• Bảng đối chứng hiển thị: Thuật toán CRC32 bị qua mặt hoàn toàn (báo xanh hợp lệ sai do tính chất tuyến tính).\n"
         "• Ngược lại, hàm băm SHA-256 phát hiện gian lận chính xác nhờ tính phi tuyến mạnh.\n"
         "Ví dụ trực quan & Điểm nói lý thuyết: Mã CRC32 giống như việc đếm tổng số chữ cái trong bức thư — kẻ gian bớt chữ này và bù chữ khác cho đủ số thì qua mặt được. "
         "Đúng như Halpin phân tích tại Mục 1 & 2: Mã tuyến tính như CRC32 chỉ bắt được lỗi ngẫu nhiên vô tình, hoàn toàn bất lực trước kẻ thù có chủ đích. "
         "Còn SHA-256 là hàm băm phi tuyến có hiệu ứng tuyết lở (~50% bit đảo khi sửa 1 bit), kẻ địch không thể dùng đại số giải phương trình để bù mã được."),

        ("Bước 5: Hoàn nguyên dữ liệu sạch (Rollback All)",
         "Tại giao diện Red Team (Port 8081), bấm 'ROLLBACK ALL' để khôi phục toàn bộ tệp từ thư mục sao lưu an toàn .adversary_backups/.\n"
         "Quay lại Blue Team bấm kiểm tra lại: Hệ thống trở về trạng thái an toàn màu xanh nguyên trạng.")
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
        "   Kết quả & Điểm nói lý thuyết: Đo đạc tính phi tuyến của SHA-256, kiểm chứng hiệu ứng tuyết lở SAC đạt xấp xỉ 50% bit đảo, "
        "chứng minh kịch bản bẻ gãy mã tuyến tính CRC32 trên trường GF(2), và kiểm chứng rò rỉ kênh kề thời gian (Kocher 1996 - Mục 5 bài báo) "
        "khi so sánh chuỗi ngắt sớm so với hàm thời gian hằng số hmac.compare_digest.\n\n"
        "2. Diễn tập 2 Terminal Red Team vs Blue Team:\n"
        "   • Terminal 1 (Red Team): Chạy 'python src/adversary/red_team_cli.py'. Chọn phím '1' để can thiệp tệp, hoặc phím '4' để bù mã CRC32.\n"
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
         "1. Tính toàn vẹn là điều kiện tiên quyết: Trong tam giác an ninh CIA (Bảo mật - Toàn vẹn - Khả dụng), Halpin nhấn mạnh Tính toàn vẹn (Integrity) là nền tảng số một. "
         "Một tệp tin hay thông điệp có thể không cần giữ kín bí mật, nhưng bắt buộc phải toàn vẹn thì người nhận mới có thể tin cậy sử dụng (ví dụ: một bản cập nhật phần mềm nếu bị sửa đổi chèn mã độc thì cực kỳ nguy hiểm dù mã nguồn có công khai hay không).\n"
         "2. Đưa triết lý vào kiểm định thực tế: Halpin phê phán việc toán học hóa trừu tượng đã làm mật mã học xa rời thực tế vật lý. "
         "Bài toán kiểm tra toàn vẹn tệp tin trên đĩa cứng là môi trường thực nghiệm lý tưởng để đưa các khái niệm trừu tượng vào đo đạc thời gian, bộ nhớ và khả năng bắt lỗi đối kháng rõ ràng.\n"
         "3. Minh họa trực quan các mô hình kẻ tấn công: Hệ thống tệp cho phép mô phỏng sống động các dạng đối thủ được Halpin đề cập: "
         "Kẻ can thiệp đường truyền Dolev-Yao, kẻ đo thời gian rò rỉ Kocher và kẻ chủ đích giả mạo dữ liệu."),

        ("Câu hỏi 2: Bài báo có 134 tài liệu tham khảo, khi thẩm định và tái lập thì có bắt buộc phải làm theo toàn bộ không?",
         "Luận điểm phản biện:\n"
         "1. Bản chất là bài tổng quan học thuật: Bài báo của Halpin là bài khảo cứu triết học dài 17 trang, với 134 tài liệu tham khảo là kho tàng lịch sử từ năm 1883 đến 2024 "
         "(bao gồm cả sách triết học cổ điển, hồ sơ giải mật chính phủ về NSA, v.v.). Không có một công trình kỹ thuật nào có thể hay bắt buộc phải tái lập toàn bộ 134 tài liệu nền đó.\n"
         "2. Tinh lọc đúng các trụ cột an ninh cốt lõi: Nhóm đã tinh lọc 6 trụ cột kỹ thuật trực tiếp liên quan đến bài toán an ninh: "
         "Shannon (Mã tuyến tính vs Hàm băm), Goldwasser-Micali (Trò chơi an ninh PPT), Kocher (Kênh kề thời gian), Bellare-Rogaway (Random Oracle), "
         "Dolev-Yao (Mạng thù địch), Snowden (Đối thủ thể chế), đồng thời nhóm bổ sung chuẩn IETF RFC 6962 về phân tách miền để hoàn thiện giải pháp kỹ thuật."),

        ("Câu hỏi 3: Tập dữ liệu kaggle_dataset từ đâu ra? Có xuất hiện trong bài báo gốc không?",
         "Luận điểm phản biện:\n"
         "1. Khẳng định minh bạch: Bài báo gốc của Harry Halpin là công trình khảo luận nhận thức luận thuần túy, hoàn toàn KHÔNG có tập dữ liệu, bảng số liệu thực nghiệm hay mã nguồn đi kèm.\n"
         "2. Mục đích đưa tập dữ liệu vào: Bộ dữ liệu Kaggle Resume Dataset (2.485 tệp tin, ~112,8 MB) do nhóm sinh viên tự chọn và tiền xử lý. "
         "Mục đích là tạo môi trường cây thư mục doanh nghiệp nhiều tầng thực tế, kích thước tệp biến thiên, phản ánh tải làm việc thực tế để đo tốc độ quét và kiểm tra khả năng bắt lỗi đối kháng, "
         "hoàn toàn là phần đóng góp thực nghiệm độc lập của sinh viên, không phải ngụy tạo số liệu tác giả."),

        ("Câu hỏi 4: Tại sao mã kiểm tra lỗi CRC32 bị qua mặt bằng GF(2), trong khi SHA-256 lại phát hiện được gian lận?",
         "Luận điểm phản biện:\n"
         "1. Tính tuyến tính của CRC32: CRC32 hoạt động dựa trên phép chia đa thức trên trường nhị phân GF(2). "
         "Do tính chất tuyến tính, giá trị kiểm tra của dữ liệu sửa đổi có thể tính trước được bằng đại số. "
         "Kẻ tấn công giải hệ phương trình tuyến tính để tìm ra 4 byte bù (patch) nối vào đuôi tệp, khiến mã CRC32 của tệp giả mạo trùng khớp hoàn toàn với tệp gốc trong kịch bản kiểm thử này.\n"
         "2. Tính phi tuyến mạnh của SHA-256: SHA-256 xây dựng trên cấu trúc lặp Merkle-Damgard với các phép toán phi tuyến phức tạp. "
         "Hàm băm đạt hiệu ứng tuyết lở (SAC ~50%): Chỉ đổi đúng 1 bit đầu vào thì trung bình 50% số bit đầu ra bị đảo ngẫu nhiên. "
         "Bài toán tìm tiền ảnh thứ hai trên SHA-256 có độ phức tạp 2^256, vượt xa năng lực tính toán của kẻ tấn công thời gian đa thức (PPT)."),

        ("Câu hỏi 5: Tấn công Kênh kề thời gian (Timing Side-Channel) trong so sánh chuỗi hoạt động thế nào và cách giải quyết?",
         "Luận điểm phản biện:\n"
         "1. Cơ chế rò rỉ qua thời gian: Lệnh so sánh chuỗi thông thường (toán tử ==) so sánh tuần tự từng ký tự từ trái sang phải và ngắt sớm (early return) ngay khi thấy ký tự sai đầu tiên. "
         "Kẻ tấn công gửi các chuỗi thử và đo thời gian phản hồi: Thời gian xử lý lâu hơn nghĩa là đã đoán đúng thêm ký tự. Nhờ đó, kẻ địch dò ra khóa bí mật mà không cần giải mã.\n"
         "2. Giải pháp kỹ thuật triệt để: Sử dụng hàm so sánh thời gian hằng số (Constant-time comparison) `hmac.compare_digest`. "
         "Hàm này luôn duyệt hết độ dài chuỗi bằng phép toán XOR tích lũy, đảm bảo thời gian xử lý luôn không đổi dù chuỗi khớp hay không, loại bỏ hoàn toàn kênh rò rỉ thời gian ở mức phần mềm ứng dụng."),

        ("Câu hỏi 6: Cây Merkle phân cấp và phân tách miền RFC 6962 đóng vai trò gì trong đề tài?",
         "Luận điểm phản biện:\n"
         "1. Đóng góp kỹ thuật độc lập của nhóm: Tác giả Halpin chỉ nêu triết lý chứ không đề xuất Cây Merkle hay chuẩn RFC 6962. "
         "Đây là phần giải pháp kỹ thuật do nhóm chủ động tích hợp nhằm giải quyết bài toán kiểm tra toàn vẹn cây thư mục dữ liệu lớn.\n"
         "2. Phân tách miền RFC 6962: Phép ghép băm thông thường H(L || R) có nguy cơ dẫn đến lỗi nhầm lẫn cấu trúc cây (Node-leaf ambiguity) "
         "khi nút lá bị kẻ gian giả mạo thành cặp băm của nút nhánh nội bộ. Nhóm áp dụng chuẩn RFC 6962: Thêm tiền tố byte 0x00 cho lá và 0x01 cho nhánh, "
         "ngăn chặn hiệu quả sự nhập nhằng cấu trúc ở cấp độ cây Merkle.\n"
         "3. Định hướng cho đồ án cuối kỳ: Nhóm tiếp tục tối ưu hóa cơ chế bộ đệm cập nhật cục bộ và đo đạc thực nghiệm trên các quy mô dữ liệu lớn hơn.")
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

    # Lưu thành file Word chính thức duy nhất
    target_docs = "docs/Kich_Ban_Thuyet_Trinh_Va_Demo_DT005.docx"
    doc.save(target_docs)
    print(f"[THÀNH CÔNG] Đã cập nhật file Word chính thức tại: {target_docs}")

if __name__ == "__main__":
    create_document()
