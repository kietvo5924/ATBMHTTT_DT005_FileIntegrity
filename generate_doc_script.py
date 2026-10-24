# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word (.docx) Kịch Bản Thuyết Trình & Hướng Dẫn Diễn Tập Demo:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn cải tiến:
- Ghi lại đầy đủ, chính xác toàn bộ nội dung hiển thị TRÊN TỪNG SLIDE (15 slide chuẩn).
- KHÔNG giải thích tràn lan mọi câu chữ; CHỈ tập trung giải thích các khái niệm/thuật ngữ trừu tượng, khó hiểu.
- Lời giải thích dùng từ ngữ bình dân, trực quan, dễ hiểu, có ví dụ đời thường/kỹ thuật để người đọc hiểu bản chất ngay lập tức.
- Giữ nguyên cấu trúc học thuật, chuyên nghiệp, không xưng hô rườm rà.
- Đầy đủ 3 phần: (1) Nội dung 15 Slide & Giải thích dễ hiểu; (2) Kịch bản Demo tinh giản theo bài báo; (3) Bộ câu hỏi phản biện chuyên sâu (Q&A).
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
                "Cách mở đầu báo cáo ngắn gọn, tự nhiên:\n"
                "Báo cáo giữa kỳ của đề tài kết hợp hai phần: Một là thẩm định học thuật bài báo triết học về Kẻ thù trong mật mã học của TS. Harry Halpin; Hai là tự xây dựng một hệ thống demo thực tế trên máy tính để kiểm tra toàn vẹn tệp tin, làm bằng chứng thực nghiệm minh họa cho các luận điểm của bài báo."
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
                ("1. Bối cảnh mật mã học:", "Khoa học về truyền thông bí mật khi có sự hiện diện của đối thủ có chủ đích trên môi trường mạng mở."),
                ("2. Nghịch lý nhận thức luận:", "Lý thuyết thông tin của Shannon (1948) được triết học nghiên cứu sâu rộng, nhưng chính mật mã học lại bị triết học bỏ quên suốt nhiều thập kỷ."),
                ("3. Cái giá của sự trừu tượng hóa:", "Xem mật mã học chỉ là các bài toán toán học thuần túy trên giấy sẽ tạo ra điểm mù nguy hiểm trước các đòn tấn công vật lý trong thế giới thực."),
                ("4. Ba đóng góp học thuật của Halpin:", "Hệ thống hóa lịch sử các mô hình đối thủ; Phê phán giới hạn của việc toán học hóa; Đề xuất 4 câu hỏi nghiên cứu mở kết hợp liên ngành.")
            ],
            "explanation": (
                "Giải thích các thuật ngữ khó hiểu trên Slide:\n"
                "• 'Nghịch lý nhận thức luận' là gì? Nghĩa là từ trước đến nay, các nhà nghiên cứu rất hay bàn về triết học của thông tin (thông tin là gì, truyền tin thế nào). Nhưng họ lại quên mất một thực tế hiển nhiên: Khi truyền tin qua mạng thì luôn có kẻ xấu rình rập để nghe lén hoặc phá hoại. Mật mã học sinh ra chính là để đối phó với kẻ xấu đó, nhưng giới triết học lại bỏ quên góc nhìn này. Bài báo của Halpin ra đời để giải quyết đúng khoảng trống đó.\n"
                "• 'Cái giá của sự trừu tượng hóa' là gì? Nghĩa là nếu ta chỉ ngồi trên giấy chứng minh công thức toán học này an toàn tuyệt đối, ta sẽ bị mù trước thực tế. Kẻ trộm ngoài đời không cần đi bẻ công thức toán, mà chúng đo xem máy tính chạy mất bao nhiêu giây, ăn bao nhiêu điện để đoán ra mật khẩu. Coi an toàn chỉ là toán học sẽ khiến hệ thống sụp đổ trước các đòn đánh vật lý."
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
                ("1. Bản thể học thông tin (Brian Cantwell Smith):", "Thế giới khởi đầu từ dòng chảy siêu hình nguyên thủy; con người định hình thực tại bằng cách 'đăng ký đối tượng' (registration of objects) giữa các tác nhân ở xa nhau."),
                ("2. Thuyết thông tin khách quan (Fred Dretske):", "Thông tin tồn tại khách quan theo quy luật tự nhiên; bản rõ mang thông tin về nguồn phát, hệ mật hoàn hảo đảm bảo bản mã không để rò rỉ bất kỳ thông tin nào về nguồn."),
                ("3. Bản chất của Kẻ đối đầu:", "Đối thủ là tác nhân có chủ đích (intentional agent), có ý chí và mục tiêu, khác hoàn toàn với nhiễu ngẫu nhiên vật lý."),
                ("4. Bản chất của Bí mật & Tri thức:", "Common Knowledge (David Lewis): Mọi người đều biết và tin trạng thái sự việc. Public Knowledge: Thông tin ai cũng có tiềm năng tiếp cận. Bí mật là 'tri thức không phổ biến', không đồng nhất với bản rõ hay bản mã.")
            ],
            "explanation": (
                "Giải thích các thuật ngữ triết học khó hiểu trên Slide:\n"
                "• 'Đăng ký đối tượng' (Smith) là gì? Thế giới thực vốn là một mớ dữ liệu liên tục không có đường biên. Để hai máy tính ở xa nhau nói chuyện an toàn được, con người phải 'đóng khung' dữ liệu lại thành các đối tượng rõ ràng: đâu là Thông điệp M, đâu là Bản mã C, đâu là Chìa khóa K. Nhờ đóng khung rạch ròi, hệ thống mới biết phần nào cần bảo vệ và truyền đi.\n"
                "• 'Thông tin khách quan' (Dretske) là gì? Dretske cho rằng thông tin tồn tại khách quan tự nhiên, giống như tiếng sấm nổ tự nó báo hiệu trời sắp mưa, dù có ai nghe hay không. Áp dụng vào mật mã: Bản rõ chứa thông tin thật; nếu mã hóa tốt thì bản mã nhìn như một mớ dữ liệu rác vô nghĩa, hoàn toàn không dính dáng gì đến thông tin gốc, khiến kẻ trộm nhìn vào không suy đoán được gì.\n"
                "• Phân biệt Common Knowledge và Public Knowledge (David Lewis 1969):\n"
                "  - Common Knowledge (Tri thức chung): Là thông tin mà mọi người đều biết, và ai cũng biết là người khác cũng biết (ví dụ: luật đèn đỏ phải dừng, ai đi đường cũng biết và tin chắc người đối diện cũng biết luật đó).\n"
                "  - Public Knowledge (Tri thức công khai): Là thông tin ai cũng có thể tìm thấy nếu muốn (ví dụ: tài liệu mật bị rò rỉ lên mạng, ai tò mò lên mạng tải về là đọc được).\n"
                "  - Bí mật (Secret): Tác giả định nghĩa bí mật là 'thứ không công khai' (uncommon knowledge). Bí mật không phải tự thân chuỗi bit có phép thuật, mà là mối quan hệ xã hội: Chừng nào nhóm của ta vẫn giữ kín được thông tin và không để lọt ra ngoài, thì chừng đó nó vẫn là bí mật."
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
                ("Giai đoạn 1 (1948–1949) — Mô hình Shannon:", "Mô hình bảo mật điểm-điểm; định nghĩa Bí mật hoàn hảo H(K) >= H(M). Giới hạn: Giản lược đối thủ thành kẻ nghe lén trên kênh trừu tượng, bỏ qua chi phí phân phối khóa."),
                ("Giai đoạn 2 (1976) — Cách mạng Khóa công khai:", "Diffie-Hellman và Merkle giải quyết bài toán phân phối khóa, chuyển gánh nặng sang độ phức tạp tính toán một chiều."),
                ("Giai đoạn 3 (1982–1984) — PPT Adversary & Security Games:", "Goldwasser-Micali định nghĩa đối thủ thời gian đa thức xác suất (PPT) và an toàn ngữ nghĩa thông qua các trò chơi an ninh."),
                ("Giai đoạn 4 (1981–1993) — Dolev-Yao & Random Oracle:", "Dolev-Yao mô hình hóa đối thủ kiểm soát toàn bộ kênh truyền mạng; Bellare-Rogaway đề xuất Random Oracle Model (ROM) cho hàm băm."),
                ("Giai đoạn 5 (2013–2025) — Phê phán hậu Snowden:", "Snowden phơi bày năng lực thực tế của NSA: Kênh kề, cửa sau chuẩn mật mã, tấn công chuỗi cung ứng và giám sát siêu dữ liệu.")
            ],
            "explanation": (
                "Giải thích các khái niệm khó hiểu trên Slide:\n"
                "• Bí mật hoàn hảo của Shannon H(K) >= H(M) có hạn chế gì? Shannon chứng minh muốn bảo mật tuyệt đối thì chìa khóa phải dài tối thiểu bằng bức thư (như mã dùng một lần One-Time Pad). Nhưng hạn chế ngoài đời là: Nếu ta muốn gửi bức thư 10 GB qua mạng mà lại phải đem theo cái ổ cứng chứa chìa khóa 10 GB đi giao tận tay người nhận thì cực kỳ tốn kém và không thể áp dụng cho mạng Internet toàn cầu.\n"
                "• 'PPT Adversary & Security Games' (1984) là gì? Vì không thể dùng chìa khóa dài vô tận của Shannon, các nhà khoa học hiện đại chấp nhận thực tế: Kẻ trộm chỉ có máy tính chạy thời gian đa thức (PPT - tức máy tính có sức mạnh hữu hạn, không thể thử hàng tỷ tỷ năm). Họ đặt ra một 'Trò chơi an ninh': Đưa cho kẻ trộm hai gói tin mã hóa, nếu kẻ trộm không đoán được gói nào chứa tin nhắn nào (tỷ lệ đoán chỉ bằng 50% tung đồng xu) thì thuật toán đó được coi là an toàn."
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
                ("Panel Trái — Mô hình điểm-điểm Shannon (1949):", "Alice gửi bản rõ M mã hóa bằng khóa K qua kênh truyền đến Bob. Đối thủ Eve chỉ là kẻ nghe lén thụ động trên đường truyền. Bỏ qua chi phí quản lý khóa và bề mặt tấn công vật lý."),
                ("Panel Phải — Mạng xã hội - kỹ thuật phức hợp (Halpin 2025):", "Thông tin luân chuyển giữa Nhóm giữ bí mật và Hạ tầng mở / ISP / Đám mây. Đối thủ thể chế (NSA, Five Eyes) tấn công qua 4 con đường: (1) Kênh kề; (2) Giám sát siêu dữ liệu; (3) Cửa sau phần cứng; (4) Lỗi phản bội nội bộ Byzantine.")
            ],
            "explanation": (
                "Giải thích bản chất trực quan của Sơ đồ đối sánh:\n"
                "• Tại sao mô hình Shannon không còn đủ cho thế giới ngày nay? Thời của Shannon (1949), người ta chỉ nghĩ đơn giản là kéo một sợi dây điện từ nhà Alice sang nhà Bob, và có một kẻ xấu Eve đứng kẹp dây nghe trộm. Nhưng Internet ngày nay phức tạp hơn thế: Gói tin phải đi qua nhà mạng Viettel/VNPT, qua máy chủ đám mây của Google, Amazon. Kẻ nghe lén hiện nay là các cơ quan tình báo khổng lồ. Họ không chỉ nghe lén nội dung tin nhắn, mà họ soi 'Siêu dữ liệu' (Metadata): Tức là ai nhắn tin cho ai, gửi lúc mấy giờ, tệp nặng bao nhiêu MB. Chỉ cần nắm siêu dữ liệu là họ biết hết quan hệ và hành vi của người dùng mà không cần tốn công bẻ khóa mật mã."
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
                ("1. Mô hình PPT Adversary (Goldwasser-Micali 1984):", "Đối thủ là thuật toán Turing xác suất chạy thời gian đa thức; an toàn được định nghĩa qua Trò chơi An ninh với lợi thế đoán đúng là hàm không đáng kể; chứng minh bằng quy dẫn toán học."),
                ("2. Mô hình Đối thủ mạng Logic (Dolev-Yao 1983):", "Kẻ thù kiểm soát toàn bộ kênh truyền (đọc, chặn, sửa, xóa, phát lại gói tin), nhưng xem mật mã là 'hộp đen đại số hoàn hảo', không thể giải mã nếu không có khóa."),
                ("3. Mô hình Tiên tri Ngẫu nhiên (Bellare-Rogaway 1993):", "Giả định hàm băm thực tế hoạt động như một cỗ máy ngẫu nhiên lý tưởng (Random Oracle); thực tế không có hàm băm nào tương đương toán học với một Random Oracle thực sự.")
            ],
            "explanation": (
                "Giải thích cốt lõi của 3 mô hình toán học:\n"
                "• Mô hình Dolev-Yao (1983) coi kẻ tấn công như 'ông trùm đường truyền Internet': Muốn chặn gói tin, sửa tin hay phát lại gói tin cũ lúc nào cũng được. Nhưng mô hình này lại ngây thơ coi mật mã như một 'hộp đen thần kỳ', tức là kẻ địch không bao giờ phá được nếu không có khóa.\n"
                "• Mô hình Random Oracle (ROM 1993) coi hàm băm (như SHA-256) hoàn hảo như một cỗ máy bốc thăm ngẫu nhiên tuyệt đối của trời cho. Nhưng thực tế ngoài đời, SHA-256 chỉ là một đoạn code toán học cố định do con người viết ra, nên không thể ngẫu nhiên tuyệt đối 100% được.\n"
                "• Tác giả phê phán điều gì? Cả 3 mô hình này đều quá lý tưởng trên giấy tờ. Kẻ thù ngoài đời không ngoan ngoãn giải toán như các nhà khoa học giả định, mà chúng tấn công vào điểm mù phần cứng và con người."
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
                ("Bước 1 — Giả định toán học hình thức:", "Hộp đen mật mã lý tưởng, thời gian thực thi vô hình, độ dài bản mã không lộ nội dung, kênh truyền độc lập."),
                ("Bước 2 — Yếu tố vật lý bị trừu tượng hóa bỏ qua:", "Thời gian vi xử lý CPU, điện năng tiêu thụ, bức xạ điện từ, bộ nhớ đệm Cache, lỗi lập trình phần mềm."),
                ("Bước 3 — Con đường khai thác của đối thủ thế giới thực:", "Tấn công kênh kề thời gian (Paul Kocher 1996), phân tích lưu lượng và độ dài siêu dữ liệu, khai thác lỗi cài đặt phần mềm và quản trị (Ross Anderson 2001)."),
                ("Bước 4 — Kết luận nhận thức luận của Halpin:", "Sự trừu tượng hóa quá mức làm cho mật mã học mù quáng trước chính những năng lực thực tế của đối thủ nằm ngoài mô hình lý thuyết.")
            ],
            "explanation": (
                "Giải thích trực quan về Tấn công Kênh kề (Side-channel Attack):\n"
                "• Tấn công kênh kề thời gian (Kocher 1996) hoạt động thế nào? Trong lập trình, khi so sánh hai chuỗi mật khẩu, máy tính thường dùng lệnh so sánh tuần tự từ trái qua phải: hễ thấy ký tự đầu tiên bị sai là nó dừng lại ngay lập tức (ngắt sớm - early return). Kẻ tấn công lợi dụng điều này bằng cách gửi thử các ký tự và dùng đồng hồ siêu chính xác đo thời gian máy tính trả lời: Nếu thấy máy xử lý lâu hơn một chút xíu (vài phần triệu giây), nghĩa là chữ cái đầu tiên đã đoán đúng! Bằng cách đó, kẻ trộm dò ra từng chữ cái của chìa khóa mà không cần phải bẻ khóa thuật toán.\n"
                "• Rò rỉ qua độ dài siêu dữ liệu là gì? Dù tin nhắn được mã hóa kín 100%, nhưng nếu ta gửi một tệp nặng đúng 4.5 MB vào lúc 2 giờ sáng, kẻ theo dõi mạng đối chiếu dung lượng là biết ngay ta vừa gửi bài hát hay đoạn video cụ thể nào."
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
                ("1. Khủng hoảng thực tiễn (Snowden 2013):", "NSA không cần giải các bài toán toán học khó của RSA/AES mà tấn công vào toàn bộ hệ sinh thái xung quanh (đánh cắp khóa, thâm nhập cáp quang biển)."),
                ("2. Năng lực đối thủ thể chế:", "Khai thác Siêu dữ liệu để dựng mạng xã hội; Tấn công chuỗi cung ứng bằng cách cài cửa sau vào chuẩn NIST (Dual_EC_DRBG); Ép buộc pháp lý các tập đoàn công nghệ lớn."),
                ("3. Tính tương đối của Kẻ thù:", "'Kẻ thù' mang tính vị thế quyền lực; Mật mã học buộc phải nghiên cứu phòng chống cửa sau (Cliptography) thay vì giả định kênh sạch.")
            ],
            "explanation": (
                "Giải thích vụ việc cài cửa sau Dual_EC_DRBG nổi tiếng của NSA:\n"
                "• Chuyện gì đã xảy ra trong vụ Dual_EC_DRBG? Bộ sinh số ngẫu nhiên là trái tim của mật mã: Chìa khóa bí mật bắt buộc phải được tạo ra ngẫu nhiên thì người khác mới không đoán được. Thay vì đi phá từng chiếc chìa khóa, cơ quan an ninh NSA đã bí mật tác động vào Viện Tiêu chuẩn Mỹ (NIST) để biến một thuật toán tạo số ngẫu nhiên bị lỗi (Dual_EC_DRBG) thành tiêu chuẩn quốc tế. NSA nắm giữ một công thức toán học bí mật đằng sau thuật toán này, do đó bất kỳ công ty hay tổ chức nào dùng chuẩn này để tạo khóa thì NSA đều tính toán ra được chìa khóa đó trong chớp mắt.\n"
                "• 'Cliptography' là gì? Là ngành nghiên cứu cách tạo ra hệ thống bảo mật an toàn ngay cả khi chính con chip hoặc phần mềm mã hóa đã bị kẻ xấu cài cắm cửa sau từ trước."
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
                ("Câu hỏi 1 (RQ1) — Mạng xã hội động:", "Nghiên cứu hành vi thực tế của con người khi giữ bí mật; kiểm tra xem chính sách an ninh có bền vững trước sự cố Byzantine nội bộ."),
                ("Câu hỏi 2 (RQ2) — Mô hình hóa bằng thống kê:", "Dùng thống kê thực nghiệm và mô phỏng Monte Carlo để lượng hóa năng lực thực tế của đối thủ và liên minh giám sát."),
                ("Câu hỏi 3 (RQ3) — Phương pháp liên ngành:", "Ứng dụng lý thuyết hệ thống phức hợp và động lực học mạng (như mô hình lây lan dịch bệnh) vào nghiên cứu sự phát tán thông tin."),
                ("Câu hỏi 4 (RQ4) — Triết học chất vấn ngầm định:", "Vận dụng triết học để xem xét lại các giả định nền tảng 70 năm qua từ thời Shannon, thích ứng với kỷ nguyên số và AI.")
            ],
            "explanation": (
                "Ý nghĩa tóm gọn của 4 câu hỏi nghiên cứu:\n"
                "Halpin muốn nhắn nhủ thông điệp: Các nhà khoa học không thể mãi ngồi trong phòng thí nghiệm giải toán một mình được nữa. Mật mã học tương lai phải: (1) Quan sát xem con người thực tế giữ bí mật như thế nào; (2) Thu thập số liệu thống kê để đo lường sức mạnh thực sự của các cơ quan tình báo và tin tặc; (3) Mượn mô hình lây lan dịch bệnh để nghiên cứu xem thông tin mật bị rò rỉ ra sao; và (4) Dùng triết học để xem lại xem các giả định từ 70 năm trước có còn phù hợp với thế giới Internet và AI ngày nay hay không."
            ),
            "footer": "Trích dẫn nguyên văn 4 câu hỏi nghiên cứu mở: Harry Halpin (2025), Section 7, tr. 12–14."
        },

        # SLIDE 10
        {
            "num": 10,
            "title": "PHƯƠNG PHÁP LUẬN & KHẢ NĂNG TÁI LẬP BÀI BÁO",
            "time": "1:00",
            "header": "ĐÁNH GIÁ PHƯƠNG PHÁP LUẬN & HƯỚNG TIẾP CẬN THỰC NGHIỆM",
            "lead": "Làm rõ tính khả thi tái lập đối với một công trình triết học và định hướng thí nghiệm độc lập",
            "bullets": [
                ("1. Đánh giá tính khả thi tái lập:", "Bài báo của Halpin là công trình khảo luận triết học thuần túy, hoàn toàn KHÔNG có mã nguồn, không có tập dữ liệu và không có số liệu đo đạc benchmark ban đầu. Nhóm thực hiện thẩm định học thuật (Scholarly Audit) đối chiếu toàn bộ 17 trang bản in PDF và 134 tài liệu tham khảo."),
                ("2. Định hướng thực nghiệm độc lập của sinh viên:", "Sinh viên tự xây dựng hệ thống kiểm tra toàn vẹn tệp để minh họa trực quan các khái niệm lý thuyết; thiết kế các kịch bản đối kháng kiểm chứng CRC32, Dolev-Yao, kênh kề thời gian. Quy trình kiểm toán lưu vết tại MIDTERM_CODE_AUDIT.md.")
            ],
            "explanation": (
                "Giải thích trả lời câu hỏi phản biện then chốt:\n"
                "• Tại sao bài báo không có code mà đề tài lại viết ra phần mềm demo? Vì bài báo của TS. Harry Halpin là một bài nghiên cứu triết học lý thuyết (toàn chữ và lập luận, không có một dòng code hay bảng số liệu thực nghiệm nào). Nhóm sinh viên không thể 'chạy lại code của tác giả' vì tác giả không công bố code. Do đó, nhóm đã tự tay thiết kế và lập trình một hệ thống kiểm tra toàn vẹn tệp độc lập, sử dụng tập dữ liệu thực tế 2.485 tệp tin để làm bằng chứng thực nghiệm trực quan minh họa cho các luận điểm của bài báo."
            ),
            "footer": "Báo cáo phân định phương pháp luận độc lập của nhóm nghiên cứu đề tài DT005."
        },

        # SLIDE 11
        {
            "num": 11,
            "title": "MINH HỌA PAPER-TO-CODE ĐỘC LẬP & GIỚI HẠN DIỄN GIẢI",
            "time": "1:15",
            "header": "SƠ ĐỒ 4: PHÂN ĐỊNH LÝ THUYẾT & THÍ NGHIỆM ĐỘC LẬP",
            "lead": "Đối sánh minh bạch giữa các luận điểm bài báo và thí nghiệm minh họa do sinh viên xây dựng",
            "bullets": [
                ("Ví dụ 1 — Nhiễu ngẫu nhiên vs Đối thủ có chủ đích:", "Halpin chỉ rõ mã kiểm tra lỗi chỉ chống được nhiễu ngẫu nhiên. DT005 thử nghiệm kịch bản giải hệ phương trình tuyến tính trên GF(2) bù 4 byte qua mặt CRC32; SHA-256 phát hiện tức thì nhờ tính phi tuyến. Giới hạn: Không khẳng định CRC32 bất lực trong kiểm tra lỗi ngẫu nhiên vật lý."),
                ("Ví dụ 2 — Phân tách miền nút lá (RFC 6962):", "Halpin nhấn mạnh điểm mù cấu trúc dữ liệu nhưng không đề cập cây Merkle. DT005 chủ động cài đặt phân tách miền theo chuẩn IETF RFC 6962 (tiền tố byte 0x00 cho lá, 0x01 cho nhánh) để phòng ngừa va chạm cấu trúc. Giới hạn: Đây là đóng góp độc lập của sinh viên, không phải đề xuất của Halpin."),
                ("Ví dụ 3 — Rò rỉ kênh kề thời gian:", "Halpin dẫn chứng thời gian thực thi làm rò rỉ bí mật (Kocher 1996). DT005 đo đạc so sánh chuỗi ngắt sớm so với hàm thời gian hằng số hmac.compare_digest. Giới hạn: Đo trên Python có jitter hệ điều hành, chỉ minh họa mức phần mềm.")
            ],
            "explanation": (
                "Giải thích sâu 2 kỹ thuật then chốt trên Slide 11:\n"
                "• Tại sao mã CRC32 bị bẻ gãy dễ dàng bằng phép toán GF(2), còn SHA-256 thì không? CRC32 (mã kiểm tra lỗi dạng vòng) vốn chỉ được tạo ra để phát hiện lỗi ngẫu nhiên trên dây cáp hay đĩa xước. Công thức của CRC32 là phép chia đa thức tuyến tính rất đơn giản. Kẻ tấn công chỉ cần dùng toán đại số lớp 10 (giải hệ phương trình bậc nhất trên trường nhị phân GF(2)) là tính ra được đúng 4 byte vá víu dán vào đuôi tệp tin, khiến nội dung tệp bị sửa đổi hoàn toàn nhưng mã CRC32 tính ra vẫn trùng khít 100% với tệp gốc! Trong khi đó, SHA-256 là hàm băm mật mã học phi tuyến, có hiệu ứng tuyết lở (~50%): Chỉ cần sửa 1 bit dữ liệu là hơn nửa số bit của mã băm bị đảo lộn ngẫu nhiên, kẻ xấu không thể nào giải phương trình để làm giả được.\n"
                "• Lỗ hổng nhập nhằng nút lá (Node-leaf ambiguity) và chuẩn RFC 6962 là gì? Trong cây Merkle thông thường, các nút lá và nút nhánh đều được băm bằng công thức giống nhau. Kẻ xấu có thể lấy mã băm của một nút nhánh bên trên rồi giả mạo nó thành nội dung của một tệp tin lá ở dưới, làm hệ thống kiểm tra bị nhầm lẫn cấu trúc. Để khắc phục, nhóm sinh viên áp dụng chuẩn quốc tế IETF RFC 6962: Tệp tin lá thì bắt buộc thêm byte 0x00 vào đầu trước khi băm, còn nút nhánh thì thêm byte 0x01. Giống như dán nhãn phân loại hàng hóa, hệ thống sẽ không bao giờ bị lẫn lộn giữa lá và cành nữa."
            ),
            "footer": "Bảng đối chiếu minh bạch do sinh viên tự xây dựng; chi tiết mã nguồn lưu tại MIDTERM_CODE_AUDIT.md."
        },

        # SLIDE 12
        {
            "num": 12,
            "title": "SLIDE CHUYỂN TIẾP: DEMO",
            "time": "0:15",
            "header": "GIAO DIỆN TỐI GIẢN NỀN SÁNG CHUẨN BỊ DIỄN TẬP TRỰC TIẾP",
            "lead": "Chữ hiển thị trên Slide: DEMO",
            "bullets": [
                ("Văn bản trung tâm:", "DEMO (Được thiết kế tối giản, chữ to rõ ràng trên nền sáng thanh lịch để tạo nhịp nghỉ trang trọng chuyển tiếp phần thuyết trình sang phần thực nghiệm).")
            ],
            "explanation": (
                "Lời thoại chuyển tiếp vào phần Demo:\n"
                "Sau khi đã làm rõ các cơ sở triết học và phân định phương pháp luận trên lý thuyết, nhóm xin phép chuyển sang phần Diễn tập An ninh Thực nghiệm trực tiếp (Live Demo) thông qua mô hình đối kháng hai màn hình Red Team và Blue Team, tập trung kiểm chứng thực tế các luận điểm mấu chốt của bài báo."
            ),
            "footer": "Slide chuyển tiếp phần diễn tập thực hành đối kháng an ninh."
        },

        # SLIDE 13
        {
            "num": 13,
            "title": "KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG: RED TEAM VS BLUE TEAM",
            "time": "1:00",
            "header": "SƠ ĐỒ 5: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG (DEMO ARCHITECTURE)",
            "lead": "Mô hình đối kháng thực tế 2 màn hình trên tập dữ liệu kaggle_dataset (2.485 tệp tin / ~112,81 MiB theo benchmark)",
            "bullets": [
                ("1. Hạ tầng diễn tập 2 Màn hình:", "Tương tác thời gian thực giữa Hệ thống giám sát phòng thủ Blue Team (Port 8080) và Động cơ tấn công đối kháng Red Team (Port 8081)."),
                ("2. Tập dữ liệu thực nghiệm:", "kaggle_dataset gồm 2.485 tệp tin, dung lượng ~112,81 MiB (phiên bản nạp kiểm thử cục bộ hiện tại gồm 2.483 tệp) đóng vai trò môi trường kiểm thử thực tế nhiều tầng thư mục."),
                ("3. Blue Team (Port 8080):", "Nạp dữ liệu, thiết lập mốc tin cậy sạch ban đầu H_0; quét kiểm tra toàn vẹn định vị chính xác vị trí tệp tin bị biến đổi với tốc độ quét thực tế vượt 2.000 tệp/giây."),
                ("4. Red Team (Port 8081):", "Mô phỏng các vector tấn công có chủ đích: Sửa byte tệp ngẫu nhiên (In-place Tamper) và giải hệ phương trình tuyến tính trên GF(2) bù 4 byte qua mặt mã kiểm tra CRC32."),
                ("5. Cơ chế hoàn nguyên an toàn:", "Thư mục sao lưu an toàn .adversary_backups/ cho phép phục hồi nguyên vẹn 100% dữ liệu chỉ với 1 click chuột Rollback.")
            ],
            "explanation": (
                "Giải thích cơ chế vận hành của buổi Demo trực tiếp:\n"
                "Mô hình được thiết kế như một cuộc đấu an ninh trực tiếp giữa hai màn hình:\n"
                "• Màn hình bên phải là Red Team (Kẻ tấn công có chủ đích): Đóng vai tin tặc, bấm nút sửa tệp Resume/Resume.csv hoặc tính toán 4 byte bù để qua mặt CRC32.\n"
                "• Màn hình bên trái là Blue Team (Hệ thống giám sát phòng thủ): Nạp tập dữ liệu thực tế 2.485 tệp tin. Khi bấm kiểm tra, thuật toán duyệt cây phân cấp sẽ rà soát và định vị chính xác vị trí tệp tin bị sửa đổi chỉ trong tích tắc (tốc độ quét trên 2.000 tệp/giây), phát còi báo động đỏ.\n"
                "• Sau khi diễn tập xong, nút Rollback All sẽ tự động phục hồi toàn bộ dữ liệu gốc từ thư mục sao lưu an toàn .adversary_backups/."
            ),
            "footer": "Sơ đồ kiến trúc diễn tập đối kháng độc lập do sinh viên thiết kế minh họa các khái niệm an ninh."
        },

        # SLIDE 14
        {
            "num": 14,
            "title": "TỔNG KẾT BÁO CÁO GIỮA KỲ & ĐỊNH HƯỚNG NGHIÊN CỨU",
            "time": "0:45",
            "header": "SECTION 7 — KẾT LUẬN & THÔNG ĐIỆP HỌC THUẬT",
            "lead": "Tổng kết các đóng góp lý thuyết của Halpin và định hướng nghiên cứu tiếp theo",
            "bullets": [
                ("1. Thông điệp học thuật cốt lõi (Harry Halpin 2025):", "Vượt qua sự trừu tượng hóa: Các hệ thống an toàn thế giới thực phải nhận thức khoảng cách giữa mô hình lý tưởng và thực tế vật lý (kênh kề, siêu dữ liệu, chuỗi cung ứng). Bản chất đối thủ là tác nhân có chủ đích. Halpin đặt nền móng cho khoa học an ninh liên ngành."),
                ("2. Hoàn thành giữa kỳ & Lộ trình cuối kỳ (DT005):", "Nghiệm thu giữa kỳ: Hoàn thành thẩm định học thuật sâu sắc toàn văn bài báo gốc (17 trang PDF), phân tích đầy đủ 7 mục nội dung. Minh bạch phương pháp luận giữa lý thuyết và phần mềm thực nghiệm. Định hướng cuối kỳ: Tiếp tục nghiên cứu hoàn thiện cấu trúc cây phân cấp tối ưu cho kiểm tra toàn vẹn dữ liệu lớn.")
            ],
            "explanation": (
                "Tóm tắt kết luận ngắn gọn:\n"
                "Giữa kỳ nhóm đã hoàn thành tốt 2 nhiệm vụ: (1) Thẩm định bài báo một cách có hệ thống, hiểu rõ bản chất triết học của đối thủ; và (2) Dựng xong bộ công cụ demo trực quan minh chứng cho lý thuyết. Định hướng cuối kỳ sẽ tiếp tục nghiên cứu sâu hơn về thuật toán tối ưu hóa cấu trúc cây Merkle để xử lý các tập dữ liệu khổng lồ."
            ),
            "footer": "Tài liệu tham khảo chính: Halpin (2025), Shannon (1949), Goldwasser-Micali (1984), Dolev-Yao (1983), Kocher (1996), RFC 6962."
        },

        # SLIDE 15
        {
            "num": 15,
            "title": "CẢM ƠN & PHIÊN PHẢN BIỆN (Q&A)",
            "time": "0:15",
            "header": "Trường ĐH Tài nguyên và Môi trường TP.HCM — Khoa Công nghệ thông tin",
            "lead": "CẢM ƠN THẦY VÀ CÁC BẠN ĐÃ LẮNG NGHE!",
            "bullets": [
                ("Thông điệp kết thúc:", "Q & A — Xin đón nhận các câu hỏi và ý kiến đóng góp của Thầy và Hội đồng."),
                ("Thông tin đề tài:", "Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143.")
            ],
            "explanation": (
                "Lời kết thúc báo cáo:\n"
                "Phần trình bày báo cáo giữa kỳ của đề tài DT005 đến đây là kết thúc. Xin chân thành cảm ơn sự chú ý lắng nghe của Thầy cùng các bạn. Nhóm xin sẵn sàng lắng nghe các ý kiến nhận xét và giải đáp các câu hỏi trong phiên thảo luận."
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

        # Hộp giải thích dễ hiểu (chỉ dành cho các slide có khái niệm cần giải thích)
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
        "Mục tiêu diễn tập giữa kỳ: Chỉ tập trung trình diễn các kịch bản đối kháng cốt lõi phản ánh trực tiếp luận điểm bài báo Harry Halpin (2025). "
        "Các chức năng nghiên cứu mở rộng phục vụ đồ án cuối kỳ (như đánh tráo Merkle Proof phân tán, đo đạc độ sâu động, nén CompactProof) được lược bỏ để đảm bảo thời lượng và làm nổi bật trọng tâm khoa học."
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
         "Mở terminal và kích hoạt song song 2 dịch vụ:\n"
         "• Blue Team Service:  python -m src.demo.web_demo --port 8080\n"
         "• Red Team Service:   python -m src.demo.red_team_web --port 8081\n"
         "Mở 2 cửa sổ trình duyệt chia đôi màn hình (Split Screen): Cửa sổ bên trái truy cập http://localhost:8080 (Blue Team); cửa sổ bên phải truy cập http://localhost:8081 (Red Team)."),

        ("Bước 2: Thiết lập mốc tin cậy sạch ban đầu tại Blue Team (Port 8080)",
         "Tại cổng 8080, chọn thư mục dữ liệu thực tế 'kaggle_dataset' (~112.8 MB).\n"
         "Bấm nút 'Nạp & Tính toán Toàn vẹn'. Hệ thống tính toán và hiển thị Root Hash ban đầu H_0 với trạng thái màu xanh an toàn: '100% NGUYÊN VẸN'.\n"
         "Ý nghĩa học thuật: Minh họa khái niệm 'đăng ký đối tượng' (registration of objects) theo Cantwell Smith và thiết lập trạng thái tin cậy khách quan theo Dretske."),

        ("Bước 3: Phát động Kịch bản 1 — Can thiệp tệp tin có chủ đích tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm nút 'Can thiệp tệp tin ngẫu nhiên' (In-place Tampering). Red Team sửa đổi trực tiếp dữ liệu tệp tin trên ổ đĩa (ví dụ: Resume/Resume.csv).\n"
         "Quay lại Blue Team (Port 8080), bấm nút 'Kiểm tra toàn vẹn':\n"
         "• Hệ thống ngay lập tức đổi sang trạng thái cảnh báo đỏ: 'CẢNH BÁO: PHÁT HIỆN CAN THIỆP DỮ LIỆU!'.\n"
         "• Bảng truy vết chỉ ra chính xác tệp tin Resume/Resume.csv bị tác động.\n"
         "Ý nghĩa học thuật: Minh chứng luận điểm Halpin — Kẻ đối đầu mật mã học là tác nhân có chủ đích (intentional agent), hàm băm mật mã học lập tức bắt quả tang mọi sai lệch dù chỉ 1 bit."),

        ("Bước 4: Phát động Kịch bản 2 — Tấn công bù 4 byte GF(2) qua mặt CRC32 tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm nút 'Tấn công bù 4 byte GF(2)' (CRC32 Forge Attack). Động cơ Red Team giải hệ phương trình tuyến tính, làm sai lệch nội dung tệp nhưng bù 4 byte patch khiến cho mã CRC32 không đổi.\n"
         "Quay lại Blue Team (Port 8080), bấm nút 'Kiểm tra toàn vẹn':\n"
         "• Bảng đối chứng hiển thị rõ: Thuật toán kiểm tra lỗi CRC32 bị qua mặt hoàn toàn (báo khớp sai sự thật do tính chất tuyến tính của đa thức chia).\n"
         "• Ngược lại, hàm băm SHA-256 phát hiện vi phạm ngay lập tức nhờ tính phi tuyến và hiệu ứng tuyết lở (~50%).\n"
         "Ý nghĩa học thuật: Trực quan hóa luận điểm mấu chốt của bài báo Halpin — phân biệt rạch ròi giữa mã phát hiện lỗi ngẫu nhiên của Shannon và mật mã học trước kẻ thù có chủ đích."),

        ("Bước 5: Hoàn nguyên dữ liệu sạch 100% (Rollback All)",
         "Tại giao diện Red Team (Port 8081), bấm nút 'ROLLBACK ALL'. Toàn bộ các tệp tin bị sửa đổi được khôi phục nguyên vẹn từ thư mục sao lưu an toàn .adversary_backups/.\n"
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
        "   Kết quả: Đo đạc thực tế tính phi tuyến của SHA-256, hiệu ứng tuyết lở SAC ~50%, kịch bản bẻ gãy mã tuyến tính CRC32 trên GF(2), và kênh kề thời gian so sánh chuỗi.\n\n"
        "2. Diễn tập 2 Terminal Red Team vs Blue Team:\n"
        "   • Terminal 1 (Red Team): Chạy 'python src/adversary/red_team_cli.py'. Bấm phím '1' để can thiệp tệp, hoặc phím '4' để bù mã CRC32.\n"
        "   • Terminal 2 (Blue Team): Chạy 'python -m src.baselines.sha256_manifest --dir kaggle_dataset --verify'. Blue Team phát hiện vi phạm ngay lập tức: MODIFIED_FILE: Resume/Resume.csv.\n"
        "   • Terminal 1: Bấm phím '9' để Rollback All khôi phục đĩa cứng an toàn."
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
        ("Câu hỏi 1: Bài báo của Halpin là một bài khảo luận triết học, tại sao đề tài lại chọn bài toán Kiểm tra toàn vẹn tệp tin?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Thuộc tính an ninh cốt lõi: Trong cấu trúc tam giác an ninh thông tin (Confidentiality, Integrity, Availability), Halpin phân tích sâu sắc rằng Tính toàn vẹn (Integrity) là điều kiện tiên quyết mang tính bản thể luận. Một thông điệp có thể không bí mật nhưng bắt buộc phải toàn vẹn thì tác nhân mới có thể hành động dựa trên thông tin đó.\n"
         "2. Lời kêu gọi về một 'Khoa học an ninh thực nghiệm': Halpin phê phán việc toán học hóa trừu tượng đã làm mật mã học xa rời thực tế vật lý. Tác giả kêu gọi xây dựng một nền an ninh thực nghiệm. Bài toán kiểm tra toàn vẹn tệp tin trên hệ thống đĩa cứng thực là bài toán kinh điển, cho phép đưa các khái niệm trừu tượng vào kiểm định vật lý.\n"
         "3. Minh họa các đối thủ trong bài báo: Hệ thống toàn vẹn tệp tin cho phép mô hình hóa và kiểm chứng trực quan các dạng đối thủ được Halpin mổ xẻ: Từ kẻ nghe lén của Shannon, kẻ can thiệp mạng Dolev-Yao, kẻ khai thác kênh kề Kocher, đến các đòn tấn công can thiệp dữ liệu và chèn backdoor chuỗi cung ứng."),

        ("Câu hỏi 2: Bài báo có 134 tài liệu tham khảo, khi thẩm định và tái lập thì có bắt buộc phải làm theo toàn bộ 134 tài liệu đó không?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Bản chất bài khảo cứu: Bài báo của Halpin là một bài tổng quan triết học (philosophical survey) dài 17 trang, 134 trích dẫn là kho tàng tài liệu lịch sử trải dài từ năm 1883 đến 2024, bao gồm cả sách triết học siêu hình (Dretske, Cantwell Smith), tài liệu lưu trữ chính phủ về NSA, và các hệ thống chứng minh định lý tự động. Về mặt phương pháp luận nghiên cứu, không một công trình kỹ thuật nào có thể hay bắt buộc phải tái lập toàn bộ 134 tài liệu nền đó.\n"
         "2. Phương pháp tinh lọc trọng tâm: Thẩm định một bài báo lý thuyết là tinh lọc các trụ cột kỹ thuật trực tiếp liên quan đến bài toán an ninh. Đề tài DT005 đã tinh lọc và kết nối với các tài liệu kỹ thuật cốt lõi: Ref [2, 3] của Shannon (Mã tuyến tính vs Hàm băm), Ref [21] của Biba (Toàn vẹn phân cấp), Ref [33, 34] của Goldwasser-Micali (Trò chơi an ninh PPT), Ref [36] của Kocher (Kênh kề thời gian), Ref [40] của Bellare-Rogaway (Random Oracle), Ref [52] của Dolev-Yao (Mạng thù địch), Ref [100, 101] về Snowden/NSA (Đối thủ thể chế), và bổ sung tài liệu kỹ thuật chuẩn IETF RFC 6962 về phân tách miền cho Cây Merkle."),

        ("Câu hỏi 3: Tập dữ liệu kaggle_dataset từ đâu ra? Có xuất hiện trong bài báo gốc không?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Khẳng định minh bạch khoa học: Bài báo gốc của Harry Halpin hoàn toàn KHÔNG CÓ bất kỳ tập dữ liệu, bảng số liệu thực nghiệm hay mã nguồn nào. Bài báo là công trình khảo luận nhận thức luận trên tạp chí lý thuyết.\n"
         "2. Mục đích đưa tập dữ liệu vào nghiên cứu: Tập dữ liệu Kaggle Resume Dataset (báo cáo benchmark ghi nhận 2.485 tệp tin, ~112.81 MiB) do nhóm sinh viên tự chọn và tiền xử lý. Lý do chọn tập dữ liệu này là để cung cấp một môi trường cây thư mục doanh nghiệp nhiều tầng thực tế, kích thước tệp biến thiên, phản ánh trung thực tải làm việc thực tế nhằm đo đạc năng lực định vị sai lệch và khả năng chịu tải đối kháng, hoàn toàn không phải ngụy tạo số liệu của tác giả."),

        ("Câu hỏi 4: Tại sao mã kiểm tra lỗi CRC32 bị qua mặt trong kịch bản giải hệ trên GF(2), trong khi SHA-256 lại phát hiện được gian lận?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Tính chất tuyến tính của CRC32: CRC32 hoạt động dựa trên phép chia đa thức trên trường nhị phân hữu hạn GF(2). Do tính chất tuyến tính, giá trị CRC của một thông điệp sửa đổi CRC(M XOR Delta) có thể tính toán từ CRC(M) và CRC(Delta). Kẻ đối đầu có chủ đích có thể giải hệ phương trình đại số tuyến tính bằng thuật toán khử Gauss để tìm ra 4 byte bù (patch) nối vào cuối tệp tin, khiến cho CRC32 của tệp giả mạo trùng khớp với tệp gốc ban đầu trong kịch bản kiểm thử này.\n"
         "2. Tính phi tuyến của SHA-256: SHA-256 là hàm băm mật mã học xây dựng trên cấu trúc lặp Merkle-Damgard và phép nén Davies-Meyer với các hàm phi tuyến (Ch, Maj, Sigma). Hàm băm đạt hiệu ứng tuyết lở (Strict Avalanche Criterion - SAC ~50%): Chỉ cần thay đổi 1 bit đầu vào thì trung bình 50% số bit đầu ra bị đảo ngẫu nhiên. Bài toán tìm tiền ảnh thứ hai trên SHA-256 có độ phức tạp 2^256, vượt quá năng lực tính toán của đối thủ đa thức xác suất (PPT Adversary)."),

        ("Câu hỏi 5: Tấn công Kênh kề thời gian (Timing Side-Channel) trong so sánh chuỗi hoạt động như thế nào và cách giải quyết?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Cơ chế rò rỉ: Hàm so sánh chuỗi thông thường (như toán tử == trong nhiều ngôn ngữ) thực hiện so sánh tuần tự từng ký tự từ trái sang phải và ngắt sớm (early return) khi phát hiện ký tự sai khác đầu tiên. Kẻ tấn công gửi các chuỗi thử nghiệm và đo thời gian phản hồi: Thời gian thực thi càng dài đồng nghĩa số lượng ký tự tiền tố đoán đúng càng nhiều. Từ đó, đối thủ khôi phục được mã băm hoặc khóa bí mật mà không cần giải mã.\n"
         "2. Giải pháp kỹ thuật: Chuyển sang sử dụng hàm so sánh thời gian hằng số (Constant-time comparison) `hmac.compare_digest`. Hàm này luôn duyệt qua toàn bộ độ dài của chuỗi bằng phép toán bitwise XOR tích lũy, đảm bảo thời gian thực thi độc lập với vị trí xuất hiện ký tự sai khác, triệt tiêu kênh rò rỉ thời gian ở mức phần mềm."),

        ("Câu hỏi 6: Giải pháp Cây Merkle phân cấp và phân tách miền RFC 6962 đóng vai trò gì trong đề tài?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Minh họa độc lập của nhóm: Bài báo của Halpin không đề xuất Cây Merkle hay chuẩn RFC 6962. Đây là phần nghiên cứu và triển khai độc lập của nhóm DT005 nhằm giải quyết bài toán toàn vẹn dữ liệu tệp tin.\n"
         "2. Phân tách miền RFC 6962: Phép ghép băm thông thường H(L || R) có nguy cơ dẫn đến lỗ hổng nhập nhằng nút lá (Node-leaf ambiguity) khi một nút lá có cấu trúc 64 byte bị kẻ tấn công đánh tráo thành cặp băm của nút nhánh nội bộ. Nhóm áp dụng chuẩn RFC 6962: Tiền tố 0x00 cho nút lá và 0x01 cho nút nhánh nội bộ, triệt tiêu hoàn toàn nhập nhằng cấu trúc ở cấp độ cây.\n"
         "3. Định hướng hoàn thiện cuối kỳ: Nhóm tiếp tục tối ưu hóa cơ chế bộ đệm cập nhật cục bộ và đo đạc thực nghiệm độ bền (Ablation Study) trên các quy mô dữ liệu lớn hơn.")
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
