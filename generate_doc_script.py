# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word (.docx) Kịch Bản Thuyết Trình & Hướng Dẫn Diễn Tập Demo:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn cập nhật:
- Ghi lại trung thực, đầy đủ toàn bộ nội dung xuất hiện TRÊN TỪNG SLIDE (15 slide hoàn chỉnh).
- Phần mở rộng thuyết minh được ghi TIẾP NỐI NGAY SAU câu nội dung của slide, đảm bảo ngắn gọn, súc tích, không làm loãng bài.
- Giữ nguyên cấu trúc học thuật, khách quan, không dùng xưng hô rườm rà ("Thưa Thầy", "Trả lời Thầy"...).
- Bao gồm đầy đủ 3 phần: (1) Nội dung chi tiết 15 Slide; (2) Kịch bản Demo tinh giản theo bài báo; (3) Bộ câu hỏi phản biện chuyên sâu (Q&A).
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
    r = h1.add_run("PHẦN 1: NỘI DUNG TRÊN SLIDE & LỜI THOẠI THUYẾT MINH CHI TIẾT (15 SLIDE)")
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
            "sections": [
                ("Thông tin báo cáo trên Slide:", [
                    ("Đề tài:", "Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động."),
                    ("Bài báo khảo sát:", "The adversary: the philosophy of cryptography — TS. Harry Halpin (Journal of Cybersecurity, Oxford Academic, 2025)."),
                    ("Thông tin quản lý:", "Học phần: An toàn và bảo mật hệ thống thông tin | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143 — Lớp: 11_CNPM2.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Báo cáo giữa kỳ đề tài DT005 tập trung vào 3 trọng tâm: (1) Thẩm định nhận thức luận về Kẻ thù trong mật mã học của Halpin; (2) Phân định phương pháp luận minh bạch giữa khảo cứu triết học lý thuyết của tác giả và hệ thống phần mềm minh họa độc lập do sinh viên xây dựng; (3) Diễn tập thực nghiệm đối kháng an ninh Red Team vs Blue Team trên tập dữ liệu thực tế.",
            "footer": "Bản trình chiếu chính thức phục vụ nghiệm thu đồ án giữa kỳ."
        },

        # SLIDE 2
        {
            "num": 2,
            "title": "BỐI CẢNH, MỤC TIÊU & BA ĐÓNG GÓP CỦA BÀI BÁO",
            "time": "1:00",
            "header": "SECTION 1 — GIỚI THIỆU & MỤC TIÊU BÀI BÁO (TRANG 1–2)",
            "lead": "Khảo cứu nhận thức luận về Kẻ thù trong Mật mã học và Chương trình nghiên cứu tương lai",
            "sections": [
                ("1. Bối Cảnh Mật Mã Học & Nghịch Lý Nhận Thức Luận:", [
                    ("Mật mã học trong kỷ nguyên số:", "Khoa học về truyền thông bí mật khi có sự hiện diện của đối thủ có chủ đích. Internet biến không gian mạng thành nơi tương tác với vô vàn đối thủ tiềm tàng, xa lạ — qua đó cho thấy giả định về kênh truyền đóng kín không còn phù hợp."),
                    ("Nghịch lý nhận thức luận:", "Lý thuyết thông tin của Shannon (1948) được triết học khảo cứu sâu sắc, nhưng chính mật mã học — cái nôi khai sinh lý thuyết thông tin — lại bị triết học bỏ quên suốt nhiều thập kỷ — bài báo của Halpin là công trình tiên phong giải quyết khoảng trống này."),
                    ("Cái giá của sự trừu tượng hóa:", "Khi mật mã học biến thành các bài toán toán học thuần túy trên giấy, nó tạo ra điểm mù nguy hiểm trước các đòn tấn công vật lý trong thế giới thực — khiến hệ thống dù có chứng minh an toàn toán học vẫn sụp đổ trước các kênh rò rỉ phụ.")
                ]),
                ("2. Ba Đóng Góp Học Thuật Chính Của Harry Halpin:", [
                    ("Khảo sát Lịch sử & Triết học:", "Hệ thống hóa toàn diện quá trình tiến hóa mô hình đối thủ: từ Shannon (1949), PPT Adversary (1984), Dolev-Yao (1983) đến Random Oracle Model (1993) — làm rõ sự chuyển dịch từ an toàn thông tin sang độ phức tạp tính toán."),
                    ("Phê phán Giới hạn Hình thức hóa:", "Chỉ rõ sự sụp đổ của các mô hình toán học trước thực tiễn: rò rỉ kênh kề (Kocher 1996), độ dài bản mã, rò rỉ siêu dữ liệu và giám sát hàng loạt của NSA (Snowden 2013) — khẳng định an ninh lý thuyết không đồng nghĩa với an ninh vật lý."),
                    ("Đề xuất Chương trình Mới:", "Thiết lập 4 câu hỏi nghiên cứu mở (Section 7) nhằm mở rộng mật mã học từ một nhánh toán học biệt lập thành một khoa học an ninh hoàn chỉnh kết hợp thực nghiệm — định hướng liên ngành giữa triết học, an ninh mạng và khoa học xã hội.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Ba đóng góp này tạo nên sợi chỉ đỏ xuyên suốt: Phê phán sự cô lập của mô hình toán học thuần túy và kêu gọi đưa mật mã học trở về với thế giới vật lý thực tế.",
            "footer": "Trích dẫn: Harry Halpin (2025), Section 1: Introduction, Journal of Cybersecurity, tr. 1–2."
        },

        # SLIDE 3
        {
            "num": 3,
            "title": "CƠ SỞ TRIẾT HỌC: THÔNG TIN, BÍ MẬT VÀ TÁC NHÂN",
            "time": "1:00",
            "header": "SECTION 2 & 3 — CƠ SỞ TRIẾT HỌC VỀ THÔNG TIN (TRANG 2–5)",
            "lead": "Khung nhận thức luận của Fred Dretske và Brian Cantwell Smith ứng dụng trong Mật mã học",
            "sections": [
                ("1. Bản Thể Học Thông Tin — Brian Cantwell Smith (1996):", [
                    ("Dòng chảy nguyên thủy:", "Thế giới khởi đầu từ dòng chảy siêu hình không biên giới (primordial metaphysical flux) — mọi ranh giới đều do con người và hệ thống kiến tạo."),
                    ("Phân ranh giới:", "Con người và hệ thống định hình thực tại bằng cách đăng ký đối tượng (registration of objects) — trong mật mã, thông điệp M, bản mã C và khóa K được đăng ký để kiểm soát luồng dữ liệu."),
                    ("Tác nhân xa (Distal):", "Thông tin mật mã ra đời để truyền đạt trạng thái giữa các thực thể ở xa nhau mà không thể tương tác trực tiếp — tạo cơ sở bản thể luận cho bài toán liên lạc từ xa qua mạng.")
                ]),
                ("2. Thuyết Thông Tin Khách Quan — Fred Dretske (1981):", [
                    ("Thông tin khách quan:", "Thông tin là tính quy luật của trạng thái sự vật (state of affairs), hoàn toàn độc lập với việc có ai đọc hay không — mang tính quy luật tự nhiên."),
                    ("Thông tin và Bí mật:", "Bản rõ chứa đựng thông tin về nguồn; bản mã khi được mã hóa hoàn hảo sẽ không làm rò rỉ bất kỳ thông tin nào về nguồn — ngăn chặn đối thủ suy diễn bản rõ."),
                    ("Dòng chảy thông tin:", "Nhiệm vụ của mật mã là kiểm soát dòng chảy thông tin chỉ đi đến người nhận được ủy quyền — ngăn chặn phát tán đến các tác nhân trái phép.")
                ]),
                ("3. Tác Nhân Có Chủ Đích — Bản chất của đối thủ:", [
                    ("Tính chủ đích (Intentionality):", "Đối thủ không phải là 'nhiễu ngẫu nhiên' mà là thực thể có niềm tin, mục tiêu và ý chí khai thác lỗ hổng — phân biệt triệt để mật mã với truyền thông vật lý thông thường."),
                    ("Tính tương đối của bí mật:", "Bí mật không tồn tại tự thân; nó là mối quan hệ giữa nhóm người giữ bí mật và thế giới bên ngoài trong một khoảng thời gian — sẽ mất đi khi mối quan hệ này bị phá vỡ."),
                    ("Public vs Common Knowledge:", "Common Knowledge (David Lewis 1969): Mọi tác nhân đều biết và tin trạng thái sự việc. Public Knowledge: Thông tin mà bất kỳ ai cũng có tiềm năng tiếp cận (như khi bị rò rỉ). Bí mật là 'tri thức không phổ biến', không đồng nhất với bản rõ hay bản mã.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Hai trụ cột triết học này xác lập rằng Kẻ đối đầu mật mã học là tác nhân có chủ đích, và bí mật là mối quan hệ mang tính vị thế xã hội chứ không phải thuộc tính tự thân của chuỗi bit.",
            "footer": "Trích dẫn khoa học: Dretske (1981, Ref [1]); Cantwell Smith (1996, Ref [5]); Halpin (2025), tr. 2–5."
        },

        # SLIDE 4
        {
            "num": 4,
            "title": "TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ TRONG MẬT MÃ",
            "time": "1:00",
            "header": "SƠ ĐỒ 1: DÒNG THỜI GIAN TIẾN HÓA MÔ HÌNH ĐỐI THỦ",
            "lead": "Sự dịch chuyển từ Lý thuyết thông tin sang Giả định độ phức tạp tính toán và Phê phán hậu Snowden",
            "sections": [
                ("Nội dung dòng thời gian 5 giai đoạn:", [
                    ("Giai đoạn 1 (1948–1949) — [MÔ HÌNH SHANNON]:", "Mô hình truyền thông bảo mật điểm - điểm; định nghĩa Bí mật hoàn hảo H(K) >= H(M). Giới hạn: Giản lược đối thủ thành kẻ nghe lén trên kênh trừu tượng, bỏ qua chi phí phân phối khóa và thực tế xã hội - kỹ thuật — đặt nền móng lý thuyết nhưng cô lập với thực tế triển khai."),
                    ("Giai đoạn 2 (1976) — [CÁCH MẠNG KHÓA CÔNG KHAI]:", "Diffie-Hellman và Merkle giải quyết bài toán phân phối khóa; chuyển gánh nặng bảo mật sang độ phức tạp tính toán một chiều — khởi xướng thời kỳ bảo mật dựa trên độ khó toán học."),
                    ("Giai đoạn 3 (1982–1984) — [PPT ADVERSARY & SECURITY GAMES]:", "Goldwasser-Micali định nghĩa an toàn ngữ nghĩa và đối thủ thời gian đa thức xác suất (PPT). Chứng minh an toàn bằng quy dẫn toán học — chuẩn hóa phương pháp luận chứng minh hình thức."),
                    ("Giai đoạn 4 (1981–1993) — [DOLEV-YAO & RANDOM ORACLE]:", "Dolev-Yao mô hình hóa đối thủ mạng kiểm soát hoàn toàn kênh truyền. Bellare-Rogaway đề xuất Random Oracle Model (ROM) ứng dụng thực hành — áp dụng kiểm chứng logic cho giao thức và hàm băm."),
                    ("Giai đoạn 5 (2013–2025) — [PHÊ PHÁN HẬU SNOWDEN]:", "Tiết lộ của Snowden phơi bày năng lực thực tế của NSA: Kênh kề, cửa sau chuẩn mật mã, tấn công chuỗi cung ứng và giám sát siêu dữ liệu — bộc lộ khoảng cách giữa mô hình toán học và đối thủ thực tế.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Dòng thời gian thể hiện quy luật biện chứng: Khi mật mã học nâng độ trừu tượng toán học lên đỉnh cao thì thực tiễn an ninh thế giới thực lại kéo ngành này trở về với các bài toán vật lý và thể chế.",
            "footer": "Sơ đồ tổng hợp dựa trên Section 3–6, Harry Halpin (2025), Journal of Cybersecurity, tr. 3–12."
        },

        # SLIDE 5
        {
            "num": 5,
            "title": "MÔ HÌNH TRUYỀN THÔNG CỦA SHANNON VÀ MẠNG THỰC TẾ",
            "time": "1:00",
            "header": "SƠ ĐỒ 2: ĐỐI SÁNH MÔ HÌNH TRUYỀN THÔNG ĐIỂM-ĐIỂM VÀ MẠNG XÃ HỘI-KỸ THUẬT",
            "lead": "Khoảng cách giữa mô hình điểm-điểm đóng kín của Shannon và mạng phức hợp không đồng nhất",
            "sections": [
                ("Nội dung đối sánh 2 Panel trên Sơ đồ:", [
                    ("Panel Trái — Mô hình điểm-điểm cổ điển của Shannon (1949):", "Alice gửi bản rõ M mã hóa bằng khóa K qua kênh truyền đến Bob. Đối thủ Eve chỉ là kẻ nghe lén thụ động (passive eavesdropper) trên đường truyền. Giới hạn: Trừu tượng hóa đối thủ trên kênh toán học độc lập, bỏ qua chi phí lưu trữ/phân phối khóa và không bao quát các bề mặt tấn công vật lý — chỉ phù hợp với liên lạc dây dẫn cô lập ban đầu."),
                    ("Panel Phải — Mạng xã hội - kỹ thuật phức hợp theo Halpin (2025):", "Thông tin luân chuyển giữa Nhóm giữ bí mật (Secret Keepers) và Hạ tầng mở / ISP / Đám mây. Đối thủ thể chế (NSA, Five Eyes) là tác nhân toàn cục cả thụ động lẫn chủ động, tấn công qua 4 con đường: (1) Kênh kề thời gian/điện năng; (2) Giám sát siêu dữ liệu luồng gói tin; (3) Cài cắm cửa sau chuỗi cung ứng; (4) Lỗi phản bội nội bộ Byzantine — đòi hỏi phòng thủ hiện đại phải bảo vệ cả siêu dữ liệu và kháng kênh kề.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Sự đối sánh này làm rõ luận điểm của Halpin: An ninh không chỉ là mã hóa nội dung giữa hai điểm, mà là bảo vệ toàn bộ mạng xã hội và hạ tầng trước sự giám sát toàn cục.",
            "footer": "Sơ đồ tổng hợp dựa trên Shannon (1949) và Harry Halpin (2025), Journal of Cybersecurity, tr. 3–5."
        },

        # SLIDE 6
        {
            "num": 6,
            "title": "BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI",
            "time": "1:00",
            "header": "SECTION 4 — HÌNH THỨC HÓA ĐỐI THỦ (TRANG 5–8)",
            "lead": "Tiến trình toán học hóa kẻ thù: Độ phức tạp tính toán đa thức, Logic kênh truyền và Tiên tri ngẫu nhiên",
            "sections": [
                ("1. Mô Hình PPT Adversary — Goldwasser-Micali (1984):", [
                    ("Khái niệm:", "Đối thủ là thuật toán Turing xác suất chạy trong thời gian đa thức (Probabilistic Polynomial-Time) — giới hạn năng lực đối thủ ở mức khả thi tính toán."),
                    ("Security Games:", "Định nghĩa an toàn thông qua các trò chơi thách đấu. Lợi thế Adv(A) phải là hàm không đáng kể (negligible) — đối thủ không đoán tốt hơn ngẫu nhiên."),
                    ("Quy dẫn toán học:", "Chứng minh an toàn bằng cách quy dẫn: bẻ gãy hệ mật tương đương giải bài toán khó (phân tích thừa số nguyên, logarit rời rạc) — gắn an toàn với giả thuyết toán học.")
                ]),
                ("2. Mô Hình Đối Thủ Mạng Logic — Dolev-Yao (1983):", [
                    ("Khái niệm:", "Đối thủ kiểm soát toàn bộ kênh truyền mạng: có quyền đọc, chặn, sửa đổi, xóa, tráo đổi và phát lại gói tin — chuẩn an toàn cho mạng thù địch."),
                    ("Hộp đen mật mã:", "Mô hình Dolev-Yao xem mật mã như các hàm đại số hoàn hảo (Perfect Black-box). Kẻ thù không thể giải mã nếu không có khóa — tách biệt logic mạng với thuật toán."),
                    ("Ứng dụng:", "Kiểm chứng hình thức các giao thức mạng phức tạp thông qua các công cụ logic tự động — chứng minh tính đúng đắn khi trao đổi khóa.")
                ]),
                ("3. Mô Hình Tiên Tri Ngẫu Nhiên — Bellare-Rogaway (1993):", [
                    ("Khái niệm:", "Giả định hàm băm thực tế hoạt động như một hàm ngẫu nhiên lý tưởng (Random Oracle Model - ROM) trả về chuỗi ngẫu nhiên đồng đều — mốc tham chiếu toán học lý tưởng."),
                    ("Vai trò thực tế:", "Cho phép xây dựng các hệ mật khóa công khai và chữ ký số hiệu năng cao có chứng minh an toàn — thiết kế giao thức thực tế vận hành nhanh."),
                    ("Giới hạn triết học:", "Trong thế giới thực, không có hàm băm cụ thể nào (kể cả SHA-256) có thể tương đương toán học với một Random Oracle thực sự — do mọi hàm băm thực tế đều có cấu trúc toán học xác định.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Cả ba mô hình đều đem lại công cụ toán học mạnh mẽ, nhưng sự hoàn hảo của chúng dựa trên các giả định trừu tượng mà thế giới thực thường xuyên vi phạm.",
            "footer": "Trích dẫn: Goldwasser-Micali (1984); Dolev-Yao (1983); Bellare-Rogaway (1993); Halpin (2025), tr. 5–8."
        },

        # SLIDE 7
        {
            "num": 7,
            "title": "GIỚI HẠN CỦA SỰ HÌNH THỨC HÓA TRONG MẬT MÃ HỌC",
            "time": "1:00",
            "header": "SECTION 5 — GIỚI HẠN TRỪU TƯỢNG HÓA (TRANG 8–10)",
            "lead": "Khoảng cách giữa các giả định toán học lý tưởng và thế giới vật lý thực tế",
            "sections": [
                ("Tiến trình 4 bước từ Giả định đến Thực tế:", [
                    ("BƯỚC 1: GIẢ ĐỊNH TOÁN HỌC HÌNH THỨC:", "Hộp đen mật mã lý tưởng, thời gian thực thi là vô hình, độ dài bản mã công khai không rò rỉ nội dung, kênh truyền điểm-điểm độc lập — thế giới phẳng hoàn hảo trên chứng minh giấy tờ."),
                    ("BƯỚC 2: CÁC YẾU TỐ VẬT LÝ BỊ TRỪU TƯỢNG HÓA LƯỢC BỎ:", "Thời gian vi xử lý CPU, dao động điện năng tiêu thụ, phát xạ điện từ, kiến trúc bộ nhớ đệm (Cache), lỗi lập trình phần mềm — các hiện tượng vật lý tất yếu khi chạy trên phần cứng."),
                    ("BƯỚC 3: CON ĐƯỜNG KHAI THÁC CỦA ĐỐI THỦ THẾ GIỚI THỰC:", "Tấn công kênh kề (Timing Attack - Kocher 1996), phân tích độ dài và lưu lượng siêu dữ liệu (Traffic Analysis), khai thác lỗi cài đặt phần mềm và quản lý (Anderson 2001) — kẻ thù chiếm ưu thế mà không cần bẻ gãy toán học."),
                    ("BƯỚC 4: KẾT LUẬN NHẬN THỨC LUẬN CỦA HARRY HALPIN:", "Cái giá của sự trừu tượng hóa là làm cho mật mã học mù quáng trước chính những năng lực thực tế của đối thủ nằm ngoài mô hình lý thuyết — buộc an ninh mạng phải mở rộng ra kiểm định thực nghiệm.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Sơ đồ 4 bước này khẳng định luận điểm cốt tử của Section 5: Hệ thống an toàn trên giấy tờ toán học hoàn toàn có thể bị đánh bại bởi các kênh vật lý bên ngoài mô hình.",
            "footer": "Sơ đồ tổng hợp dựa trên Kocher (1996) và Halpin (2025), Section 5, tr. 8–10."
        },

        # SLIDE 8
        {
            "num": 8,
            "title": "SNOWDEN VÀ ĐỐI THỦ THỂ CHẾ (THE REAL ADVERSARY)",
            "time": "1:00",
            "header": "SECTION 6 — TIẾT LỘ CỦA SNOWDEN (TRANG 10–12)",
            "lead": "Khủng hoảng mô hình an ninh trước đối thủ giám sát hàng loạt quy mô quốc gia",
            "sections": [
                ("1. Khủng Hoảng Thực Tiễn — Tiết lộ Snowden (2013):", [
                    ("Sự kiện bước ngoặt:", "Các tài liệu Edward Snowden tiết lộ về NSA đã kéo mật mã học vào một cuộc khủng hoảng thực tiễn sâu sắc — làm lung lay niềm tin vào sự an toàn tuyệt đối của chuẩn mật mã."),
                    ("Không phá toán học:", "NSA không cần giải các bài toán toán học khó của RSA hay AES mà tấn công trực tiếp vào toàn bộ hệ sinh thái xung quanh — bằng cách đánh cắp khóa và trích xuất siêu dữ liệu."),
                    ("Mô hình thực tế:", "Đối thủ bị động toàn cầu (Global Passive Adversary) và tấn công can thiệp vào đường truyền là mối đe dọa hoàn toàn có thật — vượt ngoài khuôn khổ bảo mật điểm-điểm.")
                ]),
                ("2. Năng Lực Đối Thủ Thể Chế — Vượt ngoài hình thức:", [
                    ("Giám sát Siêu dữ liệu:", "NSA tập trung khai thác Metadata (ai liên lạc với ai, thời gian, kích thước) để dựng lại mạng xã hội mà không cần bẻ khóa bản mã — khẳng định siêu dữ liệu nguy hiểm ngang nội dung."),
                    ("Tấn công Chuỗi cung ứng:", "Chèn cửa sau vào phần cứng, làm suy yếu chuẩn mật mã quốc tế (tiêu chuẩn NIST Dual_EC_DRBG bị cài cửa sau ngầm) — làm xói mòn lòng tin vào cơ quan tiêu chuẩn hóa."),
                    ("Áp lực Thể chế:", "Cưỡng chế pháp lý các tập đoàn công nghệ lớn bàn giao khóa bí mật và quyền truy cập dữ liệu máy chủ — biến rào cản kỹ thuật thành thỏa hiệp pháp lý kín.")
                ]),
                ("3. Tính Tương Đối Của Kẻ Thù — Bản chất nhận thức luận:", [
                    ("Khái niệm tương đối:", "Khái niệm 'Kẻ thù' mang tính vị thế: Đối với người dân, NSA là đối thủ giám sát; nhưng đối với NSA, Snowden lại là đối thủ làm rò rỉ bí mật — an ninh luôn gắn với góc nhìn quyền lực."),
                    ("Dân chủ hóa mật mã:", "Mật mã khóa công khai từng giúp cá nhân tự vệ, nhưng công cụ gián điệp của NSA nay đã lan rộng sang tội phạm mạng và các nhóm phi nhà nước — tạo thế trận đối kháng đa cực."),
                    ("Hệ quả khoa học:", "Mật mã học buộc phải từ bỏ giả định kênh truyền sạch để nghiên cứu phòng chống cửa sau (Cliptography) — xây dựng hệ thống an toàn ngay cả khi thuật toán bị cài cắm ngầm.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Sự kiện Snowden chứng minh rằng đối thủ thế giới thực không phải là các thuật toán trừu tượng mà là các thể chế quyền lực có nguồn lực tài chính, chính trị và kỹ thuật gần như không giới hạn.",
            "footer": "Trích dẫn: Harry Halpin (2025), Section 6: Snowden and the real adversary, tr. 10–12."
        },

        # SLIDE 9
        {
            "num": 9,
            "title": "CHƯƠNG TRÌNH NGHIÊN CỨU TƯƠNG LAI: BỐN CÂU HỎI MỞ",
            "time": "1:00",
            "header": "SECTION 7 — KẾT LUẬN & CHƯƠNG TRÌNH NGHIÊN CỨU (TRANG 12–14)",
            "lead": "Đề xuất hòa giải giữa trừu tượng hóa toán học và dữ liệu thực nghiệm thế giới thực",
            "sections": [
                ("Bốn câu hỏi nghiên cứu mở nguyên văn từ Halpin:", [
                    ("Câu hỏi 1: Mạng Xã Hội Động của Truyền Thông Bí Mật:", '"Can we study dynamic social networks of secret communication modeled on the empirical behavior of secret keepers?" — Nghiên cứu tổ chức thông tin dưới dạng mạng xã hội động giữa những người giữ bí mật. Kiểm tra xem các bất biến chính sách bảo vệ bí mật có được duy trì qua các biến động cấu trúc mạng và các sự cố Byzantine hay không — giúp an ninh thích ứng linh hoạt trước thay đổi tác nhân.'),
                    ("Câu hỏi 2: Mô Hình Hóa Năng Lực và Liên Minh Đối Thủ Bằng Thống Kê:", '"By incorporating statistics, can new models of the adversary take into account more accurate data about their capabilities?" — Sử dụng thống kê thực nghiệm và mô phỏng Monte Carlo để mô hình hóa chính xác hơn năng lực đối thủ; phân tích kinh tế chính trị của các liên minh giám sát (Five Eyes, Anonymous) — lượng hóa nguy cơ bằng dữ liệu định lượng thực tế.'),
                    ("Câu hỏi 3: Ứng Dụng Phương Pháp Liên Ngành & Hệ Thống Phức Hợp:", '"Can cryptography incorporate interdisciplinary methods and insights that can capture the complexity of secure communication?" — Áp dụng các công cụ từ lý thuyết hệ thống phức hợp, động lực học mạng và dịch tễ học để mô hình hóa sự lan truyền thông tin, nhận diện các điểm chuyển pha khi bị tấn công — mở rộng mật mã sang khoa học tự nhiên.'),
                    ("Câu hỏi 4: Triết Học Giúp Tái Cấu Trúc Các Giả Định Ngầm Định:", '"Will philosophy let us make explicit and then reformulate the hidden assumptions of cryptographic communication, and vice versa?" — Vận dụng triết học để chất vấn các giả định nền tảng của Shannon; xem xét ranh giới động giữa con người và môi trường kỹ thuật số theo Giả thuyết Tâm trí Mở rộng (Extended Mind Hypothesis) — tái cấu trúc nhận thức luận về an toàn thông tin.')
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: 4 câu hỏi này là bản thiết kế mở cho tương lai, định hướng mật mã học chuyển mình thành một khoa học an ninh toàn diện và thực nghiệm.",
            "footer": "Trích dẫn nguyên văn 4 câu hỏi nghiên cứu mở: Harry Halpin (2025), Section 7, tr. 12–14."
        },

        # SLIDE 10
        {
            "num": 10,
            "title": "PHƯƠNG PHÁP LUẬN & KHẢ NĂNG TÁI LẬP BÀI BÁO",
            "time": "1:00",
            "header": "ĐÁNH GIÁ PHƯƠNG PHÁP LUẬN & HƯỚNG TIẾP CẬN THỰC NGHIỆM",
            "lead": "Làm rõ tính khả thi tái lập đối với một công trình triết học và định hướng thí nghiệm độc lập",
            "sections": [
                ("1. Đánh Giá Khả Năng Tái Lập (Replication Feasibility):", [
                    ("Bản chất công trình:", "Bài báo của TS. Harry Halpin là một luận văn tổng quan triết học và phân tích nhận thức luận, không phải là một công trình kỹ thuật phần mềm — khơi mở tư duy chứ không cung cấp mã nguồn."),
                    ("Không có tài nguyên phần mềm:", "Bài báo HOÀN TOÀN KHÔNG có mã nguồn triển khai, không có tập dữ liệu thực nghiệm, không có các bảng đo thời gian (ms) hay dung lượng benchmark — không tồn tại hệ chuẩn kỹ thuật gốc từ tác giả."),
                    ("Phương pháp thẩm định:", "Nhóm sinh viên thực hiện thẩm định học thuật (Scholarly Audit) đối chiếu toàn bộ 17 trang bản in PDF, 7 mục nội dung và 134 tài liệu tham khảo — đảm bảo nắm bắt chuẩn xác và có phê phán."),
                    ("Tính trung thực học thuật:", "Không bịa đặt số liệu gốc của tác giả; không tuyên bố 'tái lập 100% benchmark bài báo' khi tác giả không công bố benchmark — giữ vững đạo đức khoa học chuẩn mực.")
                ]),
                ("2. Định Hướng Thí Nghiệm Minh Họa Độc Lập:", [
                    ("Mục tiêu thực nghiệm:", "Sinh viên tự xây dựng hệ thống kiểm tra toàn vẹn tệp nhằm chuyển hóa các luận điểm lý thuyết trừu tượng của Halpin thành các kịch bản quan sát được — cung cấp bằng chứng thực nghiệm trực quan."),
                    ("Minh họa trực quan:", "Thiết kế các kịch bản kiểm thử: đối đầu toán học trên mã tuyến tính CRC32; can thiệp đường truyền Dolev-Yao; rò rỉ kênh kề thời gian so sánh chuỗi — đưa người nghe vào trải nghiệm an ninh thực tế."),
                    ("Phân định ranh giới rõ ràng:", "Mã nguồn và hệ thống kiểm tra toàn vẹn là đóng góp kỹ thuật độc lập của sinh viên; không được nhầm lẫn là thuật toán do Halpin đề xuất — minh bạch tuyệt đối về bản quyền ý tưởng."),
                    ("Kiểm toán khoa học:", "Tất cả quy trình kiểm toán mã nguồn và đối chiếu thực nghiệm được lưu vết kỹ thuật tại MIDTERM_CODE_AUDIT.md — sẵn sàng cho hội đồng thẩm định chi tiết.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Slide 10 khẳng định nguyên tắc nghiên cứu trung thực: Bài báo là nền tảng triết học lý thuyết; hệ thống phần mềm là công cụ minh họa độc lập do sinh viên tự phát triển.",
            "footer": "Báo cáo phân định phương pháp luận độc lập của nhóm nghiên cứu đề tài DT005."
        },

        # SLIDE 11
        {
            "num": 11,
            "title": "MINH HỌA PAPER-TO-CODE ĐỘC LẬP & GIỚI HẠN DIỄN GIẢI",
            "time": "1:15",
            "header": "SƠ ĐỒ 4: PHÂN ĐỊNH LÝ THUYẾT & THÍ NGHIỆM ĐỘC LẬP",
            "lead": "Đối sánh minh bạch giữa các luận điểm bài báo và thí nghiệm minh họa do sinh viên xây dựng",
            "sections": [
                ("1. NHIỄU vs ĐỐI THỦ CHỦ ĐÍCH:", [
                    ("Luận điểm Halpin (Sec. 1–2):", "Đối thủ là tác nhân có chủ đích, không phải nhiễu ngẫu nhiên. Mã kiểm tra lỗi tuyến tính được thiết kế cho nhiễu ngẫu nhiên, không phòng vệ được trước đối thủ chủ đích."),
                    ("Minh họa độc lập (DT005):", "Thử nghiệm kịch bản bù 4 byte GF(2): Kẻ tấn công giải hệ phương trình tuyến tính để qua mặt CRC32 trong tệp bị can thiệp; hàm băm phi tuyến SHA-256 phát hiện tức thì."),
                    ("Giới hạn diễn giải:", "Chỉ minh họa tính phi tuyến của hàm băm trong kịch bản giả mạo đã kiểm thử; không kết luận CRC32 bất lực trong vai trò phát hiện lỗi ngẫu nhiên vật lý — bởi CRC32 vẫn giữ nguyên giá trị kiểm tra lỗi mạng và phần cứng truyền thống.")
                ]),
                ("2. PHÂN TÁCH MIỀN NÚT LÁ (RFC 6962):", [
                    ("Gốc lý thuyết bài báo:", "Halpin nhấn mạnh điểm mù khi triển khai cấu trúc dữ liệu trong thực tế, nhưng KHÔNG trực tiếp đề cập cấu trúc cây Merkle hay kỹ thuật phân tách miền."),
                    ("Minh họa độc lập (DT005):", "Sinh viên chủ động mở rộng bài toán toàn vẹn: Cài đặt phân tách miền theo chuẩn IETF RFC 6962 (tiền tố byte 0x00 cho lá, 0x01 cho nhánh) để phòng ngừa va chạm cấu trúc tầng cây."),
                    ("Giới hạn diễn giải:", "Kỹ thuật RFC 6962 do sinh viên tự nghiên cứu tích hợp, không phải đề xuất của Halpin; chỉ phòng ngừa va chạm cấu trúc, không thay đổi độ phức tạp 2^256 của SHA-256 — đây là đóng góp công nghệ bổ trợ của nhóm.")
                ]),
                ("3. RÒ RỈ KÊNH KỀ THỜI GIAN:", [
                    ("Luận điểm Halpin (Sec. 5):", "Mô hình toán học trừu tượng xem mật mã là hàm lý tưởng, nhưng thời gian thực thi thực tế làm rò rỉ dữ liệu bí mật (dẫn theo Paul Kocher 1996)."),
                    ("Minh họa độc lập (DT005):", "Sinh viên xây dựng kịch bản kiểm thử đo đạc thời gian so sánh chuỗi tuần tự (dừng sớm khi gặp byte lệch) so với việc sử dụng hàm thời gian hằng số hmac.compare_digest."),
                    ("Giới hạn diễn giải:", "Đo đạc trên môi trường máy tính đơn có jitter hệ điều hành; chỉ minh họa rò rỉ logic ở mức phần mềm ứng dụng, không đại diện cho kênh kề phần cứng chuyên sâu — phục vụ mục đích trực quan hóa lỗ hổng kênh kề mức mã nguồn.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Slide 11 đối chiếu 3 cột rạch ròi: Luận điểm lý thuyết của bài báo, Mã nguồn minh họa độc lập của sinh viên, và Giới hạn diễn giải khoa học chính xác.",
            "footer": "Bảng đối chiếu minh bạch do sinh viên tự xây dựng; chi tiết mã nguồn lưu tại MIDTERM_CODE_AUDIT.md."
        },

        # SLIDE 12
        {
            "num": 12,
            "title": "SLIDE CHUYỂN TIẾP: DEMO",
            "time": "0:15",
            "header": "GIAO DIỆN TỐI GIẢN NỀN SÁNG CHUẨN BỊ DIỄN TẬP TRỰC TIẾP",
            "lead": "Chữ hiển thị trên Slide: DEMO",
            "sections": [
                ("Nội dung thể hiện trên Slide:", [
                    ("Văn bản trung tâm:", "DEMO (Được thiết kế tối giản, chữ to rõ ràng trên nền sáng thanh lịch để tạo nhịp nghỉ trang trọng chuyển tiếp phần thuyết trình sang phần thực nghiệm).")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Sau khi hoàn thành phần đối chiếu lý thuyết, nhóm xin chuyển sang phần Diễn tập An ninh Thực nghiệm trực tiếp (Live Demo) thông qua mô hình đối kháng 2 màn hình Red Team và Blue Team, tập trung kiểm chứng các luận điểm mấu chốt của Halpin.",
            "footer": "Slide chuyển tiếp phần diễn tập thực hành đối kháng an ninh."
        },

        # SLIDE 13
        {
            "num": 13,
            "title": "KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG: RED TEAM VS BLUE TEAM",
            "time": "1:00",
            "header": "SƠ ĐỒ 5: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG (DEMO ARCHITECTURE)",
            "lead": "Mô hình đối kháng thực tế 2 màn hình trên tập dữ liệu kaggle_dataset (2.485 tệp tin / ~112,81 MiB theo benchmark)",
            "sections": [
                ("Nội dung thành phần kiến trúc diễn tập trên Sơ đồ:", [
                    ("Hạ tầng diễn tập 2 Màn hình:", "Tương tác thời gian thực giữa Hệ thống giám sát phòng thủ Blue Team (Port 8080) và Động cơ tấn công đối kháng Red Team (Port 8081)."),
                    ("Tập dữ liệu thực nghiệm:", "kaggle_dataset gồm 2.485 tệp tin, dung lượng ~112,81 MiB (phiên bản nạp kiểm thử cục bộ hiện tại gồm 2.483 tệp) đóng vai trò môi trường kiểm thử thực tế — phản ánh cấu trúc cây thư mục doanh nghiệp nhiều tầng."),
                    ("Blue Team (Port 8080):", "Nạp dữ liệu, thiết lập mốc tin cậy sạch ban đầu H_0; khi hệ thống kiểm tra toàn vẹn, cơ chế duyệt cây phân cấp định vị chính xác vị trí tệp tin bị biến đổi với tốc độ quét thực tế vượt 2.000 tệp/giây — bắt quả tang vi phạm toàn vẹn tức thì."),
                    ("Red Team (Port 8081):", "Mô phỏng các vector can thiệp thực tế: Sửa byte tệp ngẫu nhiên (In-place Tamper) và giải hệ phương trình tuyến tính trên GF(2) bù 4 byte qua mặt mã kiểm tra CRC32 — hiện thực hóa năng lực kẻ đối đầu có chủ đích."),
                    ("Cơ chế hoàn nguyên an toàn:", "Thư mục sao lưu an toàn .adversary_backups/ cho phép phục hồi nguyên vẹn 100% dữ liệu chỉ với 1 click chuột Rollback — đảm bảo ổ đĩa luôn được bảo toàn sau diễn tập.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Sơ đồ kiến trúc thể hiện sự kết nối hoàn chỉnh giữa phòng thủ và tấn công, cho phép kiểm chứng các luận điểm của Halpin một cách trực quan và an toàn tuyệt đối.",
            "footer": "Sơ đồ kiến trúc diễn tập đối kháng độc lập do sinh viên thiết kế minh họa các khái niệm an ninh."
        },

        # SLIDE 14
        {
            "num": 14,
            "title": "TỔNG KẾT BÁO CÁO GIỮA KỲ & ĐỊNH HƯỚNG NGHIÊN CỨU",
            "time": "0:45",
            "header": "SECTION 7 — KẾT LUẬN & THÔNG ĐIỆP HỌC THUẬT",
            "lead": "Tổng kết các đóng góp lý thuyết của Halpin và định hướng nghiên cứu tiếp theo",
            "sections": [
                ("1. Thông Điệp Học Thuật Cốt Lõi (Harry Halpin 2025):", [
                    ("Vượt qua sự trừu tượng hóa:", "Mật mã học không chỉ là toán học thuần túy trên giấy tờ. Các hệ thống an toàn trong thế giới thực phải nhận thức rõ khoảng cách giữa mô hình lý tưởng và thực tế vật lý (kênh kề, siêu dữ liệu, chuỗi cung ứng, áp lực thể chế) — đây là chìa khóa để xây dựng an ninh thực chiến."),
                    ("Bản chất của đối thủ:", "Đối thủ không phải là 'nhiễu ngẫu nhiên' mà là tác nhân có chủ đích, liên tục tìm kiếm các kênh rò rỉ bị mô hình toán học bỏ qua — an ninh phải dựa trên tư duy đối kháng toàn diện."),
                    ("Chương trình nghiên cứu mở:", "Halpin đặt nền móng triết học và vạch ra 4 câu hỏi nghiên cứu liên ngành nhằm mở rộng mật mã học thành một khoa học an ninh hoàn chỉnh — tạo cầu nối giữa toán học, máy tính và triết học.")
                ]),
                ("2. Hoàn Thành Giữa Kỳ & Lộ Trình Cuối Kỳ (Nhóm DT005):", [
                    ("Nghiệm thu tiến độ giữa kỳ:", "Hoàn thành thẩm định học thuật sâu sắc toàn văn bài báo gốc (17 trang PDF), phân tích đầy đủ 7 mục nội dung và làm rõ khoảng cách giữa lý thuyết trừu tượng với an ninh thực tế — đáp ứng chuẩn mực đồ án giữa kỳ."),
                    ("Minh bạch phương pháp luận:", "Không ngụy tạo số liệu tác giả; phân định rõ ràng giữa công trình khảo cứu triết học của Halpin và các kịch bản thực nghiệm minh họa độc lập do sinh viên tự xây dựng — khẳng định trách nhiệm khoa học."),
                    ("Định hướng nghiên cứu cuối kỳ:", "Tiếp tục nghiên cứu và hoàn thiện mô hình cấu trúc cây phân cấp tối ưu cho bài toán kiểm tra toàn vẹn dữ liệu, tiến hành đo đạc và đánh giá thực nghiệm toàn diện — hướng đến bảo vệ đồ án tốt nghiệp cuối kỳ.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Đề tài giữa kỳ đã khép lại trọn vẹn phần khảo cứu học thuật và chuẩn bị nền móng kỹ thuật vững chắc cho việc phát triển các cấu trúc dữ liệu xác thực tối ưu ở giai đoạn cuối kỳ.",
            "footer": "Tài liệu tham khảo chính: Halpin (2025), Shannon (1949), Goldwasser-Micali (1984), Dolev-Yao (1983), Kocher (1996), RFC 6962."
        },

        # SLIDE 15
        {
            "num": 15,
            "title": "CẢM ƠN & PHIÊN PHẢN BIỆN (Q&A)",
            "time": "0:15",
            "header": "Trường ĐH Tài nguyên và Môi trường TP.HCM — Khoa Công nghệ thông tin",
            "lead": "CẢM ƠN THẦY VÀ CÁC BẠN ĐÃ LẮNG NGHE!",
            "sections": [
                ("Thông tin kết thúc trên Slide:", [
                    ("Thông điệp:", "Q & A — Xin đón nhận các câu hỏi và ý kiến đóng góp của Thầy và Hội đồng."),
                    ("Đề tài:", "Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động."),
                    ("Thông tin quản lý:", "Học phần: An toàn và bảo mật hệ thống thông tin | GVHD: ThS. Phạm Trọng Huynh | SVTH: Võ Anh Kiệt — MSSV: 1150080143 — Lớp: 11_CNPM2.")
                ])
            ],
            "expansion": "Lời thuyết minh tiếp nối: Phần trình bày báo cáo giữa kỳ của đề tài DT005 đến đây là kết thúc. Xin chân thành cảm ơn sự chú ý lắng nghe của Thầy cùng các bạn. Nhóm xin sẵn sàng lắng nghe các ý kiến nhận xét và giải đáp các câu hỏi trong phiên thảo luận.",
            "footer": "Phiên hỏi đáp học thuật và đóng góp ý kiến hoàn thiện đồ án."
        }
    ]

    for s in slides_data:
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
        r_hd = p_hd.add_run(f"• Tiêu đề phụ / Section trên Slide: {s['header']}\n  {s['lead']}")
        r_hd.font.size = Pt(9.5)
        r_hd.font.italic = True
        r_hd.font.color.rgb = RGBColor(71, 85, 105)

        # Các khối nội dung trên slide
        for sec_title, items in s['sections']:
            p_sec = doc.add_paragraph()
            p_sec.paragraph_format.space_before = Pt(3)
            p_sec.paragraph_format.space_after = Pt(1)
            p_sec.paragraph_format.left_indent = Inches(0.2)
            r_st = p_sec.add_run(sec_title)
            r_st.font.size = Pt(10.5)
            r_st.font.bold = True
            r_st.font.color.rgb = RGBColor(15, 23, 42)

            for h, b in items:
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

        # Lời thuyết minh mở rộng tiếp nối
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.space_before = Pt(3)
        p_exp.paragraph_format.space_after = Pt(2)
        p_exp.paragraph_format.left_indent = Inches(0.2)
        r_ex = p_exp.add_run(s['expansion'])
        r_ex.font.size = Pt(10)
        r_ex.font.italic = True
        r_ex.font.color.rgb = RGBColor(30, 64, 175)

        # Footer của slide
        p_ft = doc.add_paragraph()
        p_ft.paragraph_format.space_before = Pt(1)
        p_ft.paragraph_format.space_after = Pt(8)
        p_ft.paragraph_format.left_indent = Inches(0.2)
        r_ft = p_ft.add_run(f"• Ghi chú chân Slide: {s['footer']}")
        r_ft.font.size = Pt(9)
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
         "Ý nghĩa học thuật tiếp nối: Minh họa khái niệm 'đăng ký đối tượng' (registration of objects) theo Cantwell Smith và thiết lập trạng thái tin cậy khách quan theo Dretske."),

        ("Bước 3: Phát động Kịch bản 1 — Can thiệp tệp tin có chủ đích tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm nút 'Can thiệp tệp tin ngẫu nhiên' (In-place Tampering). Red Team sửa đổi trực tiếp dữ liệu tệp tin trên ổ đĩa (ví dụ: Resume/Resume.csv).\n"
         "Quay lại Blue Team (Port 8080), bấm nút 'Kiểm tra toàn vẹn':\n"
         "• Hệ thống ngay lập tức đổi sang trạng thái cảnh báo đỏ: 'CẢNH BÁO: PHÁT HIỆN CAN THIỆP DỮ LIỆU!'.\n"
         "• Bảng truy vết chỉ ra chính xác tệp tin Resume/Resume.csv bị tác động.\n"
         "Ý nghĩa học thuật tiếp nối: Minh chứng luận điểm Halpin — Kẻ đối đầu mật mã học là tác nhân có chủ đích (intentional agent), hàm băm mật mã học lập tức bắt quả tang mọi sai lệch dù chỉ 1 bit."),

        ("Bước 4: Phát động Kịch bản 2 — Tấn công bù 4 byte GF(2) qua mặt CRC32 tại Red Team (Port 8081)",
         "Tại cổng 8081, bấm nút 'Tấn công bù 4 byte GF(2)' (CRC32 Forge Attack). Động cơ Red Team giải hệ phương trình tuyến tính, làm sai lệch nội dung tệp nhưng bù 4 byte patch khiến cho mã CRC32 không đổi.\n"
         "Quay lại Blue Team (Port 8080), bấm nút 'Kiểm tra toàn vẹn':\n"
         "• Bảng đối chứng hiển thị rõ: Thuật toán kiểm tra lỗi CRC32 bị qua mặt hoàn toàn (báo khớp sai sự thật do tính chất tuyến tính của đa thức chia).\n"
         "• Ngược lại, hàm băm SHA-256 phát hiện vi phạm ngay lập tức nhờ tính phi tuyến và hiệu ứng tuyết lở (~50%).\n"
         "Ý nghĩa học thuật tiếp nối: Trực quan hóa luận điểm mấu chốt của bài báo Halpin — phân biệt rạch ròi giữa mã phát hiện lỗi ngẫu nhiên của Shannon và mật mã học trước kẻ thù có chủ đích."),

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
