# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word (.docx) Kịch Bản Thuyết Trình & Hướng Dẫn Diễn Tập Demo:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn:
- 100% bám sát cấu trúc 15 Slide hoàn chỉnh của Bao_Cao_Giua_Ky_DT05_Final.pptx / Bao_Cao_Giua_Ky_DT05_Final.pdf.
- Không dùng xưng hô rườm rà (loại bỏ toàn bộ "Thưa Thầy", "Trả lời Thầy", "Em xin kính chào...").
- Ngôn ngữ học thuật, khách quan, súc tích, mạch lạc.
- Tuân thủ nghiêm ngặt 9 yêu cầu sửa đổi:
  1. Slide 3: Định nghĩa Public Knowledge vs Common Knowledge theo chuẩn Halpin / Lewis (1969). Không đồng nhất với bản mã / bản rõ.
  2. Slide 4: Bỏ câu Shannon giả định sai lệch do nhiễu ngẫu nhiên. Mô tả chính xác mô hình bảo mật truyền thông Shannon và giới hạn đối thủ trừu tượng.
  3. Slide 7 & 11: Không gán Node-Leaf Confusion cho Halpin. Ghi rõ là minh họa độc lập do DT005 xây dựng theo RFC 6962.
  4. Slide 11: Phân biệt rõ luận điểm bài báo và mã nguồn DT005. Không tuyên bố CRC32 bất lực trong mọi trường hợp, chỉ mô tả kịch bản kiểm thử cụ thể GF(2).
  5. Slide 13: Đổi badge thành "KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG (DEMO ARCHITECTURE)". Đối chiếu số liệu dataset (2.485 tệp, ~112.81 MiB theo benchmark report; 2.483 tệp kiểm thử cục bộ). Bỏ các tuyên bố 100/100 hay 1.06 ms chưa xác minh.
  6. Slide 10: Xác nhận tài liệu kiểm toán MIDTERM_CODE_AUDIT.md tồn tại trong mã nguồn.
  7. Slide 14: Loại bỏ HD-Merkle, Dirty-Path Cache, CompactProof O(1) khỏi trọng tâm kết luận giữa kỳ; thay bằng định hướng nghiên cứu cuối kỳ khái quát.
  8. Slide 12: Sử dụng ảnh chụp dashboard thực tế từ lần chạy thành công (Blue Team bắt quả tang sửa đổi Resume.csv trên cổng 8080).
  9. Không đưa thang điểm 6-8, 9-10; không tuyên bố tái lập 100% bài báo.
"""

import sys
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

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_document():
    doc = docx.Document()

    # Thiết lập lề 1 inch chuẩn
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
    run_t.font.size = Pt(18)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(30, 64, 175) # Royal Blue

    run_sub = p_title.add_run("Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp\n")
    run_sub.font.size = Pt(13)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(15, 23, 42)

    run_paper = p_title.add_run('Khảo sát và Thẩm định học thuật bài báo: "The adversary: the philosophy of cryptography"\nTS. Harry Halpin — Journal of Cybersecurity, Oxford Academic, Vol. 11, No. 1, 2025\n')
    run_paper.font.size = Pt(10.5)
    run_paper.font.italic = True
    run_paper.font.color.rgb = RGBColor(100, 116, 139)

    # THÔNG TIN SINH VIÊN & HỌC PHẦN
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.paragraph_format.space_after = Pt(18)
    run_meta = p_info.add_run("Học phần: An toàn và bảo mật hệ thống thông tin | Giảng viên hướng dẫn: ThS. Phạm Trọng Huynh\nSinh viên thực hiện: Võ Anh Kiệt — MSSV: 1150080143 — Lớp: 11_CNPM2")
    run_meta.font.size = Pt(10.5)
    run_meta.font.bold = True
    run_meta.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # =========================================================================
    # PHẦN 1: KỊCH BẢN THUYẾT MINH THEO TỪNG SLIDE (15 SLIDE HOÀN CHỈNH)
    # =========================================================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("PHẦN 1: KỊCH BẢN THUYẾT MINH CHI TIẾT THEO 15 SLIDE")
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 64, 175)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(10)

    slides_content = [
        ("SLIDE 1: BÌA BÁO CÁO GIỮA KỲ", "0:30",
         "Nội dung trình bày:",
         "Báo cáo tiến độ đồ án giữa kỳ đề tài DT005: Xây dựng hệ thống kiểm tra toàn vẹn tệp dựa trên SHA-256, HMAC và Cây Merkle phân cấp; kết hợp thẩm định học thuật bài báo 'The adversary: the philosophy of cryptography' của Tiến sĩ Harry Halpin, xuất bản năm 2025 trên tạp chí Journal of Cybersecurity thuộc Đại học Oxford.\nNội dung báo cáo gồm 3 trọng tâm: (1) Khảo cứu các luận điểm nhận thức luận và tiến trình hình thức hóa kẻ đối đầu trong mật mã học; (2) Phân định phương pháp luận giữa khảo sát triết học lý thuyết của tác giả và hệ thống thực nghiệm độc lập do nhóm tự phát triển; (3) Diễn tập thực tế mô hình đối kháng an ninh Red Team vs Blue Team trên tập dữ liệu thực tế.",
         "Bố cục rõ ràng, nhận diện trang trọng, xác định chuẩn mục tiêu đồ án giữa kỳ."),

        ("SLIDE 2: BỐI CẢNH, MỤC TIÊU & BA ĐÓNG GÓP CỦA BÀI BÁO (SECTION 1)", "1:00",
         "Nội dung trình bày:",
         "Mật mã học trong kỷ nguyên số là khoa học về việc duy trì liên lạc bảo mật khi có sự hiện diện của đối thủ có chủ đích trên môi trường mạng mở.\nBài báo của Halpin chỉ ra một nghịch lý nhận thức luận sâu sắc: Trong khi lý thuyết thông tin của Claude Shannon được triết học khảo cứu rộng rãi suốt nhiều thập kỷ, thì chính mật mã học — cái nôi khai sinh ra khái niệm thông tin và tính ngẫu nhiên — lại bị giới triết học bỏ quên.\nBài báo đưa ra ba đóng góp học thuật cốt lõi: Thứ nhất, định hình lại lịch sử mật mã qua lăng kính triết học bản thể luận về 'Kẻ thù' (The Adversary). Thứ hai, phê phán sự trừu tượng hóa toán học thái quá khi biến đối thủ thành các hàm lý tưởng và bỏ qua các vector tấn công thế giới thực (kênh kề, siêu dữ liệu, lỗi phần cứng). Thứ ba, đề xuất một chương trình nghiên cứu liên ngành mở với 4 câu hỏi định hướng cho khoa học an ninh tương lai.",
         "Nêu bật nghịch lý triết học và 3 đóng góp lớn của Halpin."),

        ("SLIDE 3: CƠ SỞ TRIẾT HỌC VỀ THÔNG TIN & BẢN THỂ HỌC (SECTION 2)", "1:00",
         "Nội dung trình bày:",
         "Nền tảng lý thuyết của Halpin dựa trên các trụ cột triết học nhận thức luận và xã hội:\n1. Bản thể học thông tin của Brian Cantwell Smith (1996): Thực tại khởi nguồn từ một dòng chảy siêu hình nguyên thủy không biên giới. Con người và hệ thống định hình thế giới bằng cách 'đăng ký đối tượng' (registration of objects). Trong mật mã học, các đối tượng như thông điệp M, bản mã C và khóa K được đăng ký nhằm kiểm soát dòng thông tin giữa các tác nhân ở xa (distal agents).\n2. Thuyết thông tin khách quan của Fred Dretske (1981): Thông tin là một trạng thái khách quan gắn liền với quy luật tự nhiên, độc lập với người quan sát.\n3. Phân biệt tri thức công khai (Public Knowledge) và tri thức chung (Common Knowledge) theo David Lewis (1969): Tri thức chung là trạng thái mọi tác nhân trong một quần thể đều biết thông tin và đều biết rằng những người khác cũng biết. Trong khi đó, Tri thức công khai là khái niệm yếu hơn, chỉ thông tin mà bất kỳ tác nhân nào cũng có khả năng tiếp cận (chẳng hạn qua rò rỉ dữ liệu). Bí mật (Secrets) không phải là bản rõ hay bản mã, mà được định nghĩa là 'Tri thức không phổ biến' (Uncommon Knowledge) — nó là một mối quan hệ xã hội giữa người giữ bí mật và đối thủ bên ngoài.",
         "Phân định chuẩn xác Public Knowledge vs Common Knowledge theo Halpin và David Lewis; không đồng nhất với bản mã hay bản rõ."),

        ("SLIDE 4: TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ [SƠ ĐỒ 1: TIMELINE]", "1:00",
         "Nội dung trình bày:",
         "Sơ đồ 1 hệ thống hóa tiến trình 5 giai đoạn tiến hóa mô hình đối thủ trong hơn 70 năm qua:\n• Giai đoạn 1 (1948–1949): Mô hình Claude Shannon. Định nghĩa Bí mật hoàn hảo (Perfect Secrecy: H(M|C) = H(M), dẫn đến H(K) >= H(M) qua One-Time Pad). Shannon mô hình hóa hệ thống truyền thông bảo mật điểm-điểm, nhưng việc giản lược đối thủ thành một kẻ nghe lén trừu tượng trên đường truyền đã bỏ qua chi phí phân phối/quản lý khóa và các bề mặt tấn công vật lý thế giới thực.\n• Giai đoạn 2 (1976): Cách mạng khóa công khai của Diffie-Hellman và Merkle, giải quyết bài toán phân phối khóa bằng các bài toán độ phức tạp tính toán một chiều.\n• Giai đoạn 3 (1982–1984): Mô hình PPT Adversary và Trò chơi An ninh (Security Games) của Goldwasser và Micali, hình thức hóa an toàn ngữ nghĩa và chứng minh an toàn bằng quy dẫn toán học đa thức xác suất.\n• Giai đoạn 4 (1981–1993): Dolev-Yao mô hình hóa kẻ thù kiểm soát hoàn toàn kênh mạng logic; Bellare-Rogaway đề xuất Random Oracle Model (ROM) ứng dụng cho hàm băm thực hành.\n• Giai đoạn 5 (2013–2025): Phê phán hậu Snowden, bộc lộ khoảng cách giữa mô hình toán học trên giấy và năng lực thực tế của đối thủ thể chế (kênh kề, backdoor chuẩn NIST, can thiệp chuỗi cung ứng).",
         "Mô tả chính xác mô hình Shannon và giới hạn trừu tượng hóa, không diễn giải sai lệch về nhiễu ngẫu nhiên."),

        ("SLIDE 5: MÔ HÌNH TRUYỀN THÔNG SHANNON vs MẠNG PHỨC HỢP [SƠ ĐỒ 2: DRAW.IO]", "1:00",
         "Nội dung trình bày:",
         "Sơ đồ 2 so sánh trực diện hai mô hình truyền thông:\n• Panel trái: Hệ thống bảo mật điểm-điểm cổ điển của Shannon (1949). Alice gửi bản rõ M mã hóa bằng khóa K qua kênh truyền đến Bob. Đối thủ Eve được mô hình hóa như một kẻ nghe lén thụ động (passive eavesdropper) trên đường truyền. Giới hạn là sự cô lập toán học, bỏ qua chi phí phân phối khóa bí mật và không bao quát các bề mặt tấn công vật lý.\n• Panel phải: Thực tế mạng xã hội - kỹ thuật phức hợp theo Halpin (2025). Hệ thống thông tin luân chuyển giữa Nhóm giữ bí mật (Secret Keepers) và Hạ tầng mở / ISP / Đám mây. Đối thủ thể chế là tác nhân toàn cục cả thụ động lẫn chủ động, tấn công qua 4 con đường thực tế: (1) Kênh kề thời gian/điện năng; (2) Giám sát siêu dữ liệu luồng gói tin; (3) Cài cắm cửa sau chuỗi cung ứng phần cứng; (4) Lỗi phản bội nội bộ Byzantine. Mục tiêu phòng thủ hiện đại vì thế phải chuyển dịch sang bảo vệ siêu dữ liệu và kháng kênh kề.",
         "Thuyết minh bám sát sơ đồ draw.io đã nhúng trên slide, phân tích rõ sự tương phản giữa 2 bên."),

        ("SLIDE 6: BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI (SECTION 4)", "1:00",
         "Nội dung trình bày:",
         "Section 4 của bài báo phân tích ba mô hình toán học hóa kẻ đối đầu nền tảng:\n1. Mô hình PPT Adversary (Goldwasser-Micali 1984): Đối thủ là thuật toán Turing xác suất chạy trong thời gian đa thức. An toàn được định nghĩa qua Trò chơi An ninh (Security Games), trong đó lợi thế đoán đúng của đối thủ phải là một hàm không đáng kể (negligible function). Chứng minh an toàn đạt được bằng phương pháp quy dẫn toán học về bài toán khó.\n2. Mô hình Dolev-Yao (1983): Kẻ thù kiểm soát hoàn toàn kênh mạng (có quyền nghe lén, chặn, sửa, xóa, phát lại thông điệp), nhưng xem mật mã học là 'hộp đen đại số hoàn hảo'. Đối thủ không thể bẻ khóa nếu không có khóa bí mật.\n3. Mô hình Tiên tri Ngẫu nhiên - ROM (Bellare-Rogaway 1993): Giả định hàm băm thực tế hoạt động như một hàm ngẫu nhiên lý tưởng để chứng minh an toàn hiệu năng cao. Tuy nhiên, về mặt triết học, không có hàm băm cụ thể nào (kể cả SHA-256) có thể tương đương toán học với một Random Oracle thực sự.",
         "Định nghĩa chính xác 3 mô hình hình thức toán học cốt lõi."),

        ("SLIDE 7: GIỚI HẠN CỦA HÌNH THỨC HÓA & THẾ GIỚI THỰC [SƠ ĐỒ 3] (SECTION 5)", "1:00",
         "Nội dung trình bày:",
         "Section 5 của Halpin nêu bật các giới hạn thực tế mà các mô hình toán học trừu tượng bỏ sót:\n• Tấn công kênh kề (Side-channel attacks): Paul Kocher (1996) chứng minh việc đo đạc thời gian thực thi, bức xạ điện từ và tiêu thụ điện năng có thể khôi phục khóa mã mà không cần bẻ gãy thuật toán toán học.\n• Rò rỉ qua độ dài bản mã và siêu dữ liệu: Các giao thức như HTTPS, SSH vẫn làm lộ kích thước thông điệp và siêu dữ liệu luồng truyền thông, cho phép đối thủ suy đoán nội dung truyền tải.\n• Phê phán của Ross Anderson (2001): Đại đa số hệ thống an ninh thất bại trong thực tế không phải do thuật toán mật mã bị bẻ gãy, mà do lỗi quản lý khóa, lỗi giao thức, lỗi giao diện người dùng, cấu hình sai và tấn công kỹ nghệ xã hội (Social Engineering).",
         "Không gán Node-Leaf Confusion cho Halpin tại slide này. Tập trung vào Kocher, rò rỉ độ dài và Anderson."),

        ("SLIDE 8: TIẾT LỘ CỦA SNOWDEN & ĐỐI THỦ THỂ CHẾ NSA (SECTION 6)", "1:00",
         "Nội dung trình bày:",
         "Section 6 của Halpin mổ xẻ bước ngoặt lịch sử năm 2013 từ tài liệu Edward Snowden:\n• Phơi bày năng lực thực sự của đối thủ thể chế (NSA và liên minh Five Eyes): Ngân sách khổng lồ, năng lực giám sát cáp quang biển toàn cầu (chương trình PRISM, MUSCULAR).\n• Can thiệp tiêu chuẩn mật mã học: Điển hình là vụ việc cài cửa sau vào bộ sinh số ngẫu nhiên Dual_EC_DRBG được chuẩn hóa bởi NIST, cho phép cơ quan tình báo tính toán trước khóa bí mật.\n• Bài học nhận thức luận: An toàn mật mã không chỉ là bài toán giữa Alice, Bob và Eve trong phòng thí nghiệm; an toàn hệ thống phải chịu đựng được áp lực từ các thực thể có nguồn lực tài chính, chính trị và công nghệ quy mô nhà nước.",
         "Nhấn mạnh ý nghĩa lịch sử của sự kiện Snowden đối với nhận thức luận an ninh."),

        ("SLIDE 9: BỐN CÂU HỎI NGHIÊN CỨU TƯƠNG LAI CỦA HALPIN (RQ1–RQ4)", "1:00",
         "Nội dung trình bày:",
         "Halpin khép lại phần lý thuyết bằng việc đề xuất 4 câu hỏi nghiên cứu mở định hình tương lai:\n• RQ1 (Triết học khoa học): Làm thế nào để xây dựng một phương pháp luận khoa học an ninh hoàn chỉnh, tích hợp cả chứng minh toán học hình thức lẫn kiểm định thực nghiệm vật lý?\n• RQ2 (Bản thể học Kẻ thù): Liệu có thể xây dựng một bản thể học hình thức thống nhất về Đối thủ, bao gồm cả rò rỉ kênh kề, suy thoái phần cứng và lỗi Byzantine nội bộ?\n• RQ3 (Giám sát Siêu dữ liệu & Mixnet): Làm thế nào để thiết kế các kiến trúc định tuyến ẩn danh (như Nym, Tor) bảo vệ siêu dữ liệu trước đối thủ toàn cục kiểm soát hạ tầng mạng?\n• RQ4 (Chính trị học Mật mã): Mật mã học định hình lại cấu trúc quyền lực và tự do dân sự trong kỷ nguyên số như thế nào?",
         "Trình bày cô đọng 4 câu hỏi nghiên cứu lớn của tác giả."),

        ("SLIDE 10: PHÂN TÍCH PHƯƠNG PHÁP LUẬN & GIỚI HẠN TÁI LẬP", "1:00",
         "Nội dung trình bày:",
         "Slide 10 đưa ra tuyên bố phương pháp luận minh bạch của nhóm nghiên cứu đề tài DT005:\n• Đặc thù của bài báo: Bài báo của Harry Halpin (2025) là một công trình khảo cứu nhận thức luận triết học thuần túy (philosophical survey), xuất bản trên tạp chí chuyên ngành liên ngành. Bài báo chứa 0 dòng code, 0 bộ dữ liệu benchmark và 0 kết quả đo đạc định lượng.\n• Nguyên tắc tái lập khoa học: Nhóm không ngụy tạo kết quả của tác giả. Thay vào đó, nhóm thực hiện tái lập theo phương pháp: Trích xuất các khái niệm lý thuyết cốt lõi của bài báo và tự thiết kế độc lập một hệ thống phần mềm thực nghiệm để kiểm chứng và minh họa các khái niệm đó.\n• Bộ dữ liệu và các kịch bản đối kháng là sản phẩm tự xây dựng của sinh viên (được kiểm toán chi tiết trong MIDTERM_CODE_AUDIT.md), đóng vai trò công cụ minh họa trực quan cho các lập luận của Halpin.",
         "Khẳng định trung thực học thuật: Bài báo là lý thuyết, phần mềm là mô hình minh họa độc lập. Xác nhận MIDTERM_CODE_AUDIT.md tồn tại."),

        ("SLIDE 11: MINH HỌA PAPER-TO-CODE & GIỚI HẠN DIỄN GIẢI [SƠ ĐỒ 4]", "1:15",
         "Nội dung trình bày:",
         "Sơ đồ 4 đối chiếu minh bạch ba ví dụ chuyển đổi từ luận điểm bài báo sang mã nguồn thực nghiệm, phân định rõ luận điểm gốc và đóng góp của sinh viên:\n• Ví dụ 1 (Nhiễu ngẫu nhiên vs Kẻ đối đầu có chủ đích): Halpin chỉ rõ mã kiểm tra tuyến tính chỉ được thiết kế để chống lại lỗi ngẫu nhiên. Nhóm sinh viên minh họa độc lập bằng kịch bản tấn công chủ đích: Giải hệ phương trình tuyến tính trên GF(2) để chèn 4 byte bù (patch), làm giả nội dung tệp mà giá trị CRC32 vẫn trùng khớp 100%. Ngược lại, hàm băm SHA-256 phát hiện thay đổi ngay lập tức nhờ tính phi tuyến. Giới hạn diễn giải: Không khẳng định CRC32 hoàn toàn bất lực trong mọi trường hợp (nó vẫn hiệu quả bắt lỗi ngẫu nhiên trong phần cứng/mạng); thí nghiệm chỉ mô tả đúng kịch bản can thiệp có chủ đích đã kiểm thử.\n• Ví dụ 2 (Cây Merkle & Phân tách miền): Halpin không đề xuất kỹ thuật này. Đây là phần minh họa độc lập do DT005 xây dựng, trích dẫn khuyến nghị từ IETF RFC 6962 (Certificate Transparency): Dùng tiền tố 0x00 cho nút lá và 0x01 cho nút nhánh nội bộ để triệt tiêu nguy cơ nhập nhằng nút lá (Node-leaf ambiguity). Giới hạn diễn giải: Kỹ thuật này giải quyết vấn đề cấu trúc cây, không triệt tiêu xác suất va chạm ngẫu nhiên của hàm băm nền tảng.\n• Ví dụ 3 (Tấn công kênh kề thời gian): Halpin dẫn chứng Paul Kocher (1996) về việc thời gian xử lý làm rò rỉ khóa. Nhóm sinh viên minh họa độc lập bằng việc so sánh chuỗi tuần tự có cơ chế ngắt sớm (early return) làm rò rỉ tiền tố, và áp dụng `hmac.compare_digest` để đạt thời gian thực thi hằng số. Giới hạn diễn giải: Phép đo trên Python chịu ảnh hưởng độ trễ lập lịch của hệ điều hành, chỉ minh họa ở mức thư viện phần mềm.",
         "Phân định rạch ròi giữa luận điểm Halpin và mã nguồn DT005; ghi rõ RFC 6962 là minh họa độc lập, không nói Halpin đề xuất; không khẳng định CRC32 bất lực trong mọi trường hợp."),

        ("SLIDE 12: SLIDE CHUYỂN TIẾP: DEMO", "0:15",
         "Nội dung trình bày:",
         "Chuyển sang phần Diễn tập An ninh Thực nghiệm trực tiếp (Live Demo). Nhóm sử dụng mô hình đối kháng hai màn hình giữa Hệ thống phòng thủ Blue Team (Port 8080) và Động cơ tấn công đối kháng Red Team (Port 8081) để trực quan hóa và kiểm chứng thực tế các luận điểm triết học cốt lõi của bài báo Harry Halpin.",
         "Slide chuyển tiếp tối giản nền sáng, chỉ hiển thị đúng một chữ DEMO ở chính giữa để tạo nhịp nghỉ trang trọng trước khi bắt đầu phần thực hành."),

        ("SLIDE 13: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG RED TEAM VS BLUE TEAM [SƠ ĐỒ 5: DRAW.IO]", "1:00",
         "Nội dung trình bày:",
         "Sơ đồ 5 mô tả kiến trúc diễn tập đối kháng (Demo Architecture):\n• Bộ dữ liệu thực nghiệm: Đối chiếu số liệu kỹ thuật, báo cáo benchmark trong mã nguồn ghi nhận tập dữ liệu gồm 2.485 tệp tin với dung lượng khoảng 112.81 MiB (phiên bản nạp kiểm thử cục bộ hiện tại gồm 2.483 tệp).\n• Blue Team (Port 8080 - Hệ thống phòng thủ & giám sát): Nạp tập dữ liệu để thiết lập Mốc tin cậy ban đầu (Root Hash H_0). Khi hệ thống kiểm tra toàn vẹn, cơ chế duyệt cây phân cấp định vị chính xác vị trí tệp tin bị biến đổi với tốc độ quét thực tế vượt 2.000 tệp/giây.\n• Red Team (Port 8081 - Động cơ tấn công đối kháng): Mô phỏng các vector can thiệp thực tế: Sửa byte tệp ngẫu nhiên (In-place Tamper), bù 4 byte GF(2) qua mặt CRC32, đánh tráo Merkle Proof trên đường truyền Dolev-Yao, và chèn backdoor/xóa tệp.\n• Cơ chế Rollback an toàn: Thư mục sao lưu .adversary_backups/ cho phép phục hồi nguyên trạng dữ liệu gốc sau mỗi kịch bản diễn tập.",
         "Tiêu đề chuẩn Demo Architecture, số liệu đối chiếu chuẩn xác với benchmark_results.json, không đưa số liệu chưa xác minh."),

        ("SLIDE 14: TỔNG KẾT BÁO CÁO GIỮA KỲ & LỘ TRÌNH CUỐI KỲ", "0:45",
         "Nội dung trình bày:",
         "Tổng kết báo cáo giữa kỳ đề tài DT005:\n1. Về mặt học thuật: Hoàn thành thẩm định khoa học bài báo của Halpin (2025), phân tích đúng đắn cơ sở triết học và các mô hình kẻ đối đầu qua các thời kỳ lịch sử.\n2. Về mặt thực nghiệm: Xây dựng thành công bộ công cụ diễn tập đối kháng Red Team vs Blue Team minh họa trực quan các luận điểm an ninh chính của bài báo.\n3. Về định hướng nghiên cứu cuối kỳ: Nhóm tiếp tục nghiên cứu cấu trúc dữ liệu xác thực phân cấp (authenticated data structures), cơ chế bộ đệm cập nhật cục bộ nhằm giảm chi phí tính toán lại, hoàn thiện chuẩn hóa phân tách miền và mở rộng kiểm định đối kháng trên các quy mô dữ liệu lớn hơn.",
         "Đã loại bỏ các tuyên bố chưa kiểm chứng như O(1), ~0.24 ms khỏi kết luận giữa kỳ; trình bày định hướng cuối kỳ mang tính học thuật khái quát."),

        ("SLIDE 15: CẢM ƠN & PHIÊN PHẢN BIỆN (Q&A)", "0:15",
         "Nội dung trình bày:",
         "Phần trình bày báo cáo giữa kỳ của đề tài DT005 đến đây là kết thúc. Xin trân trọng cảm ơn quý vị đã chú ý theo dõi. Nhóm sẵn sàng trao đổi và giải đáp các câu hỏi phản biện.",
         "Kết thúc tự tin, sẵn sàng chuyển sang phần Demo hoặc trả lời câu hỏi.")
    ]

    for title, duration, section_label, speech, note in slides_content:
        p_s = doc.add_paragraph()
        p_s.paragraph_format.space_before = Pt(10)
        p_s.paragraph_format.space_after = Pt(2)
        r_title = p_s.add_run(f"► {title} ")
        r_title.font.size = Pt(12)
        r_title.font.bold = True
        r_title.font.color.rgb = RGBColor(30, 64, 175)

        r_dur = p_s.add_run(f"[{duration}]")
        r_dur.font.size = Pt(10)
        r_dur.font.bold = True
        r_dur.font.color.rgb = RGBColor(100, 116, 139)

        p_speech = doc.add_paragraph()
        p_speech.paragraph_format.space_before = Pt(2)
        p_speech.paragraph_format.space_after = Pt(4)
        p_speech.paragraph_format.left_indent = Inches(0.2)
        r_lbl = p_speech.add_run(f"{section_label}\n")
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(10.5)
        r_lbl.font.color.rgb = RGBColor(15, 23, 42)

        r_sp = p_speech.add_run(f'"{speech}"')
        r_sp.font.size = Pt(10.5)
        r_sp.font.italic = False
        r_sp.font.color.rgb = RGBColor(51, 65, 85)

        p_note = doc.add_paragraph()
        p_note.paragraph_format.space_before = Pt(2)
        p_note.paragraph_format.space_after = Pt(8)
        p_note.paragraph_format.left_indent = Inches(0.2)
        r_nt = p_note.add_run(f"• Lưu ý trình bày: {note}")
        r_nt.font.size = Pt(9.5)
        r_nt.font.italic = True
        r_nt.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_page_break()

    # =========================================================================
    # PHẦN 2: HƯỚNG DẪN THAO TÁC DIỄN TẬP DEMO THỰC TẾ
    # =========================================================================
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("PHẦN 2: HƯỚNG DẪN THAO TÁC DIỄN TẬP DEMO THỰC TẾ")
    r2.font.size = Pt(14)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(30, 64, 175)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(10)

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
    r_op1.font.size = Pt(12)
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
    r_op2.font.size = Pt(12)
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
    r3.font.size = Pt(14)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(30, 64, 175)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(10)

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
