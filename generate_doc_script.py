# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word (.docx) Kịch Bản Thuyết Trình & Hướng Dẫn Diễn Tập Demo:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn:
- 100% bám sát cấu trúc 15 Slide hoàn chỉnh của Bao_Cao_Giua_Ky_DT05_Final.pptx.
- Không dùng xưng hô rườm rà (loại bỏ toàn bộ "Thưa Thầy", "Trả lời Thầy", "Em xin kính chào...").
- Ngôn ngữ học thuật, khách quan, súc tích, mạch lạc.
- Định dạng Word chuyên nghiệp (Tiêu đề, Heading rõ ràng, bảng biểu đối sánh, box ghi chú).
"""

import sys
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
    sections = doc.sections
    for section in sections:
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

    run_sub = p_title.add_run("Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động\n")
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
         "Báo cáo tiến độ đồ án giữa kỳ đề tài DT005: Xây dựng hệ thống kiểm tra toàn vẹn tệp dựa trên SHA-256, HMAC và Cây Merkle phân cấp động; kết hợp thẩm định học thuật bài báo 'The adversary: the philosophy of cryptography' của Tiến sĩ Harry Halpin, xuất bản năm 2025 trên tạp chí Journal of Cybersecurity thuộc Đại học Oxford.\nNội dung báo cáo gồm 3 trọng tâm: (1) Khảo cứu các luận điểm nhận thức luận và tiến trình hình thức hóa kẻ đối đầu trong mật mã học; (2) Phân định phương pháp luận giữa khảo sát triết học lý thuyết của tác giả và hệ thống thực nghiệm độc lập do nhóm tự phát triển; (3) Diễn tập thực tế mô hình đối kháng an ninh Red Team vs Blue Team trên tập dữ liệu thực tế 2.483 tệp tin.",
         "Điểm nhấn trình bày: Bố cục rõ ràng, nhận diện trang trọng, xác định chuẩn mục tiêu đồ án giữa kỳ."),

        ("SLIDE 2: BỐI CẢNH, MỤC TIÊU & BA ĐÓNG GÓP CỦA BÀI BÁO (SECTION 1)", "1:00",
         "Nội dung trình bày:",
         "Mật mã học trong kỷ nguyên số là khoa học về việc duy trì liên lạc bảo mật khi có sự hiện diện của đối thủ có chủ đích trên môi trường mạng mở.\nBài báo của Halpin chỉ ra một nghịch lý nhận thức luận sâu sắc: Trong khi lý thuyết thông tin của Claude Shannon được triết học khảo cứu rộng rãi suốt nhiều thập kỷ, thì chính mật mã học — cái nôi khai sinh ra khái niệm thông tin và tính ngẫu nhiên — lại bị giới triết học bỏ quên.\nBài báo đưa ra ba đóng góp học thuật cốt lõi: Thứ nhất, định hình lại lịch sử mật mã qua lăng kính triết học bản thể luận về 'Kẻ thù' (The Adversary). Thứ hai, phê phán sự trừu tượng hóa toán học thái quá khi biến đối thủ thành các hàm lý tưởng và bỏ qua các vector tấn công thế giới thực (kênh kề, siêu dữ liệu, lỗi phần cứng). Thứ ba, đề xuất một chương trình nghiên cứu liên ngành mở với 4 câu hỏi định hướng cho khoa học an ninh tương lai.",
         "Điểm nhấn trình bày: Nêu bật nghịch lý triết học và 3 đóng góp lớn của Halpin."),

        ("SLIDE 3: CƠ SỞ TRIẾT HỌC VỀ THÔNG TIN & BẢN THỂ HỌC (SECTION 2)", "1:00",
         "Nội dung trình bày:",
         "Nền tảng lý thuyết của Halpin dựa trên 2 trụ cột triết học nhận thức luận:\n1. Bản thể học thông tin của Brian Cantwell Smith (1996): Thực tại khởi nguồn từ một dòng chảy siêu hình nguyên thủy không biên giới. Con người và hệ thống định hình thế giới bằng cách 'đăng ký đối tượng' (registration of objects). Trong mật mã học, các đối tượng như thông điệp, bản mã và khóa được đăng ký nhằm kiểm soát dòng thông tin giữa các tác nhân ở xa nhau (distal agents).\n2. Thuyết thông tin khách quan của Fred Dretske (1981): Thông tin là một trạng thái khách quan gắn liền với quy luật tự nhiên, độc lập với người quan sát. Bản rõ chứa đựng thông tin về nguồn phát; hệ mật mã hoàn hảo đảm bảo bản mã không để rò rỉ bất kỳ thông tin nào về bản rõ nếu không có khóa.\nKhác với nhiễu ngẫu nhiên trong vật lý, Kẻ đối đầu mật mã là một tác nhân có chủ đích (intentional agent), có ý chí và năng lực khai thác lỗ hổng cấu trúc hệ thống.",
         "Điểm nhấn trình bày: Phân tích sự chuyển dịch từ đối tượng vật lý sang bản thể thông tin mật mã."),

        ("SLIDE 4: TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ [SƠ ĐỒ 1: TIMELINE]", "1:00",
         "Nội dung trình bày:",
         "Sơ đồ 1 hệ thống hóa tiến trình 5 giai đoạn tiến hóa mô hình đối thủ trong hơn 70 năm qua:\n• Giai đoạn 1 (1948–1949): Mô hình Shannon, định nghĩa Bí mật hoàn hảo (One-Time Pad) dựa trên lý thuyết xác suất và giả định sai lệch bắt nguồn từ nhiễu ngẫu nhiên.\n• Giai đoạn 2 (1976): Cách mạng khóa công khai của Diffie-Hellman và Merkle, giải quyết bài toán phân phối khóa bằng các bài toán độ phức tạp tính toán một chiều.\n• Giai đoạn 3 (1982–1984): Mô hình PPT Adversary và Trò chơi An ninh (Security Games) của Goldwasser và Micali, hình thức hóa an toàn ngữ nghĩa và chứng minh an toàn bằng quy dẫn toán học đa thức xác suất.\n• Giai đoạn 4 (1981–1993): Dolev-Yao mô hình hóa kẻ thù kiểm soát hoàn toàn kênh mạng logic; Bellare-Rogaway đề xuất Random Oracle Model (ROM) ứng dụng cho hàm băm thực hành.\n• Giai đoạn 5 (2013–2025): Phê phán hậu Snowden, bộc lộ khoảng cách giữa mô hình toán học trên giấy và năng lực thực tế của đối thủ thể chế (kênh kề, backdoor chuẩn NIST, can thiệp chuỗi cung ứng).",
         "Điểm nhấn trình bày: Trình bày mạch lạc 5 mốc lịch sử từ Shannon đến Snowden."),

        ("SLIDE 5: MÔ HÌNH TRUYỀN THÔNG SHANNON vs MẠNG PHỨC HỢP [SƠ ĐỒ 2: DRAW.IO]", "1:00",
         "Nội dung trình bày:",
         "Sơ đồ 2 so sánh trực diện hai mô hình truyền thông:\n• Panel trái: Mô hình điểm-điểm cổ điển của Shannon (1949). Alice gửi bản rõ M mã hóa bằng khóa K qua kênh truyền đến Bob. Đối thủ Eve chỉ là kẻ nghe lén thụ động (passive eavesdropper) trên đường truyền. Điểm mù của Shannon là giả định sai lệch do nhiễu ngẫu nhiên và bỏ qua chi phí lưu trữ, phân phối khóa cũng như rò rỉ vật lý.\n• Panel phải: Thực tế mạng xã hội - kỹ thuật phức hợp theo Halpin (2025). Hệ thống không còn là đường truyền đóng kín đơn thuần, mà thông tin luân chuyển giữa Nhóm giữ bí mật (Secret Keepers) và Hạ tầng mở/ISP/Đám mây. Đối thủ thể chế (NSA, Five Eyes) là tác nhân toàn cục cả thụ động lẫn chủ động, tấn công qua 4 con đường thực tế: (1) Kênh kề thời gian/điện năng; (2) Giám sát siêu dữ liệu luồng gói tin; (3) Cài cắm cửa sau chuỗi cung ứng phần cứng; (4) Lỗi phản bội nội bộ Byzantine. Mục tiêu phòng thủ hiện đại vì thế phải chuyển dịch sang bảo vệ siêu dữ liệu và kháng kênh kề.",
         "Điểm nhấn trình bày: Thuyết minh bám sát sơ đồ draw.io đã nhúng trên slide, phân tích rõ sự tương phản giữa 2 bên."),

        ("SLIDE 6: BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI (SECTION 4)", "1:00",
         "Nội dung trình bày:",
         "Section 4 của bài báo phân tích ba mô hình toán học hóa kẻ đối đầu nền tảng:\n1. Mô hình PPT Adversary (Goldwasser-Micali 1984): Đối thủ là thuật toán Turing xác suất chạy trong thời gian đa thức. An toàn được định nghĩa qua Trò chơi An ninh (Security Games), trong đó lợi thế đoán đúng của đối thủ phải là một hàm không đáng kể (negligible function). Chứng minh an toàn đạt được bằng phương pháp quy dẫn toán học về bài toán khó.\n2. Mô hình Dolev-Yao (1983): Kẻ thù kiểm soát hoàn toàn kênh mạng (có quyền nghe lén, chặn, sửa, xóa, phát lại thông điệp), nhưng xem mật mã học là 'hộp đen đại số hoàn hảo'. Đối thủ không thể bẻ khóa nếu không có khóa bí mật.\n3. Mô hình Tiên tri Ngẫu nhiên - ROM (Bellare-Rogaway 1993): Giả định hàm băm thực tế hoạt động như một hàm ngẫu nhiên lý tưởng để chứng minh an toàn hiệu năng cao. Tuy nhiên, về mặt triết học, không có hàm băm cụ thể nào (kể cả SHA-256) có thể tương đương toán học với một Random Oracle thực sự.",
         "Điểm nhấn trình bày: Định nghĩa chính xác 3 mô hình hình thức toán học cốt lõi."),

        ("SLIDE 7: GIỚI HẠN CỦA HÌNH THỨC HÓA & THẾ GIỚI THỰC [SƠ ĐỒ 3] (SECTION 5)", "1:00",
         "Nội dung trình bày:",
         "Section 5 nêu bật các giới hạn thực tế mà các mô hình toán học trừu tượng bỏ sót:\n• Tấn công kênh kề (Side-channel attacks): Paul Kocher (1996) chứng minh việc đo đạc thời gian thực thi, bức xạ điện từ và tiêu thụ điện năng có thể khôi phục khóa mã mà không cần bẻ gãy thuật toán toán học.\n• Rò rỉ qua độ dài bản mã: Các giao thức như HTTPS, SSH vẫn làm lộ kích thước thông điệp và siêu dữ liệu luồng truyền thông, cho phép đối thủ suy đoán loại nội dung truyền tải.\n• Phê phán của Ross Anderson (2001): Đại đa số hệ thống an ninh thất bại trong thực tế không phải do thuật toán mật mã bị bẻ gãy, mà do lỗi quản lý khóa, lỗi giao diện người dùng, cấu hình sai và tấn công kỹ nghệ xã hội (Social Engineering).",
         "Điểm nhấn trình bày: Làm rõ khoảng cách giữa bài toán trên giấy và thực tế triển khai hệ thống."),

        ("SLIDE 8: TIẾT LỘ CỦA SNOWDEN & ĐỐI THỦ THỂ CHẾ NSA (SECTION 6)", "1:00",
         "Nội dung trình bày:",
         "Section 6 của Halpin mổ xẻ bước ngoặt lịch sử năm 2013 từ tài liệu Edward Snowden:\n• Phơi bày năng lực thực sự của đối thủ thể chế (NSA và liên minh Five Eyes): Ngân sách khổng lồ, năng lực giám sát cáp quang biển toàn cầu (chương trình PRISM, MUSCULAR).\n• Can thiệp tiêu chuẩn mật mã học: Điển hình là vụ việc cài cửa sau vào bộ sinh số ngẫu nhiên Dual_EC_DRBG được chuẩn hóa bởi NIST, cho phép cơ quan tình báo tính toán trước khóa bí mật.\n• Bài học nhận thức luận: An toàn mật mã không chỉ là bài toán giữa Alice, Bob và Eve trong phòng thí nghiệm; an toàn hệ thống phải chịu đựng được áp lực từ các thực thể có nguồn lực tài chính, chính trị và công nghệ quy mô nhà nước.",
         "Điểm nhấn trình bày: Nhấn mạnh ý nghĩa lịch sử của sự kiện Snowden đối với nhận thức luận an ninh."),

        ("SLIDE 9: BỐN CÂU HỎI NGHIÊN CỨU TƯƠNG LAI CỦA HALPIN (RQ1–RQ4)", "1:00",
         "Nội dung trình bày:",
         "Halpin khép lại phần lý thuyết bằng việc đề xuất 4 câu hỏi nghiên cứu mở định hình tương lai:\n• RQ1 (Triết học khoa học): Làm thế nào để xây dựng một phương pháp luận khoa học an ninh hoàn chỉnh, tích hợp cả chứng minh toán học hình thức lẫn kiểm định thực nghiệm vật lý?\n• RQ2 (Bản thể học Kẻ thù): Liệu có thể xây dựng một bản thể học hình thức thống nhất về Đối thủ, bao gồm cả rò rỉ kênh kề, suy thoái phần cứng và lỗi Byzantine nội bộ?\n• RQ3 (Giám sát Siêu dữ liệu & Mixnet): Làm thế nào để thiết kế các kiến trúc định tuyến ẩn danh (như Nym, Tor) bảo vệ siêu dữ liệu trước đối thủ toàn cục kiểm soát hạ tầng mạng?\n• RQ4 (Chính trị học Mật mã): Mật mã học định hình lại cấu trúc quyền lực và tự do dân sự trong kỷ nguyên số như thế nào?",
         "Điểm nhấn trình bày: Trình bày cô đọng 4 câu hỏi nghiên cứu lớn của tác giả."),

        ("SLIDE 10: PHÂN TÍCH PHƯƠNG PHÁP LUẬN & GIỚI HẠN TÁI LẬP", "1:00",
         "Nội dung trình bày:",
         "Slide 10 đưa ra tuyên bố phương pháp luận minh bạch của nhóm nghiên cứu đề tài DT005:\n• Đặc thù của bài báo: Bài báo của Harry Halpin (2025) là một công trình khảo cứu nhận thức luận triết học thuần túy (philosophical survey), xuất bản trên tạp chí chuyên ngành liên ngành. Bài báo chứa 0 dòng code, 0 bộ dữ liệu benchmark và 0 kết quả đo đạc định lượng.\n• Nguyên tắc tái lập khoa học: Nhóm không ngụy tạo kết quả của tác giả. Thay vào đó, nhóm thực hiện tái lập theo phương pháp: Trích xuất các khái niệm lý thuyết cốt lõi của bài báo và tự thiết kế độc lập một hệ thống phần mềm thực nghiệm để kiểm chứng và minh họa các khái niệm đó.\n• Bộ dữ liệu Kaggle Resume Dataset (2.483 tệp, 112 MB) và bộ công cụ đo đạc đối kháng là sản phẩm tự phát triển của sinh viên, đóng vai trò công cụ minh họa trực quan cho các lập luận của Halpin.",
         "Điểm nhấn trình bày: Khẳng định trung thực học thuật: Bài báo là lý thuyết, phần mềm là mô hình minh họa độc lập."),

        ("SLIDE 11: MINH HỌA PAPER-TO-CODE & GIỚI HẠN DIỄN GIẢI [SƠ ĐỒ 4]", "1:15",
         "Nội dung trình bày:",
         "Sơ đồ 4 đối chiếu minh bạch ba ví dụ chuyển đổi từ luận điểm bài báo sang mã nguồn thực nghiệm:\n• Ví dụ 1 (Nhiễu vs Kẻ thù chủ đích): Halpin chỉ rõ mã kiểm tra lỗi CRC32 chỉ chống được nhiễu ngẫu nhiên. Nhóm triển khai thuật toán giải hệ phương trình bù trên trường nhị phân GF(2), tìm 4 byte patch làm giả tệp tin mà CRC32 vẫn trùng khớp 100%. Ngược lại, hàm băm SHA-256 phát hiện thay đổi ngay lập tức. Giới hạn: Thí nghiệm chỉ minh họa tính phi tuyến của hàm băm, không đại diện cho an toàn toàn diện của toàn hệ thống.\n• Ví dụ 2 (Nhập nhằng nút lá trong cây Merkle): Phép ghép thô H(L || R) có nguy cơ làm lẫn lộn nút nhánh nội bộ với nút lá. Nhóm cài đặt chuẩn phân tách miền RFC 6962: Tiền tố 0x00 cho lá và 0x01 cho nhánh, triệt tiêu hoàn toàn nhập nhằng cấu trúc. Giới hạn: Chỉ giải quyết va chạm cấu trúc tầng cây, không loại trừ va chạm ngẫu nhiên 2^256 của SHA-256.\n• Ví dụ 3 (Tấn công kênh kề thời gian): Dừng sớm khi so sánh chuỗi làm rò rỉ dữ liệu qua thời gian thực thi. Nhóm đo đạc so sánh tuần tự và chứng minh hàm hmac.compare_digest triệt tiêu độ lệch thời gian. Giới hạn: Phép đo trên Python có độ trễ hệ điều hành, chỉ minh họa ở mức thư viện phần mềm.",
         "Điểm nhấn trình bày: Bám sát 3 cột trên slide: Luận điểm bài báo, Mã nguồn DT005, Giới hạn diễn giải."),

        ("SLIDE 12: SLIDE CHUYỂN TIẾP: DEMO TRỰC TIẾP", "0:15",
         "Nội dung trình bày:",
         "Sau phần phân tích lý thuyết và đối chiếu phương pháp luận, nhóm chuyển sang phần Diễn tập An ninh Thực nghiệm trực tiếp: Mô hình đối kháng hai màn hình giữa Hệ thống phòng thủ Blue Team (Port 8080) và Động cơ tấn công đối kháng Red Team (Port 8081) trên tập dữ liệu thực tế 2.483 tệp tin.",
         "Điểm nhấn trình bày: Chuyển tiếp dứt khoát, giới thiệu các cổng demo trực quan."),

        ("SLIDE 13: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG RED TEAM VS BLUE TEAM [SƠ ĐỒ 5: DRAW.IO]", "1:00",
         "Nội dung trình bày:",
         "Sơ đồ 5 mô tả kiến trúc tương tác thời gian thực giữa hai bên:\n• Blue Team (Port 8080 - Hệ thống phòng thủ & giám sát): Nạp toàn bộ tập dữ liệu 2.483 tệp (112 MB) để thiết lập Mốc tin cậy ban đầu (Root Hash H_0 sạch). Khi phát hiện can thiệp, cơ chế Dirty Path O(log N) định vị chính xác vị trí tệp tin bị biến đổi trong ~1.06 ms (tốc độ quét đạt 2.325 tệp/giây).\n• Red Team (Port 8081 - Động cơ tấn công đối kháng Adversary Engine): Mô phỏng các vector can thiệp thực tế: Sửa byte tệp ngẫu nhiên (In-place Tamper), bù 4 byte GF(2) qua mặt CRC32, đánh tráo Merkle Proof trên đường truyền Dolev-Yao, và chèn backdoor/xóa tệp.\n• Cơ chế Rollback an toàn: Thư mục sao lưu .adversary_backups/ cho phép phục hồi nguyên trạng 100% dữ liệu gốc chỉ với 1 cú click chuột sau mỗi đợt diễn tập.",
         "Điểm nhấn trình bày: Thuyết minh theo sơ đồ draw.io kiến trúc Red vs Blue, nhấn mạnh cơ chế phát hiện và rollback an toàn."),

        ("SLIDE 14: TỔNG KẾT BÁO CÁO GIỮA KỲ & LỘ TRÌNH CUỐI KỲ", "0:45",
         "Nội dung trình bày:",
         "Tổng kết báo cáo giữa kỳ đề tài DT005:\n1. Về mặt học thuật: Hoàn thành thẩm định khoa học bài báo của Halpin (2025), phân tích đúng đắn cơ sở triết học và các mô hình kẻ đối đầu qua các thời kỳ lịch sử.\n2. Về mặt thực nghiệm: Xây dựng thành công bộ công cụ diễn tập đối kháng Red Team vs Blue Team minh họa sinh động 3 luận điểm kỹ thuật chính của bài báo.\n3. Đề xuất lộ trình nghiên cứu cuối kỳ: Nhóm tiếp tục hoàn thiện và tối ưu hóa giải pháp thuật toán Cây Merkle phân cấp động (HD-MerkleTree): Cơ chế cập nhật cục bộ Dirty-Path Cache đạt độ phức tạp O(log N) (~0.24 ms/tệp), chuẩn phân tách miền RFC 6962, tối ưu hóa kích thước bằng chứng CompactProof (256 bytes) và thực hiện kiểm thử độ bền (Ablation Study) toàn diện.",
         "Điểm nhấn trình bày: Tóm tắt 3 kết quả đạt được giữa kỳ và lộ trình kỹ thuật cuối kỳ rõ ràng."),

        ("SLIDE 15: CẢM ƠN & PHIÊN PHẢN BIỆN (Q&A)", "0:15",
         "Nội dung trình bày:",
         "Phần trình bày báo cáo giữa kỳ của đề tài DT005 đến đây là kết thúc. Xin trân trọng cảm ơn quý vị đã chú ý theo dõi. Nhóm sẵn sàng trao đổi và giải đáp các câu hỏi phản biện.",
         "Điểm nhấn trình bày: Kết thúc tự tin, sẵn sàng chuyển sang phần Demo hoặc trả lời câu hỏi.")
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
    p_intro_demo.add_run("Hệ thống demo hỗ trợ 2 hình thức diễn tập tùy thuộc vào yêu cầu của buổi báo cáo: Giao diện Web Trực quan (Khuyến nghị ưu tiên) và Dòng lệnh Terminal CLI độc lập.")

    # PHƯƠNG ÁN 1: GIAO DIỆN WEB TRỰC QUAN
    p_opt1 = doc.add_paragraph()
    p_opt1.paragraph_format.space_before = Pt(8)
    r_op1 = p_opt1.add_run("Phương án 1: Diễn tập Đối kháng 2 Màn hình trên Trình duyệt Web (Ưu tiên hàng đầu)")
    r_op1.font.bold = True
    r_op1.font.size = Pt(12)
    r_op1.font.color.rgb = RGBColor(15, 23, 42)

    demo_steps = [
        ("Bước 1: Khởi động 2 máy chủ Web trong nền",
         "Mở terminal và kích hoạt song song 2 dịch vụ:\n"
         "• Blue Team Service:  python -m src.demo.web_demo --port 8080\n"
         "• Red Team Service:   python -m src.demo.red_team_web --port 8081\n"
         "Mở 2 cửa sổ trình duyệt đặt cạnh nhau (Split Screen): Cửa sổ trái mở http://localhost:8080; cửa sổ phải mở http://localhost:8081."),

        ("Bước 2: Thiết lập mốc tin cậy sạch tại Blue Team Dashboard (Port 8080)",
         "Tại cổng 8080, chọn thư mục 'kaggle_dataset' (chứa 2.483 tệp tin, dung lượng 112.81 MB).\n"
         "Bấm nút 'Nạp & Tính toán Toàn vẹn'. Hệ thống tính toán và hiển thị Root Hash ban đầu H_0 sạch với trạng thái xanh: '100% NGUYÊN VẸN'."),

        ("Bước 3: Phát động tấn công tại Red Team Adversary Engine (Port 8081)",
         "Tại cổng 8081, thực hiện lần lượt các kịch bản tấn công đối kháng:\n"
         "• Kịch bản 1 (In-place Tamper): Bấm nút 'Can thiệp tệp tin ngẫu nhiên'. Red Team tự động chọn một tệp (ví dụ: Resume/Resume.csv) và sửa đổi nội dung trực tiếp trên ổ đĩa.\n"
         "• Kịch bản 2 (CRC32 Forge): Bấm nút 'Tấn công bù 4 byte GF(2)'. Hệ thống giải hệ phương trình tuyến tính, làm sai lệch tệp mà mã CRC32 vẫn giữ nguyên 100%.\n"
         "• Kịch bản 3 (Dolev-Yao MitM): Bấm nút 'Đánh tráo Merkle Proof trên mạng'. Mô phỏng kẻ thù chặn và thay thế chuỗi chứng minh toàn vẹn."),

        ("Bước 4: Kiểm tra và bắt quả tang tại Blue Team Dashboard (Port 8080)",
         "Quay trở lại cổng 8080, bấm nút 'Kiểm tra toàn vẹn'.\n"
         "• Hệ thống đổi sang trạng thái cảnh báo đỏ: 'CẢNH BÁO: PHÁT HIỆN CAN THIỆP DỮ LIỆU!'.\n"
         "• Cột Dirty Path O(log N) chỉ ra chính xác đường dẫn nhánh cây bị sai lệch chỉ trong ~1.06 ms.\n"
         "• Bảng đối chứng hiển thị: Thuật toán CRC32 bị qua mặt hoàn toàn (Báo khớp sai sự thật), trong khi SHA-256 và Merkle Tree phát hiện tức thì."),

        ("Bước 5: Hoàn nguyên dữ liệu sạch 100% (Rollback All)",
         "Tại giao diện Red Team (Port 8081), bấm nút 'ROLLBACK ALL'. Toàn bộ các tệp tin bị tác động trên đĩa được khôi phục nguyên vẹn từ thư mục .adversary_backups/.\n"
         "Tại Blue Team bấm kiểm tra lại: Hệ thống trở về trạng thái xanh an toàn 100%.")
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

    # PHƯƠNG ÁN 2: DÒNG LỆNH CLI
    p_opt2 = doc.add_paragraph()
    p_opt2.paragraph_format.space_before = Pt(10)
    r_op2 = p_opt2.add_run("Phương án 2: Diễn tập Dòng lệnh Terminal CLI (Dự phòng trường hợp không dùng UI)")
    r_op2.font.bold = True
    r_op2.font.size = Pt(12)
    r_op2.font.color.rgb = RGBColor(15, 23, 42)

    cli_text = (
        "1. Chạy Suite 7 Kiểm định Toán học Tự động:\n"
        "   Lệnh: python -m src.proposed_model.security_games\n"
        "   Kết quả: Thực hiện toàn bộ 7 phép đo bám sát bài báo (tính phi tuyến, hiệu ứng tuyết lở SAC ~50%, bẻ gãy CRC32 trên GF(2), phân tách miền RFC 6962, và kênh kề thời gian so sánh chuỗi). Toàn bộ 7 kiểm định hoàn thành đạt chuẩn trong ~2.885 ms.\n\n"
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
         "3. Mô hình hóa trọn vẹn các đối thủ trong bài báo: Hệ thống toàn vẹn tệp tin cho phép mô hình hóa và kiểm chứng đầy đủ các dạng đối thủ được Halpin mổ xẻ: Từ kẻ nghe lén của Shannon, kẻ can thiệp mạng Dolev-Yao, kẻ khai thác kênh kề Kocher, đến các đòn tấn công can thiệp dữ liệu và chèn backdoor chuỗi cung ứng."),

        ("Câu hỏi 2: Bài báo có 134 tài liệu tham khảo, khi thẩm định và tái lập thì có bắt buộc phải làm theo toàn bộ 134 tài liệu đó không?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Bản chất bài khảo cứu: Bài báo của Halpin là một bài tổng quan triết học (philosophical survey) dài 17 trang, 134 trích dẫn là kho tàng tài liệu lịch sử trải dài từ năm 1883 đến 2024, bao gồm cả sách triết học siêu hình (Dretske, Cantwell Smith), tài liệu lưu trữ chính phủ về NSA, và các hệ thống chứng minh định lý tự động (như Isabelle chứng minh giao thức TLS của Paulson). Về mặt phương pháp luận nghiên cứu, không một công trình kỹ thuật nào có thể hay bắt buộc phải tái lập toàn bộ 134 tài liệu nền đó.\n"
         "2. Phương pháp tinh lọc trọng tâm: Thẩm định một bài báo lý thuyết là tinh lọc các trụ cột kỹ thuật trực tiếp liên quan đến bài toán an ninh. Đề tài DT005 đã tinh lọc và hiện thực hóa thành công 8 tài liệu kỹ thuật cốt lõi: Ref [2, 3] của Shannon (Mã tuyến tính vs Hàm băm), Ref [21] của Biba (Toàn vẹn phân cấp), Ref [33, 34] của Goldwasser-Micali (Trò chơi an ninh PPT), Ref [36] của Kocher (Kênh kề thời gian), Ref [40] của Bellare-Rogaway (Random Oracle), Ref [52] của Dolev-Yao (Mạng thù địch), Ref [100, 101] về Snowden/NSA (Đối thủ thể chế), và Ref [109] của Ben Laurie (RFC 6962 Certificate Transparency)."),

        ("Câu hỏi 3: Tập dữ liệu kaggle_dataset (2.483 tệp tin, 112 MB) từ đâu ra? Có xuất hiện trong bài báo gốc không?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Khẳng định minh bạch khoa học: Bài báo gốc của Harry Halpin hoàn toàn KHÔNG CÓ bất kỳ tập dữ liệu, bảng số liệu thực nghiệm hay mã nguồn nào. Bài báo là công trình khảo luận nhận thức luận trên tạp chí lý thuyết.\n"
         "2. Mục đích đưa tập dữ liệu vào nghiên cứu: Tập dữ liệu Kaggle Resume Dataset (gồm 2.483 tệp tin PDF và CSV thuộc 24 ngành nghề) do nhóm sinh viên tự chọn và tiền xử lý. Lý do chọn tập dữ liệu này là để cung cấp một môi trường cây thư mục doanh nghiệp nhiều tầng thực tế, kích thước tệp biến thiên, phản ánh trung thực tải làm việc thực tế nhằm đo đạc năng lực định vị sai lệch O(log N) và khả năng chịu tải đối kháng, hoàn toàn không phải ngụy tạo số liệu của tác giả."),

        ("Câu hỏi 4: Tại sao mã kiểm tra lỗi tuyến tính CRC32 bị bẻ gãy 100% trên trường nhị phân GF(2), trong khi SHA-256 lại phát hiện được gian lận?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Tính chất tuyến tính của CRC32: CRC32 hoạt động dựa trên phép chia đa thức trên trường nhị phân hữu hạn GF(2). Do tính chất tuyến tính, giá trị CRC của một thông điệp sửa đổi CRC(M XOR Delta) có thể tính toán trực tiếp từ CRC(M) và CRC(Delta). Kẻ đối đầu có chủ đích chỉ cần giải một hệ phương trình đại số tuyến tính bằng thuật toán khử Gauss để tìm ra 4 byte bù (patch) nối vào cuối tệp tin, khiến cho CRC32 của tệp giả mạo trùng khớp 100% với tệp gốc ban đầu.\n"
         "2. Tính phi tuyến của SHA-256: SHA-256 là hàm băm mật mã học xây dựng trên cấu trúc lặp Merkle-Damgard và phép nén Davies-Meyer với các hàm phi tuyến (Ch, Maj, Sigma). Hàm băm đạt hiệu ứng tuyết lở (Strict Avalanche Criterion - SAC ~50%): Chỉ cần thay đổi 1 bit đầu vào thì trung bình 50% số bit đầu ra bị đảo ngẫu nhiên. Bài toán tìm tiền ảnh thứ hai trên SHA-256 có độ phức tạp 2^256, vượt quá hoàn toàn năng lực tính toán của đối thủ đa thức xác suất (PPT Adversary)."),

        ("Câu hỏi 5: Tấn công Kênh kề thời gian (Timing Side-Channel) trong so sánh chuỗi hoạt động như thế nào và cách giải quyết?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Cơ chế rò rỉ: Hàm so sánh chuỗi thông thường (như toán tử == trong nhiều ngôn ngữ) thực hiện so sánh tuần tự từng ký tự từ trái sang phải và dừng ngay lập tức (early-return) khi phát hiện ký tự sai khác đầu tiên. Kẻ tấn công gửi các chuỗi thử nghiệm và đo thời gian phản hồi: Thời gian thực thi càng dài đồng nghĩa số lượng ký tự tiền tố đoán đúng càng nhiều. Từ đó, đối thủ khôi phục được mã băm hoặc khóa bí mật mà không cần giải mã.\n"
         "2. Giải pháp kỹ thuật: Chuyển sang sử dụng hàm so sánh thời gian hằng số (Constant-time comparison) `hmac.compare_digest`. Hàm này luôn duyệt qua toàn bộ độ dài của chuỗi bằng phép toán bitwise XOR tích lũy, đảm bảo thời gian thực thi độc lập hoàn toàn với vị trí xuất hiện ký tự sai khác, triệt tiêu kênh rò rỉ thời gian ở mức phần mềm."),

        ("Câu hỏi 6: Giải pháp Cây Merkle phân cấp động (HD-MerkleTree) của đề tài có ưu thế kỹ thuật gì so với Cây Merkle truyền thống?",
         "Luận điểm phản biện & Căn cứ khoa học:\n"
         "1. Cắt giảm độ phức tạp cập nhật từ O(N) xuống O(log N): Cây Merkle nhị phân truyền thống làm phẳng toàn bộ danh mục tệp lên một mảng nút lá. Khi có 1 tệp tin được thêm mới hoặc xóa đi, cấu trúc chẵn lẻ của mảng bị phá vỡ, buộc hệ thống phải Rebuild lại toàn bộ cây với chi phí O(N) mất ~283 ms trên 2.483 tệp. HD-Merkle ánh xạ trực tiếp cấu trúc cây thư mục doanh nghiệp; cơ chế Dirty-Path Cache chỉ tính toán lại duy nhất nhánh cha bị tác động từ nút lá lên nút gốc với chi phí logarit O(log N) chỉ mất ~0.24 ms (nhanh hơn ~1.200 lần).\n"
         "2. Chuẩn phân tách miền RFC 6962: HD-Merkle áp dụng tiền tố 0x00 cho nút lá và 0x01 cho nút nhánh nội bộ, triệt tiêu 100% lỗ hổng va chạm tiền ảnh thứ hai cấp cấu trúc (Node-Leaf confusion).\n"
         "3. Tối ưu hóa băng thông bằng chứng: Gói bằng chứng CompactProof được nén xuống 256 Bytes, tiết kiệm hơn 99% băng thông so với việc truyền tải toàn bộ danh mục băm phẳng (253 KB).")
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

    # Lưu file Word
    target_path = "docs/Kich_Ban_Thuyet_Trinh_Va_Demo_DT005.docx"
    doc.save(target_path)
    print(f"[THÀNH CÔNG] Đã tạo tài liệu Word hoàn chỉnh tại: {target_path}")

if __name__ == "__main__":
    create_document()
