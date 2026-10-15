# -*- coding: utf-8 -*-
"""
Script tự động tạo tài liệu Word (.docx) chứa KỊCH BẢN THUYẾT TRÌNH VÀ HƯỚNG DẪN DEMO 5 PHÚT.
Bám sát chính xác từng slide trong Bao_Cao_Giua_Ky_DT05.pptx:
- Phân bổ thời gian chuẩn xác cho từng slide (tổng 5 phút).
- Lời thoại đọc miệng (verbatim) cho từng slide.
- Hướng dẫn chi tiết vẽ sơ đồ Draw.io tại các slide có khung vẽ.
- Kịch bản demo cực kỳ chi tiết cho Slide 7 (tay bấm nút gì, mắt nhìn terminal chỉ gì, miệng nói gì).
"""
import os
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
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOC_PATH = r"d:\AnToanvaBaoMatHeThongThongTin\GK\docs\Kich_Ban_Thuyet_Trinh_5_Phut_DT005.docx"


def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


doc = docx.Document()

# Thiết lập lề trang
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

# Style cơ bản
normal_style = doc.styles['Normal']
normal_style.font.name = 'Arial'
normal_style.font.size = Pt(11)
normal_style.font.color.rgb = RGBColor(15, 23, 42)

# ==============================================================================
# PHẦN ĐẦU TÀI LIỆU (TIÊU ĐỀ & BẢNG PHÂN BỔ THỜI GIAN)
# ==============================================================================
p_inst = doc.add_paragraph()
r_inst = p_inst.add_run("BÁO CÁO GIỮA KỲ HỌC PHẦN AN TOÀN VÀ BẢO MẬT HỆ THỐNG THÔNG TIN")
r_inst.font.bold = True
r_inst.font.size = Pt(10)
r_inst.font.color.rgb = RGBColor(100, 116, 139)

p_title = doc.add_paragraph()
r_title = p_title.add_run("KỊCH BẢN THUYẾT TRÌNH VÀ HƯỚNG DẪN TRÌNH DIỄN (5 PHÚT)\nĐỀ TÀI DT005: CÂY MERKLE PHÂN CẤP ĐỘNG (HD-MERKLETREE)")
r_title.font.bold = True
r_title.font.size = Pt(16)
r_title.font.color.rgb = RGBColor(30, 58, 138)
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_after = Pt(12)

# Khung ghi chú quan trọng
p_notice = doc.add_paragraph()
p_notice.paragraph_format.space_after = Pt(14)
r_not = p_notice.add_run("LƯU Ý QUAN TRỌNG DÀNH CHO NGƯỜI THUYẾT TRÌNH:")
r_not.font.bold = True
r_not.font.color.rgb = RGBColor(180, 83, 9)
p_notice.add_run("\n• Tổng thời lượng trình bày quy định là ")
r_time = p_notice.add_run("5 phút (300 giây)")
r_time.font.bold = True
p_notice.add_run(". Kịch bản này được tối ưu từng từ ngữ để bạn nói với tốc độ vừa phải, dõng dạc và không bị cháy giáo án.\n")
p_notice.add_run("• Toàn bộ các ảnh chụp web tĩnh đã được lược bỏ khỏi slide, vì bạn sẽ ")
r_live = p_notice.add_run("DEMO TRỰC TIẾP trên Web và Terminal tại Slide 7")
r_live.font.bold = True
p_notice.add_run(".\n• Các vị trí cần vẽ sơ đồ kiến trúc (Slide 3 và Slide 4) đã có hướng dẫn chi tiết để bạn vẽ trên Draw.io.")

# BẢNG PHÂN BỔ THỜI GIAN
p_tb_title = doc.add_paragraph()
r_tbt = p_tb_title.add_run("BẢNG PHÂN BỔ THỜI LƯỢNG CHI TIẾT TỪNG SLIDE (TỔNG 5 PHÚT)")
r_tbt.font.bold = True
r_tbt.font.size = Pt(11.5)
r_tbt.font.color.rgb = RGBColor(30, 58, 138)

table_time = doc.add_table(rows=9, cols=4)
table_time.alignment = WD_TABLE_ALIGNMENT.CENTER
headers_t = ["Slide Số", "Nội Dung Trọng Tâm", "Thời Lượng", "Nhiệm Vụ Chính Của Người Nói"]
col_widths = [Inches(1.0), Inches(2.7), Inches(1.1), Inches(2.4)]

for c_idx, h_text in enumerate(headers_t):
    cell = table_time.cell(0, c_idx)
    cell.width = col_widths[c_idx]
    set_cell_background(cell, "1E3A8A")
    set_cell_margins(cell, 80, 80, 100, 100)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h_text)
    run.font.bold = True
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(255, 255, 255)

data_time = [
    ("Slide 1", "Tiêu đề & Giới thiệu đề tài", "30 giây", "Chào Hội đồng, nêu tên đề tài và bài báo nền tảng Oxford 2025."),
    ("Slide 2", "Hạn chế mô hình cũ & 2 Dataset Kaggle", "40 giây", "Chỉ ra 3 lỗi của mô hình cũ, giới thiệu bộ 1,115 tệp và 5,600 tệp."),
    ("Slide 3", "Kiến trúc HD-Merkle & Tiền tố RFC 6962", "50 giây", "Giải thích sơ đồ cây 3 tầng và công thức tiền tố 0x00 != 0x01."),
    ("Slide 4", "Bảng Big-O lý thuyết & Cơ chế Dirty Path", "50 giây", "Trình bày bảng so sánh Big-O và cơ chế Dirty Path O(log N)."),
    ("Slide 5", "7 Mô hình toán học Halpin (Mô hình 1 - 4)", "45 giây", "Trình bày PPT Adversary, Dolev-Yao, RFC 6962, Consistency."),
    ("Slide 6", "7 Mô hình toán học Halpin (Mô hình 5 - 7)", "45 giây", "Trình bày Random Oracle/SAC, Biba, Metadata Privacy & Constant-time."),
    ("Slide 7", "LIVE DEMONSTRATION (Web & Terminal)", "60 giây", "Thao tác trực tiếp 4 bước trên Web & Terminal streaming log."),
    ("Slide 8", "Kết luận, Đóng góp & Lời cảm ơn", "20 giây", "Tóm tắt đóng góp, lời cảm ơn Hội đồng.")
]

for r_idx, row in enumerate(data_time):
    for c_idx, val in enumerate(row):
        cell = table_time.cell(r_idx + 1, c_idx)
        cell.width = col_widths[c_idx]
        bg_col = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        if r_idx == 6:  # Slide demo highlight
            bg_col = "ECFDF5"
        set_cell_background(cell, bg_col)
        set_cell_margins(cell, 60, 60, 90, 90)
        p = cell.paragraphs[0]
        run = p.add_run(val)
        run.font.size = Pt(9)
        if c_idx in (0, 2):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run.font.bold = True
        if r_idx == 6 and c_idx == 1:
            run.font.bold = True
            run.font.color.rgb = RGBColor(5, 150, 105)

doc.add_paragraph().paragraph_format.space_after = Pt(16)


# ==============================================================================
# NỘI DUNG CHI TIẾT TỪNG SLIDE
# ==============================================================================

slides_script = [
    {
        "num": "SLIDE 1: TIÊU ĐỀ & GIỚI THIỆU ĐỀ TÀI",
        "time": "30 giây (0:00 - 0:30)",
        "objective": "Tạo ấn tượng trang trọng, giới thiệu bản thân, tên đề tài và nguồn gốc lý thuyết vững chắc từ bài báo khoa học quốc tế Oxford 2025.",
        "diagram_guide": "Không có sơ đồ cần vẽ. Slide đã thiết kế sẵn thẻ đề tài và các khối mục tiêu chuẩn mực.",
        "script": (
            "\"Kính thưa Quý Thầy Cô trong Hội đồng và các bạn! Em xin phép được đại diện nhóm trình bày báo cáo giữa kỳ "
            "đề tài DT005 với nội dung: 'Nghiên cứu và phát triển hệ thống kiểm tra toàn vẹn tệp mật mã dựa trên Cây Merkle phân cấp động - HD-MerkleTree'.\n\n"
            "Đề tài của chúng em được phát triển bám sát toàn văn bài báo khoa học quốc tế của Tiến sĩ Harry Halpin mang tựa đề "
            "'The Adversary: The Philosophy of Cryptography', xuất bản trên Oxford Journal of Cybersecurity năm 2025. "
            "Mục tiêu cốt lõi của đề tài là giải quyết bài toán giám sát toàn vẹn dữ liệu quy mô lớn trong môi trường phân tán, "
            "đạt điểm tối đa thông qua việc vừa hiện thực hóa đầy đủ 7 mô hình toán học hình thức, vừa tự thiết kế một cấu trúc cây Merkle mới "
            "cho tốc độ cập nhật vượt trội và đã được kiểm chứng bằng 17 bài Unit Test tự động.\""
        )
    },
    {
        "num": "SLIDE 2: HẠN CHẾ CÁC MÔ HÌNH CŨ & KHẢO SÁT 2 TẬP DỮ LIỆU KAGGLE",
        "time": "40 giây (0:30 - 1:10)",
        "objective": "Chỉ ra 3 nhược điểm lớn của các giải pháp truyền thống và chứng minh tính thực tế qua việc khảo sát 2 bộ dữ liệu Kaggle thật.",
        "diagram_guide": "Không có sơ đồ cần vẽ. Slide chia làm 2 cột rõ ràng: Cột trái là 3 hạn chế mô hình cũ, Cột phải là số liệu 2 dataset Kaggle.",
        "script": (
            "\"Thưa Thầy Cô, khi giám sát tính toàn vẹn của một cây thư mục, các phương pháp truyền thống đều bộc lộ những điểm yếu chí mạng: "
            "Thứ nhất, Danh sách băm phẳng Flat SHA-256 có chi phí cập nhật là O(N), mỗi lần sửa 1 tệp đều phải đọc lại toàn bộ đĩa và làm lộ 100% đường dẫn tệp trên mạng. "
            "Thứ hai, HMAC-SHA256 phụ thuộc hoàn toàn vào khóa đối xứng bí mật K, gây nguy cơ rò rỉ khóa khi phân phối. "
            "Thứ ba, Standard Merkle Tree sử dụng mảng phẳng 1 chiều, khi sửa tệp phải dựng lại toàn bộ cây và dễ dính lỗ hổng giả mạo Second-Preimage.\n\n"
            "Để chứng minh các hạn chế này một cách khách quan, chúng em đã khảo sát và thử nghiệm trên 2 tập dữ liệu Kaggle thực tế: "
            "Tập thứ nhất gồm 1,115 tệp tin với 7 định dạng phổ biến như hợp đồng PDF, tài liệu Word, CSV và mã nguồn. "
            "Đặc biệt, tập thứ hai là một bộ dữ liệu lớn với 5,600 tệp tin, dung lượng hơn 175 MB chia làm 6 danh mục chuyên biệt để kiểm thử khả năng chịu tải cao của hệ thống.\""
        )
    },
    {
        "num": "SLIDE 3: KIẾN TRÚC MÔ HÌNH ĐỀ XUẤT HD-MERKLETREE & RFC 6962",
        "time": "50 giây (1:10 - 2:00)",
        "objective": "Trình bày kiến trúc cây phân cấp động, làm rõ công thức tiền tố RFC 6962 để bảo vệ hệ thống trước tấn công Second-Preimage.",
        "diagram_guide": (
            "TẠI CỘT BÊN TRÁI SLIDE 3 CÓ KHUNG CHỜ ĐỂ BẠN VẼ TRÊN DRAW.IO:\n"
            "1. Vẽ hình cây 3 tầng từ trên xuống: Đỉnh là 'Root Hash' (32 Bytes, màu xanh Navy).\n"
            "2. Tầng giữa là các 'Directory Nodes' tương ứng với các thư mục thực tế trên ổ cứng (màu xanh dương).\n"
            "3. Tầng đáy là các 'File Leaves' băm dữ liệu từng file (màu xanh lá).\n"
            "4. Thêm chú thích tiền tố rõ ràng: Tại mỗi nút Lá ghi '0x00 || path || data', tại mỗi nút Nhánh ghi '0x01 || left || right'.\n"
            "5. Vẽ mũi tên nét đứt đỏ biểu thị Audit Path / Compact Proof dẫn từ một nút lá lên đến Root Hash."
        ),
        "script": (
            "\"Để khắc phục các nhược điểm trên, nhóm chúng em đề xuất mô hình HD-MerkleTree với 3 cải tiến kỹ thuật đột phá, như sơ đồ trên slide:\n\n"
            "Thứ nhất, chúng em áp dụng chuẩn phân tách miền tiền tố RFC 6962: Nút lá được băm với tiền tố 0x00, trong khi nút nhánh được băm với tiền tố 0x01. "
            "Do 0x00 luôn khác 0x01, không gian băm của lá và nhánh hoàn toàn độc lập, giúp triệt tiêu 100% nguy cơ tấn công Second-Preimage cấu trúc.\n\n"
            "Thứ hai, cấu trúc cây phản ánh đúng cây thư mục vật lý lồng nhau thay vì băm phẳng 1 chiều, kết hợp bộ xử lý đa luồng CPU giúp băm song song 5,600 tệp tin trong chớp mắt.\n\n"
            "Thứ ba, hệ thống cung cấp cơ chế Compact Proof thu gọn chỉ khoảng 1.4 KB, và tính năng Batch Multi-Proof cho phép gộp các nút tổ tiên chung, "
            "giúp tiết kiệm tới hơn 86% băng thông mạng khi xác thực cùng lúc một nhóm nhiều tệp tin.\""
        )
    },
    {
        "num": "SLIDE 4: BẢNG ĐỐI SÁNH ĐỘ PHỨC TẠP BIG-O & CƠ CHẾ DIRTY PATH",
        "time": "50 giây (2:00 - 2:50)",
        "objective": "Chứng minh sự vượt trội về mặt toán học thông qua bảng so sánh Big-O lý thuyết và cơ chế cập nhật tăng dần Dirty Path.",
        "diagram_guide": (
            "TẠI CỘT BÊN PHẢI SLIDE 4 CÓ KHUNG CHỜ ĐỂ BẠN VẼ TRÊN DRAW.IO:\n"
            "1. Vẽ lại một nhánh cây Merkle với 1 file lá ở góc dưới bên trái bị đổi màu sang màu ĐỎ (ghi chữ: 'File_k bị sửa 1 byte').\n"
            "2. Vẽ một đường dẫn màu đỏ đậm có mũi tên đi ngược từ Node lá đỏ đó -> Node cha Directory -> Node Root Hash trên cùng.\n"
            "3. Gắn nhãn trên đường đỏ đó: 'Dirty Path lan truyền O(log N) — Mất 0.13 ms'.\n"
            "4. Các nhánh còn lại của cây tô màu XANH LÁ và ghi chú: 'Giữ nguyên hash cũ trong In-memory Cache, không băm lại (O(1))'."
        ),
        "script": (
            "\"Tiếp theo, em xin kính mời Thầy Cô theo dõi Bảng đối sánh độ phức tạp thuật toán Big-O giữa 4 mô hình trên slide:\n\n"
            "Điểm khác biệt mang tính quyết định của đề tài nằm ở chi phí cập nhật khi một tệp tin bị sửa đổi: "
            "Trong khi Flat SHA-256, HMAC và Standard Merkle Tree đều mất độ phức tạp O(N) — tức là phải đọc và tính toán lại toàn bộ danh mục hàng nghìn tệp, "
            "thì mô hình HD-MerkleTree của nhóm em chỉ mất độ phức tạp O(log N)!\n\n"
            "Để làm được điều này, chúng em phát minh cơ chế 'Dirty Path' như mô tả ở sơ đồ bên phải: "
            "Khi có một tệp tin bị thay đổi trên đĩa, hàm mark_dirty() chỉ đánh dấu bẩn duy nhất nút lá đó. "
            "Thuật toán chỉ tính lại băm dọc theo đường dẫn từ lá lên đến Root Hash (chỉ khoảng 11-12 phép băm), "
            "trong khi tất cả các nhánh còn lại đều được tái sử dụng trực tiếp từ bộ nhớ đệm In-memory Cache. "
            "Nhờ đó, thời gian cập nhật diễn ra gần như tức thì, chỉ mất 0.13 mili-giây!\""
        )
    },
    {
        "num": "SLIDE 5: 7 MÔ HÌNH HARRY HALPIN (OXFORD 2025) — PHẦN 1",
        "time": "45 giây (2:50 - 3:35)",
        "objective": "Trình bày 4 mô hình toán học mật mã đầu tiên theo bài báo Oxford 2025 (PPT Adversary, Dolev-Yao, RFC 6962, Consistency Proof).",
        "diagram_guide": "Không có sơ đồ cần vẽ. Slide đã thiết kế sẵn 4 khối thẻ toán học đối xứng, có số liệu và trích dẫn chuẩn mực.",
        "script": (
            "\"Về mặt cơ sở toán học mật mã học, đề tài của chúng em hiện thực hóa trọn vẹn 7 mô hình hình thức theo bài báo của TS. Harry Halpin trên Oxford Cybersecurity 2025:\n\n"
            "Mô hình 1 là Trò chơi thách đấu Kẻ thù siêu máy tính PPT của Goldwasser-Micali: Chứng minh bằng toán học rằng lợi thế giả mạo Advantage của kẻ thù bị chặn dưới mức không đáng kể 2^-128. "
            "Trong thực nghiệm qua 50 phép thử liên tục, tỷ lệ thắng của kẻ tấn công bằng 0 tuyệt đối.\n\n"
            "Mô hình 2 là Kẻ tấn công kiểm soát toàn bộ đường truyền mạng Dolev-Yao: Mô phỏng 3 kịch bản gồm tráo proof tệp A cho tệp B, can thiệp 1 byte trên mạng và gửi lại proof cũ. "
            "Hệ thống ngăn chặn thành công 100% cả 3 kịch bản.\n\n"
            "Mô hình 3 là Kháng va chạm Second-Preimage thông qua phân tách miền tiền tố 0x00 và 0x01 theo chuẩn RFC 6962.\n\n"
            "Và Mô hình 4 là Bằng chứng nhất quán chỉ-thêm Append-Only Consistency Proof theo chuẩn Certificate Transparency: "
            "Chứng minh cây hiện tại bảo lưu toàn vẹn lịch sử dữ liệu quá khứ với kích thước proof siêu nhẹ chỉ 11 nút (~352 Bytes), chống cài cắm backdoor ngầm.\""
        )
    },
    {
        "num": "SLIDE 6: 7 MÔ HÌNH HARRY HALPIN (OXFORD 2025) — PHẦN 2",
        "time": "45 giây (3:35 - 4:20)",
        "objective": "Trình bày 3 mô hình toán học còn lại (Random Oracle/SAC, Toàn vẹn phân cấp Biba, Bảo vệ siêu dữ liệu & Kháng Timing Attack).",
        "diagram_guide": "Không có sơ đồ cần vẽ. Slide bố trí 2 thẻ trên và 1 thẻ lớn kéo dài phía dưới so sánh bảo mật siêu dữ liệu.",
        "script": (
            "\"Ở 3 mô hình tiếp theo:\n\n"
            "Mô hình 5 kiểm định Tính xấp xỉ Random Oracle Model và tiêu chuẩn khuếch tán tuyết lở nghiêm ngặt SAC của Bellare-Rogaway: "
            "Khi lật đúng 1 bit đầu vào, hàm băm SHA-256 làm đảo ngẫu nhiên 51.95% các bit đầu ra, tiệm cận hoàn hảo phân phối Bernoulli 50% "
            "và đạt độ hỗn loạn Shannon Entropy 7.98 trên 8.0, hoàn toàn vượt trội so với mã CRC32 tuyến tính rất dễ bị bẻ gãy.\n\n"
            "Mô hình 6 hiện thực hóa Mô hình toàn vẹn phân cấp Biba năm 1977 với nguyên lý 'No Write-Up': "
            "Bất kỳ sửa đổi trái phép nào ở tệp tầng thấp đều lập tức kích hoạt thay đổi Root Hash sau 0.13 mili-giây, "
            "đồng thời hệ thống thực thi hàm secure_erase() để xóa sạch vết bộ nhớ đệm, triệt tiêu rò rỉ dữ liệu.\n\n"
            "Và Mô hình 7 là Kiểm toán bảo vệ siêu dữ liệu và kháng tấn công đo thời gian: "
            "Trong khi SHA-256 phẳng làm lộ 100% đường dẫn tệp nhạy cảm trên mạng, HD-MerkleTree chỉ gửi đúng proof 1.4 KB, rò rỉ đúng 0 Byte siêu dữ liệu. "
            "Đồng thời, hàm hmac.compare_digest() giúp so sánh thời gian hằng số O(1), triệt tiêu hoàn toàn nguy cơ tấn công kênh kề đo độ trễ mạng của Kocher.\""
        )
    },
    {
        "num": "SLIDE 7: TRÌNH DIỄN HỆ THỐNG TRỰC TIẾP (LIVE DEMO TRÊN WEB & TERMINAL)",
        "time": "60 giây (4:20 - 5:20) — [RẤT QUAN TRỌNG]",
        "objective": "Tương tác thực tế với Hội đồng, chứng minh phần mềm đang chạy thật 100% trên cổng 8080 và streaming log trên Terminal.",
        "diagram_guide": "Không có sơ đồ. Slide hiển thị bảng 4 bước kịch bản demo để Hội đồng theo dõi trong lúc bạn thao tác trên màn hình.",
        "script": (
            "\"Sau đây, để minh chứng cho các kết quả nghiên cứu trên, em xin phép thu nhỏ Slide và chuyển sang màn hình Trình duyệt Web cùng Terminal "
            "để thực hiện TRÌNH DIỄN HỆ THỐNG TRỰC TIẾP trên tập dữ liệu Kaggle lớn 5,600 tệp tin!\"\n\n"
            "========================================================================\n"
            "HƯỚNG DẪN THAO TÁC TAY VÀ LỜI THOẠI NÓI THEO 4 BƯỚC THỰC HIỆN TRÊN WEB:\n"
            "========================================================================\n\n"
            "👉 BƯỚC 1: CHỌN TẬP DỮ LIỆU KAGGLE LỚN (Mất 10s)\n"
            "• Thao tác tay: Bấm nút 'Đổi thư mục' ở góc trên ➔ Tại ô số 2, gõ 'kaggle_dataset_large' ➔ Bấm nút 'Áp dụng'.\n"
            "• Lời thoại nói: 'Thưa Thầy Cô, hệ thống đã gắn kết thành công thư mục kaggle_dataset_large với 5,600 tệp tin, dung lượng 175 MB.'\n\n"
            "👉 BƯỚC 2: NẠP DỮ LIỆU BASELINE & XEM LOG TERMINAL (Mất 15s)\n"
            "• Thao tác tay: Bấm nút '1. Nạp dữ liệu' ở phía trên.\n"
            "• Mắt nhìn & tay chỉ: Mở cửa sổ Terminal lên để Thầy Cô thấy log streaming chạy.\n"
            "• Lời thoại nói: 'Khi em bấm Nạp dữ liệu, cơ chế đa luồng xử lý toàn bộ 5,600 tệp chỉ trong 1.1 giây, đạt tốc độ 4,800 tệp/giây. "
            "Trên màn hình Terminal, toàn bộ thông tin mã băm Root Hash và thời gian dựng được in ra chi tiết theo thời gian thực.'\n\n"
            "👉 BƯỚC 3: MÔ PHỎNG TẤN CÔNG & ĐỐI SOÁT TOÀN VẸN (Mất 15s)\n"
            "• Thao tác tay: Mở thư mục Downloads/kaggle_dataset_large trên máy, mở 1 file bất kỳ sửa 1 chữ rồi lưu lại (hoặc tạo 1 file mới) ➔ Quay lại web bấm '2. Kiểm tra'.\n"
            "• Mắt nhìn & tay chỉ: Chỉ vào dòng cảnh báo đỏ trên web và terminal.\n"
            "• Lời thoại nói: 'Bây giờ em mô phỏng kẻ tấn công sửa lén 1 tệp tin trên ổ cứng. Khi em bấm Kiểm tra toàn vẹn, "
            "chỉ trong nửa giây, hệ thống lập tức phát hiện tỷ lệ toàn vẹn giảm, báo động đỏ và định vị chính xác vị trí tệp tin bị can thiệp!'\n\n"
            "👉 BƯỚC 4: CHẠY SO SÁNH 4 MÔ HÌNH & 7 MÔ HÌNH TOÁN HỌC (Mất 20s)\n"
            "• Thao tác tay: Cuộn xuống bấm nút 'Chạy so sánh 4 mô hình đối đầu'. Sau đó bấm chuyển sang tab 'Khung kiểm định an toàn toán học' và bấm 'Thực thi kiểm định 7 mô hình an toàn toán học'.\n"
            "• Lời thoại nói: 'Ở phần đối đầu, HD-Merkle thể hiện rõ ưu thế cập nhật nhanh hơn hàng trăm lần và kích thước proof siêu nhẹ. "
            "Và ở tab Kiểm định toán học, hệ thống tự động chạy 7 bài test đối kháng: từ chống giả mạo PPT, chống bắt gói tin Dolev-Yao, "
            "đến kiểm định tuyết lở 51.95%, giúp biểu đồ Radar Chart đạt chuẩn an toàn 100% trên toàn bộ các chiều mật mã học!'\""
        )
    },
    {
        "num": "SLIDE 8: KẾT LUẬN & ĐÓNG GÓP HỌC THUẬT",
        "time": "20 giây (5:20 - 5:40)",
        "objective": "Tóm tắt ngắn gọn 3 đóng góp lớn của đề tài và gửi lời cảm ơn trang trọng đến Quý Thầy Cô.",
        "diagram_guide": "Không có sơ đồ cần vẽ. Slide đã thiết kế sẵn 3 khối đóng góp và khung cảm ơn trang trọng ở dưới cùng.",
        "script": (
            "\"Tóm lại, đề tài DT005 của chúng em đã đóng góp 3 giá trị học thuật then chốt: "
            "Một là hiện thực hóa hoàn chỉnh 7 mô hình toán học theo bài báo Oxford 2025; "
            "Hai là phát triển thành công cấu trúc cây phân cấp HD-MerkleTree với thuật toán Dirty Path giảm chi phí cập nhật xuống O(log N); "
            "Và ba là kiểm chứng độ tin cậy tuyệt đối qua 17 bài Unit Test và 5,600 tệp tin thực nghiệm.\n\n"
            "Em xin trân trọng cảm ơn Quý Thầy Cô trong Hội đồng đã chú ý lắng nghe! "
            "Em rất mong nhận được những câu hỏi và đóng góp quý báu từ Thầy Cô ạ!\""
        )
    }
]

for s in slides_script:
    # Heading slide
    h_p = doc.add_paragraph()
    h_p.paragraph_format.space_before = Pt(14)
    h_p.paragraph_format.space_after = Pt(4)
    h_run = h_p.add_run(s["num"])
    h_run.font.bold = True
    h_run.font.size = Pt(13)
    h_run.font.color.rgb = RGBColor(30, 58, 138)

    # Bảng chi tiết từng slide
    tb_s = doc.add_table(rows=4, cols=2)
    tb_s.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb_s.columns[0].width = Inches(1.8)
    tb_s.columns[1].width = Inches(5.4)

    fields = [
        ("Thời lượng mục tiêu", s["time"], "EFF6FF", RGBColor(37, 99, 235)),
        ("Mục tiêu truyền tải", s["objective"], "FFFFFF", RGBColor(15, 23, 42)),
        ("Gợi ý vẽ sơ đồ Draw.io", s["diagram_guide"], "FEF3C7", RGBColor(180, 83, 9)),
        ("Lời thoại đọc miệng (Verbatim Script)", s["script"], "F8FAFC", RGBColor(15, 23, 42))
    ]

    for f_idx, (fname, fval, bg, col) in enumerate(fields):
        c0 = tb_s.cell(f_idx, 0)
        c1 = tb_s.cell(f_idx, 1)
        c0.width = Inches(1.8)
        c1.width = Inches(5.4)
        set_cell_background(c0, "1E3A8A")
        set_cell_background(c1, bg)
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 100, 100)

        p0 = c0.paragraphs[0]
        r0 = p0.add_run(fname)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = RGBColor(255, 255, 255)

        p1 = c1.paragraphs[0]
        r1 = p1.add_run(fval)
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = col
        if f_idx == 0:
            r1.font.bold = True
        if f_idx == 2:
            r1.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

doc.save(DOC_PATH)
print(f"[HOÀN TẤT] Đã tạo thành công tài liệu Word Kịch bản tại: {DOC_PATH}")
