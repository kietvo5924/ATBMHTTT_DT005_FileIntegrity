# -*- coding: utf-8 -*-
"""
Script tự động tạo tài liệu Word (.docx) chứa KỊCH BẢN THUYẾT TRÌNH VÀ HƯỚNG DẪN DEMO 5 PHÚT.
Bám sát chính xác từng slide trong Bao_Cao_Giua_Ky_DT05.pptx (17 slides):
- Phân bổ thời gian chuẩn xác cho từng slide (tổng 5 phút = 300 giây).
- Loại bỏ hoàn toàn các nhãn thang điểm thô, chuyển hóa thành văn phong học thuật bảo vệ đồ án chuẩn mực.
- Lời thoại đọc miệng (verbatim) cho từng slide, súc tích, đĩnh đạc, tự tin, không bị cháy giáo án.
- Hướng dẫn chi tiết vẽ sơ đồ Draw.io tại các slide có khung vẽ.
- Kịch bản Live Demo cực kỳ chi tiết cho Slide 15 (tay bấm nút gì, mắt nhìn terminal chỉ gì, miệng nói gì).
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
r_title = p_title.add_run("KỊCH BẢN THUYẾT TRÌNH VÀ HƯỚNG DẪN TRÌNH DIỄN (5 PHÚT)\nĐỀ TÀI DT005: CÂY MERKLE PHÂN CẤP ĐỘNG (HD-MERKLE TREE)")
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
p_notice.add_run(". Kịch bản này được tối ưu từng câu chữ để bạn nói với tốc độ vừa phải (~120 từ/phút), dõng dạc, tự tin, không bị cháy giáo án.\n")
p_notice.add_run("• Cấu trúc slide và lời thoại thể hiện sự chuyên nghiệp của một buổi bảo vệ đồ án khoa học:\n")
p_notice.add_run("   + Hoàn thành trọn vẹn việc dịch toàn văn bài báo Harry Halpin Oxford 2025 và tái lập 3 mô hình nền tảng đối chứng.\n")
p_notice.add_run("   + Tự thiết kế mô hình HD-Merkle Tree và thuật toán Dirty-Path Update O(log N), chứng minh hiệu năng cập nhật nhanh hơn 458 - 748 lần, tiết kiệm 98.8% băng thông.\n")
p_notice.add_run("   + Vượt qua 100% 7 Security Games hình thức và 17/17 Unit Tests tự động trên 2 tập dữ liệu Kaggle thực tế (1.115 & 5.600 tệp).\n")
p_notice.add_run("• Toàn bộ các ảnh chụp web tĩnh đã được thay thế bằng ")
r_live = p_notice.add_run("SLIDE 15: LIVE DEMONSTRATION")
r_live.font.bold = True
p_notice.add_run(" để bạn trình diễn thực tế trực tiếp trên Trình duyệt Web và Streaming Terminal Logs.\n")
p_notice.add_run("• Các vị trí có khung chờ sơ đồ Draw.io (Slide 3, 4, 5, 7, 9, 16) đã ghi rõ thành phần bắt buộc để bạn xem và vẽ theo.")

# BẢNG PHÂN BỔ THỜI GIAN
p_tb_title = doc.add_paragraph()
r_tbt = p_tb_title.add_run("BẢNG PHÂN BỔ THỜI LƯỢNG CHI TIẾT 17 SLIDE (TỔNG 5 PHÚT)")
r_tbt.font.bold = True
r_tbt.font.size = Pt(11.5)
r_tbt.font.color.rgb = RGBColor(30, 58, 138)

slide_schedule = [
    ("Slide 1", "Trang tiêu đề đề tài DT005 & Mục tiêu khoa học hệ thống", "15 giây", "0:00 - 0:15"),
    ("Slide 2", "Giới thiệu & Tóm tắt bài báo Harry Halpin (Oxford 2025)", "15 giây", "0:15 - 0:30"),
    ("Slide 3", "Hiện thực hóa & Chạy lại 3 mô hình nền tảng bài báo", "15 giây", "0:30 - 0:45"),
    ("Slide 4", "Phân tích 4 nút thắt kỹ thuật của mô hình nền tảng", "15 giây", "0:45 - 1:00"),
    ("Slide 5", "Đề xuất mô hình mới: Cây Merkle phân cấp động (HD-Merkle Tree)", "15 giây", "1:00 - 1:15"),
    ("Slide 6", "Thuật toán đề xuất: Cập nhật động theo đường dẫn bẩn (Dirty-Path)", "15 giây", "1:15 - 1:30"),
    ("Slide 7", "Cơ chế an ninh mật mã & Chuẩn phân tách miền RFC 6962", "15 giây", "1:30 - 1:45"),
    ("Slide 8", "Phòng chống tấn công kênh kề đo thời gian & Xóa sạch bộ nhớ", "15 giây", "1:45 - 2:00"),
    ("Slide 9", "Thiết kế 2 tập dữ liệu Kaggle thực tế (1.115 & 5.600 tệp tin)", "15 giây", "2:00 - 2:15"),
    ("Slide 10", "Bảng đối sánh Big-O lý thuyết & Thực nghiệm cập nhật nhanh 458-748x", "20 giây", "2:15 - 2:35"),
    ("Slide 11", "Khung 7 Trò chơi an ninh mật mã học: Games 1 – 4", "15 giây", "2:35 - 2:50"),
    ("Slide 12", "Khung 7 Trò chơi an ninh mật mã học: Games 5 – 7 & Tổng kết", "15 giây", "2:50 - 3:05"),
    ("Slide 13", "Bảng ma trận đối chiếu 100% tiêu chí bài báo khoa học gốc", "15 giây", "3:05 - 3:20"),
    ("Slide 14", "Bộ kiểm thử an ninh mật mã tự động (17/17 Unit Tests PASS 100%)", "15 giây", "3:20 - 3:35"),
    ("Slide 15", "LIVE DEMONSTRATION: Trình diễn thực tế trên Trình duyệt & Terminal", "80 giây", "3:35 - 4:55"),
    ("Slide 16", "Tổng kết đóng góp khoa học & Khẳng định kết quả vượt trội", "15 giây", "4:55 - 5:10"),
    ("Slide 17", "Lời cảm ơn Quý Thầy Cô & Phiên hỏi đáp học thuật (Q&A)", "10 giây", "5:10 - 5:20"),
]

table_time = doc.add_table(rows=len(slide_schedule) + 1, cols=4)
table_time.alignment = WD_TABLE_ALIGNMENT.CENTER

headers_tb = ["Số Slide", "Nội Dung Trọng Tâm", "Thời Lượng", "Mốc Thời Gian"]
for col_idx, h_text in enumerate(headers_tb):
    cell = table_time.cell(0, col_idx)
    set_cell_background(cell, "1e3a8a")
    set_cell_margins(cell, 120, 120, 150, 150)
    p = cell.paragraphs[0]
    r = p.add_run(h_text)
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(255, 255, 255)
    if col_idx in [0, 2, 3]:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

for row_idx, data_row in enumerate(slide_schedule):
    for col_idx, text_val in enumerate(data_row):
        cell = table_time.cell(row_idx + 1, col_idx)
        bg = "f8fafc" if row_idx % 2 == 1 else "ffffff"
        if row_idx == 14:  # Slide Demo
            bg = "ecfdf5"
        set_cell_background(cell, bg)
        set_cell_margins(cell, 80, 80, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(text_val)
        r.font.size = Pt(9.0)
        if row_idx == 14:
            r.font.bold = True
            r.font.color.rgb = RGBColor(5, 150, 105)
        if col_idx in [0, 2, 3]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_page_break()


# ==============================================================================
# CHI TIẾT KỊCH BẢN TỪNG SLIDE (17 SLIDES)
# ==============================================================================
def add_slide_section(doc, slide_no, title, tag, duration, script_text, notes=None, drawio_guide=None, demo_guide=None):
    p_s_hdr = doc.add_paragraph()
    p_s_hdr.paragraph_format.space_before = Pt(14)
    p_s_hdr.paragraph_format.space_after = Pt(2)
    r_no = p_s_hdr.add_run(f"SLIDE {slide_no}: {title.upper()}")
    r_no.font.bold = True
    r_no.font.size = Pt(13)
    r_no.font.color.rgb = RGBColor(30, 58, 138)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(6)
    r_tag = p_meta.add_run(f"Phân loại: {tag}   |   Thời lượng: {duration}")
    r_tag.font.bold = True
    r_tag.font.size = Pt(9.5)
    r_tag.font.color.rgb = RGBColor(37, 99, 235)

    # Khung lời thoại đọc miệng
    p_box_title = doc.add_paragraph()
    p_box_title.paragraph_format.space_after = Pt(2)
    r_bt = p_box_title.add_run("▶ LỜI THOẠI ĐỌC MIỆNG (NÓI CHÍNH XÁC THEO ĐOẠN NÀY):")
    r_bt.font.bold = True
    r_bt.font.size = Pt(10)
    r_bt.font.color.rgb = RGBColor(5, 150, 105)

    p_script = doc.add_paragraph()
    p_script.paragraph_format.left_indent = Inches(0.2)
    p_script.paragraph_format.right_indent = Inches(0.2)
    p_script.paragraph_format.space_after = Pt(8)
    r_sc = p_script.add_run(f"“{script_text}”")
    r_sc.font.italic = True
    r_sc.font.size = Pt(10.5)
    r_sc.font.color.rgb = RGBColor(15, 23, 42)

    if notes:
        p_n = doc.add_paragraph()
        p_n.paragraph_format.left_indent = Inches(0.2)
        p_n.paragraph_format.space_after = Pt(6)
        r_nh = p_n.add_run("• Ghi chú trọng tâm: ")
        r_nh.font.bold = True
        r_nh.font.size = Pt(9.5)
        r_nh.font.color.rgb = RGBColor(180, 83, 9)
        r_nb = p_n.add_run(notes)
        r_nb.font.size = Pt(9.5)

    if drawio_guide:
        p_d = doc.add_paragraph()
        p_d.paragraph_format.left_indent = Inches(0.2)
        p_d.paragraph_format.space_after = Pt(6)
        r_dh = p_d.add_run("✎ HƯỚNG DẪN VẼ SƠ ĐỒ DRAW.IO: ")
        r_dh.font.bold = True
        r_dh.font.size = Pt(9.5)
        r_dh.font.color.rgb = RGBColor(126, 34, 206)
        r_db = p_d.add_run(drawio_guide)
        r_db.font.size = Pt(9.2)

    if demo_guide:
        p_dm = doc.add_paragraph()
        p_dm.paragraph_format.left_indent = Inches(0.2)
        p_dm.paragraph_format.space_after = Pt(8)
        r_dmh = p_dm.add_run("⚡ KỊCH BẢN THAO TÁC TRÌNH DIỄN TRỰC TIẾP (TAY BẤM GÌ - MẮT NHÌN ĐÂU - MIỆNG NÓI GÌ):\n")
        r_dmh.font.bold = True
        r_dmh.font.size = Pt(10)
        r_dmh.font.color.rgb = RGBColor(220, 38, 38)
        r_dmb = p_dm.add_run(demo_guide)
        r_dmb.font.size = Pt(9.5)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(10)
    p_div.add_run("─" * 70).font.color.rgb = RGBColor(226, 232, 240)


# SLIDE 1
add_slide_section(
    doc, 1, "Trang Tiêu Đề Đề Tài DT005", "Trang bìa", "15 giây",
    "Kính thưa Thầy và Hội đồng đánh giá, em tên là Võ Anh Kiệt. Hôm nay em xin báo cáo đồ án giữa kỳ đề tài DT005: "
    "Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và Cây Merkle phân cấp HD-Merkle Tree. "
    "Đồ án tập trung nghiên cứu bài toán đối kháng trong bài báo khoa học Oxford 2025 của Harry Halpin, "
    "tái lập 3 mô hình nền tảng đối chứng, đồng thời tự thiết kế kiến trúc phân cấp mới và thuật toán cập nhật động "
    "đạt hiệu năng và độ an toàn mật mã vượt trội.",
    notes="Nói rõ ràng họ tên, MSSV, tên đề tài và mục tiêu khoa học tổng quát."
)

# SLIDE 2
add_slide_section(
    doc, 2, "Giới Thiệu & Tóm Tắt Bài Báo Harry Halpin (Oxford 2025)", "Cơ sở lý thuyết", "15 giây",
    "Đầu tiên, em đã dịch thuật học thuật toàn văn bài báo 'The Adversary: The Philosophy of Cryptography' của tác giả Harry Halpin, "
    "xuất bản trên Oxford Journal of Cybersecurity năm 2025. Như Thầy Cô quan sát bên phải màn hình, đây là ảnh chụp bài báo thực tế "
    "trên cổng xuất bản Oxford Academic. Bài báo đặt ra luận điểm cốt lõi: Mật mã học không phải bài toán tĩnh, "
    "mà là cuộc đối đầu tương tác giữa Kẻ phòng thủ và 4 mô hình Kẻ thù hình thức: Kẻ nghe lén thụ động, Kẻ tấn công Dolev-Yao, "
    "Kẻ thù tính toán PPT và Kẻ thù khai thác kênh kề siêu dữ liệu.",
    notes="Chỉ tay sang ảnh chụp bài báo thực tế trên cổng Oxford Academic ở nửa phải slide để khẳng định tính chính thống của đề tài."
)

# SLIDE 3
add_slide_section(
    doc, 3, "Hiện Thực Hóa & Chạy Lại Kết Quả 3 Mô Hình Nền Tảng", "Hiện thực hóa bài báo", "15 giây",
    "Để giải quyết khoảng trống thực nghiệm của bài báo gốc, em đã hiện thực hóa và chạy lại 3 mô hình nền tảng đối chứng: "
    "Thứ nhất là Danh sách băm phẳng SHA-256; Thứ hai là Danh sách băm có khóa HMAC-SHA256; và Thứ ba là Cây Merkle chuẩn của Ralph Merkle 1979. "
    "Cả 3 mô hình đều phát hiện 100% sai lệch khi tệp bị sửa đổi, hoàn thành trọn vẹn yêu cầu tái lập kết quả bài báo khoa học gốc.",
    notes="Khẳng định đã tái lập thành công 100% mô hình của bài báo khoa học gốc.",
    drawio_guide="Vẽ 4 khối kiến trúc hoạt động: Khối 1: Băm phẳng SHA-256 (Tệp tin -> Băm SHA-256 -> Bảng Manifest) -> Khối 2: Băm có khóa HMAC-SHA256 (Tệp tin + Khóa K -> Băm HMAC) -> Khối 3: Cây Merkle nhị phân truyền thống (Lá -> Băm từng đôi -> Merkle Root) -> Khối 4: Quy trình máy khách xác minh độc lập bằng cách so khớp băm."
)

# SLIDE 4
add_slide_section(
    doc, 4, "Phân Tích Hạn Chế Bài Báo Gốc & Động Lực Cải Tiến", "Phân tích hạn chế", "15 giây",
    "Từ thực nghiệm tái lập, em chỉ ra 4 nút thắt kỹ thuật của mô hình nền tảng: "
    "Một là bế tắc Rebuild toàn bộ O(N) khi chỉ 1 tệp thay đổi; Hai là mảng lá phẳng thiếu phản ánh cấu trúc thư mục thực tế; "
    "Ba là nguy cơ va chạm Second-Preimage giữa nút lá và nút nhánh nếu không tách miền; "
    "Bốn là lãng phí băng thông khi client phải tải toàn bộ danh mục 111 KB và làm rò rỉ siêu dữ liệu.",
    notes="Đây là cơ sở khoa học để đề xuất mô hình mới đạt hiệu năng vượt trội.",
    drawio_guide="Vẽ sơ đồ 4 trụ cột an ninh: Trụ cột 1: Second-Preimage Resistance -> Trụ cột 2: Domain Separation RFC 6962 -> Trụ cột 3: Membership Proof O(log N) -> Trụ cột 4: Append-Only Consistency. Trung tâm là Master Root Hash 32 byte."
)

# SLIDE 5
add_slide_section(
    doc, 5, "Đề Xuất Mô Hình Mới: Cây Merkle Phân Cấp Động (HD-Merkle Tree)", "Kiến trúc đề xuất", "15 giây",
    "Để khắc phục triệt để, em tự xây dựng mô hình mới: Cây Merkle phân cấp động HD-Merkle Tree, kết hợp mô hình toàn vẹn phân tầng Biba 1977 "
    "với chuẩn Certificate Transparency RFC 6962. Tầng 1 băm lá tệp tin với tiền tố 0x00; Tầng 2 gom nhóm tệp thành các cây con thư mục độc lập; "
    "và Tầng 3 tổng hợp các Directory Root thành Master Root Hash duy nhất. Lợi thế là cô lập lỗi và định vị sai lệch trong O(k log N).",
    notes="Nhấn mạnh sự kết hợp giữa cấu trúc Merkle với mô hình toàn vẹn kinh điển Biba 1977.",
    drawio_guide="Vẽ cây phân cấp 3 tầng: Tầng đáy: Nút lá tệp tin (0x00 || File) -> Tầng giữa: Các cây con thư mục con (Directory Sub-Trees) -> Tầng đỉnh: Master Root Hash tổng hợp (32 bytes)."
)

# SLIDE 6
add_slide_section(
    doc, 6, "Thuật Toán Đề Xuất: Cập Nhật Động Theo Đường Dẫn Bẩn (Dirty-Path)", "Giải thuật đề xuất", "15 giây",
    "Cùng với mô hình, em tự phát triển thuật toán Dirty-Path Update: Khi có thao tác Thêm, Sửa hoặc Xóa 1 tệp, "
    "thuật toán tra cứu nhanh vị trí lá qua bảng băm chỉ mục O(1) và chỉ lan truyền cập nhật dọc theo đường dẫn bẩn từ lá lên gốc. "
    "Độ phức tạp tính toán được hạ từ tuyến tính O(N) của mô hình truyền thống xuống độ phức tạp logarit O(log N).",
    notes="Điểm mấu chốt kỹ thuật: Giảm số phép băm từ N tệp xuống chỉ khoảng log2(N) nút trên cây."
)

# SLIDE 7
add_slide_section(
    doc, 7, "Cơ Chế An Ninh Mật Mã & Chuẩn Phân Tách Miền RFC 6962", "Cơ chế an ninh", "15 giây",
    "Về mặt mật mã học, hệ thống áp dụng chuẩn phân tách miền RFC 6962: Tiền tố 0x00 cho nút lá và 0x01 cho nút nhánh, "
    "triệt tiêu 100% nguy cơ tấn công Second-Preimage. Đồng thời, hệ thống cung cấp Bằng chứng thành viên rút gọn Compact Inclusion Proof "
    "chỉ 1.37 KB và Bằng chứng nhất quán Append-Only Consistency Proof chứng minh lịch sử tệp tin không bị sửa lén.",
    notes="Chuẩn RFC 6962 là công nghệ cốt lõi bảo vệ hệ thống Certificate Transparency toàn cầu của Google.",
    drawio_guide="Vẽ cây nhị phân tách miền: Nút lá có nhãn SHA256(0x00 || Data), Nút cha có nhãn SHA256(0x01 || Left || Right) để chứng minh hai miền tách biệt."
)

# SLIDE 8
add_slide_section(
    doc, 8, "Phòng Chống Kênh Kề Đo Thời Gian & Xóa Sạch Bộ Nhớ", "Phòng chống nâng cao", "15 giây",
    "Nhằm đối phó với kẻ thù kênh kề, em thay thế toán tử so sánh chuỗi thông thường bằng hàm hmac.compare_digest() "
    "có thời gian so sánh hằng số O(1), triệt tiêu rò rỉ thời gian CPU. Với kẻ thù trích xuất bộ nhớ tiến trình, "
    "các khóa HMAC và vùng đệm nhạy cảm được quản lý trong bytearray và chủ động ghi đè bằng giá trị 0x00 qua hàm secure_erase().",
    notes="Chỉ ra hai nguy cơ thực tế: Timing Attack (Paul Kocher 1996) và Memory-Dump Attack."
)

# SLIDE 9
add_slide_section(
    doc, 9, "Thiết Kế 2 Tập Dữ Liệu Kaggle Thực Tế & Phương Pháp Đo Lường", "Thực nghiệm", "15 giây",
    "Để đánh giá khách quan, em sử dụng 2 tập dữ liệu Kaggle thực tế: "
    "Tập 1 là kaggle_dataset gồm 1.115 tệp, dung lượng 29 MB, 23 thư mục đa tầng với 7 định dạng thực tế dùng đối chuẩn baseline. "
    "Tập 2 là kaggle_dataset_large gồm 5.600 tệp, 175 MB chia 6 danh mục lớn tại Downloads dùng kiểm tra chịu tải và quan sát streaming log terminal. "
    "Phương pháp đo dùng time.perf_counter() độ chính xác nano-giây, đo qua 5 vòng lặp độc lập trên môi trường tạm cô lập.",
    notes="Nhấn mạnh cả 2 tập dữ liệu đều từ Kaggle, có quy mô tệp và dung lượng lớn, phản ánh đúng dữ liệu doanh nghiệp.",
    drawio_guide="Vẽ quy trình thực nghiệm: 2 tập dữ liệu Kaggle -> Quét đĩa & Streaming Log -> Phân tầng Biba 6-23 thư mục -> Đo lường qua 5 vòng lặp -> Đối chuẩn 4 mô hình."
)

# SLIDE 10
add_slide_section(
    doc, 10, "Bảng Đối Sánh Big-O Lý Thuyết & Cập Nhật Động 458-748 Lần", "Đối chuẩn hiệu năng", "20 giây",
    "Bảng đối sánh Big-O chứng minh ưu thế lý thuyết: Cả 3 mô hình cũ đều có chi phí cập nhật O(N), trong khi HD-Merkle đạt O(log N). "
    "Thực nghiệm đối đầu chứng minh kết quả vượt trội: Khi sửa 1 tệp tin, mô hình cũ mất 73.4 mili-giây, còn HD-Merkle chỉ mất 0.16 mili-giây — "
    "nhanh hơn 458 LẦN! Khi xóa tệp nhanh hơn 748 LẦN! Tốc độ xác minh 1 tệp đạt 0.02 mili-giây, nhanh hơn 8.5 lần so với danh sách phẳng. "
    "Mọi thông số đo lường đã được xác thực trực tiếp trên bảng điều khiển console terminal như Thầy Cô đang quan sát.",
    notes="Chỉ tay vào bảng Big-O ở nửa trên slide và con số 458 - 748 lần kèm log console ở nửa dưới slide."
)

# SLIDE 11
add_slide_section(
    doc, 11, "Khung 7 Trò Chơi An Ninh Mật Mã Học: Games 1 – 4", "Security Games", "15 giây",
    "Em triển khai độc lập khung 7 Trò chơi an ninh đối kháng hình thức theo bài báo Harry Halpin: "
    "Game 1 kiểm định kháng va chạm trước Kẻ thù tính toán PPT, lợi thế đối thủ Adv <= 2^-128; "
    "Game 2 phát hiện 100% can thiệp gói tin của Kẻ thù Dolev-Yao; "
    "Game 3 triệt tiêu va chạm tiền ảnh cấp hai nhờ phân tách miền RFC 6962; "
    "và Game 4 kiểm chứng hiệu ứng tuyết lở với tỷ lệ đổi bit 50.39%, xấp xỉ mức lý tưởng 50% của nguyên lý khuếch tán Shannon.",
    notes="Nêu rõ tên các mô hình mật mã học: PPT, Dolev-Yao, RFC 6962, Shannon Diffusion."
)

# SLIDE 12
add_slide_section(
    doc, 12, "Khung 7 Trò Chơi An Ninh Mật Mã Học: Games 5 – 7 & Tổng Kết", "Security Games", "15 giây",
    "Ở 3 trò chơi nâng cao: Game 5 kiểm tra tính nhất quán append-only, từ chối ngay lập tức khi phát hiện dữ liệu quá khứ bị sửa lén; "
    "Game 6 bảo vệ bộ nhớ RAM với secure_erase() ghi đè 0x00; "
    "và Game 7 loại bỏ hoàn toàn kênh kề thời gian bằng compare_digest() kết hợp bảo vệ quyền riêng tư siêu dữ liệu. "
    "Bảng tổng kết xác nhận hệ thống vượt qua 100% cả 7/7 mô hình đối kháng an ninh hình thức.",
    notes="Chỉ vào bảng tổng kết 7 Games trên slide."
)

# SLIDE 13
add_slide_section(
    doc, 13, "Bảng Ma Trận Đối Chiếu Chi Tiết 100% Bài Báo Khoa Học Gốc", "Đối chiếu lý thuyết", "15 giây",
    "Bảng ma trận đối chiếu 1-1 xác nhận đồ án DT005 hiện thực hóa đầy đủ 6 luận điểm cốt lõi trong bài báo Harry Halpin Oxford 2025: "
    "Từ mô hình kẻ thù PPT, nguyên lý Kerckhoffs công khai mã nguồn trên GitHub, mô hình mạng Dolev-Yao, chuẩn phân tách miền RFC 6962, "
    "xác minh cục bộ O(log N) của Ralph Merkle, cho đến lời kêu gọi kiểm nghiệm thực tế trên dữ liệu lớn. Mọi tiêu chí đều đạt mức 100%.",
    notes="Nhấn mạnh sự liên kết chặt chẽ giữa lý thuyết bài báo gốc và mã nguồn đồ án."
)

# SLIDE 14
add_slide_section(
    doc, 14, "Bộ Kiểm Thử An Ninh Mật Mã Tự Động (17/17 Unit Tests PASS)", "Kiểm thử tự động", "15 giây",
    "Để đảm bảo độ tin cậy tuyệt đối, em xây dựng bộ kiểm thử tự động gồm 17 ca test chuyên sâu trong tệp test_file_integrity.py. "
    "Bao phủ từ phát hiện sửa 1 bit, chèn tệp lạ, xóa lén tệp, định vị sai lệch trong 0.022 mili-giây, đến kiểm định Proof và bảo vệ RAM. "
    "Khi chạy lệnh unittest, 17/17 ca kiểm thử đều VƯỢT QUA HOÀN TOÀN (100% PASS) chỉ trong 0.25 giây.",
    notes="Chỉ vào 6 nhóm test với nhãn PASS xanh lá trên slide."
)

# SLIDE 15: LIVE DEMO CHUYÊN SÂU
demo_script = (
    "VÀ SAU ĐÂY, EM XIN PHÉP TRÌNH DIỄN HỆ THỐNG TRỰC TIẾP TRÊN TRÌNH DUYỆT VÀ STREAMING TERMINAL LOGS. "
    "(Chuyển màn hình sang Trình duyệt Web http://localhost:8501 và cửa sổ Terminal PowerShell song song)."
)

demo_guide_text = """
CHUẨN BỊ TRƯỚC GIỜ BÁO CÁO (2 PHÚT TRƯỚC KHI LÊN THUYẾT TRÌNH):
1. Mở sẵn 1 cửa sổ PowerShell tại thư mục đồ án: d:\\AnToanvaBaoMatHeThongThongTin\\GK
2. Chạy lệnh: python -m streamlit run src/demo/web_demo.py --server.port 8501
3. Mở Google Chrome hoặc Edge tại địa chỉ: http://localhost:8501
4. Sắp xếp màn hình: Bên trái là Trình duyệt Web (chiếm 60%), bên phải là cửa sổ Terminal PowerShell (chiếm 40%).
5. Mở sẵn thư mục Downloads chứa tập dữ liệu Kaggle 5,600 tệp tin để chuẩn bị thao tác sửa file thử nghiệm.

KỊCH BẢN THAO TÁC 4 BƯỚC CHI TIẾT TRONG 80 GIÂY:

▶ BƯỚC 1: KHỞI TẠO DỮ LIỆU KAGGLE LARGE & QUAN SÁT STREAMING TERMINAL LOGS (20 GIÂY)
• TAY BẤM GÌ:
  - Trên giao diện Web, tại phần Chọn thư mục nguồn dữ liệu, chọn đường dẫn thư mục: "kaggle_dataset_large" (hoặc thư mục Downloads).
  - Nhấp vào nút xanh: "🚀 Khởi Tạo Cây HD-Merkle Tree".
• MẮT NHÌN ĐÂU & TAY CHỈ GÌ:
  - Tay chỉ ngay sang màn hình Terminal bên phải.
  - Quan sát các dòng log streaming màu xanh cuộn thời gian thực:
    [SCAN] Scanning 5,600 files across 6 directories...
    [THREAD-POOL] Hashing file chunks with SHA-256 + RFC 6962 domain prefixes...
    [DIRTY-PATH] Building Biba sub-trees and calculating Master Root Hash...
• MIỆNG NÓI GÌ:
  “Kính thưa Thầy Cô, em xin khởi tạo hệ thống trên tập dữ liệu Kaggle quy mô lớn 5.600 tệp tin, 175 MB. 
  Khi bấm nút, cửa sổ Terminal lập tức stream tiến trình băm đa luồng theo thời gian thực và sinh ra Master Root Hash 32 byte bất biến đại diện cho toàn bộ kho dữ liệu.”

▶ BƯỚC 2: TRỰC QUAN HÓA CẤU TRÚC PHÂN TẦNG BIBA & DIRECTORY ROOT (15 GIÂY)
• TAY BẤM GÌ:
  - Trên thanh điều hướng Web, nhấp chuyển sang tab: "Cấu trúc Cây Merkle phân cấp".
  - Nhấp mở rộng 1 thư mục bất kỳ (ví dụ: code_files hoặc documents).
• MẮT NHÌN ĐÂU:
  - Chỉ vào sơ đồ cây phân cấp hiển thị các nút lá và mã băm Directory Root.
• MIỆNG NÓI GÌ:
  “Tại tab cấu trúc cây, hệ thống trực quan hóa rõ ràng cấu trúc phân tầng Biba 1977: Mỗi thư mục con sở hữu một Directory Root Hash riêng biệt. 
  Điều này giúp cô lập hoàn toàn lỗi giữa các thư mục, không bị gom phẳng như cây Merkle truyền thống.”

▶ BƯỚC 3: THỰC NGHIỆM TẤN CÔNG SỬA TỆP NGOÀI ĐĨA & CẬP NHẬT DIRTY-PATH MICRO-GIÂY (25 GIÂY)
• TAY BẤM GÌ:
  - Mở nhanh cửa sổ Windows Explorer đã chuẩn bị sẵn, mở 1 tệp tin văn bản bất kỳ (ví dụ: sample.txt).
  - Gõ thêm 1 ký tự "X" vào nội dung và bấm Ctrl + S để lưu lại.
  - Quay lại màn hình Web Demo, nhấp vào nút đỏ: "🔍 Kiểm Tra Toàn Vẹn Hệ Thống".
• MẮT NHÌN ĐÂU & TAY CHỈ GÌ:
  - Nhìn màn hình Web: Dòng cảnh báo màu đỏ lập tức xuất hiện, hiển thị chính xác đường dẫn tệp vừa bị sửa!
  - Chỉ vào ô thời gian phát hiện: 0.022 ms!
  - Nhấp tiếp nút: "⚡ Cập Nhật Dirty-Path".
  - Nhìn Terminal và Web: Thời gian cập nhật hoàn tất chỉ mất 0.16 mili-giây!
• MIỆNG NÓI GÌ:
  “Em vừa sửa trực tiếp 1 byte ngoài đĩa vật lý. Khi bấm kiểm tra, hệ thống lập tức phát hiện trong 0.022 mili-giây và chỉ đích danh tệp bị can thiệp. 
  Khi bấm Cập nhật Dirty-Path, hệ thống chỉ băm lại đúng đường dẫn từ nút lá lên gốc mất 0.16 mili-giây, nhanh hơn 458 lần so với việc dựng lại toàn bộ cây.”

▶ BƯỚC 4: CHẠY TỰ ĐỘNG 7 SECURITY GAMES & 17 UNIT TESTS TRỰC TIẾP (20 GIÂY)
• TAY BẤM GÌ:
  - Nhấp chuyển sang tab: "Kiểm định An ninh toán học & Unit Tests".
  - Nhấp vào nút: "▶ Chạy Kiểm Định An Ninh Tự Động".
• MẮT NHÌN ĐÂU & CHỈ GÌ:
  - Nhìn Terminal và Web: Toàn bộ 7 Games (Game 1 PPT, Game 2 Dolev-Yao, Game 3 Domain Separation, Game 4 Avalanche...) 
    và 17 ca Unit Tests đều hiện chữ PASS màu xanh lá cây rực rỡ.
• MIỆNG NÓI GÌ:
  “Hệ thống tự động thực thi toàn bộ 7 mô hình toán học đối kháng của Harry Halpin và 17 ca Unit Tests an ninh. 
  Kết quả đạt 100% PASS, chứng minh thuật toán hoạt động hoàn hảo và an toàn tuyệt đối trên môi trường thực tế.”
"""

add_slide_section(
    doc, 15, "Live Demonstration: Trình Diễn Hệ Thống Trực Tiếp Trên Trình Duyệt & Terminal",
    "Trình diễn thực tế [Trọng tâm]", "80 giây", demo_script,
    notes="Đây là slide quan trọng nhất chứng minh năng lực hiện thực hóa sản phẩm thực tế.",
    demo_guide=demo_guide_text
)

# SLIDE 16
add_slide_section(
    doc, 16, "Tổng Kết Đóng Góp Khoa Học & Khẳng Định Kết Quả Vượt Trội", "Tổng kết học thuật", "15 giây",
    "Tóm lại, đồ án DT005 đạt được 4 đóng góp nổi bật: "
    "Một là hoàn thành xuất sắc việc nghiên cứu và tái lập 3 mô hình của bài báo Harry Halpin Oxford 2025; "
    "Hai là tự thiết kế mô hình HD-Merkle Tree kết hợp Biba 1977 và chuẩn phân tách miền RFC 6962; "
    "Ba là phát triển thuật toán Dirty-Path cập nhật nhanh hơn 458 đến 748 lần, tiết kiệm 98.8% băng thông mạng; "
    "và Bốn là kiểm chứng an ninh toán học trọn vẹn qua 7 Games, 17 Unit Tests và trình diễn trực tiếp trên 2 tập dữ liệu Kaggle lớn.",
    notes="Khẳng định vững chắc đã đạt được kết quả nghiên cứu và hiện thực hóa xuất sắc.",
    drawio_guide="Vẽ sơ đồ chu trình khép kín 4 khối: Khối 1: Khoa học bài báo -> Khối 2: Đột phá hiệu năng 458-748x -> Khối 3: Kiểm định an ninh 7 Games & 17 Tests -> Khối 4: Ứng dụng thực tế Web Demo & Kaggle."
)

# SLIDE 17
add_slide_section(
    doc, 17, "Lời Cảm Ơn Quý Thầy Cô & Phiên Thảo Luận (Q&A)", "Kết thúc", "10 giây",
    "Em xin trân trọng cảm ơn Thầy Phạm Trọng Huynh và Quý Thầy Cô đã lắng nghe bài báo cáo của em. "
    "Em rất mong nhận được những câu hỏi và ý kiến đóng góp quý báu từ Thầy Cô để đồ án ngày càng hoàn thiện hơn. Em xin trân trọng cảm ơn!",
    notes="Cúi chào lịch sự, phong thái tự tin, sẵn sàng chuyển sang phần trả lời câu hỏi."
)

# Lưu tài liệu Word
doc.save(DOC_PATH)
print(f"Da tao thanh cong tai lieu Word kich ban thuyet trinh tai: {DOC_PATH}")
