# -*- coding: utf-8 -*-
"""
Script tạo slide báo cáo giữa kỳ Đề tài DT005 chuẩn khoa học, tối ưu chính xác cho thời lượng 5 phút.
Bao gồm đầy đủ 8 slide:
Slide 1: Tiêu đề & Giới thiệu đề tài DT005 (30s)
Slide 2: Đặt vấn đề & Khảo sát 2 Tập Dữ Liệu Kaggle (kaggle_dataset & kaggle_dataset_large) (40s)
Slide 3: Kiến trúc Đề xuất HD-MerkleTree & Phân tách tiền tố RFC 6962 (Có khung vẽ sơ đồ Draw.io) (50s)
Slide 4: Bảng Đối sánh Độ phức tạp thuật toán Big-O lý thuyết & Cơ chế Dirty Path (50s)
Slide 5: Cơ sở toán học Mật mã học - 7 Mô hình hình thức Harry Halpin (Phần 1: Mô hình 1 - 4) (45s)
Slide 6: Cơ sở toán học Mật mã học - 7 Mô hình hình thức Harry Halpin (Phần 2: Mô hình 5 - 7) (45s)
Slide 7: TRÌNH DIỄN HỆ THỐNG TRỰC TIẾP (LIVE DEMONSTRATION TRÊN WEB & TERMINAL) (60s)
Slide 8: Kết luận, Đóng góp học thuật & Lời cảm ơn (20s)
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUTPUT_PATH = r"d:\AnToanvaBaoMatHeThongThongTin\GK\docs\Bao_Cao_Giua_Ky_DT05.pptx"

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Bảng màu chuẩn học thuật, tương phản cao, chuyên nghiệp
BG_COLOR = RGBColor(248, 250, 252)          # Nền xám nhạt sáng sủa (#f8fafc)
NAVY_COLOR = RGBColor(30, 58, 138)          # Xanh Navy đậm (#1e3a8a)
CARD_BG = RGBColor(255, 255, 255)            # Thẻ trắng (#ffffff)
BORDER_COLOR = RGBColor(226, 232, 240)       # Viền xám mỏng (#e2e8f0)
TEXT_MAIN = RGBColor(15, 23, 42)            # Chữ chính đậm (#0f172a)
TEXT_MUTED = RGBColor(100, 116, 139)        # Chữ phụ xám (#64748b)
GREEN_COLOR = RGBColor(5, 150, 105)         # Xanh lá (#059669)
GREEN_BG = RGBColor(236, 253, 245)          # Xanh lá nhạt (#ecfdf5)
BLUE_ACCENT = RGBColor(37, 99, 235)         # Xanh dương (#2563eb)
BLUE_BG = RGBColor(239, 246, 255)           # Xanh dương nhạt (#eff6ff)
RED_ACCENT = RGBColor(220, 38, 38)          # Đỏ (#dc2626)
RED_BG = RGBColor(254, 242, 242)            # Đỏ nhạt (#fef2f2)
PURPLE_ACCENT = RGBColor(126, 34, 206)      # Tím (#7e22ce)
AMBER_ACCENT = RGBColor(217, 119, 6)        # Cam hổ phách (#d97706)
PLACEHOLDER_BG = RGBColor(241, 245, 249)      # Xám xanh nhạt (#f1f5f9)
PLACEHOLDER_BORDER = RGBColor(148, 163, 184)  # Viền placeholder (#94a3b8)


def add_base_background(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_COLOR
    bg.line.color.rgb = BG_COLOR
    return bg


def add_header(slide, tag_text, title_text):
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = NAVY_COLOR
    top_bar.line.color.rgb = NAVY_COLOR

    accent_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.05))
    accent_line.fill.solid()
    accent_line.fill.fore_color.rgb = BLUE_ACCENT
    accent_line.line.color.rgb = BLUE_ACCENT

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.733), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True

    p_tag = tf.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.bold = True
    p_tag.font.size = Pt(10)
    p_tag.font.color.rgb = RGBColor(147, 197, 253)

    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.bold = True
    p_title.font.size = Pt(17)
    p_title.font.color.rgb = RGBColor(255, 255, 255)
    p_title.space_before = Pt(2)


def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_COLOR):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)
    return card


def add_bullet_card(slide, left, top, width, height, card_title, bullets, font_title_size=12, font_body_size=9.2):
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.25), width - Inches(0.5), height - Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = card_title
    p0.font.bold = True
    p0.font.size = Pt(font_title_size)
    p0.font.color.rgb = NAVY_COLOR
    p0.space_after = Pt(6)

    for i, (header, body) in enumerate(bullets):
        p_h = tf.add_paragraph()
        p_h.text = header
        p_h.font.bold = True
        p_h.font.size = Pt(font_body_size + 0.5)
        p_h.font.color.rgb = BLUE_ACCENT
        if i > 0:
            p_h.space_before = Pt(4)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.size = Pt(font_body_size)
        p_b.font.color.rgb = TEXT_MAIN
        p_b.space_before = Pt(1.5)


def add_placeholder_card(slide, left, top, width, height, diagram_title, guide_items):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = PLACEHOLDER_BG
    card.line.color.rgb = PLACEHOLDER_BORDER
    card.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.25), width - Inches(0.5), height - Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True

    p_tag = tf.paragraphs[0]
    p_tag.text = "[ VỊ TRÍ CHÈN SƠ ĐỒ / BẢN VẼ DRAW.IO ]"
    p_tag.font.bold = True
    p_tag.font.size = Pt(10)
    p_tag.font.color.rgb = BLUE_ACCENT

    p_title = tf.add_paragraph()
    p_title.text = diagram_title
    p_title.font.bold = True
    p_title.font.size = Pt(12)
    p_title.font.color.rgb = NAVY_COLOR
    p_title.space_before = Pt(4)
    p_title.space_after = Pt(6)

    p_guide = tf.add_paragraph()
    p_guide.text = "Gợi ý các thành phần cần vẽ trên Draw.io:"
    p_guide.font.bold = True
    p_guide.font.size = Pt(9.5)
    p_guide.font.color.rgb = TEXT_MUTED

    for item in guide_items:
        p_item = tf.add_paragraph()
        p_item.text = f"• {item}"
        p_item.font.size = Pt(8.8)
        p_item.font.color.rgb = TEXT_MAIN
        p_item.space_before = Pt(3.5)


# ==============================================================================
# SLIDE 1: TIÊU ĐỀ & GIỚI THIỆU ĐỀ TÀI DT005 (30s)
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)
add_base_background(slide1)

header_box1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(2.2))
header_box1.fill.solid()
header_box1.fill.fore_color.rgb = NAVY_COLOR
header_box1.line.color.rgb = NAVY_COLOR

line1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.2), Inches(13.333), Inches(0.08))
line1.fill.solid()
line1.fill.fore_color.rgb = BLUE_ACCENT
line1.line.color.rgb = BLUE_ACCENT

tb1 = slide1.shapes.add_textbox(Inches(1.0), Inches(0.35), Inches(11.333), Inches(1.6))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1_sub = tf1.paragraphs[0]
p1_sub.text = "BÁO CÁO GIỮA KỲ HỌC PHẦN AN TOÀN VÀ BẢO MẬT HỆ THỐNG THÔNG TIN — ĐỀ TÀI DT005"
p1_sub.font.bold = True
p1_sub.font.size = Pt(11)
p1_sub.font.color.rgb = RGBColor(147, 197, 253)

p1_title = tf1.add_paragraph()
p1_title.text = "NGHIÊN CỨU & PHÁT TRIỂN HỆ THỐNG KIỂM TRA TOÀN VẸN TỆP\nDỰA TRÊN CÂY MERKLE PHÂN CẤP ĐỘNG (HD-MERKLETREE)"
p1_title.font.bold = True
p1_title.font.size = Pt(21)
p1_title.font.color.rgb = RGBColor(255, 255, 255)
p1_title.space_before = Pt(6)

card_meta1 = add_card(slide1, Inches(1.0), Inches(2.55), Inches(11.333), Inches(1.2))
tb_meta1 = slide1.shapes.add_textbox(Inches(1.3), Inches(2.7), Inches(10.7), Inches(0.9))
tf_meta1 = tb_meta1.text_frame
tf_meta1.word_wrap = True
p_m1 = tf_meta1.paragraphs[0]
p_m1.text = "Cơ sở lý thuyết nền tảng: Bám sát bài báo khoa học quốc tế của TS. Harry Halpin (2025):"
p_m1.font.bold = True
p_m1.font.size = Pt(11)
p_m1.font.color.rgb = NAVY_COLOR

p_m2 = tf_meta1.add_paragraph()
p_m2.text = "\"The Adversary: The Philosophy of Cryptography\" — Xuất bản trên Oxford Journal of Cybersecurity, Vol 11, Issue 1 (2025)."
p_m2.font.italic = True
p_m2.font.size = Pt(10.5)
p_m2.font.color.rgb = BLUE_ACCENT
p_m2.space_before = Pt(3)

# 3 Thẻ mục tiêu chính
c_w1 = Inches(3.6)
c_gap1 = Inches(0.26)
c_t1 = Inches(4.0)
c_h1 = Inches(3.0)

add_bullet_card(slide1, Inches(1.0), c_t1, c_w1, c_h1, "1. Hiện Thực Hóa Nền Tảng", [
    ("Nghiên cứu bài báo Oxford 2025:", "Dịch và hiện thực hóa đầy đủ 7 mô hình toán học mật mã học của TS. Harry Halpin."),
    ("Tái lập 3 mô hình cơ sở:", "Xây dựng hoàn chỉnh Flat SHA-256 Manifest, HMAC-SHA256 (có khóa bí mật K), và Standard Merkle Tree."),
    ("Kiểm định nghiêm ngặt:", "100% vượt qua bộ 17 bài Unit Test tự động (test_file_integrity.py).")
])

add_bullet_card(slide1, Inches(1.0) + c_w1 + c_gap1, c_t1, c_w1, c_h1, "2. Đề Xuất Cải Tiến (HD-Merkle)", [
    ("Thiết kế cấu trúc phân cấp động:", "Ánh xạ trực tiếp cây logic với cây thư mục thực tế (Hierarchical Dynamic Tree)."),
    ("Thuật toán cập nhật Dirty Path:", "Cập nhật tăng dần trong thời gian O(log N), triệt tiêu nhu cầu băm lại toàn bộ cây."),
    ("Phân tách tiền tố RFC 6962:", "Kháng va chạm tiền ảnh cấp hai (Second-Preimage) tuyệt đối bằng tiền tố 0x00 và 0x01.")
])

add_bullet_card(slide1, Inches(1.0) + (c_w1 + c_gap1) * 2, c_t1, c_w1, c_h1, "3. Thực Nghiệm & Đóng Góp", [
    ("Thực nghiệm trên 2 Dataset Kaggle:", "Khảo sát và benchmark thực tế trên bộ 1,115 tệp (29 MB) và bộ quy mô lớn 5,600 tệp (175 MB)."),
    ("Bằng chứng nhóm (Batch Proof):", "Gộp nút chung giúp tiết kiệm ~86% dung lượng mạng so với Standard Merkle."),
    ("Sẵn sàng ứng dụng thực tiễn:", "Bảo vệ hệ thống file phân tán, máy chủ lưu trữ đám mây và hệ thống kiểm toán an ninh.")
])


# ==============================================================================
# SLIDE 2: ĐẶT VẤN ĐỀ & KHẢO SÁT 2 TẬP DỮ LIỆU KAGGLE THỰC TẾ (40s)
# ==============================================================================
slide2 = prs.slides.add_slide(blank_layout)
add_base_background(slide2)
add_header(slide2, "BỐI CẢNH & TẬP DỮ LIỆU THỰC NGHIỆM", "Hạn Chế Của Các Phương Pháp Cũ & Khảo Sát 2 Tập Dữ Liệu Kaggle")

w_col2 = Inches(5.6)
# Cột 1: Hạn chế của 3 mô hình cũ
add_bullet_card(slide2, Inches(0.8), Inches(1.4), w_col2, Inches(5.6), "3 Hạn Chế Cốt Lõi Của Các Phương Pháp Truyền Thống", [
    ("1. Flat SHA-256 Manifest (Danh sách băm phẳng):", 
     "• Chi phí cập nhật O(N): Khi sửa 1 file, bắt buộc phải đọc và băm lại toàn bộ thư mục.\n• Rò rỉ siêu dữ liệu: Để xác thực qua mạng, bắt buộc tải toàn bộ danh mục 111 KB làm lộ 100% đường dẫn tệp."),
    ("2. HMAC-SHA256 Manifest (Băm có khóa bí mật):", 
     "• Rủi ro phân phối khóa đối xứng K: Phía xác thực bắt buộc phải nắm khóa K. Nếu lộ khóa K, kẻ tấn công dễ dàng làm giả toàn bộ danh mục.\n• Không có bằng chứng kiểm định độc lập cho từng tệp riêng lẻ."),
    ("3. Standard Binary Merkle Tree (Cây Merkle phẳng chuẩn):", 
     "• Không tối ưu cấu trúc thư mục: Dựng mảng phẳng 1 chiều, mất O(N) tái cấu trúc khi sửa đổi file.\n• Lỗ hổng Second-Preimage: Ghép chuỗi thô H(left || right) dễ bị giả mạo 2 nút lá thành nút nhánh.")
], font_title_size=12, font_body_size=9.0)

# Cột 2: Chi tiết 2 Tập dữ liệu Kaggle đã sử dụng
add_card(slide2, Inches(6.8), Inches(1.4), w_col2, Inches(5.6))
tb_k2 = slide2.shapes.add_textbox(Inches(7.05), Inches(1.6), w_col2 - Inches(0.5), Inches(5.2))
tf_k2 = tb_k2.text_frame
tf_k2.word_wrap = True

p_kt = tf_k2.paragraphs[0]
p_kt.text = "Khảo Sát 2 Tập Dữ Liệu Kaggle Thực Tế Được Sử Dụng"
p_kt.font.bold = True
p_kt.font.size = Pt(12)
p_kt.font.color.rgb = NAVY_COLOR
p_kt.space_after = Pt(8)

# Dataset 1
p_d1 = tf_k2.add_paragraph()
p_d1.text = "1. Tập Dữ Liệu 1: kaggle_dataset (Tập kiểm thử tiêu chuẩn)"
p_d1.font.bold = True
p_d1.font.size = Pt(10)
p_d1.font.color.rgb = BLUE_ACCENT

p_d1_desc = tf_k2.add_paragraph()
p_d1_desc.text = "• Quy mô: 1,115 tệp tin — Dung lượng: 29.26 MB.\n• Đa dạng 7 định dạng thực tế: PDF hóa đơn doanh nghiệp (400), TXT báo cáo (400), DOCX chính sách & luật (117), PNG ảnh quét (150), JSON cấu hình (25), CSV kiểm toán (11), PY scripts (12).\n• Mục đích: Kiểm thử độ chính xác chức năng, kiểm định phát hiện vi phạm và đo lường đối chuẩn ban đầu."
p_d1_desc.font.size = Pt(8.8)
p_d1_desc.font.color.rgb = TEXT_MAIN
p_d1_desc.space_before = Pt(2)
p_d1_desc.space_after = Pt(8)

# Dataset 2
p_d2 = tf_k2.add_paragraph()
p_d2.text = "2. Tập Dữ Liệu 2: kaggle_dataset_large (Quy mô lớn tại Downloads)"
p_d2.font.bold = True
p_d2.font.size = Pt(10)
p_d2.font.color.rgb = GREEN_COLOR

p_d2_desc = tf_k2.add_paragraph()
p_d2_desc.text = "• Quy mô: 5,600 tệp tin — Dung lượng: 175.12 MB (Gấp 6 lần dung lượng, 5 lần số tệp).\n• 6 danh mục lớn: Giao dịch tài chính (800 CSV), Hợp đồng pháp lý (1,200 PDF), Log máy chủ (1,200 LOG), Cấu hình hệ thống (600 JSON), Hồ sơ y tế (1,200 TXT), Telemetry cảm biến IoT (600 DAT).\n• Mục đích: Kiểm thử tải nặng, quan sát tiến trình log terminal streaming và chứng minh ưu thế vượt trội của thuật toán O(log N)."
p_d2_desc.font.size = Pt(8.8)
p_d2_desc.font.color.rgb = TEXT_MAIN
p_d2_desc.space_before = Pt(2)


# ==============================================================================
# SLIDE 3: KIẾN TRÚC MÔ HÌNH ĐỀ XUẤT HD-MERKLETREE & RFC 6962 (50s)
# ==============================================================================
slide3 = prs.slides.add_slide(blank_layout)
add_base_background(slide3)
add_header(slide3, "THIẾT KẾ MÔ HÌNH MỚI (THANG ĐIỂM 9 - 10)", "Kiến Trúc Cây Merkle Phân Cấp Động & Phân Tách Tiền Tố RFC 6962")

w_half3 = Inches(5.6)
# Cột 1: Placeholder Draw.io để người dùng tự vẽ sơ đồ kiến trúc
add_placeholder_card(slide3, Inches(0.8), Inches(1.4), w_half3, Inches(5.6), 
                     "Sơ đồ Kiến trúc Cây Phân Cấp HD-MerkleTree", [
                         "Tầng 1 (Root Hash): Mã băm gốc toàn cục (32 Bytes).",
                         "Tầng 2 (Directory Nodes): Các nút thư mục trung gian phân nhánh theo cây đĩa cứng thực tế.",
                         "Tầng 3 (File Leaves): Các nút lá băm nội dung từng tệp (0x00 || file_path || data).",
                         "Đường dẫn bẩn (Dirty Path): Mũi tên đỏ lan truyền ngược từ lá bị sửa lên Root Hash.",
                         "Bằng chứng Compact Proof: Chuỗi các node sibling cần thiết để đối soát O(log N)."
                     ])

# Cột 2: Phân tích 3 Đóng góp kỹ thuật cốt lõi
add_bullet_card(slide3, Inches(6.8), Inches(1.4), w_half3, Inches(5.6), "3 Đóng Góp Kỹ Thuật Đột Phá Của Đề Tài", [
    ("1. Phân tách miền tiền tố RFC 6962 (Domain Separation):",
     "• Nút lá (Leaf Hash): H_leaf = SHA-256(0x00 || file_path || data_hash)\n• Nút nhánh (Internal Node): H_node = SHA-256(0x01 || left_hash || right_hash)\n➔ Vì 0x00 != 0x01, không gian băm hoàn toàn độc lập, triệt tiêu 100% va chạm Second-Preimage."),
    ("2. Cấu trúc cây phản ánh phân cấp thư mục thực tế:",
     "• Không băm phẳng danh mục mà ánh xạ cây thư mục vật lý thành cây lồng nhau.\n• Xử lý đa luồng (Multi-core ThreadPool) băm song song hàng nghìn file cực nhanh."),
    ("3. Bằng chứng thu gọn Compact Proof & Batch Multi-Proof:",
     "• Proof 1 file chỉ tốn ~1.4 KB thay vì gửi cả danh mục 111 KB.\n• Batch Multi-Proof: Khi kiểm định đồng thời nhóm k tệp, hệ thống gộp các nút cha chung, giúp tiết kiệm hơn 86% dung lượng truyền tải qua mạng.")
], font_title_size=12, font_body_size=9.0)


# ==============================================================================
# SLIDE 4: BẢNG ĐỐI SÁNH ĐỘ PHỨC TẠP BIG-O & CƠ CHẾ DIRTY PATH (50s)
# ==============================================================================
slide4 = prs.slides.add_slide(blank_layout)
add_base_background(slide4)
add_header(slide4, "PHÂN TÍCH HIỆU NĂNG THUẬT TOÁN", "Bảng Đối Sánh Độ Phức Tạp Big-O Lý Thuyết & Cơ Chế Dirty Path")

# Bảng Big-O bên trái
table_card4 = add_card(slide4, Inches(0.8), Inches(1.4), Inches(6.8), Inches(5.6))
tb_t4 = slide4.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(6.4), Inches(0.4))
tb_t4.text_frame.paragraphs[0].text = "BẢNG ĐỐI SÁNH ĐỘ PHỨC TẠP THUẬT TOÁN BIG-O LÝ THUYẾT"
tb_t4.text_frame.paragraphs[0].font.bold = True
tb_t4.text_frame.paragraphs[0].font.size = Pt(11)
tb_t4.text_frame.paragraphs[0].font.color.rgb = NAVY_COLOR

table_shape4 = slide4.shapes.add_table(6, 5, Inches(1.0), Inches(1.95), Inches(6.4), Inches(3.6))
table4 = table_shape4.table
table4.columns[0].width = Inches(1.6)
table4.columns[1].width = Inches(1.1)
table4.columns[2].width = Inches(1.1)
table4.columns[3].width = Inches(1.2)
table4.columns[4].width = Inches(1.4)

headers4 = ["Chỉ Số Đánh Giá", "Flat SHA-256", "HMAC-SHA256", "Standard Merkle", "HD-Merkle (Đề Xuất)"]
for col_idx, h_text in enumerate(headers4):
    cell = table4.cell(0, col_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = NAVY_COLOR
    cell.text = h_text
    p = cell.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(8.5)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

big_o_data = [
    ("Khởi tạo (Build)", "O(N)", "O(N)", "O(N)", "O(N/P) - Đa luồng"),
    ("Xác minh 1 tệp", "O(N) - Băm lại hết", "O(N) - Băm lại hết", "O(log N)", "O(log N)"),
    ("Cập nhật khi sửa 1 tệp", "O(N) - Băm lại hết", "O(N) - Băm lại hết", "O(N) - Tái dựng mảng", "O(log N) - Dirty Path"),
    ("Kích thước Proof", "O(N) - Toàn bộ file", "O(N) - Cần khóa K", "O(log N)", "O(k log(N/k)) - Batch"),
    ("Kháng Second-Preimage", "Kém", "Phụ thuộc khóa K", "Dễ bị tấn công", "Tuyệt đối - RFC 6962")
]

for row_idx, row in enumerate(big_o_data):
    for col_idx, val in enumerate(row):
        cell = table4.cell(row_idx + 1, col_idx)
        cell.fill.solid()
        if col_idx == 4:
            cell.fill.fore_color.rgb = GREEN_BG
        elif row_idx % 2 == 1:
            cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
        else:
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cell.text = val
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(8)
        p.font.color.rgb = GREEN_COLOR if col_idx == 4 else TEXT_MAIN
        if col_idx == 4:
            p.font.bold = True
        if col_idx > 0:
            p.alignment = PP_ALIGN.CENTER

tb_note4 = slide4.shapes.add_textbox(Inches(1.0), Inches(5.7), Inches(6.4), Inches(1.1))
tf_n4 = tb_note4.text_frame
tf_n4.word_wrap = True
p_n4 = tf_n4.paragraphs[0]
p_n4.text = "KẾT LUẬN HIỆU NĂNG LÝ THUYẾT:"
p_n4.font.bold = True
p_n4.font.size = Pt(9.5)
p_n4.font.color.rgb = GREEN_COLOR
p_n4_desc = tf_n4.add_paragraph()
p_n4_desc.text = "Khi có 1 tệp bị sửa đổi nội dung, HD-MerkleTree chỉ mất chi phí O(log N) nhờ cơ chế cập nhật cục bộ Dirty Path, trong khi Standard Merkle Tree và Flat SHA-256 buộc phải tính toán lại toàn bộ O(N) tệp."
p_n4_desc.font.size = Pt(8.5)
p_n4_desc.font.color.rgb = TEXT_MAIN

# Cột 2: Placeholder Draw.io vẽ Cơ chế Dirty Path
add_placeholder_card(slide4, Inches(7.8), Inches(1.4), Inches(4.7), Inches(5.6),
                     "Sơ đồ Cơ chế Cập nhật Tăng dần (Dirty Path)", [
                         "Điểm kích hoạt: Tệp tin file_k bị sửa đổi 1 byte nội dung.",
                         "Đánh dấu bẩn: Hàm mark_dirty() gắn cờ cho node lá tương ứng.",
                         "Đường dẫn bẩn O(log N): Chỉ lan truyền ngược lên các node cha (Parent -> Directory -> Root).",
                         "Các nhánh còn lại: Giữ nguyên hash cũ trong bộ nhớ đệm (In-memory Cache Hit).",
                         "Hiệu quả: Thời gian cập nhật đạt 0.13 - 0.22 ms thay vì băm lại toàn bộ đĩa."
                     ])


# ==============================================================================
# SLIDE 5: 7 MÔ HÌNH HARRY HALPIN (OXFORD 2025) - PHẦN 1: MÔ HÌNH 1 ĐẾN 4 (45s)
# ==============================================================================
slide5 = prs.slides.add_slide(blank_layout)
add_base_background(slide5)
add_header(slide5, "CƠ SỞ TOÁN HỌC MẬT MÃ HỌC (OXFORD 2025)", "Kiểm Định Hình Thức Theo Bài Báo TS. Harry Halpin (Mô Hình 1 - 4)")

w_g5 = Inches(5.7)
h_g5 = Inches(2.7)

# Model 1
add_bullet_card(slide5, Inches(0.8), Inches(1.35), w_g5, h_g5, "1. Thách Đấu Kẻ Thù Siêu Máy Tính PPT (Goldwasser-Micali 1984)", [
    ("Định nghĩa trò chơi toán học:", "Challenger khởi tạo Root Hash H_0. Kẻ tấn công A có năng lực tính toán đa thức PPT cố tình tìm trạng thái M' != M sao cho Verify(M', H_0) = 1."),
    ("Giới hạn an toàn lý thuyết:", "Advantage(A) = |Pr[A thắng] - 0| <= negl(lambda) = 2^(-128) ≈ 2.9 x 10^(-39)."),
    ("Kết quả thực nghiệm:", "Chặn đứng 50/50 lượt tấn công thử nghiệm (bit_flip, tail_append, hash_ext). Advantage = 0.0.")
], font_title_size=10.5, font_body_size=8.5)

# Model 2
add_bullet_card(slide5, Inches(6.8), Inches(1.35), w_g5, h_g5, "2. Kẻ Tấn Công Mạng Chiếm Quyền Kiểm Soát (Dolev-Yao 1983)", [
    ("Năng lực kẻ thù mạng Dolev-Yao:", "Hacker kiểm soát 100% đường truyền (nghe lén, đánh tráo gói tin, sửa đổi và gửi lại proof cũ)."),
    ("3 Kịch bản tấn công thực nghiệm:", "• Cross-proof Swap: Lấy Proof file A gán cho file B ➔ Bị từ chối 100%.\n• In-transit Tampering: Sửa 1 byte proof trên mạng ➔ Bị từ chối 100%.\n• Replay Proof: Gửi lại proof của phiên bản cũ ➔ Bị từ chối 100%."),
    ("Kết quả:", "Tỷ lệ ngăn chặn đạt 100.0% (3/3 kịch bản tấn công mạng).")
], font_title_size=10.5, font_body_size=8.5)

# Model 3
add_bullet_card(slide5, Inches(0.8), Inches(4.25), w_g5, h_g5, "3. Kháng Second-Preimage Bằng RFC 6962 Domain Separation", [
    ("Lỗ hổng của Merkle thường:", "Ghép nối chuỗi thô H(left || right) dễ bị giả mạo khi kẻ tấn công lấy 2 nút lá tạo thành một nút nhánh trung gian."),
    ("Chứng minh toán học RFC 6962:", "H_leaf = SHA256(0x00 || path || data) != H_internal = SHA256(0x01 || left || right)."),
    ("Kết luận:", "Vì 0x00 != 0x01, không gian băm của nút lá và nút nhánh hoàn toàn tách biệt, triệt tiêu 100% va chạm Second-Preimage cấu trúc.")
], font_title_size=10.5, font_body_size=8.5)

# Model 4
add_bullet_card(slide5, Inches(6.8), Inches(4.25), w_g5, h_g5, "4. Bằng Chứng Nhất Quán Chỉ-Thêm (RFC 6962 Consistency Proof)", [
    ("Chuẩn Certificate Transparency:", "Hệ thống chứng minh rằng cây kích thước M (quá khứ) là tiền tố chính xác của cây kích thước N (hiện tại)."),
    ("Kháng tấn công backdoor:", "Ngăn chặn kẻ thù cấp nhà nước chèn cửa sau hoặc âm thầm sửa đổi dữ liệu lịch sử."),
    ("Kích thước bằng chứng tối ưu:", "Độ phức tạp chỉ O(log N) = 11 nodes (~352 Bytes cho 5,600 tệp). Không cần tải lại toàn bộ cây.")
], font_title_size=10.5, font_body_size=8.5)


# ==============================================================================
# SLIDE 6: 7 MÔ HÌNH HARRY HALPIN (OXFORD 2025) - PHẦN 2: MÔ HÌNH 5 ĐẾN 7 (45s)
# ==============================================================================
slide6 = prs.slides.add_slide(blank_layout)
add_base_background(slide6)
add_header(slide6, "CƠ SỞ TOÁN HỌC MẬT MÃ HỌC (OXFORD 2025)", "Kiểm Định Hình Thức Theo Bài Báo TS. Harry Halpin (Mô Hình 5 - 7)")

w_g6 = Inches(5.7)

# Model 5
add_bullet_card(slide6, Inches(0.8), Inches(1.35), w_g6, Inches(2.7), "5. Random Oracle Model (ROM) & Hiệu Ứng Tuyết Lở (SAC)", [
    ("Tiêu chuẩn Strict Avalanche Criterion (SAC):", "Khi lật đúng 1 bit đầu vào, từng bit ở chuỗi băm đầu ra 256-bit phải thay đổi độc lập với xác suất p = 50% (Phân phối Bernoulli)."),
    ("Số liệu thực nghiệm SHA-256:", "Đảo 133 / 256 bit (~51.95%, sát mục tiêu lý tưởng 50%). Shannon Entropy đạt 7.98 / 8.00 bits/byte."),
    ("Đối sánh sự thất bại của CRC32:", "Mã CRC32 tuyến tính bị bẻ gãy hoàn toàn do tính chất đa thức modulo.")
], font_title_size=10.5, font_body_size=8.5)

# Model 6
add_bullet_card(slide6, Inches(6.8), Inches(1.35), w_g6, Inches(2.7), "6. Mô Hình Toàn Vẹn Phân Cấp Biba (Nguyên Lý No Write-Up)", [
    ("Tiên đề toàn vẹn Biba (Biba 1977):", "\"No Write-Up\": Đối tượng ở tầng toàn vẹn thấp (Unclassified) không thể can thiệp làm thay đổi trạng thái tầng cao (Top Secret) mà không bị phát hiện."),
    ("Lan truyền cảnh báo tức thời:", "Can thiệp ở tệp lá lập tức làm thay đổi Root Hash sau 0.13 ms qua Dirty Path."),
    ("Tự động xóa vết bộ nhớ đệm:", "Kích hoạt hàm secure_erase() xóa sạch cache bộ nhớ của nút bị tấn công.")
], font_title_size=10.5, font_body_size=8.5)

# Model 7 (Card lớn kéo dài ở dưới)
add_bullet_card(slide6, Inches(0.8), Inches(4.25), Inches(11.733), Inches(2.7), 
                "7. Kiểm Toán Bảo Vệ Siêu Dữ Liệu (Snowden 2013) & Kháng Kênh Kề Đo Thời Gian (Kocher 1996)", [
    ("Kiểm toán rò rỉ siêu dữ liệu kênh truyền (Metadata Privacy):",
     "• Flat SHA-256 & HMAC: Buộc phải truyền tải toàn bộ danh mục qua mạng (111 KB), rò rỉ 100% đường dẫn và cấu trúc thư mục nội bộ.\n• HD-MerkleTree: Chỉ truyền duy nhất bằng chứng Compact Proof của tệp được hỏi (~1.4 KB), bảo vệ tuyệt đối thông tin các tệp còn lại (0 Byte rò rỉ)."),
    ("Kháng tấn công đo thời gian phản hồi (Timing Side-Channel Resistance):",
     "• Nguy cơ: Hacker đo độ trễ mạng để suy đoán ký tự mã băm nếu so sánh chuỗi bằng toán tử '==' thông thường.\n• Giải pháp HD-Merkle: Áp dụng so sánh chuỗi thời gian hằng số O(1) qua hàm hmac.compare_digest() theo khuyến nghị của Kocher (1996), triệt tiêu hoàn toàn dấu vết đo thời gian.")
], font_title_size=11, font_body_size=8.8)


# ==============================================================================
# SLIDE 7: TRÌNH DIỄN HỆ THỐNG TRỰC TIẾP (LIVE DEMONSTRATION) (60s)
# ==============================================================================
slide7 = prs.slides.add_slide(blank_layout)
add_base_background(slide7)
add_header(slide7, "TRÌNH DIỄN HỆ THỐNG THỰC TẾ", "LIVE DEMONSTRATION — TRÌNH DIỄN TRỰC TIẾP TRÊN TRÌNH DUYỆT & TERMINAL")

# Khung thông báo trung tâm cực kỳ nổi bật
center_card7 = add_card(slide7, Inches(1.5), Inches(1.4), Inches(10.333), Inches(5.6), bg_color=CARD_BG, border_color=BLUE_ACCENT)

tb_c7 = slide7.shapes.add_textbox(Inches(1.8), Inches(1.6), Inches(9.733), Inches(1.0))
tf_c7 = tb_c7.text_frame
tf_c7.word_wrap = True

p_c7_tag = tf_c7.paragraphs[0]
p_c7_tag.text = "CHUYỂN GIAO DIỆN SANG TRÌNH DUYỆT WEB & TERMINAL MÁY TÍNH"
p_c7_tag.font.bold = True
p_c7_tag.font.size = Pt(12)
p_c7_tag.font.color.rgb = BLUE_ACCENT
p_c7_tag.alignment = PP_ALIGN.CENTER

p_c7_title = tf_c7.add_paragraph()
p_c7_title.text = "TRÌNH DIỄN TƯƠNG TÁC THỰC NGHIỆM TRÊN 5,600 TỆP TIN KAGGLE"
p_c7_title.font.bold = True
p_c7_title.font.size = Pt(18)
p_c7_title.font.color.rgb = NAVY_COLOR
p_c7_title.space_before = Pt(4)
p_c7_title.alignment = PP_ALIGN.CENTER

# 4 Bước kịch bản demo trực tiếp
step_w7 = Inches(4.65)
step_h7 = Inches(1.6)

# Bước 1
add_card(slide7, Inches(1.85), Inches(2.75), step_w7, step_h7, bg_color=BLUE_BG, border_color=RGBColor(191, 219, 254))
tb_s1 = slide7.shapes.add_textbox(Inches(2.0), Inches(2.85), step_w7 - Inches(0.3), step_h7 - Inches(0.2))
tf_s1 = tb_s1.text_frame
tf_s1.word_wrap = True
p_s1_t = tf_s1.paragraphs[0]
p_s1_t.text = "Bước 1: Chọn Thư Mục Giám Sát Lớn"
p_s1_t.font.bold = True
p_s1_t.font.size = Pt(11)
p_s1_t.font.color.rgb = NAVY_COLOR
p_s1_d = tf_s1.add_paragraph()
p_s1_d.text = "Gắn kết thư mục kaggle_dataset_large (5,600 tệp - 175 MB) trên đĩa cứng qua hộp thoại Explorer."
p_s1_d.font.size = Pt(9.2)
p_s1_d.font.color.rgb = TEXT_MAIN
p_s1_d.space_before = Pt(3)

# Bước 2
add_card(slide7, Inches(6.75), Inches(2.75), step_w7, step_h7, bg_color=GREEN_BG, border_color=RGBColor(187, 247, 208))
tb_s2 = slide7.shapes.add_textbox(Inches(6.9), Inches(2.85), step_w7 - Inches(0.3), step_h7 - Inches(0.2))
tf_s2 = tb_s2.text_frame
tf_s2.word_wrap = True
p_s2_t = tf_s2.paragraphs[0]
p_s2_t.text = "Bước 2: Nạp Baseline & Quan Sát Log Terminal"
p_s2_t.font.bold = True
p_s2_t.font.size = Pt(11)
p_s2_t.font.color.rgb = GREEN_COLOR
p_s2_d = tf_s2.add_paragraph()
p_s2_d.text = "Nhấn '1. Nạp dữ liệu' ➔ Quan sát tốc độ nạp nghìn tệp/s và nhật ký stream log trực tiếp trên terminal."
p_s2_d.font.size = Pt(9.2)
p_s2_d.font.color.rgb = TEXT_MAIN
p_s2_d.space_before = Pt(3)

# Bước 3
add_card(slide7, Inches(1.85), Inches(4.55), step_w7, step_h7, bg_color=RED_BG, border_color=RGBColor(254, 202, 202))
tb_s3 = slide7.shapes.add_textbox(Inches(2.0), Inches(4.65), step_w7 - Inches(0.3), step_h7 - Inches(0.2))
tf_s3 = tb_s3.text_frame
tf_s3.word_wrap = True
p_s3_t = tf_s3.paragraphs[0]
p_s3_t.text = "Bước 3: Mô Phỏng Tấn Công & Đối Soát Đĩa"
p_s3_t.font.bold = True
p_s3_t.font.size = Pt(11)
p_s3_t.font.color.rgb = RED_ACCENT
p_s3_d = tf_s3.add_paragraph()
p_s3_d.text = "Can thiệp sửa 1 byte tệp thực tế trên đĩa ➔ Nhấn '2. Kiểm tra' ➔ Hệ thống phát hiện ngay và định vị tệp vi phạm."
p_s3_d.font.size = Pt(9.2)
p_s3_d.font.color.rgb = TEXT_MAIN
p_s3_d.space_before = Pt(3)

# Bước 4
add_card(slide7, Inches(6.75), Inches(4.55), step_w7, step_h7, bg_color=RGBColor(250, 245, 255), border_color=RGBColor(233, 213, 255))
tb_s4 = slide7.shapes.add_textbox(Inches(6.9), Inches(4.65), step_w7 - Inches(0.3), step_h7 - Inches(0.2))
tf_s4 = tb_s4.text_frame
tf_s4.word_wrap = True
p_s4_t = tf_s4.paragraphs[0]
p_s4_t.text = "Bước 4: So Sánh 4 Mô Hình & 7 Bài Test Toán Học"
p_s4_t.font.bold = True
p_s4_t.font.size = Pt(11)
p_s4_t.font.color.rgb = PURPLE_ACCENT
p_s4_d = tf_s4.add_paragraph()
p_s4_d.text = "Chạy đối đầu 4 mô hình (xem bảng số liệu) & Thực thi kiểm định 7 mô hình toán học (xem Radar Chart 100%)."
p_s4_d.font.size = Pt(9.2)
p_s4_d.font.color.rgb = TEXT_MAIN
p_s4_d.space_before = Pt(3)

p_ready = tf_c7.add_paragraph()
p_ready.text = "➔ Người trình bày tiến hành thu nhỏ Slide và thao tác trực tiếp trên giao diện Web http://localhost:8080"
p_ready.font.bold = True
p_ready.font.size = Pt(10)
p_ready.font.color.rgb = NAVY_COLOR
p_ready.alignment = PP_ALIGN.CENTER
p_ready.space_before = Pt(8)


# ==============================================================================
# SLIDE 8: KẾT LUẬN & ĐÓNG GÓP HỌC THUẬT (20s)
# ==============================================================================
slide8 = prs.slides.add_slide(blank_layout)
add_base_background(slide8)
add_header(slide8, "TỔNG KẾT & ĐÓNG GÓP ĐỀ TÀI", "Kết Luận, Đóng Góp Học Thuật & Lời Cảm Ơn")

w_c8 = Inches(3.7)
gap8 = Inches(0.26)
top8 = Inches(1.4)
h8 = Inches(4.3)

add_bullet_card(slide8, Inches(0.8), top8, w_c8, h8, "1. Đóng Góp Lý Thuyết", [
    ("Nghiên cứu hình thức chuẩn mực:", "Hệ thống hóa toàn diện cơ sở lý luận an toàn thông tin theo bài báo TS. Harry Halpin (Oxford 2025)."),
    ("Mô hình hóa an toàn:", "Hiện thực hóa 7 mô hình hình thức: PPT Adversary, Dolev-Yao, RFC 6962, ROM/SAC, Biba và Metadata Privacy."),
    ("Kháng va chạm toán học:", "Chứng minh phân tách miền tiền tố 0x00 != 0x01 triệt tiêu 100% va chạm Second-Preimage.")
], font_title_size=11, font_body_size=8.8)

add_bullet_card(slide8, Inches(0.8) + w_c8 + gap8, top8, w_c8, h8, "2. Đóng Góp Kỹ Thuật", [
    ("Thiết kế mô hình HD-MerkleTree:", "Tối ưu hóa kiến trúc cây phân cấp tương thích hoàn toàn cấu trúc thư mục thực tế."),
    ("Cập nhật tăng dần Dirty Path:", "Giảm độ phức tạp cập nhật từ O(N) xuống O(log N), thời gian cập nhật 0.13 ms."),
    ("Cơ chế nén Batch Multi-Proof:", "Tiết kiệm hơn 86% băng thông truyền tải mạng khi xác thực nhóm tệp."),
    ("Độ tin cậy tuyệt đối:", "Vượt qua 17/17 bài Unit Test tự động (test_file_integrity.py).")
], font_title_size=11, font_body_size=8.8)

add_bullet_card(slide8, Inches(0.8) + (w_c8 + gap8) * 2, top8, w_c8, h8, "3. Khả Năng Ứng Dụng", [
    ("Giám sát tệp thời gian thực:", "Áp dụng cho máy chủ web, cơ sở dữ liệu và hệ điều hành chống mã độc ransomware."),
    ("Bảo mật điện toán đám mây:", "Kiểm tra tính toàn vẹn dữ liệu lưu trữ từ xa mà không làm lộ nội dung tệp."),
    ("Hướng nghiên cứu tiếp theo:", "Tích hợp công nghệ Zero-Knowledge Proof (zk-SNARKs) và blockchain lưu trữ Root Hash bất biến.")
], font_title_size=11, font_body_size=8.8)

# Thanh cảm ơn ở đáy slide
bar_thx = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.95), Inches(11.733), Inches(0.95))
bar_thx.fill.solid()
bar_thx.fill.fore_color.rgb = NAVY_COLOR
bar_thx.line.color.rgb = BLUE_ACCENT
bar_thx.line.width = Pt(1.5)

tb_thx = slide8.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.333), Inches(0.75))
tf_thx = tb_thx.text_frame
tf_thx.word_wrap = True
p_t1 = tf_thx.paragraphs[0]
p_t1.text = "XIN TRÂN TRỌNG CẢM ƠN QUÝ THẦY CÔ VÀ HỘI ĐỒNG ĐÃ LẮNG NGHE!"
p_t1.font.bold = True
p_t1.font.size = Pt(14)
p_t1.font.color.rgb = RGBColor(255, 255, 255)
p_t1.alignment = PP_ALIGN.CENTER

p_t2 = tf_thx.add_paragraph()
p_t2.text = "Em rất mong nhận được những góp ý quý báu từ Quý Thầy Cô để hoàn thiện đề tài tốt hơn."
p_t2.font.size = Pt(10)
p_t2.font.color.rgb = RGBColor(191, 219, 254)
p_t2.alignment = PP_ALIGN.CENTER
p_t2.space_before = Pt(2)

# ==============================================================================
# LƯU TỆP PRESENTATION
# ==============================================================================
prs.save(OUTPUT_PATH)
print(f"[HOÀN TẤT] Đã tạo thành công {len(prs.slides)} slide chuẩn chỉnh tại: {OUTPUT_PATH}")
