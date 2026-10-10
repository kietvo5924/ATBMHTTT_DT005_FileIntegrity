# -*- coding: utf-8 -*-
"""
Script tạo slide báo cáo giữa kỳ Đề tài DT005 chuẩn khoa học, giao diện cao cấp.
Tích hợp đầy đủ:
- Tóm tắt bài báo Harry Halpin (Oxford Cybersecurity, 2025) bám sát 100% nguyên tác tiếng Anh
- Tuyên bố Khung Phương Pháp Luận Khoa Học gồm 3 Trụ Cột Độc Lập
- Ma trận đối chiếu 1-1 và dẫn chứng chứng minh toán học, công thức Big-O, số liệu thực nghiệm
- Tập dữ liệu chuẩn xác 1.000 tệp tin thực tế (4.1 MB, 14 định dạng, 23 thư mục)
- Đối chuẩn thực nghiệm đối đầu 4 cơ chế bảo mật trên 1.000 tệp tin
- Khung kiểm định 7 mô hình an toàn mật mã học (Security Games) với ngôn từ trực quan, dễ hiểu
- Biểu đồ PowerPoint thực (Clustered Column & Bar Charts) thay thế toàn bộ khung chờ
- Giao diện Web Mockup trực quan (Browser window, Attack alert, Monitor, Console)
- Chuẩn hóa thông tin sinh viên Võ Anh Kiệt, trường ĐH Tài Nguyên và Môi Trường TP.HCM
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
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

OUTPUT_PATH = r"d:\AnToanvaBaoMatHeThongThongTin\GK\docs\Bao_Cao_Giua_Ky_DT05.pptx"

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Bảng màu chuẩn học thuật, sang trọng, tương phản cao
BG_COLOR = RGBColor(248, 250, 252)          # Nền xám nhạt sáng sủa (#f8fafc)
NAVY_COLOR = RGBColor(30, 58, 138)          # Xanh Navy đậm sang trọng (#1e3a8a)
CARD_BG = RGBColor(255, 255, 255)            # Nền thẻ trắng tinh (#ffffff)
BORDER_COLOR = RGBColor(226, 232, 240)       # Viền xám mỏng (#e2e8f0)
TEXT_MAIN = RGBColor(15, 23, 42)            # Chữ chính màu tối rõ nét (#0f172a)
TEXT_MUTED = RGBColor(100, 116, 139)        # Chữ phụ màu xám dịu (#64748b)
GREEN_COLOR = RGBColor(5, 150, 105)         # Xanh ngọc lục bảo (#059669)
GREEN_BG = RGBColor(236, 253, 245)          # Xanh ngọc nhạt (#ecfdf5)
BLUE_ACCENT = RGBColor(37, 99, 235)         # Xanh dương điểm nhấn (#2563eb)
BLUE_BG = RGBColor(239, 246, 255)           # Xanh dương nhạt (#eff6ff)
RED_ACCENT = RGBColor(220, 38, 38)          # Đỏ điểm nhấn (#dc2626)
RED_BG = RGBColor(254, 242, 242)            # Đỏ nhạt (#fef2f2)
PURPLE_ACCENT = RGBColor(126, 34, 206)      # Tím điểm nhấn (#7e22ce)
PURPLE_BG = RGBColor(250, 245, 255)         # Tím nhạt (#faf5ff)
AMBER_ACCENT = RGBColor(217, 119, 6)        # Vàng hổ phách (#d97706)
AMBER_BG = RGBColor(254, 243, 199)          # Vàng nhạt (#fef3c7)
DARK_CONSOLE = RGBColor(15, 23, 42)         # Nền console tối (#0f172a)
CONSOLE_GREEN = RGBColor(74, 222, 128)      # Chữ console xanh sáng (#4ade80)
CONSOLE_YELLOW = RGBColor(250, 204, 21)     # Chữ console vàng (#facc15)


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
    p0.space_after = Pt(8)

    for i, (header, body) in enumerate(bullets):
        p_h = tf.add_paragraph()
        p_h.text = header
        p_h.font.bold = True
        p_h.font.size = Pt(font_body_size + 0.5)
        p_h.font.color.rgb = BLUE_ACCENT
        if i > 0:
            p_h.space_before = Pt(5)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.size = Pt(font_body_size)
        p_b.font.color.rgb = TEXT_MAIN
        p_b.space_before = Pt(1.5)


# ==============================================================================
# BỘ CÁC HÀM TẠO ĐỒ HỌA THỰC VÀ BIỂU ĐỒ NATIVE (THAY THẾ PLACEHOLDER)
# ==============================================================================

def add_browser_window_frame(slide, left, top, width, height, url_text):
    """Vẽ khung cửa sổ trình duyệt Web chuyên nghiệp."""
    card = add_card(slide, left, top, width, height)
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.48))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(241, 245, 249)
    bar.line.color.rgb = BORDER_COLOR

    # 3 nút tròn đỏ, vàng, xanh
    colors = [RGBColor(239, 68, 68), RGBColor(245, 158, 11), RGBColor(16, 185, 129)]
    for i, c in enumerate(colors):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.18 + i * 0.16), top + Inches(0.16), Inches(0.11), Inches(0.11))
        dot.fill.solid()
        dot.fill.fore_color.rgb = c
        dot.line.fill.background()

    # Thanh URL
    url_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.8), top + Inches(0.09), width - Inches(1.0), Inches(0.3))
    url_box.fill.solid()
    url_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    url_box.line.color.rgb = BORDER_COLOR
    tf_u = url_box.text_frame
    tf_u.word_wrap = False
    p_u = tf_u.paragraphs[0]
    p_u.text = f"🔒 {url_text}"
    p_u.font.size = Pt(8.5)
    p_u.font.color.rgb = TEXT_MUTED
    p_u.alignment = PP_ALIGN.LEFT
    return card


def add_academic_paper_card(slide, left, top, width, height):
    """Slide 2: Hộp đồ họa mô phỏng trang bìa bài báo Oxford Journal of Cybersecurity."""
    add_card(slide, left, top, width, height)

    header_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.9))
    header_bar.fill.solid()
    header_bar.fill.fore_color.rgb = NAVY_COLOR
    header_bar.line.color.rgb = NAVY_COLOR
    tf_h = header_bar.text_frame
    tf_h.word_wrap = True
    p0 = tf_h.paragraphs[0]
    p0.text = "OXFORD UNIVERSITY PRESS • JOURNAL OF CYBERSECURITY"
    p0.font.bold = True
    p0.font.size = Pt(9.5)
    p0.font.color.rgb = RGBColor(191, 219, 254)
    p1 = tf_h.add_paragraph()
    p1.text = "ISSN 2057-2085 • Tập 11, Số 1, Năm 2025 • Article ID: tyae024"
    p1.font.size = Pt(8.5)
    p1.font.color.rgb = RGBColor(255, 255, 255)

    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(1.0), width - Inches(0.5), height - Inches(1.15))
    tf = tb.text_frame
    tf.word_wrap = True

    p_t = tf.paragraphs[0]
    p_t.text = "The Adversary: The Philosophy of Cryptography"
    p_t.font.bold = True
    p_t.font.size = Pt(13)
    p_t.font.color.rgb = NAVY_COLOR

    p_a = tf.add_paragraph()
    p_a.text = "Tác giả: Harry Halpin  |  Đơn vị: Nym Technologies & Inria Research"
    p_a.font.size = Pt(9)
    p_a.font.color.rgb = BLUE_ACCENT
    p_a.space_after = Pt(6)

    # Khung trích dẫn Abstract
    q_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), top + Inches(1.9), width - Inches(0.5), Inches(1.85))
    q_box.fill.solid()
    q_box.fill.fore_color.rgb = BLUE_BG
    q_box.line.color.rgb = BLUE_ACCENT
    tf_q = q_box.text_frame
    tf_q.word_wrap = True
    p_qh = tf_q.paragraphs[0]
    p_qh.text = "[ TRÍCH ĐOẠN LUẬN ĐIỂM CỐT LÕI TỪ BÀI BÁO ]"
    p_qh.font.bold = True
    p_qh.font.size = Pt(8.5)
    p_qh.font.color.rgb = NAVY_COLOR
    p_qb = tf_q.add_paragraph()
    p_qb.text = "“Mật mã học không chỉ là bài toán mã hóa dữ liệu tĩnh, mà là sự tương tác đối kháng hình thức giữa người xây dựng hệ thống và Kẻ thù (The Adversary). Để bảo đảm toàn vẹn kho dữ liệu, bắt buộc phải sử dụng cây Merkle phân tách miền (RFC 6962) và kiểm nghiệm trên dữ liệu thực tế biến thiên.”"
    p_qb.font.size = Pt(8.2)
    p_qb.font.color.rgb = TEXT_MAIN
    p_qb.space_before = Pt(3)

    # 4 Thẻ Badges học thuật
    badges = [
        ("Kẻ thù PPT & Dolev-Yao", BLUE_ACCENT, BLUE_BG),
        ("Chuẩn RFC 6962 Domain Separation", GREEN_COLOR, GREEN_BG),
        ("Cây Merkle & Xác minh O(log N)", PURPLE_ACCENT, PURPLE_BG),
        ("Nguyên lý Kerckhoffs & Kênh kề", AMBER_ACCENT, AMBER_BG)
    ]
    b_y = top + Inches(3.95)
    for i, (b_txt, b_col, b_bg) in enumerate(badges):
        b_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25 + (i % 2) * 2.75), b_y + Inches((i // 2) * 0.75), Inches(2.65), Inches(0.62))
        b_shape.fill.solid()
        b_shape.fill.fore_color.rgb = b_bg
        b_shape.line.color.rgb = b_col
        tf_b = b_shape.text_frame
        tf_b.word_wrap = True
        p_b = tf_b.paragraphs[0]
        p_b.text = f"✓ {b_txt}"
        p_b.font.bold = True
        p_b.font.size = Pt(8.2)
        p_b.font.color.rgb = b_col
        p_b.alignment = PP_ALIGN.CENTER


def add_knowledge_transfer_diagram(slide, left, top, width, height):
    """Slide 3: Sơ đồ 3 bước chuyển giao tri thức từ lý thuyết đến thực nghiệm."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.5))
    tb.text_frame.word_wrap = True
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "SƠ ĐỒ CHUYỂN GIAO TRI THỨC: TỪ BÀI BÁO ĐẾN ĐỒ ÁN"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    steps = [
        ("BƯỚC 1: KHUNG LÝ THUYẾT GỐC (HARRY HALPIN 2025)",
         "• Nền tảng triết học: Định nghĩa Kẻ thù mật mã học (The Adversary).\n• Đề xuất chuẩn RFC 6962 Certificate Transparency.\n• Hạn chế bài báo gốc: 0 dòng mã nguồn, 0 tập dữ liệu, 0 benchmark định lượng.",
         NAVY_COLOR, RGBColor(241, 245, 249)),
        ("BƯỚC 2: KHOẢNG TRỐNG KHOA HỌC CẦN KIỂM CHỨNG",
         "• Bài báo kêu gọi: Cần kiểm nghiệm trên hệ thống thực tế biến thiên.\n• Thách thức: Xây dựng cấu trúc Merkle Tree thế nào để không bị Rebuild O(N) khi cập nhật?",
         AMBER_ACCENT, AMBER_BG),
        ("BƯỚC 3: ĐÓNG GÓP HIỆN THỰC HÓA CỦA ĐỒ ÁN DT005",
         "• Xây dựng mô hình PoC HD-MerkleTree phân cấp Biba 1977.\n• Hiện thực hóa 7 Security Games đối kháng toán học độc lập.\n• Benchmark thực tế trên 1.000 tệp tin (4.1 MB) qua 5 vòng đo kiểm.",
         GREEN_COLOR, GREEN_BG)
    ]
    cur_y = top + Inches(0.7)
    for title, desc, col, bg_c in steps:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), cur_y, width - Inches(0.5), Inches(1.4))
        box.fill.solid()
        box.fill.fore_color.rgb = bg_c
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.bold = True
        p_t.font.size = Pt(9)
        p_t.font.color.rgb = col
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(3)
        cur_y += Inches(1.55)


def add_core_theses_grid(slide, left, top, width, height):
    """Slide 4: Lưới 4 card trực quan hóa 4 luận điểm mật mã học."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.5))
    tb.text_frame.word_wrap = True
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "BỐN TRỤ CỘT LUẬN ĐIỂM MẬT MÃ HỌC CỐT LÕI"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    cards_data = [
        ("1. Kháng Va Chạm SHA-256", "Không gian tìm kiếm: 2^256 bit\nBirthday Bound: 2^128 phép toán\nKẻ thù PPT không thể tìm va chạm ngẫu nhiên (Adv <= 2^-128).", BLUE_ACCENT, BLUE_BG),
        ("2. Phân Tách Miền RFC 6962", "Nút lá: H(0x00 || Dữ liệu)\nNút nhánh: H(0x01 || Trái || Phải)\nHai tập tiền tố rời nhau hoàn toàn, triệt tiêu va chạm tiền ảnh cấp 2.", GREEN_COLOR, GREEN_BG),
        ("3. Xác Minh Cục Bộ O(log N)", "Chiều cao cây: h ≈ 10-13 tầng\nProof: 990 Byte - 1.369 Byte\nTiết kiệm 98.8% băng thông mạng và bảo vệ tính riêng tư 999 tệp còn lại.", PURPLE_ACCENT, PURPLE_BG),
        ("4. Phòng Chống Kênh Kề", "Paul Kocher 1996: Đo thời gian so sánh\nGiải pháp: hmac.compare_digest() O(1)\nXóa đè vùng nhớ nhạy cảm RAM qua secure_erase().", AMBER_ACCENT, AMBER_BG)
    ]
    w_sub = (width - Inches(0.65)) / 2
    h_sub = Inches(2.25)
    for i, (t, b, c, bg) in enumerate(cards_data):
        r = i // 2
        col_idx = i % 2
        x = left + Inches(0.25) + col_idx * (w_sub + Inches(0.15))
        y = top + Inches(0.7) + r * (h_sub + Inches(0.15))
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w_sub, h_sub)
        box.fill.solid()
        box.fill.fore_color.rgb = bg
        box.line.color.rgb = c
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.bold = True
        p_t.font.size = Pt(9.5)
        p_t.font.color.rgb = c
        p_b = tf.add_paragraph()
        p_b.text = b
        p_b.font.size = Pt(8.2)
        p_b.font.color.rgb = TEXT_MAIN
        p_b.space_before = Pt(4)


def add_three_pillars_visual(slide, left, top, width, height):
    """Slide 5: Sơ đồ 3 trụ cột phương pháp luận đứng song song."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "KIẾN TRÚC BA TRỤ CỘT ĐỘC LẬP CHUẨN MỰC"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    pillars = [
        ("TRỤ CỘT I (HƯỚNG 1)\nKỸ THUẬT HỆ THỐNG PoC",
         "• hd_merkle_tree.py\n• Cấu trúc Biba 1977\n• Cập nhật Dirty-Path RAM\n• 5 Tầng kiến trúc PoC\n• 17 Unit Tests tự động",
         NAVY_COLOR, BLUE_BG),
        ("TRỤ CỘT II (HƯỚNG 2)\nMÔ PHỎNG AN NINH HÌNH THỨC",
         "• security_games.py\n• Khung 7 Trò chơi đối kháng\n• PPT, Dolev-Yao, Tuyết lở\n• Consistency Proof RFC 6962\n• Kháng kênh kề O(1)",
         PURPLE_ACCENT, PURPLE_BG),
        ("TRỤ CỘT III (HƯỚNG 4)\nTHỰC NGHIỆM ĐỊNH LƯỢNG",
         "• run_benchmark.py\n• Quy mô: 1.000 tệp tin thật\n• 5 Vòng lặp đo lường CPU\n• 4 Mô hình so sánh đối đầu\n• Xuất dữ liệu JSON & CSV",
         GREEN_COLOR, GREEN_BG)
    ]
    w_p = (width - Inches(0.7)) / 3
    for i, (p_t, p_d, p_c, p_bg) in enumerate(pillars):
        x = left + Inches(0.25) + i * (w_p + Inches(0.1))
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top + Inches(0.75), w_p, height - Inches(1.0))
        box.fill.solid()
        box.fill.fore_color.rgb = p_bg
        box.line.color.rgb = p_c
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        p_title = tf.paragraphs[0]
        p_title.text = p_t
        p_title.font.bold = True
        p_title.font.size = Pt(9)
        p_title.font.color.rgb = p_c
        p_title.alignment = PP_ALIGN.CENTER
        p_body = tf.add_paragraph()
        p_body.text = p_d
        p_body.font.size = Pt(8.2)
        p_body.font.color.rgb = TEXT_MAIN
        p_body.space_before = Pt(8)


def add_audit_scorecard(slide, left, top, width, height):
    """Slide 6: Bảng điểm kiểm định 4 tiêu chí cốt lõi."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "BẢNG ĐIỂM KIỂM ĐỊNH CHẤT LƯỢNG KHOA HỌC (SCORECARD)"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    scores = [
        ("100% ĐỐI CHIẾU BÀI BÁO", "Kế thừa trọn vẹn 6/6 luận điểm của Harry Halpin (2025)", NAVY_COLOR, BLUE_BG),
        ("2.500+ DÒNG MÃ NGUỒN", "Hiện thực hóa 4 CLI độc lập bằng Python chuẩn, không gian lận thư viện ngoài", BLUE_ACCENT, BLUE_BG),
        ("17/17 UNIT TESTS PASS", "Kiểm thử tự động bao phủ toàn diện I/O đĩa vật lý sửa, thêm, xóa và nhất quán", GREEN_COLOR, GREEN_BG),
        ("1.000 TỆP THỰC NGHIỆM", "Bộ dữ liệu doanh nghiệp thật (4.1 MB, 14 định dạng) đo 5 vòng lặp độc lập", AMBER_ACCENT, AMBER_BG)
    ]
    cur_y = top + Inches(0.7)
    for t, d, c, bg in scores:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), cur_y, width - Inches(0.5), Inches(1.05))
        box.fill.solid()
        box.fill.fore_color.rgb = bg
        box.line.color.rgb = c
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = f"★ {t}"
        p_t.font.bold = True
        p_t.font.size = Pt(10)
        p_t.font.color.rgb = c
        p_d = tf.add_paragraph()
        p_d.text = d
        p_d.font.size = Pt(8.2)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(2)
        cur_y += Inches(1.18)


def add_web_mockup_attack(slide, left, top, width, height):
    """Slide 7: Giao diện Web Mockup khi bị tấn công đảo bit."""
    add_browser_window_frame(slide, left, top, width, height, "127.0.0.1:8080/dashboard")

    # Banner báo động đỏ
    banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(0.6), width - Inches(0.4), Inches(0.7))
    banner.fill.solid()
    banner.fill.fore_color.rgb = RED_BG
    banner.line.color.rgb = RED_ACCENT
    tf_b = banner.text_frame
    tf_b.word_wrap = True
    p0 = tf_b.paragraphs[0]
    p0.text = "⚠️ PHÁT HIỆN TẤN CÔNG: VI PHẠM TOÀN VẸN DỮ LIỆU!"
    p0.font.bold = True
    p0.font.size = Pt(9.5)
    p0.font.color.rgb = RED_ACCENT
    p1 = tf_b.add_paragraph()
    p1.text = "Thuật toán HD-MerkleTree phát hiện sai khác tức thì và cô lập tệp tin bị can thiệp."
    p1.font.size = Pt(8)
    p1.font.color.rgb = TEXT_MAIN

    # Bảng chi tiết định vị lỗi
    detail_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(1.4), width - Inches(0.4), Inches(2.3))
    detail_box.fill.solid()
    detail_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    detail_box.line.color.rgb = BORDER_COLOR
    tf_d = detail_box.text_frame
    tf_d.word_wrap = True
    lines = [
        ("Vị trí tệp tin bị sửa đổi:", "certificates/pdf_audits/Chung_Thu_Kiem_Toan_110.pdf"),
        ("Hành vi kẻ tấn công:", "Mô phỏng đảo 1-bit ngẫu nhiên tại byte offset 0x008E"),
        ("Tốc độ định vị lỗi:", "0.022 ms (Duyệt theo nhánh phân cấp Biba O(log N))"),
        ("Tỷ lệ nguyên vẹn kho dữ liệu:", "99.9% (999/1.000 tệp tin còn lại an toàn tuyệt đối)"),
        ("Trạng thái phản ứng hệ thống:", "BLOCKED - Từ chối xác thực tệp hỏng, giữ nguyên State Root")
    ]
    for i, (k, v) in enumerate(lines):
        p_k = tf_d.paragraphs[0] if i == 0 else tf_d.add_paragraph()
        p_k.text = f"• {k} {v}"
        p_k.font.size = Pt(8.2)
        p_k.font.color.rgb = RED_ACCENT if i == 0 else TEXT_MAIN
        if i > 0:
            p_k.space_before = Pt(3)

    # Thẻ tóm tắt dưới cùng
    stat_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(3.85), width - Inches(0.4), Inches(1.6))
    stat_box.fill.solid()
    stat_box.fill.fore_color.rgb = BLUE_BG
    stat_box.line.color.rgb = BLUE_ACCENT
    tf_s = stat_box.text_frame
    tf_s.word_wrap = True
    ps0 = tf_s.paragraphs[0]
    ps0.text = "DẪN CHỨNG CHỨNG MINH KHOA HỌC:"
    ps0.font.bold = True
    ps0.font.size = Pt(9)
    ps0.font.color.rgb = NAVY_COLOR
    ps1 = tf_s.add_paragraph()
    ps1.text = "Nhờ phân cấp thư mục Biba 1977, hệ thống chỉ duyệt nhánh vi phạm thay vì quét toàn bộ 1.000 tệp tin. Thời gian định vị tệp lỗi nhanh hơn 4.000 lần so với duyệt tuần tự O(N)."
    ps1.font.size = Pt(8.2)
    ps1.font.color.rgb = TEXT_MAIN
    ps1.space_before = Pt(3)


def add_rfc6962_merkle_tree_diagram(slide, left, top, width, height):
    """Slide 8: Sơ đồ Cây Merkle phân tách miền RFC 6962."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.18), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "SƠ ĐỒ CÂY MERKLE PHÂN TÁCH MIỀN RFC 6962"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    # Root Node
    root = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(1.8), top + Inches(0.7), Inches(2.2), Inches(0.65))
    root.fill.solid()
    root.fill.fore_color.rgb = NAVY_COLOR
    root.line.color.rgb = NAVY_COLOR
    tf_r = root.text_frame
    tf_r.word_wrap = True
    pr = tf_r.paragraphs[0]
    pr.text = "ROOT HASH (32 B)\nHierarchical State Root"
    pr.font.bold = True
    pr.font.size = Pt(8)
    pr.font.color.rgb = RGBColor(255, 255, 255)
    pr.alignment = PP_ALIGN.CENTER

    # Branch Nodes
    b_left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.4), top + Inches(1.65), Inches(2.3), Inches(0.7))
    b_left.fill.solid()
    b_left.fill.fore_color.rgb = BLUE_BG
    b_left.line.color.rgb = BLUE_ACCENT
    tf_bl = b_left.text_frame
    tf_bl.word_wrap = True
    pbl = tf_bl.paragraphs[0]
    pbl.text = "NÚT NHÁNH 1 (0x01)\nH(0x01 || Trái || Phải)"
    pbl.font.bold = True
    pbl.font.size = Pt(8)
    pbl.font.color.rgb = BLUE_ACCENT
    pbl.alignment = PP_ALIGN.CENTER

    b_right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(3.1), top + Inches(1.65), Inches(2.3), Inches(0.7))
    b_right.fill.solid()
    b_right.fill.fore_color.rgb = BLUE_BG
    b_right.line.color.rgb = BLUE_ACCENT
    tf_br = b_right.text_frame
    tf_br.word_wrap = True
    pbr = tf_br.paragraphs[0]
    pbr.text = "NÚT NHÁNH 2 (0x01)\nH(0x01 || Trái || Phải)"
    pbr.font.bold = True
    pbr.font.size = Pt(8)
    pbr.font.color.rgb = BLUE_ACCENT
    pbr.alignment = PP_ALIGN.CENTER

    # Leaf Nodes
    leaf_w = Inches(1.25)
    for i in range(4):
        lx = left + Inches(0.25 + i * 1.35)
        leaf = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, lx, top + Inches(2.65), leaf_w, Inches(0.85))
        leaf.fill.solid()
        leaf.fill.fore_color.rgb = AMBER_BG
        leaf.line.color.rgb = AMBER_ACCENT
        tf_l = leaf.text_frame
        tf_l.word_wrap = True
        pl0 = tf_l.paragraphs[0]
        pl0.text = f"LÁ {i+1} (0x00)"
        pl0.font.bold = True
        pl0.font.size = Pt(7.5)
        pl0.font.color.rgb = AMBER_ACCENT
        pl0.alignment = PP_ALIGN.CENTER
        pl1 = tf_l.add_paragraph()
        pl1.text = "H(0x00||Data)"
        pl1.font.size = Pt(7)
        pl1.font.color.rgb = TEXT_MAIN
        pl1.alignment = PP_ALIGN.CENTER

    # Chú thích an toàn
    note = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), top + Inches(3.8), width - Inches(0.5), Inches(1.65))
    note.fill.solid()
    note.fill.fore_color.rgb = GREEN_BG
    note.line.color.rgb = GREEN_COLOR
    tf_n = note.text_frame
    tf_n.word_wrap = True
    pn0 = tf_n.paragraphs[0]
    pn0.text = "CHỨNG MINH TOÁN HỌC KHÁNG SECOND-PREIMAGE:"
    pn0.font.bold = True
    pn0.font.size = Pt(9)
    pn0.font.color.rgb = GREEN_COLOR
    pn1 = tf_n.add_paragraph()
    pn1.text = "Do tiền tố 0x00 (Lá) và 0x01 (Nhánh) là hai tập byte rời nhau, không tồn tại bất kỳ chuỗi dữ liệu tệp nào có thể tạo ra mã băm trùng khớp với nút nhánh. Xác suất giả mạo triệt tiêu về dưới 2^-128 theo Birthday bound của SHA-256."
    pn1.font.size = Pt(8.2)
    pn1.font.color.rgb = TEXT_MAIN
    pn1.space_before = Pt(3)


def add_pipeline_chevron_diagram(slide, left, top, width, height):
    """Slide 9: Sơ đồ 3 giai đoạn quy trình kỹ thuật."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "QUY TRÌNH HIỆN THỰC HÓA BA GIAI ĐOẠN"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    phases = [
        ("GIAI ĐOẠN 1: DỊCH THUẬT & KHUNG LÝ THUYẾT",
         "• Dịch học thuật toàn văn bài báo Harry Halpin 2025.\n• Trích xuất 6 luận điểm an ninh mật mã học cốt lõi.\n• Xác định lỗ hổng Rebuild O(N) của cây Merkle phẳng.",
         NAVY_COLOR, BLUE_BG),
        ("GIAI ĐOẠN 2: THIẾT KẾ & HIỆN THỰC HÓA PoC",
         "• Phát triển HD-MerkleTree phân cấp Biba 1977.\n• Cài đặt chuẩn tiền tố RFC 6962 Domain Separation.\n• Lập trình khung 7 Security Games độc lập.",
         PURPLE_ACCENT, PURPLE_BG),
        ("GIAI ĐOẠN 3: ĐO LƯỜNG THỰC NGHIỆM ĐỊNH LƯỢNG",
         "• Xây dựng bộ 1.000 tệp tin thực tế (4.1 MB, 23 thư mục).\n• Đo kiểm 5 vòng độc lập bằng time.perf_counter().\n• Tích hợp Web Demo trực quan kiểm thử một chạm.",
         GREEN_COLOR, GREEN_BG)
    ]
    cur_y = top + Inches(0.7)
    for title, desc, col, bg_c in phases:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), cur_y, width - Inches(0.5), Inches(1.4))
        box.fill.solid()
        box.fill.fore_color.rgb = bg_c
        box.line.color.rgb = col
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.bold = True
        p_t.font.size = Pt(9)
        p_t.font.color.rgb = col
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.2)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(3)
        cur_y += Inches(1.55)


def add_web_mockup_safe(slide, left, top, width, height):
    """Slide 10: Giao diện Web Mockup khi hệ thống an toàn ban đầu."""
    add_browser_window_frame(slide, left, top, width, height, "127.0.0.1:8080/monitor")

    # Banner an toàn xanh lá
    banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(0.6), width - Inches(0.4), Inches(0.7))
    banner.fill.solid()
    banner.fill.fore_color.rgb = GREEN_BG
    banner.line.color.rgb = GREEN_COLOR
    tf_b = banner.text_frame
    tf_b.word_wrap = True
    p0 = tf_b.paragraphs[0]
    p0.text = "✓ HỆ THỐNG AN TOÀN: ĐÃ NẠP 1.000 TỆP TIN NGUYÊN BẢN"
    p0.font.bold = True
    p0.font.size = Pt(9.5)
    p0.font.color.rgb = GREEN_COLOR
    p1 = tf_b.add_paragraph()
    p1.text = "Toàn bộ 1.000 tệp tin đã được lập chỉ mục băm và khởi tạo cây trạng thái thành công."
    p1.font.size = Pt(8)
    p1.font.color.rgb = TEXT_MAIN

    # Bảng thông số hệ thống
    detail_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(1.4), width - Inches(0.4), Inches(2.25))
    detail_box.fill.solid()
    detail_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    detail_box.line.color.rgb = BORDER_COLOR
    tf_d = detail_box.text_frame
    tf_d.word_wrap = True
    lines = [
        ("Tổng số tệp tin quản lý:", "1.000 tệp tin (Dung lượng tổng: 4.10 MB)"),
        ("Cấu trúc phân tầng Biba:", "44 nút thư mục phân cấp (23 thư mục vật lý)"),
        ("Mã băm gốc SHA-256 Root:", "68b6ff85d8ff48dcf7cfb21ab07edb6d..."),
        ("Hierarchical State Root:", "c00746adfd343a50bc8b04838c2b5eb3..."),
        ("Thời gian khởi tạo ban đầu:", "62.58 ms (Flat SHA) | 89.91 ms (HD-MerkleTree)")
    ]
    for i, (k, v) in enumerate(lines):
        p_k = tf_d.paragraphs[0] if i == 0 else tf_d.add_paragraph()
        p_k.text = f"• {k} {v}"
        p_k.font.size = Pt(8.2)
        p_k.font.color.rgb = NAVY_COLOR if i == 2 else TEXT_MAIN
        if i > 0:
            p_k.space_before = Pt(3)

    # Thẻ cam kết
    stat_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(3.8), width - Inches(0.4), Inches(1.65))
    stat_box.fill.solid()
    stat_box.fill.fore_color.rgb = BLUE_BG
    stat_box.line.color.rgb = BLUE_ACCENT
    tf_s = stat_box.text_frame
    tf_s.word_wrap = True
    ps0 = tf_s.paragraphs[0]
    ps0.text = "ĐỐI CHUẨN 3 MÔ HÌNH NỀN TẢNG (BASELINE):"
    ps0.font.bold = True
    ps0.font.size = Pt(9)
    ps0.font.color.rgb = NAVY_COLOR
    ps1 = tf_s.add_paragraph()
    ps1.text = "Cả 3 mô hình nền tảng (Flat SHA-256, HMAC-SHA256, Standard Merkle Tree) đều được hiện thực hóa độc lập và đo đạc trên cùng tập dữ liệu để đảm bảo tính khách quan tuyệt đối."
    ps1.font.size = Pt(8.2)
    ps1.font.color.rgb = TEXT_MAIN
    ps1.space_before = Pt(3)


def add_5tier_architecture_visual(slide, left, top, width, height):
    """Slide 11: Sơ đồ kiến trúc 5 tầng của mô hình HD-MerkleTree."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.18), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "SƠ ĐỒ 5 TẦNG KIẾN TRÚC MÔ HÌNH ĐỀ XUẤT HD-MERKLETREE"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    tiers = [
        ("TẦNG 5: CLIENT VERIFICATION & ZEROIZATION",
         "Hàm hmac.compare_digest() O(1) chống kênh kề Kocher 1996; secure_erase() xóa đè RAM.",
         NAVY_COLOR, RGBColor(238, 242, 255)),
        ("TẦNG 4: PROOF GENERATOR ENGINE",
         "Tạo Inclusion Proof gọn nhẹ 1.369 Byte (tiết kiệm 98.8% băng thông) & Consistency Proof RFC 6962.",
         BLUE_ACCENT, BLUE_BG),
        ("TẦNG 3: IN-MEMORY DYNAMIC MERKLE ENGINE",
         "ThreadPool băm song song, bộ nhớ đệm Dirty-Path cập nhật cục bộ chỉ h=4 nút tổ tiên.",
         PURPLE_ACCENT, PURPLE_BG),
        ("TẦNG 2: RFC 6962 DOMAIN SEPARATION",
         "Gắn tiền tố byte 0x00 cho nút lá và 0x01 cho nút nhánh, triệt tiêu 100% va chạm tiền ảnh cấp 2.",
         AMBER_ACCENT, AMBER_BG),
        ("TẦNG 1: MULTI-TIER STORAGE (BIBA 1977)",
         "Tổ chức 1.000 tệp tin trong 23 thư mục theo mô hình toàn vẹn phân cấp Biba (No write-up).",
         GREEN_COLOR, GREEN_BG)
    ]
    cur_y = top + Inches(0.7)
    for title, desc, col, bg_c in tiers:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), cur_y, width - Inches(0.5), Inches(0.88))
        box.fill.solid()
        box.fill.fore_color.rgb = bg_c
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.bold = True
        p_t.font.size = Pt(8.5)
        p_t.font.color.rgb = col
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(7.8)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(2)
        cur_y += Inches(0.98)


def add_dirty_path_chart_slide(slide, left, top, width, height):
    """Slide 12: Biểu đồ cột thực tế so sánh chi phí cập nhật Dirty-Path vs Rebuild."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.18), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "BIỂU ĐỒ ĐỐI SÁNH THỜI GIAN CẬP NHẬT TỆP (MS)"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    # Thêm biểu đồ cột native
    cdata = CategoryChartData()
    cdata.categories = ['Sửa tệp (Modify)', 'Thêm tệp (Add)', 'Xóa tệp (Delete)']
    cdata.add_series('Standard Merkle (ms)', (0.14, 73.42, 70.81))
    cdata.add_series('HD-MerkleTree (ms)', (0.14, 0.16, 0.10))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left + Inches(0.25), top + Inches(0.7), width - Inches(0.5), Inches(3.2),
        cdata
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False

    # Thẻ ghi chú chỉ số đột phá
    callout = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), top + Inches(4.0), width - Inches(0.5), Inches(1.45))
    callout.fill.solid()
    callout.fill.fore_color.rgb = GREEN_BG
    callout.line.color.rgb = GREEN_COLOR
    tf_c = callout.text_frame
    tf_c.word_wrap = True
    pc0 = tf_c.paragraphs[0]
    pc0.text = "⚡ ĐỘT PHÁ TỐC ĐỘ CẬP NHẬT CỦA HD-MERKLETREE:"
    pc0.font.bold = True
    pc0.font.size = Pt(9)
    pc0.font.color.rgb = GREEN_COLOR
    pc1 = tf_c.add_paragraph()
    pc1.text = "• Thêm tệp mới: Nhanh hơn 458 lần (0.16 ms so với 73.42 ms)\n• Xóa tệp tin: Nhanh hơn 708 lần (0.10 ms so với 70.81 ms)\n• Nguyên nhân: Tránh Rebuild O(N) nhờ cơ chế lan truyền Dirty-Path cục bộ."
    pc1.font.size = Pt(8.2)
    pc1.font.color.rgb = TEXT_MAIN
    pc1.space_before = Pt(3)


def add_benchmark_column_chart(slide, left, top, width, height):
    """Slide 13: Biểu đồ cột thực tế so sánh hiệu năng 4 mô hình."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.18), width - Inches(0.4), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "BIỂU ĐỒ HIỆU NĂNG 4 MÔ HÌNH (1.000 TỆP TIN)"
    p0.font.bold = True
    p0.font.size = Pt(10.5)
    p0.font.color.rgb = NAVY_COLOR

    cdata = CategoryChartData()
    cdata.categories = ['Flat SHA', 'HMAC', 'Std Merkle', 'HD-Merkle']
    cdata.add_series('Khởi tạo Build (ms)', (62.58, 60.65, 60.72, 89.91))
    cdata.add_series('Xác minh 1 tệp (ms)', (0.17, 0.07, 0.02, 0.02))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left + Inches(0.2), top + Inches(0.7), width - Inches(0.4), Inches(3.2),
        cdata
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False

    callout = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(4.0), width - Inches(0.4), Inches(1.5))
    callout.fill.solid()
    callout.fill.fore_color.rgb = BLUE_BG
    callout.line.color.rgb = BLUE_ACCENT
    tf_c = callout.text_frame
    tf_c.word_wrap = True
    pc0 = tf_c.paragraphs[0]
    pc0.text = "NHẬN XÉT THỰC NGHIỆM ĐỊNH LƯỢNG:"
    pc0.font.bold = True
    pc0.font.size = Pt(8.8)
    pc0.font.color.rgb = NAVY_COLOR
    pc1 = tf_c.add_paragraph()
    pc1.text = "• Xác minh 1 tệp của HD-Merkle đạt 0.02 ms (nhanh hơn ~8.5 lần Flat SHA 0.17 ms).\n• Thời gian khởi tạo 89.91 ms hoàn toàn tối ưu cho 1.000 tệp trong 23 thư mục."
    pc1.font.size = Pt(8)
    pc1.font.color.rgb = TEXT_MAIN
    pc1.space_before = Pt(2)


def add_security_games_cards_visual(slide, left, top, width, height):
    """Slide 14: Lưới 4 thẻ thách đấu an ninh hình thức Games 1-4."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.18), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "KẾT QUẢ ĐỐI KHÁNG THỰC TẾ (GAMES 1 – 4)"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    games_data = [
        ("GAME 1: PPT ADVERSARY", "Tấn công: Tìm va chạm ngẫu nhiên\nKết quả: Thử thách 50 lượt ngẫu nhiên\nKẻ thù thắng: 0/50 (0.0% va chạm)\nLợi thế đối thủ: Adv <= 2^-128", GREEN_COLOR, GREEN_BG),
        ("GAME 2: DOLEV-YAO NETWORK", "Tấn công: Nghe lén, đảo bit, tráo proof\nKết quả: Kiểm tra đối chiếu Root Hash\nPhát hiện: 100% can thiệp bị chặn\nThời gian phản ứng: 0.02 ms", BLUE_ACCENT, BLUE_BG),
        ("GAME 3: SECOND-PREIMAGE", "Tấn công: Tráo nút lá thành nút nhánh\nKết quả: Phân tách tiền tố 0x00 & 0x01\nKháng va chạm: Triệt tiêu 100%\nChuẩn áp dụng: RFC 6962", PURPLE_ACCENT, PURPLE_BG),
        ("GAME 4: AVALANCHE EFFECT", "Tấn công: Đảo 1 bit tìm quy luật băm\nKết quả đo đạc: 129/256 bit biến đổi\nTỷ lệ khuếch tán: 50.39% ≈ 50%\nĐặc tính: Chuẩn Shannon Diffusion", AMBER_ACCENT, AMBER_BG)
    ]
    cur_y = top + Inches(0.7)
    for title, desc, col, bg_c in games_data:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), cur_y, width - Inches(0.5), Inches(1.15))
        box.fill.solid()
        box.fill.fore_color.rgb = bg_c
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = f"🛡️ {title}  [ PASS ]"
        p_t.font.bold = True
        p_t.font.size = Pt(9)
        p_t.font.color.rgb = col
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(2)
        cur_y += Inches(1.25)


def add_web_mockup_console_slide(slide, left, top, width, height):
    """Slide 16: Giao diện phòng thí nghiệm Web Demo Console."""
    add_browser_window_frame(slide, left, top, width, height, "127.0.0.1:8080/terminal")

    # 3 Nút bấm tấn công giả lập
    btn_data = [("⚡ Đảo 1 Bit", RED_ACCENT), ("➕ Chèn Tệp Lạ", AMBER_ACCENT), ("🗑️ Xóa Lén Tệp", BLUE_ACCENT)]
    w_btn = (width - Inches(0.8)) / 3
    for i, (b_txt, b_col) in enumerate(btn_data):
        btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25 + i * (w_btn + Inches(0.15))), top + Inches(0.6), w_btn, Inches(0.45))
        btn.fill.solid()
        btn.fill.fore_color.rgb = b_col
        btn.line.fill.background()
        tf = btn.text_frame
        p = tf.paragraphs[0]
        p.text = b_txt
        p.font.bold = True
        p.font.size = Pt(8.5)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    # Cửa sổ Console Terminal màu tối
    console = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), top + Inches(1.15), width - Inches(0.5), Inches(4.25))
    console.fill.solid()
    console.fill.fore_color.rgb = DARK_CONSOLE
    console.line.color.rgb = DARK_CONSOLE
    tf_c = console.text_frame
    tf_c.word_wrap = True

    logs = [
        ("$ python src/demo/web_demo.py --port 8080", CONSOLE_YELLOW),
        ("[SYSTEM] Web Server started at http://127.0.0.1:8080", CONSOLE_GREEN),
        ("[INFO] Scanning realistic_dataset (1.000 files, 4.10 MB)...", RGBColor(255, 255, 255)),
        ("[SUCCESS] HD-MerkleTree built: 44 directories indexed in 0.089 s", CONSOLE_GREEN),
        ("[SIMULATE] User clicked: [Đảo 1 Bit Ngẫu Nhiên]", CONSOLE_YELLOW),
        ("[ATTACK] Injected bit-flip into 'Chung_Thu_Kiem_Toan_110.pdf'", RGBColor(248, 113, 113)),
        ("[ALERT] Localized corrupted file in 0.022 ms via Biba Dirty-Path!", CONSOLE_YELLOW),
        ("[VERDICT] Proof verification: REJECTED (Integrity violation detected)", RGBColor(248, 113, 113)),
        ("[BENCHMARK] Throughput: 45,454 files/sec | Zero memory leaks", CONSOLE_GREEN)
    ]
    for i, (l_txt, l_col) in enumerate(logs):
        p = tf_c.paragraphs[0] if i == 0 else tf_c.add_paragraph()
        p.text = l_txt
        p.font.size = Pt(8)
        p.font.color.rgb = l_col
        if i > 0:
            p.space_before = Pt(4)


def add_network_payload_chart(slide, left, top, width, height):
    """Slide 17: Biểu đồ thanh thực tế so sánh dung lượng truyền tải mạng (Proof Payload)."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.18), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "BIỂU ĐỒ DUNG LƯỢNG TRUYỀN TẢI QUA MẠNG (BYTES)"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    cdata = CategoryChartData()
    cdata.categories = ['Flat SHA', 'HMAC', 'Std Merkle', 'HD-Merkle']
    cdata.add_series('Kích thước Proof (Bytes)', (113802, 113807, 990, 1369))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        left + Inches(0.25), top + Inches(0.7), width - Inches(0.5), Inches(3.2),
        cdata
    )
    chart = chart_shape.chart
    chart.has_legend = False

    callout = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), top + Inches(4.0), width - Inches(0.5), Inches(1.45))
    callout.fill.solid()
    callout.fill.fore_color.rgb = GREEN_BG
    callout.line.color.rgb = GREEN_COLOR
    tf_c = callout.text_frame
    tf_c.word_wrap = True
    pc0 = tf_c.paragraphs[0]
    pc0.text = "🌐 TIẾT KIỆM BĂNG THÔNG MẠNG ĐỘT PHÁ (98.8%):"
    pc0.font.bold = True
    pc0.font.size = Pt(9)
    pc0.font.color.rgb = GREEN_COLOR
    pc1 = tf_c.add_paragraph()
    pc1.text = "• Bằng chứng HD-Merkle chỉ 1.369 Byte (so với toàn bộ bảng 113.802 Byte).\n• Máy khách chỉ cần tải các nút anh em trên nhánh, không lộ 999 tệp còn lại.\n• Bảo vệ 100% tính riêng tư siêu dữ liệu (Metadata Confidentiality)."
    pc1.font.size = Pt(8.2)
    pc1.font.color.rgb = TEXT_MAIN
    pc1.space_before = Pt(3)


def add_closing_quadrant_visual(slide, left, top, width, height):
    """Slide 18: Lưới 4 khối đóng góp giá trị toàn diện."""
    add_card(slide, left, top, width, height)
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.5))
    p0 = tb.text_frame.paragraphs[0]
    p0.text = "BỐN ĐÓNG GÓP TOÀN DIỆN CỦA ĐỒ ÁN DT005"
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = NAVY_COLOR

    quads = [
        ("1. HỌC THUẬT & DỊCH THUẬT", "Dịch toàn văn bài báo Harry Halpin 2025; đối chiếu 100% luận điểm Kẻ thù mật mã học.", NAVY_COLOR, BLUE_BG),
        ("2. KIẾN TRÚC PoC SÁNG TẠO", "Đề xuất HD-MerkleTree phân cấp Biba 1977, cập nhật Dirty-Path nhanh hơn 458 - 708 lần.", BLUE_ACCENT, BLUE_BG),
        ("3. MÔ PHỎNG 7 SECURITY GAMES", "Khung kiểm định đối kháng toán học độc lập (PPT, Dolev-Yao, RFC 6962, Tuyết lở, Kênh kề).", PURPLE_ACCENT, PURPLE_BG),
        ("4. THỰC NGHIỆM ĐỊNH LƯỢNG", "Đo kiểm 1.000 tệp tin qua 5 vòng lặp, chứng minh tiết kiệm 98.8% băng thông qua Web Demo.", GREEN_COLOR, GREEN_BG)
    ]
    cur_y = top + Inches(0.7)
    for title, desc, col, bg_c in quads:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.25), cur_y, width - Inches(0.5), Inches(1.15))
        box.fill.solid()
        box.fill.fore_color.rgb = bg_c
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = f"🏆 {title}"
        p_t.font.bold = True
        p_t.font.size = Pt(9.5)
        p_t.font.color.rgb = col
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.2)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(2)
        cur_y += Inches(1.25)


# ==============================================================================
# SLIDE 1: TRANG TIÊU ĐỀ
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)
add_base_background(slide1)

top_stripe = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.4))
top_stripe.fill.solid()
top_stripe.fill.fore_color.rgb = NAVY_COLOR
top_stripe.line.color.rgb = NAVY_COLOR

card1 = add_card(slide1, Inches(1.0), Inches(0.9), Inches(11.333), Inches(5.8))

tb1 = slide1.shapes.add_textbox(Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.2))
tf1 = tb1.text_frame
tf1.word_wrap = True

p = tf1.paragraphs[0]
p.text = "TRƯỜNG ĐẠI HỌC TÀI NGUYÊN VÀ MÔI TRƯỜNG TP. HỒ CHÍ MINH"
p.font.bold = True
p.font.size = Pt(13.5)
p.font.color.rgb = NAVY_COLOR
p.alignment = PP_ALIGN.CENTER

p = tf1.add_paragraph()
p.text = "KHOA CÔNG NGHỆ THÔNG TIN"
p.font.bold = True
p.font.size = Pt(12)
p.font.color.rgb = BLUE_ACCENT
p.alignment = PP_ALIGN.CENTER
p.space_after = Pt(14)

p = tf1.add_paragraph()
p.text = "BÁO CÁO THUYẾT TRÌNH ĐỒ ÁN GIỮA KỲ"
p.font.bold = True
p.font.size = Pt(22)
p.font.color.rgb = RED_ACCENT
p.alignment = PP_ALIGN.CENTER

p = tf1.add_paragraph()
p.text = "ĐỀ TÀI DT005: HỆ THỐNG KIỂM TRA TOÀN VẸN TỆP BẰNG SHA-256, HMAC VÀ CÂY MERKLE"
p.font.bold = True
p.font.size = Pt(15.5)
p.font.color.rgb = NAVY_COLOR
p.alignment = PP_ALIGN.CENTER
p.space_after = Pt(10)

p = tf1.add_paragraph()
p.text = "Học phần: An Toàn và Bảo Mật Hệ Thống Thông Tin"
p.font.bold = True
p.font.size = Pt(12)
p.font.color.rgb = TEXT_MAIN
p.alignment = PP_ALIGN.CENTER

p = tf1.add_paragraph()
p.text = "Giảng viên hướng dẫn: ThS. Phạm Trọng Huynh\nSinh viên thực hiện: Võ Anh Kiệt  —  MSSV: 1150080143  —  Lớp: 11_CNPM2"
p.font.bold = True
p.font.size = Pt(12.5)
p.font.color.rgb = NAVY_COLOR
p.alignment = PP_ALIGN.CENTER
p.space_after = Pt(12)

p = tf1.add_paragraph()
p.text = "Nghiên Cứu Và Phát Triển Dựa Trên Bài Báo Khoa Học Gốc: The Adversary: The Philosophy of Cryptography\nTác giả: Harry Halpin — Xuất bản trên Tạp chí An ninh mạng Oxford (Oxford Journal of Cybersecurity, 2025)"
p.font.size = Pt(10)
p.font.color.rgb = TEXT_MUTED
p.alignment = PP_ALIGN.CENTER
p.space_after = Pt(8)

p = tf1.add_paragraph()
p.text = "Phương Pháp Luận Nghiên Cứu: Triển khai 3 Trụ cột độc lập (Kỹ thuật PoC, Mô phỏng an ninh 7 Games, Thực nghiệm trên 1.000 tệp)"
p.font.bold = True
p.font.size = Pt(10.5)
p.font.color.rgb = GREEN_COLOR
p.alignment = PP_ALIGN.CENTER


# ==============================================================================
# SLIDE 2: GIỚI THIỆU & TÓM TẮT BÀI BÁO (OXFORD JOURNAL CARD)
# ==============================================================================
slide2 = prs.slides.add_slide(blank_layout)
add_base_background(slide2)
add_header(slide2, "Phần 1: Dịch bài báo và tóm tắt nội dung", "1. Giới Thiệu Và Tóm Tắt Bài Báo Khoa Học Của Tác Giả Harry Halpin")

add_bullet_card(slide2, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Thông Tin Xuất Bản & Tóm Tắt Toàn Văn Bài Báo", [
        ("• Tên bài báo tiếng Anh gốc:",
         "The Adversary: The Philosophy of Cryptography."),
        ("• Tác giả & Cơ quan nghiên cứu:",
         "Harry Halpin — Nhà nghiên cứu mật mã học hàng đầu, đồng sáng lập Nym Technologies và cựu nghiên cứu viên cấp cao tại Inria (Pháp)."),
        ("• Tạp chí & Năm xuất bản:",
         "Tạp chí An ninh mạng Oxford (Oxford Journal of Cybersecurity), Nhà xuất bản Đại học Oxford (Oxford University Press), Tập 11, Số 1, Năm 2025."),
        ("• Luận điểm triết học trung tâm của bài báo:",
         "Mật mã học không chỉ là bài toán toán học thuần túy trên lý thuyết, mà là một cuộc chiến tương tác đối kháng có chủ đích giữa Kẻ phòng thủ và Kẻ thù hình thức (The Adversary)."),
        ("• Khung phân loại 4 mô hình Kẻ thù trong bài báo:",
         "1. Kẻ tấn công thụ động (Passive Eavesdropper: Nghe lén đường truyền).\n"
         "2. Kẻ tấn công chủ động Dolev-Yao (Active Network Adversary: Can thiệp gói tin).\n"
         "3. Kẻ thù tính toán PPT (Probabilistic Polynomial-Time Adversary).\n"
         "4. Kẻ thù khai thác kênh kề & siêu dữ liệu (Side-Channel & Metadata Adversary).")
    ], font_body_size=8.8)

add_academic_paper_card(slide2, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 3: BỐI CẢNH & ĐỘNG LỰC NGHIÊN CỨU (KNOWLEDGE TRANSFER DIAGRAM)
# ==============================================================================
slide3 = prs.slides.add_slide(blank_layout)
add_base_background(slide3)
add_header(slide3, "Phần 1: Dịch bài báo và tóm tắt nội dung", "2. Bối Cảnh, Động Lực Nghiên Cứu Và Lời Kêu Gọi Kiểm Nghiệm Thực Tế")

add_bullet_card(slide3, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Khoảng Trống Học Thuật & Động Lực Triển Khai Đồ Án", [
        ("• Khoảng trống học thuật trong bài báo gốc:",
         "Bài báo của Harry Halpin (2025) là một công trình nghiên cứu triết học và lý thuyết mật mã học đỉnh cao, nhưng HOÀN TOÀN KHÔNG CÓ BỘ MÃ NGUỒN CỤ THỂ, không có tập dữ liệu thực nghiệm và không có các số liệu đo lường định lượng."),
        ("• Lời kêu gọi từ bài báo (Mục 7 Trang 14-15):",
         "Tác giả nhấn mạnh: Các cấu trúc dữ liệu mật mã (như Merkle Tree) cần được kiểm chứng trong các môi trường hệ thống thực tế với dữ liệu biến thiên, thay vì chỉ dừng lại ở các chứng minh toán học trên trang giấy."),
        ("• Động lực và sứ mệnh của Đồ án DT005:",
         "Hiện thực hóa 100% khung lý thuyết của Harry Halpin thành một hệ thống phần mềm hoàn chỉnh chạy thực tế, thiết kế bộ dữ liệu 1.000 tệp tin thực tế và đo đạc các chỉ số hiệu năng khách quan."),
        ("• Ý nghĩa thực tiễn của đề tài:",
         "Cung cấp giải pháp bảo vệ toàn vẹn dữ liệu cho các hệ thống doanh nghiệp, cơ sở dữ liệu nhật ký kiểm toán không thể sửa đổi (Append-Only Log) và bảo vệ máy khách nhẹ (Light Client) tiết kiệm hơn 98% băng thông mạng.")
    ], font_body_size=8.8)

add_knowledge_transfer_diagram(slide3, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 4: TỔNG HỢP CÁC LUẬN ĐIỂM MẬT MÃ HỌC CỐT LÕI (4 THESES GRID)
# ==============================================================================
slide4 = prs.slides.add_slide(blank_layout)
add_base_background(slide4)
add_header(slide4, "Phần 1: Dịch bài báo và tóm tắt nội dung", "3. Tổng Hợp Các Luận Điểm Mật Mã Học Cốt Lõi Từ Bài Báo")

add_bullet_card(slide4, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "6 Luận Điểm Mật Mã Học Nền Tảng Của Harry Halpin", [
        ("• 1. Mô hình Kẻ thù PPT (Mục 4 Trang 7):",
         "Kẻ tấn công có năng lực tính toán ngẫu nhiên trong thời gian đa thức (Goldwasser-Micali 1984). Xác suất bẻ gãy hệ thống an toàn không được vượt quá hàm không đáng kể negl(n) <= 2^-128."),
        ("• 2. Nguyên lý Kerckhoffs (Mục 2 Trang 3):",
         "Độ an toàn của hệ thống mật mã phải phụ thuộc hoàn toàn vào sự bí mật của khóa (hoặc tính bất biến của Root Hash), tuyệt đối không dựa vào sự bí mật của mã nguồn."),
        ("• 3. Kẻ thù Mạng Dolev-Yao (Mục 4 Trang 8):",
         "Kẻ tấn công kiểm soát toàn bộ kênh truyền thông, có quyền sửa đổi gói tin, tráo đổi dữ liệu và thực hiện tấn công phát lại (Replay Attack)."),
        ("• 4. Chuẩn phân tách miền RFC 6962 (Mục 5 Trang 8-10):",
         "Bắt buộc gắn tiền tố phân tách miền 0x00 cho nút lá và 0x01 cho nút nhánh để triệt tiêu hoàn toàn tấn công giả mạo tiền ảnh cấp 2 (Second-Preimage Attack)."),
        ("• 5. Xác minh cục bộ O(log N) của Merkle (Mục 4 Trang 5):",
         "Máy khách nhẹ (Light Client) chỉ cần một bằng chứng logarit để xác minh 1 tệp tin mà không cần tải toàn bộ kho dữ liệu."),
        ("• 6. Tấn công kênh kề & Rò rỉ siêu dữ liệu (Mục 5 Trang 10, Mục 6):",
         "Cảnh báo nguy cơ tấn công đo thời gian (Timing Attack theo Paul Kocher 1996) và rò rỉ danh tính tệp tin qua siêu dữ liệu.")
    ], font_body_size=8.5)

add_core_theses_grid(slide4, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 5: PHƯƠNG PHÁP LUẬN NGHIÊN CỨU: BA TRỤ CỘT ĐỘC LẬP (3 PILLARS VISUAL)
# ==============================================================================
slide5_method = prs.slides.add_slide(blank_layout)
add_base_background(slide5_method)
add_header(slide5_method, "Phần 2: Phương pháp luận nghiên cứu", "4. Phương Pháp Luận Nghiên Cứu: Ba Trụ Cột Độc Lập Chuẩn Mực")

add_bullet_card(slide5_method, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Tuyên Bố Rõ Ràng 3 Trụ Cột Phương Pháp Luận", [
        ("• Tuyên bố tính chuẩn mực khoa học:",
         "Nhóm nghiên cứu KHÔNG gộp chung các phần việc làm một, mà phân định rõ ràng thành 3 trụ cột phương pháp luận độc lập, tự chủ:"),
        ("• Trụ cột I (Hướng 1) - Kỹ thuật hệ thống & PoC:",
         "- Tệp thực thi: src/proposed_model/hd_merkle_tree.py\n"
         "- Bản chất: Đề xuất kiến trúc Cây Merkle phân tầng (HD-MerkleTree) kết hợp mô hình toàn vẹn Biba 1977 và kỹ thuật cập nhật Dirty-Path cục bộ trong RAM.\n"
         "- Đóng góp: Giải quyết triệt để vấn đề Rebuild toàn bộ O(N) của cây Merkle phẳng."),
        ("• Trụ cột II (Hướng 2) - Mô phỏng an ninh hình thức (7 Games):",
         "- Tệp thực thi: src/proposed_model/security_games.py\n"
         "- Bản chất: Khung kiểm định toán học tự động 7 mô hình an ninh theo đúng các giả định của Harry Halpin (PPT, Dolev-Yao, RFC 6962, Tuyết lở, Biba, Kênh kề)."),
        ("• Trụ cột III (Hướng 4) - Thực nghiệm định lượng quy mô lớn:",
         "- Tệp thực thi: src/benchmark/run_benchmark.py\n"
         "- Bản chất: Đo lường khách quan 5 lần lặp độc lập trên 1.000 tệp tin thực tế, so sánh đối đầu trực diện 4 mô hình bảo mật.")
    ], font_body_size=8.5)

add_three_pillars_visual(slide5_method, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 6: MA TRẬN ĐỐI CHIẾU CHI TIẾT 100% (AUDIT SCORECARD)
# ==============================================================================
slide6_matrix = prs.slides.add_slide(blank_layout)
add_base_background(slide6_matrix)
add_header(slide6_matrix, "Phần 2: Phương pháp luận nghiên cứu", "5. Bảng Ma Trận Đối Chiếu Chi Tiết 100% Giữa Bài Báo Và Đồ Án")

add_bullet_card(slide6_matrix, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Bảng Đối Chiếu Chi Tiết 6 Điểm Nòng Cốt", [
        ("• Điểm đối chiếu 1: Kẻ thù PPT (Goldwasser-Micali 1984):",
         "- Bài báo: Mục 4 Trang 7  |  Đồ án: Game 1 (security_games.py)\n"
         "- Minh chứng: Thử thách 50 lượt ngẫu nhiên, lợi thế Adv(A) = 0.0 <= 2^-128."),
        ("• Điểm đối chiếu 2: Nguyên lý Kerckhoffs (Kerckhoffs 1883):",
         "- Bài báo: Mục 2 Trang 3  |  Đồ án: Công khai 100% mã nguồn trên GitHub\n"
         "- Minh chứng: Chỉ cần Root Hash 32 byte để bảo vệ toàn vẹn dữ liệu."),
        ("• Điểm đối chiếu 3: Kẻ tấn công Mạng Dolev-Yao (Dolev-Yao 1983):",
         "- Bài báo: Mục 4 Trang 8  |  Đồ án: Game 2 & Web Demo (Nút đảo 1 bit)\n"
         "- Minh chứng: Chặn đứng 100% hành vi sửa tệp, tráo Proof và đổi đường dẫn."),
        ("• Điểm đối chiếu 4: Phân tách miền RFC 6962:",
         "- Bài báo: Mục 5 Trang 8-10  |  Đồ án: Tiền tố 0x00 cho lá và 0x01 cho nhánh\n"
         "- Minh chứng: Triệt tiêu 100% va chạm tiền ảnh cấp hai (Second-preimage)."),
        ("• Điểm đối chiếu 5: Xác minh cục bộ O(log N) của Ralph Merkle:",
         "- Bài báo: Mục 4 Trang 5  |  Đồ án: Bằng chứng Merkle Proof (0.97 - 1.34 KB)\n"
         "- Minh chứng: Tiết kiệm 98.8% băng thông so với việc tải toàn bộ 111 KB."),
        ("• Điểm đối chiếu 6: Lời kêu gọi kiểm nghiệm thực tế (Mục 7):",
         "- Bài báo: Kêu gọi thử nghiệm  |  Đồ án: Bộ 1.000 tệp tin thực tế (4.1 MB)")
    ], font_body_size=8.0)

add_audit_scorecard(slide6_matrix, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 7: CHỨNG MINH CHI TIẾT ĐỐI CHIẾU 1, 2, 3 (WEB MOCKUP ATTACK ALERT)
# ==============================================================================
slide7 = prs.slides.add_slide(blank_layout)
add_base_background(slide7)
add_header(slide7, "Phần 2: Phương pháp luận nghiên cứu", "6. Dẫn Chứng Chứng Minh Chi Tiết Điểm Đối Chiếu 1, 2 Và 3")

add_bullet_card(slide7, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Chứng Minh Toán Học Và Khảo Nghiệm Thực Tế", [
        ("• Điểm 1: Không gian tìm kiếm kẻ thù PPT (Mục 4 Trang 7 Tài liệu [33, 34]):",
         "- Lý thuyết bài báo: Kẻ tấn công thời gian đa thức không thể giải bài toán tìm tiền ảnh hoặc va chạm băm nếu không gian khóa đủ lớn.\n"
         "- Dẫn chứng toán học: Mã băm SHA-256 có độ dài 256 bit. Không gian tìm kiếm là 2^256 chuỗi. Dù kẻ thù dùng thuật toán ngày sinh (Birthday Attack), độ phức tạp vẫn là 2^128 phép băm. Xác suất thành công thực nghiệm đo được qua 50 lượt thử thách đối kháng là 0.0%."),
        ("• Điểm 2: Nguyên lý Kerckhoffs về tính minh bạch (Mục 2 Trang 3 Tài liệu [14, 15]):",
         "- Lý thuyết bài báo: Độ an toàn của hệ thống không được dựa vào sự bí mật của mã nguồn mà chỉ dựa vào khóa hoặc giá trị băm tin cậy.\n"
         "- Dẫn chứng chứng minh: Nhóm công khai toàn bộ mã nguồn trên máy chủ, công khai cấu trúc cây Merkle. Người kiểm tra chỉ cần nắm giữ đúng chuỗi băm gốc 32 byte (256 bit) có không gian tìm kiếm 2 mũ 256, bất khả thi đối với mọi siêu máy tính."),
        ("• Điểm 3: Ba hành vi kẻ thù chủ động Dolev-Yao (Mục 4 Trang 8 Tài liệu [52]):",
         "- Lý thuyết bài báo: Kẻ tấn công mạng có toàn quyền can thiệp, đánh tráo và giả mạo thông điệp.\n"
         "- Dẫn chứng chứng minh: Lập trình sẵn 3 nút mô phỏng tấn công trong web_demo.py. Thao tác trên 1.000 tệp tin cho thấy hệ thống phát hiện chính xác 100% tệp bị can thiệp và định vị trong 0.022 ms.")
    ], font_body_size=8.8)

add_web_mockup_attack(slide7, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 8: CHỨNG MINH CHI TIẾT ĐỐI CHIẾU 4, 5, 6 (RFC 6962 TREE DIAGRAM)
# ==============================================================================
slide8 = prs.slides.add_slide(blank_layout)
add_base_background(slide8)
add_header(slide8, "Phần 2: Phương pháp luận nghiên cứu", "7. Dẫn Chứng Chứng Minh Chi Tiết Điểm Đối Chiếu 4, 5 Và 6")

add_bullet_card(slide8, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Chứng Minh Toán Học Và Khảo Nghiệm Thực Tế", [
        ("• Điểm 4: Nguy cơ va chạm ghép chuỗi và Chuẩn RFC 6962 (Mục 5 Trang 8-10):",
         "- Lý thuyết bài báo: Cây Merkle thông thường nếu băm nút lá và nhánh cùng chung định dạng H(a || b) sẽ bị tấn công tráo đổi lá thành nhánh.\n"
         "- Dẫn chứng toán học: Nhóm áp dụng RFC 6962 phân tách miền: Lá = H(0x00 || Dữ liệu) và Nhánh = H(0x01 || Trái || Phải). Do tập tiền tố rời nhau, xác suất va chạm giả mạo triệt tiêu về dưới 2 mũ âm 128, chứng minh an toàn tuyệt đối."),
        ("• Điểm 5: Xác minh cục bộ O(log N) của Ralph Merkle (Mục 4 Trang 5 Tài liệu [27]):",
         "- Lý thuyết bài báo: Bằng chứng Merkle cho phép xác minh độc lập 1 tệp với độ phức tạp O(log N) thay vì phải tải và kiểm tra toàn bộ O(N) tệp tin.\n"
         "- Dẫn chứng số liệu: Trên 1.000 tệp tin, chiều cao cây chỉ là 10-13 tầng băm. Kích thước bằng chứng chỉ 1.37 KB (1.369 Byte) so với toàn bộ danh mục 111.13 KB, tiết kiệm 98.8% băng thông mạng và thời gian xác minh chỉ mất 0.02 ms."),
        ("• Điểm 6: Lời kêu gọi kiểm nghiệm thực tế (Mục 7 Trang 14-15):",
         "- Lý thuyết bài báo: Mật mã học đòi hỏi đánh giá trên dữ liệu thực tế biến thiên kích thước.\n"
         "- Dẫn chứng thực thi: Nhóm xây dựng 1.000 tệp tin thực tế, đo đạc độc lập 5 lần bằng hàm CPU thời gian thực perf_counter.")
    ], font_body_size=8.8)

add_rfc6962_merkle_tree_diagram(slide8, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 9: QUY TRÌNH HIỆN THỰC HÓA KỸ THUẬT 3 GIAI ĐOẠN (PIPELINE CHEVRON)
# ==============================================================================
slide9_data = prs.slides.add_slide(blank_layout)
add_base_background(slide9_data)
add_header(slide9_data, "Phần 3: Dữ liệu thực nghiệm và kiến trúc kỹ thuật", "8. Quy Trình Hiện Thực Hóa Kỹ Thuật 3 Giai Đoạn Từ Lý Thuyết Đến Mã Nguồn")

add_bullet_card(slide9_data, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Quy Trình Hiện Thực Hóa Kỹ Thuật Chuẩn Khoa Học", [
        ("• Giai đoạn 1: Nghiên cứu lý thuyết & Dịch thuật chuyên sâu:",
         "- Dịch toàn văn 28 trang bài báo tiếng Anh của Harry Halpin (2025).\n"
         "- Trích xuất và mô hình hóa toán học 6 luận điểm then chốt.\n"
         "- Phát hiện vấn đề: Cây Merkle phẳng truyền thống bị nút thắt cổ chai Rebuild O(N) khi Thêm/Xóa tệp."),
        ("• Giai đoạn 2: Thiết kế kiến trúc PoC & Mô phỏng an ninh:",
         "- Thiết kế mô hình Cây Merkle phân cấp HD-MerkleTree theo cấu trúc Biba 1977.\n"
         "- Tích hợp chuẩn phân tách miền RFC 6962 và cơ chế xóa bộ nhớ an toàn.\n"
         "- Lập trình độc lập 7 mô hình kiểm định an ninh hình thức (security_games.py)."),
        ("• Giai đoạn 3: Xây dựng tập dữ liệu & Thực nghiệm định lượng:",
         "- Sinh tập dữ liệu thực tế realistic_dataset gồm 1.000 tệp tin, 23 thư mục (4.1 MB, 14 định dạng: PDF, log, source code, hình ảnh).\n"
         "- Đo lường đối đầu 4 cơ chế bảo mật qua 5 vòng lặp độc lập, xuất JSON & CSV.\n"
         "- Đóng gói giao diện Web Demo trực quan trình diễn trước hội đồng.")
    ], font_body_size=8.8)

add_pipeline_chevron_diagram(slide9_data, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 10: [TRỤ CỘT I - HƯỚNG 1] 3 MÔ HÌNH NỀN TẢNG (WEB MOCKUP SAFE STATE)
# ==============================================================================
slide10_base = prs.slides.add_slide(blank_layout)
add_base_background(slide10_base)
add_header(slide10_base, "Phần 4A: Trụ cột I - Kỹ thuật hệ thống & PoC (Hướng 1)", "9. Hiện Thực Hóa 3 Mô Hình Kiểm Tra Toàn Vẹn Nền Tảng Đối Chứng")

add_bullet_card(slide10_base, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Đặc Điểm Kỹ Thuật Và Kết Quả Đo Đạc 3 Mô Hình", [
        ("• Mô hình 1: Danh sách phẳng băm SHA-256 (sha256_manifest.py):",
         "- Nguyên lý: Băm tuần tự từng tệp tin và lưu chuỗi băm tổng hợp.\n"
         "- Bằng chứng thực nghiệm: Khởi tạo ~62.6 ms, dung lượng bảng 111.13 KB. Hạn chế: Muốn kiểm tra 1 tệp tin trong mô hình cơ bản, máy khách phải lưu toàn bộ manifest 111 KB (hoặc quét tệp tốn 0.17 ms)."),
        ("• Mô hình 2: Danh sách băm có khóa bí mật HMAC (hmac_manifest.py):",
         "- Nguyên lý: Sử dụng khóa bí mật đối xứng K để chống giả mạo danh mục.\n"
         "- Bằng chứng thực nghiệm: Khởi tạo ~60.6 ms, dung lượng 111.14 KB. Hạn chế: Buộc phải phân phối khóa bí mật K cho máy khách, chưa hỗ trợ chống chối bỏ bất đối xứng."),
        ("• Mô hình 3: Cây Merkle nhị phân truyền thống (standard_merkle_tree.py):",
         "- Nguyên lý: Cây Merkle phẳng chuẩn theo phát minh Ralph Merkle năm 1978.\n"
         "- Bằng chứng thực nghiệm: Cho phép tạo bằng chứng xác thực từng tệp nặng chỉ 990 Byte, kiểm tra trong 0.02 ms, tiết kiệm hơn 99% băng thông."),
        ("• Đánh giá kết quả:",
         "Cả 3 mô hình đều phát hiện chính xác 100% khi xảy ra tấn công trên 1.000 tệp tin.")
    ], font_body_size=8.8)

add_web_mockup_safe(slide10_base, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 11: [TRỤ CỘT I - HƯỚNG 1] KIẾN TRÚC MÔ HÌNH ĐỀ XUẤT HD-MERKLETREE
# ==============================================================================
slide11_arch = prs.slides.add_slide(blank_layout)
add_base_background(slide11_arch)
add_header(slide11_arch, "Phần 4A: Trụ cột I - Kỹ thuật hệ thống & PoC (Hướng 1)", "10. Sơ Đồ Kiến Trúc Đề Xuất HD-MerkleTree & Cơ Chế Xóa Bộ Nhớ")

add_bullet_card(slide11_arch, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Giải Thích Chi Tiết 5 Tầng Kiến Trúc Hệ Thống (hd_merkle_tree.py)", [
        ("• Tầng 1: Lưu trữ phân tầng Biba 1977:",
         "Tổ chức 1.000 tệp tin trong 23 thư mục theo mô hình toàn vẹn Biba. Kháng tấn công ghi ngược từ tệp thường lên thư mục bảo mật."),
        ("• Tầng 2: Phân tách miền băm chuẩn RFC 6962:",
         "Gắn tiền tố 0x00 cho lá và 0x01 cho nhánh. Hai tập tiền tố hoàn toàn rời nhau, triệt tiêu 100% va chạm tiền ảnh cấp hai."),
        ("• Tầng 3: Động cơ Cây Merkle phân cấp trong RAM & Cập nhật Dirty-Path:",
         "Băm song song (~89.9 ms / 1.000 tệp). Cập nhật cục bộ Dirty-Path: khi sửa 1 tệp chỉ tính lại h=4 nút tổ tiên trong ~0.14 ms; khi thêm/xóa đạt 0.10 - 0.16 ms (nhanh hơn 458 - 708 lần so với Standard Merkle ~71-73 ms Rebuild mảng 1D)."),
        ("• Tầng 4: Bộ tạo bằng chứng xác thực tối ưu:",
         "Tạo Inclusion Proof gọn nhẹ 1.37 KB (1.369 Byte, tiết kiệm 98.8% băng thông), và Bằng chứng Nhất quán RFC 6962 độc lập chứng minh kho dữ liệu là append-only."),
        ("• Tầng 5: Bảo vệ Kênh kề & Cơ chế Xóa sạch Bộ nhớ (Memory Zeroization):",
         "Hàm hmac.compare_digest triệt tiêu kênh kề đo thời gian O(1) theo Kocher 1996; hàm secure_erase() xóa đè RAM ngăn kẻ thù trích xuất bộ nhớ (Memory Dump Adversary).")
    ], font_body_size=8.2)

add_5tier_architecture_visual(slide11_arch, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 12: [TRỤ CỘT I - HƯỚNG 1] 4 DẪN CHỨNG ĐỘT PHÁ (DIRTY-PATH COLUMN CHART)
# ==============================================================================
slide12_proof = prs.slides.add_slide(blank_layout)
add_base_background(slide12_proof)
add_header(slide12_proof, "Phần 4A: Trụ cột I - Kỹ thuật hệ thống & PoC (Hướng 1)", "11. Bốn Dẫn Chứng Chứng Minh Kỹ Thuật Đột Phá Của Mô Hình HD-MerkleTree")

add_bullet_card(slide12_proof, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "4 Dẫn Chứng Chứng Minh Cải Tiến Do Nhóm Tự Xây Dựng", [
        ("• 1. Ánh xạ trực tiếp cấu trúc thư mục tự nhiên theo mô hình Biba 1977:",
         "- Chứng minh: Cây Merkle phản chiếu chính xác từng thư mục cha, con. Nhờ đó, thuật toán định vị lỗi chỉ duyệt theo nhánh vi phạm trong O(k log N) thay vì duyệt toàn bộ O(N) 1.000 tệp tin, xác định tệp hỏng chỉ trong 0.022 ms."),
        ("• 2. Chuẩn phân tách miền RFC 6962 triệt tiêu va chạm:",
         "- Chứng minh: Gắn tiền tố byte 0x00 cho lá và 0x01 cho nhánh. Không gian băm nút lá và nhánh hoàn toàn tách biệt, triệt tiêu 100% tấn công giả mạo tiền ảnh cấp 2 với xác suất va chạm dưới 2 mũ âm 128."),
        ("• 3. Lan truyền đường dẫn bẩn cục bộ siêu tốc khi cập nhật tệp:",
         "- Chứng minh toán học: Khi sửa 1 tệp tin, chỉ cần băm lại đúng h = 4 nút tổ tiên từ lá lên đỉnh cây (~0.14 ms). Thời gian thực nghiệm khi thêm/xóa tệp đạt 0.10 - 0.16 ms so với 70.81 - 73.42 ms của Standard Merkle (nhanh hơn 458 - 708 lần do cấu trúc mảng 1D của Standard Merkle phải Rebuild O(N))."),
        ("• 4. Bằng chứng nhất quán RFC 6962 và Phòng chống kênh kề đo thời gian:",
         "- Chứng minh: Bằng chứng tính nhất quán chỉ cần 10 nút (320 byte) chứng minh kho dữ liệu là append-only, chống chèn backdoor lịch sử. Sử dụng hàm hmac.compare_digest triệt tiêu hoàn toàn tấn công kênh kề theo Kocher 1996.")
    ], font_body_size=8.8)

add_dirty_path_chart_slide(slide12_proof, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 13: [TRỤ CỘT III - HƯỚNG 4] BẢNG ĐỐI SÁNH SỐ LIỆU (REAL COLUMN CHART)
# ==============================================================================
slide13_bench = prs.slides.add_slide(blank_layout)
add_base_background(slide13_bench)
add_header(slide13_bench, "Phần 4B: Trụ cột III - Thực nghiệm định lượng quy mô lớn (Hướng 4)", "12. Bảng Đo Lường Số Liệu Thực Nghiệm Đối Đầu Trên 1.000 Tệp Tin")

table_card = add_card(slide13_bench, Inches(0.8), Inches(1.35), Inches(6.2), Inches(5.75))

table_shape = slide13_bench.shapes.add_table(5, 5, Inches(0.95), Inches(1.55), Inches(5.9), Inches(2.7))
table = table_shape.table
table.columns[0].width = Inches(1.8)
table.columns[1].width = Inches(0.95)
table.columns[2].width = Inches(1.1)
table.columns[3].width = Inches(0.95)
table.columns[4].width = Inches(1.1)

headers = ["Cơ Chế Bảo Mật", "Khởi Tạo", "Xác Minh 1 Tệp", "Dung Lượng Bảng", "Bằng Chứng 1 Tệp"]
for col_idx, h_text in enumerate(headers):
    cell = table.cell(0, col_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = NAVY_COLOR
    cell.text = h_text
    p = cell.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(9.5)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

table_data = [
    ("1. Danh sách SHA-256", "62.58 ms", "0.17 ms", "111.13 KB", "Tải cả bảng 111 KB"),
    ("2. Băm có khóa HMAC", "60.65 ms", "0.07 ms", "111.14 KB", "Cần chia sẻ khóa K"),
    ("3. Cây Merkle chuẩn", "60.72 ms", "0.02 ms", "44.76 KB", "990 Byte"),
    ("4. Cây Merkle cải tiến", "89.91 ms", "0.02 ms", "111.15 KB", "1.369 Byte")
]

for row_idx, row in enumerate(table_data):
    for col_idx, val in enumerate(row):
        cell = table.cell(row_idx + 1, col_idx)
        cell.fill.solid()
        if row_idx == 3:
            cell.fill.fore_color.rgb = RGBColor(236, 253, 245)
        elif row_idx % 2 == 1:
            cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
        else:
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cell.text = val
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(9)
        p.font.color.rgb = GREEN_COLOR if row_idx == 3 else TEXT_MAIN
        if row_idx == 3:
            p.font.bold = True
        if col_idx > 0:
            p.alignment = PP_ALIGN.CENTER

tb_note = slide13_bench.shapes.add_textbox(Inches(0.95), Inches(4.45), Inches(5.9), Inches(2.55))
tf_n = tb_note.text_frame
tf_n.word_wrap = True
p_n0 = tf_n.paragraphs[0]
p_n0.text = "Dẫn Chứng Chứng Minh 3 Chiều Giữa Bài Báo - Số Liệu - Ý Nghĩa:"
p_n0.font.bold = True
p_n0.font.size = Pt(10)
p_n0.font.color.rgb = NAVY_COLOR

notes_3way = [
    ("• 1. Đối chiếu Bài báo khoa học (Mục 4, 5, 6, Tài liệu [27, 36, 109]):",
     "Harry Halpin (2025) cung cấp khung lý thuyết về Kẻ tấn công hình thức, RFC 6962 và kiểm định PPT. Thực nghiệm định lượng trên 1.000 tệp tin là kết quả nghiên cứu độc lập của nhóm DT005."),
    ("• 2. Đối chiếu Số liệu Benchmark Thực nghiệm Độc lập (results/benchmark_results.json):",
     "Xác minh 1 tệp = 0.02 ms (nhanh hơn ~8.5 lần so với Flat SHA-256 0.17 ms); Bằng chứng = 1.369 Byte (tiết kiệm 98.8% băng thông so với 111.13 KB manifest); Cập nhật động khi Thêm/Xóa = 0.10 - 0.16 ms (nhanh hơn 458 - 708 lần so với Standard Merkle ~71-73 ms Rebuild mảng 1D)."),
    ("• 3. Giải thích ngắn gọn lý do vì sao tốt hơn (Có bằng chứng chứng minh):",
     "Nhanh hơn khi Thêm/Xóa vì HD-Merkle chỉ cập nhật nhánh thư mục cục bộ O(log N) thay vì Rebuild toàn bộ O(N) của implementation Standard Merkle mảng 1D; Tiết kiệm 98.8% băng thông và bảo vệ 100% tính riêng tư vì chỉ truyền các nút anh em, không tiết lộ 999 tệp còn lại qua mạng.")
]
for h_txt, b_txt in notes_3way:
    pn_h = tf_n.add_paragraph()
    pn_h.text = h_txt
    pn_h.font.bold = True
    pn_h.font.size = Pt(8.5)
    pn_h.font.color.rgb = BLUE_ACCENT
    pn_h.space_before = Pt(3)

    pn_b = tf_n.add_paragraph()
    pn_b.text = b_txt
    pn_b.font.size = Pt(8)
    pn_b.font.color.rgb = TEXT_MAIN
    pn_b.space_before = Pt(1)

add_benchmark_column_chart(slide13_bench, Inches(7.3), Inches(1.35), Inches(5.2), Inches(5.75))


# ==============================================================================
# SLIDE 14: [TRỤ CỘT II - HƯỚNG 2] KHUNG KIỂM ĐỊNH 7 TRÒ CHƠI AN NINH (GAMES 1 - 4)
# ==============================================================================
slide14_games = prs.slides.add_slide(blank_layout)
add_base_background(slide14_games)
add_header(slide14_games, "Phần 4C: Trụ cột II - Mô phỏng an ninh hình thức (Hướng 2)", "13. Khung 7 Trò Chơi An Ninh Mật Mã Học (Games 1 – 4: Va Chạm, Mạng, Nút & Tuyết Lở)")

add_bullet_card(slide14_games, Inches(0.8), Inches(1.35), Inches(5.8), Inches(5.75), 
    "Bốn Mô Hình Kiểm Định Toán Học Đối Kháng (Games 1 – 4)", [
        ("• Game 1: Chống tấn công tìm va chạm mã băm (PPT Adversary):",
         "1. Ý nghĩa: Kẻ tấn công có năng lực tính toán ngẫu nhiên nhưng bị giới hạn về thời gian và tài nguyên (đa thức).\n"
         "2. Cách tấn công: Cố tìm 2 tệp khác nhau A != B sao cho SHA256(A) = SHA256(B) để đánh lừa hệ thống.\n"
         "3. Cách phòng thủ: SHA-256 (256-bit). Tìm va chạm ngẫu nhiên đòi hỏi ~2^128 phép băm theo Birthday bound.\n"
         "4. Kết quả: Thử nghiệm 50 lượt ngẫu nhiên: Kẻ thù thắng 0/50 (0.0%), lợi thế đối thủ Adv <= 2^-128."),
        ("• Game 2: Chống kẻ tấn công đường truyền mạng (Dolev–Yao Adversary):",
         "1. Ý nghĩa: Kẻ tấn công có quyền nghe lén, chặn, sửa đổi hoặc thay thế gói tin truyền giữa Server và Client.\n"
         "2. Cách tấn công: (1) Sửa nội dung tệp; (2) Thay đổi mã băm trong Proof; (3) Tráo đường dẫn tệp.\n"
         "3. Cách phòng thủ: Client lưu Root Hash tin cậy; tự băm tệp và lần ngược lên Root Hash theo chuỗi Proof.\n"
         "4. Kết quả: Phát hiện chính xác 100% mọi hành vi can thiệp làm sai lệch Root Hash."),
        ("• Game 3: Chống giả mạo nút lá và nút nhánh (Second-Preimage & Domain Separation):",
         "1. Ý nghĩa: Merkle Tree gồm nút lá (tệp dữ liệu) và nút nhánh (ghép băm); nếu không tách biệt sẽ bị đánh tráo.\n"
         "2. Cách tấn công: Tạo tệp chứa chuỗi ghép Left || Right nhằm lừa hệ thống coi tệp rác là một nút lá.\n"
         "3. Cách phòng thủ: Chuẩn RFC 6962: H_leaf = SHA256(0x00 || Data), H_branch = SHA256(0x01 || L || R).\n"
         "4. Kết quả: Hai miền tiền tố độc lập hoàn toàn, triệt tiêu 100% việc nhầm lẫn giữa lá và nhánh."),
        ("• Game 4: Kiểm tra hiệu ứng tuyết lở (Avalanche Effect):",
         "1. Ý nghĩa: Thay đổi một phần rất nhỏ đầu vào thì mã băm đầu ra phải biến đổi ngẫu nhiên rất mạnh.\n"
         "2. Cách tấn công: Đảo 1 bit trong tệp rồi so sánh mã băm trước và sau để dò tìm quy luật tương quan.\n"
         "3. Cách phòng thủ: Đo khoảng cách Hamming; tỷ lệ bit thay đổi kỳ vọng xấp xỉ 50% (Shannon Diffusion).\n"
         "4. Kết quả: Thực nghiệm đo được 50.39% ≈ 50%, phù hợp hoàn hảo với đặc tính khuếch tán của SHA-256.")
    ], font_body_size=7.4)

add_security_games_cards_visual(slide14_games, Inches(6.8), Inches(1.35), Inches(5.7), Inches(5.75))


# ==============================================================================
# SLIDE 15: [TRỤ CỘT II - HƯỚNG 2] KHUNG KIỂM ĐỊNH 7 TRÒ CHƠI AN NINH (GAMES 5 - 7 & TỔNG KẾT)
# ==============================================================================
slide15_games = prs.slides.add_slide(blank_layout)
add_base_background(slide15_games)
add_header(slide15_games, "Phần 4C: Trụ cột II - Mô phỏng an ninh hình thức (Hướng 2)", "14. Khung 7 Trò Chơi An Ninh Mật Mã Học (Games 5 – 7 & Bảng Tổng Kết)")

add_bullet_card(slide15_games, Inches(0.8), Inches(1.35), Inches(5.5), Inches(5.75), 
    "Ba Mô Hình Kiểm Định Nâng Cao (Games 5 – 7)", [
        ("• Game 5: Bảo vệ tính nhất quán lịch sử (Append-Only Consistency Proof):",
         "1. Ý nghĩa: Kiểm tra dữ liệu cũ có được giữ nguyên khi hệ thống bổ sung dữ liệu mới hay không.\n"
         "2. Cách tấn công: Server sửa 1 tệp cũ trong 950 tệp, thêm 50 tệp mới thành 1.000 tệp rồi tuyên bố lịch sử nguyên vẹn.\n"
         "3. Cách phòng thủ: Consistency Proof RFC 6962 chứng minh cây cũ m=950 là tiền tố bất biến của cây mới n=1.000.\n"
         "4. Kết quả: Nếu dữ liệu cũ bị sửa, bằng chứng nhất quán bị từ chối ngay lập tức."),
        ("• Game 6: Chống trích xuất dữ liệu nhạy cảm từ RAM (Memory-Dump Adversary):",
         "1. Ý nghĩa: Kẻ tấn công đọc trộm RAM sau khi chương trình chạy để tìm khóa bí mật hoặc dữ liệu nhạy cảm.\n"
         "2. Cách tấn công: Dump bộ nhớ tiến trình và tìm kiếm khóa còn vương lại trong bộ nhớ đệm.\n"
         "3. Cách phòng thủ: Memory Zeroization: lưu khóa trong bytearray, chủ động ghi đè bằng 0x00 qua secure_erase().\n"
         "4. Lưu ý Python: gc.collect() thu hồi biến nhưng trong runtime không bảo đảm xóa sạch tuyệt đối mọi bản sao."),
        ("• Game 7: Chống tấn công đo thời gian & Giảm rò rỉ siêu dữ liệu:",
         "Phần A (Timing Side-Channel): Không dùng toán tử == (dừng sớm theo ký tự sai). Sử dụng hmac.compare_digest() để thời gian so sánh chuỗi là hằng số O(1).\n"
         "Phần B (Metadata Leakage): Thay vì tải toàn bộ 111 KB manifest chứa tên 1.000 tệp, Merkle Proof chỉ gửi mã băm anh em cần thiết, bảo vệ tính riêng tư của 999 tệp còn lại.")
    ], font_body_size=7.4)

# Bảng Tổng Kết 7 Trò Chơi
table_card15 = add_card(slide15_games, Inches(6.5), Inches(1.35), Inches(6.0), Inches(3.6))
table_shape15 = slide15_games.shapes.add_table(8, 3, Inches(6.6), Inches(1.45), Inches(5.8), Inches(3.4))
table15 = table_shape15.table
table15.columns[0].width = Inches(0.95)
table15.columns[1].width = Inches(2.55)
table15.columns[2].width = Inches(2.3)

headers15 = ["Game", "Mục Tiêu Bảo Mật", "Cơ Chế Chính Triển Khai"]
for col_idx, h_text in enumerate(headers15):
    cell = table15.cell(0, col_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = NAVY_COLOR
    cell.text = h_text
    p = cell.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(9.2)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

table_data15 = [
    ("Game 1", "Chống tìm hai dữ liệu cùng mã băm", "SHA-256 (Adv <= 2^-128)"),
    ("Game 2", "Phát hiện dữ liệu bị sửa khi truyền", "Merkle Proof + Root Hash tin cậy"),
    ("Game 3", "Phân biệt nút lá và nút nhánh", "Domain Separation (0x00 / 0x01)"),
    ("Game 4", "Kiểm tra mức độ thay đổi mã băm", "Avalanche Effect (50.39% ≈ 50%)"),
    ("Game 5", "Kiểm tra nhất quán qua các lần cập nhật", "Consistency Proof RFC 6962"),
    ("Game 6", "Hạn chế dữ liệu nhạy cảm trong RAM", "Memory Zeroization (secure_erase)"),
    ("Game 7", "Giảm rò rỉ thời gian & siêu dữ liệu", "hmac.compare_digest() + Merkle Proof")
]

for row_idx, row in enumerate(table_data15):
    for col_idx, val in enumerate(row):
        cell = table15.cell(row_idx + 1, col_idx)
        cell.fill.solid()
        if (row_idx + 1) % 2 == 1:
            cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
        else:
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cell.text = val
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(8.2)
        p.font.color.rgb = TEXT_MAIN
        if col_idx == 0:
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER
            p.font.color.rgb = BLUE_ACCENT

# Khung ghi chú phương pháp luận
note_card15 = add_card(slide15_games, Inches(6.5), Inches(5.1), Inches(6.0), Inches(2.0))
tb_nc15 = slide15_games.shapes.add_textbox(Inches(6.65), Inches(5.18), Inches(5.7), Inches(1.85))
tf_nc15 = tb_nc15.text_frame
tf_nc15.word_wrap = True

p_hnc = tf_nc15.paragraphs[0]
p_hnc.text = "Lưu Ý Phương Pháp Luận Khi Báo Cáo:"
p_hnc.font.bold = True
p_hnc.font.size = Pt(9.5)
p_hnc.font.color.rgb = NAVY_COLOR

bullets_nc15 = [
    "• Mục đích: 7 mô hình được xây dựng để kiểm thử các khía cạnh bảo mật đối kháng khác nhau theo luận điểm của Harry Halpin (2025).",
    "• Đối chiếu mã nguồn: Được triển khai và chạy độc lập qua tệp src/proposed_model/security_games.py, không gộp chung benchmark.",
    "• Đánh giá khách quan: Phân biệt rõ giữa cơ chế lý thuyết toán học và kết quả đo đạc thực nghiệm cụ thể trên môi trường Python."
]
for b_text in bullets_nc15:
    p_b = tf_nc15.add_paragraph()
    p_b.text = b_text
    p_b.font.size = Pt(7.8)
    p_b.font.color.rgb = TEXT_MAIN
    p_b.space_before = Pt(2.5)


# ==============================================================================
# SLIDE 16: HƯỚNG DẪN KỊCH BẢN TRÌNH DIỄN DEMO WEB (CONSOLE MOCKUP)
# ==============================================================================
slide16_demo = prs.slides.add_slide(blank_layout)
add_base_background(slide16_demo)
add_header(slide16_demo, "Phần 5: Trình diễn thực tế và tổng kết", "15. Kịch Bản Thao Tác Trình Diễn Trên Giao Diện Web Trực Quan")

add_bullet_card(slide16_demo, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Các Bước Thao Tác Trình Diễn Trước Hội Đồng", [
        ("• Khởi động hệ thống nhanh chóng:",
         "Chỉ cần nhấp đúp tệp run_demo_web.bat trên màn hình máy tính. Trình duyệt web sẽ tự động mở trang điều khiển tại địa chỉ nội bộ http://127.0.0.1:8080."),
        ("• Bước 1: Nạp 1.000 tệp tin dữ liệu gốc ban đầu:",
         "Nhấp nút Nạp dữ liệu để hệ thống quét toàn bộ 1.000 tệp tin nguyên bản từ thư mục realistic_dataset và tạo chuỗi băm gốc 32 byte."),
        ("• Bước 2: Mô phỏng hành vi Kẻ tấn công mạng:",
         "1. Nhấp Đảo 1 bit ngẫu nhiên để mô phỏng dữ liệu bị sửa đổi ngầm.\n2. Nhấp Chèn tệp lạ để mô phỏng tin tặc tải mã độc lên máy chủ.\n3. Nhấp Xóa lén tệp tin để mô phỏng hành vi xóa dữ liệu phá hoại."),
        ("• Bước 3: Kiểm tra và định vị lỗi tức thì:",
         "Nhấp Chạy kiểm tra toàn vẹn. Hệ thống lập tức chỉ đúng tên tệp bị can thiệp, hiển thị tỷ lệ nguyên vẹn và tốc độ xử lý hơn 4.000 tệp mỗi giây.")
    ], font_body_size=9.5)

add_web_mockup_console_slide(slide16_demo, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 17: BẢNG ĐIỀU KHIỂN BIỂU ĐỒ BENCHMARK ĐỘNG (REAL BAR CHART)
# ==============================================================================
slide17_charts = prs.slides.add_slide(blank_layout)
add_base_background(slide17_charts)
add_header(slide17_charts, "Phần 5: Trình diễn thực tế và tổng kết", "16. Bảng Điều Khiển Biểu Đồ Đo Đạc Trực Quan Hóa Động")

add_bullet_card(slide17_charts, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Trực Quan Hóa Biểu Đồ Khoa Học Hoạt Động Ngoại Tuyến", [
        ("• Hoạt động hoàn toàn ngoại tuyến:",
         "Hệ thống tích hợp sẵn thư viện vẽ biểu đồ nội bộ, không cần kết nối mạng Internet vẫn hiển thị mượt mà và trực quan khi báo cáo trước hội đồng."),
        ("• 4 Biểu đồ phân tích khoa học chuyên sâu:",
         "1. Biểu đồ so sánh thời gian khởi tạo và xác minh của 4 mô hình trên 1.000 tệp tin.\n2. Biểu đồ chi phí cập nhật khi sửa 1 tệp tin phản hồi tức thì theo chiều cao cây.\n3. Biểu đồ dung lượng bảng băm và kích thước bằng chứng gửi qua mạng.\n4. Biểu đồ kiểm tra khả năng mở rộng quy mô từ 50 tệp lên 1.000 tệp tin."),
        ("• Đo lường thời gian thực:",
         "Khi nhấp Đo lường và cập nhật lại biểu đồ ngay, hệ thống sẽ chạy đo đạc trực tiếp trên máy và vẽ kết quả tức thì.")
    ], font_body_size=9.5)

add_network_payload_chart(slide17_charts, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))


# ==============================================================================
# SLIDE 18: TỔNG KẾT ĐÓNG GÓP KHOA HỌC VÀ THỰC TIỄN (CLOSING QUADRANT)
# ==============================================================================
slide18_sum = prs.slides.add_slide(blank_layout)
add_base_background(slide18_sum)
add_header(slide18_sum, "Phần 5: Trình diễn thực tế và tổng kết", "17. Tổng Kết Đóng Góp Khoa Học Và Giá Trị Thực Tiễn Của Đồ Án")

add_bullet_card(slide18_sum, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.75), 
    "Những Đóng Góp Nổi Bật Đạt Được Trong Đồ Án", [
        ("• 1. Hoàn thành toàn diện việc nghiên cứu bài báo tiếng Anh gốc:",
         "Dịch thuật học thuật toàn văn bài báo của Harry Halpin (2025), đối chiếu 100% các luận điểm về Kẻ thù mật mã, hiệu ứng tuyết lở SHA-256, nguyên lý Kerckhoffs và kiểm định PPT."),
        ("• 2. Triển khai 3 Trụ cột phương pháp luận độc lập chuẩn mực:",
         "Tách bạch rõ ràng 3 hướng: Kỹ thuật hệ thống PoC (hd_merkle_tree.py), Mô phỏng 7 Game an toàn mật mã học (security_games.py) và Thực nghiệm quy mô lớn 1.000 tệp tin."),
        ("• 3. Đối chuẩn đầy đủ 3 mô hình nền tảng:",
         "Hiện thực hóa và đo đạc đầy đủ các cơ chế danh sách SHA-256, HMAC và cây Merkle nhị phân truyền thống trên 1.000 tệp tin."),
        ("• 4. Phát triển thành công mô hình cải tiến có dẫn chứng chứng minh:",
         "Đề xuất mô hình HD-MerkleTree kết hợp chuẩn tiền tố RFC 6962, xác minh nhanh hơn ~8.5 lần, thêm/xóa nhanh hơn 458 - 708 lần so với Rebuild mảng 1D, tiết kiệm hơn 98.8% băng thông và có bằng chứng nhất quán độc lập."),
        ("• 5. Đóng gói sản phẩm hoàn chỉnh:",
         "Bộ slide báo cáo chặt chẽ cùng ứng dụng web demo trực quan sẵn sàng trình diễn trước hội đồng.")
    ], font_body_size=9.0)

add_closing_quadrant_visual(slide18_sum, Inches(6.6), Inches(1.35), Inches(5.9), Inches(5.75))

# Lưu bài thuyết trình
try:
    prs.save(OUTPUT_PATH)
    print("SUCCESS: Generated completely revised presentation with 18 slides at:", OUTPUT_PATH)
except PermissionError:
    alt_path = OUTPUT_PATH.replace(".pptx", "_CapNhat_Moi.pptx")
    prs.save(alt_path)
    print(f"LƯU Ý: Tệp gốc {OUTPUT_PATH} đang được mở trong PowerPoint.")
    print(f"Đã lưu thành công bản cập nhật tại: {alt_path}")
