# -*- coding: utf-8 -*-
"""
Script Xây Dựng Slide Thuyết Trình Báo Cáo Giữa Kỳ Chuẩn Khoa Học:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn thiết kế:
- Khổ 16:9 widescreen (13.333 x 7.5 inch).
- Light Academic Theme (nền sáng, card slate/blue/emerald/amber, chữ sắc nét).
- 14 slide hoàn chỉnh theo đúng cấu trúc học thuật của bài báo.
- TUYỆT ĐỐI KHÔNG CÓ thang điểm [6-8], [9-10] của giảng viên.
- Có đầy đủ 4 sơ đồ vector khoa học minh họa lập luận của tác giả và phân định nghiên cứu.
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Thiết lập bảng màu chuẩn Light Academic
COLOR_BG_WHITE = RGBColor(255, 255, 255)
COLOR_BG_PANEL = RGBColor(248, 250, 252)       # Slate 50
COLOR_BG_CARD = RGBColor(255, 255, 255)
COLOR_BORDER = RGBColor(226, 232, 240)         # Slate 200
COLOR_BORDER_BLUE = RGBColor(191, 219, 254)    # Blue 200

COLOR_TEXT_TITLE = RGBColor(15, 23, 42)        # Slate 900
COLOR_TEXT_PRIMARY = RGBColor(30, 41, 59)      # Slate 800
COLOR_TEXT_SECONDARY = RGBColor(71, 85, 105)   # Slate 600
COLOR_TEXT_MUTED = RGBColor(100, 116, 139)     # Slate 500

COLOR_BRAND_NAVY = RGBColor(15, 23, 42)
COLOR_BRAND_BLUE = RGBColor(30, 64, 175)       # Royal Blue
COLOR_ACCENT_BLUE = RGBColor(37, 99, 235)      # Vibrant Blue
COLOR_ACCENT_EMERALD = RGBColor(5, 150, 105)   # Emerald 600
COLOR_ACCENT_AMBER = RGBColor(217, 119, 6)     # Amber 600
COLOR_ACCENT_ROSE = RGBColor(220, 38, 38)      # Rose 600

COLOR_BADGE_BG = RGBColor(239, 246, 255)       # Blue 50
COLOR_BADGE_TEXT = RGBColor(29, 78, 216)       # Blue 700

COLOR_EMERALD_BG = RGBColor(236, 253, 245)     # Emerald 50
COLOR_EMERALD_BORDER = RGBColor(167, 243, 208) # Emerald 200
COLOR_AMBER_BG = RGBColor(255, 251, 235)       # Amber 50
COLOR_AMBER_BORDER = RGBColor(253, 230, 138)   # Amber 200


def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def add_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG_WHITE
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_badge, subtitle_text=None):
        # Category badge container
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(4.5), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = COLOR_BADGE_BG
        badge.line.color.rgb = COLOR_BORDER_BLUE
        badge.line.width = Pt(1)
        tf_b = badge.text_frame
        tf_b.word_wrap = True
        p_b = tf_b.paragraphs[0]
        p_b.text = category_badge.upper()
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_BADGE_TEXT
        p_b.alignment = PP_ALIGN.CENTER

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.82), Inches(11.7), Inches(0.55))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(21)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_TITLE

        # Subtitle
        if subtitle_text:
            sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.38), Inches(11.7), Inches(0.35))
            tf_s = sub_box.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = tf_s.margin_right = tf_s.margin_top = tf_s.margin_bottom = 0
            p_s = tf_s.paragraphs[0]
            p_s.text = subtitle_text
            p_s.font.size = Pt(12)
            p_s.font.color.rgb = COLOR_TEXT_SECONDARY

    def add_card(slide, left, top, width, height, bg_color=COLOR_BG_PANEL, border_color=COLOR_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
        return card

    # =========================================================================
    # SLIDE 1: BÌA BÁO CÁO GIỮA KỲ
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    # Accent top banner line
    top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_BRAND_BLUE
    top_bar.line.fill.background()

    # Category badge
    b1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.7), Inches(4.8), Inches(0.4))
    b1.fill.solid()
    b1.fill.fore_color.rgb = COLOR_BADGE_BG
    b1.line.color.rgb = COLOR_BORDER_BLUE
    b1.line.width = Pt(1)
    tf1 = b1.text_frame
    p1 = tf1.paragraphs[0]
    p1.text = "BÁO CÁO TIẾN ĐỘ GIỮA KỲ — ĐỀ TÀI DT005"
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_BADGE_TEXT
    p1.alignment = PP_ALIGN.CENTER

    # Main Title
    t1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.25), Inches(11.333), Inches(1.3))
    tf_main = t1.text_frame
    tf_main.word_wrap = True
    p_m1 = tf_main.paragraphs[0]
    p_m1.text = "NGHIÊN CỨU BÀI BÁO:\nTHE ADVERSARY — THE PHILOSOPHY OF CRYPTOGRAPHY"
    p_m1.font.size = Pt(25)
    p_m1.font.bold = True
    p_m1.font.color.rgb = COLOR_TEXT_TITLE

    # Subtitle
    sub1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.7), Inches(11.333), Inches(0.6))
    tf_sub = sub1.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "Khảo cứu Triết học về Kẻ thù trong Mật mã học, Phê phán Giới hạn Hình thức hóa và Thí nghiệm Minh họa Độc lập"
    p_sub.font.size = Pt(14)
    p_sub.font.color.rgb = COLOR_BRAND_BLUE

    # Card 1: Paper Meta Card
    add_card(s1, Inches(1.0), Inches(3.45), Inches(5.5), Inches(3.3), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
    t_card1 = s1.shapes.add_textbox(Inches(1.2), Inches(3.6), Inches(5.1), Inches(3.0))
    tf_c1 = t_card1.text_frame
    tf_c1.word_wrap = True
    
    p = tf_c1.paragraphs[0]
    p.text = "THÔNG TIN BÀI BÁO THAM CHIẾU"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE
    
    bullets_c1 = [
        ("Tạp chí:", "Journal of Cybersecurity, Oxford Academic (2025)"),
        ("Tác giả:", "TS. Harry Halpin (Nym Technologies, Thụy Sĩ)"),
        ("DOI:", "10.1093/cybsec/tyaf006 (Volume 11, Issue 1)"),
        ("Bản chất:", "Nghiên cứu triết học, lịch sử & nhận thức luận mật mã"),
        ("Tài nguyên:", "Bài báo thuần lý thuyết; 0% code, 0% benchmark thực nghiệm"),
        ("Định hướng:", "Phê phán trừu tượng hóa & đề xuất 4 câu hỏi nghiên cứu mở")
    ]
    for lbl, val in bullets_c1:
        p = tf_c1.add_paragraph()
        p.space_before = Pt(5)
        run_lbl = p.add_run()
        run_lbl.text = f"• {lbl} "
        run_lbl.font.bold = True
        run_lbl.font.size = Pt(11)
        run_lbl.font.color.rgb = COLOR_TEXT_PRIMARY
        run_val = p.add_run()
        run_val.text = val
        run_val.font.size = Pt(11)
        run_val.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Student / Course Meta Card
    add_card(s1, Inches(6.8), Inches(3.45), Inches(5.5), Inches(3.3), COLOR_EMERALD_BG, COLOR_EMERALD_BORDER)
    t_card2 = s1.shapes.add_textbox(Inches(7.0), Inches(3.6), Inches(5.1), Inches(3.0))
    tf_c2 = t_card2.text_frame
    tf_c2.word_wrap = True
    
    p = tf_c2.paragraphs[0]
    p.text = "THÔNG TIN HỌC PHẦN & ĐỀ TÀI DT005"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_EMERALD
    
    bullets_c2 = [
        ("Trường:", "Đại học Tài nguyên và Môi trường TP. Hồ Chí Minh"),
        ("Khoa / Ngành:", "Khoa Công nghệ Thông tin"),
        ("Học phần:", "An toàn và Bảo mật Hệ thống Thông tin"),
        ("Đề tài DT005:", "Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC, Merkle Tree"),
        ("Đóng góp nhóm:", "Kiểm toán học thuật toàn văn bài báo & Xây dựng bộ thí nghiệm minh họa thực nghiệm độc lập (Red Team vs Blue Team)"),
        ("Phạm vi giữa kỳ:", "Báo cáo lý thuyết bài báo + Demo đối kháng minh họa")
    ]
    for lbl, val in bullets_c2:
        p = tf_c2.add_paragraph()
        p.space_before = Pt(5)
        run_lbl = p.add_run()
        run_lbl.text = f"• {lbl} "
        run_lbl.font.bold = True
        run_lbl.font.size = Pt(11)
        run_lbl.font.color.rgb = COLOR_TEXT_PRIMARY
        run_val = p.add_run()
        run_val.text = val
        run_val.font.size = Pt(11)
        run_val.font.color.rgb = COLOR_TEXT_SECONDARY

    # =========================================================================
    # SLIDE 2: BỐI CẢNH & VẤN ĐỀ NGHIÊN CỨU (SEC 1, TR. 1–2)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    add_header(s2, "BỐI CẢNH BÀI BÁO & VẤN ĐỀ NGHIÊN CỨU", "SECTION 1 — GIỚI THIỆU (TRANG 1–2)",
               "Khảo sát mối quan hệ giữa Mật mã học, Kẻ thù và Triết học về Thông tin")

    # 2 Big Cards
    add_card(s2, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_BG_PANEL, COLOR_BORDER)
    t = s2.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "1. Mật Mã Học Trong Kỷ Nguyên Số"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE

    points_s2_left = [
        ("Bản chất cốt lõi:", "Mật mã học là ngành khoa học về kỹ thuật truyền thông an toàn trong sự hiện diện của các đối thủ có chủ đích (adversaries)."),
        ("Không gian vô tận:", "Sự bùng nổ của Internet toàn cầu đã biến không gian số thành môi trường tương tác với vô vàn đối thủ tiềm tàng, xa lạ và phân tán (distal agents)."),
        ("Niềm tin và Bí mật:", "Trong xã hội hiện đại, niềm tin giữa các bên không quen biết không còn dựa trên gia tộc hay bộ lạc nhỏ, mà phải dựa trên các thủ tục quan liêu, hệ thống giám sát và các công cụ toán học mật mã bảo vệ bí mật.")
    ]
    for h, b in points_s2_left:
        p = tf.add_paragraph()
        p.space_before = Pt(10)
        r1 = p.add_run()
        r1.text = f"{h} "
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p.add_run()
        r2.text = b
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    add_card(s2, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_AMBER_BG, COLOR_AMBER_BORDER)
    t = s2.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "2. Nghịch Lý Thiếu Vắng Khảo Sát Triết Học"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_AMBER

    points_s2_right = [
        ("Nghịch lý nhận thức:", "Lý thuyết thông tin của Claude Shannon (1948) đã được triết học phân tích sâu sắc (Bar-Hillel, Dretske, Floridi), nhưng chính mật mã học — cái nôi khai sinh lý thuyết thông tin — lại bị triết học bỏ quên."),
        ("Tuyên bố cách mạng:", "Mật mã học đưa ra một tuyên bố mang tính triết học cách mạng: 'Bí mật có thể được truyền đưa an toàn trong thế giới kỹ thuật số thông qua việc ứng dụng toán học'."),
        ("Lỗ hổng thực tiễn:", "Khi mật mã học chuyển mình thành khoa học toán thuần túy thông qua sự trừu tượng hóa, nó tạo ra một điểm mù nguy hiểm trước các đối thủ trong thế giới thực.")
    ]
    for h, b in points_s2_right:
        p = tf.add_paragraph()
        p.space_before = Pt(10)
        r1 = p.add_run()
        r1.text = f"{h} "
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p.add_run()
        r2.text = b
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Footer note
    f_box = s2.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn khoa học: Harry Halpin (2025), Section 1: Introduction, Journal of Cybersecurity, tr. 1–2."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 3: MỤC TIÊU & ĐÓNG GÓP HỌC THUẬT CỦA HALPIN (SEC 1–2, TR. 1–3)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    add_header(s3, "MỤC TIÊU & BA ĐÓNG GÓP HỌC THUẬT CỦA BÀI BÁO", "SECTION 1 & 2 (TRANG 1–3)",
               "Ba trụ cột đóng góp chính của công trình Harry Halpin đối với mật mã học và an ninh mạng")

    # 3 Column Cards
    col_w = Inches(3.7)
    gap = Inches(0.3)
    
    pillars = [
        ("Trụ Cột 1: Cầu Nối Hai Chiều",
         "KẾT NỐI TRIẾT HỌC & MẬT MÃ",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         [("Mục tiêu 1:", "Làm cho các hiểu biết triết học trở nên sáng tỏ và hữu ích đối với các nhà mật mã học."),
          ("Mục tiêu 2:", "Kêu gọi các triết gia áp dụng công cụ triết học vào mật mã học thay vì chỉ dừng lại ở lý thuyết thông tin thuần túy."),
          ("Ý nghĩa:", "Mở rộng mật mã học từ một nhánh toán học thuần túy thành một 'khoa học an ninh tổng quát' (general science of security).")]),
        ("Trụ Cột 2: Lịch Sử Hình Thức Hóa",
         "TỪ NGHỆ THUẬT SANG KHOA HỌC",
         COLOR_BG_PANEL, COLOR_BORDER, COLOR_TEXT_TITLE,
         [("Tiến trình lịch sử:", "Khảo cứu quá trình mật mã học chuyển mình từ một môn 'nghệ thuật bí truyền' thành một ngành 'khoa học chính xác'."),
          ("Toán học hóa đối thủ:", "Quá trình này đạt được chủ yếu nhờ việc mô hình hóa và hình thức hóa đối thủ (Shannon, PPT Game, Dolev-Yao, Random Oracle)."),
          ("Phân tích nhận thức luận:", "Làm rõ các giả định triết học ngầm định đằng sau từng bước hình thức hóa.")]),
        ("Trụ Cột 3: Phê Phán Trừu Tượng Hóa",
         "PHÊ PHÁN GIỚI HẠN MÔ HÌNH",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER,
         [("Bản đồ vs Lãnh thổ:", "Vận dụng luận điểm của Korzybski và Cantwell Smith: Sự trừu tượng hóa luôn đi kèm cái giá là lược bỏ chi tiết thực tế."),
          ("Khủng hoảng Snowden:", "Vụ rò rỉ Snowden (2013) là bằng chứng thực tế cho thấy các mô hình hình thức thất bại trước đối thủ thể chế (NSA)."),
          ("Lời giải:", "Đề xuất chương trình nghiên cứu 'hòa giải' (reconciliation) với dữ liệu thực tế và mạng xã hội-kỹ thuật phức hợp.")])
    ]

    for idx, (p_title, p_badge, bg_c, brd_c, text_c, items) in enumerate(pillars):
        left_pos = Inches(0.8) + idx * (col_w + gap)
        add_card(s3, left_pos, Inches(1.85), col_w, Inches(4.8), bg_c, brd_c)
        t = s3.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.0), col_w - Inches(0.4), Inches(4.5))
        tf = t.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_c

        sub_p = tf.add_paragraph()
        sub_p.text = p_badge
        sub_p.font.size = Pt(10)
        sub_p.font.bold = True
        sub_p.font.color.rgb = COLOR_TEXT_MUTED
        sub_p.space_after = Pt(8)

        for h, b in items:
            p_it = tf.add_paragraph()
            p_it.space_before = Pt(6)
            r1 = p_it.add_run()
            r1.text = f"• {h} "
            r1.font.bold = True
            r1.font.size = Pt(11)
            r1.font.color.rgb = COLOR_TEXT_PRIMARY
            r2 = p_it.add_run()
            r2.text = b
            r2.font.size = Pt(10.5)
            r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s3.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn khoa học: Harry Halpin (2025), Section 1–2, Journal of Cybersecurity, tr. 1–3."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 4: CƠ SỞ TRIẾT HỌC: DRETSKE, CANTWELL SMITH & BÍ MẬT (SEC 2, TR. 2–3)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4)
    add_header(s4, "CƠ SỞ TRIẾT HỌC: DRETSKE, CANTWELL SMITH & BẢN CHẤT BÍ MẬT", "SECTION 2 — TIỀN ĐỀ TRIẾT HỌC (TRANG 2–3)",
               "Xây dựng khung nhận thức luận về tác nhân có chủ đích, sự phân ranh thế giới và khái niệm bí mật")

    # 3 Horizontal Panels
    h_cards = [
        ("1. Fred Dretske (1981) — Dòng Chảy Thông Tin và Tác Nhân Có Chủ Đích",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         "Thông tin là các quy luật (regularities) trong các trạng thái sự vật. Tác nhân có chủ đích (như con người) khác biệt hoàn toàn với các vật mang tin vô tri (nhiệt kế, DNA): tác nhân sở hữu cấu trúc ý niệm, niềm tin, kế hoạch hành động và có khả năng tích hợp thông tin vào trạng thái cục bộ để ra quyết định."),
        ("2. Brian Cantwell Smith (1996) — Siêu Hình Học Dòng Chảy & Sự Phân Ranh (Cutting at the Joints)",
         COLOR_BG_PANEL, COLOR_BORDER, COLOR_TEXT_TITLE,
         "Thế giới là dòng chảy siêu hình nguyên thủy (primordial metaphysical flux). Con người phân tích thế giới bằng cách 'cắt' nó thành các hệ thống riêng biệt (proximal vs distal agents). Mọi mô hình trừu tượng đều là phép xấp xỉ; ranh giới giữa các tác nhân luôn mang tính động và tương đối."),
        ("3. Bản Chất Của Bí Mật: Bí Mật Là Gì Dưới Góc Nhìn Triết Học?",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD,
         "Bí mật KHÔNG tồn tại tự thân và không có bí mật vĩnh cửu. Bí mật là mối quan hệ xã hội tương đối giữa thông tin và nhóm người giữ bí mật (secret keepers) qua một khoảng thời gian nhất định, trong đó nhóm người này chủ đích loại trừ những kẻ khác (adversaries) khỏi quyền tiếp cận thông tin đó.")
    ]

    for idx, (title_h, bg_c, brd_c, text_c, content_str) in enumerate(h_cards):
        top_pos = Inches(1.85) + idx * Inches(1.58)
        add_card(s4, Inches(0.8), top_pos, Inches(11.733), Inches(1.42), bg_c, brd_c)
        t = s4.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.12), Inches(11.333), Inches(1.18))
        tf = t.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title_h
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_c

        p_desc = tf.add_paragraph()
        p_desc.space_before = Pt(4)
        p_desc.text = content_str
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s4.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn: Dretske (1981, Ref [1]); Cantwell Smith (1996, Ref [5]); Halpin (2025), tr. 2–3."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: SƠ ĐỒ 1 — DÒNG THỜI GIAN TIẾN HÓA MÔ HÌNH ĐỐI THỦ
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5)
    add_header(s5, "SƠ ĐỒ 1: DÒNG THỜI GIAN 7 THẬP KỶ TIẾN HÓA MÔ HÌNH ĐỐI THỦ", "SECTIONS 3–6 (TRANG 3–12)",
               "Tiến trình lịch sử từ mô hình truyền thông Shannon đến các đối thủ hình thức và khủng hoảng Snowden")

    # SƠ ĐỒ VECTOR: 7 Mốc Thời Gian
    # Vẽ thanh trục thời gian ngang
    axis = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(4.1), Inches(11.333), Inches(0.08))
    axis.fill.solid()
    axis.fill.fore_color.rgb = COLOR_BORDER_BLUE
    axis.line.fill.background()

    timeline_nodes = [
        ("1948", "Shannon", "Nhiễu ngẫu nhiên, One-Time Pad, Bí mật hoàn hảo", True),
        ("1976", "Diffie-Hellman", "Mật mã khóa công khai, Kênh truyền mở", False),
        ("1984", "Goldwasser-Micali", "Mô hình PPT Adversary, Security Games", True),
        ("1983", "Dolev-Yao", "Đối thủ mạng kiểm soát kênh, Hộp đen mật mã", False),
        ("1993", "Bellare-Rogaway", "Random Oracle Model (ROM), Băm lý tưởng", True),
        ("1996", "Paul Kocher", "Phê phán: Tấn công kênh kề (Side-Channels)", False),
        ("2013", "Snowden / NSA", "Đối thủ thể chế, Giám sát siêu dữ liệu", True)
    ]

    node_w = Inches(1.48)
    gap_x = Inches(0.18)

    for i, (yr, author, desc, is_top) in enumerate(timeline_nodes):
        x = Inches(1.0) + i * (node_w + gap_x)
        
        # Circle on axis
        circle = s5.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.59), Inches(3.99), Inches(0.3), Inches(0.3))
        circle.fill.solid()
        circle.fill.fore_color.rgb = COLOR_BRAND_BLUE if i < 5 else (COLOR_ACCENT_AMBER if i == 5 else COLOR_ACCENT_ROSE)
        circle.line.color.rgb = COLOR_BG_WHITE
        circle.line.width = Pt(2)

        # Card position: Alternate top and bottom
        card_y = Inches(1.9) if is_top else Inches(4.5)
        card_h = Inches(1.85)

        card_bg = COLOR_BG_PANEL if i < 5 else (COLOR_AMBER_BG if i == 5 else RGBColor(254, 242, 242))
        card_border = COLOR_BORDER if i < 5 else (COLOR_AMBER_BORDER if i == 5 else RGBColor(254, 202, 202))

        add_card(s5, x, card_y, node_w, card_h, card_bg, card_border)
        
        tb = s5.shapes.add_textbox(x + Inches(0.08), card_y + Inches(0.08), node_w - Inches(0.16), card_h - Inches(0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = yr
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_BRAND_BLUE if i < 5 else (COLOR_ACCENT_AMBER if i == 5 else COLOR_ACCENT_ROSE)

        p2 = tf.add_paragraph()
        p2.text = author
        p2.font.size = Pt(10)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_TEXT_TITLE

        p3 = tf.add_paragraph()
        p3.space_before = Pt(3)
        p3.text = desc
        p3.font.size = Pt(8.5)
        p3.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s5.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ tổng hợp của người trình bày dựa trên bài báo Harry Halpin (2025), Sections 3–6, tr. 3–12."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 6: SHANNON & GIỚI HẠN MÔ HÌNH NHIỄU (SƠ ĐỒ 2: MÔ HÌNH TRUYỀN THÔNG)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6)
    add_header(s6, "SHANNON VÀ GIỚI HẠN CỦA MÔ HÌNH NHIỄU", "SECTION 2 & 3 (TRANG 3–5)",
               "Đối sánh mô hình truyền thông điểm-điểm của Shannon và mạng xã hội-kỹ thuật phức hợp")

    # SƠ ĐỒ VECTOR: Diagram 2 (Shannon vs Complex Network)
    # Top Panel: Shannon Point-to-Point
    add_card(s6, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
    t = s6.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Mô hình Điểm - Điểm Cổ Điển của Shannon"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE

    # Draw diagram 2a inside
    p_box = tf.add_paragraph()
    p_box.space_before = Pt(8)
    p_box.text = "[Sender Alice]  ──────>  [Channel: Key + Ciphertext]  ──────>  [Receiver Bob]\n                                    ▲\n                                    │\n                         [Adversary Eve (Passive)]"
    p_box.font.size = Pt(9.5)
    p_box.font.name = "Consolas"
    p_box.font.color.rgb = COLOR_TEXT_PRIMARY

    shannon_pts = [
        ("Nguồn gốc mật mã:", "Lý thuyết thông tin Shannon (1948) thực chất bắt nguồn từ báo cáo mật mã tối mật năm 1945 ('A Mathematical Theory of Cryptography')."),
        ("Bí mật hoàn hảo:", "Shannon chứng minh hệ mật hoàn hảo duy nhất là One-Time Pad, nhưng đòi hỏi khóa ngẫu nhiên dài bằng thông điệp ($H(K) >= H(M)$)."),
        ("Giả định giới hạn:", "Shannon xem sự sai lệch thông tin là 'nhiễu ngẫu nhiên' (random noise). Tuy nhiên, kẻ thù mật mã là tác nhân có chủ đích (intentional agent), không phải nhiễu ngẫu nhiên!")
    ]
    for h, b in shannon_pts:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(6)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Panel: Complex Socio-Technical Network
    add_card(s6, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_EMERALD_BG, COLOR_EMERALD_BORDER)
    t = s6.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Thực Tế: Mạng Xã Hội - Kỹ Thuật Không Đồng Nhất"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_EMERALD

    p_box2 = tf.add_paragraph()
    p_box2.space_before = Pt(8)
    p_box2.text = "[Nhiều Bên Giữ Bí Mật]  ◄──►  [Hạ Tầng Viễn Thông / Đám Mây]\n           ▲                                     ▲\n           │                                     │\n[Kẻ Tấn Công Thể Chế]   ◄──►  [Lỗ Hổng Triển Khai / Kênh Kề]"
    p_box2.font.size = Pt(9.5)
    p_box2.font.name = "Consolas"
    p_box2.font.color.rgb = COLOR_TEXT_PRIMARY

    complex_pts = [
        ("Vượt qua kênh đơn:", "Trong thực tế, thông tin không đi qua 1 kênh duy nhất mà chảy qua mạng lưới phức tạp gồm con người, thiết bị, hệ điều hành và nhà mạng."),
        ("Tổ chức thông tin:", "Bảo toàn bí mật là duy trì bất biến chính sách (policy invariant) giữa nhóm người giữ bí mật trước các lỗi Byzantine và hành vi phản bội."),
        ("Điểm mù của Shannon:", "Mô hình Shannon bỏ qua hoàn toàn chi phí phân phối khóa, lưu trữ khóa và sự can thiệp trực tiếp vào hành vi của bên giữ bí mật.")
    ]
    for h, b in complex_pts:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(6)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s6.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ tổng hợp dựa trên Shannon (1948, Ref [3]) và Halpin (2025), Section 2–3, tr. 3–5."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 7: CÁC MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI (SEC 4, TR. 5–8)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7)
    add_header(s7, "BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI", "SECTION 4 — HÌNH THỨC HÓA ĐỐI THỦ (TRANG 5–8)",
               "Tiến trình toán học hóa kẻ thù: Độ phức tạp tính toán đa thức, Logic kênh truyền và Tiên tri ngẫu nhiên")

    col_w7 = Inches(3.7)
    gap7 = Inches(0.3)

    formal_models = [
        ("1. Mô Hình PPT Adversary", "GOLDWASSER-MICALI (1984)",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         [("Khái niệm:", "Đối thủ là thuật toán Turing xác suất chạy trong thời gian đa thức (Probabilistic Polynomial-Time)."),
          ("Security Games:", "Định nghĩa an toàn thông qua các trò chơi thách đấu (Challenger vs Adversary). Lợi thế Adv(A) phải là hàm không đáng kể (negligible)."),
          ("Quy dẫn toán học:", "Chứng minh an toàn bằng cách quy dẫn rằng việc phá vỡ hệ mật tương đương với giải các bài toán toán học khó (phân tích thừa số nguyên, logarit rời rạc).")]),
        ("2. Mô Hình Đối Thủ Mạng Logic", "DOLEV-YAO (1983)",
         COLOR_BG_PANEL, COLOR_BORDER, COLOR_TEXT_TITLE,
         [("Khái niệm:", "Đối thủ là thực thể kiểm soát toàn bộ kênh truyền mạng: có quyền đọc, chặn, sửa đổi, xóa, tráo đổi và phát lại gói tin."),
          ("Hộp đen mật mã:", "Mô hình Dolev-Yao xem các thuật toán mật mã như các hàm đại số hoàn hảo (Perfect Black-box). Kẻ thù không thể giải mã nếu không có khóa."),
          ("Ứng dụng:", "Kiểm chứng hình thức các giao thức mạng phức tạp thông qua các công cụ logic tự động.")]),
        ("3. Mô Hình Tiên Tri Ngẫu Nhiên", "BELLARE-ROGAWAY (1993)",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER,
         [("Khái niệm:", "Giả định hàm băm thực tế hoạt động như một hàm ngẫu nhiên lý tưởng (Random Oracle Model - ROM) trả về chuỗi ngẫu nhiên đồng đều."),
          ("Vai trò thực tế:", "Cho phép xây dựng các hệ mật khóa công khai và chữ ký số hiệu năng cao có chứng minh an toàn."),
          ("Giới hạn triết học:", "Trong thế giới thực, không có hàm băm cụ thể nào (kể cả SHA-256) có thể tương đương toán học với một Random Oracle thực sự.")])
    ]

    for idx, (m_title, m_badge, bg_c, brd_c, text_c, items) in enumerate(formal_models):
        left_pos = Inches(0.8) + idx * (col_w7 + gap7)
        add_card(s7, left_pos, Inches(1.85), col_w7, Inches(4.8), bg_c, brd_c)
        t = s7.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.0), col_w7 - Inches(0.4), Inches(4.5))
        tf = t.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = m_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_c

        sub_p = tf.add_paragraph()
        sub_p.text = m_badge
        sub_p.font.size = Pt(9.5)
        sub_p.font.bold = True
        sub_p.font.color.rgb = COLOR_TEXT_MUTED
        sub_p.space_after = Pt(8)

        for h, b in items:
            p_it = tf.add_paragraph()
            p_it.space_before = Pt(6)
            r1 = p_it.add_run()
            r1.text = f"• {h} "
            r1.font.bold = True
            r1.font.size = Pt(11)
            r1.font.color.rgb = COLOR_TEXT_PRIMARY
            r2 = p_it.add_run()
            r2.text = b
            r2.font.size = Pt(10)
            r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s7.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn: Goldwasser-Micali (1984, Ref [34]); Dolev-Yao (1983, Ref [52]); Bellare-Rogaway (1993, Ref [40]); Halpin (2025), tr. 5–8."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 8: GIỚI HẠN CỦA HÌNH THỨC HÓA (SƠ ĐỒ 3: KHOẢNG CÁCH MÔ HÌNH VÀ THỰC TẾ)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8)
    add_header(s8, "GIỚI HẠN CỦA SỰ HÌNH THỨC HÓA TRONG MẬT MÃ HỌC", "SECTION 5 — GIỚI HẠN TRỪU TƯỢNG HÓA (TRANG 8–10)",
               "Khoảng cách giữa các giả định toán học lý tưởng và thế giới vật lý thực tế")

    # SƠ ĐỒ VECTOR: Diagram 3 (4 tầng luồng khai thác)
    flow_steps = [
        ("BƯỚC 1: GIẢ ĐỊNH TOÁN HỌC HÌNH THỨC",
         "Hộp đen mật mã lý tưởng, thời gian thực thi là vô hình, độ dài bản mã công khai không rò rỉ nội dung, kênh truyền điểm-điểm độc lập.",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE),
        ("BƯỚC 2: CÁC YẾU TỐ VẬT LÝ BỊ TRỪU TƯỢNG HÓA LƯỢC BỎ",
         "Thời gian vi xử lý CPU, dao động điện năng tiêu thụ, phát xạ điện từ, kiến trúc bộ nhớ đệm (Cache), lỗi lập trình phần mềm.",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER),
        ("BƯỚC 3: CON ĐƯỜNG KHAI THÁC CỦA ĐỐI THỦ THẾ GIỚI THỰC",
         "Tấn công kênh kề (Timing Attack - Kocher 1996), phân tích lưu lượng siêu dữ liệu, tấn công tiền ảnh cấp hai cấu trúc (Node-Leaf Confusion).",
         RGBColor(254, 242, 242), RGBColor(254, 202, 202), COLOR_ACCENT_ROSE),
        ("BƯỚC 4: KẾT LUẬN NHẬN THỨC LUẬN CỦA HARRY HALPIN",
         "Cái giá của sự trừu tượng hóa là làm cho mật mã học mù quáng trước chính những năng lực thực tế của đối thủ nằm ngoài mô hình lý thuyết.",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD)
    ]

    for idx, (step_t, step_d, bg_c, brd_c, text_c) in enumerate(flow_steps):
        top_pos = Inches(1.85) + idx * Inches(1.18)
        add_card(s8, Inches(0.8), top_pos, Inches(11.733), Inches(1.05), bg_c, brd_c)
        t = s8.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.08), Inches(11.333), Inches(0.9))
        tf = t.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = step_t
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = text_c

        p_d = tf.add_paragraph()
        p_d.space_before = Pt(3)
        p_d.text = step_d
        p_d.font.size = Pt(10.5)
        p_d.font.color.rgb = COLOR_TEXT_PRIMARY

    f_box = s8.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ tổng hợp dựa trên Kocher (1996, Ref [36]) và Halpin (2025), Section 5, tr. 8–10."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 9: SNOWDEN & ĐỐI THỦ THỂ CHẾ (SEC 6, TR. 10–12)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9)
    add_header(s9, "SNOWDEN VÀ ĐỐI THỦ THỂ CHẾ (THE REAL ADVERSARY)", "SECTION 6 — TIẾT LỘ CỦA SNOWDEN (TRANG 10–12)",
               "Khủng hoảng mô hình an ninh trước đối thủ giám sát hàng loạt quy mô quốc gia")

    col_w9 = Inches(3.7)
    gap9 = Inches(0.3)

    snowden_pillars = [
        ("1. Khủng Hoảng Thực Tiễn", "TIẾT LỘ SNOWDEN (2013)",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         [("Sự kiện bước ngoặt:", "Các tài liệu Edward Snowden tiết lộ về NSA đã lột bỏ tấm mặt nạ, kéo mật mã học vào một cuộc khủng hoảng thực tiễn sâu sắc."),
          ("Không phá toán học:", "NSA không cần giải các bài toán toán học khó của RSA hay AES mà tấn công trực tiếp vào toàn bộ hệ sinh thái xung quanh."),
          ("Mô hình thực tế:", "Đối thủ bị động toàn cầu (Global Passive Adversary) và tấn công chủ động vào kênh truyền là mối đe dọa hoàn toàn có thật.")]),
        ("2. Năng Lực Đối Thủ Thể Chế", "VƯỢT NGOÀI HÌNH THỨC",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER,
         [("Giám sát Siêu dữ liệu:", "NSA tập trung khai thác Metadata (ai liên lạc với ai, thời gian, kích thước) để dựng lại mạng xã hội mà không cần bẻ khóa bản mã."),
          ("Tấn công Chuỗi cung ứng:", "Chèn cửa sau vào phần cứng, làm suy yếu chuẩn mật mã quốc tế (tiêu chuẩn NIST Dual_EC_DRBG bị cài cửa sau ngầm)."),
          ("Áp lực Thể chế:", "Cưỡng chế pháp lý các tập đoàn công nghệ lớn bàn giao khóa bí mật và quyền truy cập dữ liệu máy chủ.")]),
        ("3. Tính Tương Đối Của Kẻ Thù", "BẢN CHẤT NHẬN THỨC LUẬN",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD,
         [("Khái niệm tương đối:", "Khái niệm 'Kẻ thù' mang tính vị thế: Đối với người dân, NSA là đối thủ giám sát; nhưng đối với NSA, Snowden lại là đối thủ (insider) làm rò rỉ bí mật."),
          ("Dân chủ hóa mật mã:", "Mật mã khóa công khai từng giúp cá nhân tự vệ, nhưng công cụ gián điệp của NSA nay đã lan rộng sang tội phạm mạng và các nhóm phi nhà nước."),
          ("Hệ quả khoa học:", "Mật mã học buộc phải từ bỏ giả định kênh truyền sạch để nghiên cứu phòng chống cửa sau (Cliptography).")])
    ]

    for idx, (s_title, s_badge, bg_c, brd_c, text_c, items) in enumerate(snowden_pillars):
        left_pos = Inches(0.8) + idx * (col_w9 + gap9)
        add_card(s9, left_pos, Inches(1.85), col_w9, Inches(4.8), bg_c, brd_c)
        t = s9.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.0), col_w9 - Inches(0.4), Inches(4.5))
        tf = t.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = s_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_c

        sub_p = tf.add_paragraph()
        sub_p.text = s_badge
        sub_p.font.size = Pt(9.5)
        sub_p.font.bold = True
        sub_p.font.color.rgb = COLOR_TEXT_MUTED
        sub_p.space_after = Pt(8)

        for h, b in items:
            p_it = tf.add_paragraph()
            p_it.space_before = Pt(6)
            r1 = p_it.add_run()
            r1.text = f"• {h} "
            r1.font.bold = True
            r1.font.size = Pt(11)
            r1.font.color.rgb = COLOR_TEXT_PRIMARY
            r2 = p_it.add_run()
            r2.text = b
            r2.font.size = Pt(10)
            r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s9.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn khoa học: Harry Halpin (2025), Section 6: Snowden and the real adversary, tr. 10–12."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 10: ĐÓNG GÓP & 4 CÂU HỎI NGHIÊN CỨU MỞ CỦA HALPIN (SEC 7, TR. 12–14)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10)
    add_header(s10, "CHƯƠNG TRÌNH NGHIÊN CỨU TƯƠNG LAI: BỐN CÂU HỎI MỞ", "SECTION 7 — KẾT LUẬN & CHƯƠNG TRÌNH NGHIÊN CỨU (TRANG 12–14)",
               "Đề xuất hòa giải giữa trừu tượng hóa toán học và dữ liệu thực nghiệm thế giới thực")

    # 4 Research Questions Box
    q_boxes = [
        ("Câu hỏi 1: Mạng Xã Hội Động của Truyền Thông Bí Mật",
         "Can we study dynamic social networks of secret communication modeled on the empirical behavior of secret keepers?",
         "Nghiên cứu tổ chức thông tin dưới dạng mạng xã hội động giữa những người giữ bí mật. Kiểm tra xem các bất biến chính sách bảo vệ bí mật có được duy trì qua các biến động cấu trúc mạng và các sự cố Byzantine hay không."),
        ("Câu hỏi 2: Mô Hình Hóa Năng Lực và Liên Minh Đối Thủ Bằng Thống Kê",
         "By incorporating statistics, can new models of the adversary take into account more accurate data about their capabilities?",
         "Sử dụng thống kê thực nghiệm và mô phỏng Monte Carlo để mô hình hóa chính xác hơn năng lực đối thủ; phân tích kinh tế chính trị của các liên minh giám sát (liên minh Five Eyes, Anonymous, quan hệ nhà nước - cá nhân)."),
        ("Câu hỏi 3: Ứng Dụng Phương Pháp Liên Ngành & Hệ Thống Phức Hợp",
         "Can cryptography incorporate interdisciplinary methods and insights that can capture the complexity of secure communication?",
         "Áp dụng các công cụ từ lý thuyết hệ thống phức hợp, động lực học mạng và dịch tễ học để mô hình hóa sự lan truyền thông tin, nhận diện các điểm chuyển pha (phase shifts) khi bị tấn công."),
        ("Câu hỏi 4: Triết Học Giúp Tái Cấu Trúc Các Giả Định Ngầm Định",
         "Will philosophy let us make explicit and then reformulate the hidden assumptions of cryptographic communication, and vice versa?",
         "Vận dụng triết học để chất vấn các giả định nền tảng của Shannon; xem xét ranh giới động giữa con người và môi trường kỹ thuật số theo Giả thuyết Tâm trí Mở rộng (Extended Mind Hypothesis).")
    ]

    for idx, (q_title, q_en, q_desc) in enumerate(q_boxes):
        top_pos = Inches(1.85) + idx * Inches(1.18)
        add_card(s10, Inches(0.8), top_pos, Inches(11.733), Inches(1.05), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
        t = s10.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.06), Inches(11.333), Inches(0.92))
        tf = t.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = q_title
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_BRAND_BLUE

        p_en = tf.add_paragraph()
        p_en.text = f"\"{q_en}\""
        p_en.font.size = Pt(9.5)
        p_en.font.italic = True
        p_en.font.color.rgb = COLOR_TEXT_MUTED

        p_d = tf.add_paragraph()
        p_d.space_before = Pt(2)
        p_d.text = q_desc
        p_d.font.size = Pt(10)
        p_d.font.color.rgb = COLOR_TEXT_PRIMARY

    f_box = s10.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn nguyên văn 4 câu hỏi nghiên cứu mở: Harry Halpin (2025), Section 7, tr. 12–14."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 11: PHƯƠNG PHÁP LUẬN & GIỚI HẠN TÁI LẬP BÀI BÁO
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_bg(s11)
    add_header(s11, "PHƯƠNG PHÁP LUẬN & KHẢ NĂNG TÁI LẬP BÀI BÁO", "ĐÁNH GIÁ PHƯƠNG PHÁP LUẬN & HƯỚNG TIẾP CẬN THỰC NGHIỆM",
               "Làm rõ tính khả thi tái lập đối với một công trình triết học và định hướng thí nghiệm độc lập")

    add_card(s11, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_AMBER_BG, COLOR_AMBER_BORDER)
    t = s11.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "1. Đánh Giá Khả Năng Tái Lập (Replication Feasibility)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_AMBER

    replicate_pts = [
        ("Bản chất công trình:", "Bài báo của TS. Harry Halpin là một luận văn tổng quan triết học và phân tích nhận thức luận, không phải là một công trình kỹ thuật phần mềm."),
        ("Không có tài nguyên phần mềm:", "Bài báo HOÀN TOÀN KHÔNG có mã nguồn triển khai, không có tập dữ liệu thực nghiệm, không có các bảng đo thời gian (ms) hay dung lượng benchmark."),
        ("Phương pháp thẩm định:", "Nhóm sinh viên thực hiện thẩm định học thuật (Scholarly Audit) đối chiếu toàn bộ 17 trang bản in PDF, 7 mục nội dung và 134 tài liệu tham khảo."),
        ("Tính trung thực học thuật:", "Không bịa đặt số liệu gốc của tác giả; không tuyên bố 'tái lập 100% benchmark bài báo' khi tác giả không công bố benchmark.")
    ]
    for h, b in replicate_pts:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(8)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    add_card(s11, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_EMERALD_BG, COLOR_EMERALD_BORDER)
    t = s11.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "2. Định Hướng Thí Nghiệm Minh Họa Độc Lập"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_EMERALD

    independent_pts = [
        ("Mục tiêu thực nghiệm:", "Sinh viên tự xây dựng hệ thống kiểm tra toàn vẹn tệp nhằm chuyển hóa các luận điểm lý thuyết trừu tượng của Halpin thành các kịch bản quan sát được."),
        ("Minh họa trực quan:", "Thiết kế các thí nghiệm minh họa: sự bất lực của mã tuyến tính CRC32 trước kẻ thù toán học; can thiệp đường truyền Dolev-Yao; rò rỉ siêu dữ liệu kênh truyền."),
        ("Phân định ranh giới rõ ràng:", "Mã nguồn và hệ thống kiểm tra toàn vẹn là đóng góp kỹ thuật độc lập của sinh viên; không được nhầm lẫn là thuật toán do Halpin đề xuất."),
        ("Báo cáo khoa học nghiêm túc:", "Tất cả kết quả đo đạc được lưu vết kiểm toán tại file MIDTERM_CODE_AUDIT.md.")
    ]
    for h, b in independent_pts:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(8)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s11.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Báo cáo phân định phương pháp luận của nhóm nghiên cứu đề tài DT005."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 12: MINH HỌA PAPER-TO-CODE ĐỘC LẬP (SƠ ĐỒ 4 & BẢNG ĐỐI SÁNH)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_bg(s12)
    add_header(s12, "MINH HỌA PAPER-TO-CODE ĐỘC LẬP & GIỚI HẠN DIỄN GIẢI", "SƠ ĐỒ 4: PHÂN ĐỊNH LÝ THUYẾT & THÍ NGHIỆM ĐỘC LẬP",
               "Đối sánh minh bạch giữa các luận điểm bài báo và thí nghiệm minh họa do sinh viên xây dựng")

    # SƠ ĐỒ 4: Two independent streams
    # Top Stream Header
    add_card(s12, Inches(0.8), Inches(1.85), Inches(5.7), Inches(0.85), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
    t = s12.shapes.add_textbox(Inches(0.9), Inches(1.9), Inches(5.5), Inches(0.75))
    tf = t.text_frame
    p = tf.paragraphs[0]
    p.text = "LUỒNG LÝ THUYẾT BÀI BÁO (HARRY HALPIN 2025)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE
    p_sub = tf.add_paragraph()
    p_sub.text = "Khảo sát triết học, phê phán trừu tượng hóa, phân tích Dolev-Yao & NSA"
    p_sub.font.size = Pt(9.5)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom Stream Header
    add_card(s12, Inches(6.8), Inches(1.85), Inches(5.7), Inches(0.85), COLOR_EMERALD_BG, COLOR_EMERALD_BORDER)
    t = s12.shapes.add_textbox(Inches(6.9), Inches(1.9), Inches(5.5), Inches(0.75))
    tf = t.text_frame
    p = tf.paragraphs[0]
    p.text = "LUỒNG THỰC NGHIỆM ĐỘC LẬP (SINH VIÊN DT005)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_EMERALD
    p_sub = tf.add_paragraph()
    p_sub.text = "Hệ thống kiểm tra toàn vẹn tệp, 7 thí nghiệm minh họa, đối kháng Red/Blue"
    p_sub.font.size = Pt(9.5)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    # Bảng 3 ví dụ tiêu biểu đã kiểm tra
    add_card(s12, Inches(0.8), Inches(2.85), Inches(11.733), Inches(3.85), COLOR_BG_PANEL, COLOR_BORDER)
    t = s12.shapes.add_textbox(Inches(1.0), Inches(2.95), Inches(11.333), Inches(3.65))
    tf = t.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "BẢNG ĐỐI CHIẾU BA MINH HỌA THỰC NGHIỆM TIÊU BIỂU VÀ GIỚI HẠN DIỄN GIẢI"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_TITLE

    table_data = [
        ("1. Noise vs Adversary (Sec. 1 & 2)",
         "Kẻ thù có chủ đích khác với nhiễu ngẫu nhiên; mã tuyến tính CRC32 bất lực.",
         "Code giải mã bù đại số trên trường nhị phân GF(2), tìm 4 byte bù qua mặt CRC32 100%; SHA-256 phát hiện tức thì.",
         "Bằng chứng đã chạy:", "orig_crc == forged_crc (Trùng 100%), SHA-256 lệch hoàn toàn.",
         "Giới hạn diễn giải:", "Minh họa tính phi tuyến của hàm băm; không chứng minh an toàn toàn diện cho toàn hệ thống."),
        ("2. Node-Leaf Confusion (Sec. 5)",
         "Ghép chuỗi thô H(L||R) gây va chạm tiền ảnh cấp hai cấu trúc (nút cha bị nhầm với nút lá).",
         "Triển khai phân tách miền chuẩn RFC 6962: Tiền tố 0x00 cho lá, 0x01 cho nhánh nội bộ.",
         "Bằng chứng đã chạy:", "Standard Merkle Tree ghép thô bị va chạm; HD-Merkle phân tách miền chặn đứng hoàn toàn.",
         "Giới hạn diễn giải:", "Chỉ giải quyết va chạm cấu trúc tầng cây; không loại trừ va chạm ngẫu nhiên 2^256 của SHA-256."),
        ("3. Timing Side-Channel (Sec. 5)",
         "Rò rỉ thông tin qua thời gian so sánh chuỗi (Early-return comparison leak - Kocher 1996).",
         "Khảo sát đo lường 5 vòng so sánh chuỗi naive vs hàm thời gian hằng số hmac.compare_digest.",
         "Bằng chứng đã chạy:", "Naive so sánh tuần tự lộ độ lệch trung vị ~0.08 µs; hmac.compare_digest triệt tiêu độ lệch.",
         "Giới hạn diễn giải:", "Phép đo trên Python máy đơn có jitter; chỉ minh họa mức thư viện C, không chứng minh miễn nhiễm kênh kề phần cứng.")
    ]

    for item in table_data:
        p_row = tf.add_paragraph()
        p_row.space_before = Pt(6)
        r_t = p_row.add_run()
        r_t.text = f"{item[0]} — "
        r_t.font.bold = True
        r_t.font.size = Pt(10.5)
        r_t.font.color.rgb = COLOR_BRAND_BLUE

        r_claim = p_row.add_run()
        r_claim.text = f"Luận điểm Halpin: {item[1]} | "
        r_claim.font.size = Pt(10)
        r_claim.font.color.rgb = COLOR_TEXT_PRIMARY

        r_code = p_row.add_run()
        r_code.text = f"Code DT005: {item[2]} "
        r_code.font.size = Pt(10)
        r_code.font.color.rgb = COLOR_TEXT_SECONDARY

        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(2)
        r_ev = p_sub.add_run()
        r_ev.text = f"   ► {item[3]} {item[4]}  |  "
        r_ev.font.bold = True
        r_ev.font.size = Pt(9.5)
        r_ev.font.color.rgb = COLOR_ACCENT_EMERALD

        r_lim = p_sub.add_run()
        r_lim.text = f"▲ {item[5]} {item[6]}"
        r_lim.font.size = Pt(9.5)
        r_lim.font.color.rgb = COLOR_ACCENT_AMBER

    f_box = s12.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ và bảng đối chiếu do sinh viên tự xây dựng; chi tiết mã nguồn và kiểm toán lưu tại MIDTERM_CODE_AUDIT.md."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 13: DEMO THỰC NGHIỆM ĐỐI KHÁNG (RED TEAM VS BLUE TEAM)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_bg(s13)
    add_header(s13, "DEMO THỰC NGHIỆM ĐỐI KHÁNG: RED TEAM VS BLUE TEAM", "THÍ NGHIỆM MINH HỌA ĐỘC LẬP DO SINH VIÊN THIẾT KẾ",
               "Diễn tập an ninh mạng trực quan minh họa các luận điểm đối thủ trên tập dữ liệu thực tế")

    # Left: Architecture & Scenarios
    add_card(s13, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
    t = s13.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "1. Kiến Trúc Diễn Tập Đối Kháng Trực Quan"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE

    p_arch = tf.add_paragraph()
    p_arch.space_before = Pt(6)
    p_arch.text = "[BLUE TEAM - Cổng 8080]        [RED TEAM - Cổng 8081]\n- Giám sát kaggle_dataset (2.483 tệp)  - 6 Vector tấn công đối kháng\n- Quét mốc tin cậy ban đầu H_0         - Sửa tệp, CRC32, Dolev-Yao\n- Phát hiện sai lệch & Dirty Path     - Phục hồi an toàn (Rollback)"
    p_arch.font.size = Pt(9.5)
    p_arch.font.name = "Consolas"
    p_arch.font.color.rgb = COLOR_TEXT_PRIMARY

    demo_scenarios = [
        ("Kịch bản 1 (CRC32 Forge):", "Red Team sửa giao dịch, bù 4 byte GF(2) qua mặt CRC32; Blue Team dùng SHA-256 bắt quả tang tức thì."),
        ("Kịch bản 2 (In-place Tamper):", "Sửa nội dung tệp; Blue Team phát hiện FILE_MODIFIED và định vị chính xác qua Dirty Path O(log N)."),
        ("Kịch bản 3 (Dolev-Yao MitM):", "Đánh tráo Merkle Proof trên đường truyền mạng; bộ xác thực từ chối 100% chữ ký bằng chứng."),
        ("Khôi phục an toàn (Rollback):", "Nút 'ROLLBACK ALL' phục hồi 100% nguyên trạng từ thư mục an toàn .adversary_backups/.")
    ]
    for h, b in demo_scenarios:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(5)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right: Real Execution Log & Limitations
    add_card(s13, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_EMERALD_BG, COLOR_EMERALD_BORDER)
    t = s13.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "2. Bằng Chứng Chạy Thật & Giới Hạn Thực Nghiệm"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_EMERALD

    p_log = tf.add_paragraph()
    p_log.space_before = Pt(6)
    p_log.text = "[LOG THỰC TẾ: python -m src.proposed_model.security_games]\n[*] Thư mục kiểm định : kaggle_dataset (2,483 tệp tin)\n[*] PPT Integrity Game: Phát hiện 100/100 phép thử (Tỷ lệ = 0.0)\n[*] Dolev-Yao Network : Chặn đứng 3/3 kịch bản (Swap, Tamper, Replay)\n[*] Consistency Proof : Xác thực nhất quán chỉ-thêm RFC 6962\n[*] Tuyết lở SAC Monte Carlo: 127.28/256 bit biến đổi (~49.72%)\n[*] Thời gian tổng cộng: 2,885.53 ms (Đạt chuẩn 7/7 thí nghiệm)"
    p_log.font.size = Pt(9)
    p_log.font.name = "Consolas"
    p_log.font.color.rgb = COLOR_BRAND_NAVY

    limitations = [
        ("Bản chất minh họa:", "Thực nghiệm được thiết kế độc lập nhằm kiểm chứng tính đúng đắn logic của các khái niệm; không phải là việc chạy lại số liệu của tác giả."),
        ("Môi trường kiểm thử:", "Đo đạc trên máy tính cá nhân (Intel Core i7, Windows 11). Các phép đo thời gian vi mô chịu ảnh hưởng của jitter hệ điều hành."),
        ("Hướng dẫn chạy:", "Chi tiết lệnh chạy và các bước thao tác màn hình kép được hướng dẫn trong file MIDTERM_DEMO.md.")
    ]
    for h, b in limitations:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(5)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s13.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Số liệu thực tế đo đạc trực tiếp từ mã nguồn đề tài DT005; xem hướng dẫn chạy tại MIDTERM_DEMO.md."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 14: KẾT LUẬN & TÀI LIỆU THAM KHẢO
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_bg(s14)
    add_header(s14, "TỔNG KẾT BÁO CÁO GIỮA KỲ & TÀI LIỆU THAM KHẢO", "SECTION 7 — KẾT LUẬN & NGUỒN THAM CHIẾU HỌC THUẬT",
               "Tổng kết các luận điểm chính của Halpin, bài học thiết kế an ninh và danh mục tài liệu")

    # Left: Conclusion
    add_card(s14, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
    t = s14.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "1. Tổng Kết Thông Điệp Học Thuật Của Bài Báo"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE

    conclusions = [
        ("Bài học từ Halpin:", "Mật mã học không chỉ là toán học thuần túy trên giấy tờ. Các hệ thống an toàn trong thế giới thực phải nhận thức rõ khoảng cách giữa mô hình trừu tượng và thực tế triển khai (kênh kề, siêu dữ liệu, chuỗi cung ứng, thể chế)."),
        ("Đóng góp của Halpin:", "Đặt nền móng triết học cho mật mã học, phân tích nhận thức luận về kẻ thù và vạch ra 4 hướng nghiên cứu mở nhằm mở rộng mật mã thành một khoa học an ninh hoàn chỉnh."),
        ("Hoàn thành mục tiêu giữa kỳ:", "Nhóm sinh viên DT005 đã thẩm định học thuật sâu sắc toàn văn bài báo gốc (17 trang PDF) và xây dựng thành công bộ thí nghiệm minh họa thực nghiệm sống động phục vụ báo cáo hội đồng.")
    ]
    for h, b in conclusions:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(8)
        r1 = p_it.add_run()
        r1.text = f"• {h} "
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_TEXT_PRIMARY
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right: References
    add_card(s14, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_BG_PANEL, COLOR_BORDER)
    t = s14.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "2. Tài Liệu Tham Khảo Chọn Lọc (Key References)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_TITLE

    refs = [
        ("[1] H. Halpin,", "The adversary: the philosophy of cryptography, Journal of Cybersecurity, Oxford Academic, Vol. 11, No. 1, tyaf006, 2025."),
        ("[2] C. E. Shannon,", "A mathematical theory of communication, Bell Syst. Tech. J., 1948."),
        ("[3] C. E. Shannon,", "Communication theory of secrecy systems, Bell Syst. Tech. J., 1949."),
        ("[4] S. Goldwasser & S. Micali,", "Probabilistic encryption, JCSS, 1984."),
        ("[5] D. Dolev & A. Yao,", "On the security of public key protocols, IEEE TIT, 1983."),
        ("[6] M. Bellare & P. Rogaway,", "Random oracles are practical, ACM CCS, 1993."),
        ("[7] P. Kocher,", "Timing attacks on implementations of Diffie-Hellman, RSA, DSS, Crypto, 1996."),
        ("[8] B. Laurie et al.,", "Certificate Transparency, IETF RFC 6962, 2013.")
    ]
    for h, b in refs:
        p_it = tf.add_paragraph()
        p_it.space_before = Pt(4)
        r1 = p_it.add_run()
        r1.text = f"{h} "
        r1.font.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_BRAND_BLUE
        r2 = p_it.add_run()
        r2.text = b
        r2.font.size = Pt(9)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s14.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Đề tài DT005 — An Toàn và Bảo Mật Hệ Thống Thông Tin — Báo cáo Tiến độ Giữa kỳ."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # Lưu file
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Bao_Cao_Giua_Ky_DT05_Final.pptx")
    prs.save(out_path)
    
    # Lưu thêm 1 bản tại thư mục gốc để dễ truy cập
    root_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bao_Cao_Giua_Ky_DT05_Final.pptx")
    prs.save(root_path)

    print(f"[THÀNH CÔNG] Đã tạo bài thuyết trình giữa kỳ hoàn chỉnh tại:\n  1. {out_path}\n  2. {root_path}")
    print(f"Tổng số slide: {len(prs.slides)} slide chuẩn 16:9 widescreen.")


if __name__ == "__main__":
    create_deck()
