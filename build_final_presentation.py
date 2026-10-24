# -*- coding: utf-8 -*-
"""
Script Xây Dựng Slide Thuyết Trình Báo Cáo Giữa Kỳ Chuẩn Khoa Học:
Đề tài DT005 - Thẩm định Học thuật và Minh họa Thực nghiệm Bài báo Harry Halpin (2025):
"The adversary: the philosophy of cryptography" (Journal of Cybersecurity, Oxford Academic).

Quy chuẩn thiết kế:
- Khổ 16:9 widescreen (13.333 x 7.5 inch).
- Light Academic Theme (nền sáng, card slate/blue/emerald/amber, chữ sắc nét).
- Đúng 13 slide tinh gọn, hoàn chỉnh:
  + Slide 1: Bìa chuẩn nhận diện (Thanh dọc xanh, Trường, Khoa, Báo cáo đồ án, Tên đề tài, GVHD, SV Võ Anh Kiệt).
  + Slide 2-9: Nội dung học thuật bài báo Halpin (Bối cảnh, Dretske, Shannon, Dolev-Yao, Tiến hóa đối thủ, Giới hạn hình thức hóa, Snowden, 4 Câu hỏi mở).
  + Slide 10: Kiến trúc Diễn tập Đối kháng Red Team vs Blue Team (Sơ đồ trực quan + Ghi chú minh họa độc lập).
  + Slide 11: Slide chuyển tiếp tối giản: DEMO.
  + Slide 12: Bảng Tổng Kết Đóng Góp Học Thuật & Bài Học An Ninh Thực Tế (Bảng đối sánh tổng hợp).
  + Slide 13: Slide Cảm Ơn & Q&A (Đồng bộ nhận diện với Slide 1).
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

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.82), Inches(11.7), Inches(0.55))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(21)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_TITLE

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
        return card

    def set_shape_text(shape, lines, font_size=Pt(9.5), bold_first=True, color=COLOR_TEXT_PRIMARY):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.04)
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = line
            p.font.size = font_size
            p.font.bold = bold_first if i == 0 else False
            p.font.color.rgb = color
            p.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 1: BÌA BÁO CÁO GIỮA KỲ (THIẾT KẾ ĐỒNG BỘ 100% THEO ẢNH MẪU CỦA USER)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    # Dải màu xanh dương dọc bên lề trái (nhận diện thương hiệu như ảnh đính kèm)
    left_bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.8), Inches(0.18), Inches(5.9))
    left_bar1.fill.solid()
    left_bar1.fill.fore_color.rgb = COLOR_ACCENT_BLUE
    left_bar1.line.fill.background()

    # Khối tiêu đề Trường & Khoa ở phần trên
    t_inst = s1.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11.0), Inches(0.9))
    tf_inst = t_inst.text_frame
    tf_inst.word_wrap = True
    p_uni = tf_inst.paragraphs[0]
    p_uni.alignment = PP_ALIGN.CENTER
    p_uni.text = "Trường Đại học Tài nguyên và Môi trường TP. Hồ Chí Minh"
    p_uni.font.size = Pt(15)
    p_uni.font.bold = True
    p_uni.font.color.rgb = COLOR_TEXT_PRIMARY

    p_fac = tf_inst.add_paragraph()
    p_fac.alignment = PP_ALIGN.CENTER
    p_fac.space_before = Pt(4)
    p_fac.text = "Khoa Công nghệ thông tin"
    p_fac.font.size = Pt(14)
    p_fac.font.bold = True
    p_fac.font.color.rgb = COLOR_TEXT_SECONDARY

    # Tiêu đề báo cáo lớn màu xanh dương ở giữa
    t_main = s1.shapes.add_textbox(Inches(1.2), Inches(2.5), Inches(11.0), Inches(2.2))
    tf_main = t_main.text_frame
    tf_main.word_wrap = True

    p_report = tf_main.paragraphs[0]
    p_report.alignment = PP_ALIGN.CENTER
    p_report.text = "Báo cáo thuyết trình đồ án giữa kỳ"
    p_report.font.size = Pt(28)
    p_report.font.bold = True
    p_report.font.color.rgb = COLOR_ACCENT_BLUE

    p_topic = tf_main.add_paragraph()
    p_topic.alignment = PP_ALIGN.CENTER
    p_topic.space_before = Pt(12)
    p_topic.text = "Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động"
    p_topic.font.size = Pt(16)
    p_topic.font.bold = True
    p_topic.font.color.rgb = COLOR_TEXT_TITLE

    p_sub = tf_main.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.space_before = Pt(8)
    p_sub.text = "Khảo sát và Thẩm định học thuật bài báo: \"The adversary: the philosophy of cryptography\"\nTS. Harry Halpin — Oxford Academic, Journal of Cybersecurity, Vol. 11, No. 1, 2025 (DOI: 10.1093/cybsec/tyaf006)"
    p_sub.font.size = Pt(11)
    p_sub.font.italic = True
    p_sub.font.color.rgb = COLOR_BRAND_BLUE

    # Khối thông tin học phần, giảng viên và sinh viên ở góc dưới bên trái
    t_info = s1.shapes.add_textbox(Inches(1.4), Inches(5.1), Inches(10.0), Inches(1.6))
    tf_info = t_info.text_frame
    tf_info.word_wrap = True

    p_c = tf_info.paragraphs[0]
    p_c.text = "Học phần: An toàn và bảo mật hệ thống thông tin"
    p_c.font.size = Pt(13)
    p_c.font.color.rgb = COLOR_TEXT_PRIMARY

    p_lec = tf_info.add_paragraph()
    p_lec.space_before = Pt(4)
    p_lec.text = "Giảng viên hướng dẫn: ThS. Phạm Trọng Huynh"
    p_lec.font.size = Pt(13)
    p_lec.font.color.rgb = COLOR_TEXT_PRIMARY

    p_stu = tf_info.add_paragraph()
    p_stu.space_before = Pt(4)
    r_s1 = p_stu.add_run()
    r_s1.text = "Sinh viên thực hiện: "
    r_s1.font.size = Pt(13)
    r_s1.font.color.rgb = COLOR_TEXT_PRIMARY
    r_s2 = p_stu.add_run()
    r_s2.text = "Võ Anh Kiệt"
    r_s2.font.bold = True
    r_s2.font.size = Pt(13)
    r_s2.font.color.rgb = COLOR_BRAND_BLUE
    r_s3 = p_stu.add_run()
    r_s3.text = " — MSSV: 1150080143 — Lớp: 11_CNPM2"
    r_s3.font.size = Pt(13)
    r_s3.font.color.rgb = COLOR_TEXT_PRIMARY

    # =========================================================================
    # SLIDE 2: BỐI CẢNH, MỤC TIÊU & ĐÓNG GÓP BÀI BÁO HALPIN (SEC 1, TR. 1–2)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    add_header(s2, "BỐI CẢNH, MỤC TIÊU & BA ĐÓNG GÓP CỦA BÀI BÁO", "SECTION 1 — GIỚI THIỆU & MỤC TIÊU BÀI BÁO (TRANG 1–2)",
               "Khảo cứu nhận thức luận về Kẻ thù trong Mật mã học và Chương trình nghiên cứu tương lai")

    add_card(s2, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_BG_PANEL, COLOR_BORDER)
    t = s2.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "1. Bối Cảnh Mật Mã Học & Nghịch Lý Nhận Thức Luận"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_BRAND_BLUE

    points_s2_left = [
        ("Mật mã học trong kỷ nguyên số:", "Khoa học về truyền thông bí mật khi có sự hiện diện của đối thủ có chủ đích. Internet biến không gian mạng thành nơi tương tác với vô vàn đối thủ tiềm tàng, xa lạ."),
        ("Nghịch lý nhận thức luận:", "Lý thuyết thông tin của Shannon (1948) được triết học khảo cứu sâu sắc, nhưng chính mật mã học — cái nôi khai sinh lý thuyết thông tin — lại bị triết học bỏ quên suốt nhiều thập kỷ."),
        ("Cái giá của sự trừu tượng hóa:", "Khi mật mã học biến thành các bài toán toán học thuần túy trên giấy, nó tạo ra điểm mù nguy hiểm trước các đòn tấn công vật lý trong thế giới thực.")
    ]
    for h, b in points_s2_left:
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

    add_card(s2, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8), COLOR_EMERALD_BG, COLOR_EMERALD_BORDER)
    t = s2.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(4.5))
    tf = t.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "2. Ba Đóng Góp Học Thuật Chính Của Harry Halpin"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_EMERALD

    points_s2_right = [
        ("Khảo sát Lịch sử & Triết học:", "Hệ thống hóa toàn diện quá trình tiến hóa mô hình đối thủ: từ Shannon (1949), PPT Adversary (1984), Dolev-Yao (1983) đến Random Oracle Model (1993)."),
        ("Phê phán Giới hạn Hình thức hóa:", "Chỉ rõ sự sụp đổ của các mô hình toán học trước thực tiễn: rò rỉ kênh kề (Kocher 1996), độ dài bản mã, rò rỉ siêu dữ liệu và giám sát hàng loạt của NSA (Snowden 2013)."),
        ("Đề xuất Chương trình Mới:", "Thiết lập 4 câu hỏi nghiên cứu mở (Section 7) nhằm mở rộng mật mã học từ một nhánh toán học biệt lập thành một khoa học an ninh hoàn chỉnh kết hợp thực nghiệm.")
    ]
    for h, b in points_s2_right:
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

    f_box = s2.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn: Harry Halpin (2025), Section 1: Introduction, Journal of Cybersecurity, tr. 1–2."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 3: CƠ SỞ TRIẾT HỌC: DRETSKE, SMITH & TÁC NHÂN (SEC 2 & 3, TR. 2–5)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    add_header(s3, "CƠ SỞ TRIẾT HỌC: THÔNG TIN, BÍ MẬT VÀ TÁC NHÂN", "SECTION 2 & 3 — CƠ SỞ TRIẾT HỌC VỀ THÔNG TIN (TRANG 2–5)",
               "Khung nhận thức luận của Fred Dretske và Brian Cantwell Smith ứng dụng trong Mật mã học")

    col_w = Inches(3.7)
    gap = Inches(0.3)
    cards_data = [
        ("1. Bản Thể Học Thông Tin", "BRIAN CANTWELL SMITH (1996)",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         [("Dòng chảy nguyên thủy:", "Thế giới khởi đầu từ dòng chảy siêu hình không biên giới (primordial metaphysical flux)."),
          ("Phân ranh giới:", "Con người và hệ thống định hình thực tại bằng cách đăng ký đối tượng (registration of objects)."),
          ("Tác nhân xa (Distal):", "Thông tin mật mã ra đời để truyền đạt trạng thái giữa các thực thể ở xa nhau mà không thể tương tác trực tiếp.")]),
        ("2. Thuyết Thông Tin Khách Quan", "FRED DRETSKE (1981)",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER,
         [("Thông tin khách quan:", "Thông tin là tính quy luật của trạng thái sự vật (state of affairs), hoàn toàn độc lập với việc có ai đọc hay không."),
          ("Thông tin và Bí mật:", "Bản rõ chứa đựng thông tin về nguồn; bản mã khi được mã hóa hoàn hảo sẽ không làm rò rỉ bất kỳ thông tin nào về nguồn."),
          ("Dòng chảy thông tin:", "Nhiệm vụ của mật mã là kiểm soát dòng chảy thông tin chỉ đi đến người nhận được ủy quyền.")]),
        ("3. Tác Nhân Có Chủ Đích", "BẢN CHẤT CỦA ĐỐI THỦ",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD,
         [("Tính chủ đích (Intentionality):", "Đối thủ không phải là 'nhiễu ngẫu nhiên' mà là thực thể có niềm tin, mục tiêu và ý chí khai thác lỗ hổng."),
          ("Tính tương đối của bí mật:", "Bí mật không tồn tại tự thân; nó là mối quan hệ giữa nhóm người giữ bí mật và thế giới bên ngoài trong một khoảng thời gian."),
          ("Public vs Common Knowledge:", "Common Knowledge (Lewis): Mọi tác nhân đều biết và tin trạng thái sự việc. Public Knowledge: Thông tin mà bất kỳ ai cũng có tiềm năng tiếp cận (như khi bí mật bị rò rỉ - leak).")]),
    ]

    for idx, (c_title, c_badge, bg_c, brd_c, text_c, items) in enumerate(cards_data):
        left_pos = Inches(0.8) + idx * (col_w + gap)
        add_card(s3, left_pos, Inches(1.85), col_w, Inches(4.8), bg_c, brd_c)
        t = s3.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.0), col_w - Inches(0.4), Inches(4.5))
        tf = t.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = c_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_c

        sub_p = tf.add_paragraph()
        sub_p.text = c_badge
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

    f_box = s3.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn khoa học: Dretske (1981, Ref [1]); Cantwell Smith (1996, Ref [5]); Halpin (2025), tr. 2–5."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 4: TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ (SƠ ĐỒ 1: TIMELINE)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4)
    add_header(s4, "TIẾN TRÌNH LỊCH SỬ MÔ HÌNH HÓA ĐỐI THỦ TRONG MẬT MÃ", "SƠ ĐỒ 1: DÒNG THỜI GIAN TIẾN HÓA MÔ HÌNH ĐỐI THỦ",
               "Sự dịch chuyển từ Lý thuyết thông tin sang Giả định độ phức tạp tính toán và Phê phán hậu Snowden")

    timeline_phases = [
        ("Giai đoạn 1 (1948–1949)", "MÔ HÌNH SHANNON",
         "Mô hình truyền thông bảo mật điểm - điểm; định nghĩa Bí mật hoàn hảo H(K) >= H(M). Giới hạn: Giản lược đối thủ thành kẻ nghe lén trên kênh trừu tượng, bỏ qua chi phí phân phối khóa và thực tế xã hội - kỹ thuật.",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE),
        ("Giai đoạn 2 (1976)", "CÁCH MẠNG KHÓA CÔNG KHAI",
         "Diffie-Hellman và Merkle giải quyết bài toán phân phối khóa; chuyển gánh nặng bảo mật sang độ phức tạp tính toán một chiều.",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER),
        ("Giai đoạn 3 (1982–1984)", "PPT ADVERSARY & SECURITY GAMES",
         "Goldwasser-Micali định nghĩa an toàn ngữ nghĩa và đối thủ thời gian đa thức xác suất (PPT). Chứng minh an toàn bằng quy dẫn toán học.",
         COLOR_BG_PANEL, COLOR_BORDER, COLOR_TEXT_TITLE),
        ("Giai đoạn 4 (1981–1993)", "DOLEV-YAO & RANDOM ORACLE",
         "Dolev-Yao mô hình hóa đối thủ mạng kiểm soát hoàn toàn kênh truyền. Bellare-Rogaway đề xuất Random Oracle Model (ROM) ứng dụng thực hành.",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD),
        ("Giai đoạn 5 (2013–2025)", "PHÊ PHÁN HẬU SNOWDEN",
         "Tiết lộ của Snowden phơi bày năng lực thực tế của NSA: Kênh kề, cửa sau chuẩn mật mã, tấn công chuỗi cung ứng và giám sát siêu dữ liệu.",
         RGBColor(254, 242, 242), RGBColor(254, 202, 202), COLOR_ACCENT_ROSE)
    ]

    for idx, (t_time, t_badge, t_desc, bg_c, brd_c, text_c) in enumerate(timeline_phases):
        top_pos = Inches(1.85) + idx * Inches(0.95)
        add_card(s4, Inches(0.8), top_pos, Inches(11.733), Inches(0.85), bg_c, brd_c)
        t = s4.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.05), Inches(11.333), Inches(0.75))
        tf = t.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r_time = p.add_run()
        r_time.text = f"{t_time} — "
        r_time.font.bold = True
        r_time.font.size = Pt(11.5)
        r_time.font.color.rgb = text_c

        r_badge = p.add_run()
        r_badge.text = f"[{t_badge}]: "
        r_badge.font.bold = True
        r_badge.font.size = Pt(11)
        r_badge.font.color.rgb = COLOR_TEXT_PRIMARY

        r_desc = p.add_run()
        r_desc.text = t_desc
        r_desc.font.size = Pt(10.5)
        r_desc.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s4.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ tổng hợp dựa trên Section 3–6, Harry Halpin (2025), Journal of Cybersecurity, tr. 3–12."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: SHANNON & MẠNG PHỨC HỢP (SƠ ĐỒ 2: MÔ HÌNH TRUYỀN THÔNG)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5)
    add_header(s5, "MÔ HÌNH TRUYỀN THÔNG CỦA SHANNON VÀ MẠNG THỰC TẾ", "SƠ ĐỒ 2: ĐỐI SÁNH MÔ HÌNH TRUYỀN THÔNG ĐIỂM-ĐIỂM VÀ MẠNG XÃ HỘI-KỸ THUẬT",
               "Khoảng cách giữa mô hình điểm-điểm đóng kín của Shannon và mạng phức hợp không đồng nhất")

    img_s5_path = os.path.join(os.path.dirname(__file__), "docs", "Slide6.drawio.png")
    if os.path.exists(img_s5_path):
        img_w = Inches(10.3)
        img_h = Inches(4.84)
        img_l = (prs.slide_width - img_w) / 2
        img_t = Inches(1.82)
        s5.shapes.add_picture(img_s5_path, img_l, img_t, img_w, img_h)

    f_box = s5.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ tổng hợp dựa trên Shannon (1949) và Harry Halpin (2025), Journal of Cybersecurity, tr. 3–5."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 6: CÁC MÔ HÌNH ĐỐI THỦ HÌNH THỨC (SEC 4, TR. 5–8)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6)
    add_header(s6, "BA MÔ HÌNH ĐỐI THỦ HÌNH THỨC TRONG MẬT MÃ HIỆN ĐẠI", "SECTION 4 — HÌNH THỨC HÓA ĐỐI THỦ (TRANG 5–8)",
               "Tiến trình toán học hóa kẻ thù: Độ phức tạp tính toán đa thức, Logic kênh truyền và Tiên tri ngẫu nhiên")

    col_w7 = Inches(3.7)
    gap7 = Inches(0.3)
    formal_models = [
        ("1. Mô Hình PPT Adversary", "GOLDWASSER-MICALI (1984)",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         [("Khái niệm:", "Đối thủ là thuật toán Turing xác suất chạy trong thời gian đa thức (Probabilistic Polynomial-Time)."),
          ("Security Games:", "Định nghĩa an toàn thông qua các trò chơi thách đấu. Lợi thế Adv(A) phải là hàm không đáng kể (negligible)."),
          ("Quy dẫn toán học:", "Chứng minh an toàn bằng cách quy dẫn: bẻ gãy hệ mật tương đương giải bài toán khó (phân tích thừa số nguyên, logarit rời rạc).")]),
        ("2. Mô Hình Đối Thủ Mạng Logic", "DOLEV-YAO (1983)",
         COLOR_BG_PANEL, COLOR_BORDER, COLOR_TEXT_TITLE,
         [("Khái niệm:", "Đối thủ kiểm soát toàn bộ kênh truyền mạng: có quyền đọc, chặn, sửa đổi, xóa, tráo đổi và phát lại gói tin."),
          ("Hộp đen mật mã:", "Mô hình Dolev-Yao xem mật mã như các hàm đại số hoàn hảo (Perfect Black-box). Kẻ thù không thể giải mã nếu không có khóa."),
          ("Ứng dụng:", "Kiểm chứng hình thức các giao thức mạng phức tạp thông qua các công cụ logic tự động.")]),
        ("3. Mô Hình Tiên Tri Ngẫu Nhiên", "BELLARE-ROGAWAY (1993)",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER,
         [("Khái niệm:", "Giả định hàm băm thực tế hoạt động như một hàm ngẫu nhiên lý tưởng (Random Oracle Model - ROM) trả về chuỗi ngẫu nhiên đồng đều."),
          ("Vai trò thực tế:", "Cho phép xây dựng các hệ mật khóa công khai và chữ ký số hiệu năng cao có chứng minh an toàn."),
          ("Giới hạn triết học:", "Trong thế giới thực, không có hàm băm cụ thể nào (kể cả SHA-256) có thể tương đương toán học với một Random Oracle thực sự.")])
    ]

    for idx, (m_title, m_badge, bg_c, brd_c, text_c, items) in enumerate(formal_models):
        left_pos = Inches(0.8) + idx * (col_w7 + gap7)
        add_card(s6, left_pos, Inches(1.85), col_w7, Inches(4.8), bg_c, brd_c)
        t = s6.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.0), col_w7 - Inches(0.4), Inches(4.5))
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

    f_box = s6.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn: Goldwasser-Micali (1984); Dolev-Yao (1983); Bellare-Rogaway (1993); Halpin (2025), tr. 5–8."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 7: GIỚI HẠN CỦA HÌNH THỨC HÓA (SƠ ĐỒ 3: KHOẢNG CÁCH VỚI THỰC TẾ)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7)
    add_header(s7, "GIỚI HẠN CỦA SỰ HÌNH THỨC HÓA TRONG MẬT MÃ HỌC", "SECTION 5 — GIỚI HẠN TRỪU TƯỢNG HÓA (TRANG 8–10)",
               "Khoảng cách giữa các giả định toán học lý tưởng và thế giới vật lý thực tế")

    flow_steps = [
        ("BƯỚC 1: GIẢ ĐỊNH TOÁN HỌC HÌNH THỨC",
         "Hộp đen mật mã lý tưởng, thời gian thực thi là vô hình, độ dài bản mã công khai không rò rỉ nội dung, kênh truyền điểm-điểm độc lập.",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE),
        ("BƯỚC 2: CÁC YẾU TỐ VẬT LÝ BỊ TRỪU TƯỢNG HÓA LƯỢC BỎ",
         "Thời gian vi xử lý CPU, dao động điện năng tiêu thụ, phát xạ điện từ, kiến trúc bộ nhớ đệm (Cache), lỗi lập trình phần mềm.",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER),
        ("BƯỚC 3: CON ĐƯỜNG KHAI THÁC CỦA ĐỐI THỦ THẾ GIỚI THỰC",
         "Tấn công kênh kề (Timing Attack - Kocher 1996), phân tích độ dài và lưu lượng siêu dữ liệu (Traffic Analysis), khai thác lỗi cài đặt phần mềm và quản lý (Anderson 2001).",
         RGBColor(254, 242, 242), RGBColor(254, 202, 202), COLOR_ACCENT_ROSE),
        ("BƯỚC 4: KẾT LUẬN NHẬN THỨC LUẬN CỦA HARRY HALPIN",
         "Cái giá của sự trừu tượng hóa là làm cho mật mã học mù quáng trước chính những năng lực thực tế của đối thủ nằm ngoài mô hình lý thuyết.",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD)
    ]

    for idx, (step_t, step_d, bg_c, brd_c, text_c) in enumerate(flow_steps):
        top_pos = Inches(1.85) + idx * Inches(1.18)
        add_card(s7, Inches(0.8), top_pos, Inches(11.733), Inches(1.05), bg_c, brd_c)
        t = s7.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.08), Inches(11.333), Inches(0.9))
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

    f_box = s7.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ tổng hợp dựa trên Kocher (1996) và Halpin (2025), Section 5, tr. 8–10."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 8: TIẾT LỘ CỦA SNOWDEN & ĐỐI THỦ THỂ CHẾ (SEC 6, TR. 10–12)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8)
    add_header(s8, "SNOWDEN VÀ ĐỐI THỦ THỂ CHẾ (THE REAL ADVERSARY)", "SECTION 6 — TIẾT LỘ CỦA SNOWDEN (TRANG 10–12)",
               "Khủng hoảng mô hình an ninh trước đối thủ giám sát hàng loạt quy mô quốc gia")

    col_w9 = Inches(3.7)
    gap9 = Inches(0.3)
    snowden_pillars = [
        ("1. Khủng Hoảng Thực Tiễn", "TIẾT LỘ SNOWDEN (2013)",
         COLOR_BG_PANEL, COLOR_BORDER_BLUE, COLOR_BRAND_BLUE,
         [("Sự kiện bước ngoặt:", "Các tài liệu Edward Snowden tiết lộ về NSA đã kéo mật mã học vào một cuộc khủng hoảng thực tiễn sâu sắc."),
          ("Không phá toán học:", "NSA không cần giải các bài toán toán học khó của RSA hay AES mà tấn công trực tiếp vào toàn bộ hệ sinh thái xung quanh."),
          ("Mô hình thực tế:", "Đối thủ bị động toàn cầu (Global Passive Adversary) và tấn công can thiệp vào đường truyền là mối đe dọa hoàn toàn có thật.")]),
        ("2. Năng Lực Đối Thủ Thể Chế", "VƯỢT NGOÀI HÌNH THỨC",
         COLOR_AMBER_BG, COLOR_AMBER_BORDER, COLOR_ACCENT_AMBER,
         [("Giám sát Siêu dữ liệu:", "NSA tập trung khai thác Metadata (ai liên lạc với ai, thời gian, kích thước) để dựng lại mạng xã hội mà không cần bẻ khóa bản mã."),
          ("Tấn công Chuỗi cung ứng:", "Chèn cửa sau vào phần cứng, làm suy yếu chuẩn mật mã quốc tế (tiêu chuẩn NIST Dual_EC_DRBG bị cài cửa sau ngầm)."),
          ("Áp lực Thể chế:", "Cưỡng chế pháp lý các tập đoàn công nghệ lớn bàn giao khóa bí mật và quyền truy cập dữ liệu máy chủ.")]),
        ("3. Tính Tương Đối Của Kẻ Thù", "BẢN CHẤT NHẬN THỨC LUẬN",
         COLOR_EMERALD_BG, COLOR_EMERALD_BORDER, COLOR_ACCENT_EMERALD,
         [("Khái niệm tương đối:", "Khái niệm 'Kẻ thù' mang tính vị thế: Đối với người dân, NSA là đối thủ giám sát; nhưng đối với NSA, Snowden lại là đối thủ làm rò rỉ bí mật."),
          ("Dân chủ hóa mật mã:", "Mật mã khóa công khai từng giúp cá nhân tự vệ, nhưng công cụ gián điệp của NSA nay đã lan rộng sang tội phạm mạng và các nhóm phi nhà nước."),
          ("Hệ quả khoa học:", "Mật mã học buộc phải từ bỏ giả định kênh truyền sạch để nghiên cứu phòng chống cửa sau (Cliptography).")])
    ]

    for idx, (s_title, s_badge, bg_c, brd_c, text_c, items) in enumerate(snowden_pillars):
        left_pos = Inches(0.8) + idx * (col_w9 + gap9)
        add_card(s8, left_pos, Inches(1.85), col_w9, Inches(4.8), bg_c, brd_c)
        t = s8.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.0), col_w9 - Inches(0.4), Inches(4.5))
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

    f_box = s8.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn: Harry Halpin (2025), Section 6: Snowden and the real adversary, tr. 10–12."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 9: BỐN CÂU HỎI NGHIÊN CỨU MỞ CỦA HALPIN (SEC 7, TR. 12–14)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9)
    add_header(s9, "CHƯƠNG TRÌNH NGHIÊN CỨU TƯƠNG LAI: BỐN CÂU HỎI MỞ", "SECTION 7 — KẾT LUẬN & CHƯƠNG TRÌNH NGHIÊN CỨU (TRANG 12–14)",
               "Đề xuất hòa giải giữa trừu tượng hóa toán học và dữ liệu thực nghiệm thế giới thực")

    q_boxes = [
        ("Câu hỏi 1: Mạng Xã Hội Động của Truyền Thông Bí Mật",
         "Can we study dynamic social networks of secret communication modeled on the empirical behavior of secret keepers?",
         "Nghiên cứu tổ chức thông tin dưới dạng mạng xã hội động giữa những người giữ bí mật. Kiểm tra xem các bất biến chính sách bảo vệ bí mật có được duy trì qua các biến động cấu trúc mạng và các sự cố Byzantine hay không."),
        ("Câu hỏi 2: Mô Hình Hóa Năng Lực và Liên Minh Đối Thủ Bằng Thống Kê",
         "By incorporating statistics, can new models of the adversary take into account more accurate data about their capabilities?",
         "Sử dụng thống kê thực nghiệm và mô phỏng Monte Carlo để mô hình hóa chính xác hơn năng lực đối thủ; phân tích kinh tế chính trị của các liên minh giám sát (liên minh Five Eyes, Anonymous, quan hệ nhà nước - cá nhân)."),
        ("Câu hỏi 3: Ứng Dụng Phương Pháp Liên Ngành & Hệ Thống Phức Hợp",
         "Can cryptography incorporate interdisciplinary methods and insights that can capture the complexity of secure communication?",
         "Áp dụng các công cụ từ lý thuyết hệ thống phức hợp, động lực học mạng và dịch tễ học để mô hình hóa sự lan truyền thông tin, nhận diện các điểm chuyển pha khi bị tấn công."),
        ("Câu hỏi 4: Triết Học Giúp Tái Cấu Trúc Các Giả Định Ngầm Định",
         "Will philosophy let us make explicit and then reformulate the hidden assumptions of cryptographic communication, and vice versa?",
         "Vận dụng triết học để chất vấn các giả định nền tảng của Shannon; xem xét ranh giới động giữa con người và môi trường kỹ thuật số theo Giả thuyết Tâm trí Mở rộng (Extended Mind Hypothesis).")
    ]

    for idx, (q_title, q_en, q_desc) in enumerate(q_boxes):
        top_pos = Inches(1.85) + idx * Inches(1.18)
        add_card(s9, Inches(0.8), top_pos, Inches(11.733), Inches(1.05), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
        t = s9.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.06), Inches(11.333), Inches(0.92))
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

    f_box = s9.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Trích dẫn nguyên văn 4 câu hỏi nghiên cứu mở: Harry Halpin (2025), Section 7, tr. 12–14."
    p_f.font.size = Pt(10)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 10: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG RED TEAM VS BLUE TEAM (PRE-DEMO)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10)
    add_header(s10, "KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG: RED TEAM VS BLUE TEAM", "SƠ ĐỒ 4: KIẾN TRÚC DIỄN TẬP ĐỐI KHÁNG (PRE-DEMO ARCHITECTURE)",
               "Mô hình đối kháng 2 màn hình trên tập dữ liệu kaggle_dataset (2.485 tệp tin) kiểm chứng trực quan luận điểm bài báo")

    img_s10_path = os.path.join(os.path.dirname(__file__), "docs", "Slide14.drawio.png")
    if os.path.exists(img_s10_path):
        img_w = Inches(9.6)
        img_h = Inches(4.15)
        img_l = (prs.slide_width - img_w) / 2
        img_t = Inches(1.72)
        s10.shapes.add_picture(img_s10_path, img_l, img_t, img_w, img_h)

    # Khối ghi chú tích hợp phương pháp luận & 3 kịch bản kiểm chứng (gọn gàng, tinh tế)
    add_card(s10, Inches(0.8), Inches(5.98), Inches(11.733), Inches(0.82), COLOR_BG_PANEL, COLOR_BORDER_BLUE)
    t_note = s10.shapes.add_textbox(Inches(1.0), Inches(6.02), Inches(11.333), Inches(0.74))
    tf_n = t_note.text_frame
    tf_n.word_wrap = True
    
    p_n1 = tf_n.paragraphs[0]
    r_n1 = p_n1.add_run()
    r_n1.text = "Mô hình hóa thực nghiệm độc lập: "
    r_n1.font.bold = True
    r_n1.font.size = Pt(10.5)
    r_n1.font.color.rgb = COLOR_BRAND_BLUE
    r_n2 = p_n1.add_run()
    r_n2.text = "Bài báo Halpin là khảo luận triết học thuần túy (không có mã nguồn). Nhóm tự phát triển hệ thống kiểm tra toàn vẹn trên tập dữ liệu kaggle_dataset để kiểm chứng trực quan 3 luận điểm cốt lõi:"
    r_n2.font.size = Pt(9.5)
    r_n2.font.color.rgb = COLOR_TEXT_PRIMARY

    p_n2 = tf_n.add_paragraph()
    p_n2.space_before = Pt(2)
    r_n3 = p_n2.add_run()
    r_n3.text = "• (1) Nhiễu vs Kẻ thù (Mã CRC32 bị qua mặt trên GF(2) vs SHA-256 phi tuyến) | • (2) Phân tách miền nút lá RFC 6962 | • (3) Rò rỉ kênh kề thời gian (Kocher 1996)."
    r_n3.font.size = Pt(9.5)
    r_n3.font.color.rgb = COLOR_TEXT_SECONDARY

    f_box = s10.shapes.add_textbox(Inches(0.8), Inches(6.88), Inches(11.7), Inches(0.35))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Sơ đồ kiến trúc diễn tập đối kháng độc lập do sinh viên thiết kế minh họa các khái niệm an ninh."
    p_f.font.size = Pt(9.5)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 11: SLIDE CHUYỂN TIẾP: DEMO (NỀN SÁNG, ĐÚNG MỘ CHỮ "DEMO")
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    bg11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = RGBColor(248, 250, 252) # Nền sáng thanh lịch
    bg11.line.fill.background()

    demo_box = s11.shapes.add_textbox(Inches(2.0), Inches(2.6), Inches(9.333), Inches(2.2))
    tf11 = demo_box.text_frame
    tf11.word_wrap = True
    p11 = tf11.paragraphs[0]
    p11.alignment = PP_ALIGN.CENTER
    p11.text = "DEMO"
    p11.font.name = "Calibri"
    p11.font.size = Pt(88)
    p11.font.bold = True
    p11.font.color.rgb = RGBColor(30, 64, 175) # Royal Blue nổi bật

    # =========================================================================
    # SLIDE 12: BẢNG TỔNG KẾT ĐÓNG GÓP HỌC THUẬT (HARRY HALPIN 2025)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_bg(s12)
    add_header(s12, "TỔNG KẾT BÁO CÁO: ĐÓNG GÓP HỌC THUẬT & BÀI HỌC AN NINH", "SECTION 7 — TỔNG KẾT TOÀN VĂN BÀI BÁO (HARRY HALPIN 2025)",
               "Bảng đối sánh giữa các giả định hình thức truyền thống và nhận thức luận của Harry Halpin")

    # Tạo bảng tổng kết 5 hàng x 4 cột
    rows_cnt = 5
    cols_cnt = 4
    tbl_shape = s12.shapes.add_table(rows_cnt, cols_cnt, Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.7))
    tbl = tbl_shape.table

    # Độ rộng 4 cột
    tbl.columns[0].width = Inches(2.1)
    tbl.columns[1].width = Inches(3.1)
    tbl.columns[2].width = Inches(3.3)
    tbl.columns[3].width = Inches(3.233)

    # Tiêu đề các cột
    headers_text = [
        "Trụ Cột Khảo Cứu",
        "Mô Hình Toán Hình Thức\n(Shannon & Hộp Đen)",
        "Phê Phán Nhận Thức Luận\n(Harry Halpin 2025)",
        "Bài Học & Định Hướng\nAn Ninh Thực Tế"
    ]
    for c_idx, h_text in enumerate(headers_text):
        c = tbl.cell(0, c_idx)
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_BRAND_BLUE
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_c = c.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = Inches(0.08)
        tf_c.margin_right = Inches(0.08)
        tf_c.margin_top = Inches(0.05)
        tf_c.margin_bottom = Inches(0.05)
        p_c = tf_c.paragraphs[0]
        p_c.alignment = PP_ALIGN.CENTER
        p_c.text = h_text
        p_c.font.bold = True
        p_c.font.size = Pt(10.5)
        p_c.font.color.rgb = RGBColor(255, 255, 255)

    # Dữ liệu 4 hàng nội dung
    table_data = [
        ("1. Bản chất thông tin &\nBí mật an toàn (Sec 1–2)",
         "Coi thông tin là xác suất thống kê; kênh truyền chỉ chịu nhiễu ngẫu nhiên; an toàn tuyệt đối là One-Time Pad.",
         "Thông tin mang ngữ nghĩa khách quan (Dretske); đối thủ là tác nhân có chủ đích, liên tục tìm cách khai thác.",
         "Phân định rạch ròi giữa bắt lỗi ngẫu nhiên và an ninh đối kháng; mã tuyến tính (CRC32) không thể thay thế hàm băm phi tuyến (SHA-256)."),

        ("2. Giới hạn của sự\nhình thức hóa (Sec 5)",
         "Coi thuật toán mật mã là hộp đen toán học lý tưởng; thời gian và phần cứng thực thi là vô hình.",
         "Sự trừu tượng hóa làm mật mã học mù quáng trước thế giới vật lý: bỏ qua thời gian CPU, năng lượng, lỗi cài đặt.",
         "Xây dựng hệ thống an toàn đa tầng; áp dụng hàm so sánh thời gian hằng số (hmac.compare_digest) triệt tiêu rò rỉ kênh kề."),

        ("3. Đối thủ trong\nthế giới thực (Sec 6)",
         "Đối thủ bị giới hạn trong năng lực tính toán đa thức (PPT) hoặc tấn công vào đường truyền (Dolev-Yao).",
         "Snowden tiết lộ đối thủ thể chế (NSA) có nguồn lực khổng lồ: khai thác siêu dữ liệu, cài cửa sau chuẩn NIST.",
         "Mở rộng nghiên cứu mật mã chống cửa sau (Cliptography); bảo vệ siêu dữ liệu và cấu trúc dữ liệu minh bạch (RFC 6962)."),

        ("4. Chương trình nghiên cứu\ntương lai (Sec 7)",
         "Mật mã học là phân nhánh độc lập thuần túy của toán học rời rạc và khoa học máy tính lý thuyết.",
         "Đặt ra 4 câu hỏi mở: tích hợp phương pháp liên ngành, dùng thống kê lượng hóa đối thủ, mô hình hóa hệ thống phức hợp.",
         "Kết hợp toán học với triết học nhận thức và dữ liệu thực nghiệm để xây dựng một khoa học an ninh hoàn chỉnh.")
    ]

    for r_idx, (col0, col1, col2, col3) in enumerate(table_data, start=1):
        bg_row = RGBColor(248, 250, 252) if r_idx % 2 == 1 else RGBColor(255, 255, 255)
        row_vals = [col0, col1, col2, col3]
        for c_idx, val in enumerate(row_vals):
            cell = tbl.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_row
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf_cell = cell.text_frame
            tf_cell.word_wrap = True
            tf_cell.margin_left = Inches(0.08)
            tf_cell.margin_right = Inches(0.08)
            tf_cell.margin_top = Inches(0.05)
            tf_cell.margin_bottom = Inches(0.05)
            p_cell = tf_cell.paragraphs[0]
            p_cell.text = val
            if c_idx == 0:
                p_cell.font.bold = True
                p_cell.font.size = Pt(10)
                p_cell.font.color.rgb = COLOR_BRAND_BLUE
            else:
                p_cell.font.size = Pt(9.5)
                p_cell.font.color.rgb = COLOR_TEXT_PRIMARY

    f_box = s12.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Tài liệu tham khảo chính: Halpin (2025), Shannon (1949), Goldwasser-Micali (1984), Dolev-Yao (1983), Kocher (1996), RFC 6962."
    p_f.font.size = Pt(9.5)
    p_f.font.italic = True
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 13: SLIDE CẢM ƠN & Q&A (ĐỒNG BỘ THIẾT KẾ 100% VỚI SLIDE BÌA)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_bg(s13)

    # Dải màu xanh dương dọc bên lề trái đồng bộ nhận diện
    left_bar13 = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.8), Inches(0.18), Inches(5.9))
    left_bar13.fill.solid()
    left_bar13.fill.fore_color.rgb = COLOR_ACCENT_BLUE
    left_bar13.line.fill.background()

    # Khối tiêu đề Trường & Khoa ở phần trên
    t_inst13 = s13.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11.0), Inches(0.9))
    tf_inst13 = t_inst13.text_frame
    tf_inst13.word_wrap = True
    p_uni13 = tf_inst13.paragraphs[0]
    p_uni13.alignment = PP_ALIGN.CENTER
    p_uni13.text = "Trường Đại học Tài nguyên và Môi trường TP. Hồ Chí Minh"
    p_uni13.font.size = Pt(15)
    p_uni13.font.bold = True
    p_uni13.font.color.rgb = COLOR_TEXT_PRIMARY

    p_fac13 = tf_inst13.add_paragraph()
    p_fac13.alignment = PP_ALIGN.CENTER
    p_fac13.space_before = Pt(4)
    p_fac13.text = "Khoa Công nghệ thông tin"
    p_fac13.font.size = Pt(14)
    p_fac13.font.bold = True
    p_fac13.font.color.rgb = COLOR_TEXT_SECONDARY

    # Khối cảm ơn & Q&A ở giữa
    t_thank = s13.shapes.add_textbox(Inches(1.2), Inches(2.6), Inches(11.0), Inches(2.0))
    tf_thank = t_thank.text_frame
    tf_thank.word_wrap = True

    p_thk = tf_thank.paragraphs[0]
    p_thk.alignment = PP_ALIGN.CENTER
    p_thk.text = "CẢM ƠN THẦY VÀ CÁC BẠN ĐÃ LẮNG NGHE!"
    p_thk.font.size = Pt(32)
    p_thk.font.bold = True
    p_thk.font.color.rgb = COLOR_ACCENT_BLUE

    p_qa = tf_thank.add_paragraph()
    p_qa.alignment = PP_ALIGN.CENTER
    p_qa.space_before = Pt(14)
    p_qa.text = "Q & A — Xin đón nhận các câu hỏi và ý kiến đóng góp của Thầy và Hội đồng"
    p_qa.font.size = Pt(16)
    p_qa.font.italic = True
    p_qa.font.color.rgb = COLOR_TEXT_SECONDARY

    # Khối thông tin đề tài & sinh viên ở góc dưới
    t_info13 = s13.shapes.add_textbox(Inches(1.4), Inches(4.9), Inches(10.0), Inches(1.8))
    tf_info13 = t_info13.text_frame
    tf_info13.word_wrap = True

    p_t13 = tf_info13.paragraphs[0]
    p_t13.text = "Đề tài DT005: Hệ thống kiểm tra toàn vẹn tệp bằng SHA-256, HMAC và cây Merkle phân cấp động"
    p_t13.font.size = Pt(13)
    p_t13.font.bold = True
    p_t13.font.color.rgb = COLOR_TEXT_TITLE

    p_c13 = tf_info13.add_paragraph()
    p_c13.space_before = Pt(4)
    p_c13.text = "Học phần: An toàn và bảo mật hệ thống thông tin"
    p_c13.font.size = Pt(12.5)
    p_c13.font.color.rgb = COLOR_TEXT_PRIMARY

    p_lec13 = tf_info13.add_paragraph()
    p_lec13.space_before = Pt(4)
    p_lec13.text = "Giảng viên hướng dẫn: ThS. Phạm Trọng Huynh"
    p_lec13.font.size = Pt(12.5)
    p_lec13.font.color.rgb = COLOR_TEXT_PRIMARY

    p_stu13 = tf_info13.add_paragraph()
    p_stu13.space_before = Pt(4)
    r13_1 = p_stu13.add_run()
    r13_1.text = "Sinh viên thực hiện: "
    r13_1.font.size = Pt(12.5)
    r13_1.font.color.rgb = COLOR_TEXT_PRIMARY
    r13_2 = p_stu13.add_run()
    r13_2.text = "Võ Anh Kiệt"
    r13_2.font.bold = True
    r13_2.font.size = Pt(12.5)
    r13_2.font.color.rgb = COLOR_BRAND_BLUE
    r13_3 = p_stu13.add_run()
    r13_3.text = " — MSSV: 1150080143 — Lớp: 11_CNPM2"
    r13_3.font.size = Pt(12.5)
    r13_3.font.color.rgb = COLOR_TEXT_PRIMARY

    # Lưu file
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Bao_Cao_Giua_Ky_DT05_Final.pptx")
    prs.save(out_path)
    
    root_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bao_Cao_Giua_Ky_DT05_Final.pptx")
    prs.save(root_path)

    print(f"[THÀNH CÔNG] Đã tạo bài thuyết trình giữa kỳ hoàn chỉnh tại:\n  1. {out_path}\n  2. {root_path}")
    print(f"Tổng số slide: {len(prs.slides)} slide chuẩn 16:9 widescreen.")


if __name__ == "__main__":
    create_deck()
