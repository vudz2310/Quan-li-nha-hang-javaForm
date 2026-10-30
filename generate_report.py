# -*- coding: utf-8 -*-
"""
Script to generate the complete Word Document (.docx) for the Java Restaurant Management Project report.
University: Trường Đại học Kinh doanh và Công nghệ Hà Nội (HUBT)
Instructor: Thầy Mai Xuân Hùng
Student: 1 student (All MSSV removed)
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="D0D7DE", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_page_number(run):
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)

def create_report():
    doc = docx.Document()

    # Configure A4 Page and Margins (Left 3.0cm, Right 2.0cm, Top 2.0cm, Bottom 2.0cm)
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)
        
        # Header & Footer setup
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Trường ĐH Kinh doanh & Công nghệ Hà Nội — Đồ án Lập trình Java")
        hrun.font.name = "Times New Roman"
        hrun.font.size = Pt(9)
        hrun.font.italic = True
        hrun.font.color.rgb = RGBColor(120, 120, 120)

        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run()
        frun.font.name = "Times New Roman"
        frun.font.size = Pt(10)
        frun.font.color.rgb = RGBColor(80, 80, 80)
        add_page_number(frun)

    # Set base styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(13)
    normal_style.font.color.rgb = RGBColor(30, 30, 30)

    # Helper functions for content
    def add_title(text, size=20, bold=True, color=(27, 54, 93), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.2
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.line_spacing = 1.25
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = RGBColor(27, 54, 93) # Dark Navy
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.line_spacing = 1.25
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(44, 82, 130) # Slate Blue
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.line_spacing = 1.25
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.italic = True
        run.font.color.rgb = RGBColor(45, 55, 72)
        return p

    def add_p(text, bold_prefix="", indent=True):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.3
        if indent:
            p.paragraph_format.first_line_indent = Cm(1.0)
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(13)
            run_b.font.bold = True
            run_b.font.color.rgb = RGBColor(30, 30, 30)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(30, 30, 30)
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.25
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(13)
            run_b.font.bold = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        return p

    def add_img(path, caption, width_in=5.6):
        if os.path.exists(path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            run = p.add_run()
            run.add_picture(path, width=Inches(width_in))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(8)
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = 'Times New Roman'
            run_cap.font.size = Pt(11)
            run_cap.font.italic = True
            run_cap.font.bold = True
            run_cap.font.color.rgb = RGBColor(80, 80, 80)
        else:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_cap = p_cap.add_run(f"[{caption} - Vui lòng chèn ảnh: {path}]")
            run_cap.font.italic = True
            run_cap.font.color.rgb = RGBColor(180, 50, 50)

    def add_styled_table(headers, rows_data, col_widths=None):
        table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_borders(table, color="B0BEC5", sz="4")

        # Header Row
        hdr_row = table.rows[0]
        # Repeat header on new pages
        trPr = hdr_row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        for idx, title in enumerate(headers):
            cell = hdr_row.cells[idx]
            set_cell_background(cell, "1B365D") # Navy blue
            set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(title)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

        # Body Rows
        for r_idx, row in enumerate(rows_data):
            row_cells = table.rows[r_idx + 1]
            bg_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row):
                cell = row_cells.cells[c_idx]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                if c_idx == 0 or val in ["PK", "FK", "NUMBER", "VARCHAR2", "DATE", "Pass", "Đạt"]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(str(val))
                run.font.name = 'Times New Roman'
                run.font.size = Pt(10.5)
                run.font.color.rgb = RGBColor(40, 40, 40)

        # Set widths if specified
        if col_widths:
            for row in table.rows:
                for idx, w in enumerate(col_widths):
                    row.cells[idx].width = Inches(w)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ==================== TRANG BÌA CHÍNH ====================
    add_title("TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI", size=14, bold=True, color=(0, 0, 0), space_after=4)
    add_title("KHOA CÔNG NGHỆ THÔNG TIN", size=15, bold=True, color=(27, 54, 93), space_after=24)

    # Logo
    if os.path.exists("src/Icons/logo_register.png"):
        add_img("src/Icons/logo_register.png", "", width_in=1.8)

    add_title("BÁO CÁO ĐỒ ÁN MÔN HỌC", size=16, bold=True, color=(100, 100, 100), space_after=6)
    add_title("MÔN: LẬP TRÌNH JAVA / CÔNG NGHỆ JAVA", size=15, bold=True, color=(27, 54, 93), space_after=18)

    add_title("ĐỀ TÀI:", size=14, bold=False, color=(50, 50, 50), space_after=6)
    add_title("XÂY DỰNG HỆ THỐNG QUẢN LÝ NHÀ HÀNG\nROYAL THE DREAMERS RESTAURANT", size=19, bold=True, color=(192, 0, 0), space_after=40)

    # Info table for cover
    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    info_p.paragraph_format.left_indent = Cm(3.5)
    info_p.paragraph_format.line_spacing = 1.4
    info_p.paragraph_format.space_after = Pt(50)

    def add_info_line(p, label, value):
        r1 = p.add_run(f"{label:<28}: ")
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(13)
        r1.font.bold = True
        r2 = p.add_run(f"{value}\n")
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(13)

    add_info_line(info_p, "Giảng viên hướng dẫn", "Thầy Mai Xuân Hùng")
    add_info_line(info_p, "Sinh viên thực hiện", "[Họ và tên sinh viên]")
    add_info_line(info_p, "Lớp", "[Lớp sinh viên]")
    add_info_line(info_p, "Khóa học", "K30")
    add_info_line(info_p, "Ngành đào tạo", "Công nghệ Thông tin")

    add_title("Hà Nội, Năm 2026", size=13, bold=False, color=(60, 60, 60), space_after=0)
    doc.add_page_break()

    # ==================== TRANG BÌA PHỤ ====================
    add_title("TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI", size=14, bold=True, color=(0, 0, 0), space_after=4)
    add_title("KHOA CÔNG NGHỆ THÔNG TIN", size=15, bold=True, color=(27, 54, 93), space_after=30)

    add_title("BÁO CÁO ĐỒ ÁN MÔN HỌC", size=15, bold=True, color=(100, 100, 100), space_after=6)
    add_title("XÂY DỰNG HỆ THỐNG QUẢN LÝ NHÀ HÀNG\nROYAL THE DREAMERS RESTAURANT", size=18, bold=True, color=(27, 54, 93), space_after=40)

    info_p2 = doc.add_paragraph()
    info_p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    info_p2.paragraph_format.left_indent = Cm(3.5)
    info_p2.paragraph_format.line_spacing = 1.4
    info_p2.paragraph_format.space_after = Pt(80)

    add_info_line(info_p2, "Giảng viên hướng dẫn", "Thầy Mai Xuân Hùng")
    add_info_line(info_p2, "Sinh viên thực hiện", "[Họ và tên sinh viên]")
    add_info_line(info_p2, "Ngành", "Công nghệ Thông tin")
    add_info_line(info_p2, "Hệ đào tạo", "Đại học chính quy")

    add_title("Hà Nội, Năm 2026", size=13, bold=False, color=(60, 60, 60), space_after=0)
    doc.add_page_break()

    # ==================== LỜI CẢM ƠN ====================
    add_title("LỜI CẢM ƠN", size=16, bold=True, color=(27, 54, 93), space_after=18)
    add_p("Lời đầu tiên, em xin gửi lời cảm ơn trân trọng và sâu sắc nhất đến Thầy Mai Xuân Hùng – Giảng viên trực tiếp giảng dạy và hướng dẫn bộ môn Lập trình Java tại Trường Đại học Kinh doanh và Công nghệ Hà Nội (HUBT). Trong suốt quá trình học tập và triển khai đồ án, với sự tận tình chỉ bảo, định hướng chuyên môn sâu sắc và những ý kiến đóng góp vô cùng quý báu của Thầy, em đã từng bước tháo gỡ được những vướng mắc về kiến trúc phần mềm, kỹ thuật kết nối cơ sở dữ liệu và xử lý giao diện người dùng để hoàn thành đồ án một cách hoàn thiện nhất.")
    add_p("Em cũng xin trân trọng gửi lời cảm ơn tới Ban Giám hiệu, các thầy cô giáo trong Khoa Công nghệ Thông tin – Trường Đại học Kinh doanh và Công nghệ Hà Nội đã luôn tâm huyết giảng dạy, truyền đạt cho em nền tảng kiến thức công nghệ vững vàng trong suốt các kỳ học vừa qua, tạo tiền đề thuận lợi để em vận dụng vào bài toán thực tiễn.")
    add_p("Do kiến thức và kinh nghiệm thực tế của bản thân còn có những hạn chế nhất định, đề tài không thể tránh khỏi những thiếu sót. Em rất mong nhận được những ý kiến nhận xét, chỉ dẫn và góp ý quý giá từ Thầy Mai Xuân Hùng cùng các quý thầy cô để sản phẩm phần mềm cũng như vốn kiến thức chuyên ngành của em ngày càng được hoàn thiện hơn.")
    add_p("Em xin kính chúc Thầy luôn dồi dào sức khỏe, niềm vui và luôn gặt hái được nhiều thành công hơn nữa trong sự nghiệp giáo dục cao quý!")
    
    sign_p = doc.add_paragraph()
    sign_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sign_p.paragraph_format.space_before = Pt(20)
    sign_p.paragraph_format.space_after = Pt(4)
    r_date = sign_p.add_run("Hà Nội, ngày ... tháng ... năm 2026\nSinh viên thực hiện\n\n\n\n")
    r_date.font.italic = True
    r_name = sign_p.add_run("[Họ và tên sinh viên]")
    r_name.font.bold = True
    doc.add_page_break()

    # ==================== NHẬN XÉT CỦA GIẢNG VIÊN ====================
    add_title("NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN", size=16, bold=True, color=(27, 54, 93), space_after=16)
    add_p("Họ và tên giảng viên: Thầy Mai Xuân Hùng", bold_prefix="1. ", indent=False)
    add_p("Cơ quan công tác: Khoa Công nghệ Thông tin – Trường Đại học Kinh doanh và Công nghệ Hà Nội", bold_prefix="2. ", indent=False)
    add_p("Tên đề tài đồ án: Xây dựng hệ thống quản lý nhà hàng Royal The Dreamers Restaurant", bold_prefix="3. ", indent=False)
    add_p("Sinh viên thực hiện: [Họ và tên sinh viên]", bold_prefix="4. ", indent=False)
    
    add_p("Ý thức, thái độ học tập và tinh thần kỷ luật của sinh viên:", bold_prefix="5. ", indent=False)
    add_p("..........................................................................................................................................................", indent=False)
    add_p("..........................................................................................................................................................", indent=False)
    add_p("Đánh giá về chất lượng nội dung báo cáo và tính khoa học:", bold_prefix="6. ", indent=False)
    add_p("..........................................................................................................................................................", indent=False)
    add_p("..........................................................................................................................................................", indent=False)
    add_p("Đánh giá về sản phẩm phần mềm (Mã nguồn, tính năng, độ ổn định):", bold_prefix="7. ", indent=False)
    add_p("..........................................................................................................................................................", indent=False)
    add_p("..........................................................................................................................................................", indent=False)
    add_p("Kết luận và Điểm đánh giá:", bold_prefix="8. ", indent=False)
    add_p("- Điểm bằng số: .......................................   Điểm bằng chữ: ..............................................................", indent=False)

    sign_gv = doc.add_paragraph()
    sign_gv.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sign_gv.paragraph_format.space_before = Pt(30)
    r_gv1 = sign_gv.add_run("Hà Nội, ngày ... tháng ... năm 2026\nGiảng viên hướng dẫn\n(Ký và ghi rõ họ tên)\n\n\n\n")
    r_gv1.font.italic = True
    r_gv2 = sign_gv.add_run("Thầy Mai Xuân Hùng")
    r_gv2.font.bold = True
    doc.add_page_break()

    # ==================== DANH MỤC THUẬT NGỮ VIẾT TẮT ====================
    add_title("DANH MỤC CÁC TỪ VIẾT TẮT", size=16, bold=True, color=(27, 54, 93), space_after=14)
    abbr_headers = ["Ký hiệu viết tắt", "Cụm từ tiếng Anh / Thuật ngữ", "Ý nghĩa / Diễn giải"]
    abbr_rows = [
        ["HUBT", "Hanoi University of Business and Technology", "Trường Đại học Kinh doanh và Công nghệ Hà Nội"],
        ["JDK", "Java Development Kit", "Bộ công cụ phát triển ứng dụng Java"],
        ["JRE", "Java Runtime Environment", "Môi trường thực thi ứng dụng Java"],
        ["IDE", "Integrated Development Environment", "Môi trường phát triển tích hợp (NetBeans, IntelliJ, VS Code)"],
        ["MVC", "Model - View - Controller", "Mô hình kiến trúc phân tầng trong công nghệ phần mềm"],
        ["DBMS / RDBMS", "Relational Database Management System", "Hệ quản trị cơ sở dữ liệu quan hệ (Oracle Database)"],
        ["JDBC", "Java Database Connectivity", "API chuẩn kết nối ứng dụng Java với hệ cơ sở dữ liệu"],
        ["SQL", "Structured Query Language", "Ngôn ngữ truy vấn dữ liệu có cấu trúc"],
        ["ERD", "Entity Relationship Diagram", "Sơ đồ quan hệ thực thể liên kết"],
        ["OTP", "One-Time Password", "Mật khẩu xác minh bảo mật sử dụng một lần"],
        ["SMTP", "Simple Mail Transfer Protocol", "Giao thức truyền tải thư điện tử đơn giản"],
        ["UI / UX", "User Interface / User Experience", "Giao diện người dùng / Trải nghiệm người dùng"],
        ["CTHD", "Chi tiết Hóa đơn", "Bảng dữ liệu lưu các món ăn của từng hóa đơn"],
        ["PNK / CTNK", "Phiếu Nhập Kho / Chi tiết Nhập Kho", "Quản lý hoạt động nhập nguyên liệu vào kho"],
        ["PXK / CTXK", "Phiếu Xuất Kho / Chi tiết Xuất Kho", "Quản lý hoạt động xuất nguyên liệu phục vụ chế biến"],
        ["HR", "Human Resources", "Phân hệ quản lý nhân sự, nhân viên nhà hàng"]
    ]
    add_styled_table(abbr_headers, abbr_rows, col_widths=[1.5, 2.5, 2.6])
    doc.add_page_break()

    # ==================== DANH MỤC HÌNH VẼ ====================
    add_title("DANH MỤC HÌNH VẼ MINH HỌA", size=16, bold=True, color=(27, 54, 93), space_after=14)
    img_list = [
        ["Hình 1.1", "Mô hình kiến trúc MVC (Model - View - Controller) ứng dụng trong dự án"],
        ["Hình 2.1", "Sơ đồ Use Case tổng quát của Hệ thống Quản lý Nhà hàng Royal The Dreamers"],
        ["Hình 2.2", "Sơ đồ Thực thể Liên kết (ERD - Entity Relationship Diagram) của cơ sở dữ liệu"],
        ["Hình 2.3", "Cấu trúc gói (Package Explorer) của dự án trong môi trường phát triển Java"],
        ["Hình 3.1", "Giao diện Đăng ký tài khoản khách hàng mới"],
        ["Hình 3.2", "Giao diện Đăng nhập hệ thống phân quyền tự động"],
        ["Hình 3.3", "Giao diện Giới thiệu nhà hàng (About Us)"],
        ["Hình 3.4", "Giao diện Khách hàng đặt món ăn theo 12 Cung hoàng đạo"],
        ["Hình 3.5", "Giao diện Đổi điểm tích lũy lấy Voucher ưu đãi"],
        ["Hình 3.6", "Giao diện Xem chi tiết hóa đơn thanh toán"],
        ["Hình 3.7", "Giao diện Quản lý thông tin cá nhân khách hàng"],
        ["Hình 3.8", "Giao diện Quản lý trạng thái bàn ăn dành cho Nhân viên phục vụ"],
        ["Hình 3.9", "Giao diện Lập phiếu gọi món cho khách tại bàn"],
        ["Hình 3.10", "Giao diện Thanh toán hóa đơn và in phiếu"],
        ["Hình 3.11", "Giao diện Tra cứu thông tin và điểm tích lũy khách hàng"],
        ["Hình 3.12", "Giao diện Quản lý nguyên vật liệu và kho hàng"],
        ["Hình 3.13", "Giao diện Quản lý danh mục món ăn dành cho Quản trị viên"],
        ["Hình 3.14", "Giao diện Chỉnh sửa và cập nhật thực đơn món ăn"],
        ["Hình 3.15", "Giao diện Quản lý hồ sơ nhân sự nhà hàng"],
        ["Hình 3.16", "Giao diện Thêm mới nhân viên và phân quyền"],
        ["Hình 3.17", "Giao diện Báo cáo doanh thu và biểu đồ thống kê kinh doanh"]
    ]
    add_styled_table(["Mã hình", "Tên gọi hình ảnh minh họa"], img_list, col_widths=[1.5, 5.1])
    doc.add_page_break()

    # ==================== DANH MỤC BẢNG BIỂU ====================
    add_title("DANH MỤC BẢNG BIỂU", size=16, bold=True, color=(27, 54, 93), space_after=14)
    tbl_list = [
        ["Bảng 1.1", "Bảng khảo sát các chức năng yêu cầu theo từng phân hệ người dùng"],
        ["Bảng 2.1", "Đặc tả các tác nhân (Actors) tham gia vào hệ thống"],
        ["Bảng 2.2", "Đặc tả Use Case Đăng ký tài khoản và xác thực OTP Email"],
        ["Bảng 2.3", "Đặc tả Use Case Đặt món tại bàn"],
        ["Bảng 2.4", "Đặc tả Use Case Thanh toán hóa đơn và áp dụng Voucher"],
        ["Bảng 2.5", "Đặc tả chi tiết cấu trúc Bảng NguoiDung"],
        ["Bảng 2.6", "Đặc tả chi tiết cấu trúc Bảng NhanVien"],
        ["Bảng 2.7", "Đặc tả chi tiết cấu trúc Bảng KhachHang"],
        ["Bảng 2.8", "Đặc tả chi tiết cấu trúc Bảng MonAn"],
        ["Bảng 2.9", "Đặc tả chi tiết cấu trúc Bảng Ban"],
        ["Bảng 2.10", "Đặc tả chi tiết cấu trúc Bảng Voucher"],
        ["Bảng 2.11", "Đặc tả chi tiết cấu trúc Bảng HoaDon"],
        ["Bảng 2.12", "Đặc tả chi tiết cấu trúc Bảng CTHD"],
        ["Bảng 2.13", "Đặc tả chi tiết cấu trúc Bảng NguyenLieu và Bảng Kho"],
        ["Bảng 2.14", "Đặc tả chi tiết cấu trúc Bảng PhieuNK, CTNK, PhieuXK, CTXK"],
        ["Bảng 3.1", "Bảng thông số môi trường cài đặt và cấu hình phần mềm"],
        ["Bảng 3.2", "Kịch bản kiểm thử (Test Cases) các chức năng chính của hệ thống"]
    ]
    add_styled_table(["Mã bảng", "Tên bảng thống kê / đặc tả"], tbl_list, col_widths=[1.5, 5.1])
    doc.add_page_break()

    # ==================== MỤC LỤC TỔNG QUAN ====================
    add_title("MỤC LỤC NỘI DUNG", size=16, bold=True, color=(27, 54, 93), space_after=14)
    toc_data = [
        ["LỜI CẢM ƠN", "i"],
        ["NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN", "ii"],
        ["DANH MỤC CÁC TỪ VIẾT TẮT", "iii"],
        ["DANH MỤC HÌNH VẼ MINH HỌA", "iv"],
        ["DANH MỤC BẢNG BIỂU", "v"],
        ["LỜI MỞ ĐẦU", "1"],
        ["CHƯƠNG 1: TỔNG QUAN VÀ KHẢO SÁT HIỆN TRẠNG ĐỀ TÀI", "3"],
        ["  1.1. Giới thiệu tổng quan về đề tài quản lý nhà hàng", "3"],
        ["  1.2. Khảo sát hiện trạng và quy trình nghiệp vụ thực tế", "4"],
        ["    1.2.1. Quy trình tiếp đón khách và quản lý bàn ăn", "4"],
        ["    1.2.2. Quy trình đặt món và phục vụ tại bàn", "5"],
        ["    1.2.3. Quy trình thanh toán, tích điểm và voucher ưu đãi", "5"],
        ["    1.2.4. Quy trình quản lý kho và nguyên vật liệu", "6"],
        ["    1.2.5. Quy trình quản lý nhân sự và phân quyền người dùng", "7"],
        ["    1.2.6. Quy trình thống kê doanh thu và báo cáo tài chính", "7"],
        ["  1.3. Yêu cầu của hệ thống", "8"],
        ["    1.3.1. Yêu cầu chức năng", "8"],
        ["    1.3.2. Yêu cầu phi chức năng", "9"],
        ["  1.4. Công nghệ và môi trường phát triển", "10"],
        ["    1.4.1. Ngôn ngữ Java và Java Swing Framework", "10"],
        ["    1.4.2. Hệ quản trị CSDL Oracle và thư viện JDBC ojdbc8", "11"],
        ["    1.4.3. Mô hình kiến trúc MVC (Model - View - Controller)", "12"],
        ["    1.4.4. Các thư viện bổ trợ (JavaMail API, POI, Chart, JCalendar)", "13"],
        ["CHƯƠNG 2: PHÂN TÍCH VÀ THIẾT KẾ HỆ THỐNG", "14"],
        ["  2.1. Phân tích chức năng và Sơ đồ Use Case", "14"],
        ["    2.1.1. Sơ đồ Use Case tổng quát", "14"],
        ["    2.1.2. Đặc tả các Use Case trọng tâm của hệ thống", "16"],
        ["  2.2. Phân tích và thiết kế Cơ sở dữ liệu (Database Design)", "19"],
        ["    2.2.1. Mô hình thực thể liên kết (ERD)", "19"],
        ["    2.2.2. Đặc tả chi tiết các bảng trong cơ sở dữ liệu Oracle", "20"],
        ["    2.2.3. Các ràng buộc toàn vẹn và khóa ngoại", "25"],
        ["  2.3. Thiết kế kiến trúc phần mềm và cấu trúc Class", "26"],
        ["    2.3.1. Phân tầng kiến trúc MVC trong dự án", "26"],
        ["    2.3.2. Thiết kế tầng Model (Data Transfer Objects)", "27"],
        ["    2.3.3. Thiết kế tầng Controller và Service nghiệp vụ", "28"],
        ["    2.3.4. Thiết kế tầng View và Giao diện người dùng", "29"],
        ["  2.4. Thiết kế giao diện và trải nghiệm người dùng (UI/UX Concept)", "30"],
        ["CHƯƠNG 3: CÀI ĐẶT THỬ NGHIỆM VÀ ĐÁNH GIÁ HỆ THỐNG", "31"],
        ["  3.1. Môi trường cài đặt và cấu hình hệ thống", "31"],
        ["    3.1.1. Cấu hình cơ sở dữ liệu Oracle Database", "31"],
        ["    3.1.2. Cấu hình kết nối JDBC qua DatabaseConnection Singleton", "32"],
        ["    3.1.3. Cấu hình dịch vụ gửi email xác minh OTP qua SMTP Gmail", "33"],
        ["  3.2. Kết quả cài đặt các phân hệ và giao diện thực tế", "34"],
        ["    3.2.1. Phân hệ Xác thực tài khoản (Đăng nhập, Đăng ký OTP)", "34"],
        ["    3.2.2. Phân hệ Khách hàng (Đặt món, Đổi voucher, Xem hóa đơn)", "36"],
        ["    3.2.3. Phân hệ Nhân viên phục vụ (Quản lý bàn, Gọi món, Thanh toán)", "40"],
        ["    3.2.4. Phân hệ Nhân viên kho (Quản lý nguyên liệu, Nhập - Xuất kho)", "43"],
        ["    3.2.5. Phân hệ Quản trị viên (Thực đơn, Nhân sự, Báo cáo thống kê)", "45"],
        ["  3.3. Kiểm thử phần mềm và đánh giá chất lượng (Testing)", "49"],
        ["  3.4. Đánh giá ưu điểm và hạn chế của hệ thống", "51"],
        ["KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", "53"],
        ["TÀI LIỆU THAM KHẢO", "55"]
    ]
    add_styled_table(["Mục nội dung", "Trang"], toc_data, col_widths=[5.5, 1.1])
    doc.add_page_break()

    # ==================== LỜI MỞ ĐẦU ====================
    add_title("LỜI MỞ ĐẦU", size=16, bold=True, color=(27, 54, 93), space_after=14)
    
    add_h2("1. Tính cấp thiết của đề tài")
    add_p("Trong bối cảnh nền kinh tế số và thời đại cách mạng công nghiệp 4.0 phát triển mạnh mẽ, công nghệ thông tin đã và đang len lỏi sâu rộng vào mọi lĩnh vực của đời sống xã hội. Trong đó, ngành kinh doanh ẩm thực và dịch vụ ăn uống (F&B - Food and Beverage) là một trong những lĩnh vực có tốc độ phát triển nhanh chóng và cạnh tranh khốc liệt nhất hiện nay. Nhà hàng không chỉ đơn thuần là nơi khách hàng tìm đến để thưởng thức những món ăn ngon, mà còn là không gian giao lưu bạn bè, gặp gỡ đối tác, hẹn hò và tổ chức các sự kiện văn hóa, tiệc tùng quan trọng.")
    add_p("Tuy nhiên, phần lớn các nhà hàng vừa và nhỏ hiện nay vẫn đang duy trì phương thức quản lý thủ công truyền thống: ghi chép gọi món trên giấy, chuyển giao order giữa các bộ phận bằng lời nói hoặc nhân viên chạy bàn, tính toán hóa đơn thủ công và kiểm kê nguyên vật liệu định kỳ bằng sổ sách kế toán. Cách tiếp cận này bộc lộ rất nhiều hạn chế nghiêm trọng như: thời gian chờ đợi của khách hàng kéo dài, dễ xảy ra nhầm lẫn món ăn, thất thoát nguyên liệu trong kho, số liệu doanh thu thiếu chính xác và người quản lý không thể nắm bắt tức thời tình hình kinh doanh của nhà hàng.")
    add_p("Xuất phát từ thực tiễn cấp bách đó, việc nghiên cứu, thiết kế và phát triển một hệ thống phần mềm quản lý nhà hàng hiện đại, chuyên nghiệp, khép kín từ khâu đón tiếp khách hàng, gọi món tại bàn, thanh toán hóa đơn đến quản lý xuất nhập kho và thống kê tài chính là vô cùng cần thiết. Đề tài 'Xây dựng hệ thống quản lý nhà hàng Royal The Dreamers Restaurant' được sinh viên lựa chọn triển khai nhằm giải quyết triệt để các bài toán nghiệp vụ nêu trên, đem lại hiệu quả vận hành tối ưu cho doanh nghiệp và nâng cao trải nghiệm của thực khách.")

    add_h2("2. Mục tiêu nghiên cứu của đề tài")
    add_bullet("Nghiên cứu tổng quan về quy trình vận hành thực tế tại các nhà hàng ẩm thực hiện đại, phân tích các dòng lưu chuyển thông tin giữa thực khách, nhân viên tiếp tân, nhân viên phục vụ, bếp, thủ kho và người quản lý.", "Về mặt lý thuyết: ");
    add_bullet("Vận dụng toàn diện kiến thức đã học trong môn Lập trình Java, Lập trình hướng đối tượng (OOP), Cơ sở dữ liệu và Thiết kế giao diện phần mềm để xây dựng ứng dụng hoàn chỉnh trên nền tảng Java Swing kết nối CSDL Oracle.", "Về mặt kỹ thuật: ");
    add_bullet("Xây dựng thành công phần mềm Quản lý nhà hàng 'Royal The Dreamers' với đầy đủ 4 phân hệ chuyên biệt: Khách hàng tự gọi món tại bàn, Nhân viên phục vụ/tiếp tân, Nhân viên quản lý kho nguyên vật liệu và Quản trị viên/Chủ nhà hàng điều hành kinh doanh.", "Về mặt ứng dụng: ");

    add_h2("3. Đối tượng và phạm vi nghiên cứu")
    add_bullet("Đối tượng nghiên cứu: Nghiệp vụ quản lý bán hàng, đặt bàn, thực đơn món ăn, xuất nhập tồn kho nguyên vật liệu, tích lũy điểm thưởng khách hàng thân thiết và lập báo cáo tài chính trong môi trường kinh doanh dịch vụ ăn uống.", "");
    add_bullet("Phạm vi ứng dụng: Hệ thống phần mềm dạng ứng dụng Desktop chạy trên nền tảng Java Virtual Machine (JVM), kết nối máy chủ cơ sở dữ liệu quan hệ Oracle Database, phục vụ nội bộ trong phạm vi nhà hàng Royal The Dreamers.", "");

    add_h2("4. Phương pháp nghiên cứu")
    add_bullet("Phương pháp thu thập tài liệu: Tìm hiểu các tài liệu giáo trình Lập trình Java, các đặc tả kỹ thuật của Oracle Database, JDBC và các dự án mã nguồn mở về giao diện người dùng Java Swing.", "");
    add_bullet("Phương pháp phân tích và thiết kế hệ thống: Sử dụng ngôn ngữ mô hình hóa UML để vẽ sơ đồ Use Case, mô hình hóa dữ liệu quan hệ bằng sơ đồ ERD, áp dụng mô hình kiến trúc MVC để phân tách rõ ràng các tầng nghiệp vụ.", "");
    add_bullet("Phương pháp thực nghiệm: Cài đặt viết mã nguồn (coding) từng mô-đun, tích hợp hệ thống, thực hiện kiểm thử chức năng (Black-box testing) và đánh giá độ ổn định của sản phẩm.", "");

    add_h2("5. Bố cục của đồ án")
    add_p("Báo cáo đồ án môn học được cấu trúc thành 3 chương chính như sau:")
    add_bullet("Trình bày lý do chọn đề tài, khảo sát chi tiết hiện trạng các quy trình nghiệp vụ trong nhà hàng, thu thập các yêu cầu chức năng và phi chức năng, đồng thời giới thiệu các công nghệ nền tảng được lựa chọn sử dụng trong dự án.", "Chương 1 - Tổng quan và khảo sát hiện trạng đề tài: ");
    add_bullet("Trình bày phân tích thiết kế hệ thống theo chuẩn công nghệ phần mềm: xây dựng sơ đồ Use Case và đặc tả các ca sử dụng chính, thiết kế mô hình thực thể liên kết ERD, đặc tả chi tiết cấu trúc các bảng dữ liệu Oracle và thiết kế kiến trúc phân tầng MVC.", "Chương 2 - Phân tích và thiết kế hệ thống: ");
    add_bullet("Mô tả chi tiết môi trường cài đặt, quy trình cấu hình cơ sở dữ liệu Oracle và kết nối JDBC; minh họa toàn bộ các giao diện thực tế của phần mềm kèm hình ảnh minh chứng; xây dựng kịch bản kiểm thử chức năng và đánh giá ưu/nhược điểm của hệ thống.", "Chương 3 - Cài đặt thử nghiệm và đánh giá hệ thống: ");
    add_bullet("Tổng kết những kết quả đã đạt được, chỉ ra những hạn chế còn tồn đọng và định hướng các nội dung cần nâng cấp mở rộng trong tương lai.", "Phần Kết luận và Hướng phát triển: ");
    doc.add_page_break()

    # ==================== CHƯƠNG 1 ====================
    add_title("CHƯƠNG 1: TỔNG QUAN VÀ KHẢO SÁT HIỆN TRẠNG ĐỀ TÀI", size=16, bold=True, color=(27, 54, 93), space_after=14)

    add_h2("1.1. Giới thiệu tổng quan về đề tài quản lý nhà hàng")
    add_p("Kinh doanh nhà hàng ẩm thực là một ngành dịch vụ đòi hỏi tính chuyên nghiệp, sự chính xác và tốc độ phục vụ rất cao. Trong một ca làm việc điển hình, nhà hàng phải tiếp nhận hàng chục đến hàng trăm lượt khách hàng, xử lý đồng thời nhiều lượt gọi món, điều phối món ăn từ bếp đến đúng vị trí bàn, tính toán hóa đơn kèm theo các chương trình khuyến mãi, giảm giá và quản lý dòng tiền thu chi minh bạch. Bên cạnh đó, bộ phận bếp và kho phải liên tục theo dõi lượng nguyên liệu xuất - nhập mỗi ngày để đảm bảo luôn sẵn sàng nguyên liệu chế biến, tránh tình trạng ứ đọng gây hư hỏng hoặc thiếu hụt làm gián đoạn phục vụ.")
    add_p("Tại nhiều cơ sở kinh doanh F&B truyền thống, các công việc trên chủ yếu dựa vào trí nhớ của nhân viên và các loại hóa đơn giấy viết tay. Khi lượng khách tăng đột biến vào các khung giờ cao điểm (buổi trưa hoặc buổi tối cuối tuần), mô hình quản lý thủ công bộc lộ những nhược điểm chí mạng:")
    add_bullet("Nhân viên phục vụ phải di chuyển liên tục giữa bàn ăn của khách và khu vực quầy thu ngân hoặc quầy bếp để chuyển phiếu order, làm tăng thời gian chờ đợi của khách.", "Tốc độ xử lý chậm: ");
    add_bullet("Tình trạng ghi sai tên món, sai số lượng, nhầm lẫn giữa các bàn ăn hoặc thất lạc phiếu order giấy thường xuyên xảy ra, dẫn đến trải nghiệm tiêu cực cho khách hàng.", "Dễ xảy ra nhầm lẫn: ");
    add_bullet("Khó khăn trong việc kiểm soát định lượng nguyên vật liệu, dễ xảy ra hao hụt, thất thoát mà không rõ nguyên nhân.", "Thất thoát nguyên vật liệu: ");
    add_bullet("Người quản lý phải mất hàng giờ đồng hồ vào cuối ngày hoặc cuối tháng để cộng dồn các hóa đơn giấy, tính toán doanh thu và chi phí, số liệu dễ bị sai lệch và không thể đưa ra các quyết định kinh doanh kịp thời.", "Thiếu tính thời sự trong báo cáo: ");
    add_p("Nhà hàng 'Royal The Dreamers' được xây dựng với định hướng phong cách ẩm thực Hoàng gia sang trọng, lấy cảm hứng từ 12 chòm sao hoàng đạo kết hợp không gian thưởng thức tinh tế. Để tương xứng với đẳng cấp và mô hình vận hành đó, việc tin học hóa toàn diện quy trình quản trị là yếu tố then chốt quyết định sự thành bại. Ứng dụng Quản lý nhà hàng Royal The Dreamers ra đời nhằm số hóa toàn bộ các hoạt động nghiệp vụ, mang đến một quy trình khép kín, hiện đại và chuẩn mực.")

    add_h2("1.2. Khảo sát hiện trạng và quy trình nghiệp vụ thực tế")
    add_p("Qua quá trình khảo sát thực tế tại các mô hình nhà hàng ẩm thực, các luồng nghiệp vụ cốt lõi đã được tổng hợp và chuẩn hóa như sau:")

    add_h3("1.2.1. Quy trình tiếp đón khách và quản lý bàn ăn")
    add_p("Khi khách hàng đến nhà hàng, nhân viên tiếp tân hoặc nhân viên phục vụ sẽ kiểm tra sơ đồ bàn ăn trên hệ thống. Các bàn ăn được định danh rõ ràng theo số hiệu, khu vực vị trí và phân biệt qua các trạng thái trực quan:")
    add_bullet("Bàn hiện đang không có khách, sẵn sàng để tiếp đón khách mới.", "Trạng thái 'Còn trống': ");
    add_bullet("Bàn đã có khách đang ngồi dùng bữa và đang có hóa đơn mở.", "Trạng thái 'Đang dùng bữa': ");
    add_bullet("Bàn đã được khách liên hệ đặt trước theo khung giờ xác định.", "Trạng thái 'Đã đặt trước': ");
    add_p("Nhân viên có thể linh hoạt chuyển đổi trạng thái bàn, xếp chỗ cho khách hoặc thực hiện chuyển bàn theo yêu cầu một cách nhanh chóng.")

    add_h3("1.2.2. Quy trình đặt món và phục vụ tại bàn")
    add_p("Hệ thống cung cấp hai phương thức gọi món linh hoạt:")
    add_bullet("Khách hàng có thể sử dụng màn hình gọi món tại bàn (Main_Customer_Frame). Khách hàng tự do duyệt xem thực đơn phong phú được phân loại theo 12 cung hoàng đạo (Aries, Taurus, Gemini, Cancer, Leo, Virgo, Libra, Scorpio, Sagittarius, Capricorn, Aquarius, Pisces). Mỗi món ăn đều có hình ảnh minh họa sống động, tên món, đơn giá chi tiết. Khách hàng lựa chọn số lượng món ăn và nhấn 'Thêm vào giỏ', sau đó xác nhận 'Đặt món'. Dữ liệu order ngay lập tức được đồng bộ vào hệ thống cơ sở dữ liệu.", "Phương thức 1 - Khách hàng tự gọi món: ");
    add_bullet("Nhân viên phục vụ sử dụng giao diện phân hệ nhân viên (Main_Staff_Frame) để chọn bàn, tìm kiếm món ăn theo tên hoặc mã món, chọn số lượng và ghi nhận order thay cho khách hàng.", "Phương thức 2 - Nhân viên hỗ trợ gọi món: ");

    add_h3("1.2.3. Quy trình thanh toán, tích điểm và voucher ưu đãi")
    add_p("Khi khách hàng kết thúc bữa ăn và yêu cầu thanh toán:")
    add_p("1. Nhân viên thu ngân/phục vụ truy cập chức năng 'Thanh toán', chọn số bàn tương ứng. Hệ thống tự động truy xuất toàn bộ danh sách các món ăn đã gọi trong phiên dùng bữa, tính tổng tiền món ăn.", indent=False)
    add_p("2. Áp dụng mã giảm giá (Voucher): Khách hàng có thể sử dụng các mã Voucher tích lũy được từ điểm thưởng hoặc các chương trình khuyến mãi. Hệ thống tự động kiểm tra tính hợp lệ của mã, tính số tiền được giảm theo tỷ lệ phần trăm tương ứng và khấu trừ vào hóa đơn.", indent=False)
    add_p("3. Cập nhật điểm tích lũy: Sau khi hóa đơn được thanh toán thành công, hệ thống tự động cộng dồn doanh số chi tiêu của khách hàng vào tài khoản thành viên, đồng thời quy đổi điểm tích lũy (ví dụ: mỗi 100.000 VNĐ chi tiêu tương ứng với số điểm thưởng nhất định). Điểm thưởng này có thể dùng để đổi các Voucher ưu đãi cho những lần dùng bữa tiếp theo.", indent=False)
    add_p("4. Chuyển trạng thái bàn về 'Còn trống' và hỗ trợ in hóa đơn chi tiết gửi cho khách hàng.", indent=False)

    add_h3("1.2.4. Quy trình quản lý kho và nguyên vật liệu")
    add_p("Kho nguyên vật liệu là bộ phận giữ vai trò huyết mạch của bếp nhà hàng. Quy trình quản lý kho bao gồm:")
    add_bullet("Khi nguyên vật liệu từ nhà cung cấp được chuyển đến, nhân viên kho kiểm tra quy cách, số lượng và tạo 'Phiếu Nhập Kho' (PhieuNK). Hệ thống lưu lại thông tin ngày nhập, nhân viên lập phiếu và chi tiết từng loại nguyên liệu kèm số lượng, đơn giá nhập. Số lượng tồn kho (SLTon trong bảng Kho) ngay lập tức được cộng dồn tương ứng.", "Quy trình Nhập kho: ");
    add_bullet("Hàng ngày hoặc theo từng ca, bếp yêu cầu xuất các loại rau củ, thịt cá, gia vị để sơ chế. Nhân viên kho tạo 'Phiếu Xuất Kho' (PhieuXK), ghi nhận số lượng từng nguyên liệu xuất ra. Hệ thống kiểm tra số lượng tồn hiện có, nếu đủ sẽ trừ trực tiếp vào tồn kho.", "Quy trình Xuất kho: ");
    add_bullet("Nhân viên kho có thể tra cứu nhanh danh mục nguyên liệu, số lượng tồn kho còn lại để kịp thời đề xuất kế hoạch mua sắm, tránh tình trạng đứt gãy nguồn cung.", "Kiểm kê cảnh báo: ");

    add_h3("1.2.5. Quy trình quản lý nhân sự và phân quyền người dùng")
    add_p("Nhà hàng phân chia lực lượng lao động theo các vị trí công việc rõ ràng:")
    add_bullet("Nhân viên phục vụ bàn, nhân viên tiếp tân đón khách, nhân viên thu ngân, đầu bếp, nhân viên kho và quản lý nhà hàng.", "Các chức vụ: ");
    add_bullet("Hệ thống quản lý chặt chẽ thông tin từng nhân sự: Mã NV, Tên NV, Ngày vào làm, Số điện thoại, Chức vụ, Trạng thái làm việc (Đang làm việc / Đã nghỉ việc) và Người quản lý trực tiếp.", "Hồ sơ nhân viên: ");
    add_bullet("Mỗi nhân viên được cấp một tài khoản đăng nhập tương ứng với vai trò (Role): 'Quan Ly' (Toàn quyền quản trị), 'Nhan Vien' (Phục vụ/Tiếp tân), 'Nhan Vien Kho' (Thủ kho) và 'Khach Hang' (Thực khách).", "Phân quyền: ");

    add_h3("1.2.6. Quy trình thống kê doanh thu và báo cáo tài chính")
    add_p("Dành riêng cho Quản trị viên (Admin/Chủ nhà hàng):")
    add_bullet("Theo dõi doanh thu thực tế phát sinh theo từng ngày, từng tháng hoặc từng năm.", "Báo cáo doanh thu: ");
    add_bullet("Biểu diễn trực quan tốc độ tăng trưởng doanh thu qua biểu đồ đường (Line Chart) hoặc biểu đồ cột, giúp người quản trị dễ dàng so sánh hiệu quả kinh doanh giữa các mốc thời gian.", "Biểu đồ trực quan: ");
    add_bullet("Thống kê số lượng khách hàng mới, số lượng món ăn bán chạy nhất để có kế hoạch điều chỉnh thực đơn phù hợp với thị hiếu khách hàng.", "Phân tích sản phẩm: ");

    add_h2("1.3. Yêu cầu của hệ thống")
    add_h3("1.3.1. Yêu cầu chức năng (Functional Requirements)")
    add_p("Dựa trên kết quả khảo sát nghiệp vụ, các yêu cầu chức năng được phân loại chi tiết theo 4 nhóm tác nhân sử dụng trong Bảng 1.1 dưới đây:")

    req_headers = ["Phân hệ / Tác nhân", "Mã YC", "Tên chức năng", "Mô tả chi tiết yêu cầu nghiệp vụ"]
    req_rows = [
        ["Khách hàng\n(Customer)", "YC-KH01", "Đăng ký & Xác thực OTP", "Cho phép khách hàng đăng ký tài khoản với Email, nhận mã xác thực OTP 6 số qua hộp thư để kích hoạt tài khoản."],
        ["Khách hàng", "YC-KH02", "Đăng nhập hệ thống", "Xác thực Email và Mật khẩu, điều hướng trực tiếp vào giao diện Main_Customer_Frame."],
        ["Khách hàng", "YC-KH03", "Xem thực đơn 12 Cung", "Hiển thị món ăn theo 12 Cung hoàng đạo, lọc món ăn, xem chi tiết giá và hình ảnh món."],
        ["Khách hàng", "YC-KH04", "Đặt món tại bàn", "Chọn món, chọn số lượng, thêm vào giỏ hàng và gửi yêu cầu đặt món lên máy chủ."],
        ["Khách hàng", "YC-KH05", "Đổi điểm tích lũy Voucher", "Sử dụng điểm tích lũy hiện có để đổi các mã giảm giá (10%, 20%, 50%...) theo quy định."],
        ["Khách hàng", "YC-KH06", "Xem lịch sử hóa đơn", "Tra cứu danh sách hóa đơn đã dùng bữa, xem chi tiết các món ăn và tổng chi phí."],
        ["Khách hàng", "YC-KH07", "Quản lý thông tin cá nhân", "Cập nhật tên, mật khẩu, xem tổng doanh số chi tiêu và điểm tích lũy cá nhân."],
        ["Nhân viên\n(Staff)", "YC-NV01", "Quản lý sơ đồ bàn ăn", "Xem trạng thái trực quan các bàn (Còn trống, Đang dùng bữa, Đã đặt trước), đổi trạng thái bàn."],
        ["Nhân viên", "YC-NV02", "Lập order cho khách", "Tạo hóa đơn mới gắn với số bàn, thêm món ăn và số lượng theo yêu cầu của khách tại bàn."],
        ["Nhân viên", "YC-NV03", "Thanh toán & Áp Voucher", "Tính tổng tiền, nhập mã Voucher giảm giá, xác nhận thanh toán và in phiếu thanh toán."],
        ["Nhân viên", "YC-NV04", "Quản lý thông tin khách", "Tra cứu thông tin khách hàng, số điện thoại, điểm tích lũy để phục vụ khách chu đáo."],
        ["Nhân viên Kho\n(Warehouse)", "YC-KHO01", "Quản lý danh mục nguyên liệu", "Thêm mới, cập nhật tên nguyên liệu, đơn vị tính (g, kg, ml, l), đơn giá nguyên liệu."],
        ["Nhân viên Kho", "YC-KHO02", "Quản lý tồn kho", "Theo dõi số lượng tồn kho thực tế của từng loại nguyên liệu trong nhà hàng."],
        ["Nhân viên Kho", "YC-KHO03", "Lập phiếu nhập kho", "Tạo phiếu nhập kho, chọn nguyên liệu, nhập số lượng, tự động cộng dồn tồn kho."],
        ["Nhân viên Kho", "YC-KHO04", "Lập phiếu xuất kho", "Tạo phiếu xuất kho phục vụ chế biến, kiểm tra số lượng và trừ tồn kho tự động."],
        ["Quản trị viên\n(Admin)", "YC-AD01", "Quản lý thực đơn", "Thêm món mới, cập nhật đơn giá, loại món (Cung hoàng đạo), chuyển trạng thái kinh doanh."],
        ["Quản trị viên", "YC-AD02", "Quản lý hồ sơ nhân viên", "Thêm nhân viên mới, sửa thông tin, phân quyền chức vụ, cập nhật trạng thái làm việc."],
        ["Quản trị viên", "YC-AD03", "Báo cáo doanh thu & Biểu đồ", "Xem doanh thu hôm nay, doanh thu tháng, vẽ biểu đồ tăng trưởng doanh số, xuất file Excel."]
    ]
    add_styled_table(req_headers, req_rows, col_widths=[1.4, 0.9, 1.8, 2.5])

    add_h3("1.3.2. Yêu cầu phi chức năng (Non-Functional Requirements)")
    add_bullet("Phần mềm phải có tốc độ phản hồi các thao tác truy vấn dữ liệu dưới 1 giây. Việc cập nhật trạng thái bàn và thêm món ăn phải được thực thi tức thời.", "Hiệu năng và tốc độ xử lý: ");
    add_bullet("Hệ thống chỉ cho phép đăng nhập sau khi tài khoản đã được xác thực OTP qua Email. Phân quyền chặt chẽ các vai trò (Role-based Access Control), không để nhân viên can thiệp vào báo cáo tài chính của quản trị viên.", "Tính bảo mật: ");
    add_bullet("Dữ liệu về hóa đơn, khách hàng, nguyên vật liệu và tồn kho được lưu trữ tập trung trên cơ sở dữ liệu quan hệ Oracle, đảm bảo tính toàn vẹn (ACID), không xảy ra tình trạng mất mát khi phần mềm gặp sự cố đột ngột.", "Độ tin cậy và toàn vẹn dữ liệu: ");
    add_bullet("Giao diện xây dựng theo phong cách hiện đại, phối màu hài hòa, có hình ảnh minh họa sống động, các nút bấm thao tác rõ ràng, có thông báo hướng dẫn người dùng trực quan.", "Tính khả dụng (Usability): ");

    add_h2("1.4. Công nghệ và môi trường phát triển")
    add_h3("1.4.1. Ngôn ngữ lập trình Java và Java Swing Framework")
    add_p("Dự án được xây dựng hoàn toàn bằng ngôn ngữ lập trình Java (phiên bản Java SE 17 LTS). Java là ngôn ngữ lập trình hướng đối tượng mạnh mẽ, có tính độc lập nền tảng cao nhờ cơ chế 'Write Once, Run Anywhere' (Viết một lần, chạy mọi nơi).")
    add_p("Giao diện đồ họa (GUI) được phát triển bằng thư viện Java Swing. Swing cung cấp bộ linh kiện phong phú (JFrame, JPanel, JButton, JTable, JDialog...) cùng khả năng tùy biến cao (Custom Painting với Graphics2D), giúp xây dựng các thành phần giao diện uốn lượn góc bo tròn, đổ bóng, chuyển màu mượt mà theo phong cách Flat Design và hiệu ứng chuyển động sống động.")

    add_h3("1.4.2. Hệ quản trị cơ sở dữ liệu Oracle Database và JDBC")
    add_p("Hệ quản trị CSDL quan hệ Oracle Database (phiên bản 19c/21c) được chọn làm nền tảng lưu trữ dữ liệu trung tâm của nhà hàng. Oracle nổi tiếng toàn cầu về độ tin cậy, khả năng quản trị khối lượng dữ liệu lớn, tính bảo mật cao và tốc độ thực thi giao dịch vượt trội.")
    add_p("Để kết nối ứng dụng Java với Oracle, dự án sử dụng thư viện điều khiển chuẩn JDBC 'ojdbc8.jar' (Oracle JDBC Driver). Mọi câu lệnh truy vấn dữ liệu đều được thực hiện thông qua PreparedStatement nhằm tối ưu hóa bộ nhớ đệm (Cache) của máy chủ CSDL và ngăn chặn hoàn toàn nguy cơ tấn công SQL Injection.")

    add_h3("1.4.3. Mô hình kiến trúc MVC (Model - View - Controller)")
    add_p("Kiến trúc phần mềm được thiết kế tuân thủ nghiêm ngặt theo mô hình MVC:")
    add_bullet("Chứa các lớp thực thể đại diện cho các bảng dữ liệu trong CSDL Oracle (ModelMonAn, ModelHoaDon, ModelNhanVien, ModelKhachHang...). Các lớp này chỉ đóng vai trò đóng gói dữ liệu (Data Transfer Object) với các trường thuộc tính và hàm Getter/Setter.", "Tầng Model (Dữ liệu): ");
    add_bullet("Bao gồm toàn bộ các Form hiển thị, Panel, Dialog và các thành phần giao diện Java Swing trực tiếp tương tác với người dùng.", "Tầng View (Giao diện): ");
    add_bullet("Đóng vai trò là cầu nối xử lý logic nghiệp vụ và điều phối dữ liệu giữa View và Model. Tầng này được tổ chức thành các lớp Service chuyên trách (ServiceUser, ServiceCustomer, ServiceStaff, ServiceAdmin, ServiceMail) và lớp DatabaseConnection quản lý kết nối CSDL theo mẫu thiết kế Singleton.", "Tầng Controller / Service (Điều khiển): ");

    add_h3("1.4.4. Các thư viện bổ trợ khác")
    add_bullet("Cung cấp API gửi thư điện tử thông qua máy chủ SMTP của Google (smtp.gmail.com:587) với cơ chế bảo mật TLS, phục vụ việc gửi mã xác minh OTP 6 chữ số khi khách hàng đăng ký tài khoản.", "JavaMail API (javax.mail): ");
    add_bullet("Hỗ trợ xuất hóa đơn thanh toán và các báo cáo thống kê doanh thu ra tệp tin bảng tính Excel (.xlsx) chuẩn.", "Apache POI: ");
    add_bullet("Thư viện hỗ trợ vẽ biểu đồ doanh thu dạng đường cong và cột đa chiều, nâng cao năng lực phân tích cho nhà quản lý.", "Thư viện biểu đồ (Chart): ");
    doc.add_page_break()

    # ==================== CHƯƠNG 2 ====================
    add_title("CHƯƠNG 2: PHÂN TÍCH VÀ THIẾT KẾ HỆ THỐNG", size=16, bold=True, color=(27, 54, 93), space_after=14)

    add_h2("2.1. Phân tích chức năng và Sơ đồ Use Case")
    add_h3("2.1.1. Sơ đồ Use Case tổng quát")
    add_p("Hệ thống Quản lý nhà hàng Royal The Dreamers được phân rã thành các gói chức năng tương ứng với 4 tác nhân chính (Actors): Khách hàng (Customer), Nhân viên phục vụ/Tiếp tân (Staff), Nhân viên quản lý kho (Warehouse Staff) và Quản trị viên (Admin).")

    actor_headers = ["Tên Tác nhân (Actor)", "Mô tả vai trò trong hệ thống", "Quyền hạn truy cập chính"]
    actor_rows = [
        ["Khách hàng (Customer)", "Là thực khách đến trải nghiệm ẩm thực tại nhà hàng, sử dụng hệ thống để tự phục vụ hoặc xem thông tin cá nhân.", "Đăng ký tài khoản, đăng nhập, đặt món tại bàn, đổi điểm lấy Voucher, xem lịch sử hóa đơn, cập nhật thông tin cá nhân."],
        ["Nhân viên phục vụ (Staff)", "Bao gồm nhân viên tiếp tân đón khách, nhân viên bàn và thu ngân trực tiếp phục vụ tại sảnh nhà hàng.", "Quản lý trạng thái bàn ăn, lập order gọi món cho khách, thanh toán hóa đơn, tra cứu thông tin tích lũy của khách hàng."],
        ["Nhân viên Kho (Warehouse)", "Là thủ kho chịu trách nhiệm quản lý dòng nguyên vật liệu đầu vào và đầu ra của bếp.", "Quản lý danh mục nguyên liệu, kiểm tra số lượng tồn kho, lập phiếu nhập kho nguyên liệu, lập phiếu xuất kho nguyên liệu."],
        ["Quản trị viên (Admin)", "Là chủ nhà hàng hoặc người quản lý cấp cao điều hành toàn bộ hoạt động kinh doanh.", "Quản trị danh mục thực đơn món ăn, quản trị hồ sơ nhân sự, xem báo cáo doanh thu tài chính và biểu đồ tăng trưởng."]
    ]
    add_styled_table(actor_headers, actor_rows, col_widths=[1.8, 2.5, 2.3])

    add_p("Dưới đây là mô tả trực quan các nhóm Use Case chính của từng tác nhân:")
    add_bullet("Đăng ký tài khoản -> Nhận mã OTP Email -> Xác minh tài khoản -> Đăng nhập hệ thống -> Xem danh mục món ăn 12 Cung hoàng đạo -> Đặt món ăn -> Đổi điểm tích lũy lấy Voucher giảm giá -> Xem lịch sử hóa đơn ăn uống -> Quản lý thông tin cá nhân.", "Nhóm Use Case Khách hàng: ");
    add_bullet("Đăng nhập -> Xem sơ đồ bàn ăn -> Đổi trạng thái bàn (Còn trống/Đang dùng bữa/Đã đặt trước) -> Lập phiếu gọi món cho bàn -> Áp dụng mã Voucher giảm giá -> Thực hiện thanh toán hóa đơn -> Tra cứu điểm tích lũy khách hàng.", "Nhóm Use Case Nhân viên phục vụ: ");
    add_bullet("Đăng nhập -> Quản lý danh mục nguyên liệu -> Lập phiếu nhập kho nguyên liệu mới -> Lập phiếu xuất kho cho bếp -> Kiểm tra số lượng tồn kho.", "Nhóm Use Case Nhân viên kho: ");
    add_bullet("Đăng nhập quyền Quản lý -> Quản lý thực đơn (Thêm món, Sửa món, Ngừng kinh doanh) -> Quản lý nhân sự (Thêm nhân viên, Sửa thông tin, Sa thải/Nghỉ việc, Phân quyền) -> Xem báo cáo doanh thu ngày/tháng/năm -> Xem biểu đồ tăng trưởng.", "Nhóm Use Case Quản trị viên: ");

    add_h3("2.1.2. Đặc tả các Use Case trọng tâm của hệ thống")

    # UC01
    add_p("Bảng 2.2: Đặc tả Use Case Đăng ký tài khoản và Xác thực OTP", bold_prefix="Đặc tả Use Case 1: ", indent=False)
    uc01_headers = ["Thuộc tính Use Case", "Nội dung đặc tả chi tiết"]
    uc01_rows = [
        ["Tên Use Case", "Đăng ký tài khoản người dùng mới (Register & Verify Account)"],
        ["Tác nhân (Actor)", "Khách hàng (Customer)"],
        ["Mục đích", "Tạo tài khoản mới trong hệ thống và kích hoạt tài khoản thông qua mã xác minh OTP gửi về hộp thư điện tử."],
        ["Điều kiện tiên quyết", "Khách hàng chưa có tài khoản và có một địa chỉ Email thực tế đang hoạt động."],
        ["Luồng sự kiện chính", "1. Khách hàng truy cập màn hình Đăng ký, nhập Tên người dùng, Email và Mật khẩu.\n2. Hệ thống kiểm tra: Email không được để trống, đúng định dạng và chưa tồn tại trong bảng NguoiDung.\n3. Hệ thống tự động sinh một mã OTP ngẫu nhiên gồm 6 chữ số.\n4. Hệ thống gọi ServiceMail gửi mã OTP đến địa chỉ Email của khách hàng qua máy chủ SMTP Gmail.\n5. Lưu thông tin người dùng vào bảng NguoiDung với trạng thái VerifyCode và Trangthai=''.\n6. Hiển thị hộp thoại (Dialog) yêu cầu người dùng nhập mã xác minh 6 số.\n7. Khách hàng kiểm tra email, nhập đúng mã xác minh và nhấn Xác nhận.\n8. Hệ thống kiểm tra mã khớp, cập nhật Trangthai='Verified' và thông báo đăng ký thành công."],
        ["Luồng ngoại lệ", "- Nếu Email đã tồn tại: Hệ thống báo lỗi 'Email đã được đăng ký'.\n- Nếu mã OTP nhập không khớp: Hệ thống báo lỗi 'Mã xác minh không chính xác' và cho phép thử lại."]
    ]
    add_styled_table(uc01_headers, uc01_rows, col_widths=[2.0, 4.6])

    # UC02
    add_p("Bảng 2.3: Đặc tả Use Case Khách hàng đặt món tại bàn", bold_prefix="Đặc tả Use Case 2: ", indent=False)
    uc02_headers = ["Thuộc tính Use Case", "Nội dung đặc tả chi tiết"]
    uc02_rows = [
        ["Tên Use Case", "Đặt món ăn tại bàn (Order Food)"],
        ["Tác nhân (Actor)", "Khách hàng (Customer)"],
        ["Mục đích", "Cho phép thực khách tự xem thực đơn theo 12 Cung hoàng đạo và chọn món ăn đưa vào đơn đặt."],
        ["Điều kiện tiên quyết", "Khách hàng đã đăng nhập thành công vào phân hệ Main_Customer_Frame."],
        ["Luồng sự kiện chính", "1. Khách hàng chuyển đến tab 'Đặt món'.\n2. Hệ thống tải danh mục các món ăn tương ứng với loại cung hoàng đạo được chọn (mặc định hoặc lọc).\n3. Khách hàng nhấn vào thẻ món ăn để xem chi tiết tên, hình ảnh và đơn giá.\n4. Khách hàng chọn số lượng món và nhấn nút 'Thêm vào giỏ hàng'.\n5. Danh sách các món ăn đã chọn hiển thị bên bảng giỏ hàng kèm tổng số tiền tạm tính.\n6. Khách hàng kiểm tra lại danh sách, nhấn nút 'Xác nhận đặt món'.\n7. Hệ thống tạo mới hoặc cập nhật hóa đơn của bàn vào cơ sở dữ liệu."],
        ["Luồng ngoại lệ", "Nếu khách hàng chưa chọn số lượng hoặc danh sách món trống: Hệ thống hiển thị cảnh báo yêu cầu chọn ít nhất một món ăn."]
    ]
    add_styled_table(uc02_headers, uc02_rows, col_widths=[2.0, 4.6])

    # UC03
    add_p("Bảng 2.4: Đặc tả Use Case Nhân viên thanh toán hóa đơn và áp dụng Voucher", bold_prefix="Đặc tả Use Case 3: ", indent=False)
    uc03_headers = ["Thuộc tính Use Case", "Nội dung đặc tả chi tiết"]
    uc03_rows = [
        ["Tên Use Case", "Thanh toán hóa đơn và Áp dụng Voucher (Payment Bill)"],
        ["Tác nhân (Actor)", "Nhân viên phục vụ / Thu ngân (Staff)"],
        ["Mục đích", "Kiểm tra tổng tiền món ăn, giảm trừ tiền theo Voucher ưu đãi và xác nhận hoàn tất thanh toán hóa đơn."],
        ["Điều kiện tiên quyết", "Bàn ăn đang ở trạng thái 'Đang dùng bữa' và đã có hóa đơn mở."],
        ["Luồng sự kiện chính", "1. Nhân viên chọn bàn cần thanh toán trong danh sách bàn đang phục vụ.\n2. Hệ thống tải toàn bộ chi tiết hóa đơn (CTHD) gồm danh sách món ăn, đơn giá, số lượng và tổng tiền món.\n3. Nếu khách hàng có mã Voucher, nhân viên nhập mã vào ô Voucher. Hệ thống kiểm tra tính hợp lệ và phần trăm giảm giá của Voucher.\n4. Hệ thống tính toán lại Tiền giảm và Tổng tiền thực thu = Tiền món ăn - Tiền giảm.\n5. Nhân viên nhận tiền từ khách và nhấn 'Thanh toán'.\n6. Hệ thống cập nhật bảng HoaDon: Trangthai='Da thanh toan', cộng điểm tích lũy cho khách hàng tương ứng với giá trị hóa đơn, chuyển trạng thái bàn về 'Còn trống'.\n7. Hệ thống hiển thị tùy chọn in hóa đơn thanh toán cho khách."],
        ["Luồng ngoại lệ", "Nếu mã Voucher không tồn tại hoặc đã hết số lượng: Hệ thống hiển thị cảnh báo 'Mã giảm giá không hợp lệ hoặc đã hết lượt sử dụng'."]
    ]
    add_styled_table(uc03_headers, uc03_rows, col_widths=[2.0, 4.6])

    add_h2("2.2. Phân tích và thiết kế Cơ sở dữ liệu (Database Design)")
    add_h3("2.2.1. Mô hình thực thể liên kết (ERD)")
    add_p("Cơ sở dữ liệu của Hệ thống Quản lý nhà hàng Royal The Dreamers được xây dựng trên hệ quản trị cơ sở dữ liệu quan hệ Oracle Database. Hệ thống bao gồm 12 bảng thực thể chính, được liên kết chặt chẽ qua các mối quan hệ 1-N và N-N với các khóa chính (PK) và khóa ngoại (FK) bảo đảm tính toàn vẹn dữ liệu.")
    add_p("Dưới đây là Sơ đồ Thực thể Liên kết (ERD) chi tiết của toàn bộ hệ thống:")

    # Chèn ảnh ERD từ src/Icons/ERD.png
    add_img("src/Icons/ERD.png", "Hình 2.2: Sơ đồ Thực thể Liên kết (ERD) Hệ thống Quản lý Nhà hàng Royal The Dreamers", width_in=5.8)

    add_h3("2.2.2. Đặc tả chi tiết các bảng trong cơ sở dữ liệu Oracle")
    add_p("Dưới đây là bảng đặc tả cấu trúc chi tiết của toàn bộ các bảng trong cơ sở dữ liệu:")

    # Bảng 1: NguoiDung
    add_p("Bảng 2.5: Cấu trúc Bảng NguoiDung (Lưu trữ thông tin tài khoản đăng nhập)", bold_prefix="1. Bảng NguoiDung: ", indent=False)
    nd_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    nd_rows = [
        ["ID_ND", "NUMBER(8,0)", "PK", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của người dùng"],
        ["Email", "VARCHAR2(50)", "", "NOT NULL, CHECK (is not null)", "Địa chỉ hòm thư điện tử dùng để đăng nhập"],
        ["Matkhau", "VARCHAR2(20)", "", "NOT NULL", "Mật khẩu đăng nhập hệ thống"],
        ["VerifyCode", "VARCHAR2(10)", "", "DEFAULT NULL", "Mã OTP 6 số xác minh tài khoản qua Email"],
        ["Trangthai", "VARCHAR2(10)", "", "DEFAULT ''", "Trạng thái xác minh ('Verified' hoặc '')"],
        ["Vaitro", "VARCHAR2(20)", "", "CHECK ('Khach Hang','Nhan Vien','Nhan Vien Kho','Quan Ly')", "Vai trò người dùng trong hệ thống"]
    ]
    add_styled_table(nd_headers, nd_rows, col_widths=[1.2, 1.4, 0.6, 2.0, 1.4])

    # Bảng 2: NhanVien
    add_p("Bảng 2.6: Cấu trúc Bảng NhanVien (Lưu trữ thông tin nhân sự)", bold_prefix="2. Bảng NhanVien: ", indent=False)
    nv_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    nv_rows = [
        ["ID_NV", "NUMBER(8,0)", "PK", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của nhân viên"],
        ["TenNV", "VARCHAR2(50)", "", "NOT NULL", "Họ và tên đầy đủ của nhân viên"],
        ["NgayVL", "DATE", "", "NOT NULL", "Ngày nhân viên chính thức vào làm việc"],
        ["SDT", "VARCHAR2(50)", "", "NOT NULL", "Số điện thoại liên lạc của nhân viên"],
        ["Chucvu", "VARCHAR2(50)", "", "CHECK ('Phuc vu','Tiep tan','Thu ngan','Bep','Kho','Quan ly')", "Chức vụ đảm nhiệm tại nhà hàng"],
        ["ID_ND", "NUMBER(8,0)", "FK", "REFERENCES NguoiDung(ID_ND)", "Mã tài khoản đăng nhập tương ứng"],
        ["ID_NQL", "NUMBER(8,0)", "FK", "REFERENCES NhanVien(ID_NV)", "Mã của người quản lý trực tiếp"],
        ["Tinhtrang", "VARCHAR2(20)", "", "CHECK ('Dang lam viec','Da nghi viec')", "Tình trạng công tác hiện tại"]
    ]
    add_styled_table(nv_headers, nv_rows, col_widths=[1.1, 1.4, 0.6, 2.1, 1.4])

    # Bảng 3: KhachHang
    add_p("Bảng 2.7: Cấu trúc Bảng KhachHang (Lưu trữ thông tin thành viên thực khách)", bold_prefix="3. Bảng KhachHang: ", indent=False)
    kh_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    kh_rows = [
        ["ID_KH", "NUMBER(8,0)", "PK", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của khách hàng"],
        ["TenKH", "VARCHAR2(50)", "", "NOT NULL", "Họ và tên đầy đủ của khách hàng"],
        ["Ngaythamgia", "DATE", "", "NOT NULL", "Ngày khách hàng đăng ký thành viên"],
        ["Doanhso", "NUMBER(10,0)", "", "DEFAULT 0, NOT NULL", "Tổng số tiền khách hàng đã chi tiêu"],
        ["Diemtichluy", "NUMBER(5,0)", "", "DEFAULT 0, NOT NULL", "Số điểm thưởng tích lũy hiện có để đổi voucher"],
        ["ID_ND", "NUMBER(8,0)", "FK", "REFERENCES NguoiDung(ID_ND), NOT NULL", "Mã tài khoản đăng nhập của khách"]
    ]
    add_styled_table(kh_headers, kh_rows, col_widths=[1.1, 1.4, 0.6, 2.1, 1.4])

    # Bảng 4: MonAn
    add_p("Bảng 2.8: Cấu trúc Bảng MonAn (Danh mục các món ăn trong thực đơn)", bold_prefix="4. Bảng MonAn: ", indent=False)
    ma_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    ma_rows = [
        ["ID_MonAn", "NUMBER(8,0)", "PK", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của món ăn"],
        ["TenMon", "VARCHAR2(50)", "", "NOT NULL", "Tên gọi của món ăn trong thực đơn"],
        ["DonGia", "NUMBER(8,0)", "", "NOT NULL", "Đơn giá niêm yết của món ăn (VNĐ)"],
        ["Loai", "VARCHAR2(50)", "", "CHECK (Loai in 12 Cung hoàng đạo)", "Phân loại món ăn theo 12 Cung hoàng đạo"],
        ["TrangThai", "VARCHAR2(30)", "", "CHECK ('Dang kinh doanh','Ngung kinh doanh')", "Trạng thái sẵn sàng phục vụ của món"]
    ]
    add_styled_table(ma_headers, ma_rows, col_widths=[1.2, 1.4, 0.6, 2.0, 1.4])

    # Bảng 5: Ban
    add_p("Bảng 2.9: Cấu trúc Bảng Ban (Quản lý các bàn ăn trong sảnh nhà hàng)", bold_prefix="5. Bảng Ban: ", indent=False)
    ban_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    ban_rows = [
        ["ID_Ban", "NUMBER(8,0)", "PK", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của bàn ăn"],
        ["TenBan", "VARCHAR2(50)", "", "NOT NULL", "Tên hoặc số hiệu của bàn (ví dụ: Bàn 01, Bàn 02)"],
        ["Vitri", "VARCHAR2(50)", "", "NOT NULL", "Vị trí đặt bàn (Khu vực Tầng 1, Tầng 2, Sân vườn...)"],
        ["Trangthai", "VARCHAR2(50)", "", "CHECK ('Con trong','Dang dung bua','Da dat truoc')", "Trạng thái hiện tại của bàn ăn"]
    ]
    add_styled_table(ban_headers, ban_rows, col_widths=[1.2, 1.4, 0.6, 2.0, 1.4])

    # Bảng 6: Voucher
    add_p("Bảng 2.10: Cấu trúc Bảng Voucher (Mã giảm giá ưu đãi)", bold_prefix="6. Bảng Voucher: ", indent=False)
    vc_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    vc_rows = [
        ["Code_Voucher", "VARCHAR2(10)", "PK", "PRIMARY KEY, NOT NULL", "Mã code của voucher (ví dụ: VOUCHER10, VIP20)"],
        ["Mota", "VARCHAR2(50)", "", "NOT NULL", "Mô tả nội dung chương trình giảm giá"],
        ["Phantram", "NUMBER(3,0)", "", "CHECK (Phantram > 0 AND Phantram <= 100)", "Tỷ lệ phần trăm giảm giá (%)"],
        ["LoaiMA", "VARCHAR2(50)", "", "CHECK (LoaiMA in 'All' hoặc 12 Cung)", "Loại món ăn được phép áp dụng giảm giá"],
        ["SoLuong", "NUMBER(3,0)", "", "DEFAULT 0", "Số lượng mã voucher phát hành còn lại"],
        ["Diem", "NUMBER(8,0)", "", "DEFAULT 0", "Số điểm tích lũy cần có để đổi lấy voucher này"]
    ]
    add_styled_table(vc_headers, vc_rows, col_widths=[1.3, 1.3, 0.6, 2.0, 1.4])

    # Bảng 7: HoaDon & CTHD
    add_p("Bảng 2.11: Cấu trúc Bảng HoaDon (Hóa đơn dùng bữa của khách)", bold_prefix="7. Bảng HoaDon: ", indent=False)
    hd_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    hd_rows = [
        ["ID_HoaDon", "NUMBER(8,0)", "PK", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của hóa đơn"],
        ["ID_KH", "NUMBER(8,0)", "FK", "REFERENCES KhachHang(ID_KH)", "Mã khách hàng sở hữu hóa đơn"],
        ["ID_Ban", "NUMBER(8,0)", "FK", "REFERENCES Ban(ID_Ban)", "Mã bàn ăn gắn với hóa đơn"],
        ["NgayHD", "DATE", "", "NOT NULL", "Ngày giờ lập hóa đơn"],
        ["TienMonAn", "NUMBER(8,0)", "", "DEFAULT 0", "Tổng số tiền món ăn trước khi giảm giá"],
        ["Code_Voucher", "VARCHAR2(10)", "", "DEFAULT NULL", "Mã voucher giảm giá được áp dụng"],
        ["TienGiam", "NUMBER(8,0)", "", "DEFAULT 0", "Số tiền được giảm giá theo voucher"],
        ["Tongtien", "NUMBER(10,0)", "", "DEFAULT 0", "Tổng số tiền thực tế khách hàng phải trả"],
        ["Trangthai", "VARCHAR2(50)", "", "CHECK ('Chua thanh toan','Da thanh toan')", "Trạng thái thanh toán của hóa đơn"]
    ]
    add_styled_table(hd_headers, hd_rows, col_widths=[1.2, 1.4, 0.6, 2.0, 1.4])

    add_p("Bảng 2.12: Cấu trúc Bảng CTHD (Chi tiết từng món trong hóa đơn)", bold_prefix="8. Bảng CTHD: ", indent=False)
    cthd_headers = ["Tên trường", "Kiểu dữ liệu", "Khóa", "Ràng buộc (Constraints)", "Diễn giải ý nghĩa"]
    cthd_rows = [
        ["ID_HoaDon", "NUMBER(8,0)", "PK, FK", "REFERENCES HoaDon(ID_HoaDon)", "Mã hóa đơn"],
        ["ID_MonAn", "NUMBER(8,0)", "PK, FK", "REFERENCES MonAn(ID_MonAn)", "Mã món ăn được gọi"],
        ["SoLuong", "NUMBER(3,0)", "", "NOT NULL, CHECK (SoLuong > 0)", "Số lượng phần món ăn được đặt"],
        ["Thanhtien", "NUMBER(10,0)", "", "NOT NULL", "Thành tiền = SoLuong * DonGia"]
    ]
    add_styled_table(cthd_headers, cthd_rows, col_widths=[1.2, 1.4, 0.8, 2.0, 1.2])

    # Bảng 9, 10, 11, 12: Kho & Nguyên Liệu
    add_p("Bảng 2.13: Cấu trúc Bảng NguyenLieu và Bảng Kho (Quản lý tồn kho)", bold_prefix="9. Bảng NguyenLieu và Bảng Kho: ", indent=False)
    nl_headers = ["Bảng", "Tên trường", "Kiểu dữ liệu", "Khóa / Ràng buộc", "Diễn giải ý nghĩa"]
    nl_rows = [
        ["NguyenLieu", "ID_NL", "NUMBER(8,0)", "PK, NOT NULL", "Mã định danh duy nhất của nguyên liệu"],
        ["NguyenLieu", "TenNL", "VARCHAR2(50)", "NOT NULL", "Tên gọi của nguyên vật liệu thực phẩm"],
        ["NguyenLieu", "Dongia", "NUMBER(8,0)", "NOT NULL", "Đơn giá mua vào của nguyên liệu"],
        ["NguyenLieu", "Donvitinh", "VARCHAR2(50)", "CHECK ('g','kg','ml','l')", "Đơn vị đo lường nguyên liệu"],
        ["Kho", "ID_NL", "NUMBER(8,0)", "PK, FK REFERENCES NguyenLieu(ID_NL)", "Mã nguyên liệu trong kho"],
        ["Kho", "SLTon", "NUMBER(3,0)", "DEFAULT 0", "Số lượng tồn kho thực tế hiện tại"]
    ]
    add_styled_table(nl_headers, nl_rows, col_widths=[1.2, 1.2, 1.4, 1.6, 1.2])

    add_p("Bảng 2.14: Cấu trúc Bảng Phiếu Nhập Kho (PhieuNK, CTNK) và Phiếu Xuất Kho (PhieuXK, CTXK)", bold_prefix="10. Bảng Nhập / Xuất Kho: ", indent=False)
    nk_headers = ["Bảng dữ liệu", "Tên trường", "Kiểu dữ liệu", "Khóa / Ràng buộc", "Diễn giải ý nghĩa"]
    nk_rows = [
        ["PhieuNK", "ID_NK", "NUMBER(8,0)", "PK, NOT NULL", "Mã định danh duy nhất của phiếu nhập kho"],
        ["PhieuNK", "ID_NV", "NUMBER(8,0)", "FK REFERENCES NhanVien(ID_NV)", "Mã nhân viên lập phiếu nhập kho"],
        ["PhieuNK", "NgayNK", "DATE", "NOT NULL", "Ngày giờ thực hiện nhập kho nguyên liệu"],
        ["PhieuNK", "Tongtien", "NUMBER(10,0)", "DEFAULT 0", "Tổng số tiền nhập hàng trong phiếu"],
        ["CTNK", "ID_NK, ID_NL", "NUMBER(8,0)", "PK, FK liên kết PhieuNK và NguyenLieu", "Khóa chính tổ hợp chi tiết nhập kho"],
        ["CTNK", "SoLuong", "NUMBER(3,0)", "NOT NULL", "Số lượng nguyên liệu thực tế nhập vào"],
        ["CTNK", "Thanhtien", "NUMBER(10,0)", "NOT NULL", "Thành tiền nhập của nguyên liệu đó"],
        ["PhieuXK", "ID_XK", "NUMBER(8,0)", "PK, NOT NULL", "Mã định danh duy nhất của phiếu xuất kho"],
        ["PhieuXK", "ID_NV", "NUMBER(8,0)", "FK REFERENCES NhanVien(ID_NV)", "Mã nhân viên lập phiếu xuất kho"],
        ["PhieuXK", "NgayXK", "DATE", "NOT NULL", "Ngày giờ thực hiện xuất kho nguyên liệu"],
        ["CTXK", "ID_XK, ID_NL", "NUMBER(8,0)", "PK, FK liên kết PhieuXK và NguyenLieu", "Khóa chính tổ hợp chi tiết xuất kho"],
        ["CTXK", "SoLuong", "NUMBER(3,0)", "NOT NULL", "Số lượng nguyên liệu xuất cho bếp chế biến"]
    ]
    add_styled_table(nk_headers, nk_rows, col_widths=[1.2, 1.2, 1.4, 1.6, 1.2])

    add_h2("2.3. Thiết kế kiến trúc phần mềm và cấu trúc Class (Mô hình MVC)")
    add_p("Toàn bộ mã nguồn ứng dụng được phân chia một cách khoa học theo mẫu kiến trúc Model - View - Controller (MVC), nằm trong package cha `RTDRestaurant`. Cấu trúc cây thư mục mã nguồn bao gồm:")
    add_bullet("src/RTDRestaurant/Model: Chứa các lớp thực thể dữ liệu.", "");
    add_bullet("src/RTDRestaurant/Controller/Connection: Quản lý kết nối cơ sở dữ liệu.", "");
    add_bullet("src/RTDRestaurant/Controller/Service: Các lớp xử lý logic nghiệp vụ và truy vấn SQL.", "");
    add_bullet("src/RTDRestaurant/Controller/Event: Các Interface bắt sự kiện thao tác người dùng.", "");
    add_bullet("src/RTDRestaurant/View/Main_Frame: Các cửa sổ chính phân theo quyền.", "");
    add_bullet("src/RTDRestaurant/View/Form: Các Panel giao diện chức năng con.", "");
    add_bullet("src/RTDRestaurant/View/Dialog: Các hộp thoại tương tác (thông báo, xác nhận, OTP).", "");
    add_bullet("src/RTDRestaurant/View/Component: Các thành phần UI tùy biến (Card, Header, Menu).", "");
    add_bullet("src/RTDRestaurant/View/Swing: Các linh kiện đồ họa tùy biến cao.", "");

    add_h3("2.3.1. Thiết kế tầng Model (Entities Data Objects)")
    add_p("Tầng Model bao gồm 20 lớp đại diện cho các thực thể kinh doanh và các đối tượng truyền tải dữ liệu:")
    add_bullet("Đại diện cho tài khoản đăng nhập, chứa các trường: userID, email, password, role.", "ModelNguoiDung: ");
    add_bullet("Chứa thông tin món ăn, hình ảnh minh họa (ImageIcon), tên món, đơn giá và loại cung hoàng đạo.", "ModelMonAn: ");
    add_bullet("Quản lý thông tin hóa đơn, tổng tiền, tiền giảm, voucher và danh sách các món ăn trong hóa đơn.", "ModelHoaDon: ");
    add_bullet("Thông tin khách hàng thành viên, điểm tích lũy và doanh số chi tiêu.", "ModelKhachHang: ");
    add_bullet("Thông tin nhân sự, ngày vào làm, số điện thoại, chức vụ và trạng thái làm việc.", "ModelNhanVien: ");
    add_bullet("Thông tin bàn ăn, tên bàn, vị trí sảnh và trạng thái (Còn trống, Đang dùng bữa, Đã đặt trước).", "ModelBan: ");
    add_bullet("Thông tin mã voucher, phần trăm giảm giá, loại món ăn áp dụng và số điểm quy đổi.", "ModelVoucher: ");
    add_bullet("Quản lý nguyên vật liệu, số lượng tồn kho và các phiếu nhập/xuất kho.", "ModelNguyenLieu, ModelKho, ModelPNK, ModelPXK: ");

    add_h3("2.3.2. Thiết kế tầng Controller và Service")
    add_bullet("Triển khai theo mẫu thiết kế Singleton (Singleton Pattern). Chỉ khởi tạo một thể hiện kết nối duy nhất (Connection) tới CSDL Oracle thông qua URL `jdbc:oracle:thin:@localhost:1521:orcl` với Username 'Doan' và Password '123'. Giúp tối ưu hóa tài nguyên kết nối và quản lý tập trung.", "Lớp DatabaseConnection: ");
    add_bullet("Chuyên trách xác thực đăng nhập (login), kiểm tra tính hợp lệ của email/mật khẩu, đăng ký tài khoản khách hàng mới (insertUser) và kiểm tra mã xác thực OTP (verifyCodeWithUser).", "Lớp ServiceUser: ");
    add_bullet("Cung cấp các API phục vụ khách hàng: truy vấn thực đơn theo cung hoàng đạo (MenuFood), lọc món theo thứ tự giá/tên (MenuFoodOrder), lấy danh sách voucher hợp lệ (MenuVoucher), đổi điểm tích lũy lấy voucher (CashVoucher), lấy thông tin cá nhân và lịch sử hóa đơn.", "Lớp ServiceCustomer: ");
    add_bullet("Cung cấp các API cho nhân viên phục vụ: lấy sơ đồ danh sách bàn (MenuTable), đổi trạng thái bàn (SetTableStatus), tạo hóa đơn mới (InsertHoaDon), thêm chi tiết món ăn (InsertCTHD), thanh toán hóa đơn và in phiếu.", "Lớp ServiceStaff: ");
    add_bullet("Cung cấp các API cho nhà quản trị: thống kê doanh thu theo ngày/tháng/năm, tính toán số liệu biểu đồ doanh thu (getMonthlyRevenue), quản lý thêm/sửa món ăn thực đơn và quản lý hồ sơ nhân viên.", "Lớp ServiceAdmin: ");
    add_bullet("Cấu hình phiên làm việc JavaMail với giao thức SMTP bảo mật TLS qua cổng 587 của Google, tự động định dạng và gửi email chứa mã OTP xác minh tài khoản cho khách hàng.", "Lớp ServiceMail: ");

    add_h3("2.3.3. Thiết kế tầng View và Giao diện người dùng")
    add_bullet("Cửa sổ Đăng nhập & Đăng ký tích hợp hiệu ứng trượt mượt mà (Slide panel) giữa 2 form và hiển thị Dialog nhập mã OTP.", "Main_LoginAndRegister: ");
    add_bullet("Giao diện chính cho khách hàng với thanh điều hướng bên trái (Menu Bar) gồm: Về chúng tôi (About Us), Đặt món (Order Food), Đổi Voucher, Lịch sử Hóa đơn (Bill) và Thông tin cá nhân (Profile).", "Main_Customer_Frame: ");
    add_bullet("Giao diện cho nhân viên phục vụ với các tab: Quản lý bàn ăn (sơ đồ trực quan các bàn), Gọi món tại bàn, Thanh toán hóa đơn và Tra cứu thông tin khách hàng.", "Main_Staff_Frame: ");
    add_bullet("Giao diện cho nhân viên kho: Quản lý danh mục nguyên liệu, Quản lý kho tồn, Lập phiếu nhập kho và Lập phiếu xuất kho.", "Main_WarehouseStaff_Frame: ");
    add_bullet("Giao diện dành riêng cho Quản trị viên: Quản lý danh mục thực đơn món ăn, Chỉnh sửa món ăn, Quản lý hồ sơ nhân viên (Thêm/Sửa/Sa thải) và Báo cáo & Thống kê doanh thu trực quan.", "Main_Admin_Frame: ");

    add_h2("2.4. Thiết kế giao diện và trải nghiệm người dùng (UI/UX Concept)")
    add_p("Hệ thống được thiết kế hướng tới trải nghiệm người dùng tối ưu (User-Centric Design), lấy ý tưởng từ phong cách ẩm thực Hoàng gia sang trọng (Royal Concept). Bảng màu chủ đạo sử dụng sắc màu vàng kim kết hợp nền xanh đậm quý phái (#1B365D), tạo cảm giác ấm cúng, đẳng cấp và chuyên nghiệp.")
    add_p("Các bảng dữ liệu JTable đều được thiết kế lại với tiêu đề nổi bật, các dòng xen kẽ màu nền để dễ quét thông tin. Các nút bấm (JButton) đều có hiệu ứng hover đổi màu khi rê chuột. Các thẻ món ăn (Food Card) hiển thị hình ảnh món ăn sắc nét bo góc mềm mại, đi kèm nhãn giá rõ ràng, giúp thực khách dễ dàng đưa ra quyết định gọi món.")
    doc.add_page_break()

    # ==================== CHƯƠNG 3 ====================
    add_title("CHƯƠNG 3: CÀI ĐẶT THỬ NGHIỆM VÀ ĐÁNH GIÁ HỆ THỐNG", size=16, bold=True, color=(27, 54, 93), space_after=14)

    add_h2("3.1. Môi trường cài đặt và cấu hình hệ thống")
    add_p("Để triển khai và vận hành phần mềm Quản lý nhà hàng Royal The Dreamers, hệ thống đòi hỏi môi trường phần mềm và cấu hình như trong Bảng 3.1:")

    env_headers = ["Thành phần", "Thông số kỹ thuật / Phiên bản", "Ghi chú cấu hình"]
    env_rows = [
        ["Hệ điều hành", "Microsoft Windows 10 / Windows 11 (64-bit)", "Tương thích tốt với JVM và Swing UI"],
        ["Bộ công cụ Java (JDK)", "Oracle Java SE Development Kit 17 (JDK 17 LTS)", "Cấu hình JAVA_HOME trỏ đến thư mục cài đặt JDK"],
        ["Môi trường IDE", "Apache NetBeans IDE 17+ / IntelliJ IDEA / VS Code", "Dùng để biên dịch, đóng gói file JAR và chạy dự án"],
        ["Hệ quản trị CSDL", "Oracle Database 19c / 21c Enterprise/Express Edition", "Chạy trên cổng mặc định 1521, SID: 'orcl'"],
        ["Trình điều khiển CSDL", "ojdbc8.jar (Oracle JDBC Driver)", "Đã tích hợp trong thư mục src/External_Library"],
        ["Tài khoản CSDL", "Username: 'Doan' / Password: '123'", "Gán quyền DBA hoặc CONNECT, RESOURCE, CREATE VIEW"],
        ["Dịch vụ Email OTP", "Google SMTP Server (smtp.gmail.com), Cổng 587 TLS", "Sử dụng App Password của tài khoản thedreamersrestaurant@gmail.com"]
    ]
    add_styled_table(env_headers, env_rows, col_widths=[1.8, 2.5, 2.3])

    add_h3("3.1.1. Cấu hình cơ sở dữ liệu Oracle Database")
    add_p("Quy trình thiết lập cơ sở dữ liệu trên máy chủ Oracle:")
    add_p("Bước 1: Khởi động công cụ Oracle SQL*Plus hoặc Oracle SQL Developer, đăng nhập bằng tài khoản quản trị hệ thống SYS (SYSDBA).", indent=False)
    add_p("Bước 2: Tiến hành tạo User mới tên 'Doan' và thiết lập quyền hạn bằng các câu lệnh SQL:", indent=False)
    add_bullet("CREATE USER Doan IDENTIFIED BY 123 DEFAULT TABLESPACE users TEMPORARY TABLESPACE temp;", "");
    add_bullet("GRANT CONNECT, RESOURCE, DBA TO Doan;", "");
    add_bullet("GRANT UNLIMITED TABLESPACE TO Doan;", "");
    add_p("Bước 3: Đăng nhập vào tài khoản 'Doan', mở tệp tin mã nguồn `src/DataBase/Database.sql` và thực thi toàn bộ script. Kịch bản SQL sẽ tự động khởi tạo 12 bảng thực thể, toàn bộ các ràng buộc khóa chính, khóa ngoại, các ràng buộc Check Constraint, đồng thời nạp đầy đủ dữ liệu mẫu (món ăn 12 cung hoàng đạo, voucher, danh mục nguyên liệu, tài khoản người dùng ban đầu).", indent=False)

    add_h3("3.1.2. Cấu hình kết nối JDBC qua DatabaseConnection Singleton")
    add_p("Trong gói `RTDRestaurant.Controller.Connection`, lớp `DatabaseConnection.java` chịu trách nhiệm kết nối ứng dụng với CSDL Oracle thông qua chuỗi kết nối chuẩn:")
    add_bullet("Chuỗi kết nối URL: jdbc:oracle:thin:@localhost:1521:orcl", "");
    add_bullet("Username: Doan", "");
    add_bullet("Password: 123", "");
    add_p("Mẫu thiết kế Singleton đảm bảo chỉ tồn tại duy nhất một đối tượng kết nối `connection` trong suốt vòng đời của ứng dụng, tránh tình trạng rò rỉ bộ nhớ (memory leak) hoặc nghẽn kết nối cơ sở dữ liệu.")

    add_h3("3.1.3. Cấu hình dịch vụ gửi email xác minh OTP qua SMTP Gmail")
    add_p("Lớp `ServiceMail.java` được cấu hình để gửi mã xác thực tự động tới hộp thư của khách hàng khi đăng ký tài khoản:")
    add_bullet("Máy chủ SMTP: smtp.gmail.com", "");
    add_bullet("Cổng dịch vụ: 587 (Bật starttls.enable = true)", "");
    add_bullet("Tài khoản gửi: thedreamersrestaurant@gmail.com", "");
    add_bullet("Cơ chế xác thực: Mật khẩu ứng dụng (App Password) gồm 16 ký tự bảo mật, đảm bảo gửi mail ổn định, không bị chặn bởi cơ chế bảo mật 2 lớp của Google.", "");

    add_h2("3.2. Kết quả cài đặt các phân hệ và giao diện thực tế")
    add_p("Dưới đây là hình ảnh chụp màn hình và thuyết minh chi tiết kết quả cài đặt thực tế của tất cả các phân hệ chức năng trong phần mềm Quản lý nhà hàng Royal The Dreamers:")

    # 3.2.1. Xác thực
    add_h3("3.2.1. Phân hệ Xác thực tài khoản (Đăng nhập & Đăng ký xác thực OTP)")
    add_p("Phân hệ xác thực được đóng gói trong cửa sổ `Main_LoginAndRegister.java`. Điểm nhấn nổi bật của màn hình này là hiệu ứng đồ họa chuyển đổi mượt mà giữa form Đăng ký và form Đăng nhập mà không cần mở nhiều cửa sổ.")
    add_p("Khi khách hàng đăng ký tài khoản mới, hệ thống tự động sinh mã OTP 6 số ngẫu nhiên, gửi tới hộp thư của khách hàng và bật hộp thoại yêu cầu nhập mã xác minh. Khách hàng chỉ có thể đăng nhập sau khi hoàn tất xác minh mã này.")

    add_img("src/Demo/SignUp.png", "Hình 3.1: Giao diện Đăng ký tài khoản khách hàng mới", width_in=5.4)
    add_img("src/Demo/SignIn.png", "Hình 3.2: Giao diện Đăng nhập hệ thống phân quyền tự động", width_in=5.4)

    # 3.2.2. Khách hàng
    add_h3("3.2.2. Phân hệ dành cho Khách hàng (Customer Module)")
    add_p("Khách hàng sau khi đăng nhập thành công sẽ được điều hướng vào cửa sổ `Main_Customer_Frame.java`. Tại đây, khách hàng có thể khám phá không gian ẩm thực của nhà hàng:")
    add_bullet("Trang 'About Us' giới thiệu câu chuyện thương hiệu nhà hàng Royal The Dreamers và triết lý ẩm thực lấy cảm hứng từ 12 chòm sao hoàng đạo.", "");
    add_bullet("Trang 'Đặt món' cho phép duyệt thực đơn theo các cung hoàng đạo (Aries, Taurus, Gemini...), xem giá tiền, thêm món vào giỏ và đặt món trực tiếp tại bàn.", "");
    add_bullet("Trang 'Đổi Voucher' cho phép khách hàng theo dõi điểm tích lũy của mình và dùng điểm để đổi lấy các mã giảm giá 10%, 20%, 50%...", "");
    add_bullet("Trang 'Hóa đơn' cho phép thực khách tra cứu toàn bộ lịch sử các lần dùng bữa, chi tiết từng món ăn và tổng chi phí đã thanh toán.", "");
    add_bullet("Trang 'Thông tin cá nhân' hiển thị hồ sơ khách hàng, ngày tham gia thành viên, doanh số chi tiêu và điểm tích lũy hiện có.", "");

    add_img("src/Demo/Customer/AboutUs.png", "Hình 3.3: Giao diện Giới thiệu nhà hàng (About Us)", width_in=5.4)
    add_img("src/Demo/Customer/OrderFood.png", "Hình 3.4: Giao diện Khách hàng đặt món ăn theo 12 Cung hoàng đạo", width_in=5.4)
    add_img("src/Demo/Customer/voucher.png", "Hình 3.5: Giao diện Đổi điểm tích lũy lấy Voucher ưu đãi", width_in=5.4)
    add_img("src/Demo/Customer/Bill.png", "Hình 3.6: Giao diện Xem chi tiết hóa đơn thanh toán", width_in=5.4)
    add_img("src/Demo/Customer/Profile.png", "Hình 3.7: Giao diện Quản lý hồ sơ thông tin cá nhân khách hàng", width_in=5.4)

    # 3.2.3. Nhân viên
    add_h3("3.2.3. Phân hệ dành cho Nhân viên Phục vụ (Staff Module)")
    add_p("Nhân viên phục vụ và tiếp tân làm việc trên cửa sổ `Main_Staff_Frame.java` với các tính năng chuyên trách:")
    add_bullet("Quản lý sơ đồ bàn ăn: Hệ thống hiển thị trực quan toàn bộ danh sách bàn với các trạng thái màu sắc rõ ràng (Còn trống - Xanh lá; Đang dùng bữa - Đỏ; Đã đặt trước - Vàng). Nhân viên có thể nhấp chuột vào từng bàn để chuyển trạng thái hoặc xem hóa đơn tương ứng.", "");
    add_bullet("Gọi món cho khách: Nhân viên có thể tìm kiếm món ăn nhanh chóng, chọn số lượng phần ăn và ghi nhận đơn đặt món cho khách tại bàn.", "");
    add_bullet("Thanh toán hóa đơn: Giao diện thanh toán tự động tổng hợp danh sách món ăn của bàn, hỗ trợ nhập mã Voucher giảm giá, tính tiền thừa và in phiếu thanh toán cho khách hàng.", "");
    add_bullet("Tra cứu thông tin khách hàng: Nhân viên có thể tìm kiếm thông tin khách hàng thân thiết theo số điện thoại hoặc mã khách để áp dụng chế độ chăm sóc khách hàng đặc biệt.", "");

    add_img("src/Demo/Staff/Manage_Table.png", "Hình 3.8: Giao diện Quản lý trạng thái sơ đồ bàn ăn trực quan", width_in=5.4)
    add_img("src/Demo/Staff/Order_Food.png", "Hình 3.9: Giao diện Lập phiếu gọi món cho khách tại bàn", width_in=5.4)
    add_img("src/Demo/Staff/Payment.png", "Hình 3.10: Giao diện Thanh toán hóa đơn và in phiếu", width_in=5.4)
    add_img("src/Demo/Staff/Customer_Info.png", "Hình 3.11: Giao diện Tra cứu thông tin và điểm tích lũy khách hàng", width_in=5.4)

    # 3.2.4. Nhân viên kho
    add_h3("3.2.4. Phân hệ dành cho Nhân viên Kho (Warehouse Staff Module)")
    add_p("Nhân viên kho sử dụng cửa sổ `Main_WarehouseStaff_Frame.java` để theo dõi và điều phối nguồn nguyên vật liệu trong nhà hàng:")
    add_bullet("Theo dõi tồn kho: Hiển thị bảng tổng hợp số lượng tồn thực tế của tất cả các nguyên liệu (thịt bò, cá hồi, rau củ, gia vị...) kèm đơn vị đo lường tương ứng.", "");
    add_bullet("Nhập kho: Lập phiếu nhập kho khi tiếp nhận lô hàng mới, ghi nhận đơn giá nhập và tự động tăng số lượng tồn kho.", "");
    add_bullet("Xuất kho: Lập phiếu xuất kho cấp phát nguyên liệu cho khu vực bếp, tự động khấu trừ số lượng tồn kho sau khi xuất.", "");

    add_img("src/Demo/Staff/Manage_Ingredient.png", "Hình 3.12: Giao diện Quản lý nguyên vật liệu và kho hàng", width_in=5.4)

    # 3.2.5. Quản trị viên
    add_h3("3.2.5. Phân hệ dành cho Quản trị viên (Admin Module)")
    add_p("Quản trị viên nhà hàng truy cập vào cửa sổ `Main_Admin_Frame.java`, sở hữu toàn quyền quản trị cao nhất:")
    add_bullet("Quản lý thực đơn: Xem danh sách món ăn, thêm món ăn mới vào thực đơn, điều chỉnh đơn giá, thay đổi hình ảnh đại diện và chuyển đổi trạng thái 'Đang kinh doanh' hoặc 'Ngừng kinh doanh'.", "");
    add_bullet("Quản lý nhân sự: Xem danh sách nhân viên, thông tin chức vụ, ngày vào làm và số điện thoại. Hỗ trợ thêm mới nhân viên, cập nhật thông tin và xử lý cho thôi việc.", "");
    add_bullet("Báo cáo và Thống kê doanh thu: Bảng điều khiển (Dashboard) hiển thị tổng doanh thu hôm nay, tổng số hóa đơn đã thanh toán, biểu đồ đường biểu diễn doanh thu tăng trưởng theo từng tháng trong năm, hỗ trợ đắc lực cho việc hoạch định chiến lược kinh doanh.", "");

    add_img("src/Demo/Admin/Manage_Food.png", "Hình 3.13: Giao diện Quản lý danh mục món ăn dành cho Quản trị viên", width_in=5.4)
    add_img("src/Demo/Admin/Edit_Food.png", "Hình 3.14: Giao diện Chỉnh sửa và cập nhật thực đơn món ăn", width_in=5.4)
    add_img("src/Demo/Admin/Mange_HR.png", "Hình 3.15: Giao diện Quản lý hồ sơ nhân sự nhà hàng", width_in=5.4)
    add_img("src/Demo/Admin/Add_HR.png", "Hình 3.16: Giao diện Thêm mới nhân viên và phân quyền", width_in=5.4)
    add_img("src/Demo/Admin/Statistic.png", "Hình 3.17: Giao diện Báo cáo doanh thu và biểu đồ thống kê kinh doanh", width_in=5.4)

    add_h2("3.3. Kiểm thử phần mềm và đánh giá chất lượng (Testing)")
    add_p("Để đảm bảo phần mềm vận hành ổn định, chính xác và an toàn trước khi đưa vào khai thác, quá trình kiểm thử phần mềm được tiến hành bằng phương pháp Kiểm thử hộp đen (Black-box Testing) với các ca kiểm thử (Test Cases) bao quát toàn bộ các luồng nghiệp vụ chính.")
    add_p("Dưới đây là Bảng 3.2 tổng hợp kết quả thực thi các kịch bản kiểm thử:")

    tc_headers = ["Mã TC", "Chức năng kiểm thử", "Dữ liệu đầu vào (Input)", "Kết quả kỳ vọng (Expected Output)", "Thực tế", "Đánh giá"]
    tc_rows = [
        ["TC01", "Đăng ký tài khoản", "Nhập họ tên, email hợp lệ, mật khẩu", "Gửi mã OTP 6 số về email, mở popup nhập OTP", "Nhận được email OTP, hiển thị popup nhập mã", "Pass"],
        ["TC02", "Xác minh OTP", "Nhập đúng mã 6 số nhận từ email", "Tài khoản chuyển 'Verified', thông báo thành công", "Xác thực thành công, tài khoản kích hoạt", "Pass"],
        ["TC03", "Xác minh OTP sai", "Nhập mã ngẫu nhiên sai", "Báo lỗi 'Mã xác minh không chính xác'", "Báo lỗi chính xác, cho phép nhập lại", "Pass"],
        ["TC04", "Đăng nhập sai pass", "Email đúng, nhập mật khẩu sai", "Hiển thị thông báo đăng nhập thất bại", "Báo sai thông tin, không cho phép truy cập", "Pass"],
        ["TC05", "Đăng nhập đúng quyền", "Nhập tài khoản Quản lý / Nhân viên", "Điều hướng đúng Frame tương ứng theo quyền", "Điều hướng chính xác Main_Admin/Staff_Frame", "Pass"],
        ["TC06", "Khách đặt món", "Chọn món, chọn số lượng = 2, bấm đặt", "Thêm vào giỏ hàng, tổng tiền tính đúng", "Món hiển thị trong giỏ, tổng tiền = 2 * Đơn giá", "Pass"],
        ["TC07", "Đổi Voucher", "Đủ 50 điểm, bấm đổi voucher 10%", "Điểm bị trừ 50, voucher thêm vào tài khoản", "Điểm tích lũy trừ đúng, voucher lưu vào DB", "Pass"],
        ["TC08", "Đổi Voucher khi thiếu điểm", "Có 20 điểm, bấm đổi voucher 50 điểm", "Hiển thị cảnh báo không đủ điểm tích lũy", "Báo lỗi không đủ điểm, không trừ điểm", "Pass"],
        ["TC09", "Đổi trạng thái bàn", "Bấm vào bàn trống, chọn 'Đặt trước'", "Bàn đổi màu vàng, DB lưu 'Da dat truoc'", "Bàn đổi màu trực quan, DB cập nhật ngay", "Pass"],
        ["TC10", "Thanh toán & Áp Voucher", "Hóa đơn 500k, nhập voucher 20%", "Tiền giảm = 100k, Tổng tiền = 400k", "Tính toán tiền giảm chính xác, lưu hóa đơn", "Pass"],
        ["TC11", "Tích điểm khi thanh toán", "Thanh toán hóa đơn 1.000.000 VNĐ", "Cộng điểm thưởng vào tài khoản khách hàng", "Doanh số và điểm thưởng của khách tăng đúng", "Pass"],
        ["TC12", "Nhập kho nguyên liệu", "Nhập 10kg thịt bò vào kho", "Số lượng tồn kho cộng thêm 10kg", "Tồn kho trong bảng Kho tăng thêm 10 đơn vị", "Pass"],
        ["TC13", "Xuất kho nguyên liệu", "Xuất 5kg thịt bò cho bếp", "Số lượng tồn kho giảm đi 5kg", "Tồn kho trong bảng Kho giảm đúng 5 đơn vị", "Pass"],
        ["TC14", "Thêm món ăn mới", "Admin nhập tên món, giá, loại, ảnh", "Món mới xuất hiện trên thực đơn khách hàng", "Món lưu vào DB MonAn, hiển thị đúng ảnh", "Pass"],
        ["TC15", "Biểu đồ thống kê", "Admin mở tab Thống kê doanh thu", "Biểu đồ hiển thị doanh thu theo các tháng", "Biểu đồ vẽ mượt mà, số liệu khớp với DB", "Pass"]
    ]
    add_styled_table(tc_headers, tc_rows, col_widths=[0.7, 1.4, 1.4, 1.5, 1.1, 0.5])

    add_h2("3.4. Đánh giá ưu điểm và hạn chế của hệ thống")
    add_h3("3.4.1. Những ưu điểm nổi bật")
    add_bullet("Giao diện người dùng Java Swing được thiết kế tỉ mỉ, hiện đại, hình ảnh món ăn sắc nét, hiệu ứng hoạt họa mượt mà, mang đậm bản sắc sang trọng của nhà hàng Hoàng gia Royal The Dreamers.", "Giao diện thẩm mỹ cao và thân thiện: ");
    add_bullet("Bao quát đầy đủ tất cả các quy trình nghiệp vụ thực tế trong nhà hàng ẩm thực: tự phục vụ gọi món tại bàn, đổi điểm tích lũy, phân quyền nhân sự, quản lý xuất nhập tồn kho và thống kê tài chính.", "Quy trình nghiệp vụ khép kín, hoàn chỉnh: ");
    add_bullet("Tích hợp thành công dịch vụ gửi mã xác minh OTP qua hộp thư điện tử Gmail (JavaMail API), nâng cao tính bảo mật và xác thực chính chủ tài khoản khách hàng.", "Bảo mật và xác thực OTP tiên tiến: ");
    add_bullet("Cơ sở dữ liệu Oracle Database được thiết kế chuẩn mực (dạng chuẩn 3NF), toàn vẹn dữ liệu thông qua các ràng buộc khóa ngoại, Check constraint và trigger tự động.", "Cơ sở dữ liệu mạnh mẽ, toàn vẹn: ");
    add_bullet("Kiến trúc mã nguồn phân tầng MVC rõ ràng, tách bạch giữa giao diện hiển thị, lớp điều khiển nghiệp vụ và đối tượng dữ liệu, tạo thuận lợi tối đa cho việc bảo trì và nâng cấp mở rộng sau này.", "Kiến trúc mã nguồn chuẩn mực: ");

    add_h3("3.4.2. Một số hạn chế còn tồn tại")
    add_bullet("Hệ thống hiện tại mới được triển khai dưới dạng ứng dụng Desktop (Java Swing). Khách hàng cần sử dụng thiết bị máy tính bảng hoặc máy tính bàn đặt tại chỗ để gọi món, chưa có phiên bản ứng dụng di động (Mobile App) chạy trên Android hoặc iOS.", "Nền tảng triển khai: ");
    add_bullet("Chức năng thanh toán hiện tại hỗ trợ thanh toán tiền mặt và ghi nhận hóa đơn trên hệ thống, chưa tích hợp trực tiếp với các cổng thanh toán điện tử trực tuyến phổ biến (VNPay QR, MoMo, ZaloPay, thẻ tín dụng ngân hàng).", "Phương thức thanh toán trực tuyến: ");
    add_bullet("Khi mở rộng ra chuỗi nhiều chi nhánh nhà hàng đồng thời, hệ thống cần được nâng cấp kiến trúc lên mô hình Web API Microservices (Spring Boot) kết hợp đồng bộ dữ liệu đám mây (Cloud Database).", "Khả năng nhân rộng quy mô chuỗi: ");
    doc.add_page_break()

    # ==================== KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN ====================
    add_title("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", size=16, bold=True, color=(27, 54, 93), space_after=14)

    add_h2("1. Kết quả đạt được của đồ án")
    add_p("Sau một thời gian nghiêm túc học tập, nghiên cứu và triển khai với sự hướng dẫn tận tình của Thầy Mai Xuân Hùng, sinh viên đã hoàn thành đồ án môn học 'Xây dựng hệ thống quản lý nhà hàng Royal The Dreamers Restaurant' với các kết quả cụ thể như sau:")
    add_bullet("Đã nắm vững và vận dụng thành thạo các kiến thức nền tảng của ngôn ngữ lập trình Java, các kỹ thuật lập trình hướng đối tượng (OOP), lập trình giao diện đồ họa người dùng với Java Swing và kỹ thuật kết nối cơ sở dữ liệu JDBC.", "Về mặt kiến thức: ");
    add_bullet("Đã làm chủ quy trình phân tích và thiết kế hệ thống thông tin quản lý: từ khâu khảo sát hiện trạng nghiệp vụ thực tế, xây dựng sơ đồ Use Case, thiết kế cơ sở dữ liệu quan hệ ERD chuẩn mực, đến việc cài đặt kiến trúc phân tầng MVC trong mã nguồn dự án.", "Về mặt kỹ năng: ");
    add_bullet("Xây dựng thành công ứng dụng phần mềm Quản lý nhà hàng hoàn chỉnh với 4 phân hệ hoạt động trơn tru: Khách hàng (Đăng ký OTP, đặt món 12 cung hoàng đạo, đổi voucher tích lũy), Nhân viên phục vụ (Quản lý sơ đồ bàn, gọi món, thanh toán), Nhân viên kho (Quản lý nguyên liệu, nhập/xuất kho) và Quản trị viên (Quản lý thực đơn, nhân sự, báo cáo doanh thu trực quan).", "Về mặt sản phẩm: ");
    add_bullet("Toàn bộ các ca kiểm thử chức năng đều đạt kết quả tốt, hệ thống vận hành ổn định, không phát sinh lỗi nghiêm trọng và bảo toàn tính toàn vẹn của dữ liệu trong CSDL Oracle.", "Về mặt thử nghiệm: ");

    add_h2("2. Hướng phát triển và nâng cấp trong tương lai")
    add_p("Nhằm nâng cao hơn nữa giá trị thực tiễn và khả năng ứng dụng đại trà của sản phẩm, trong thời gian tới, hệ thống có thể được tiếp tục nghiên cứu và mở rộng theo các hướng sau:")
    add_bullet("Phát triển phiên bản ứng dụng di động (Mobile App) chạy trên hệ điều hành Android và iOS để khách hàng có thể cài đặt trực tiếp trên điện thoại cá nhân, quét mã QR tại bàn để xem thực đơn và gọi món nhanh chóng.", "1. Mở rộng nền tảng Mobile: ");
    add_bullet("Kết nối với các cổng thanh toán trực tuyến thông dụng như VNPay, MoMo, VietQR để khách hàng có thể quét mã thanh toán không dùng tiền mặt tức thời ngay tại bàn ăn.", "2. Tích hợp cổng thanh toán trực tuyến: ");
    add_bullet("Chuyển đổi tầng xử lý nghiệp vụ sang kiến trúc RESTful Web API sử dụng Spring Boot Framework, đưa cơ sở dữ liệu lên dịch vụ máy chủ đám mây (Cloud Database) để hỗ trợ quản lý chuỗi nhiều chi nhánh nhà hàng tập trung.", "3. Nâng cấp kiến trúc Cloud & Web API: ");
    add_bullet("Áp dụng các thuật toán Trí tuệ nhân tạo (AI/Machine Learning) để phân tích thói quen ăn uống của thực khách, từ đó đưa ra các gợi ý món ăn (Recommendation System) cá nhân hóa theo từng cung hoàng đạo và sở thích của khách hàng.", "4. Ứng dụng AI gợi ý món ăn: ");
    doc.add_page_break()

    # ==================== TÀI LIỆU THAM KHẢO ====================
    add_title("TÀI LIỆU THAM KHẢO", size=16, bold=True, color=(27, 54, 93), space_after=14)
    refs = [
        "[1] Mai Xuân Hùng, Giáo trình Lập trình Java cơ bản và nâng cao, Trường Đại học Kinh doanh và Công nghệ Hà Nội (HUBT).",
        "[2] Herbert Schildt (2022), Java: The Complete Reference, Twelfth Edition, McGraw-Hill Education.",
        "[3] Cay S. Horstmann (2021), Core Java Volume I--Fundamentals, 12th Edition, Prentice Hall.",
        "[4] Oracle Corporation, Java SE 17 Documentation & Java Swing API Specifications, https://docs.oracle.com/en/java/javase/17/.",
        "[5] Oracle Corporation, Oracle Database SQL Language Reference 19c / 21c, https://docs.oracle.com/en/database/oracle/oracle-database/.",
        "[6] Oracle Corporation, Oracle Database JDBC Developer's Guide (ojdbc8), https://docs.oracle.com/en/database/oracle/oracle-database/19/jjdbc/.",
        "[7] Jakarta Mail / JavaMail API Documentation, Eclipse Foundation, https://eclipse-ee4j.github.io/mail/.",
        "[8] Apache Software Foundation, Apache POI - the Java API for Microsoft Documents, https://poi.apache.org/.",
        "[9] DJ-Raven, Java Swing UI Design - Register and Verify Code With Email, https://github.com/DJ-Raven/java-swing-login-ui-001.",
        "[10] DJ-Raven, Java Swing UI Design - School Management Dashboard, https://github.com/DJ-Raven/java-swing-school-management-dashboard.",
        "[11] DJ-Raven, Java UI Design - Dashboard Desktop Application, https://github.com/DJ-Raven/java-ui-dashboard-008.",
        "[12] Gamma, E., Helm, R., Johnson, R., & Vlissides, J. (1994), Design Patterns: Elements of Reusable Object-Oriented Software (Mô hình Singleton, MVC), Addison-Wesley."
    ]
    for r in refs:
        add_p(r, indent=False)

    # Save document
    output_path = "Bao_Cao_Do_An_Java_Quan_Ly_Nha_Hang.docx"
    doc.save(output_path)
    print(f"Report successfully generated and saved to: {output_path}")

if __name__ == "__main__":
    create_report()
