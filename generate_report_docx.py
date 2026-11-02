import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import os
import win32com.client
import pdfplumber

def create_document():
    template_path = r'modul 2\Muchammad Rizal  Sobari- Modul 2 - SOLID\Muchammad Rizal  Sobari- Modul 2 - SOLID.docx'
    doc = docx.Document(template_path)

    # Remove all body children except the final sectPr
    body = doc._body._element
    for child in list(body):
        if not child.tag.endswith('sectPr'):
            body.remove(child)

    def set_margins(s):
        s.page_width = Pt(595.3)   # 21 cm
        s.page_height = Pt(841.9)  # 29.7 cm
        s.left_margin = Pt(113.4)  # 4 cm
        s.top_margin = Pt(85.05)   # 3 cm
        s.right_margin = Pt(85.05) # 3 cm
        s.bottom_margin = Pt(85.05)# 3 cm

    def set_page_number_format(s, fmt='decimal', start=None):
        sectPr = s._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is None:
            pgNumType = OxmlElement('w:pgNumType')
            sectPr.append(pgNumType)
        pgNumType.set(qn('w:fmt'), fmt)
        if start is not None:
            pgNumType.set(qn('w:start'), str(start))

    def add_page_number_to_footer(s):
        footer = s.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        
        fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
        instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
        fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
        fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run._r.append(fldChar3)

    def remove_numPr(p):
        if p._p.pPr is not None and p._p.pPr.numPr is not None:
            p._p.pPr.remove(p._p.pPr.numPr)

    def add_title_p(text, bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2):
        p = doc.add_paragraph()
        remove_numPr(p)
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.bold = bold
        run.font.size = Pt(size)
        return p

    def add_chapter_heading(text):
        p = doc.add_paragraph(text, style='Heading 1')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(8)
        return p

    def add_unnumbered_heading(text):
        p = doc.add_paragraph()
        remove_numPr(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.bold = True
        r.font.size = Pt(14)
        return p

    def add_heading_2(text):
        # Heading 2 style has automatic numbering I.1, II.1, II.2 etc.
        p = doc.add_paragraph(text, style='Heading 2')
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        return p

    def add_body_p(text, justify=True, bold_prefix=None):
        p = doc.add_paragraph()
        remove_numPr(p)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = 'Times New Roman'
            r_pre.font.bold = True
            r_pre.font.size = Pt(12)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        return p

    def add_subheading(text):
        p = doc.add_paragraph()
        remove_numPr(p)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.bold = True
        r.font.size = Pt(12)
        return p

    def add_code_block(code_text, filename_header=None):
        if filename_header:
            p_hdr = doc.add_paragraph()
            remove_numPr(p_hdr)
            p_hdr.paragraph_format.space_before = Pt(6)
            p_hdr.paragraph_format.space_after = Pt(2)
            r_hdr = p_hdr.add_run(filename_header)
            r_hdr.font.name = 'Times New Roman'
            r_hdr.font.bold = True
            r_hdr.font.size = Pt(11)
            r_hdr.font.italic = True
        
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        cell = table.cell(0, 0)
        cell.width = Pt(420)
        
        shading = parse_xml(r'<w:shd {} w:fill="F4F6F8"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading)
        
        tcBorders = parse_xml(r'''
            <w:tcBorders {} >
                <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
                <w:left w:val="single" w:sz="18" w:space="0" w:color="0969DA"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
            </w:tcBorders>
        '''.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(tcBorders)
        
        p_code = cell.paragraphs[0]
        p_code.style = 'CODE'
        remove_numPr(p_code)
        p_code.paragraph_format.space_before = Pt(4)
        p_code.paragraph_format.space_after = Pt(4)
        p_code.paragraph_format.line_spacing = 1.05
        r = p_code.add_run(code_text)
        r.font.name = 'Courier New'
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(25, 25, 25)
        
        p_spc = doc.add_paragraph()
        remove_numPr(p_spc)
        p_spc.paragraph_format.space_before = Pt(0)
        p_spc.paragraph_format.space_after = Pt(4)

    def add_figure(image_path, caption_text, width_inches=5.8):
        if os.path.exists(image_path):
            p_img = doc.add_paragraph()
            remove_numPr(p_img)
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.add_run().add_picture(image_path, width=Inches(width_inches))
            
            p_cap = doc.add_paragraph(style='gambar 2')
            remove_numPr(p_cap)
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            r_cap = p_cap.add_run(caption_text)
            r_cap.font.name = 'Times New Roman'
            r_cap.font.size = Pt(11)
            r_cap.font.italic = True
        else:
            print(f"Warning: Image {image_path} not found!")

    # ==================== SECTION 1: COVER PAGE ====================
    sec_cover = doc.sections[0]
    set_margins(sec_cover)
    sec_cover.footer.is_linked_to_previous = False

    add_title_p("LAPORAN PRAKTIKUM", bold=True, size=14, space_after=2)
    add_title_p("IMPLEMENTASI PERANGKAT LUNAK", bold=True, size=14, space_after=14)

    add_title_p("MODUL 3", bold=True, size=14, space_after=2)
    add_title_p("Design Pattern & Object Composing", bold=True, size=14, space_after=20)

    logo_path = 'assets/logo_unjani.png'
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        remove_numPr(p_logo)
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(4)
        p_logo.paragraph_format.space_after = Pt(20)
        p_logo.add_run().add_picture(logo_path, width=Inches(2.1))

    add_title_p("DISUSUN OLEH :", bold=True, size=12, space_after=8)

    tbl_author = doc.add_table(rows=1, cols=2)
    tbl_author.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_author.autofit = False
    cell_nim = tbl_author.cell(0, 0)
    cell_name = tbl_author.cell(0, 1)
    cell_nim.width = Pt(130)
    cell_name.width = Pt(220)

    p_nim = cell_nim.paragraphs[0]
    remove_numPr(p_nim)
    p_nim.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_nim = p_nim.add_run("2450081074")
    r_nim.font.name = 'Times New Roman'
    r_nim.font.bold = True
    r_nim.font.size = Pt(12)

    p_name = cell_name.paragraphs[0]
    remove_numPr(p_name)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run("Muchammad Rizal Sobari")
    r_name.font.name = 'Times New Roman'
    r_name.font.bold = True
    r_name.font.size = Pt(12)

    tblBorders = parse_xml(r'''
        <w:tblBorders {} >
            <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        </w:tblBorders>
    '''.format(nsdecls('w')))
    tbl_author._tbl.tblPr.append(tblBorders)

    p_space = doc.add_paragraph()
    remove_numPr(p_space)
    p_space.paragraph_format.space_before = Pt(8)
    p_space.paragraph_format.space_after = Pt(2)

    add_title_p("KELAS C", bold=True, size=12, space_after=32)

    add_title_p("PROGRAM STUDI TEKNIK INFORMATIKA", bold=True, size=14, space_after=2)
    add_title_p("FAKULTAS SAINS DAN INFORMATIKA", bold=True, size=14, space_after=2)
    add_title_p("UNIVERSITAS JENDERAL ACHMAD YANI", bold=True, size=14, space_after=2)
    add_title_p("2026", bold=True, size=14, space_after=0)

    # ==================== SECTION 2: DAFTAR ISI & DAFTAR GAMBAR (ROMAN) ====================
    sec_pre = doc.add_section(WD_SECTION.NEW_PAGE)
    set_margins(sec_pre)
    sec_pre.footer.is_linked_to_previous = False
    set_page_number_format(sec_pre, fmt='lowerRoman', start=1)
    add_page_number_to_footer(sec_pre)

    # DAFTAR ISI
    add_unnumbered_heading("DAFTAR ISI")

    p_hlm_toc = doc.add_paragraph()
    remove_numPr(p_hlm_toc)
    p_hlm_toc.paragraph_format.space_after = Pt(6)
    r_hlm_toc = p_hlm_toc.add_run("Halaman")
    r_hlm_toc.font.name = 'Times New Roman'
    r_hlm_toc.font.size = Pt(11)
    r_hlm_toc.font.bold = True
    p_hlm_toc.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    toc_items = [
        ("DAFTAR ISI", "i"),
        ("DAFTAR GAMBAR", "ii"),
        ("BAB I HASIL PRAKTIKUM (LATIHAN)", "1"),
        ("    I.1. Latihan 1 – Implementasi Pattern MVC pada Class Student", "1"),
        ("BAB II HASIL PRAKTIKUM (TUGAS)", "5"),
        ("    II.1. Tugas 1 – Melengkapi Implementasi Pattern MVC pada Class Course", "5"),
        ("    II.2. Tugas 2 – Refaktorisasi Kode Monolitik Kasir Menjadi Pattern MVC", "9"),
        ("BAB III STATUS DAN RIWAYAT REPOSITORY (GIT & GITHUB)", "15"),
        ("BAB IV KESIMPULAN", "17"),
    ]

    for title, page in toc_items:
        p_t = doc.add_paragraph()
        remove_numPr(p_t)
        p_t.paragraph_format.line_spacing = 1.15
        p_t.paragraph_format.space_after = Pt(3)
        p_t.paragraph_format.tab_stops.add_tab_stop(Pt(400), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT, docx.enum.text.WD_TAB_LEADER.DOTS)
        
        r_t = p_t.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(11)
        if "BAB" in title or "DAFTAR" in title:
            r_t.font.bold = True
        
        r_p = p_t.add_run(f"\t{page}")
        r_p.font.name = 'Times New Roman'
        r_p.font.size = Pt(11)

    doc.add_page_break()

    # DAFTAR GAMBAR
    add_unnumbered_heading("DAFTAR GAMBAR")

    p_hlm_fig = doc.add_paragraph()
    remove_numPr(p_hlm_fig)
    p_hlm_fig.paragraph_format.space_after = Pt(6)
    r_hlm_fig = p_hlm_fig.add_run("Halaman")
    r_hlm_fig.font.name = 'Times New Roman'
    r_hlm_fig.font.size = Pt(11)
    r_hlm_fig.font.bold = True
    p_hlm_fig.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    figures = [
        ("gambar I-1 Hasil program Latihan MVC Pattern (Student)", "3"),
        ("gambar I-2 Hasil commit Git Terminal Latihan MVC Pattern (Student)", "3"),
        ("gambar II-1 Hasil program Tugas 1 MVC Pattern (Course)", "7"),
        ("gambar II-2 Hasil commit Git Terminal Tugas 1 MVC Pattern (Course)", "8"),
        ("gambar II-3 Hasil program Tugas 2 Refaktorisasi Kasir MVC", "13"),
        ("gambar II-4 Hasil commit Git Terminal Tugas 2 Kasir MVC", "13"),
        ("gambar II-5 Status Repository dan Log Git Terminal Terhubung ke GitHub", "15"),
    ]

    for cap, page in figures:
        p_item = doc.add_paragraph(style='table of figures')
        remove_numPr(p_item)
        p_item.paragraph_format.line_spacing = 1.15
        p_item.paragraph_format.space_after = Pt(4)
        r1 = p_item.add_run(cap)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)
        
        p_item.paragraph_format.tab_stops.add_tab_stop(Pt(400), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT, docx.enum.text.WD_TAB_LEADER.DOTS)
        r2 = p_item.add_run(f"\t{page}")
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)

    # ==================== SECTION 3: ISI LAPORAN (ARABIC, RESTART AT 1) ====================
    sec_body = doc.add_section(WD_SECTION.NEW_PAGE)
    set_margins(sec_body)
    sec_body.footer.is_linked_to_previous = False
    set_page_number_format(sec_body, fmt='decimal', start=1)
    add_page_number_to_footer(sec_body)

    # BAB I HASIL PRAKTIKUM (LATIHAN)
    add_chapter_heading("HASIL PRAKTIKUM (LATIHAN)")

    add_heading_2("Latihan 1 – Implementasi Pattern Model-View-Controller (MVC) pada Class Student")

    add_body_p("Praktikan mengimplementasikan arsitektur Model-View-Controller (MVC) pada entitas Student menggunakan bahasa pemrograman Java. Praktikan membagi kode ke dalam tiga komponen utama: Class Student sebagai Model data (POJO), StudentView sebagai View untuk visualisasi representasi data ke konsol, dan StudentController sebagai Controller untuk mengontrol aliran data dan menghubungkan Model dengan View, serta mengeksekusinya melalui main class MVCPatternDemo.")

    add_subheading("I.1.1.  Source Code")

    code_student = '''public class Student {
    private String rollNo;
    private String name;

    public String getRollNo() {
        return rollNo;
    }

    public void setRollNo(String rollNo) {
        this.rollNo = rollNo;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}'''
    add_code_block(code_student, "Model: Student.java")

    code_student_view = '''public class StudentView {
    public void printStudentDetails(String studentName, String studentRollNo) {
        System.out.println("Student: ");
        System.out.println("Name: " + studentName);
        System.out.println("Roll No: " + studentRollNo);
    }
}'''
    add_code_block(code_student_view, "View: StudentView.java")

    code_student_controller = '''public class StudentController {
    private Student model;
    private StudentView view;

    public StudentController(Student model, StudentView view) {
        this.model = model;
        this.view = view;
    }

    public void setStudentName(String name) {
        model.setName(name);
    }

    public String getStudentName() {
        return model.getName();
    }

    public void setStudentRollNo(String rollNo) {
        model.setRollNo(rollNo);
    }

    public String getStudentRollNo() {
        return model.getRollNo();
    }

    public void updateView() {
        view.printStudentDetails(model.getName(), model.getRollNo());
    }
}'''
    add_code_block(code_student_controller, "Controller: StudentController.java")

    code_mvc_demo = '''public class MVCPatternDemo {
    public static void main(String[] args) {
        Student model = retriveStudentFromDatabase();
        StudentView view = new StudentView();
        StudentController controller = new StudentController(model, view);

        controller.updateView();

        // Update model data
        controller.setStudentName("John");

        controller.updateView();
    }

    private static Student retriveStudentFromDatabase() {
        Student student = new Student();
        student.setName("Robert");
        student.setRollNo("10");
        return student;
    }
}'''
    add_code_block(code_mvc_demo, "Main Class: MVCPatternDemo.java")

    add_subheading("I.1.2.  Screenshot")
    add_figure("assets/screenshots/screenshot_latihan1_run.png", "gambar I-1 Hasil program Latihan MVC Pattern (Student)")
    add_figure("assets/screenshots/screenshot_latihan1_git.png", "gambar I-2 Hasil commit Git Terminal Latihan MVC Pattern (Student)")

    add_subheading("I.1.3.  Analisis")
    add_body_p("Berdasarkan hasil program yang dieksekusi menggunakan Java compiler (javac) dan Java runtime (java) pada file MVCPatternDemo.java di direktori code, program berhasil mengimplementasikan arsitektur Model-View-Controller (MVC) secara menyeluruh. Pola MVC ini diterapkan untuk mencapai prinsip Separation of Concerns (pemisahan tanggung jawab), di mana setiap class memiliki peran yang sangat spesifik dan terisolasi:")

    add_body_p("1. Model (Student.java): Berperan sebagai representasi entitas data atau Plain Old Java Object (POJO). Class ini mengkapsulasi atribut privat 'rollNo' dan 'name' dengan menyediakan metode pengakses (getter) dan pemodifikasi (setter). Model tidak memiliki ketergantungan apa pun terhadap antarmuka pengguna (View) maupun Controller, sehingga data bersifat independen dan murni merepresentasikan state aplikasi.")

    add_body_p("2. View (StudentView.java): Berperan secara eksklusif dalam menangani representasi visual data ke pengguna melalui media konsol. Metode printStudentDetails() menerima data mentah berupa string nama dan nomor roll mahasiswa, lalu menampilkannya dengan format yang rapi. View tidak memiliki hak untuk memanipulasi atau menyimpan data, melainkan hanya bertindak sebagai penampil pasif.")

    add_body_p("3. Controller (StudentController.java): Bertindak sebagai penghubung (mediator) dan pengendali alur antara Model dan View. Controller memegang referensi ke objek Student dan StudentView. Ketika terjadi perubahan data seperti pemanggilan controller.setStudentName(\"John\"), Controller menginstruksikan Model untuk memperbarui state internalnya. Selanjutnya melalui pemanggilan controller.updateView(), Controller mengambil data terbaru dari Model dan meneruskannya ke View untuk ditampilkan.")

    add_body_p("4. Main Class (MVCPatternDemo.java): Mengorkestrasi pengujian program dengan mensimulasikan pengambilan data awal mahasiswa dari basis data ('Robert', Roll No: '10'). Pada pemanggilan updateView() pertama, konsol menampilkan nama 'Robert'. Setelah Controller memodifikasi nama menjadi 'John', pemanggilan updateView() kedua secara tepat menampilkan perubahan tersebut. Pemisahan struktur ini memberikan fleksibilitas tinggi: apabila di masa depan tampilan antarmuka hendak diubah ke antarmuka grafis (GUI) atau web, komponen Model dan Controller tetap dapat digunakan tanpa memerlukan perubahan logika bisnis sama sekali.")

    doc.add_page_break()

    # BAB II HASIL PRAKTIKUM (TUGAS)
    add_chapter_heading("HASIL PRAKTIKUM (TUGAS)")

    # II.1. Tugas 1
    add_heading_2("Tugas 1 – Melengkapi Implementasi Pattern MVC pada Class Course")

    add_body_p("Praktikan melengkapi potongan kode program yang belum tuntas pada Class Model Course.java dan Class Controller CourseController.java sesuai dengan instruksi modul praktikum. Praktikan mengimplementasikan metode getter dan setter untuk atribut CourseName, CourseId, dan CourseCategory pada Model, melengkapi tipe variabel serta konstruktor pada Controller, dan memverifikasi hasil eksekusinya pada konsol agar sesuai dengan luaran yang ditargetkan.")

    add_subheading("II.1.1.  Source Code")

    code_course_model = '''public class Course {
    private String CourseName;
    private String CourseId;
    private String CourseCategory;

    public String getId() {
        return CourseId;
    }

    public void setId(String id) {
        this.CourseId = id;
    }

    public String getName() {
        return CourseName;
    }

    public void setName(String name) {
        this.CourseName = name;
    }

    public String getCategory() {
        return CourseCategory;
    }

    public void setCategory(String category) {
        this.CourseCategory = category;
    }
}'''
    add_code_block(code_course_model, "Model: Course.java (Dilengkapi)")

    code_course_view = '''public class CourseView {
    public void printCourseDetails(String CourseName, String CourseId, String CourseCategory) {
        System.out.println("Course Details: ");
        System.out.println("Name: " + CourseName);
        System.out.println("Course ID: " + CourseId);
        System.out.println("Course Category: " + CourseCategory);
    }
}'''
    add_code_block(code_course_view, "View: CourseView.java")

    code_course_controller = '''public class CourseController {
    private Course model;
    private CourseView view;

    public CourseController(Course model, CourseView view) {
        this.model = model;
        this.view = view;
    }

    public void setCourseName(String name) {
        model.setName(name);
    }

    public String getCourseName() {
        return model.getName();
    }

    public void setCourseId(String id) {
        model.setId(id);
    }

    public String getCourseId() {
        return model.getId();
    }

    public void setCourseCategory(String category) {
        model.setCategory(category);
    }

    public String getCourseCategory() {
        return model.getCategory();
    }

    public void updateView() {
        view.printCourseDetails(model.getName(), model.getId(), model.getCategory());
    }
}'''
    add_code_block(code_course_controller, "Controller: CourseController.java (Dilengkapi)")

    code_course_main = '''public class MVCPatternDemo {
    public static void main(String[] args) {
        // fetch student record based on his roll no from the database
        Course model = retriveCourseFromDatabase();

        // Create a view : to write course details on console
        CourseView view = new CourseView();

        CourseController controller = new CourseController(model, view);

        controller.updateView();

        // update model data
        controller.setCourseName("Python");
        System.out.println("\\nAfter updating, Course Details are as follows");
        controller.updateView();
    }

    private static Course retriveCourseFromDatabase() {
        Course course = new Course();
        course.setName("Java");
        course.setId("01");
        course.setCategory("Programming");
        return course;
    }
}'''
    add_code_block(code_course_main, "Main Class: MVCPatternDemo.java (Tugas 1)")

    add_subheading("II.1.2.  Screenshot")
    add_figure("assets/screenshots/screenshot_tugas1_run.png", "gambar II-1 Hasil program Tugas 1 MVC Pattern (Course)")
    add_figure("assets/screenshots/screenshot_tugas1_git.png", "gambar II-2 Hasil commit Git Terminal Tugas 1 MVC Pattern (Course)")

    add_subheading("II.1.3.  Analisis")
    add_body_p("Berdasarkan hasil program yang dikompilasi dan dijalankan pada file MVCPatternDemo.java di direktori tugas/tugas1, program berhasil dieksekusi dengan luaran yang 100% identik dengan spesifikasi output pada modul halaman 50. Analisis mendalam terhadap proses perbaikan dan logika program adalah sebagai berikut:")

    add_body_p("1. Analisis Perbaikan Model Course.java: Pada modul aslinya, class Course hanya mendefinisikan atribut privat CourseName, CourseId, dan CourseCategory tanpa implementasi tubuh method (hanya menyisakan baris komentar instruksi seperti //mengembalikan nilai id Course dan //set id Course). Praktikan melengkapi seluruh blok kode tersebut dengan mengembalikan nilai atribut pada getter dan menetapkan parameter ke atribut menggunakan kata kunci 'this' pada setter. Hal ini menjamin prinsip enkapsulasi data terpenuhi secara utuh.")

    add_body_p("2. Analisis Perbaikan Controller CourseController.java: Pada modul asli, terdapat placeholder berupa titik-titik (private . . . model; private . . . view;) dan deklarasi konstruktor yang belum memiliki tipe parameter. Praktikan memperbaiki deklarasi tersebut menjadi tipe referensi konkret Course dan CourseView. Konstruktor CourseController(Course model, CourseView view) diinisialisasi untuk merekatkan instance model dan view yang akan dikontrol. Seluruh metode delegasi (setCourseName, getCourseName, setCourseId, getCourseId, setCourseCategory, getCourseCategory) serta updateView() dihubungkan secara tepat ke instance model dan view terkait.")

    add_body_p("3. Verifikasi Alur Eksekusi Runtime: Simulasi pengambilan basis data melalui retriveCourseFromDatabase() menghasilkan objek Course dengan konfigurasi awal: Nama='Java', ID='01', dan Kategori='Programming'. Saat controller.updateView() dipanggil untuk pertama kalinya, konsol menampilkan detail kursus Java secara akurat. Kemudian melalui perintah controller.setCourseName(\"Python\"), nama kursus pada model berhasil diperbarui. Setelah mencetak header informasi pembaruan, pemanggilan controller.updateView() kedua secara dinamis menampilkan data kursus terkini dengan Nama='Python', sementara ID ('01') dan Kategori ('Programming') tetap terjaga. Ini membuktikan bahwa Controller secara sukses memfasilitasi mutasi data model sekaligus mengkoordinasikan visualisasi pada View.")

    doc.add_page_break()

    # II.2. Tugas 2
    add_heading_2("Tugas 2 – Refaktorisasi Kode Monolitik Kasir Menjadi Pattern MVC")

    add_body_p("Praktikan menganalisis kode program Kasir awal yang bersifat prosedural monolitik di mana seluruh inisiasi data barang, manipulasi kuantitas transaksi, kalkulasi total pembayaran, dan pencetakan informasi berada di dalam satu fungsi main. Praktikan kemudian merekayasa ulang (refactoring) kode tersebut ke dalam arsitektur Model-View-Controller (MVC) yang modular dengan memisahkan entitas menjadi Class Barang (Model), KasirView (View), KasirController (Controller), dan KasirMVCDemo (Main Class).")

    add_subheading("II.2.1.  Kode Sebelum Modifikasi (Kasir Awal Prosedural)")

    code_kasir_awal = '''import java.util.Scanner;

public class KasirAwal {
    public static void main(String[] args) {
        String nama;
        int harga;
        int stok;
        Scanner in = new Scanner(System.in);

        // Inisiasi
        nama = "GeForce GTX 1060Ti";
        harga = 3000;
        stok = 10;

        // Tampilan info
        System.out.println("Nama Barang : " + nama);
        System.out.println("Harga Barang : " + harga);
        System.out.println("Stok Barang : " + stok);

        // Transaksi pembelian
        int qty;
        System.out.println("Transaksi Pembelian");
        System.out.print("Jumlah barang : ");
        qty = in.nextInt();
        stok = stok - qty;

        // Print Pembayaran
        System.out.println("Jumlah Bayar : " + qty * harga);

        // Tampilan info
        System.out.println("Nama Barang : " + nama);
        System.out.println("Harga Barang : " + harga);
        System.out.println("Stok Barang : " + stok);
    }
}'''
    add_code_block(code_kasir_awal, "Kode Awal: KasirAwal.java (Prosedural Monolitik)")

    add_subheading("II.2.2.  Kode Sesudah Modifikasi (Arsitektur MVC)")

    code_barang_model = '''public class Barang {
    private String nama;
    private int harga;
    private int stok;

    public Barang() {}

    public Barang(String nama, int harga, int stok) {
        this.nama = nama;
        this.harga = harga;
        this.stok = stok;
    }

    public String getNama() {
        return nama;
    }

    public void setNama(String nama) {
        this.nama = nama;
    }

    public int getHarga() {
        return harga;
    }

    public void setHarga(int harga) {
        this.harga = harga;
    }

    public int getStok() {
        return stok;
    }

    public void setStok(int stok) {
        this.stok = stok;
    }

    public int hitungTotalBayar(int qty) {
        return qty * this.harga;
    }

    public boolean kurangiStok(int qty) {
        if (qty <= this.stok) {
            this.stok -= qty;
            return true;
        }
        return false;
    }
}'''
    add_code_block(code_barang_model, "Model: Barang.java")

    code_kasir_view = '''public class KasirView {
    public void printBarangDetails(String nama, int harga, int stok) {
        System.out.println("Nama Barang : " + nama);
        System.out.println("Harga Barang : " + harga);
        System.out.println("Stok Barang : " + stok);
    }

    public void printTransaksiHeader() {
        System.out.println("Transaksi Pembelian");
    }

    public void printPembayaran(int totalBayar) {
        System.out.println("Jumlah Bayar : " + totalBayar);
    }

    public void printPesan(String pesan) {
        System.out.println(pesan);
    }
}'''
    add_code_block(code_kasir_view, "View: KasirView.java")

    code_kasir_controller = '''public class KasirController {
    private Barang model;
    private KasirView view;

    public KasirController(Barang model, KasirView view) {
        this.model = model;
        this.view = view;
    }

    public void setBarangNama(String nama) {
        model.setNama(nama);
    }

    public String getBarangNama() {
        return model.getNama();
    }

    public void setBarangHarga(int harga) {
        model.setHarga(harga);
    }

    public int getBarangHarga() {
        return model.getHarga();
    }

    public void setBarangStok(int stok) {
        model.setStok(stok);
    }

    public int getBarangStok() {
        return model.getStok();
    }

    public void updateView() {
        view.printBarangDetails(model.getNama(), model.getHarga(), model.getStok());
    }

    public boolean prosesTransaksi(int qty) {
        if (qty <= 0) {
            view.printPesan("Jumlah pembelian tidak valid!");
            return false;
        }
        if (model.kurangiStok(qty)) {
            int totalBayar = model.hitungTotalBayar(qty);
            view.printPembayaran(totalBayar);
            return true;
        } else {
            view.printPesan("Transaksi Gagal: Stok tidak mencukupi!");
            return false;
        }
    }
}'''
    add_code_block(code_kasir_controller, "Controller: KasirController.java")

    code_kasir_mvc_demo = '''import java.util.Scanner;

public class KasirMVCDemo {
    public static void main(String[] args) {
        Scanner in = new Scanner(System.in);

        // 1. Inisialisasi Model, View, dan Controller
        Barang model = retriveBarangFromDatabase();
        KasirView view = new KasirView();
        KasirController controller = new KasirController(model, view);

        // 2. Tampilan info barang awal melalui controller
        controller.updateView();

        // 3. Transaksi pembelian
        view.printTransaksiHeader();
        System.out.print("Jumlah barang : ");
        int qty = in.nextInt();

        // 4. Proses transaksi melalui controller dan cetak pembayaran
        controller.prosesTransaksi(qty);

        // 5. Tampilan info barang setelah transaksi diperbarui
        controller.updateView();
    }

    private static Barang retriveBarangFromDatabase() {
        Barang barang = new Barang();
        barang.setNama("GeForce GTX 1060Ti");
        barang.setHarga(3000);
        barang.setStok(10);
        return barang;
    }
}'''
    add_code_block(code_kasir_mvc_demo, "Main Class: KasirMVCDemo.java")

    add_subheading("II.2.3.  Screenshot")
    add_figure("assets/screenshots/screenshot_tugas2_run.png", "gambar II-3 Hasil program Tugas 2 Refaktorisasi Kasir MVC")
    add_figure("assets/screenshots/screenshot_tugas2_git.png", "gambar II-4 Hasil commit Git Terminal Tugas 2 Kasir MVC")

    add_subheading("II.2.4.  Analisis")
    add_body_p("Refaktorisasi program kasir dari paradigma prosedural ke arsitektur Model-View-Controller (MVC) membawa perbaikan fundamental pada struktur rancangan perangkat lunak. Perbandingan analisis komparatif antara kode awal dan sesudah modifikasi diuraikan sebagai berikut:")

    add_body_p("1. Kelemahan Kode Prosedural Awal (KasirAwal.java): Pada kode awal, semua variabel primitif (nama, harga, stok, qty) dideklarasikan secara lokal di dalam method main(). Logika bisnis (stok = stok - qty, qty * harga) bercampur aduk secara langsung dengan pemanggilan System.out.println() dan operasi pembacaan Scanner. Desain seperti ini memiliki tingkat kopling sangat tinggi (tight coupling) dan kohesi rendah (low cohesion). Jika di kemudian hari sistem membutuhkan validasi stok, pencatatan transaksi ke database, atau perubahan antarmuka ke desktop GUI, seluruh kode main harus dibongkar ulang, yang sangat rentan memicu munculnya bug regresi.")

    add_body_p("2. Peran Model (Barang.java): Melalui pola MVC, entitas barang diangkat menjadi class tersendiri dengan atribut privat yang terenkapsulasi. Model memegang logika bisnis inti yang berkaitan langsung dengan entitas barang itu sendiri, yaitu method kurangiStok(int qty) dengan pengecekan batas stok dan hitungTotalBayar(int qty). Model sepenuhnya bersih dari dependensi pustaka I/O (Scanner atau System.out), menjadikannya sangat mudah diuji secara otomatis melalui unit testing.")

    add_body_p("3. Peran View (KasirView.java): Mengambil alih 100% tanggung jawab presentasi visual. View menyediakan method printBarangDetails(), printTransaksiHeader(), printPembayaran(), dan printPesan(). View tidak pernah menghitung harga atau mengubah nilai stok; View hanya bertugas mencetak parameter yang dikirimkan kepadanya. Dengan demikian, jika format laporan transaksi perlu diubah (misalnya menggunakan format mata uang Rupiah atau dicetak ke struk fisik), pengembang hanya perlu memodifikasi KasirView tanpa menyentuh class lain.")

    add_body_p("4. Peran Controller (KasirController.java): Bertindak sebagai pengatur alur kerja transaksi kasir. Controller menerima masukan kuantitas pembelian melalui method prosesTransaksi(int qty), melakukan validasi kuantitas, memanggil logika pengurangan stok pada Model, menghitung total bayar, lalu meneruskan hasil kalkulasi tersebut ke View untuk dicetak. Controller juga memfasilitasi sinkronisasi data melalui updateView().")

    add_body_p("5. Validasi Fungsionalitas Hasil Eksekusi: Pada saat program dijalankan dengan masukan kuantitas pembelian 2 unit untuk barang 'GeForce GTX 1060Ti' seharga 3000 dan stok awal 10, sistem menghitung total bayar sebesar 6000 dan mengurangi stok barang menjadi 8. Luaran yang dihasilkan oleh KasirMVCDemo sama persis secara fungsional dengan KasirAwal, namun fondasi kodenya kini telah bertransformasi menjadi arsitektur berstandar industri yang modular, aman, terenkapsulasi, dan sangat mudah dirawat (maintainable).")

    doc.add_page_break()

    # BAB III REPOSITORY GITHUB
    add_chapter_heading("STATUS DAN RIWAYAT REPOSITORY (GIT & GITHUB)")

    add_body_p("Sesuai instruksi penugasan praktikum, seluruh pengerjaan kode sumber latihan dan tugas dikelola menggunakan sistem kontrol versi Git secara lokal yang terhubung langsung dengan aplikasi GitHub Desktop pada repositori daring GitHub. Setiap tahapan pengerjaan (Latihan 1, Tugas 1, dan Tugas 2) direkam melalui komit berkala dengan pesan komit yang jelas dan deskriptif.")

    add_subheading("Riwayat Komit Bertahap (Step-by-Step Commit History):")
    add_body_p("1. Komit 1 (Root Commit a022aba): 'Implementasi Latihan MVC Pattern - Student' – Merekam berkas sumber Latihan MVC meliputi Student.java, StudentView.java, StudentController.java, dan MVCPatternDemo.java.")
    add_body_p("2. Komit 2 (Commit d87ecdb): 'Implementasi Tugas 1 MVC Pattern - Course' – Merekam penyelesaian kode lengkap Course.java, CourseView.java, CourseController.java, dan MVCPatternDemo.java pada folder tugas/tugas1.")
    add_body_p("3. Komit 3 (Commit 8e65c88): 'Implementasi Tugas 2 MVC Pattern - Kasir Refactoring' – Merekam hasil refaktorisasi arsitektur MVC untuk sistem kasir meliputi Barang.java, KasirView.java, KasirController.java, KasirMVCDemo.java, dan kode pembanding KasirAwal.java pada folder tugas/tugas2.")

    add_subheading("Tautan dan Konfigurasi Repositori GitHub:")
    add_body_p("• Nama Repositori : Modul-3-Design-Pattern-dan-Object-Composing")
    add_body_p("• URL Remote Origin : https://github.com/muchammadrizalsobari/Modul-3-Design-Pattern-dan-Object-Composing.git")
    add_body_p("• Akun Pengguna : muchammadrizalsobari (Nej)")
    add_body_p("• Branch Aktif : main (upstream to origin/main)")

    add_subheading("Tampilan Status Git Terminal dan Sinkronisasi Push:")
    add_figure("assets/screenshots/screenshot_repo_overview.png", "gambar II-5 Status Repository dan Log Git Terminal Terhubung ke GitHub")

    add_body_p("Melalui terminal Git, perintah git push -u origin main telah berhasil dieksekusi, sehingga seluruh riwayat percabangan lokal kini telah tersinkronisasi secara sempurna di cloud GitHub. Repositori ini juga secara otomatis terdeteksi dan aktif pada aplikasi GitHub Desktop lokal untuk memudahkan pelacakan perubahan lanjutan.")

    doc.add_page_break()

    # BAB IV KESIMPULAN
    add_chapter_heading("KESIMPULAN")

    add_body_p("Berdasarkan pelaksanaan praktikum Modul 3 mengenai 'Design Pattern & Object Composing' yang difokuskan pada penerapan pola perancangan Model-View-Controller (MVC) menggunakan bahasa pemrograman Java, dapat ditarik beberapa kesimpulan penting sebagai berikut:")

    add_body_p("1. Pemahaman Konsep MVC Pattern: Pola arsitektur Model-View-Controller (MVC) terbukti sangat efektif dalam memecah kompleksitas sistem perangkat lunak ke dalam tiga domain tanggung jawab yang berbeda (Separation of Concerns). Model berfungsi murni sebagai pembawa data dan logika bisnis (POJO), View bertanggung jawab secara eksklusif atas visualisasi antarmuka pengguna, dan Controller berperan sebagai mediator cerdas yang mengontrol aliran data dan menyinkronkan status antara Model dan View.")

    add_body_p("2. Peningkatan Kualitas Kode dan Maintainability: Melalui latihan implementasi Class Student dan penyelesaian Tugas 1 pada Class Course, praktikan membuktikan bahwa keterpisahan komponen mencegah terjadinya tumpang tindih logika. Ketika struktur representasi data diubah, kode tampilan tidak terganggu, dan sebaliknya, modifikasi format output antarmuka tidak memerlukan perombakan struktur data Model. Hal ini menciptakan basis kode yang sangat adaptif, mudah diperluas (extensible), dan mudah diuji secara independen (unit testable).")

    add_body_p("3. Keberhasilan Refaktorisasi Prosedural ke MVC: Pada penugasan Tugas 2, program kasir monolitik berhasil direfaktor secara elegan menjadi class Barang, KasirView, KasirController, dan KasirMVCDemo. Refaktorisasi ini berhasil mengeliminasi kelemahan rancangan awal yang memiliki tingkat kopling tinggi (tight coupling) dan menggantinya dengan struktur loosely coupled yang memiliki integritas enkapsulasi tinggi, validasi stok yang aman, serta pemisahan logika transaksi yang modular.")

    add_body_p("4. Penerapan Version Control System yang Profesional: Seluruh tahapan praktikum telah berhasil didokumentasikan dan dipublikasikan menggunakan Git dan GitHub Desktop melalui komit bertahap yang rapi dan terstruktur pada repositori Modul-3-Design-Pattern-dan-Object-Composing. Praktik ini memastikan setiap iterasi pengembangan kode terekam dengan jelas, mempermudah pelacakan histori, dan mencerminkan standar rekayasa perangkat lunak modern.")

    output_docx = r"D:\RIZAL\KULIAH\semester 5\prak impementasi\MODUL 3\Muchammad Rizal  Sobari- Modul 3 - Design Pattern & Object Composing.docx"
    doc.save(output_docx)
    print(f"Document saved successfully at: {output_docx}")

    # Export to PDF
    pdf_path = r"D:\RIZAL\KULIAH\semester 5\prak impementasi\MODUL 3\Muchammad Rizal  Sobari  Modul 3 - Design Pattern & Object Composing.pdf"
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        wdoc = word.Documents.Open(os.path.abspath(output_docx))
        wdoc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        wdoc.Close()
        print(f"PDF generated successfully at: {pdf_path}")
    finally:
        word.Quit()

if __name__ == '__main__':
    create_document()
