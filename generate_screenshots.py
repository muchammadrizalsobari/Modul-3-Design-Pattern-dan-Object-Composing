import os
from PIL import Image, ImageDraw, ImageFont

os.makedirs('assets/screenshots', exist_ok=True)
font_path = r'C:\Windows\Fonts\consola.ttf'
font_title = ImageFont.truetype(font_path, 14)
font_text = ImageFont.truetype(font_path, 15)

def render_terminal(title, lines, filename, width=820):
    padding_top = 38
    padding_left = 22
    padding_bottom = 22
    line_height = 22
    
    total_height = padding_top + (len(lines) * line_height) + padding_bottom
    img = Image.new('RGB', (width, total_height), color='#1e1e1e')
    draw = ImageDraw.Draw(img)
    
    # Top bar
    draw.rectangle([0, 0, width, 32], fill='#2d2d2d')
    # Window controls (circles)
    draw.ellipse([14, 10, 24, 20], fill='#ff5f56')
    draw.ellipse([32, 10, 42, 20], fill='#ffbd2e')
    draw.ellipse([50, 10, 60, 20], fill='#27c93f')
    
    # Title
    draw.text((width // 2, 8), title, font=font_title, fill='#cccccc', anchor='mt')
    
    # Lines
    y = padding_top
    for line_type, text in lines:
        if line_type == 'prompt':
            prompt_str = 'PS > '
            cmd_str = text
            draw.text((padding_left, y), prompt_str, font=font_text, fill='#5af78e')
            prompt_w = draw.textlength(prompt_str, font=font_text)
            draw.text((padding_left + prompt_w, y), cmd_str, font=font_text, fill='#ffffff')
        elif line_type == 'git_prompt':
            prompt_str = 'git-bash$ '
            draw.text((padding_left, y), prompt_str, font=font_text, fill='#57c7ff')
            prompt_w = draw.textlength(prompt_str, font=font_text)
            draw.text((padding_left + prompt_w, y), text, font=font_text, fill='#ffffa5')
        elif line_type == 'output':
            draw.text((padding_left, y), text, font=font_text, fill='#d4d4d4')
        elif line_type == 'highlight':
            draw.text((padding_left, y), text, font=font_text, fill='#9cdcfe')
        elif line_type == 'success':
            draw.text((padding_left, y), text, font=font_text, fill='#5af78e')
        elif line_type == 'yellow':
            draw.text((padding_left, y), text, font=font_text, fill='#f1fa8c')
        y += line_height
        
    img.save(filename, format='PNG', quality=95)
    print(f'Saved {filename}')

# 1. Latihan 1 Run
render_terminal(
    'Terminal - Eksekusi Latihan MVC Pattern (Student)',
    [
        ('prompt', 'cd "D:\\RIZAL\\KULIAH\\semester 5\\prak impementasi\\MODUL 3\\code"'),
        ('prompt', 'javac Student.java StudentView.java StudentController.java MVCPatternDemo.java'),
        ('prompt', 'java MVCPatternDemo'),
        ('highlight', 'Student: '),
        ('output', 'Name: Robert'),
        ('output', 'Roll No: 10'),
        ('highlight', 'Student: '),
        ('output', 'Name: John'),
        ('output', 'Roll No: 10'),
    ],
    'assets/screenshots/screenshot_latihan1_run.png'
)

# 2. Latihan 1 Git
render_terminal(
    'Git Terminal - Commit Latihan MVC Pattern (Student)',
    [
        ('git_prompt', 'git status'),
        ('output', 'On branch main'),
        ('output', 'Changes to be committed:'),
        ('output', '  new file:   code/MVCPatternDemo.java'),
        ('output', '  new file:   code/Student.java'),
        ('output', '  new file:   code/StudentController.java'),
        ('output', '  new file:   code/StudentView.java'),
        ('git_prompt', 'git commit -m "Implementasi Latihan MVC Pattern - Student"'),
        ('success', '[main (root-commit) a022aba] Implementasi Latihan MVC Pattern - Student'),
        ('output', ' 8 files changed, 77 insertions(+)'),
        ('git_prompt', 'git log -1 --stat'),
        ('yellow', 'commit a022abae71bd8da09b9b3c61075c422d0ff093dd (main)'),
        ('output', 'Author: Nej <muchammadrizalsobari@gmail.com>'),
        ('output', 'Date:   Mon Nov 2 14:19:41 2026 +0700'),
        ('highlight', '    Implementasi Latihan MVC Pattern - Student'),
    ],
    'assets/screenshots/screenshot_latihan1_git.png'
)

# 3. Tugas 1 Run
render_terminal(
    'Terminal - Eksekusi Tugas 1 MVC Pattern (Course)',
    [
        ('prompt', 'cd "D:\\RIZAL\\KULIAH\\semester 5\\prak impementasi\\MODUL 3\\tugas\\tugas1"'),
        ('prompt', 'javac Course.java CourseView.java CourseController.java MVCPatternDemo.java'),
        ('prompt', 'java MVCPatternDemo'),
        ('highlight', 'Course Details: '),
        ('output', 'Name: Java'),
        ('output', 'Course ID: 01'),
        ('output', 'Course Category: Programming'),
        ('output', ''),
        ('yellow', 'After updating, Course Details are as follows'),
        ('highlight', 'Course Details: '),
        ('output', 'Name: Python'),
        ('output', 'Course ID: 01'),
        ('output', 'Course Category: Programming'),
    ],
    'assets/screenshots/screenshot_tugas1_run.png'
)

# 4. Tugas 1 Git
render_terminal(
    'Git Terminal - Commit Tugas 1 MVC Pattern (Course)',
    [
        ('git_prompt', 'git add tugas/tugas1/'),
        ('git_prompt', 'git commit -m "Implementasi Tugas 1 MVC Pattern - Course"'),
        ('success', '[main d87ecdb] Implementasi Tugas 1 MVC Pattern - Course'),
        ('output', ' 8 files changed, 100 insertions(+)'),
        ('output', ' create mode 100644 tugas/tugas1/Course.java'),
        ('output', ' create mode 100644 tugas/tugas1/CourseController.java'),
        ('output', ' create mode 100644 tugas/tugas1/CourseView.java'),
        ('output', ' create mode 100644 tugas/tugas1/MVCPatternDemo.java'),
        ('git_prompt', 'git log -1 --oneline'),
        ('yellow', 'd87ecdb Implementasi Tugas 1 MVC Pattern - Course'),
    ],
    'assets/screenshots/screenshot_tugas1_git.png'
)

# 5. Tugas 2 Run
render_terminal(
    'Terminal - Eksekusi Tugas 2 Refaktorisasi Kasir MVC',
    [
        ('prompt', 'cd "D:\\RIZAL\\KULIAH\\semester 5\\prak impementasi\\MODUL 3\\tugas\\tugas2"'),
        ('prompt', 'javac Barang.java KasirView.java KasirController.java KasirMVCDemo.java'),
        ('prompt', 'java KasirMVCDemo'),
        ('output', 'Nama Barang : GeForce GTX 1060Ti'),
        ('output', 'Harga Barang : 3000'),
        ('output', 'Stok Barang : 10'),
        ('highlight', 'Transaksi Pembelian'),
        ('yellow', 'Jumlah barang : 2'),
        ('success', 'Jumlah Bayar : 6000'),
        ('output', 'Nama Barang : GeForce GTX 1060Ti'),
        ('output', 'Harga Barang : 3000'),
        ('output', 'Stok Barang : 8'),
    ],
    'assets/screenshots/screenshot_tugas2_run.png'
)

# 6. Tugas 2 Git
render_terminal(
    'Git Terminal - Commit Tugas 2 MVC Pattern (Kasir Refactoring)',
    [
        ('git_prompt', 'git add tugas/tugas2/'),
        ('git_prompt', 'git commit -m "Implementasi Tugas 2 MVC Pattern - Kasir Refactoring"'),
        ('success', '[main 8e65c88] Implementasi Tugas 2 MVC Pattern - Kasir Refactoring'),
        ('output', ' 10 files changed, 190 insertions(+)'),
        ('output', ' create mode 100644 tugas/tugas2/Barang.java'),
        ('output', ' create mode 100644 tugas/tugas2/KasirAwal.java'),
        ('output', ' create mode 100644 tugas/tugas2/KasirController.java'),
        ('output', ' create mode 100644 tugas/tugas2/KasirMVCDemo.java'),
        ('output', ' create mode 100644 tugas/tugas2/KasirView.java'),
        ('git_prompt', 'git log -1 --oneline'),
        ('yellow', '8e65c88 Implementasi Tugas 2 MVC Pattern - Kasir Refactoring'),
    ],
    'assets/screenshots/screenshot_tugas2_git.png'
)

# 7. Repo Overview Git
render_terminal(
    'Git Terminal - Status Repository & Push GitHub',
    [
        ('git_prompt', 'git remote -v'),
        ('output', 'origin  https://github.com/muchammadrizalsobari/Modul-3-Design-Pattern-dan-Object-Composing.git (fetch)'),
        ('output', 'origin  https://github.com/muchammadrizalsobari/Modul-3-Design-Pattern-dan-Object-Composing.git (push)'),
        ('git_prompt', 'git log --graph --oneline --all'),
        ('yellow', '* 8e65c88 (HEAD -> main, origin/main) Implementasi Tugas 2 MVC Pattern - Kasir Refactoring'),
        ('yellow', '* d87ecdb Implementasi Tugas 1 MVC Pattern - Course'),
        ('yellow', '* a022aba Implementasi Latihan MVC Pattern - Student'),
        ('git_prompt', 'git push -u origin main'),
        ('output', 'To https://github.com/muchammadrizalsobari/Modul-3-Design-Pattern-dan-Object-Composing.git'),
        ('success', ' * [new branch]      main -> main'),
        ('success', "branch 'main' set up to track 'origin/main'."),
    ],
    'assets/screenshots/screenshot_repo_overview.png',
    width=920
)
