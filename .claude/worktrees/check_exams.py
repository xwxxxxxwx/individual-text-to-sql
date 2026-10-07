import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from pypdf import PdfReader

base = r'C:\Users\33238\Desktop\2026 semester1 选课\2.STATS 762\4.23 期中考试准备\过往考试'
files = [
    (r'2019\test2019.pdf', 'Midterm 2019'),
    (r'2020\test2020-1.pdf', 'Test 2020'),
    (r'2021\test2021-2.pdf', 'Test 2021'),
    (r'2025\Test 2025.pdf', 'Test 2025'),
]

for fpath, label in files:
    full = os.path.join(base, fpath)
    try:
        reader = PdfReader(full)
        first_text = reader.pages[0].extract_text()
        print('=== {} ({} pages) ==='.format(label, len(reader.pages)))
        print(first_text[:600].replace(chr(8722), '-'))
        print()
    except Exception as e:
        print('Error {}: {}'.format(label, e))
