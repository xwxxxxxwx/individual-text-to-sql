import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from pypdf import PdfReader

path = r'C:\Users\33238\Desktop\2026 semester1 选课\2.STATS 762\4.23 期中考试准备\过往考试\2025\Test 2025.pdf'
reader = PdfReader(path)
print('Pages:', len(reader.pages))
for i, page in enumerate(reader.pages):
    text = page.extract_text()
    print('=== Page {} ==='.format(i+1))
    print(text.replace(chr(8722), '-'))
    print()
