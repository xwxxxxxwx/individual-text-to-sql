import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from pypdf import PdfReader

base = r'C:\Users\33238\Desktop\2026 semester1 选课\2.STATS 762\4.23 期中考试准备\过往考试'

# Check appendix-1 from 2021 and full content of 2020 model answers
files = [
    r'2021\appendix-1.pdf',
    r'2020\test2020_modelanswers.pdf',
    r'2021\test2021ans-1.pdf',
    r'2019\test2019ans.pdf',
]

for fpath in files:
    full = os.path.join(base, fpath)
    try:
        reader = PdfReader(full)
        print('=== {} ({} pages) ==='.format(fpath, len(reader.pages)))
        # Print first page header
        text = reader.pages[0].extract_text()
        print('--- Page 1 (first 800 chars) ---')
        print(text[:800].replace(chr(8722), '-'))
        # Also print last page to see what topic it ends with
        if len(reader.pages) > 1:
            text_last = reader.pages[-1].extract_text()
            print('--- Last page (first 500 chars) ---')
            print(text_last[:500].replace(chr(8722), '-'))
        print()
    except Exception as e:
        print('Error {}: {}'.format(fpath, e))
