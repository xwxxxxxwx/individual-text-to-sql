"""Parse Word .doc template — extract structure, styles, paragraphs with formatting info."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client
import os
import json

DOC_PATH = r"C:\Users\33238\Desktop\2026 semester1 选课\4.DATASCI 791 企业项目课程\2.学校要求的交付内容\01_论文写作模板\Word Dissertation Template.doc"

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

try:
    doc = word.Documents.Open(FileName=DOC_PATH, ReadOnly=True, ConfirmConversions=False)

    print("=" * 80)
    print("DOCUMENT METADATA")
    print("=" * 80)
    print(f"Name: {doc.Name}")
    print(f"Built-in props — Author: {doc.BuiltInDocumentProperties('Author').Value}")
    print(f"Pages: {doc.ComputeStatistics(2)}")  # wdStatisticPages
    print(f"Words: {doc.ComputeStatistics(0)}")
    print(f"Paragraphs: {doc.ComputeStatistics(4)}")

    print("\n" + "=" * 80)
    print("PAGE SETUP")
    print("=" * 80)
    ps = doc.PageSetup
    # values in points (1pt = 1/72 inch, 28.35pt = 1cm)
    def pt_to_cm(pt): return round(pt / 28.3464567, 2)
    print(f"Top margin:     {pt_to_cm(ps.TopMargin)} cm  ({ps.TopMargin} pt)")
    print(f"Bottom margin:  {pt_to_cm(ps.BottomMargin)} cm  ({ps.BottomMargin} pt)")
    print(f"Left margin:    {pt_to_cm(ps.LeftMargin)} cm  ({ps.LeftMargin} pt)")
    print(f"Right margin:   {pt_to_cm(ps.RightMargin)} cm  ({ps.RightMargin} pt)")
    print(f"Gutter:         {pt_to_cm(ps.Gutter)} cm")
    print(f"Header offset:  {pt_to_cm(ps.HeaderDistance)} cm")
    print(f"Footer offset:  {pt_to_cm(ps.FooterDistance)} cm")
    print(f"Page width:     {pt_to_cm(ps.PageWidth)} cm")
    print(f"Page height:    {pt_to_cm(ps.PageHeight)} cm")
    print(f"Mirror margins: {ps.MirrorMargins}")
    print(f"Orientation:    {'Portrait' if ps.Orientation == 0 else 'Landscape'}")

    print("\n" + "=" * 80)
    print("STYLES USED (selected — heading / body / abstract)")
    print("=" * 80)
    style_names_of_interest = [
        "Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4",
        "Title", "Subtitle", "Abstract", "Caption", "Body Text", "TOC 1", "TOC 2"
    ]
    for s in doc.Styles:
        try:
            if s.NameLocal in style_names_of_interest:
                f = s.Font
                pf = s.ParagraphFormat
                print(f"\n[{s.NameLocal}]")
                print(f"  Font: {f.Name} | Size: {f.Size}pt | Bold: {f.Bold} | Italic: {f.Italic} | Color: {f.ColorIndex}")
                print(f"  Line spacing: {pf.LineSpacing}pt (rule={pf.LineSpacingRule}) | "
                      f"Space before: {pf.SpaceBefore}pt | Space after: {pf.SpaceAfter}pt")
                print(f"  Alignment: {pf.Alignment} (0=Left,1=Center,2=Right,3=Justify)")
        except Exception as e:
            pass

    print("\n" + "=" * 80)
    print("FULL PARAGRAPH DUMP — text + style + font + size")
    print("=" * 80)
    for i, para in enumerate(doc.Paragraphs, 1):
        text = para.Range.Text.replace("\r", "").replace("\x07", "[TAB]").strip()
        style = para.Style.NameLocal
        font_name = para.Range.Font.Name
        font_size = para.Range.Font.Size
        bold = para.Range.Font.Bold
        italic = para.Range.Font.Italic
        if text:
            preview = text[:200] + ("..." if len(text) > 200 else "")
            print(f"\n[{i:03d}] style='{style}' | font={font_name} {font_size}pt | B={bold} I={italic}")
            print(f"      {preview}")
        else:
            print(f"\n[{i:03d}] (empty) style='{style}'")

    doc.Close(SaveChanges=False)
finally:
    word.Quit()
