# -*- coding: utf-8 -*-
"""
==============================================================
alifbo-sssr-ocr.py - Sariq va Qadimgi Kitoblarni Wordga O'tkazuvchi
Formulalarni saqlovchi, sahifalarni saralovchi va kirilldan lotinga o'tkazuvchi
Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
==============================================================
"""

import os
import sys
import re
import traceback

# Kirill -> Lotin
CYR_TO_LAT = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo',
    'ж': 'j', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'x', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sh',
    'ъ': "'", 'ы': 'i', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
    'ў': "o'", 'ғ': "g'", 'қ': 'q', 'ҳ': 'h',
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'Yo',
    'Ж': 'J', 'З': 'Z', 'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M',
    'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U',
    'Ф': 'F', 'Х': 'X', 'Ц': 'Ts', 'Ч': 'Ch', 'SH': 'Sh', 'Щ': 'Sh',
    'Ъ': "'", 'Ы': 'I', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',
    'Ў': "O'", 'Ғ': "G'", 'Қ': 'Q', 'Ҳ': 'H'
}

MATH_SYMBOLS = set("=+-*/^<≤>≥≠≈±∑∫∬∭∮√∝∞∂∇∈∉⊂⊆∪∩∧∨¬⇒⇔∀∃θλαβγδεζηικμνξπρστυφχψωΔΣΩ")
MATH_FUNCS = {"sin", "cos", "tg", "tan", "ctg", "cot", "sec", "cosec", "ln", "lg", "log", "exp", "lim", "max", "min", "det", "arg"}

def is_math_token(token):
    t_clean = token.strip("(),;:.[]{}")
    if not t_clean:
        return False
    if any(c in MATH_SYMBOLS for c in token):
        return True
    if t_clean.lower() in MATH_FUNCS:
        return True
    if re.match(r'^[a-zA-Z][0-9a-zA-Z_^{}\+\-\*/=]*$', t_clean) and len(t_clean) <= 6:
        return True
    return False

def transliterate_text_preserving_math(text):
    """Formulalarga teginmasdan matnni lotinga o'tkazadi"""
    lines = text.split('\n')
    out_lines = []
    for line in lines:
        stripped = line.strip()
        if any(sym in stripped for sym in "=<>≤≥∑∫√"):
            words = line.split(' ')
            line_parts = []
            for w in words:
                if is_math_token(w):
                    line_parts.append(w)
                else:
                    converted = "".join(CYR_TO_LAT.get(ch, ch) for ch in w)
                    line_parts.append(converted)
            out_lines.append(" ".join(line_parts))
        else:
            converted = "".join(CYR_TO_LAT.get(ch, ch) for ch in line)
            out_lines.append(converted)
    return "\n".join(out_lines)

def parse_page_range(range_str, total_pages):
    """
    '1-100', '1, 3, 5-10', '15-' kabi formatlarni tahlil qiladi (1-based index)
    0-based indekslar ro'yxatini qaytaradi
    """
    if not range_str or not range_str.strip() or range_str.strip().lower() in ['all', 'barchasi', 'hammasi', '*']:
        return list(range(total_pages))

    pages = set()
    parts = [p.strip() for p in range_str.split(',') if p.strip()]
    for part in parts:
        if '-' in part:
            seg = part.split('-')
            start_str = seg[0].strip()
            end_str = seg[1].strip() if len(seg) > 1 else ""

            start = int(start_str) if start_str.isdigit() else 1
            end = int(end_str) if end_str.isdigit() else total_pages

            start = max(1, min(start, total_pages))
            end = max(1, min(end, total_pages))
            if start > end:
                start, end = end, start

            for p in range(start, end + 1):
                pages.add(p - 1)
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p - 1)

    return sorted(list(pages)) if pages else list(range(total_pages))

def extract_pdf_pages_fitz(src_pdf, target_pdf, page_indices):
    """PyMuPDF yordamida faqat kerakli sahifalarni ajratib yangi PDF yaratadi"""
    import fitz
    doc = fitz.open(src_pdf)
    new_doc = fitz.open()
    for idx in page_indices:
        if 0 <= idx < len(doc):
            new_doc.insert_pdf(doc, from_page=idx, to_page=idx)
    new_doc.save(target_pdf)
    new_doc.close()
    doc.close()
    return target_pdf

def convert_pdf_to_word_native(pdf_path, out_docx, page_range=""):
    """
    PDF faylni Word orqali to'g'ridan-to'g'ri va xavfsiz (reflow bilan) Wordga o'tkazish.
    Sahifalar oralig'i berilgan bo'lsa, avval PyMuPDF bilan ajratib olinadi.
    """
    temp_pdf = None
    import fitz
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    doc.close()

    indices = parse_page_range(page_range, total_pages)

    # Agar sahifalar kesilishi kerak bo'lsa
    if len(indices) < total_pages:
        temp_pdf = os.path.join(os.path.dirname(out_docx), f"_temp_pages_{os.getpid()}.pdf")
        extract_pdf_pages_fitz(pdf_path, temp_pdf, indices)
        process_pdf = temp_pdf
    else:
        process_pdf = pdf_path

    # pdf2docx yoki Word COM orqali o'tkazamiz
    try:
        from pdf2docx import Converter
        cv = Converter(process_pdf)
        cv.convert(out_docx)
        cv.close()
        return True, len(indices), total_pages
    except Exception as e_pdf2docx:
        # Fallback: Word COM
        import win32com.client
        word = None
        w_doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            w_doc = word.Documents.Open(os.path.abspath(process_pdf), False, False, False)
            w_doc.SaveAs2(os.path.abspath(out_docx), 16) # wdFormatXMLDocument = 16
            w_doc.Close(False)
            w_doc = None
            word.Quit()
            word = None
            return True, len(indices), total_pages
        except Exception as e_word:
            if w_doc:
                try: w_doc.Close(False)
                except: pass
            if word:
                try: word.Quit()
                except: pass
            raise Exception(f"Word va pdf2docx ikkisi ham xato berdi: {e_word}")
    finally:
        if temp_pdf and os.path.exists(temp_pdf):
            try: os.remove(temp_pdf)
            except: pass

def post_process_docx_latin(docx_path):
    """Tayyor .docx hujjatidagi matnlarni formulalarga teginmasdan kirilldan lotinga o'tkazish"""
    from docx import Document
    doc = Document(docx_path)
    for p in doc.paragraphs:
        if p.text and any('\u0400' <= c <= '\u04FF' for c in p.text):
            for run in p.runs:
                if run.text and any('\u0400' <= c <= '\u04FF' for c in run.text):
                    run.text = transliterate_text_preserving_math(run.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text and any('\u0400' <= c <= '\u04FF' for c in p.text):
                        for run in p.runs:
                            if run.text and any('\u0400' <= c <= '\u04FF' for c in run.text):
                                run.text = transliterate_text_preserving_math(run.text)
    doc.save(docx_path)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Foydalanish: python alifbo-sssr-ocr.py <fayl_yo'li> [sahifalar] [lotin:1/0]")
        sys.exit(1)

    src_file = sys.argv[1]
    pages_arg = sys.argv[2] if len(sys.argv) > 2 else ""
    to_latin = sys.argv[3] != "0" if len(sys.argv) > 3 else True

    dot_pos = src_file.rfind('.')
    out_docx = (src_file[:dot_pos] if dot_pos > 0 else src_file) + "_Word.docx"

    print(f"Boshlandi: {os.path.basename(src_file)}")
    print(f"Sahifalar: {pages_arg if pages_arg else 'Barchasi'}")

    ext = os.path.splitext(src_file)[1].lower()
    if ext == ".pdf":
        ok, proc_count, tot_count = convert_pdf_to_word_native(src_file, out_docx, pages_arg)
        print(f"O'tkazildi: {proc_count}/{tot_count} sahifa Word (.docx) ga aylantirildi.")
        if to_latin:
            print("Formulalar saqlanib, kirill matn lotin alifbosiga o'tkazilmoqda...")
            post_process_docx_latin(out_docx)
            print("Lotinlashtirish yakunlandi.")
    else:
        # Rasm fayl (JPG, PNG)
        from docx import Document
        from PIL import Image, ImageEnhance, ImageFilter
        img = Image.open(src_file).convert('L')
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.2)
        cleaned_path = src_file + "_cleaned.png"
        img.save(cleaned_path)
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("Qadimgi kitob sahifasi (Tozalangan va oqartirilgan):")
        r.bold = True
        doc.add_picture(cleaned_path)
        doc.save(out_docx)
        if os.path.exists(cleaned_path):
            try: os.remove(cleaned_path)
            except: pass
        print("Rasm tozalandi va Wordga saqlandi.")

    print("TAYYOR: " + out_docx)
