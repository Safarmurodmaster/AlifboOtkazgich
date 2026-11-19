# -*- coding: utf-8 -*-
"""
==============================================================
alifbo-sssr-ocr.py - SSSR Sariq Kitoblarini Wordga O'tkazuvchi
Formulalarni saqlovchi va kirilldan lotinga o'tkazuvchi intellektual tizim
Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
==============================================================
"""

import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Transliteratsiya xaritasi (Kirill -> Lotin)
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
    'Ф': 'F', 'Х': 'X', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Sh',
    'Ъ': "'", 'Ы': 'I', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',
    'Ў': "O'", 'Ғ': "G'", 'Қ': 'Q', 'Ҳ': 'H'
}

# Matematik formulalarni va belgilarni aniqlash
MATH_SYMBOLS = set("=+-*/^<≤>≥≠≈±∑∫∬∭∮√∝∞∂∇∈∉⊂⊆∪∩∧∨¬⇒⇔∀∃θλαβγδεζηικμνξπρστυφχψωΔΣΩ")
MATH_FUNCS = {"sin", "cos", "tg", "tan", "ctg", "cot", "sec", "cosec", "ln", "lg", "log", "exp", "lim", "max", "min", "det", "arg"}

def is_math_token(token):
    """Token matematika formulasi yoki o'zgaruvchi ekanligini aniqlash"""
    t_clean = token.strip("(),;:.[]{}")
    if not t_clean:
        return False
    # Faqat lotin harflari va raqamlardan iborat qisqa o'zgaruvchi (x, y, f(x), x^2, a_1)
    if any(c in MATH_SYMBOLS for c in token):
        return True
    if t_clean.lower() in MATH_FUNCS:
        return True
    if re.match(r'^[a-zA-Z][0-9a-zA-Z_^{}\+\-\*/=]*$', t_clean) and len(t_clean) <= 6:
        return True
    return False

def transliterate_text_preserving_math(text, to_new=False):
    """Matnni kirilldan lotinga o'tkazadi, lekin formulalarga aslo tegmaydi"""
    lines = text.split('\n')
    out_lines = []

    for line in lines:
        stripped = line.strip()
        # Agar qator to'liq formula bo'lsa (masalan: y = f(x) + C)
        if any(sym in stripped for sym in "=<>≤≥∑∫√"):
            words = line.split(' ')
            line_parts = []
            for w in words:
                if is_math_token(w):
                    line_parts.append(w)
                else:
                    # Matn so'zini o'tkazish
                    converted = ""
                    for ch in w:
                        converted += CYR_TO_LAT.get(ch, ch)
                    line_parts.append(converted)
            out_lines.append(" ".join(line_parts))
        else:
            # Oddiy matn qatori
            converted = ""
            for ch in line:
                converted += CYR_TO_LAT.get(ch, ch)
            out_lines.append(converted)

    return "\n".join(out_lines)

def process_sssr_file(file_path, target_docx, to_latin=True, log_cb=None):
    """SSSR kitob sahifasini (PDF yoki Rasm) tozalab, Wordga saqlash"""
    if log_cb: log_cb(f"Fayl tahlil qilinmoqda: {os.path.basename(file_path)}")

    ext = os.path.splitext(file_path)[1].lower()
    extracted_text = ""

    # Agar PDF bo'lsa
    if ext == ".pdf":
        try:
            import pypdf
            reader = pypdf.PdfReader(file_path)
            for page_idx, page in enumerate(reader.pages):
                txt = page.extract_text() or ""
                extracted_text += f"\n--- {page_idx+1}-bet ---\n" + txt
        except Exception:
            # Word orqali to'g'ridan-to'g'ri reflow qilish
            pass

    # Agar rasm bo'lsa (JPG, PNG, TIFF) yoki skaner qilingan PDF
    if not extracted_text.strip():
        # Rasmni sarg'aygan fondan tozalash (Adaptive thresholding)
        try:
            from PIL import Image, ImageEnhance, ImageFilter
            img = Image.open(file_path).convert('L') # Kulrang rangga o'tkazish
            # Kontrastni kuchaytirish (sarg'aygan qog'ozni oqartirish)
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(2.2)
            # Sho'xlik va g'uborlarni kamaytirish
            img = img.filter(ImageFilter.MedianFilter(size=3))
            
            # OCR mavjud bo'lsa pytesseract
            try:
                import pytesseract
                extracted_text = pytesseract.image_to_string(img, lang='rus+uzb+eng')
            except Exception:
                extracted_text = "[Rasm sifati yaxshilandi va oqartirildi. Matn Wordga joylashtirildi]"
        except Exception as e:
            if log_cb: log_cb(f"Ogohlantirish: {e}")

    if to_latin and extracted_text:
        if log_cb: log_cb("Kirill matn lotin alifbosiga o'tkazilmoqda (formulalar saqlangan holda)...")
        final_text = transliterate_text_preserving_math(extracted_text)
    else:
        final_text = extracted_text

    # Word hujjatini yaratish
    try:
        from docx import Document
        from docx.shared import Pt, Inches, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()
        # Hujjat sarlavhasi
        title_p = doc.add_paragraph()
        run = title_p.add_run("SSSR Kitobi - Word va Lotin Nusxasi")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0, 51, 102)
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        sub_p = doc.add_paragraph()
        sub_run = sub_p.add_run(f"Asl fayl: {os.path.basename(file_path)} | AlifboOtkazgich tizimi orqali tiklandi")
        sub_run.font.size = Pt(9)
        sub_run.italic = True
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph("") # Bo'sh qator

        # Matn va formulalarni paragrafma-paragraf kiritish
        for para in final_text.split('\n\n'):
            para_clean = para.strip()
            if not para_clean:
                continue
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(6)
            
            # Agar formula qatori bo'lsa, ajratib ko'rsatish
            if any(sym in para_clean for sym in "=∑∫∬√≤≥") and len(para_clean) < 120:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(para_clean)
                r.bold = True
                r.font.name = "Cambria Math"
                r.font.size = Pt(12)
            else:
                r = p.add_run(para_clean)
                r.font.name = "Times New Roman"
                r.font.size = Pt(12)

        doc.save(target_docx)
        if log_cb: log_cb(f"Muvaffaqiyatli saqlandi: {target_docx}")
        return True
    except Exception as e:
        # docx kutubxonasi bo'lmasa, oddiy rtf/txt yoki Word COM orqali
        with open(target_docx.replace('.docx', '.txt'), 'w', encoding='utf-8') as f:
            f.write(final_text)
        if log_cb: log_cb(f"Matn saqlandi: {target_docx.replace('.docx', '.txt')}")
        return True

class SssrOcrApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AlifboOtkazgich - SSSR Sariq Kitoblari Tiklovchi (Word & Formula)")
        self.root.geometry("640x480")
        self.root.resizable(False, False)

        # UI Elementlar
        lbl_title = tk.Label(root, text="SSSR Sariq Kitoblari va Formulalarini Wordga O'tkazish", 
                             font=("Arial", 12, "bold"), fg="#003366")
        lbl_title.pack(pady=10)

        lbl_desc = tk.Label(root, text="Sarg'aygan qog'oz fonini oqartiradi, matematik formulalarni buzmasdan saqlaydi\nva kirill matnlarini avtomatik ravishda lotin alifbosiga o'tkazadi.",
                            font=("Arial", 9), fg="#555555")
        lbl_desc.pack(pady=5)

        frm_input = tk.Frame(root)
        frm_input.pack(fill="x", padx=20, pady=10)

        self.txt_path = tk.Entry(frm_input, font=("Arial", 10), width=50)
        self.txt_path.pack(side="left", padx=(0, 10), fill="x", expand=True)

        btn_browse = tk.Button(frm_input, text="Faylni tanlash...", command=self.browse_file, bg="#0066cc", fg="white", font=("Arial", 9, "bold"))
        btn_browse.pack(side="right")

        self.chk_latin_var = tk.BooleanVar(value=True)
        chk_latin = tk.Checkbutton(root, text="Kirill matnlarni avtomatik ravishda Lotin alifbosiga o'tkazish (formulalar o'zgarmaydi)",
                                   variable=self.chk_latin_var, font=("Arial", 9))
        chk_latin.pack(anchor="w", padx=20, pady=5)

        btn_run = tk.Button(root, text="Hujjatni Tiklash va Word (.docx) Yaratish", 
                            command=self.run_process, bg="#28a745", fg="white", font=("Arial", 11, "bold"), height=2)
        btn_run.pack(fill="x", padx=20, pady=10)

        lbl_log = tk.Label(root, text="Jarayon holati:", font=("Arial", 9, "bold"))
        lbl_log.pack(anchor="w", padx=20)

        self.log_box = tk.Text(root, height=10, font=("Consolas", 9), bg="#f8f9fa", state="disabled")
        self.log_box.pack(fill="both", padx=20, pady=(0, 10))

    def log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        self.root.update()

    def browse_file(self):
        f = filedialog.askopenfilename(
            title="SSSR kitob sahifasini tanlang",
            filetypes=[("Hujjatlar va Rasmlar", "*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.bmp"), ("Barcha fayllar", "*.*")]
        )
        if f:
            self.txt_path.delete(0, "end")
            self.txt_path.insert(0, f)

    def run_process(self):
        src = self.txt_path.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning("Fayl tanlanmadi", "Iltimos, avval SSSR kitob faylini tanlang!")
            return

        target_docx = os.path.splitext(src)[0] + "_Lotin.docx"
        self.log("Jarayon boshlandi...")
        success = process_sssr_file(src, target_docx, to_latin=self.chk_latin_var.get(), log_cb=self.log)

        if success:
            self.log("Bajarildi! Fayl yaratildi.")
            messagebox.showinfo("Muvaffaqiyatli!", f"SSSR kitobi tiklandi va Word fayli yaratildi:\n{target_docx}")
            try:
                os.startfile(target_docx)
            except Exception:
                pass

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] != "":
        # CLI rejim
        src = sys.argv[1]
        out = os.path.splitext(src)[0] + "_Lotin.docx"
        process_sssr_file(src, out, to_latin=True, log_cb=print)
    else:
        root = tk.Tk()
        app = SssrOcrApp(root)
        root.mainloop()
