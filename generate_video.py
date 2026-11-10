import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import math
import os

WIDTH = 720
HEIGHT = 1280
FPS = 30
OUT_PATH = "Alifbo-Reklama.mp4"

# Shriftlar
FONT_NAME = "arial.ttf"
try:
    font_title = ImageFont.truetype("arialbd.ttf", 46)
    font_sub = ImageFont.truetype("arial.ttf", 26)
    font_bold = ImageFont.truetype("arialbd.ttf", 32)
    font_text = ImageFont.truetype("arial.ttf", 28)
    font_card = ImageFont.truetype("arialbd.ttf", 38)
    font_huge = ImageFont.truetype("arialbd.ttf", 64)
    font_btn = ImageFont.truetype("arialbd.ttf", 30)
except Exception:
    font_title = ImageFont.load_default()
    font_sub = font_title
    font_bold = font_title
    font_text = font_title
    font_card = font_title
    font_huge = font_title
    font_btn = font_title

fourcc = cv2.VideoWriter_fourcc(*'mp4v')
out = cv2.VideoWriter(OUT_PATH, fourcc, FPS, (WIDTH, HEIGHT))

def draw_header(draw, progress_idx, total=5, pct=0.0):
    # Stories progress bars
    gap = 8
    total_w = WIDTH - 60
    bar_w = (total_w - (total - 1) * gap) / total
    y = 35
    for i in range(total):
        x1 = 30 + i * (bar_w + gap)
        x2 = x1 + bar_w
        draw.rounded_rectangle([x1, y, x2, y + 6], radius=3, fill=(50, 65, 95))
        if i < progress_idx:
            draw.rounded_rectangle([x1, y, x2, y + 6], radius=3, fill=(96, 165, 250))
        elif i == progress_idx:
            curr_w = bar_w * pct
            draw.rounded_rectangle([x1, y, x1 + curr_w, y + 6], radius=3, fill=(96, 165, 250))

    # Branding
    draw.rounded_rectangle([30, 65, 80, 115], radius=14, fill=(37, 99, 235))
    draw.text((45, 72), "A", font=font_card, fill=(255, 255, 255))
    draw.text((95, 68), "AlifboO'tkazgich", font=font_bold, fill=(255, 255, 255))
    draw.text((95, 104), "Yangi o'zbek alifbosiga 1 bosishda", font=font_sub, fill=(148, 163, 184))

def make_bg(frame_num):
    # Dinamik fon
    img = Image.new("RGB", (WIDTH, HEIGHT), (10, 15, 30))
    draw = ImageDraw.Draw(img)
    # Yumshoq fon to'lqinlari
    for i in range(3):
        cx = int(WIDTH/2 + 100 * math.sin((frame_num + i*40) * 0.03))
        cy = int(HEIGHT/2 + 150 * math.cos((frame_num + i*50) * 0.02))
        r = 350 + i * 80
        draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(15 + i*8, 25 + i*10, 50 + i*15))
    return img, draw

# ================= 1-SAHNA: MUAMMO (0 - 5 soniya = 150 kadr) =================
print("1-sahna render qilinmoqda...")
typed_text = "O'zbekiston Respublikasi o'zgarishlar va yangi qo'shimcha chora-tadbirlar to'g'risida..."
for f in range(150):
    img, draw = make_bg(f)
    draw_header(draw, 0, 5, f / 150.0)

    # Emoji / badge
    draw.rounded_rectangle([WIDTH//2 - 40, 240, WIDTH//2 + 40, 320], radius=20, fill=(45, 20, 30), outline=(244, 63, 94), width=2)
    draw.text((WIDTH//2 - 20, 255), "X", font=font_huge, fill=(244, 63, 94))

    draw.rounded_rectangle([WIDTH//2 - 140, 345, WIDTH//2 + 140, 385], radius=15, fill=(76, 20, 35))
    draw.text((WIDTH//2 - 110, 352), "TANISH HOLATMI?", font=font_sub, fill=(253, 164, 175))

    draw.text((60, 420), "Yuzlab bet hujjatni yangi\nalifboga o'tkazish kerakmi?", font=font_title, fill=(255, 255, 255), align="center")
    draw.text((80, 550), "Har bitta o', g', sh, ch larni qo'lda\nqidirib, tuzatishga soatlab vaqt ketyaptimi?", font=font_text, fill=(148, 163, 184), align="center")

    # Typing qutisi
    draw.rounded_rectangle([50, 680, WIDTH - 50, 920], radius=24, fill=(15, 23, 42), outline=(51, 65, 85), width=2)
    draw.text((75, 705), "Qo'lda tahrirlash (Word):", font=font_sub, fill=(244, 63, 94))

    chars_to_show = int(len(typed_text) * min(1.0, f / 110.0))
    current_text = typed_text[:chars_to_show]
    cursor = "|" if (f // 8) % 2 == 0 else ""
    
    # Qatorlarga bo'lish
    line1 = current_text[:38]
    line2 = current_text[38:]
    draw.text((75, 760), line1 + (cursor if not line2 else ""), font=font_text, fill=(226, 232, 240))
    if line2:
        draw.text((75, 810), line2 + cursor, font=font_text, fill=(226, 232, 240))

    frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    out.write(frame)

# ================= 2-SAHNA: YECHIM (5 - 10 soniya = 150 kadr) =================
print("2-sahna render qilinmoqda...")
for f in range(150):
    img, draw = make_bg(150 + f)
    draw_header(draw, 1, 5, f / 150.0)

    # Katta chaqmoq
    draw.rounded_rectangle([WIDTH//2 - 60, 220, WIDTH//2 + 60, 340], radius=30, fill=(37, 99, 235), outline=(96, 165, 250), width=3)
    draw.text((WIDTH//2 - 22, 245), "!", font=font_huge, fill=(255, 255, 255))

    draw.rounded_rectangle([WIDTH//2 - 130, 370, WIDTH//2 + 130, 410], radius=15, fill=(30, 58, 138))
    draw.text((WIDTH//2 - 105, 377), "ENDI QIYNALMANG!", font=font_sub, fill=(147, 197, 253))

    draw.text((90, 440), "AlifboO'tkazgich", font=font_huge, fill=(96, 165, 250))
    draw.text((60, 530), "Birgina tugma bilan butun hujjatni\n1 soniyada yangi alifboga o'tkazing!", font=font_bold, fill=(255, 255, 255), align="center")

    # Kartochkalar
    y_card = 680
    draw.rounded_rectangle([60, y_card, WIDTH - 60, y_card + 90], radius=20, fill=(30, 41, 59), outline=(71, 85, 105), width=2)
    draw.text((100, y_card + 26), "o'   va   g'", font=font_card, fill=(148, 163, 184))
    draw.text((330, y_card + 26), "==>", font=font_bold, fill=(96, 165, 250))
    draw.text((450, y_card + 26), "ö   va   ğ", font=font_card, fill=(52, 211, 153))

    y_card = 795
    draw.rounded_rectangle([60, y_card, WIDTH - 60, y_card + 90], radius=20, fill=(30, 41, 59), outline=(71, 85, 105), width=2)
    draw.text((100, y_card + 26), "sh   va   ch", font=font_card, fill=(148, 163, 184))
    draw.text((330, y_card + 26), "==>", font=font_bold, fill=(96, 165, 250))
    draw.text((450, y_card + 26), "ş   va   ç", font=font_card, fill=(52, 211, 153))

    frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    out.write(frame)

# ================= 3-SAHNA: WORD ICHIDA JONLI ISHLASHI (10 - 16 soniya = 180 kadr) =================
print("3-sahna render qilinmoqda...")
for f in range(180):
    img, draw = make_bg(300 + f)
    draw_header(draw, 2, 5, f / 180.0)

    draw.rounded_rectangle([WIDTH//2 - 140, 220, WIDTH//2 + 140, 260], radius=15, fill=(49, 46, 129))
    draw.text((WIDTH//2 - 115, 227), "WORD ICHIDA 1 BOSISHDA", font=font_sub, fill=(199, 210, 254))

    draw.text((50, 280), "Hech qayerga ko'chirish shart emas!", font=font_title, fill=(255, 255, 255))
    draw.text((70, 345), "To'g'ridan-to'g'ri Word ichida bosing:", font=font_text, fill=(148, 163, 184))

    # Word interfeysi simulyatsiyasi
    wy = 420
    draw.rounded_rectangle([50, wy, WIDTH - 50, wy + 480], radius=24, fill=(15, 23, 42), outline=(51, 65, 85), width=2)
    # Word top bar
    draw.rounded_rectangle([50, wy, WIDTH - 50, wy + 70], radius=24, fill=(30, 41, 59))
    draw.text((75, wy + 20), "Word", font=font_bold, fill=(96, 165, 250))
    draw.rounded_rectangle([180, wy + 14, 280, wy + 56], radius=10, fill=(37, 99, 235))
    draw.text((195, wy + 20), "Alifbo", font=font_bold, fill=(255, 255, 255))

    # Tugma (f > 60 bo'lganda bosiladi)
    btn_clicked = (f >= 60)
    btn_color = (16, 185, 129) if btn_clicked else (5, 150, 105)
    draw.rounded_rectangle([WIDTH - 250, wy + 14, WIDTH - 75, wy + 56], radius=10, fill=btn_color)
    btn_label = "Tayyor! OK" if btn_clicked else "▶ Yangi ->"
    draw.text((WIDTH - 235, wy + 20), btn_label, font=font_btn, fill=(255, 255, 255))

    # Matn qismi
    if not btn_clicked:
        status_txt = "Holati: Eski alifbo (o', g', sh, ch)"
        status_col = (251, 191, 36)
        text_line1 = "O'zbekiston taraqqiyoti yo'lida yangi"
        text_line2 = "qo'shimcha sharoitlar va chora-tadbirlar..."
        body_col = (203, 213, 225)
    else:
        status_txt = "Holati: Yangi alifboga o'tdi! (100% to'g'ri)"
        status_col = (52, 211, 153)
        text_line1 = "Özbekiston taraqqiyoti yölida yangi"
        text_line2 = "qöşimça şaroitlar va çora-tadbirlar..."
        body_col = (52, 211, 153)

    draw.text((75, wy + 110), status_txt, font=font_sub, fill=status_col)
    draw.text((75, wy + 170), text_line1, font=font_bold, fill=body_col)
    draw.text((75, wy + 230), text_line2, font=font_bold, fill=body_col)

    # Qulaylik eslatmasi
    draw.rounded_rectangle([70, wy + 350, WIDTH - 70, wy + 440], radius=16, fill=(15, 30, 50))
    draw.text((95, wy + 375), "Format, shrift va jadvallar 100% saqlanadi!", font=font_bold, fill=(56, 189, 248))

    frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    out.write(frame)

# ================= 4-SAHNA: HAMMA DASTURLARDA (16 - 21 soniya = 150 kadr) =================
print("4-sahna render qilinmoqda...")
for f in range(150):
    img, draw = make_bg(480 + f)
    draw_header(draw, 3, 5, f / 150.0)

    draw.rounded_rectangle([WIDTH//2 - 130, 220, WIDTH//2 + 130, 260], radius=15, fill=(6, 78, 59))
    draw.text((WIDTH//2 - 105, 227), "HAMMA DASTURLARDA", font=font_sub, fill=(110, 231, 183))

    draw.text((70, 280), "Microsoft Office Paketida!", font=font_title, fill=(255, 255, 255))
    draw.text((60, 350), "Qayerda ishlasangiz ham doim yordamchi:", font=font_text, fill=(148, 163, 184))

    # 3 ta programma qutisi
    items = [
        ("W", "Word", "Hujjatlar uchun", (37, 99, 235)),
        ("X", "Excel", "Jadvallar uchun", (16, 185, 129)),
        ("P", "PowerPoint", "Slaydlar uchun", (234, 88, 12))
    ]
    for idx, (letter, name, desc, col) in enumerate(items):
        box_y = 440 + idx * 130
        draw.rounded_rectangle([60, box_y, WIDTH - 60, box_y + 110], radius=20, fill=(30, 41, 59), outline=(71, 85, 105), width=2)
        draw.rounded_rectangle([80, box_y + 18, 150, box_y + 92], radius=16, fill=col)
        draw.text((100, box_y + 28), letter, font=font_card, fill=(255, 255, 255))
        draw.text((170, box_y + 24), name, font=font_bold, fill=(255, 255, 255))
        draw.text((170, box_y + 64), desc, font=font_sub, fill=(148, 163, 184))

    # Hotkey yordamchi
    draw.rounded_rectangle([60, 860, WIDTH - 60, 970], radius=20, fill=(30, 27, 75), outline=(99, 102, 241), width=2)
    draw.text((90, 880), "Telegram, brauzer va boshqa joylarda:", font=font_sub, fill=(199, 210, 254))
    draw.text((90, 920), "Tezkor tugma:  Ctrl + Shift + V", font=font_bold, fill=(255, 255, 255))

    frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    out.write(frame)

# ================= 5-SAHNA: CALL TO ACTION (21 - 27 soniya = 180 kadr) =================
print("5-sahna render qilinmoqda...")
for f in range(180):
    img, draw = make_bg(630 + f)
    draw_header(draw, 4, 5, f / 180.0)

    draw.rounded_rectangle([WIDTH//2 - 50, 220, WIDTH//2 + 50, 320], radius=26, fill=(16, 185, 129))
    draw.text((WIDTH//2 - 25, 235), "OK", font=font_huge, fill=(255, 255, 255))

    draw.text((60, 360), "Vaqtingizni Tejang,\nIshingizni Yengillating!", font=font_title, fill=(255, 255, 255), align="center")
    draw.text((70, 480), "Dasturni bir marta o'rnatasiz va\ndoimiy rohatlanib ishlaysiz!", font=font_text, fill=(148, 163, 184), align="center")

    # Telegram aloqa qutisi
    box_y = 590
    draw.rounded_rectangle([50, box_y, WIDTH - 50, box_y + 320], radius=28, fill=(15, 23, 42), outline=(59, 130, 246), width=3)
    draw.text((WIDTH//2 - 130, box_y + 35), "DASTURNI OLISH UCHUN:", font=font_sub, fill=(147, 197, 253))

    # TG tugmasi
    draw.rounded_rectangle([80, box_y + 90, WIDTH - 80, box_y + 180], radius=20, fill=(37, 99, 235))
    draw.text((WIDTH//2 - 145, box_y + 115), "@Safar_master", font=font_huge, fill=(255, 255, 255))

    draw.text((WIDTH//2 - 160, box_y + 215), "Telefon: +998 93 953 04 73", font=font_bold, fill=(226, 232, 240))
    draw.text((WIDTH//2 - 165, box_y + 265), "Muallif: Fozliddinov Safarmurod", font=font_sub, fill=(148, 163, 184))

    frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    out.write(frame)

out.release()
print(f"Video tayyor bo'ldi -> {OUT_PATH}")
