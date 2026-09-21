#!/usr/bin/env python3
"""IST-HILLTOWN Tempe haftalık performans sunumu."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import nsmap
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt, Emu
from lxml import etree

# --- Palette (retail / Inditex-inspired, not purple/cream defaults) ---
INK = RGBColor(0x1A, 0x1A, 0x1A)
SOFT_INK = RGBColor(0x4A, 0x4A, 0x4A)
MUTED = RGBColor(0x7A, 0x7A, 0x7A)
PAPER = RGBColor(0xF7, 0xF5, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0xC4, 0x5C, 0x26)  # warm terracotta-copper for retail
POSITIVE = RGBColor(0x1F, 0x6B, 0x4A)
NEGATIVE = RGBColor(0xA3, 0x3A, 0x2C)
CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)
RULE = RGBColor(0xE4, 0xDF, 0xD8)
BAR_FILL = RGBColor(0x2C, 0x2C, 0x2C)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def set_run(run, size=18, bold=False, color=INK, font="Calibri"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_textbox(slide, left, top, width, height, text, size=18, bold=False,
                color=INK, align=PP_ALIGN.LEFT, font="Calibri", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def add_bg(slide, color=PAPER):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    # send to back
    spTree = slide.shapes._spTree
    sp = shape._element
    spTree.remove(sp)
    spTree.insert(2, sp)
    return shape


def add_accent_bar(slide, left, top, width=Inches(0.08), height=Inches(0.55)):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()
    return bar


def add_card(slide, left, top, width, height, fill=WHITE):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = fill
    card.line.color.rgb = RULE
    card.line.width = Pt(1)
    # softer corners
    try:
        card.adjustments[0] = 0.08
    except Exception:
        pass
    return card


def add_footer(slide, page, total=7):
    add_textbox(
        slide, Inches(0.6), Inches(7.05), Inches(8), Inches(0.3),
        "IST-HILLTOWN  ·  Tempe  ·  Son 7 Gün  ·  Ladies Ranking",
        size=10, color=MUTED, font="Calibri"
    )
    add_textbox(
        slide, Inches(11.5), Inches(7.05), Inches(1.3), Inches(0.3),
        f"{page} / {total}",
        size=10, color=MUTED, align=PP_ALIGN.RIGHT, font="Calibri"
    )


def metric_card(slide, left, top, width, height, label, value, delta=None, delta_positive=True):
    add_card(slide, left, top, width, height)
    add_textbox(slide, left + Inches(0.22), top + Inches(0.22), width - Inches(0.4), Inches(0.3),
                label.upper(), size=11, color=MUTED, bold=True)
    add_textbox(slide, left + Inches(0.22), top + Inches(0.55), width - Inches(0.4), Inches(0.55),
                value, size=28, color=INK, bold=True, font="Calibri")
    if delta:
        dcolor = POSITIVE if delta_positive else NEGATIVE
        add_textbox(slide, left + Inches(0.22), top + Inches(1.15), width - Inches(0.4), Inches(0.35),
                    delta, size=13, color=dcolor, bold=True)


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    # ========== 1. KAPAK ==========
    s = prs.slides.add_slide(blank)
    add_bg(s, INK)
    # left accent strip
    strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.18), SLIDE_H)
    strip.fill.solid()
    strip.fill.fore_color.rgb = ACCENT
    strip.line.fill.background()

    add_textbox(s, Inches(0.9), Inches(1.7), Inches(11), Inches(0.4),
                "IST-HILLTOWN", size=16, color=ACCENT, bold=True, font="Calibri")
    add_textbox(s, Inches(0.9), Inches(2.2), Inches(11), Inches(1.1),
                "TEMPE HAFTALIK\nPERFORMANS", size=44, color=WHITE, bold=True, font="Calibri")
    add_textbox(s, Inches(0.9), Inches(4.5), Inches(10), Inches(0.4),
                "Son 7 gün satış, kategori ve stok özeti", size=18, color=RGBColor(0xB8, 0xB8, 0xB8))
    add_textbox(s, Inches(0.9), Inches(6.4), Inches(10), Inches(0.35),
                "21 Eylül 2026  ·  Ladies Ranking  ·  Units bazlı", size=12, color=MUTED)

    # ========== 2. SATIŞ ÖZETİ ==========
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_accent_bar(s, Inches(0.6), Inches(0.45))
    add_textbox(s, Inches(0.85), Inches(0.4), Inches(10), Inches(0.55),
                "Satış Özeti", size=28, bold=True)
    add_textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.35),
                "Haftalık Tempe toplam performansı", size=14, color=SOFT_INK)

    metric_card(s, Inches(0.6), Inches(1.6), Inches(3.9), Inches(1.7),
                "Total Tempe Satış", "397 uts", "+62 uts  ·  +%18,6", True)
    metric_card(s, Inches(4.7), Inches(1.6), Inches(3.9), Inches(1.7),
                "En Düşük Gün", "Perşembe", "45 uts", False)
    metric_card(s, Inches(8.8), Inches(1.6), Inches(3.9), Inches(1.7),
                "Ana Sürücü", "Basic Tempe", "Adet & TL lideri", True)

    # insight box
    add_card(s, Inches(0.6), Inches(3.7), Inches(12.1), Inches(2.7))
    add_textbox(s, Inches(0.95), Inches(3.95), Inches(11.4), Inches(0.4),
                "Haftalık okuma", size=16, bold=True, color=ACCENT)
    bullets = [
        "Toplam satış 397 uts ile geçen haftaya göre +62 adet (+%18,6) artış gösterdi.",
        "Büyümenin ana kaynağı Basic Tempe; hem adet hem ciroda en yüksek getiriyi sağladı.",
        "Perşembe günü 45 uts ile haftanın en düşük satışı gerçekleşti.",
        "Sebep olarak dağıtılmayan 25’ler ve ertesi güne kalan sepetler öne çıkıyor.",
    ]
    y = 4.45
    for b in bullets:
        add_textbox(s, Inches(0.95), Inches(y), Inches(11.4), Inches(0.4),
                    "▸  " + b, size=14, color=SOFT_INK)
        y += 0.42
    add_footer(s, 2)

    # ========== 3. PERŞEMBE ANALİZİ ==========
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_accent_bar(s, Inches(0.6), Inches(0.45))
    add_textbox(s, Inches(0.85), Inches(0.4), Inches(10), Inches(0.55),
                "Perşembe Günü Düşüşü", size=28, bold=True)
    add_textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.35),
                "Operasyonel etki analizi", size=14, color=SOFT_INK)

    # big number
    add_card(s, Inches(0.6), Inches(1.55), Inches(4.2), Inches(4.7))
    add_textbox(s, Inches(0.95), Inches(2.0), Inches(3.5), Inches(0.35),
                "PERŞEMBE SATIŞ", size=12, color=MUTED, bold=True)
    add_textbox(s, Inches(0.95), Inches(2.5), Inches(3.5), Inches(1.0),
                "45", size=72, bold=True, color=NEGATIVE, font="Calibri")
    add_textbox(s, Inches(0.95), Inches(3.6), Inches(3.5), Inches(0.4),
                "uts", size=22, color=SOFT_INK)
    add_textbox(s, Inches(0.95), Inches(4.4), Inches(3.5), Inches(1.2),
                "Haftanın en düşük satış günü.\nOrtalama günlük seyrin altında kaldı.",
                size=14, color=SOFT_INK)

    add_card(s, Inches(5.1), Inches(1.55), Inches(7.6), Inches(4.7))
    add_textbox(s, Inches(5.5), Inches(1.9), Inches(6.8), Inches(0.4),
                "Olası nedenler", size=18, bold=True, color=ACCENT)

    causes = [
        ("Dağıtılmayan 25’ler", "Kritik bedenlerin mağazaya ulaşmaması satış kaybına yol açtı."),
        ("Ertesi güne kalan sepetler", "Tamamlanamayan satışlar bir sonraki güne sarktı."),
        ("Operasyon aksiyonu", "Teslimat takibi ve sepet kapatma disiplini güçlendirilmeli."),
    ]
    cy = 2.55
    for title, desc in causes:
        add_textbox(s, Inches(5.5), Inches(cy), Inches(6.8), Inches(0.35),
                    title, size=15, bold=True, color=INK)
        add_textbox(s, Inches(5.5), Inches(cy + 0.35), Inches(6.8), Inches(0.55),
                    desc, size=13, color=SOFT_INK)
        cy += 1.05
    add_footer(s, 3)

    # ========== 4. KATEGORİ PERFORMANSI ==========
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_accent_bar(s, Inches(0.6), Inches(0.45))
    add_textbox(s, Inches(0.85), Inches(0.4), Inches(10), Inches(0.55),
                "Kategori Performansı", size=28, bold=True)
    add_textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.35),
                "Buyer bazında adet ve TL farkı", size=14, color=SOFT_INK)

    # table-like cards
    headers = ["Kategori", "Adet (uts)", "TL", "Durum"]
    rows = [
        ("Basic Tempe", "En yüksek", "En yüksek", "Lider"),
        ("Bags", "−1", "Nötr / sınırda", "Adette eksi"),
        ("Sneakers", "−4", "−7k ₺", "Adet & TL eksi"),
        ("Diğer Tempe", "Pozitif katkı", "Pozitif katkı", "Destekleyici"),
    ]

    # header row
    x0 = Inches(0.6)
    col_w = [Inches(3.2), Inches(2.8), Inches(3.2), Inches(2.9)]
    y = Inches(1.55)
    # header bar
    hdr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x0, y, Inches(12.1), Inches(0.55))
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = INK
    hdr.line.fill.background()
    x = x0
    for i, h in enumerate(headers):
        add_textbox(s, x + Inches(0.2), y + Inches(0.1), col_w[i] - Inches(0.2), Inches(0.4),
                    h.upper(), size=12, bold=True, color=WHITE)
        x += col_w[i]

    y = Inches(2.2)
    for idx, (cat, adet, tl, durum) in enumerate(rows):
        bg = WHITE if idx % 2 == 0 else RGBColor(0xF0, 0xEC, 0xE6)
        row_bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x0, y, Inches(12.1), Inches(0.7))
        row_bg.fill.solid()
        row_bg.fill.fore_color.rgb = bg
        row_bg.line.color.rgb = RULE
        row_bg.line.width = Pt(0.5)
        vals = [cat, adet, tl, durum]
        colors = [INK, INK, INK, INK]
        if "eksi" in durum.lower() or durum == "Adette eksi":
            colors = [INK, NEGATIVE, NEGATIVE if "−7k" in tl else SOFT_INK, NEGATIVE]
        if durum == "Lider":
            colors = [INK, POSITIVE, POSITIVE, POSITIVE]
        x = x0
        for i, v in enumerate(vals):
            add_textbox(s, x + Inches(0.2), y + Inches(0.18), col_w[i] - Inches(0.2), Inches(0.4),
                        v, size=14, bold=(i == 0 or i == 3), color=colors[i])
            x += col_w[i]
        y += Inches(0.7)

    add_textbox(s, Inches(0.6), Inches(5.3), Inches(12.1), Inches(1.2),
                "Özet: Haftanın totalinde yalnızca Bags (−1) ve Sneakers (−4) adette ekside kaldı. "
                "TL tarafında ise sadece Sneakers −7k ile negatif. Hem adet hem TL’de en yüksek getiri Basic Tempe’den geldi.",
                size=14, color=SOFT_INK)
    add_footer(s, 4)

    # ========== 5. STOK ==========
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_accent_bar(s, Inches(0.6), Inches(0.45))
    add_textbox(s, Inches(0.85), Inches(0.4), Inches(10), Inches(0.55),
                "Stok Durumu", size=28, bold=True)
    add_textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.35),
                "Tempe envanter dağılımı ve aktif oran", size=14, color=SOFT_INK)

    metric_card(s, Inches(0.6), Inches(1.5), Inches(4.0), Inches(1.65),
                "Total Tempe Stok", "1.753 adet", "+413 uts", True)
    metric_card(s, Inches(4.8), Inches(1.5), Inches(4.0), Inches(1.65),
                "Tempe Aktif Oran", "%82", "Hedef üstü aktiflik", True)
    metric_card(s, Inches(9.0), Inches(1.5), Inches(3.7), Inches(1.65),
                "Pasif Stok", "320 uts", "İyileştirme alanı", False)

    # distribution bars
    add_card(s, Inches(0.6), Inches(3.45), Inches(12.1), Inches(3.0))
    add_textbox(s, Inches(0.95), Inches(3.7), Inches(11), Inches(0.35),
                "Stok kırılımı", size=16, bold=True, color=ACCENT)

    parts = [
        ("Shop", 749, ACCENT),
        ("Aktif Stok", 684, BAR_FILL),
        ("Pasif Stok", 320, MUTED),
    ]
    total = 1753
    bar_left = Inches(0.95)
    bar_top = Inches(4.3)
    bar_h = Inches(0.55)
    full_w = Inches(11.4)
    # stacked bar
    x = bar_left
    for name, qty, color in parts:
        w = int(full_w * (qty / total))
        seg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, bar_top, w, bar_h)
        seg.fill.solid()
        seg.fill.fore_color.rgb = color
        seg.line.fill.background()
        x += w

    # legends
    lx = Inches(0.95)
    for name, qty, color in parts:
        sw = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, lx, Inches(5.15), Inches(0.25), Inches(0.25))
        sw.fill.solid()
        sw.fill.fore_color.rgb = color
        sw.line.fill.background()
        pct = round(100 * qty / total)
        add_textbox(s, lx + Inches(0.35), Inches(5.1), Inches(3.2), Inches(0.35),
                    f"{name}: {qty} uts (%{pct})", size=13, color=SOFT_INK, bold=True)
        lx += Inches(3.9)

    add_textbox(s, Inches(0.95), Inches(5.7), Inches(11.4), Inches(0.45),
                "Shop 749 + Aktif 684 + Pasif 320 = 1.753. Aktif oran %82 ile stok sağlığı güçlü; pasif stok aksiyonu fırsat alanı.",
                size=13, color=SOFT_INK)
    add_footer(s, 5)

    # ========== 6. TOP ÜRÜNLER ==========
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_accent_bar(s, Inches(0.6), Inches(0.45))
    add_textbox(s, Inches(0.85), Inches(0.4), Inches(10), Inches(0.55),
                "Ladies Ranking — Top Ürünler", size=28, bold=True)
    add_textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.35),
                "Son 7 gün · Units sıralaması · In Store / Shoes", size=14, color=SOFT_INK)

    products = [
        ("01", "Mary Jane Ante", "13 uts", "₺34.970"),
        ("02", "Slingback Tacón Alto", "11 uts", "₺27.390"),
        ("03", "Zapato Tacón Media", "11 uts", "₺14.190"),
        ("04", "Shopper Asa Edited", "10 uts", "₺26.900"),
        ("05", "Destalonado Print", "10 uts", "₺26.900"),
        ("06", "Destalonado Medio", "10 uts", "₺18.900"),
        ("07", "Bota Hípica", "9 uts", "₺32.310"),
        ("08", "Shopper Metálica", "9 uts", "₺24.210"),
    ]

    # 2x4 grid
    for i, (rank, name, uts, tl) in enumerate(products):
        col = i % 4
        row = i // 4
        left = Inches(0.6) + col * Inches(3.15)
        top = Inches(1.5) + row * Inches(2.45)
        add_card(s, left, top, Inches(3.0), Inches(2.2))
        add_textbox(s, left + Inches(0.2), top + Inches(0.25), Inches(2.5), Inches(0.35),
                    rank, size=20, bold=True, color=ACCENT)
        add_textbox(s, left + Inches(0.2), top + Inches(0.7), Inches(2.6), Inches(0.55),
                    name, size=14, bold=True, color=INK)
        add_textbox(s, left + Inches(0.2), top + Inches(1.35), Inches(2.6), Inches(0.3),
                    uts, size=16, bold=True, color=POSITIVE)
        add_textbox(s, left + Inches(0.2), top + Inches(1.7), Inches(2.6), Inches(0.3),
                    tl + "  ·  7D", size=12, color=MUTED)
    add_footer(s, 6)

    # ========== 7. SONUÇ / AKSİYON ==========
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_accent_bar(s, Inches(0.6), Inches(0.45))
    add_textbox(s, Inches(0.85), Inches(0.4), Inches(10), Inches(0.55),
                "Sonuç & Aksiyonlar", size=28, bold=True)
    add_textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.35),
                "Haftalık kapanış notları", size=14, color=SOFT_INK)

    actions = [
        ("Büyüme korunsun", "Basic Tempe vitrin ve stok derinliği korunarak +%18,6 momentumu sürdürülmeli."),
        ("Perşembe operasyonu", "25’ beden teslimatları ve sepet kapatma süreci günlük takip listesine alınmalı."),
        ("Sneakers toparlama", "−4 uts / −7k ₺ için ürün görünürlüğü, beden tamlığı ve visual push gözden geçirilmeli."),
        ("Bags izleme", "Adette −1; ciro kırılgan — top shopper’lar için stok/display disiplini sürdürülmeli."),
        ("Pasif stok", "320 uts pasif stokta aktifleştirebilecek ürünler için transfer / floor push planı."),
        ("Aktif oran", "%82 aktif oran güçlü; Shop + Aktif dengesini koruyarak 1.753 stok sağlığı muhafaza edilmeli."),
    ]

    for i, (title, desc) in enumerate(actions):
        col = i % 2
        row = i // 2
        left = Inches(0.6) + col * Inches(6.35)
        top = Inches(1.5) + row * Inches(1.65)
        add_card(s, left, top, Inches(6.1), Inches(1.45))
        num = s.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.25), top + Inches(0.4), Inches(0.45), Inches(0.45))
        num.fill.solid()
        num.fill.fore_color.rgb = ACCENT
        num.line.fill.background()
        add_textbox(s, left + Inches(0.25), top + Inches(0.45), Inches(0.45), Inches(0.4),
                    str(i + 1), size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, left + Inches(0.9), top + Inches(0.3), Inches(4.9), Inches(0.35),
                    title, size=15, bold=True, color=INK)
        add_textbox(s, left + Inches(0.9), top + Inches(0.7), Inches(4.9), Inches(0.6),
                    desc, size=12, color=SOFT_INK)
    add_footer(s, 7)

    out = "/workspace/presentations/IST-HILLTOWN_Tempe_Haftalik_Performans.pptx"
    prs.save(out)
    # also copy to artifacts for easy download
    import shutil
    shutil.copy(out, "/opt/cursor/artifacts/IST-HILLTOWN_Tempe_Haftalik_Performans.pptx")
    print(out)
    return out


if __name__ == "__main__":
    build()
