import sys
import os
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    base = r"d:\Perkuliahan\Kelass\SKB"
    data_dir = os.path.join(base, "04_INTERNAL_TOOLS_DAN_DATA", "03_data_dan_aset")
    with open(os.path.join(data_dir, "config_bakmi_mimu.json"), "r", encoding="utf-8") as f:
        cfg = json.load(f)
    with open(os.path.join(data_dir, "metrics_bakmi_mimu.json"), "r", encoding="utf-8") as f:
        met = json.load(f)

    logo_path = os.path.join(data_dir, "logo_ukrida.png")
    layout_path = os.path.join(data_dir, "denah_tata_letak_ruko.png")
    out_pptx_1 = os.path.join(base, "01_TUGAS_FINAL_BAKMI_MIMU", "Presentasi_SKB_Bakmi_Mimu_Carina_Sayang.pptx")
    out_pptx_2 = os.path.join(base, "naskah utama", "Presentasi_SKB_Bakmi_Mimu_Carina_Sayang.pptx")

    prs = Presentation()
    # 16:9 Widescreen dimensions (13.333 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # PALET WARNA FORMAL KAMPUS FEB UKRIDA (CLEAN EXECUTIVE LIGHT THEME)
    # Sangat kompatibel 100% dengan Google Slides, proyektor kelas, & laptop
    # =========================================================================
    BG_LIGHT = RGBColor(250, 251, 253)       # Clean Soft White (#FAFBFD)
    CARD_BG = RGBColor(255, 255, 255)        # Pure White Card (#FFFFFF)
    CARD_BG_SOFT = RGBColor(241, 245, 249)   # Light Gray Card (#F1F5F9)
    NAVY_PRIMARY = RGBColor(15, 41, 77)      # Deep UKRIDA Navy (#0F294D)
    NAVY_LIGHT = RGBColor(30, 58, 138)       # Royal Navy Accent (#1E3A8A)
    GOLD_ACCENT = RGBColor(194, 120, 3)      # Warm Amber Gold (#C27803)
    GOLD_BG_LIGHT = RGBColor(254, 243, 199)  # Soft Gold Badge (#FEF3C7)
    GREEN_ACCENT = RGBColor(5, 150, 105)     # Feasibility Green (#059669)
    GREEN_BG_LIGHT = RGBColor(236, 253, 245) # Soft Green Badge (#ECFDF5)
    TEXT_MAIN = RGBColor(15, 23, 42)         # Deep Charcoal (#0F172A - Super Sharp)
    TEXT_MUTED = RGBColor(71, 85, 105)       # Slate Gray Body (#475569)
    BORDER_MUTED = RGBColor(203, 213, 225)   # Border Slate (#CBD5E1)
    BORDER_NAVY = RGBColor(147, 197, 253)    # Soft Blue Border (#93C5FD)
    WHITE = RGBColor(255, 255, 255)

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_LIGHT
        bg.line.fill.background()
        
        # Top decorative thin line (UKRIDA Navy + Gold)
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = NAVY_PRIMARY
        top_bar.line.fill.background()
        return bg

    def add_header(slide, title_text, subtitle_text=None, tag="STUDI KELAYAKAN BISNIS", slide_idx=1):
        # Category Tag Badge (Top Left Pill)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.6), Inches(0.35))
        pill.fill.solid()
        pill.fill.fore_color.rgb = GOLD_BG_LIGHT
        pill.line.color.rgb = GOLD_ACCENT
        pill.line.width = Pt(0.75)
        tf_pill = pill.text_frame
        tf_pill.word_wrap = True
        tf_pill.margin_left = tf_pill.margin_right = tf_pill.margin_top = tf_pill.margin_bottom = 0
        p_pill = tf_pill.paragraphs[0]
        p_pill.alignment = PP_ALIGN.CENTER
        p_pill.text = f"• {tag.upper()}"
        p_pill.font.name = "Times New Roman"
        p_pill.font.size = Pt(10)
        p_pill.font.bold = True
        p_pill.font.color.rgb = GOLD_ACCENT

        # Slide Number Badge (Top Right)
        s_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.2), Inches(0.35), Inches(1.33), Inches(0.35))
        s_badge.fill.solid()
        s_badge.fill.fore_color.rgb = CARD_BG_SOFT
        s_badge.line.color.rgb = BORDER_MUTED
        s_badge.line.width = Pt(0.75)
        tf_sb = s_badge.text_frame
        tf_sb.word_wrap = True
        tf_sb.margin_left = tf_sb.margin_right = tf_sb.margin_top = tf_sb.margin_bottom = 0
        p_sb = tf_sb.paragraphs[0]
        p_sb.alignment = PP_ALIGN.CENTER
        p_sb.text = f"SLIDE {slide_idx:02d} / 22"
        p_sb.font.name = "Times New Roman"
        p_sb.font.size = Pt(9.5)
        p_sb.font.bold = True
        p_sb.font.color.rgb = TEXT_MUTED

        # Slide Title Box
        tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.78), Inches(11.733), Inches(0.75))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(21)
        p_t.font.bold = True
        p_t.font.color.rgb = NAVY_PRIMARY

        if subtitle_text:
            p_sub = tf_t.add_paragraph()
            p_sub.text = subtitle_text
            p_sub.font.name = "Times New Roman"
            p_sub.font.size = Pt(11)
            p_sub.font.color.rgb = TEXT_MUTED

        # Bottom thin dividing rule
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.58), Inches(11.733), Inches(0.02))
        div.fill.solid()
        div.fill.fore_color.rgb = BORDER_MUTED
        div.line.fill.background()

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_MUTED, border_width=1.0):
        """Card shape yang juga bertindak langsung sebagai container text agar tidak meleset di Google Slides"""
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(border_width)
        else:
            card.line.fill.background()
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.18)
        tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.16)
        tf.margin_bottom = Inches(0.16)
        return card, tf

    def style_table_cell(cell, text, font_size=10, bold=False, color=TEXT_MAIN, align=PP_ALIGN.LEFT, bg_color=None):
        cell.text = text
        tf = cell.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)
        tf.margin_top = Inches(0.05)
        tf.margin_bottom = Inches(0.05)
        p = tf.paragraphs[0]
        p.alignment = align
        p.font.name = "Times New Roman"
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.color.rgb = color
        if bg_color:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color

    # =========================================================================
    # SLIDE 1: COVER PRESENTASI EKSEKUTIF (Clean Formal UKRIDA)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Master Frame Box
    card1, tf1 = add_card(s1, Inches(0.8), Inches(0.5), Inches(11.733), Inches(6.5), bg_color=CARD_BG, border_color=BORDER_NAVY, border_width=1.5)

    # UKRIDA Logo at top center
    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(6.066), Inches(0.75), Inches(1.2), Inches(1.2))

    # Text content on Cover Box
    tb_cov = s1.shapes.add_textbox(Inches(1.0), Inches(2.05), Inches(11.333), Inches(2.5))
    tf_cov = tb_cov.text_frame
    tf_cov.word_wrap = True
    tf_cov.margin_left = tf_cov.margin_right = tf_cov.margin_top = tf_cov.margin_bottom = 0

    p_c1 = tf_cov.paragraphs[0]
    p_c1.alignment = PP_ALIGN.CENTER
    p_c1.text = "FAKULTAS EKONOMI & BISNIS  •  UNIVERSITAS KRISTEN KRIDA WACANA"
    p_c1.font.name = "Times New Roman"
    p_c1.font.size = Pt(11)
    p_c1.font.bold = True
    p_c1.font.color.rgb = GOLD_ACCENT

    p_c2 = tf_cov.add_paragraph()
    p_c2.alignment = PP_ALIGN.CENTER
    p_c2.text = "LAPORAN PRESENTASI STUDI KELAYAKAN BISNIS (SKB)"
    p_c2.font.name = "Times New Roman"
    p_c2.font.size = Pt(14)
    p_c2.font.color.rgb = TEXT_MUTED

    p_c3 = tf_cov.add_paragraph()
    p_c3.alignment = PP_ALIGN.CENTER
    p_c3.text = f"“ {cfg['business_name'].upper()} ”"
    p_c3.font.name = "Times New Roman"
    p_c3.font.size = Pt(30)
    p_c3.font.bold = True
    p_c3.font.color.rgb = NAVY_PRIMARY

    p_c4 = tf_cov.add_paragraph()
    p_c4.alignment = PP_ALIGN.CENTER
    p_c4.text = f"- {cfg['tagline']} -"
    p_c4.font.name = "Times New Roman"
    p_c4.font.size = Pt(12.5)
    p_c4.font.italic = True
    p_c4.font.color.rgb = GOLD_ACCENT

    # 5 Mahasiswa Cards Grid
    authors = [
        ("Arthur Reezan", "312023002"),
        ("Jennese Putra Alamsyah Sukadi", "312023033"),
        ("Valendrik Dwiputra Wirawan", "312023013"),
        ("Affandy", "312023075"),
        ("Steven Putra Tjhin", "312023015")
    ]
    card_w = Inches(2.2)
    card_gap = Inches(0.12)
    start_x = Inches(1.15)
    card_y = Inches(4.75)
    card_h = Inches(1.15)

    for i, (name, nim) in enumerate(authors):
        x = start_x + i * (card_w + card_gap)
        c_m, tf_m = add_card(s1, x, card_y, card_w, card_h, bg_color=CARD_BG_SOFT, border_color=BORDER_MUTED)
        
        pm1 = tf_m.paragraphs[0]
        pm1.alignment = PP_ALIGN.CENTER
        pm1.text = f"Mahasiswa {i+1}"
        pm1.font.name = "Times New Roman"
        pm1.font.size = Pt(9)
        pm1.font.bold = True
        pm1.font.color.rgb = GOLD_ACCENT

        pm2 = tf_m.add_paragraph()
        pm2.alignment = PP_ALIGN.CENTER
        pm2.text = name
        pm2.font.name = "Times New Roman"
        pm2.font.size = Pt(8.5)
        pm2.font.bold = True
        pm2.font.color.rgb = NAVY_PRIMARY

        pm3 = tf_m.add_paragraph()
        pm3.alignment = PP_ALIGN.CENTER
        pm3.text = f"NIM: {nim}"
        pm3.font.name = "Times New Roman"
        pm3.font.size = Pt(9)
        pm3.font.color.rgb = TEXT_MUTED

    # Footer note
    tb_foot = s1.shapes.add_textbox(Inches(1.0), Inches(6.1), Inches(11.333), Inches(0.4))
    tf_foot = tb_foot.text_frame
    tf_foot.word_wrap = True
    p_f = tf_foot.paragraphs[0]
    p_f.alignment = PP_ALIGN.CENTER
    p_f.text = "Program Studi Sarjana Manajemen  |  Konsentrasi Studi Kelayakan Bisnis  |  Tahun Akademik 2024"
    p_f.font.name = "Times New Roman"
    p_f.font.size = Pt(10)
    p_f.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 2: DAFTAR ISI & AGENDA KAJIAN
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Daftar Isi & Sistematika Kajian", "Ringkasan 6 bab utama studi kelayakan bisnis sesuai standar FEB UKRIDA", tag="AGENDA PRESENTASI", slide_idx=2)

    toc_items = [
        ("01", "BAB I: PENDAHULUAN & PROFIL", "Perjalanan 1999 Taman Aries -> 2016 Duri Kosambi, basis pelanggan 2019, & ekspansi 2027/28."),
        ("02", "BAB II: ASPEK PASAR & PEMASARAN", "Segmentasi STP, 3 minyak mie, topping ayam putih/babi/casiu, benchmarking, & bauran 7P."),
        ("03", "BAB III: ASPEK MANAJEMEN", "Struktur organisasi 3 staf fungsional, job desc terstruktur, & total alokasi gaji Rp 10 Jt/bln."),
        ("04", "BAB IV: ASPEK TEKNIS & OPERASI", "Tata letak ruko 1 lantai 22 kursi, modal 2016 Rp 62 Jt, open kitchen gerobak, & adonan fresh mingguan."),
        ("05", "BAB V: ASPEK KEUANGAN TERPADU", f"Initial outlay Rp {met['initial_outlay']/1e6:.1f} Jt, proyeksi 5 tahun arus kas, & depresiasi garis lurus 20%."),
        ("06", "BAB VI: KELAYAKAN INVESTASI", f"Uji 4 kriteria: NPV Rp {met['npv']/1e9:.2f} M, IRR > 20%, Payback 0,82 Bulan, & PI {met['pi']:.2f}.")
    ]

    for idx, (num, title, desc) in enumerate(toc_items):
        col = idx % 3
        row = idx // 3
        x = Inches(0.8 + col * 3.98)
        y = Inches(1.85 + row * 2.55)
        w = Inches(3.78)
        h = Inches(2.35)

        card_t, tf_t = add_card(s2, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_MUTED)
        
        p1 = tf_t.paragraphs[0]
        p1.text = f"BAB {num}"
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_t.add_paragraph()
        p2.text = title
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(12.5)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_t.add_paragraph()
        p3.text = f"\n{desc}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10.5)
        p3.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 3: LATAR BELAKANG & EVOLUSI USAHA
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Latar Belakang & Filosofi Usaha", "Perjalanan kuliner legendaris sejak 1999 dan peluang ekspansi masa depan", tag="BAB I: PENDAHULUAN", slide_idx=3)

    story_steps = [
        ("1999", "Awal di Taman Aries", "Usaha didirikan pertama kali tahun 1999 di Taman Aries (Meruya) dengan resep autentik keluarga turun-temurun."),
        ("2016", "Relokasi ke Duri Kosambi", "Tahun 2016 berpindah ke ruko Jl. Angsoka Hijau IV Duri Kosambi untuk melayani perumahan dan civitas sekolah."),
        ("2019", "Basis Pelanggan Solid", "Tahun 2019 basis pelanggan setia terbentuk kuat, menjadi destinasi sarapan dan makan siang favorit keluarga."),
        ("2027/28", "Peluang Ekspansi Cabang", "Tingginya permintaan membuka peluang ekspansi 2027/2028 untuk mencari ruko lebih luas atau buka cabang baru.")
    ]

    for idx, (yr, ttl, dsc) in enumerate(story_steps):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(1.85)
        w = Inches(2.78)
        h = Inches(4.0)

        card_s, tf_s = add_card(s3, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 3 else BORDER_MUTED)
        
        p1 = tf_s.paragraphs[0]
        p1.text = yr
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_s.add_paragraph()
        p2.text = ttl
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_s.add_paragraph()
        p3.text = f"\n{dsc}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10.5)
        p3.font.color.rgb = TEXT_MUTED

    # Bottom Callout: Filosofi Adonan
    card_ph, tf_ph = add_card(s3, Inches(0.8), Inches(6.05), Inches(11.733), Inches(0.95), bg_color=GOLD_BG_LIGHT, border_color=GOLD_ACCENT)
    pph1 = tf_ph.paragraphs[0]
    pph1.text = "FILOSOFI MUTU: ADONAN MIE FRESH DIBUAT SETIAP MINGGU (BATCH MINGGUAN TANPA PENGAWET)"
    pph1.font.name = "Times New Roman"
    pph1.font.size = Pt(11)
    pph1.font.bold = True
    pph1.font.color.rgb = GOLD_ACCENT
    pph2 = tf_ph.add_paragraph()
    pph2.text = "Adonan mie dibuat mandiri setiap minggu menghasilkan tekstur kenyal alami berkilau (shining), bebas dari formalin dan pengawet berbahaya."
    pph2.font.name = "Times New Roman"
    pph2.font.size = Pt(10)
    pph2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 4: 5 PILAR KEUNGGULAN KOMPETITIF
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "5 Pilar Keunggulan Kompetitif", "Faktor diferensiasi penentu daya saing Bakmi Mimu Carina Sayang", tag="BAB I: PENDAHULUAN", slide_idx=4)

    pillars = [
        ("01", "Taste (Cita Rasa)", "Mie kenyal dipadu 3 pilihan minyak mie (babi, ayam, sayur+wijen), kaldu asli, & 3 topping murni: ayam putih, babi kecap, dan casiu madu (tanpa bebek & jamur)."),
        ("02", "Location (Lokasi)", "Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi, Cengkareng (persis di depan Taman TK, SD, SMP Kalam Kudus), lokasi sangat strategis & parkir motor aman."),
        ("03", "Service (Open Kitchen)", "Dapur terbuka gerobak di teras depan ruko, perebusan higienis terlihat jelas, dan dilengkapi saus botolan terpercaya (Cap Belibis & Mangga Besar)."),
        ("04", "Size (Variasi Porsi)", "Pilihan fleksibel: Porsi Kecil (27k), Reguler (29k-39k), hingga Porsi Jumbo (+100% mie dari porsi standar / 2x lipat porsi biasa 49k-59k) untuk segala tingkatan selera."),
        ("05", "Pelengkap & Minuman", "Pilihan pelengkap swikiaw kuah, pangsit rebus, baso sapi & ikan kuah (@4.5k-5.5k/pcs), baso goreng babi-udang 8k, badak sarsaparilla, susu kacang & liang teh 12k.")
    ]

    for idx, (num, tit, dsc) in enumerate(pillars):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(1.85)
        w = Inches(2.2)
        h = Inches(5.1)

        c_p, tf_p = add_card(s4, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_MUTED)
        
        p1 = tf_p.paragraphs[0]
        p1.text = f"PILAR {num}"
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_p.add_paragraph()
        p2.text = tit
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_p.add_paragraph()
        p3.text = f"\n{dsc}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 5: DAFTAR MENU & STRUKTUR HARGA (TABEL NATIVE POWERPOINT)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Daftar Menu & Struktur Harga Jual Resmi", "Struktur harga bersaing (Value-Based Pricing) sesuai daftar menu riil kedai", tag="BAB I: PENDAHULUAN", slide_idx=5)

    # Left Table: Menu Bakmi Utama
    tb_shape1 = s5.shapes.add_table(7, 3, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.9))
    tbl1 = tb_shape1.table
    tbl1.columns[0].width = Inches(2.7)
    tbl1.columns[1].width = Inches(1.3)
    tbl1.columns[2].width = Inches(1.7)

    style_table_cell(tbl1.cell(0, 0), "Menu Utama (Topping Murni)", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl1.cell(0, 1), "Harga", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl1.cell(0, 2), "Keterangan Porsi", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)

    m1_data = [
        ("Mie Campur (Ayam + Babi Kecap)", "Rp 29.000", "Reguler Favorit"),
        ("Mie Ayam Putih / Babi Saja", "Rp 29.000", "Topping Tunggal"),
        ("Mie Daging Casiu Madu", "Rp 31.000", "Casiu Panggang Gurih"),
        ("Mie Spesial 3 Isi Lengkap", "Rp 39.000", "Ayam, Babi & Casiu"),
        ("Porsi Jumbo (+100% Mie)", "Rp 49.000 - 59.000", "2x Lipat Porsi Standar"),
        ("Porsi Kecil (Sarapan Hemat)", "Rp 27.000 - 29.000", "Porsi Ringan Anak/Pagi")
    ]
    for r_idx, (m_n, m_h, m_k) in enumerate(m1_data, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl1.cell(r_idx, 0), m_n, font_size=10, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl1.cell(r_idx, 1), m_h, font_size=10, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl1.cell(r_idx, 2), m_k, font_size=9.5, color=TEXT_MUTED, bg_color=bg_r)

    # Right Table: Menu Pelengkap & Minuman
    tb_shape2 = s5.shapes.add_table(7, 3, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.9))
    tbl2 = tb_shape2.table
    tbl2.columns[0].width = Inches(2.7)
    tbl2.columns[1].width = Inches(1.3)
    tbl2.columns[2].width = Inches(1.7)

    style_table_cell(tbl2.cell(0, 0), "Pelengkap, Kuah & Minuman", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_LIGHT)
    style_table_cell(tbl2.cell(0, 1), "Harga", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_LIGHT)
    style_table_cell(tbl2.cell(0, 2), "Keterangan Item", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_LIGHT)

    m2_data = [
        ("Pangsit Rebus Kuah (5 pcs)", "Rp 22.500", "@Rp 4.500 / pcs"),
        ("Swikiaw Rebus Kuah (5 pcs)", "Rp 27.500", "@Rp 5.500 / pcs"),
        ("Baso Sapi & Baso Ikan (5 pcs)", "Rp 22.500", "Sama harga pangsit rebus"),
        ("Pangsit Goreng & Baso Goreng", "Rp 5.000 - 8.000", "Baso grg babi-udang 8k"),
        ("Badak Sarsaparilla", "Rp 12.000 - 14.000", "Soda legendaris Siantar"),
        ("Susu Kacang, Liang Teh, Teh", "Rp 1.000 - 12.000", "Saus Belibis & Mangga Bsr")
    ]
    for r_idx, (m_n, m_h, m_k) in enumerate(m2_data, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl2.cell(r_idx, 0), m_n, font_size=10, bold=True, color=NAVY_LIGHT, bg_color=bg_r)
        style_table_cell(tbl2.cell(r_idx, 1), m_h, font_size=10, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl2.cell(r_idx, 2), m_k, font_size=9.5, color=TEXT_MUTED, bg_color=bg_r)

    # =========================================================================
    # SLIDE 6: ANALISIS STP PASAR
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Analisis Pasar: STP Framework", "Segmenting, Targeting, dan Positioning pasar sasaran Bakmi Mimu", tag="BAB II: ASPEK PASAR", slide_idx=6)

    stp_cols = [
        ("SEGMENTASI (SEGMENTING)", [
            ("Geografis", "Radius 0 - 5 km meliputi Perumahan Duri Kosambi, Semanan, Taman Semanan Indah, Cengkareng, dan Puri Kembangan."),
            ("Demografis", "Pria & wanita usia 7 - 65 tahun, siswa-siswi, guru, orang tua murid Kalam Kudus, karyawan kantor, SES Menengah."),
            ("Perilaku", "Pencinta kuliner mie oriental yang mengutamakan tekstur kenyal, cita rasa bumbu otentik, dan higienitas ruang makan.")
        ]),
        ("TARGET PASAR (TARGETING)", [
            ("Keluarga Residensial", "Keluarga perumahan sekitar Duri Kosambi yang mencari santapan pagi sarapan dan makan siang lezat."),
            ("Civitas Kalam Kudus", "Siswa-siswi, guru, staf, dan orang tua murid Sekolah Kristen Kalam Kudus Kosambi tepat di depan gerai."),
            ("Pesan Antar & Komunitas", "Jemaat gereja hari Sabtu-Minggu serta pelanggan bawa pulang / pesan antar warga sekitar.")
        ]),
        ("POSISI PASAR (POSITIONING)", [
            ("Identitas Utama", "“Kedai Bakmi Otentik Legendaris dengan Resep 1999 dan Standar Kebersihan Higienis Terpercaya.”"),
            ("Value Proposition", "Memberikan kepuasan rasa resep keluarga 1999 dengan porsi mengenyangkan dan harga sangat sebanding."),
            ("Brand Perception", "Bukan sekadar warung tenda biasa, melainkan kedai bakmi keluarga yang bersih, ramah, dan terpercaya.")
        ])
    ]

    for idx, (title, points) in enumerate(stp_cols):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(1.85)
        w = Inches(3.78)
        h = Inches(5.1)

        c_stp, tf_stp = add_card(s6, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_NAVY if idx == 2 else BORDER_MUTED)
        
        p1 = tf_stp.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12.5)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        for sub, txt in points:
            psub = tf_stp.add_paragraph()
            psub.text = f"\n• {sub}:"
            psub.font.name = "Times New Roman"
            psub.font.size = Pt(10.5)
            psub.font.bold = True
            psub.font.color.rgb = GOLD_ACCENT

            ptxt = tf_stp.add_paragraph()
            ptxt.text = txt
            ptxt.font.name = "Times New Roman"
            ptxt.font.size = Pt(10)
            ptxt.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 7: PETA PERSAINGAN KOMPETITOR
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Analisis Persaingan & Peta Kompetitor", "Matriks perbandingan daya saing Bakmi Mimu terhadap pelaku usaha sejenis", tag="BAB II: ASPEK PASAR", slide_idx=7)

    comp_data = [
        ("BAKMI MIMU CARINA SAYANG", "Brand Unggulan Kita", GOLD_ACCENT, [
            ("Rasa & Resep", "Resep keluarga sejak 1999 di Taman Aries, 3 racikan minyak khas (babi, ayam, sayur+wijen), adonan fresh mingguan."),
            ("Fasilitas Kedai", "Ruko 1 lantai, sirkulasi 6 kipas angin, suasana hangat padat (22 kursi), open kitchen gerobak depan (tanpa AC/Wi-Fi)."),
            ("Varian Topping", "3 Topping murni: Ayam putih gurih, babi kecap manis, casiu madu (tanpa bebek/jamur), kuah swikiaw & pangsit."),
            ("Layanan & Saus", "Penyajian cepat < 5 menit, nota manual kasir, saus botol resmi Cap Belibis & Mangga Besar."),
            ("Harga & Value", "Rp 24.000 - Rp 59.000 (Sangat terjangkau dengan opsi porsi kecil, reguler, hingga jumbo +100% mie).")
        ]),
        ("BAKMI ALOK / BRAND BESAR", "Kompetitor Merek Terkenal", NAVY_LIGHT, [
            ("Rasa & Resep", "Cita rasa bakmi ayam rebus gurih, brand equity kuat di Jakarta Barat."),
            ("Fasilitas Kedai", "Restoran permanen ber-AC, kapasitas besar, antrean panjang saat jam makan siang."),
            ("Varian Topping", "Fokus dominan ayam kampung rebus, tidak menyediakan varian babi casiu madu."),
            ("Layanan & Saus", "Standar restoran waralaba terstruktur, layanan formal."),
            ("Harga & Value", "Rp 45.000 - Rp 65.000 (Segmen premium, relatif lebih mahal untuk porsi harian).")
        ]),
        ("WARUNG BAKMI LOKAL KOSAMBI", "Kompetitor Tradisional Sekitar", TEXT_MUTED, [
            ("Rasa & Resep", "Cita rasa standar warung tenda gerobak kaki lima, bumbu penyedap dominan."),
            ("Fasilitas Kedai", "Kios semi terbuka/tenda pinggir jalan, sirkulasi udara terbatas, parkir sempit."),
            ("Varian Topping", "Topping terbatas ayam cincang biasa, jarang menyediakan menu swikiaw rebus."),
            ("Layanan & Saus", "Kemasan plastik mika konvensional, rentan tumpah dan cepat dingin."),
            ("Harga & Value", "Rp 20.000 - Rp 25.000 (Harga murah namun kebersihan & fasilitas sangat minim).")
        ])
    ]

    for idx, (b_name, b_sub, b_col, pts) in enumerate(comp_data):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(1.85)
        w = Inches(3.78)
        h = Inches(5.1)

        c_cp, tf_cp = add_card(s7, x, y, w, h, bg_color=CARD_BG, border_color=b_col, border_width=1.5 if idx == 0 else 1.0)
        
        p1 = tf_cp.paragraphs[0]
        p1.text = b_name
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = b_col

        p2 = tf_cp.add_paragraph()
        p2.text = b_sub
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MUTED

        for p_name, p_desc in pts:
            prt = tf_cp.add_paragraph()
            prt.text = f"\n▸ {p_name}:"
            prt.font.name = "Times New Roman"
            prt.font.size = Pt(10)
            prt.font.bold = True
            prt.font.color.rgb = NAVY_PRIMARY

            prd = tf_cp.add_paragraph()
            prd.text = p_desc
            prd.font.name = "Times New Roman"
            prd.font.size = Pt(9.5)
            prd.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 8: BAURAN PEMASARAN 4P
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Strategi Bauran Pemasaran (4P Tradisional)", "Strategi eksekusi taktis: Product, Price, Place, dan Promotion", tag="BAB II: ASPEK PASAR", slide_idx=8)

    p4_cards = [
        ("PRODUCT (PRODUK)", GOLD_ACCENT, [
            "Resep autentik 1999 (Taman Aries), adonan fresh dibuat mandiri setiap minggu tanpa pengawet.",
            "3 Topping murni: Ayam Putih, Babi Kecap, & Casiu Madu gurih (tanpa daging bebek & jamur).",
            "3 Pilihan racikan minyak mie: Minyak Babi, Minyak Ayam, & Minyak Sayur + Wijen.",
            "Variasi porsi fleksibel: Porsi Kecil (27k), Reguler (29k-39k), hingga Jumbo (+100% mie, 49k-59k)."
        ]),
        ("PRICE (HARGA)", GREEN_ACCENT, [
            "Menerapkan Value-Based Pricing bersaing di kawasan residensial Duri Kosambi.",
            "Range harga bakmi: Rp 24.000 s.d. Rp 59.000 (belum termasuk menu pelengkap tambahan).",
            "Menu kuah: Pangsit rebus Rp 22.500 (5 pcs), Swikiaw Rp 27.500 (5 pcs), Baso sapi/ikan Rp 22.500.",
            "Gorengan Rp 5.000 - 8.000, minuman segar Rp 1.000 - 15.000 (Badak, Liang Teh, Susu Kacang, tanpa jeruk)."
        ]),
        ("PLACE (DISTRIBUSI / LOKASI)", NAVY_LIGHT, [
            "Lokasi ruko 1 lantai di Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi, Cengkareng, Jakarta Barat.",
            "Posisi emas persis di depan Taman TK, SD, dan SMP Sekolah Kristen Kalam Kudus.",
            "Fasilitas ruang makan berdaya tampung maksimal 22 orang padat (4 meja pendek + meja bar dinding).",
            "Fokus saluran penjualan langsung: Makan di tempat (Dine-In) dan pesanan bawa pulang (Take-Away)."
        ]),
        ("PROMOTION (PROMOSI)", NAVY_PRIMARY, [
            "Tidak mengadakan program promosi diskon buatan, kupon berhadiah, atau stamp card formal.",
            "Murni mengandalkan penjualan lokal getok tular (loyalitas pelanggan lama sejak 2019 & pembeli baru).",
            "Pemasaran organik sukarela dari ulasan pengunjung dan food vlogger/influencer kuliner (viral di TikTok).",
            "Papan nama kedai (signage) sederhana di fasad ruko yang terlihat jelas oleh para penjemput sekolah."
        ])
    ]

    for idx, (title, col, bullets) in enumerate(p4_cards):
        c = idx % 2
        r = idx // 2
        x = Inches(0.8 + c * 6.0)
        y = Inches(1.85 + r * 2.55)
        w = Inches(5.7)
        h = Inches(2.4)

        card_4p, tf_4p = add_card(s8, x, y, w, h, bg_color=CARD_BG, border_color=col)
        
        p1 = tf_4p.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = col

        for b in bullets:
            pb = tf_4p.add_paragraph()
            pb.text = f"• {b}"
            pb.font.name = "Times New Roman"
            pb.font.size = Pt(10)
            pb.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 9: BAURAN LAYANAN 3P EXTENDED
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Strategi Bauran Layanan (3P Extended)", "Pilar keunggulan operasional: People, Process, dan Physical Evidence", tag="BAB II: ASPEK PASAR", slide_idx=9)

    p3_cards = [
        ("PEOPLE (SUMBER DAYA MANUSIA)", GOLD_ACCENT, [
            "Tim operasional 3 karyawan fungsional: Koki Utama, Asisten Koki, dan Kasir-Pramusaji.",
            "Sistem remunerasi berbasis gaji pokok bulanan tetap (total alokasi gaji Rp 10 Jt/bulan).",
            "Tidak menerapkan bonus insentif penjualan, menjaga fokus tim pada konsistensi racikan rasa.",
            "Karyawan berseragam bersih, ramah menyambut tamu, dan menguasai varian menu topping."
        ]),
        ("PROCESS (PROSES OPERASIONAL)", NAVY_LIGHT, [
            "Sistem pemesanan kasir masih manual menggunakan nota fisik/kertas bon pesanan sederhana.",
            "Alur penyajian kilat: Perebusan mie fresh membutuhkan waktu 45 detik, saji di meja < 5 menit.",
            "Siklus produksi adonan mie dibuat mandiri secara terjadwal setiap minggu tanpa pengawet.",
            "Penyediaan botol saus meja standar resmi (Cap Belibis & Mangga Besar) yang selalu terisi bersih."
        ]),
        ("PHYSICAL EVIDENCE (BUKTI FISIK)", GREEN_ACCENT, [
            "Desain interior standar ruko 1 lantai yang bersih, bersahaja, rapi, dan fungsional.",
            "Fasilitas sirkulasi udara menggunakan 6 unit kipas angin dinding/plafon sejuk (tanpa fasilitas AC).",
            "Kapasitas dine-in maksimal 22 orang padat (4 meja pendek 16 kursi + 1 meja panjang dinding 6 kursi).",
            "Tidak menyediakan fasilitas Wi-Fi demi memaksimalkan perputaran meja (table turnover) di jam sibuk."
        ])
    ]

    for idx, (title, col, bullets) in enumerate(p3_cards):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(1.85)
        w = Inches(3.78)
        h = Inches(5.1)

        card_3p, tf_3p = add_card(s9, x, y, w, h, bg_color=CARD_BG, border_color=col)
        
        p1 = tf_3p.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = col

        for b in bullets:
            pb = tf_3p.add_paragraph()
            pb.text = f"\n▸ {b}"
            pb.font.name = "Times New Roman"
            pb.font.size = Pt(10)
            pb.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 10: STRUKTUR ORGANISASI & PEMBAGIAN TUGAS
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Struktur Organisasi & Pembagian Tugas", "Desain tata kelola tim operasional gerai Bakmi Mimu Carina Sayang", tag="BAB III: ASPEK MANAJEMEN", slide_idx=10)

    # Top Manager Card
    card_m, tf_m = add_card(s10, Inches(3.8), Inches(1.85), Inches(5.7), Inches(1.35), bg_color=CARD_BG_SOFT, border_color=GOLD_ACCENT, border_width=1.5)
    pm1 = tf_m.paragraphs[0]
    pm1.text = "OWNER / PENGELOLA UTAMA KEDAI (1 ORANG)"
    pm1.font.name = "Times New Roman"
    pm1.font.size = Pt(12.5)
    pm1.font.bold = True
    pm1.font.color.rgb = NAVY_PRIMARY

    pm2 = tf_m.add_paragraph()
    pm2.text = "Tanggung Jawab: Penetapan strategi usaha, pengawasan keuangan harian, kontrol cita rasa SOP resep 1999, pengadaan bahan baku daging segar, dan perencanaan ekspansi 2027/2028."
    pm2.font.name = "Times New Roman"
    pm2.font.size = Pt(10)
    pm2.font.color.rgb = TEXT_MUTED

    # 3 Subordinate Cards (3 Karyawan Tetap)
    sub_roles = [
        ("KOKI UTAMA (1 ORANG)", NAVY_LIGHT, [
            "Membuat adonan mie fresh setiap minggu.",
            "Memasak 3 topping: ayam, babi kecap, casiu.",
            "Menyiapkan 3 racikan minyak & rebus kaldu.",
            "Memimpin perebusan & peracikan mie di gerobak."
        ]),
        ("ASISTEN KOKI (1 ORANG)", GREEN_ACCENT, [
            "Menyiapkan bahan mentah & bumbu dapur.",
            "Melipat kulit pangsit & swikiaw isian babi-udang.",
            "Mencuci mangkok, sumpit, panci & alat makan.",
            "Menjaga kebersihan stasiun open kitchen gerobak."
        ]),
        ("KASIR & PRAMUSAJI (1 ORANG)", GOLD_ACCENT, [
            "Mencatat pesanan pelanggan dengan nota manual.",
            "Melayani penerimaan pembayaran tunai & QRIS.",
            "Menyajikan mangkok mie & minuman ke 22 kursi.",
            "Membersihkan meja makan & melayani take-away."
        ])
    ]

    for idx, (r_title, r_col, r_tasks) in enumerate(sub_roles):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(3.45)
        w = Inches(3.78)
        h = Inches(3.5)

        card_sb, tf_sb = add_card(s10, x, y, w, h, bg_color=CARD_BG, border_color=r_col)
        
        p1 = tf_sb.paragraphs[0]
        p1.text = r_title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = r_col

        for t in r_tasks:
            pt = tf_sb.add_paragraph()
            pt.text = f"\n• {t}"
            pt.font.name = "Times New Roman"
            pt.font.size = Pt(10)
            pt.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 11: TENAGA KERJA & SKEMA KOMPENSASI
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Tenaga Kerja & Kebijakan Kompensasi", "Kebutuhan SDM, alokasi upah bulanan, dan kebijakan ketenagakerjaan", tag="BAB III: ASPEK MANAJEMEN", slide_idx=11)

    # Top Stat Callouts
    c_s1, tf_s1 = add_card(s11, Inches(0.8), Inches(1.85), Inches(5.7), Inches(1.3), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    p1 = tf_s1.paragraphs[0]
    p1.text = "TOTAL TENAGA KERJA: 3 STAF KARYAWAN"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(14)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT
    p2 = tf_s1.add_paragraph()
    p2.text = "Formasi Fungsional: 1 Koki Utama Dapur, 1 Asisten Koki Dapur, 1 Kasir & Pramusaji."
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(10.5)
    p2.font.color.rgb = TEXT_MAIN

    c_s2, tf_s2 = add_card(s11, Inches(6.8), Inches(1.85), Inches(5.7), Inches(1.3), bg_color=CARD_BG, border_color=GREEN_ACCENT)
    p3 = tf_s2.paragraphs[0]
    p3.text = "TOTAL ANGGARAN GAJI: RP 120.000.000 / TAHUN"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(14)
    p3.font.bold = True
    p3.font.color.rgb = GREEN_ACCENT
    p4 = tf_s2.add_paragraph()
    p4.text = "Alokasi anggaran gaji tetap stabil Rp 10.000.000 per bulan (tanpa skema bonus insentif)."
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(10.5)
    p4.font.color.rgb = TEXT_MAIN

    comp_policies = [
        ("Gaji Koki Utama: Rp 3.800.000 / Bln", "Mengapresiasi keahlian khusus pembuatan adonan mie mingguan, pengolahan babi casiu panggang madu, ayam putih, dan racikan 3 minyak."),
        ("Gaji Asisten Koki: Rp 3.200.000 / Bln", "Kompensasi persiapan bahan baku, pengisian kaldu, perakitan pelengkap swikiaw/pangsit, serta sanitasi intensif alat masak dapur."),
        ("Gaji Kasir-Pramusaji: Rp 3.000.000 / Bln", "Kompensasi pencatatan nota pesanan manual, penanganan transaksi pembayaran tunai/QRIS, serta kecepatan penyajian ke 22 kursi."),
        ("Kebijakan Kompensasi Tetap (Non-Insentif)", "Menerapkan sistem upah bulanan pasti, fasilitas makan kedai harian, dan Tunjangan Hari Raya (THR) tahunan tanpa insentif penjualan.")
    ]

    for idx, (c_tit, c_txt) in enumerate(comp_policies):
        c = idx % 2
        r = idx // 2
        x = Inches(0.8 + c * 6.0)
        y = Inches(3.4 + r * 1.75)
        w = Inches(5.7)
        h = Inches(1.6)

        c_pol, tf_pol = add_card(s11, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_MUTED)
        p1 = tf_pol.paragraphs[0]
        p1.text = f"• {c_tit}"
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        p2 = tf_pol.add_paragraph()
        p2.text = c_txt
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 12: TATA LETAK RUKO 1 LANTAI & ZONASI OPERASIONAL
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "Denah Tata Letak Kedai Ruko 1 Lantai", "Visualisasi zonasi operasional, dapur terbuka teras (open kitchen), dan ruang makan (kapasitas 23 kursi)", tag="BAB IV: TEKNIS & OPERASI", slide_idx=12)

    # Top Card: Wadah Denah Arsitektural
    c_denah, tf_denah = add_card(s12, Inches(0.8), Inches(1.80), Inches(11.733), Inches(3.18), bg_color=CARD_BG, border_color=NAVY_PRIMARY)
    p_d0 = tf_denah.paragraphs[0]
    p_d0.text = "DENAH ARSITEKTURAL TATA LETAK FISIK KEDAI RUKO (ALUR SATU ARAH)"
    p_d0.font.name = "Times New Roman"
    p_d0.font.size = Pt(11)
    p_d0.font.bold = True
    p_d0.font.color.rgb = NAVY_PRIMARY

    p_d1 = tf_denah.add_paragraph()
    p_d1.text = "Alur Kerja: Pintu Masuk ➔ Open Kitchen Teras ➔ Rolling Door ➔ Ruang Makan Utama (Meja Reguler & Dinding) ➔ Fasilitas Sanitasi"
    p_d1.font.name = "Times New Roman"
    p_d1.font.size = Pt(8.5)
    p_d1.font.color.rgb = TEXT_MUTED

    if os.path.exists(layout_path):
        # Rasio aspek denah adalah 3.879 (1024 / 264)
        img_w = Inches(9.10)
        img_h = Inches(2.35)
        img_left = Inches(0.8) + (Inches(11.733) - img_w) / 2
        img_top = Inches(2.50)
        s12.shapes.add_picture(layout_path, img_left, img_top, img_w, img_h)

    # 3 Kartu Zonasi Berdampingan di Bagian Bawah
    # Card 1: Zona Depan & Dapur
    c_z1, tf_z1 = add_card(s12, Inches(0.8), Inches(5.12), Inches(3.75), Inches(2.05), bg_color=CARD_BG, border_color=NAVY_PRIMARY)
    p_z1 = tf_z1.paragraphs[0]
    p_z1.text = "ZONA 1: DAPUR TERAS & TERBUKA"
    p_z1.font.name = "Times New Roman"
    p_z1.font.size = Pt(10)
    p_z1.font.bold = True
    p_z1.font.color.rgb = NAVY_PRIMARY

    z1_items = [
        "Etalase kaca display racik mie & kompor gas perebusan",
        "Meja kuah kaldu, dispenser minum & kulkas sayur segar",
        "Meja persiapan potong sayur & pembungkusan pangsit",
        "1 Meja makan teras (4 kursi) & 1 unit freezer beku daging"
    ]
    for it in z1_items:
        p = tf_z1.add_paragraph()
        p.text = f"• {it}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(8.5)
        p.font.color.rgb = TEXT_MAIN

    # Card 2: Zona Tengah & Dining Room
    c_z2, tf_z2 = add_card(s12, Inches(4.79), Inches(5.12), Inches(3.75), Inches(2.05), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    p_z2 = tf_z2.paragraphs[0]
    p_z2.text = "ZONA 2: RUANG MAKAN (DINE-IN)"
    p_z2.font.name = "Times New Roman"
    p_z2.font.size = Pt(10)
    p_z2.font.bold = True
    p_z2.font.color.rgb = GOLD_ACCENT

    z2_items = [
        "Pembatas partisi fleksibel Rolling Door (sekat debu/udara)",
        "3 Meja makan reguler kayu (kapasitas 12 kursi santap)",
        "1 Meja makan dinding memanjang (7 kursi solo diner)",
        "Sirkulasi 6 kipas angin dinding sejuk bebas pengap"
    ]
    for it in z2_items:
        p = tf_z2.add_paragraph()
        p.text = f"• {it}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(8.5)
        p.font.color.rgb = TEXT_MAIN

    # Card 3: Zona Belakang & Sanitasi
    c_z3, tf_z3 = add_card(s12, Inches(8.78), Inches(5.12), Inches(3.75), Inches(2.05), bg_color=CARD_BG, border_color=NAVY_LIGHT)
    p_z3 = tf_z3.paragraphs[0]
    p_z3.text = "ZONA 3: FASILITAS & SANITASI"
    p_z3.font.name = "Times New Roman"
    p_z3.font.size = Pt(10)
    p_z3.font.bold = True
    p_z3.font.color.rgb = NAVY_LIGHT

    z3_items = [
        "Meja barang penyimpanan stok piring & kemasan bersih",
        "Wastafel cuci tangan higienis bagi pengunjung kedai",
        "1 Unit kamar mandi / toilet ruko tertutup higienis",
        "Total kapasitas santap kedai: 23 kursi pengunjung aktif"
    ]
    for it in z3_items:
        p = tf_z3.add_paragraph()
        p.text = f"• {it}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(8.5)
        p.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 13: ALUR PROSES PRODUKSI & SIKLUS MINGGUAN ADONAN
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_header(s13, "Alur Proses Produksi & Siklus Mingguan Adonan", "Tahapan pengolahan higienis dari adonan mingguan hingga siap disajikan", tag="BAB IV: TEKNIS & OPERASI", slide_idx=13)

    steps_prod = [
        ("LANGKAH 01", "Adonan Fresh Mingguan", "Setiap Minggu Pagi", "Adonan mie dibuat mandiri secara konsisten setiap minggu tanpa bahan pengawet kimia."),
        ("LANGKAH 02", "Pengolahan 3 Topping", "05.30 - 07.30 WIB", "Memanggang babi casiu madu, menumis babi kecap gurih, dan merebus potongan ayam putih."),
        ("LANGKAH 03", "Racikan 3 Minyak & Kaldu", "06.00 - 08.00 WIB", "Menyiapkan minyak babi, minyak ayam, dan minyak sayur+wijen serta rebusan kaldu tulang gurih."),
        ("LANGKAH 04", "Perebusan Mie Seketika", "Saat Order Masuk", "Mie direbus 45 detik dalam air mendidih saat nota pesanan diterima dari tamu/kasir."),
        ("LANGKAH 05", "Plating & Penyajian Meja", "< 5 Menit Total", "Pencampuran minyak pilihan, penataan topping daging melimpah, daun bawang, & saus meja.")
    ]

    for idx, (st, tit, tm, dsc) in enumerate(steps_prod):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(1.85)
        w = Inches(2.2)
        h = Inches(5.1)

        c_st, tf_st = add_card(s13, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 3 else BORDER_MUTED)
        
        p1 = tf_st.paragraphs[0]
        p1.text = st
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_st.add_paragraph()
        p2.text = tit
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_st.add_paragraph()
        p3.text = f"Waktu: {tm}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = NAVY_LIGHT

        p4 = tf_st.add_paragraph()
        p4.text = f"\n{dsc}"
        p4.font.name = "Times New Roman"
        p4.font.size = Pt(9.5)
        p4.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 14: JADWAL PELAKSANAAN PRA-OPERASI RELOKASI 2016
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14)
    add_header(s14, "Jadwal Pelaksanaan Pra-Operasi Relokasi 2016", "Rencana timeline 12 minggu persiapan operasional gerai Duri Kosambi", tag="BAB IV: TEKNIS & OPERASI", slide_idx=14)

    phases = [
        ("FASE 1: SURVEY & SEWA RUKO", "Minggu 1 - 4", GOLD_ACCENT, [
            "A. Observasi lokasi Jl. Angsoka Hijau IV depan Kalam Kudus.",
            "B. Negosiasi dan pembayaran sewa ruko 1 lantai Rp 20.000.000.",
            "C. Pengurusan izin lingkungan RT/RW dan retribusi kebersihan.",
            "D. Analisis potensi pasar sarapan murid dan orang tua sekolah."
        ]),
        ("FASE 2: RENOVASI & FIT-OUT KEDAI", "Minggu 5 - 8", NAVY_LIGHT, [
            "E. Pengecatan ruko 1 lantai dan perbaikan instalasi pipa air.",
            "F. Pemasangan 6 unit kipas angin dinding/plafon (@Rp 150.000).",
            "G. Penataan teras depan untuk stasiun gerobak open kitchen.",
            "H. Pemasangan spanduk dan plang nama kedai Bakmi Mimu."
        ]),
        ("FASE 3: PENGADAAN & SETTING DAPUR", "Minggu 9 - 10", GREEN_ACCENT, [
            "I. Pembuatan gerobak etalase kaca stasiun masak (Rp 8 Jt).",
            "J. Pengadaan 4 meja pendek, meja dinding, dan kursi 37 pcs.",
            "K. Pembelian 1 kulkas, 1 freezer daging, panci masak & mangkok.",
            "L. Pengadaan saus standar (Belibis & Mangga Besar) dan botol meja."
        ]),
        ("FASE 4: TRIAL RUN & PEMBUKAAN", "Minggu 11 - 12", NAVY_PRIMARY, [
            "M. Rekrutmen 3 karyawan tetap (Koki, Asisten Koki, Kasir-Pramusaji).",
            "N. Uji coba pembuatan adonan mie mingguan dan kaldu SOP 1999.",
            "O. Simulasi alur pelayanan meja dine-in 22 kursi dan nota manual.",
            "P. Pembukaan resmi gerai Duri Kosambi melayani warga sekitar."
        ])
    ]

    for idx, (p_tit, p_time, p_col, p_items) in enumerate(phases):
        c = idx % 2
        r = idx // 2
        x = Inches(0.8 + c * 6.0)
        y = Inches(1.85 + r * 2.55)
        w = Inches(5.7)
        h = Inches(2.4)

        card_ph, tf_ph = add_card(s14, x, y, w, h, bg_color=CARD_BG, border_color=p_col)
        p1 = tf_ph.paragraphs[0]
        p1.text = p_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = p_col

        p2 = tf_ph.add_paragraph()
        p2.text = f"Jadwal: {p_time}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MUTED

        for itm in p_items:
            pit = tf_ph.add_paragraph()
            pit.text = f"• {itm}"
            pit.font.name = "Times New Roman"
            pit.font.size = Pt(9.5)
            pit.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 15: RENCANA INVESTASI AWAL (INITIAL OUTLAY)
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15)
    add_header(s15, "Rencana Modal Investasi Awal (Initial Outlay)", "Alokasi anggaran modal awal relokasi 2016 untuk pembukaan kedai Duri Kosambi", tag="BAB V: ASPEK KEUANGAN", slide_idx=15)

    # Hero Stat Callout Box
    card_hero, tf_hero = add_card(s15, Inches(0.8), Inches(1.85), Inches(11.733), Inches(1.4), bg_color=CARD_BG, border_color=GOLD_ACCENT, border_width=1.5)
    p_h1 = tf_hero.paragraphs[0]
    p_h1.text = "TOTAL INITIAL OUTLAY KEBUTUHAN MODAL: RP 62.000.000"
    p_h1.font.name = "Times New Roman"
    p_h1.font.size = Pt(21)
    p_h1.font.bold = True
    p_h1.font.color.rgb = GOLD_ACCENT

    p_h2 = tf_hero.add_paragraph()
    p_h2.text = "Investasi awal didanai 100% modal sendiri tanpa utang bank. Seluruh alokasi modal difokuskan untuk pengadaan aktiva tetap gerobak & fasilitas kedai, sewa tempat ruko 1 tahun, dan pra-operasi."
    p_h2.font.name = "Times New Roman"
    p_h2.font.size = Pt(11)
    p_h2.font.color.rgb = TEXT_MAIN

    # 5 Components Cards
    outlay_items = [
        ("01. Aktiva Tetap (Capex)", "Rp 40.700.000", "65,6%", "Pengadaan gerobak etalase, meja, kursi 37 pcs, kipas angin, kulkas, freezer, panci, & mangkok."),
        ("02. Sewa Tempat Ruko", "Rp 20.000.000", "32,3%", "Alokasi biaya sewa ruko 1 lantai Duri Kosambi untuk 1 tahun masa operasional."),
        ("03. Perizinan & AMDAL", "Rp 300.000", "0,5%", "Retribusi kebersihan lingkungan RT/RW dan pengadaan tempat sampah kedai higienis."),
        ("04. Biaya Survey Pasar", "Rp 500.000", "0,8%", "Observasi demografi perumahan Kosambi dan potensi kantin sekolah Kalam Kudus."),
        ("05. Biaya Promosi Awal", "Rp 500.000", "0,8%", "Pembuatan spanduk pembukaan dan plang penunjuk arah kedai Bakmi Mimu.")
    ]

    for idx, (c_name, c_val, c_pct, c_dsc) in enumerate(outlay_items):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(3.45)
        w = Inches(2.2)
        h = Inches(3.5)

        c_box, tf_box = add_card(s15, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_MUTED)
        p1 = tf_box.paragraphs[0]
        p1.text = c_name
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        p2 = tf_box.add_paragraph()
        p2.text = c_val
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = GOLD_ACCENT

        p3 = tf_box.add_paragraph()
        p3.text = f"Porsi: {c_pct}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = GREEN_ACCENT

        p4 = tf_box.add_paragraph()
        p4.text = f"\n{c_dsc}"
        p4.font.name = "Times New Roman"
        p4.font.size = Pt(9.5)
        p4.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 16: RINCIAN AKTIVA TETAP & DEPRESIASI (TABEL NATIVE)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16)
    add_header(s16, "Aktiva Tetap & Beban Depresiasi", "Inventarisasi 12 item aktiva tetap dan penyusutan garis lurus 20% per tahun", tag="BAB V: ASPEK KEUANGAN", slide_idx=16)

    # Left Box: Parameter Depresiasi
    c_dep, tf_dep = add_card(s16, Inches(0.8), Inches(1.85), Inches(3.6), Inches(5.1), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    p1 = tf_dep.paragraphs[0]
    p1.text = "PARAMETER DEPRESIASI"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT

    p2 = tf_dep.add_paragraph()
    p2.text = f"\nTotal Nilai Perolehan:\nRp {met['capex_total']:,.0f}".replace(",", ".")
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(15)
    p2.font.bold = True
    p2.font.color.rgb = NAVY_PRIMARY

    p3 = tf_dep.add_paragraph()
    p3.text = "\nMetode Penyusutan:\nMetode Garis Lurus (Straight-Line)"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(10.5)
    p3.font.color.rgb = TEXT_MUTED

    p4 = tf_dep.add_paragraph()
    p4.text = "\nTarif Depresiasi Tahunan:\n20,00% (Masa Manfaat 5 Tahun)"
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(10.5)
    p4.font.color.rgb = TEXT_MUTED

    p5 = tf_dep.add_paragraph()
    p5.text = f"\nBeban Depresiasi / Tahun:\nRp {met['depresiasi_per_year']:,.0f}".replace(",", ".")
    p5.font.name = "Times New Roman"
    p5.font.size = Pt(15)
    p5.font.bold = True
    p5.font.color.rgb = GREEN_ACCENT

    # Right Table: 6 Kelompok Capex
    tb_shape_c = s16.shapes.add_table(7, 3, Inches(4.7), Inches(1.85), Inches(7.833), Inches(5.1))
    tbl_c = tb_shape_c.table
    tbl_c.columns[0].width = Inches(3.2)
    tbl_c.columns[1].width = Inches(1.5)
    tbl_c.columns[2].width = Inches(3.133)

    style_table_cell(tbl_c.cell(0, 0), "Kelompok Aktiva Tetap (12 Item)", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_c.cell(0, 1), "Nilai Perolehan", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_c.cell(0, 2), "Rincian Item Peralatan", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)

    capex_rows = [
        ("1. Gerobak & Etalase Kaca Depan", "Rp 8.000.000", "Stasiun masak open kitchen teras ruko"),
        ("2. Perabotan Dine-In (Meja & Kursi)", "Rp 8.300.000", "4 Meja pendek, meja dinding, kursi 37 pcs"),
        ("3. Sirkulasi Udara (6 Kipas Angin)", "Rp 900.000", "6 Unit kipas angin dinding/plafon @150k"),
        ("4. Pendingin Dapur (Kulkas & Freezer)", "Rp 6.000.000", "1 Kulkas sayur 2,5 Jt, 1 Freezer daging 3,5 Jt"),
        ("5. Panci Masak & Stok Mangkok", "Rp 6.500.000", "Panci kaldu, mangkok keramik, sendok, sumpit"),
        ("6. Renovasi Ruko & Signage Nama", "Rp 11.000.000", "Cat ruko 1 lantai, kelistrikan/air, plang nama")
    ]
    for r_idx, (c_g, c_v, c_d) in enumerate(capex_rows, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_c.cell(r_idx, 0), c_g, font_size=10, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_c.cell(r_idx, 1), c_v, font_size=10, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_c.cell(r_idx, 2), c_d, font_size=9.5, color=TEXT_MUTED, bg_color=bg_r)

    # =========================================================================
    # SLIDE 17: PROYEKSI PENDAPATAN CASH INFLOW 5 TAHUN (TABEL NATIVE)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_background(s17)
    add_header(s17, "Proyeksi Pendapatan (Cash Inflow 5 Tahun)", "Estimasi penerimaan kas berbasis 300 hari operasi tahunan dengan eskalasi volume riil", tag="BAB V: ASPEK KEUANGAN", slide_idx=17)

    tb_shape_cif = s17.shapes.add_table(6, 4, Inches(0.8), Inches(1.85), Inches(11.733), Inches(5.1))
    tbl_cif = tb_shape_cif.table
    tbl_cif.columns[0].width = Inches(2.2)
    tbl_cif.columns[1].width = Inches(2.8)
    tbl_cif.columns[2].width = Inches(1.8)
    tbl_cif.columns[3].width = Inches(4.933)

    style_table_cell(tbl_cif.cell(0, 0), "Periode Proyeksi", font_size=11, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cif.cell(0, 1), "Cash Inflow (CIF) / Tahun", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cif.cell(0, 2), "Pertumbuhan", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cif.cell(0, 3), "Asumsi & Justifikasi Operasional", font_size=11, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)

    cif_table_rows = [
        ("Tahun 1 (2016/17)", f"Rp {met['cif_years'][0]:,.0f}".replace(",", "."), "Basis Awal", "Omzet riil 300 hari operasi: rata-rata Rp 4,94 Jt/hari (menu utama + pelengkap)."),
        ("Tahun 2 (2017/18)", f"Rp {met['cif_years'][1]:,.0f}".replace(",", "."), f"+{(met['cif_years'][1]/met['cif_years'][0]-1)*100:.1f}%", "Peningkatan loyalitas pelanggan sekolah Kalam Kudus & warga perumahan sekitar."),
        ("Tahun 3 (2018/19)", f"Rp {met['cif_years'][2]:,.0f}".replace(",", "."), f"+{(met['cif_years'][2]/met['cif_years'][1]-1)*100:.1f}%", "Basis pelanggan mantap terbentuk (2019), pesanan bungkus take-away meningkat pesat."),
        ("Tahun 4 (2019/20)", f"Rp {met['cif_years'][3]:,.0f}".replace(",", "."), f"+{(met['cif_years'][3]/met['cif_years'][2]-1)*100:.1f}%", "Penjualan stabil dengan pelanggan tetap keluarga residensial perumahan Duri Kosambi."),
        ("Tahun 5 (2020/21)", f"Rp {met['cif_years'][4]:,.0f}".replace(",", "."), f"+{(met['cif_years'][4]/met['cif_years'][3]-1)*100:.1f}%", "Kapasitas kedai ruko optimal & persiapan akumulasi kas untuk rencana ekspansi 2027/28.")
    ]
    for r_idx, (y_t, y_v, y_g, y_d) in enumerate(cif_table_rows, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_cif.cell(r_idx, 0), y_t, font_size=10.5, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_cif.cell(r_idx, 1), y_v, font_size=11, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cif.cell(r_idx, 2), y_g, font_size=10, bold=True, color=GREEN_ACCENT, align=PP_ALIGN.CENTER, bg_color=bg_r)
        style_table_cell(tbl_cif.cell(r_idx, 3), y_d, font_size=10, color=TEXT_MUTED, bg_color=bg_r)

    # =========================================================================
    # SLIDE 18: PROYEKSI BIAYA & LABA BERSIH 5 TAHUN (TABEL NATIVE)
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    set_slide_background(s18)
    add_header(s18, "Proyeksi Biaya & Laba Bersih (NCF / EAT)", "Analisis arus kas keluar operasional dan laba bersih selama 5 tahun", tag="BAB V: ASPEK KEUANGAN", slide_idx=18)

    tb_shape_cof = s18.shapes.add_table(6, 5, Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.3))
    tbl_cof = tb_shape_cof.table
    tbl_cof.columns[0].width = Inches(2.133)
    tbl_cof.columns[1].width = Inches(2.4)
    tbl_cof.columns[2].width = Inches(2.4)
    tbl_cof.columns[3].width = Inches(2.4)
    tbl_cof.columns[4].width = Inches(2.4)

    style_table_cell(tbl_cof.cell(0, 0), "Tahun Operasi", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 1), "Inflow (CIF)", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 2), "Outflow (COF)", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 3), "Net Cash Flow (EAT)", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 4), "Proceed (Kas Riil)", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)

    for i in range(5):
        bg_r = CARD_BG if (i+1) % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_cof.cell(i+1, 0), f"Tahun {i+1}", font_size=10.5, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 1), f"Rp {met['cif_years'][i]:,.0f}".replace(",", "."), font_size=10, bold=True, color=TEXT_MAIN, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 2), f"Rp {met['cof_years'][i]:,.0f}".replace(",", "."), font_size=10, color=TEXT_MUTED, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 3), f"Rp {met['ncf_years'][i]:,.0f}".replace(",", "."), font_size=10.5, bold=True, color=GREEN_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 4), f"Rp {met['proceed_years'][i]:,.0f}".replace(",", "."), font_size=10.5, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)

    # Bottom Highlight Card
    c_b18, tf_b18 = add_card(s18, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.75), bg_color=GREEN_BG_LIGHT, border_color=GREEN_ACCENT)
    pb = tf_b18.paragraphs[0]
    pb.text = "KESIMPULAN KINERJA: Margin Laba Bersih Operasional Mencapai 61,97% di Tahun Pertama"
    pb.font.name = "Times New Roman"
    pb.font.size = Pt(11)
    pb.font.bold = True
    pb.font.color.rgb = GREEN_ACCENT
    pb2 = tf_b18.add_paragraph()
    pb2.text = "Arus kas masuk sangat likuid untuk menutup seluruh beban sewa ruko Rp 20 Jt, gaji 3 staf Rp 120 Jt, serta HPP bahan baku daging mingguan."
    pb2.font.name = "Times New Roman"
    pb2.font.size = Pt(9.5)
    pb2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 19: EVALUASI MODAL: NET PRESENT VALUE (TABEL NATIVE)
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    set_slide_background(s19)
    add_header(s19, "Evaluasi Modal: Net Present Value (NPV)", "Hasil perhitungan nilai sekarang bersih kas masuk dengan discount rate 20%", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=19)

    # Left Hero Box
    c_npv, tf_npv = add_card(s19, Inches(0.8), Inches(1.85), Inches(4.5), Inches(5.1), bg_color=CARD_BG, border_color=GREEN_ACCENT, border_width=2.0)
    pn1 = tf_npv.paragraphs[0]
    pn1.text = "NET PRESENT VALUE (NPV)"
    pn1.font.name = "Times New Roman"
    pn1.font.size = Pt(14)
    pn1.font.bold = True
    pn1.font.color.rgb = GOLD_ACCENT

    pn2 = tf_npv.add_paragraph()
    pn2.text = f"Rp {met['npv']:,.0f}".replace(",", ".")
    pn2.font.name = "Times New Roman"
    pn2.font.size = Pt(25)
    pn2.font.bold = True
    pn2.font.color.rgb = GREEN_ACCENT

    pn3 = tf_npv.add_paragraph()
    pn3.text = "\nSTATUS KELAYAKAN: SANGAT LAYAK (NPV > 0)"
    pn3.font.name = "Times New Roman"
    pn3.font.size = Pt(11.5)
    pn3.font.bold = True
    pn3.font.color.rgb = NAVY_PRIMARY

    pn4 = tf_npv.add_paragraph()
    pn4.text = f"\nBerdasarkan kriteria standar FEB UKRIDA, rencana investasi dinyatakan LAYAK jika NPV > 0. Nilai NPV sebesar Rp {met['npv']/1e9:.2f} Milyar membuktikan bahwa proyek Bakmi Mimu Carina Sayang mampu menutupi seluruh modal investasi awal Rp 62 Juta dan menghasilkan surplus kekayaan bersih yang luar biasa."
    pn4.font.name = "Times New Roman"
    pn4.font.size = Pt(10)
    pn4.font.color.rgb = TEXT_MUTED

    # Right Table: Discounted PV 5 Tahun
    tb_shape_pv = s19.shapes.add_table(7, 3, Inches(5.5), Inches(1.85), Inches(7.033), Inches(5.1))
    tbl_pv = tb_shape_pv.table
    tbl_pv.columns[0].width = Inches(2.6)
    tbl_pv.columns[1].width = Inches(2.2)
    tbl_pv.columns[2].width = Inches(2.233)

    style_table_cell(tbl_pv.cell(0, 0), "Periode Aliran Kas", font_size=10.5, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_pv.cell(0, 1), "Nominal Arus Kas", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_pv.cell(0, 2), "Present Value (DF 20%)", font_size=10.5, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)

    pv_table_data = [
        ("Tahun 0 (Initial Outlay)", f"-Rp {met['initial_outlay']:,.0f}".replace(",", "."), "Modal Awal Relokasi"),
        ("Tahun 1 (Proceed Thn 1)", f"Rp {met['proceed_years'][0]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][0]:,.0f}".replace(",", ".")),
        ("Tahun 2 (Proceed Thn 2)", f"Rp {met['proceed_years'][1]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][1]:,.0f}".replace(",", ".")),
        ("Tahun 3 (Proceed Thn 3)", f"Rp {met['proceed_years'][2]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][2]:,.0f}".replace(",", ".")),
        ("Tahun 4 (Proceed Thn 4)", f"Rp {met['proceed_years'][3]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][3]:,.0f}".replace(",", ".")),
        ("Tahun 5 (Proceed Thn 5)", f"Rp {met['proceed_years'][4]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][4]:,.0f}".replace(",", "."))
    ]
    for r_idx, (p_t, p_v, p_pv) in enumerate(pv_table_data, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_pv.cell(r_idx, 0), p_t, font_size=10, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_pv.cell(r_idx, 1), p_v, font_size=10, color=TEXT_MAIN, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_pv.cell(r_idx, 2), p_pv, font_size=10, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)

    # =========================================================================
    # SLIDE 20: EVALUASI MODAL: IRR & PAYBACK PERIOD
    # =========================================================================
    s20 = prs.slides.add_slide(blank_layout)
    set_slide_background(s20)
    add_header(s20, "Evaluasi Modal: IRR & Payback Period", "Tingkat pengembalian internal dan kecepatan pemulihan modal investasi", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=20)

    # Hero Stat Left: IRR
    c_irr, tf_irr = add_card(s20, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.1), bg_color=CARD_BG, border_color=GREEN_ACCENT, border_width=1.5)
    pi1 = tf_irr.paragraphs[0]
    pi1.text = "INTERNAL RATE OF RETURN (IRR)"
    pi1.font.name = "Times New Roman"
    pi1.font.size = Pt(14)
    pi1.font.bold = True
    pi1.font.color.rgb = GOLD_ACCENT

    pi2 = tf_irr.add_paragraph()
    pi2.text = "> 20,00%"
    pi2.font.name = "Times New Roman"
    pi2.font.size = Pt(36)
    pi2.font.bold = True
    pi2.font.color.rgb = GREEN_ACCENT

    pi3 = tf_irr.add_paragraph()
    pi3.text = "\nTARGET / STANDAR FEB: > 20,00% (Opportunity Cost)"
    pi3.font.name = "Times New Roman"
    pi3.font.size = Pt(11)
    pi3.font.bold = True
    pi3.font.color.rgb = NAVY_PRIMARY

    pi4 = tf_irr.add_paragraph()
    pi4.text = "\nSTATUS: SANGAT LAYAK (Jauh Melampaui Standar)"
    pi4.font.name = "Times New Roman"
    pi4.font.size = Pt(11)
    pi4.font.bold = True
    pi4.font.color.rgb = GREEN_ACCENT

    pi5 = tf_irr.add_paragraph()
    pi5.text = f"\nNilai IRR yang melampaui 20,00% membuktikan efisiensi operasional kedai yang sangat prima. Imbal hasil riil proyek jauh melampaui tingkat diskonto modal dan suku bunga deposito perbankan."
    pi5.font.name = "Times New Roman"
    pi5.font.size = Pt(10)
    pi5.font.color.rgb = TEXT_MUTED

    # Hero Stat Right: Payback Period
    c_pp, tf_pp = add_card(s20, Inches(6.8), Inches(1.85), Inches(5.7), Inches(5.1), bg_color=CARD_BG, border_color=NAVY_LIGHT, border_width=1.5)
    ppp1 = tf_pp.paragraphs[0]
    ppp1.text = "PAYBACK PERIOD (PP)"
    ppp1.font.name = "Times New Roman"
    ppp1.font.size = Pt(14)
    ppp1.font.bold = True
    ppp1.font.color.rgb = NAVY_LIGHT

    ppp2 = tf_pp.add_paragraph()
    ppp2.text = f"{met['pp_months']:.2f} BULAN"
    ppp2.font.name = "Times New Roman"
    ppp2.font.size = Pt(36)
    ppp2.font.bold = True
    ppp2.font.color.rgb = NAVY_LIGHT

    ppp3 = tf_pp.add_paragraph()
    ppp3.text = f"\nSETARA: {met['pp_years']:.2f} TAHUN (KURANG DARI 1 BULAN OPERASI)"
    ppp3.font.name = "Times New Roman"
    ppp3.font.size = Pt(11)
    ppp3.font.bold = True
    ppp3.font.color.rgb = NAVY_PRIMARY

    ppp4 = tf_pp.add_paragraph()
    ppp4.text = "\nTARGET / STANDAR FEB: < 3,00 TAHUN (36 BULAN) ➔ SANGAT LAYAK"
    ppp4.font.name = "Times New Roman"
    ppp4.font.size = Pt(11)
    ppp4.font.bold = True
    ppp4.font.color.rgb = GREEN_ACCENT

    ppp5 = tf_pp.add_paragraph()
    ppp5.text = f"\nModal awal sebesar Rp 62.000.000 telah pulih kembali sepenuhnya hanya dalam tempo {met['pp_months']:.2f} bulan (sekitar 25 hari kerja pada bulan pertama). Seluruh arus kas berikutnya murni menjadi keuntungan likuid."
    ppp5.font.name = "Times New Roman"
    ppp5.font.size = Pt(10)
    ppp5.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 21: PROFITABILITY INDEX & MATRIKS KELAYAKAN (TABEL NATIVE)
    # =========================================================================
    s21 = prs.slides.add_slide(blank_layout)
    set_slide_background(s21)
    add_header(s21, "Profitability Index & Matriks Kelayakan", "Rangkuman skor evaluasi investasi dan ketahanan terhadap risiko pasar", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=21)

    tb_shape_mat = s21.shapes.add_table(5, 4, Inches(0.8), Inches(1.85), Inches(11.733), Inches(3.2))
    tbl_mat = tb_shape_mat.table
    tbl_mat.columns[0].width = Inches(3.5)
    tbl_mat.columns[1].width = Inches(2.8)
    tbl_mat.columns[2].width = Inches(2.8)
    tbl_mat.columns[3].width = Inches(2.633)

    style_table_cell(tbl_mat.cell(0, 0), "Metode Kriteria Investasi", font_size=11, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_mat.cell(0, 1), "Hasil Perhitungan", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_mat.cell(0, 2), "Standar Kelayakan FEB", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_mat.cell(0, 3), "Status Evaluasi", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=NAVY_PRIMARY)

    crit_rows = [
        ("Net Present Value (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "> Rp 0 (Positif)", "SANGAT LAYAK"),
        ("Internal Rate of Return (IRR)", "> 20,00%", "> 20,00% (Opportunity Cost)", "SANGAT LAYAK"),
        ("Payback Period (PP)", f"{met['pp_months']:.2f} Bulan (0,07 Thn)", "< 36 Bulan (3 Tahun)", "SANGAT LAYAK"),
        ("Profitability Index (PI)", f"{met['pi']:.2f}", "> 1,20", "SANGAT LAYAK")
    ]
    for r_idx, (c_n, c_v, c_s, c_st) in enumerate(crit_rows, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_mat.cell(r_idx, 0), c_n, font_size=10.5, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_mat.cell(r_idx, 1), c_v, font_size=11, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_mat.cell(r_idx, 2), c_s, font_size=10, color=TEXT_MUTED, align=PP_ALIGN.CENTER, bg_color=bg_r)
        style_table_cell(tbl_mat.cell(r_idx, 3), c_st, font_size=10.5, bold=True, color=GREEN_ACCENT, align=PP_ALIGN.CENTER, bg_color=bg_r)

    # Bottom Sensitivity Card
    c_sens, tf_sens = add_card(s21, Inches(0.8), Inches(5.25), Inches(11.733), Inches(1.7), bg_color=CARD_BG, border_color=BORDER_NAVY)
    ps1 = tf_sens.paragraphs[0]
    ps1.text = "ANALISIS SENSITIVITAS & KETAHANAN TERHADAP RISIKO PASAR"
    ps1.font.name = "Times New Roman"
    ps1.font.size = Pt(11.5)
    ps1.font.bold = True
    ps1.font.color.rgb = NAVY_LIGHT

    ps2 = tf_sens.add_paragraph()
    ps2.text = "• Skenario Kenaikan Harga Bahan Baku Daging (+10%): Nilai NPV tetap positif di atas Rp 5,2 Milyar dengan Payback Period tetap < 1 bulan.\n" \
               "• Skenario Penurunan Volume Penjualan (-15%): Kedai tetap menghasilkan arus kas operasional surplus yang sangat sehat untuk menutup gaji dan sewa ruko."
    ps2.font.name = "Times New Roman"
    ps2.font.size = Pt(10)
    ps2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 22: KESIMPULAN AKHIR & PENUTUP (EXECUTIVE OUTRO)
    # =========================================================================
    s22 = prs.slides.add_slide(blank_layout)
    set_slide_background(s22)

    # Grand Box Container
    c_end, tf_end_box = add_card(s22, Inches(0.8), Inches(0.5), Inches(11.733), Inches(6.5), bg_color=CARD_BG, border_color=GREEN_ACCENT, border_width=2.0)

    # UKRIDA Logo at top center
    if os.path.exists(logo_path):
        s22.shapes.add_picture(logo_path, Inches(6.066), Inches(0.75), Inches(1.2), Inches(1.2))

    tb_end = s22.shapes.add_textbox(Inches(1.0), Inches(2.05), Inches(11.333), Inches(4.7))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True
    tf_end.margin_left = tf_end.margin_right = tf_end.margin_top = tf_end.margin_bottom = 0

    pe1 = tf_end.paragraphs[0]
    pe1.alignment = PP_ALIGN.CENTER
    pe1.text = "KESIMPULAN AKHIR KELAYAKAN INVESTASI"
    pe1.font.name = "Times New Roman"
    pe1.font.size = Pt(15)
    pe1.font.bold = True
    pe1.font.color.rgb = GOLD_ACCENT

    pe2 = tf_end.add_paragraph()
    pe2.alignment = PP_ALIGN.CENTER
    pe2.text = "“ BISNIS DINYATAKAN SANGAT LAYAK ”"
    pe2.font.name = "Times New Roman"
    pe2.font.size = Pt(30)
    pe2.font.bold = True
    pe2.font.color.rgb = GREEN_ACCENT

    pe3 = tf_end.add_paragraph()
    pe3.alignment = PP_ALIGN.CENTER
    pe3.text = f"Berdasarkan hasil analisis komprehensif aspek pasar, manajemen, teknis operasi, dan finansial,\n" \
               f"proyek Bakmi Mimu Carina Sayang menghasilkan NPV Rp {met['npv']:,.0f}".replace(",", ".") + \
               f", IRR > 20,00%, Payback Period {met['pp_months']:.2f} bulan (hanya 0,07 tahun), dan PI {met['pi']:.2f}."
    pe3.font.name = "Times New Roman"
    pe3.font.size = Pt(11.5)
    pe3.font.color.rgb = TEXT_MAIN

    pe4 = tf_end.add_paragraph()
    pe4.alignment = PP_ALIGN.CENTER
    pe4.text = "\nDisusun Oleh Tim Mahasiswa:\n" \
               "Arthur Reezan (312023002)   •   Jennese Putra Alamsyah Sukadi (312023033)   •   Valendrik Dwiputra Wirawan (312023013)\n" \
               "Affandy (312023075)   •   Steven Putra Tjhin (312023015)"
    pe4.font.name = "Times New Roman"
    pe4.font.size = Pt(10.5)
    pe4.font.bold = True
    pe4.font.color.rgb = NAVY_PRIMARY

    pe5 = tf_end.add_paragraph()
    pe5.alignment = PP_ALIGN.CENTER
    pe5.text = "\nSekian & Terima Kasih  |  Sesi Tanya Jawab Dibuka\n" \
               "Program Studi Manajemen  |  Fakultas Ekonomi & Bisnis  |  Universitas Kristen Krida Wacana"
    pe5.font.name = "Times New Roman"
    pe5.font.size = Pt(10)
    pe5.font.italic = True
    pe5.font.color.rgb = TEXT_MUTED

    # Save to both final destinations
    prs.save(out_pptx_1)
    prs.save(out_pptx_2)
    print(f"[SUCCESS] Presentasi Eksekutif FEB UKRIDA berhasil dibuat di:")
    print(f"  1. {out_pptx_1}")
    print(f"  2. {out_pptx_2}")

if __name__ == "__main__":
    create_presentation()
