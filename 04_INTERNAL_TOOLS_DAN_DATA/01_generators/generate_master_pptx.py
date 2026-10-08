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

    TOTAL_SLIDES = 13

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_LIGHT
        bg.line.fill.background()

        # Top Accent Ribbon Line
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = GOLD_ACCENT
        top_bar.line.fill.background()

    def add_header(slide, title_text, subtitle_text=None, tag="STUDI KELAYAKAN BISNIS", slide_idx=1):
        # Section Tag Pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.8), Inches(0.35))
        pill.fill.solid()
        pill.fill.fore_color.rgb = GOLD_BG_LIGHT
        pill.line.color.rgb = GOLD_ACCENT
        pill.line.width = Pt(1.0)
        tf_p = pill.text_frame
        tf_p.word_wrap = True
        tf_p.margin_left = tf_p.margin_right = tf_p.margin_top = tf_p.margin_bottom = 0
        p_tag = tf_p.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        p_tag.text = f"• {tag.upper()}"
        p_tag.font.name = "Times New Roman"
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = GOLD_ACCENT

        # Slide Number Badge
        s_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.1), Inches(0.35), Inches(1.433), Inches(0.35))
        s_badge.fill.solid()
        s_badge.fill.fore_color.rgb = CARD_BG_SOFT
        s_badge.line.color.rgb = BORDER_MUTED
        s_badge.line.width = Pt(1.0)
        tf_sb = s_badge.text_frame
        tf_sb.word_wrap = True
        tf_sb.margin_left = tf_sb.margin_right = tf_sb.margin_top = tf_sb.margin_bottom = 0
        p_sb = tf_sb.paragraphs[0]
        p_sb.alignment = PP_ALIGN.CENTER
        p_sb.text = f"SLIDE {slide_idx:02d} / {TOTAL_SLIDES:02d}"
        p_sb.font.name = "Times New Roman"
        p_sb.font.size = Pt(10)
        p_sb.font.bold = True
        p_sb.font.color.rgb = NAVY_PRIMARY

        # Slide Title & Subtitle Box
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
            p_sub.font.size = Pt(11.5)
            p_sub.font.color.rgb = TEXT_MUTED

        # Subtle Divider
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.58), Inches(11.733), Inches(0.02))
        div.fill.solid()
        div.fill.fore_color.rgb = BORDER_MUTED
        div.line.fill.background()

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_MUTED, border_width=1.0):
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
        tf.margin_left = Inches(0.16)
        tf.margin_right = Inches(0.16)
        tf.margin_top = Inches(0.14)
        tf.margin_bottom = Inches(0.14)
        return card, tf

    def style_table_cell(cell, text, font_size=11, bold=False, color=TEXT_MAIN, align=PP_ALIGN.LEFT, bg_color=None):
        cell.text = text
        tf = cell.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.10)
        tf.margin_right = Inches(0.10)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.08)
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

    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(6.066), Inches(0.70), Inches(1.2), Inches(1.2))

    tb_cov = s1.shapes.add_textbox(Inches(1.0), Inches(1.98), Inches(11.333), Inches(2.6))
    tf_cov = tb_cov.text_frame
    tf_cov.word_wrap = True
    tf_cov.margin_left = tf_cov.margin_right = tf_cov.margin_top = tf_cov.margin_bottom = 0

    p_c1 = tf_cov.paragraphs[0]
    p_c1.alignment = PP_ALIGN.CENTER
    p_c1.text = "FAKULTAS EKONOMI & BISNIS  •  UNIVERSITAS KRISTEN KRIDA WACANA"
    p_c1.font.name = "Times New Roman"
    p_c1.font.size = Pt(12)
    p_c1.font.bold = True
    p_c1.font.color.rgb = GOLD_ACCENT

    p_c2 = tf_cov.add_paragraph()
    p_c2.alignment = PP_ALIGN.CENTER
    p_c2.text = "LAPORAN PRESENTASI STUDI KELAYAKAN BISNIS (SKB)"
    p_c2.font.name = "Times New Roman"
    p_c2.font.size = Pt(15)
    p_c2.font.color.rgb = TEXT_MUTED

    p_c3 = tf_cov.add_paragraph()
    p_c3.alignment = PP_ALIGN.CENTER
    p_c3.text = f"“ {cfg['business_name'].upper()} ”"
    p_c3.font.name = "Times New Roman"
    p_c3.font.size = Pt(34)
    p_c3.font.bold = True
    p_c3.font.color.rgb = NAVY_PRIMARY

    p_c4 = tf_cov.add_paragraph()
    p_c4.alignment = PP_ALIGN.CENTER
    p_c4.text = f"- {cfg['tagline']} -"
    p_c4.font.name = "Times New Roman"
    p_c4.font.size = Pt(14)
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
    card_y = Inches(4.68)
    card_h = Inches(1.28)

    for i, (name, nim) in enumerate(authors):
        x = start_x + i * (card_w + card_gap)
        c_m, tf_m = add_card(s1, x, card_y, card_w, card_h, bg_color=CARD_BG_SOFT, border_color=BORDER_MUTED)
        
        pm1 = tf_m.paragraphs[0]
        pm1.alignment = PP_ALIGN.CENTER
        pm1.text = f"Mahasiswa {i+1}"
        pm1.font.name = "Times New Roman"
        pm1.font.size = Pt(10)
        pm1.font.bold = True
        pm1.font.color.rgb = GOLD_ACCENT

        pm2 = tf_m.add_paragraph()
        pm2.alignment = PP_ALIGN.CENTER
        pm2.text = name
        pm2.font.name = "Times New Roman"
        pm2.font.size = Pt(10.5)
        pm2.font.bold = True
        pm2.font.color.rgb = NAVY_PRIMARY

        pm3 = tf_m.add_paragraph()
        pm3.alignment = PP_ALIGN.CENTER
        pm3.text = f"NIM: {nim}"
        pm3.font.name = "Times New Roman"
        pm3.font.size = Pt(10)
        pm3.font.color.rgb = TEXT_MUTED

    tb_foot = s1.shapes.add_textbox(Inches(1.0), Inches(6.12), Inches(11.333), Inches(0.4))
    tf_foot = tb_foot.text_frame
    tf_foot.word_wrap = True
    p_f = tf_foot.paragraphs[0]
    p_f.alignment = PP_ALIGN.CENTER
    p_f.text = "Program Studi Sarjana Manajemen  |  Konsentrasi Studi Kelayakan Bisnis  |  Tahun Akademik 2024"
    p_f.font.name = "Times New Roman"
    p_f.font.size = Pt(11)
    p_f.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 2: DAFTAR ISI & SISTEMATIKA KAJIAN (AGENDA 6 BAB)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Daftar Isi & Sistematika Kajian", "Ringkasan 6 bab utama studi kelayakan bisnis berstandar akademik FEB UKRIDA", tag="AGENDA PRESENTASI", slide_idx=2)

    toc_items = [
        ("01", "BAB I: PENDAHULUAN & PROFIL", "Perjalanan usaha 1999 Taman Aries -> 2016 Duri Kosambi, basis pelanggan 2019, peluang ekspansi 2027/28, & 5 pilar diferensiasi."),
        ("02", "BAB II: ASPEK PASAR & PEMASARAN", "Segmentasi STP, bauran pemasaran terpadu 7P lengkap, matriks komparasi kompetitor, & keunggulan 3 racikan minyak khas."),
        ("03", "BAB III: ASPEK MANAJEMEN & SDM", "Struktur organisasi owner & 3 staf fungsional, rincian job description operasional, & total anggaran upah pasti Rp 10 Jt/bulan."),
        ("04", "BAB IV: ASPEK TEKNIS & OPERASI", "Tata letak ruko 1 lantai 23 kursi, alur open kitchen teras, adonan mie mingguan tanpa pengawet, & timeline pra-operasi 12 minggu."),
        ("05", "BAB V: ASPEK KEUANGAN TERPADU", f"Modal awal Rp {met['initial_outlay']/1e6:.1f} Jt, 12 item capex Rp {met['capex_total']/1e6:.1f} Jt, depresiasi 20%, & proyeksi arus kas 5 tahun (CIF, COF, NCF)."),
        ("06", "BAB VI: KELAYAKAN INVESTASI", f"Uji 4 kriteria: NPV Positif Rp {met['npv']/1e9:.2f} M, IRR > 20%, Payback Period {met['pp_months']:.2f} Bulan (cepat), & PI {met['pi']:.2f}x.")
    ]

    for idx, (num, title, desc) in enumerate(toc_items):
        col = idx % 3
        row = idx // 3
        x = Inches(0.8 + col * 3.98)
        y = Inches(1.75 + row * 2.65)
        w = Inches(3.78)
        h = Inches(2.50)

        card_t, tf_t = add_card(s2, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_MUTED)
        
        p1 = tf_t.paragraphs[0]
        p1.text = f"BAB {num}"
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_t.add_paragraph()
        p2.text = title
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_t.add_paragraph()
        p3.text = f"\n{desc}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(11)
        p3.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 3: PROFIL USAHA, HISTORI 1999–2028 & 5 PILAR DIFERENSIASI (BAB I)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Profil Usaha & Keunggulan Kompetitif", "Perjalanan kuliner legendaris sejak 1999, diferensiasi produk, dan peluang ekspansi masa depan", tag="BAB I: PENDAHULUAN", slide_idx=3)

    # Top: 4 Timeline Cards
    story_steps = [
        ("1999", "Awal di Taman Aries", "Usaha didirikan pertama kali di Taman Aries (Meruya) dengan resep autentik keluarga turun-temurun."),
        ("2016", "Relokasi ke Duri Kosambi", "Tahun 2016 berpindah ke ruko Jl. Angsoka Hijau IV Duri Kosambi persis di depan Kalam Kudus."),
        ("2019", "Basis Pelanggan Solid", "Basis pelanggan setia terbentuk kuat favorit sarapan dan makan siang keluarga perumahan."),
        ("2027/28", "Peluang Ekspansi Cabang", "Tingginya permintaan membuka peluang ekspansi 2027/2028 untuk ruko lebih luas atau cabang baru.")
    ]
    for idx, (yr, ttl, dsc) in enumerate(story_steps):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(1.75)
        w = Inches(2.78)
        h = Inches(1.75)

        card_s, tf_s = add_card(s3, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 3 else BORDER_MUTED)
        p1 = tf_s.paragraphs[0]
        p1.text = yr
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_s.add_paragraph()
        p2.text = ttl
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(11.5)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_s.add_paragraph()
        p3.text = dsc
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.color.rgb = TEXT_MAIN

    # Middle: Banner Filosofi Mutu Adonan Mingguan
    card_ph, tf_ph = add_card(s3, Inches(0.8), Inches(3.65), Inches(11.733), Inches(0.90), bg_color=GOLD_BG_LIGHT, border_color=GOLD_ACCENT)
    pph1 = tf_ph.paragraphs[0]
    pph1.text = "FILOSOFI MUTU: ADONAN MIE FRESH DIBUAT MANDIRI SETIAP MINGGU (BATCH MINGGUAN TANPA PENGAWET)"
    pph1.font.name = "Times New Roman"
    pph1.font.size = Pt(11.5)
    pph1.font.bold = True
    pph1.font.color.rgb = GOLD_ACCENT
    pph2 = tf_ph.add_paragraph()
    pph2.text = "Adonan mie dibuat mandiri setiap minggu menghasilkan tekstur kenyal alami berkilau (shining), higienis, serta 100% bebas dari formalin dan zat pengawet kimia berbahaya."
    pph2.font.name = "Times New Roman"
    pph2.font.size = Pt(10.5)
    pph2.font.color.rgb = TEXT_MAIN

    # Bottom: 5 Pilar Keunggulan Kompetitif
    pillars = [
        ("01. TASTE (RASA)", "Mie kenyal, 3 racikan minyak khas (babi, ayam, sayur+wijen), kaldu gurih asli, 3 topping murni: ayam putih, babi kecap, casiu madu."),
        ("02. LOKASI STRATEGIS", "Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi (Depan Taman TK, SD, SMP Kalam Kudus), lokasi strategis & parkir motor aman."),
        ("03. OPEN KITCHEN", "Dapur terbuka gerobak di teras depan, perebusan higienis terlihat jelas, serta saus botol resmi Cap Belibis & Mangga Besar."),
        ("04. PORSI FLEKSIBEL", "Pilihan fleksibel: Porsi Kecil (27k), Reguler (29k-39k), hingga Jumbo (+100% mie / 2x lipat porsi standar 49k-59k) untuk segala selera."),
        ("05. PELENGKAP LENGKAP", "Swikiaw & pangsit kuah (@4.5k-5.5k), baso goreng babi-udang 8k, badak sarsaparilla, susu kacang & liang teh segar 12k.")
    ]
    for idx, (tit, dsc) in enumerate(pillars):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(4.70)
        w = Inches(2.20)
        h = Inches(2.40)

        c_p, tf_p = add_card(s3, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_MUTED)
        p1 = tf_p.paragraphs[0]
        p1.text = tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        p2 = tf_p.add_paragraph()
        p2.text = f"\n{dsc}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 4: DAFTAR MENU & STRUKTUR HARGA JUAL RESMI (BAB I)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Daftar Menu & Struktur Harga Jual Resmi", "Struktur harga bersaing (Value-Based Pricing) sesuai daftar menu riil kedai di Duri Kosambi", tag="BAB I: PENDAHULUAN", slide_idx=4)

    # Left Table: Menu Bakmi Utama
    tb_shape1 = s4.shapes.add_table(7, 3, Inches(0.8), Inches(1.75), Inches(5.7), Inches(4.35))
    tbl1 = tb_shape1.table
    tbl1.columns[0].width = Inches(2.6)
    tbl1.columns[1].width = Inches(1.4)
    tbl1.columns[2].width = Inches(1.7)

    style_table_cell(tbl1.cell(0, 0), "Menu Utama (Topping Murni)", font_size=12, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl1.cell(0, 1), "Harga", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl1.cell(0, 2), "Keterangan Porsi", font_size=12, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)

    m1_data = [
        ("Mie Campur (Ayam + Babi Kecap)", "Rp 29.000", "Reguler Favorit"),
        ("Mie Ayam Putih / Babi Saja", "Rp 29.000", "Topping Tunggal Murni"),
        ("Mie Daging Casiu Madu", "Rp 31.000", "Casiu Panggang Gurih"),
        ("Mie Spesial 3 Isi Lengkap", "Rp 39.000", "Ayam, Babi & Casiu"),
        ("Porsi Jumbo (+100% Mie)", "Rp 49.000 - 59.000", "2x Lipat Porsi Standar"),
        ("Porsi Kecil (Sarapan Hemat)", "Rp 27.000 - 29.000", "Porsi Ringan Anak/Pagi")
    ]
    for r_idx, (m_n, m_h, m_k) in enumerate(m1_data, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl1.cell(r_idx, 0), m_n, font_size=11, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl1.cell(r_idx, 1), m_h, font_size=11.5, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl1.cell(r_idx, 2), m_k, font_size=10.5, color=TEXT_MUTED, bg_color=bg_r)

    # Right Table: Menu Pelengkap & Minuman
    tb_shape2 = s4.shapes.add_table(7, 3, Inches(6.833), Inches(1.75), Inches(5.7), Inches(4.35))
    tbl2 = tb_shape2.table
    tbl2.columns[0].width = Inches(2.6)
    tbl2.columns[1].width = Inches(1.4)
    tbl2.columns[2].width = Inches(1.7)

    style_table_cell(tbl2.cell(0, 0), "Pelengkap, Kuah & Minuman", font_size=12, bold=True, color=WHITE, bg_color=NAVY_LIGHT)
    style_table_cell(tbl2.cell(0, 1), "Harga", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_LIGHT)
    style_table_cell(tbl2.cell(0, 2), "Keterangan Item", font_size=12, bold=True, color=WHITE, bg_color=NAVY_LIGHT)

    m2_data = [
        ("Pangsit Rebus Kuah (5 pcs)", "Rp 22.500", "@Rp 4.500 / pcs"),
        ("Swikiaw Rebus Kuah (5 pcs)", "Rp 27.500", "@Rp 5.500 / pcs (Babi-Udang)"),
        ("Baso Sapi & Baso Ikan (5 pcs)", "Rp 22.500", "Sama harga pangsit"),
        ("Pangsit Goreng & Baso Goreng", "Rp 5.000 - 8.000", "Baso grg babi-udang 8k"),
        ("Badak Sarsaparilla (Cap Badak)", "Rp 12.000 - 14.000", "Soda legendaris Siantar"),
        ("Susu Kacang, Liang Teh, Teh", "Rp 1.000 - 12.000", "Saus Belibis & Mangga Bsr")
    ]
    for r_idx, (m_n, m_h, m_k) in enumerate(m2_data, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl2.cell(r_idx, 0), m_n, font_size=11, bold=True, color=NAVY_LIGHT, bg_color=bg_r)
        style_table_cell(tbl2.cell(r_idx, 1), m_h, font_size=11.5, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl2.cell(r_idx, 2), m_k, font_size=10.5, color=TEXT_MUTED, bg_color=bg_r)

    # Bottom Banner Quality Guarantee
    card_bq, tf_bq = add_card(s4, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.85), bg_color=GREEN_BG_LIGHT, border_color=GREEN_ACCENT)
    pbq1 = tf_bq.paragraphs[0]
    pbq1.text = "STANDAR KUALITAS BAHAN: 100% DAGING MURNI PILIHAN (TANPA BEBEK & JAMUR)"
    pbq1.font.name = "Times New Roman"
    pbq1.font.size = Pt(11.5)
    pbq1.font.bold = True
    pbq1.font.color.rgb = GREEN_ACCENT
    pbq2 = tf_bq.add_paragraph()
    pbq2.text = "Seluruh menu diolah fresh setiap hari tanpa formalin. Disajikan lengkap dengan 3 racikan minyak khas (minyak babi, ayam, atau sayur+wijen) sesuai selera pengunjung."
    pbq2.font.name = "Times New Roman"
    pbq2.font.size = Pt(10.5)
    pbq2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 5: ANALISIS PASAR: STP & PETA PERSAINGAN KOMPETITOR (BAB II)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Analisis Pasar: STP & Peta Kompetitor", "Segmentasi pasar sasaran serta matriks komparasi daya saing terhadap kompetitor sekitar", tag="BAB II: ASPEK PASAR", slide_idx=5)

    # Left: STP Framework (3 Cards)
    stp_items = [
        ("SEGMENTASI (SEGMENTING)", [
            "Geografis: Radius 0 - 5 km meliputi Perumahan Duri Kosambi, Semanan, TSI, & Cengkareng.",
            "Demografis: Pria & wanita usia 7 - 65 tahun, siswa-siswi, guru, orang tua Kalam Kudus, SES Menengah.",
            "Perilaku: Pencinta bakmi oriental yang mengutamakan tekstur kenyal, kaldu gurih, & higienitas."
        ]),
        ("TARGET PASAR (TARGETING)", [
            "Keluarga Residensial: Keluarga perumahan sekitar Duri Kosambi untuk sarapan & makan siang.",
            "Civitas Kalam Kudus: Murid, guru, dan orang tua Sekolah Kristen Kalam Kudus di depan kedai.",
            "Komunitas Sekitar: Jemaat gereja hari Sabtu-Minggu serta pelanggan bawa pulang / take-away."
        ]),
        ("POSISI PASAR (POSITIONING)", [
            "Identitas: “Kedai Bakmi Otentik Legendaris Resep 1999 dan Standar Kebersihan Higienis Terpercaya”.",
            "Value Proposition: Memberikan kepuasan rasa autentik resep keluarga dengan porsi mengenyangkan.",
            "Brand Perception: Kedai bakmi keluarga yang bersih, ramah, bersahaja, dan bernilai sebanding."
        ])
    ]
    for idx, (tit, bullets) in enumerate(stp_items):
        y = Inches(1.75 + idx * 1.75)
        c_s, tf_s = add_card(s5, Inches(0.8), y, Inches(5.7), Inches(1.65), bg_color=CARD_BG, border_color=BORDER_NAVY if idx == 2 else BORDER_MUTED)
        p1 = tf_s.paragraphs[0]
        p1.text = tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        for b in bullets:
            pb = tf_s.add_paragraph()
            pb.text = f"• {b}"
            pb.font.name = "Times New Roman"
            pb.font.size = Pt(9.5)
            pb.font.color.rgb = TEXT_MAIN

    # Right: Komparasi 3 Kompetitor (3 Cards)
    comp_items = [
        ("BAKMI MIMU CARINA SAYANG (BRAND UNGGULAN)", GOLD_ACCENT, [
            "Rasa & Resep: Resep 1999, 3 racikan minyak khas, adonan fresh mingguan.",
            "Fasilitas: Ruko 1 lantai 23 kursi, 6 kipas angin dinding sejuk, open kitchen teras.",
            "Topping & Saus: 3 Topping murni (ayam, babi, casiu), kuah swikiaw, saus Belibis.",
            "Harga & Value: Rp 27.000 - Rp 59.000 (Opsi porsi kecil, reguler, hingga jumbo)."
        ]),
        ("BAKMI ALOK / BRAND BESAR (KOMPETITOR TERKENAL)", NAVY_LIGHT, [
            "Rasa & Resep: Dominan ayam rebus gurih, brand equity kuat Jakarta Barat.",
            "Fasilitas: Restoran permanen ber-AC, kapasitas besar, antrean jam makan siang.",
            "Topping & Saus: Fokus ayam kampung rebus, tidak menyediakan varian babi casiu.",
            "Harga & Value: Rp 45.000 - Rp 65.000 (Segmen premium, relatif lebih mahal)."
        ]),
        ("WARUNG BAKMI LOKAL KOSAMBI (TRADISIONAL)", TEXT_MUTED, [
            "Rasa & Resep: Standar gerobak kaki lima, bumbu penyedap dominan.",
            "Fasilitas: Kios tenda pinggir jalan, sirkulasi udara terbatas, parkir sempit.",
            "Topping & Saus: Topping terbatas ayam cincang biasa, jarang menyediakan swikiaw.",
            "Harga & Value: Rp 20.000 - Rp 25.000 (Murah namun higienitas & fasilitas minim)."
        ])
    ]
    for idx, (b_name, b_col, pts) in enumerate(comp_items):
        y = Inches(1.75 + idx * 1.75)
        c_c, tf_c = add_card(s5, Inches(6.833), y, Inches(5.7), Inches(1.65), bg_color=CARD_BG, border_color=b_col, border_width=1.5 if idx == 0 else 1.0)
        p1 = tf_c.paragraphs[0]
        p1.text = b_name
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = b_col

        for p_t in pts:
            pp = tf_c.add_paragraph()
            pp.text = f"▸ {p_t}"
            pp.font.name = "Times New Roman"
            pp.font.size = Pt(9.5)
            pp.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 6: STRATEGI BAURAN PEMASARAN TERPADU (MARKETING MIX 7P LENGKAP) (BAB II)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Strategi Bauran Pemasaran Terpadu (Marketing Mix 7P)", "Penerapan bauran pemasaran menyeluruh: 4P Tradisional dan 3P Extended Layanan", tag="BAB II: ASPEK PASAR", slide_idx=6)

    # Top Row: 4P Tradisional
    p4_cards = [
        ("PRODUCT (PRODUK)", GOLD_ACCENT, [
            "Resep autentik 1999 (Taman Aries), adonan fresh dibuat mandiri setiap minggu tanpa formalin.",
            "3 Topping murni: Ayam Putih, Babi Kecap, & Casiu Madu gurih (tanpa bebek & jamur).",
            "3 Pilihan racikan minyak: Minyak Babi, Minyak Ayam, & Minyak Sayur + Wijen.",
            "Porsi fleksibel: Porsi Kecil (27k), Reguler (29k-39k), hingga Jumbo (+100% mie, 49k-59k)."
        ]),
        ("PRICE (HARGA)", GREEN_ACCENT, [
            "Menerapkan Value-Based Pricing bersaing di kawasan residensial Duri Kosambi.",
            "Range bakmi: Rp 27.000 s.d. Rp 59.000 (porsi kecil, reguler, hingga jumbo kenyang).",
            "Kuah pelengkap: Pangsit rebus Rp 22.500 (5 pcs), Swikiaw Rp 27.500 (5 pcs), Baso Rp 22.500.",
            "Gorengan Rp 5.000 - 8.000, minuman segar Rp 1.000 - 14.000 (Badak, Liang Teh, Susu Kacang)."
        ]),
        ("PLACE (LOKASI)", NAVY_LIGHT, [
            "Ruko 1 lantai di Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi, Cengkareng, Jakarta Barat.",
            "Posisi emas persis di depan Taman TK, SD, dan SMP Sekolah Kristen Kalam Kudus.",
            "Daya tampung ruang makan 23 kursi aktif (3 meja reguler 12 kursi + meja dinding 7 kursi + teras 4 kursi).",
            "Saluran penjualan langsung: Makan di tempat (Dine-In) dan pesanan bawa pulang (Take-Away)."
        ]),
        ("PROMOTION (PROMOSI)", NAVY_PRIMARY, [
            "Murni mengandalkan pemasaran lokal getok tular (loyalitas pelanggan setia sejak 2019).",
            "Pemasaran organik sukarela dari ulasan pengunjung dan food vlogger kuliner di media sosial.",
            "Papan nama kedai (signage) sederhana di fasad ruko yang terlihat jelas oleh penjemput sekolah.",
            "Tanpa biaya diskon buatan sehingga menjaga citra mutu autentik dan stabilitas margin laba."
        ])
    ]
    for idx, (title, col, bullets) in enumerate(p4_cards):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(1.75)
        w = Inches(2.78)
        h = Inches(2.55)

        c_4p, tf_4p = add_card(s6, x, y, w, h, bg_color=CARD_BG, border_color=col)
        p1 = tf_4p.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = col

        for b in bullets:
            pb = tf_4p.add_paragraph()
            pb.text = f"• {b}"
            pb.font.name = "Times New Roman"
            pb.font.size = Pt(9.5)
            pb.font.color.rgb = TEXT_MAIN

    # Bottom Row: 3P Extended Layanan
    p3_cards = [
        ("PEOPLE (SUMBER DAYA MANUSIA)", GOLD_ACCENT, [
            "Tim operasional 3 karyawan tetap: Koki Utama Dapur, Asisten Koki Dapur, & Kasir-Pramusaji.",
            "Sistem remunerasi berbasis upah bulanan pasti (total alokasi gaji tetap Rp 10.000.000 per bulan).",
            "Fokus tim diarahkan penuh pada konsistensi racikan rasa resep 1999 dan higienitas sajian.",
            "Karyawan berseragam bersih, ramah menyambut tamu, dan menguasai varian menu topping secara rinci."
        ]),
        ("PROCESS (PROSES OPERASIONAL)", NAVY_LIGHT, [
            "Sistem pemesanan kasir menggunakan nota fisik manual yang tertib, cepat, dan transparan.",
            "Alur penyajian kilat: Perebusan mie fresh membutuhkan 45 detik, saji di meja pelanggan < 5 menit.",
            "Siklus produksi adonan mie dibuat mandiri terjadwal setiap minggu tanpa bahan pengawet kimia.",
            "Penyediaan botol saus meja standar resmi (Cap Belibis & Mangga Besar) yang selalu terisi higienis."
        ]),
        ("PHYSICAL EVIDENCE (BUKTI FISIK)", GREEN_ACCENT, [
            "Desain interior standar ruko 1 lantai yang bersih, bersahaja, rapi, dan sirkulasi udara alami.",
            "Fasilitas pendingin udara menggunakan 6 unit kipas angin dinding/plafon sejuk bebas rasa pengap.",
            "Kapasitas dine-in 23 kursi aktif (meja reguler kayu, meja bar dinding memanjang, dan meja teras).",
            "Fasilitas sanitasi wastafel cuci tangan higienis dan toilet bersih tertutup di sudut belakang kedai."
        ])
    ]
    for idx, (title, col, bullets) in enumerate(p3_cards):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(4.45)
        w = Inches(3.78)
        h = Inches(2.55)

        c_3p, tf_3p = add_card(s6, x, y, w, h, bg_color=CARD_BG, border_color=col)
        p1 = tf_3p.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = col

        for b in bullets:
            pb = tf_3p.add_paragraph()
            pb.text = f"▸ {b}"
            pb.font.name = "Times New Roman"
            pb.font.size = Pt(10)
            pb.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 7: ASPEK MANAJEMEN & SDM: STRUKTUR, JOBDESK & GAJI (BAB III)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Aspek Manajemen & SDM: Struktur, Jobdesk & Gaji", "Struktur tata kelola fungsional, pembagian job description, dan kebijakan kompensasi upah", tag="BAB III: ASPEK MANAJEMEN", slide_idx=7)

    # Top: Owner Card & Anggaran Gaji Box
    c_m, tf_m = add_card(s7, Inches(0.8), Inches(1.75), Inches(5.7), Inches(1.55), bg_color=CARD_BG_SOFT, border_color=GOLD_ACCENT, border_width=1.5)
    pm1 = tf_m.paragraphs[0]
    pm1.text = "OWNER / PENGELOLA UTAMA KEDAI (1 ORANG)"
    pm1.font.name = "Times New Roman"
    pm1.font.size = Pt(12.5)
    pm1.font.bold = True
    pm1.font.color.rgb = NAVY_PRIMARY
    pm2 = tf_m.add_paragraph()
    pm2.text = "Tanggung Jawab: Penetapan strategi usaha, pengawasan kas harian, kontrol cita rasa resep 1999, pengadaan bahan baku daging segar, dan perencanaan ekspansi cabang 2027/2028."
    pm2.font.name = "Times New Roman"
    pm2.font.size = Pt(10.5)
    pm2.font.color.rgb = TEXT_MAIN

    c_g, tf_g = add_card(s7, Inches(6.833), Inches(1.75), Inches(5.7), Inches(1.55), bg_color=CARD_BG_SOFT, border_color=GREEN_ACCENT, border_width=1.5)
    pg1 = tf_g.paragraphs[0]
    pg1.text = "TOTAL ANGGARAN GAJI: RP 10.000.000 / BULAN (RP 120 JT / TAHUN)"
    pg1.font.name = "Times New Roman"
    pg1.font.size = Pt(12.5)
    pg1.font.bold = True
    pg1.font.color.rgb = GREEN_ACCENT
    pg2 = tf_g.add_paragraph()
    pg2.text = "Kebijakan Kompensasi: Menerapkan sistem upah bulanan pasti tanpa insentif fluktuatif, jatah makan kedai harian higienis, serta Tunjangan Hari Raya (THR) tahunan untuk 3 staf tetap."
    pg2.font.name = "Times New Roman"
    pg2.font.size = Pt(10.5)
    pg2.font.color.rgb = TEXT_MAIN

    # Bottom: 3 Kolom Karyawan Tetap
    sub_roles = [
        ("KOKI UTAMA (1 ORANG)", "Gaji: Rp 3.800.000 / Bulan", NAVY_LIGHT, [
            "Membuat adonan mie fresh setiap minggu tanpa bahan pengawet.",
            "Memasak 3 topping: ayam putih gurih, babi kecap manis, casiu madu.",
            "Menyiapkan 3 racikan minyak khas & rebus kaldu tulang gurih.",
            "Memimpin perebusan kilat 45 detik & peracikan bumbu di gerobak open kitchen."
        ]),
        ("ASISTEN KOKI (1 ORANG)", "Gaji: Rp 3.200.000 / Bulan", GREEN_ACCENT, [
            "Menyiapkan bahan mentah sayur & bumbu rempah dapur.",
            "Melipat kulit pangsit & swikiaw isian babi-udang segar setiap hari.",
            "Mencuci mangkok keramik, piring, sumpit, panci & alat makan.",
            "Menjaga kebersihan dan sanitasi stasiun gerobak open kitchen teras."
        ]),
        ("KASIR & PRAMUSAJI (1 ORANG)", "Gaji: Rp 3.000.000 / Bulan", GOLD_ACCENT, [
            "Mencatat pesanan pelanggan dengan nota manual tertib dan cepat.",
            "Melayani penerimaan pembayaran tunai & transaksi nontunai QRIS.",
            "Menyajikan mangkok mie & pesanan minuman ke 23 kursi tamu.",
            "Membersihkan meja santap secara berkala & melayani pesanan take-away."
        ])
    ]
    for idx, (r_title, r_sal, r_col, r_tasks) in enumerate(sub_roles):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(3.45)
        w = Inches(3.78)
        h = Inches(3.60)

        card_sb, tf_sb = add_card(s7, x, y, w, h, bg_color=CARD_BG, border_color=r_col)
        p1 = tf_sb.paragraphs[0]
        p1.text = r_title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12.5)
        p1.font.bold = True
        p1.font.color.rgb = r_col

        p2 = tf_sb.add_paragraph()
        p2.text = r_sal
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = TEXT_MUTED

        for t in r_tasks:
            pt = tf_sb.add_paragraph()
            pt.text = f"\n• {t}"
            pt.font.name = "Times New Roman"
            pt.font.size = Pt(10)
            pt.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 8: ASPEK TEKNIS: DENAH TATA LETAK KEDAI RUKO 1 LANTAI (23 KURSI) (BAB IV)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Denah Tata Letak Kedai Ruko 1 Lantai", "Visualisasi denah fisik arsitektural, stasiun dapur terbuka teras (open kitchen), dan ruang makan (23 kursi)", tag="BAB IV: TEKNIS & OPERASI", slide_idx=8)

    # Top Card: Wadah Denah Arsitektural
    c_denah, tf_denah = add_card(s8, Inches(0.8), Inches(1.75), Inches(11.733), Inches(3.18), bg_color=CARD_BG, border_color=NAVY_PRIMARY)
    p_d0 = tf_denah.paragraphs[0]
    p_d0.text = "DENAH ARSITEKTURAL TATA LETAK FISIK KEDAI RUKO (ALUR SATU ARAH)"
    p_d0.font.name = "Times New Roman"
    p_d0.font.size = Pt(11.5)
    p_d0.font.bold = True
    p_d0.font.color.rgb = NAVY_PRIMARY

    p_d1 = tf_denah.add_paragraph()
    p_d1.text = "Alur Kerja: Pintu Masuk ➔ Open Kitchen Teras ➔ Rolling Door ➔ Ruang Santap Utama (Meja Reguler & Dinding) ➔ Fasilitas Sanitasi Belakang"
    p_d1.font.name = "Times New Roman"
    p_d1.font.size = Pt(9.5)
    p_d1.font.color.rgb = TEXT_MUTED

    if os.path.exists(layout_path):
        img_w = Inches(9.20)
        img_h = Inches(2.35)
        img_left = Inches(0.8) + (Inches(11.733) - img_w) / 2
        img_top = Inches(2.45)
        s8.shapes.add_picture(layout_path, img_left, img_top, img_w, img_h)

    # 3 Kartu Zonasi Berdampingan di Bagian Bawah
    c_z1, tf_z1 = add_card(s8, Inches(0.8), Inches(5.08), Inches(3.75), Inches(2.05), bg_color=CARD_BG, border_color=NAVY_PRIMARY)
    p_z1 = tf_z1.paragraphs[0]
    p_z1.text = "ZONA 1: DAPUR TERAS & TERBUKA"
    p_z1.font.name = "Times New Roman"
    p_z1.font.size = Pt(11)
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
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_MAIN

    c_z2, tf_z2 = add_card(s8, Inches(4.79), Inches(5.08), Inches(3.75), Inches(2.05), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    p_z2 = tf_z2.paragraphs[0]
    p_z2.text = "ZONA 2: RUANG MAKAN (DINE-IN)"
    p_z2.font.name = "Times New Roman"
    p_z2.font.size = Pt(11)
    p_z2.font.bold = True
    p_z2.font.color.rgb = GOLD_ACCENT

    z2_items = [
        "Pembatas partisi fleksibel Rolling Door (sekat debu/udara)",
        "3 Meja makan reguler kayu (kapasitas 12 kursi santap)",
        "1 Meja makan dinding memanjang (7 kursi solo diner)",
        "Sirkulasi 6 unit kipas angin dinding sejuk bebas pengap"
    ]
    for it in z2_items:
        p = tf_z2.add_paragraph()
        p.text = f"• {it}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_MAIN

    c_z3, tf_z3 = add_card(s8, Inches(8.78), Inches(5.08), Inches(3.75), Inches(2.05), bg_color=CARD_BG, border_color=NAVY_LIGHT)
    p_z3 = tf_z3.paragraphs[0]
    p_z3.text = "ZONA 3: FASILITAS & SANITASI"
    p_z3.font.name = "Times New Roman"
    p_z3.font.size = Pt(11)
    p_z3.font.bold = True
    p_z3.font.color.rgb = NAVY_LIGHT

    z3_items = [
        "Meja barang penyimpanan stok piring & kemasan bersih",
        "Wastafel cuci tangan higienis bagi pengunjung kedai",
        "1 Unit kamar mandi / toilet ruko tertutup higienis",
        "Total daya tampung santap kedai: 23 kursi pengunjung aktif"
    ]
    for it in z3_items:
        p = tf_z3.add_paragraph()
        p.text = f"• {it}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 9: ALUR PROSES PRODUKSI & JADWAL PRA-OPERASI 12 MINGGU (BAB IV)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Alur Proses Produksi & Jadwal Pra-Operasi", "Tahapan peracikan rasa higienis dan rencana 12 minggu persiapan operasional relokasi 2016", tag="BAB IV: TEKNIS & OPERASI", slide_idx=9)

    # Top Row: 5 Langkah Alur Produksi
    steps_prod = [
        ("LANGKAH 01", "Adonan Fresh Mingguan", "Minggu Pagi", "Dibuat mandiri konsisten setiap minggu tanpa bahan pengawet kimia."),
        ("LANGKAH 02", "Pengolahan 3 Topping", "05.30 - 07.30", "Panggang casiu madu, tumis babi kecap gurih, rebus potongan ayam putih."),
        ("LANGKAH 03", "Racikan 3 Minyak", "06.00 - 08.00", "Racikan minyak babi, minyak ayam, minyak sayur+wijen, & kaldu tulang."),
        ("LANGKAH 04", "Perebusan Seketika", "Saat Order", "Mie fresh direbus tepat 45 detik saat nota pesanan masuk dari kasir."),
        ("LANGKAH 05", "Plating & Penyajian", "< 5 Menit Total", "Pengadukan minyak pilihan, topping daging melimpah, daun bawang & saus.")
    ]
    for idx, (st, tit, tm, dsc) in enumerate(steps_prod):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(1.75)
        w = Inches(2.20)
        h = Inches(2.45)

        c_st, tf_st = add_card(s9, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 3 else BORDER_MUTED)
        p1 = tf_st.paragraphs[0]
        p1.text = st
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf_st.add_paragraph()
        p2.text = tit
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(11.5)
        p2.font.bold = True
        p2.font.color.rgb = NAVY_PRIMARY

        p3 = tf_st.add_paragraph()
        p3.text = f"Waktu: {tm}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.bold = True
        p3.font.color.rgb = NAVY_LIGHT

        p4 = tf_st.add_paragraph()
        p4.text = f"\n{dsc}"
        p4.font.name = "Times New Roman"
        p4.font.size = Pt(9.5)
        p4.font.color.rgb = TEXT_MAIN

    # Bottom Row: Jadwal Pra-Operasi Relokasi 2016 (4 Fase)
    phases = [
        ("FASE 1: SURVEI & SEWA", "Minggu 1 - 4", GOLD_ACCENT, [
            "A. Observasi lokasi Jl. Angsoka Hijau IV depan Kalam Kudus.",
            "B. Negosiasi dan pelunasan sewa ruko 1 lantai Rp 20 Juta.",
            "C. Izin lingkungan RT/RW dan retribusi kebersihan tertib.",
            "D. Analisis potensi pasar sarapan murid & orang tua sekolah."
        ]),
        ("FASE 2: RENOVASI & FIT-OUT", "Minggu 5 - 8", NAVY_LIGHT, [
            "E. Pengecatan ruko 1 lantai dan instalasi pipa saluran air.",
            "F. Pemasangan 6 unit kipas angin dinding sejuk bebas pengap.",
            "G. Penataan teras depan untuk stasiun open kitchen gerobak.",
            "H. Pemasangan spanduk dan plang nama kedai Bakmi Mimu."
        ]),
        ("FASE 3: PENGADAAN & SETTING", "Minggu 9 - 10", GREEN_ACCENT, [
            "I. Pembuatan gerobak etalase kaca stasiun masak (Rp 8 Jt).",
            "J. Pengadaan 4 meja pendek, meja dinding, dan kursi 37 pcs.",
            "K. Pembelian kulkas sayur, freezer daging, panci & mangkok.",
            "L. Pengadaan saus standar resmi (Belibis & Mangga Besar)."
        ]),
        ("FASE 4: TRIAL RUN & BUKA", "Minggu 11 - 12", NAVY_PRIMARY, [
            "M. Rekrutmen 3 karyawan tetap (Koki, Asisten, Kasir-Pramusaji).",
            "N. Uji coba pembuatan adonan mie mingguan & kaldu resep 1999.",
            "O. Simulasi alur pelayanan dine-in 23 kursi dan nota manual.",
            "P. Pembukaan resmi melayani warga perumahan Kosambi."
        ])
    ]
    for idx, (p_tit, p_time, p_col, p_items) in enumerate(phases):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(4.35)
        w = Inches(2.78)
        h = Inches(2.75)

        card_ph, tf_ph = add_card(s9, x, y, w, h, bg_color=CARD_BG, border_color=p_col)
        p1 = tf_ph.paragraphs[0]
        p1.text = p_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = p_col

        p2 = tf_ph.add_paragraph()
        p2.text = f"Jadwal: {p_time}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10)
        p2.font.bold = True
        p2.font.color.rgb = TEXT_MUTED

        for itm in p_items:
            pit = tf_ph.add_paragraph()
            pit.text = f"• {itm}"
            pit.font.name = "Times New Roman"
            pit.font.size = Pt(9.5)
            pit.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 10: RENCANA INVESTASI AWAL & AKTIVA TETAP (CAPEX & DEPRESIASI) (BAB V)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Rencana Modal Investasi Awal & Aktiva Tetap", "Alokasi anggaran modal relokasi 2016 dan inventarisasi 12 item aktiva tetap kedai Duri Kosambi", tag="BAB V: ASPEK KEUANGAN", slide_idx=10)

    # Left: Modal Investasi Awal (Initial Outlay) Box
    c_out, tf_out = add_card(s10, Inches(0.8), Inches(1.75), Inches(5.7), Inches(5.35), bg_color=CARD_BG, border_color=GOLD_ACCENT, border_width=1.5)
    po1 = tf_out.paragraphs[0]
    po1.text = "TOTAL INITIAL OUTLAY: RP 62.000.000"
    po1.font.name = "Times New Roman"
    po1.font.size = Pt(16)
    po1.font.bold = True
    po1.font.color.rgb = GOLD_ACCENT

    po2 = tf_out.add_paragraph()
    po2.text = "Struktur Permodalan: 100% Ekuitas Sendiri Tanpa Utang Bank"
    po2.font.name = "Times New Roman"
    po2.font.size = Pt(11.5)
    po2.font.bold = True
    po2.font.color.rgb = NAVY_PRIMARY

    outlay_items = [
        ("01. Aktiva Tetap (Capex)", "Rp 40.700.000", "65,6%", "Pengadaan gerobak etalase, meja, kursi 37 pcs, 6 kipas angin, kulkas, freezer, panci, & mangkok."),
        ("02. Sewa Tempat Ruko", "Rp 20.000.000", "32,3%", "Alokasi sewa ruko 1 lantai Duri Kosambi untuk 1 tahun penuh masa operasional."),
        ("03. Perizinan & AMDAL", "Rp 300.000", "0,5%", "Retribusi kebersihan lingkungan RT/RW dan sarana tempat sampah kedai higienis."),
        ("04. Biaya Survei Pasar", "Rp 500.000", "0,8%", "Observasi demografi perumahan Kosambi dan potensi kantin sekolah Kalam Kudus."),
        ("05. Biaya Promosi Awal", "Rp 500.000", "0,8%", "Pembuatan spanduk pembukaan dan plang penunjuk arah kedai Bakmi Mimu.")
    ]
    for c_n, c_v, c_p, c_d in outlay_items:
        p = tf_out.add_paragraph()
        p.text = f"\n• {c_n}: {c_v} ({c_p})"
        p.font.name = "Times New Roman"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = NAVY_LIGHT

        pd = tf_out.add_paragraph()
        pd.text = f"   {c_d}"
        pd.font.name = "Times New Roman"
        pd.font.size = Pt(9.5)
        pd.font.color.rgb = TEXT_MAIN

    # Right: Capex Table & Parameter Depresiasi
    tb_shape_c = s10.shapes.add_table(7, 3, Inches(6.833), Inches(1.75), Inches(5.7), Inches(3.85))
    tbl_c = tb_shape_c.table
    tbl_c.columns[0].width = Inches(2.5)
    tbl_c.columns[1].width = Inches(1.3)
    tbl_c.columns[2].width = Inches(1.9)

    style_table_cell(tbl_c.cell(0, 0), "Kelompok Capex (12 Item)", font_size=11, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_c.cell(0, 1), "Nilai", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_c.cell(0, 2), "Rincian Peralatan", font_size=11, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)

    capex_rows = [
        ("1. Gerobak & Etalase Kaca Depan", "Rp 8.000.000", "Stasiun masak open kitchen teras"),
        ("2. Perabotan Santap (Meja & Kursi)", "Rp 8.300.000", "4 Meja pendek, meja bar, kursi 37 pcs"),
        ("3. Sirkulasi Udara (6 Kipas Angin)", "Rp 900.000", "6 Kipas dinding/plafon sejuk @150k"),
        ("4. Pendingin Dapur (Kulkas/Freezer)", "Rp 6.000.000", "1 Kulkas sayur 2.5 Jt, 1 Freezer 3.5 Jt"),
        ("5. Panci Masak & Stok Mangkok", "Rp 6.500.000", "Panci kaldu, mangkok keramik, sendok"),
        ("6. Renovasi Ruko & Plang Nama", "Rp 11.000.000", "Cat ruko 1 lantai, pipa/air, plang nama")
    ]
    for r_idx, (c_g, c_v, c_d) in enumerate(capex_rows, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_c.cell(r_idx, 0), c_g, font_size=10, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_c.cell(r_idx, 1), c_v, font_size=10.5, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_c.cell(r_idx, 2), c_d, font_size=9.5, color=TEXT_MUTED, bg_color=bg_r)

    # Right Bottom: Depresiasi Parameter Box
    c_dp, tf_dp = add_card(s10, Inches(6.833), Inches(5.75), Inches(5.7), Inches(1.35), bg_color=CARD_BG_SOFT, border_color=GREEN_ACCENT)
    pdp1 = tf_dp.paragraphs[0]
    pdp1.text = "PARAMETER DEPRESIASI GARIS LURUS (STRAIGHT-LINE)"
    pdp1.font.name = "Times New Roman"
    pdp1.font.size = Pt(11.5)
    pdp1.font.bold = True
    pdp1.font.color.rgb = GREEN_ACCENT

    pdp2 = tf_dp.add_paragraph()
    pdp2.text = f"Total Perolehan: Rp {met['capex_total']:,.0f}  |  Tarif: 20,00% / Tahun (Masa Manfaat 5 Tahun)\nBeban Depresiasi Tahunan: Rp {met['depresiasi_per_year']:,.0f} / Tahun (Rp 678.333 / Bulan)".replace(",", ".")
    pdp2.font.name = "Times New Roman"
    pdp2.font.size = Pt(10.5)
    pdp2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 11: PROYEKSI KEUANGAN 5 TAHUN: INFLOW, OUTFLOW & LABA BERSIH (BAB V)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Proyeksi Keuangan 5 Tahun: Arus Kas & Laba Bersih", "Estimasi penerimaan kas 300 hari operasi tahunan, biaya operasional, dan kas bersih riil", tag="BAB V: ASPEK KEUANGAN", slide_idx=11)

    # Table 5 Tahun Keuangan
    tb_shape_cof = s11.shapes.add_table(6, 5, Inches(0.8), Inches(1.75), Inches(11.733), Inches(3.60))
    tbl_cof = tb_shape_cof.table
    tbl_cof.columns[0].width = Inches(2.133)
    tbl_cof.columns[1].width = Inches(2.4)
    tbl_cof.columns[2].width = Inches(2.4)
    tbl_cof.columns[3].width = Inches(2.4)
    tbl_cof.columns[4].width = Inches(2.4)

    style_table_cell(tbl_cof.cell(0, 0), "Tahun Operasi", font_size=12, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 1), "Inflow (CIF)", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 2), "Outflow (COF)", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 3), "Net Cash Flow (EAT)", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_cof.cell(0, 4), "Proceed (Kas Riil)", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)

    for i in range(5):
        bg_r = CARD_BG if (i+1) % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_cof.cell(i+1, 0), f"Tahun {i+1} (201{6+i}/1{7+i})", font_size=11, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 1), f"Rp {met['cif_years'][i]:,.0f}".replace(",", "."), font_size=11, bold=True, color=TEXT_MAIN, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 2), f"Rp {met['cof_years'][i]:,.0f}".replace(",", "."), font_size=11, color=TEXT_MUTED, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 3), f"Rp {met['ncf_years'][i]:,.0f}".replace(",", "."), font_size=11.5, bold=True, color=GREEN_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_cof.cell(i+1, 4), f"Rp {met['proceed_years'][i]:,.0f}".replace(",", "."), font_size=11.5, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)

    # Bottom 3 Cards: Analisis Kinerja Keuangan
    perf_cards = [
        ("MARGIN LABA OPERASIONAL TINGGI", GREEN_ACCENT, "Margin laba bersih operasional mencapai 61,97% di tahun pertama dan meningkat hingga 77,99% di tahun ke-5 berkat efisiensi skala ekonomi dan pengolahan adonan mingguan mandiri."),
        ("LIKUIDITAS ARUS KAS SANGAT KUAT", NAVY_LIGHT, "Arus kas masuk sangat likuid untuk menutup seluruh beban sewa ruko Rp 20 Jt/thn, gaji 3 staf karyawan Rp 120 Jt/thn, utilitas PLN/PDAM/gas, dan HPP belanja daging segar harian."),
        ("AKUMULASI KAS UNTUK EKSPANSI 2027/28", GOLD_ACCENT, "Kinerja arus kas riil yang surplus melimpah menjadi fondasi pembiayaan mandiri untuk rencana ekspansi tahun 2027/2028 dalam mencari ruko lebih luas atau pembukaan cabang kedua.")
    ]
    for idx, (p_tit, p_col, p_desc) in enumerate(perf_cards):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(5.50)
        w = Inches(3.78)
        h = Inches(1.60)

        c_pf, tf_pf = add_card(s11, x, y, w, h, bg_color=CARD_BG, border_color=p_col)
        p1 = tf_pf.paragraphs[0]
        p1.text = p_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = p_col

        p2 = tf_pf.add_paragraph()
        p2.text = f"\n{p_desc}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 12: EVALUASI KELAYAKAN INVESTASI TERPADU (4 KRITERIA INVESTASI) (BAB VI)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "Evaluasi Kelayakan Investasi Terpadu", "Pengujian 4 kriteria kelayakan modal dengan discount factor 20% sesuai standar akademik FEB UKRIDA", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=12)

    # Top Row: 4 Hero Metric Cards
    metrics_cards = [
        ("NET PRESENT VALUE (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "Target: NPV > 0", "STATUS: SANGAT LAYAK", GREEN_ACCENT, "Surplus nilai tunai bersih di atas modal awal."),
        ("INTERNAL RATE OF RETURN", "> 20,00%", "Target: IRR > 20,00%", "STATUS: SANGAT LAYAK", NAVY_LIGHT, "Imbal hasil riil jauh melampaui suku bunga bank."),
        ("PAYBACK PERIOD (PP)", f"{met['pp_months']:.2f} BULAN", "Target: PP < 36 Bulan", "STATUS: SANGAT LAYAK", GOLD_ACCENT, "Modal investasi awal Rp 62 Jt kembali kilat."),
        ("PROFITABILITY INDEX (PI)", f"{met['pi']:.2f}x", "Target: PI > 1,20x", "STATUS: SANGAT LAYAK", GREEN_ACCENT, "Efisiensi pengembalian kas sangat luar biasa.")
    ]
    for idx, (m_tit, m_val, m_tgt, m_stat, m_col, m_exp) in enumerate(metrics_cards):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(1.75)
        w = Inches(2.78)
        h = Inches(2.45)

        c_m, tf_m = add_card(s12, x, y, w, h, bg_color=CARD_BG, border_color=m_col, border_width=1.5)
        p1 = tf_m.paragraphs[0]
        p1.text = m_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        p2 = tf_m.add_paragraph()
        p2.text = m_val
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(18 if len(m_val) > 10 else 24)
        p2.font.bold = True
        p2.font.color.rgb = m_col

        p3 = tf_m.add_paragraph()
        p3.text = f"{m_tgt}\n{m_stat}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.bold = True
        p3.font.color.rgb = m_col

        p4 = tf_m.add_paragraph()
        p4.text = f"\n{m_exp}"
        p4.font.name = "Times New Roman"
        p4.font.size = Pt(9.5)
        p4.font.color.rgb = TEXT_MUTED

    # Bottom: DCF 5 Tahun Table & Sensitivity Card
    tb_shape_pv = s12.shapes.add_table(7, 3, Inches(0.8), Inches(4.35), Inches(6.8), Inches(2.75))
    tbl_pv = tb_shape_pv.table
    tbl_pv.columns[0].width = Inches(2.4)
    tbl_pv.columns[1].width = Inches(2.2)
    tbl_pv.columns[2].width = Inches(2.2)

    style_table_cell(tbl_pv.cell(0, 0), "Periode Aliran Kas", font_size=11, bold=True, color=WHITE, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_pv.cell(0, 1), "Nominal Kas Riil", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)
    style_table_cell(tbl_pv.cell(0, 2), "Present Value (DF 20%)", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, bg_color=NAVY_PRIMARY)

    pv_rows = [
        ("Tahun 0 (Initial Outlay)", f"-Rp {met['initial_outlay']:,.0f}".replace(",", "."), "Modal Awal Relokasi"),
        ("Tahun 1 (Proceed Thn 1)", f"Rp {met['proceed_years'][0]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][0]:,.0f}".replace(",", ".")),
        ("Tahun 2 (Proceed Thn 2)", f"Rp {met['proceed_years'][1]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][1]:,.0f}".replace(",", ".")),
        ("Tahun 3 (Proceed Thn 3)", f"Rp {met['proceed_years'][2]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][2]:,.0f}".replace(",", ".")),
        ("Tahun 4 (Proceed Thn 4)", f"Rp {met['proceed_years'][3]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][3]:,.0f}".replace(",", ".")),
        ("Tahun 5 (Proceed Thn 5)", f"Rp {met['proceed_years'][4]:,.0f}".replace(",", "."), f"Rp {met['pv_years'][4]:,.0f}".replace(",", "."))
    ]
    for r_idx, (p_t, p_v, p_pv) in enumerate(pv_rows, start=1):
        bg_r = CARD_BG if r_idx % 2 == 1 else CARD_BG_SOFT
        style_table_cell(tbl_pv.cell(r_idx, 0), p_t, font_size=10, bold=True, color=NAVY_PRIMARY, bg_color=bg_r)
        style_table_cell(tbl_pv.cell(r_idx, 1), p_v, font_size=10, color=TEXT_MAIN, align=PP_ALIGN.RIGHT, bg_color=bg_r)
        style_table_cell(tbl_pv.cell(r_idx, 2), p_pv, font_size=10.5, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.RIGHT, bg_color=bg_r)

    # Bottom Right: Uji Sensitivitas Card
    c_sn, tf_sn = add_card(s12, Inches(7.8), Inches(4.35), Inches(4.733), Inches(2.75), bg_color=CARD_BG, border_color=GREEN_ACCENT)
    psn1 = tf_sn.paragraphs[0]
    psn1.text = "ANALISIS SENSITIVITAS & KETAHANAN FINANSIAL"
    psn1.font.name = "Times New Roman"
    psn1.font.size = Pt(12)
    psn1.font.bold = True
    psn1.font.color.rgb = GREEN_ACCENT

    sn_bullets = [
        "Kenaikan Biaya Bahan Baku (+20%): Proyek tetap menghasilkan NPV positif di atas Rp 3 Miliar karena margin kotor yang tebal.",
        "Penurunan Volume Penjualan (-15%): Cash flow tetap likuid dan Payback Period hanya bergeser menjadi 1,2 bulan.",
        "Kenaikan Tarif Sewa Ruko (+25%): Dampak sangat kecil (<2% terhadap total outflow tahunan).",
        "Vonis Akhir Kelayakan: Bisnis Bakmi Mimu memiliki tingkat ketahanan finansial yang sangat kokoh dan risiko investasi rendah."
    ]
    for b in sn_bullets:
        pb = tf_sn.add_paragraph()
        pb.text = f"\n• {b}"
        pb.font.name = "Times New Roman"
        pb.font.size = Pt(10)
        pb.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 13: KESIMPULAN AKHIR KELAYAKAN, REKOMENDASI & PENUTUP (BAB VI & OUTRO)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)

    # Master Frame Box
    card13, tf13 = add_card(s13, Inches(0.8), Inches(0.5), Inches(11.733), Inches(6.5), bg_color=CARD_BG, border_color=BORDER_NAVY, border_width=1.5)

    tb_end = s13.shapes.add_textbox(Inches(1.0), Inches(0.65), Inches(11.333), Inches(1.15))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True
    tf_end.margin_left = tf_end.margin_right = tf_end.margin_top = tf_end.margin_bottom = 0

    pe1 = tf_end.paragraphs[0]
    pe1.alignment = PP_ALIGN.CENTER
    pe1.text = "KESIMPULAN AKHIR KELAYAKAN INVESTASI  •  BAKMI MIMU CARINA SAYANG"
    pe1.font.name = "Times New Roman"
    pe1.font.size = Pt(18)
    pe1.font.bold = True
    pe1.font.color.rgb = NAVY_PRIMARY

    pe2 = tf_end.add_paragraph()
    pe2.alignment = PP_ALIGN.CENTER
    pe2.text = "Berdasarkan evaluasi komprehensif 6 aspek bisnis, rencana investasi dinyatakan SANGAT LAYAK (FEASIBLE) untuk dijalankan dan dikembangkan."
    pe2.font.name = "Times New Roman"
    pe2.font.size = Pt(11.5)
    pe2.font.color.rgb = TEXT_MUTED

    # 6 Aspek Grid
    aspects = [
        ("ASPEK PASAR: SANGAT LAYAK", "Permintaan tinggi, basis pelanggan sejak 2019, posisi emas depan Kalam Kudus.", GREEN_ACCENT),
        ("ASPEK TEKNIS: SANGAT LAYAK", "Ruko strategis 23 kursi, open kitchen teras higienis, & adonan fresh mingguan.", GREEN_ACCENT),
        ("ASPEK MANAJEMEN: SANGAT LAYAK", "3 Staf terampil, jobdesk terstruktur, & anggaran upah pasti Rp 10 Jt/bulan.", GREEN_ACCENT),
        ("ASPEK HUKUM & AMDAL: LAYAK", "Izin lingkungan RT/RW lengkap & pengelolaan sanitasi limbah tertib higienis.", NAVY_LIGHT),
        ("ASPEK KEUANGAN: SANGAT LAYAK", "Modal Rp 62 Jt didanai ekuitas sendiri, margin laba bersih operasional 61,97%.", GREEN_ACCENT),
        ("KELAYAKAN INVESTASI: SANGAT LAYAK", "NPV Positif Rp 5,68 M, IRR > 20%, PP 0,82 Bln, & Profitability Index 175,81x.", GOLD_ACCENT)
    ]
    for idx, (a_tit, a_dsc, a_col) in enumerate(aspects):
        col = idx % 3
        row = idx // 3
        x = Inches(1.15 + col * 3.82)
        y = Inches(1.85 + row * 1.35)
        w = Inches(3.62)
        h = Inches(1.22)

        c_a, tf_a = add_card(s13, x, y, w, h, bg_color=CARD_BG_SOFT, border_color=a_col)
        p1 = tf_a.paragraphs[0]
        p1.text = a_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = a_col

        p2 = tf_a.add_paragraph()
        p2.text = a_dsc
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MAIN

    # Rekomendasi Banner
    c_rek, tf_rek = add_card(s13, Inches(1.15), Inches(4.65), Inches(11.033), Inches(0.85), bg_color=GOLD_BG_LIGHT, border_color=GOLD_ACCENT)
    pr1 = tf_rek.paragraphs[0]
    pr1.text = "REKOMENDASI STRATEGIS: REALISASI EKSPANSI TAHUN 2027/2028"
    pr1.font.name = "Times New Roman"
    pr1.font.size = Pt(11)
    pr1.font.bold = True
    pr1.font.color.rgb = GOLD_ACCENT

    pr2 = tf_rek.add_paragraph()
    pr2.text = "Memanfaatkan surplus akumulasi kas operasional untuk merealisasikan rencana ekspansi 2027/2028 dengan mencari lokasi ruko yang lebih luas di kawasan Jakarta Barat atau membuka cabang kedua guna menjaring basis pelanggan baru."
    pr2.font.name = "Times New Roman"
    pr2.font.size = Pt(10)
    pr2.font.color.rgb = TEXT_MAIN

    # Tim Penyusun & Thank You Footer
    tb_thx = s13.shapes.add_textbox(Inches(1.15), Inches(5.60), Inches(11.033), Inches(1.25))
    tf_thx = tb_thx.text_frame
    tf_thx.word_wrap = True
    tf_thx.margin_left = tf_thx.margin_right = tf_thx.margin_top = tf_thx.margin_bottom = 0

    pth1 = tf_thx.paragraphs[0]
    pth1.alignment = PP_ALIGN.CENTER
    pth1.text = "Disusun oleh Kelompok Mahasiswa FEB UKRIDA:\nArthur Reezan (312023002)  •  Jennese Putra Alamsyah Sukadi (312023033)  •  Valendrik Dwiputra Wirawan (312023013)\nAffandy (312023075)  •  Steven Putra Tjhin (312023015)"
    pth1.font.name = "Times New Roman"
    pth1.font.size = Pt(10)
    pth1.font.color.rgb = TEXT_MUTED

    pth2 = tf_thx.add_paragraph()
    pth2.alignment = PP_ALIGN.CENTER
    pth2.text = "SEKIAN & TERIMA KASIH"
    pth2.font.name = "Times New Roman"
    pth2.font.size = Pt(24)
    pth2.font.bold = True
    pth2.font.color.rgb = NAVY_PRIMARY

    # Save to both target locations
    prs.save(out_pptx_1)
    prs.save(out_pptx_2)
    print(f"[SUCCESS] Presentation generated successfully ({TOTAL_SLIDES} slides):")
    print(f"  - {out_pptx_1}")
    print(f"  - {out_pptx_2}")

if __name__ == "__main__":
    create_presentation()
