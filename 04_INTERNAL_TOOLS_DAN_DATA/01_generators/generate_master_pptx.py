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
    out_pptx = os.path.join(base, "01_TUGAS_FINAL_BAKMI_MIMU", "Presentasi_SKB_Bakmi_Mimu_Carina_Sayang.pptx")

    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Curated Executive Culinary Palette (from pptx-skills-main guidelines)
    DARK_BG = RGBColor(15, 23, 42)        # Deep Obsidian Slate (#0F172A)
    CARD_BG = RGBColor(30, 41, 59)        # Rich Surface Slate (#1E293B)
    CARD_BG_LIGHT = RGBColor(40, 53, 75)  # Hover/Accent Card Surface
    GOLD_ACCENT = RGBColor(245, 158, 11)  # Warm Gold Accent (#F59E0B)
    GOLD_LIGHT = RGBColor(251, 191, 36)   # Light Wheat Gold (#FBBF24)
    GREEN_ACCENT = RGBColor(16, 185, 129) # Emerald Feasibility (#10B981)
    GREEN_BG = RGBColor(20, 50, 40)       # Dark Emerald Container
    BLUE_ACCENT = RGBColor(56, 189, 248)  # Cyan/Sky Accent (#38BDF8)
    WHITE = RGBColor(248, 250, 252)       # Pure Crisp White (#F8FAFC)
    GRAY_TEXT = RGBColor(148, 163, 184)   # Slate Gray Secondary (#94A3B8)
    MUTED_BORDER = RGBColor(51, 65, 85)   # Border Gray (#334155)

    def add_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BG
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, subtitle_text=None, tag="STUDI KELAYAKAN BISNIS", slide_idx=1):
        # Category Badge (Top Left Pill)
        tb_tag = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(4.5), Inches(0.35))
        tf_tag = tb_tag.text_frame
        tf_tag.margin_left = tf_tag.margin_right = tf_tag.margin_top = tf_tag.margin_bottom = 0
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = f"• {tag.upper()}"
        p_tag.font.name = "Times New Roman"
        p_tag.font.size = Pt(10.5)
        p_tag.font.bold = True
        p_tag.font.color.rgb = GOLD_ACCENT

        # Slide Number Badge (Top Right)
        tb_num = slide.shapes.add_textbox(Inches(10.5), Inches(0.4), Inches(2.0), Inches(0.35))
        tf_num = tb_num.text_frame
        tf_num.margin_left = tf_num.margin_right = tf_num.margin_top = tf_num.margin_bottom = 0
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"SLIDE {slide_idx:02d} / 22"
        p_num.font.name = "Times New Roman"
        p_num.font.size = Pt(10.5)
        p_num.font.bold = True
        p_num.font.color.rgb = GRAY_TEXT

        # Title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.85))
        tf_title = tb_title.text_frame
        tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.name = "Times New Roman"
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE

        if subtitle_text:
            p_sub = tf_title.add_paragraph()
            p_sub.text = subtitle_text
            p_sub.font.name = "Times New Roman"
            p_sub.font.size = Pt(12)
            p_sub.font.color.rgb = GRAY_TEXT

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=MUTED_BORDER, border_width=1.0):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(border_width)
        else:
            card.line.fill.background()
        return card

    # =================================================================
    # SLIDE 1: COVER PAGE (Type 1: Executive Cover)
    # =================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_background(s1)

    # Big outer card container
    c_cov = add_card(s1, Inches(1.0), Inches(0.5), Inches(11.333), Inches(6.5), bg_color=CARD_BG, border_color=GOLD_ACCENT, border_width=1.5)

    # UKRIDA Logo at top center
    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(6.066), Inches(0.75), Inches(1.2), Inches(1.2))

    # Cover Title Box
    tb_c = s1.shapes.add_textbox(Inches(1.3), Inches(2.05), Inches(10.733), Inches(2.4))
    tf_c = tb_c.text_frame
    tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = 0

    p1 = tf_c.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = "FAKULTAS EKONOMI & BISNIS  •  UNIVERSITAS KRISTEN KRIDA WACANA"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT

    p2 = tf_c.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = "LAPORAN STUDI KELAYAKAN BISNIS (SKB)"
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(15)
    p2.font.color.rgb = GRAY_TEXT

    p3 = tf_c.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.text = f"“ {cfg['business_name'].upper()} ”"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(32)
    p3.font.bold = True
    p3.font.color.rgb = WHITE

    p4 = tf_c.add_paragraph()
    p4.alignment = PP_ALIGN.CENTER
    p4.text = f"- {cfg['tagline']} -"
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(13)
    p4.font.italic = True
    p4.font.color.rgb = GOLD_LIGHT

    # Authors structured 5-card horizontal grid
    authors_data = [
        ("Arthur Reezan", "312023002"),
        ("Jennese Putra Alamsyah Sukadi", "312023033"),
        ("Valendrik Dwiputra Wirawan", "312023013"),
        ("Affandy", "312023075"),
        ("Steven Putra Tjhin", "312023015")
    ]
    card_w = Inches(2.05)
    card_gap = Inches(0.12)
    start_x = Inches(1.35)
    card_y = Inches(4.7)
    card_h = Inches(1.1)

    for i, (name, nim) in enumerate(authors_data):
        x = start_x + i * (card_w + card_gap)
        add_card(s1, x, card_y, card_w, card_h, bg_color=DARK_BG, border_color=MUTED_BORDER)
        tb_m = s1.shapes.add_textbox(x + Inches(0.08), card_y + Inches(0.12), card_w - Inches(0.16), card_h - Inches(0.24))
        tf_m = tb_m.text_frame
        tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = 0
        
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
        pm2.font.color.rgb = WHITE

        pm3 = tf_m.add_paragraph()
        pm3.alignment = PP_ALIGN.CENTER
        pm3.text = f"NIM: {nim}"
        pm3.font.name = "Times New Roman"
        pm3.font.size = Pt(9.5)
        pm3.font.color.rgb = GRAY_TEXT

    # Footer note
    tb_foot = s1.shapes.add_textbox(Inches(1.3), Inches(6.0), Inches(10.733), Inches(0.5))
    tf_foot = tb_foot.text_frame
    p_f = tf_foot.paragraphs[0]
    p_f.alignment = PP_ALIGN.CENTER
    p_f.text = "Program Studi Sarjana Manajemen  |  Konsentrasi Studi Kelayakan Bisnis  |  Tahun Akademik 2024"
    p_f.font.name = "Times New Roman"
    p_f.font.size = Pt(10)
    p_f.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 2: TABLE OF CONTENTS / AGENDA (Type 2: 6-Chapter Structured Grid)
    # =================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_background(s2)
    add_header(s2, "Daftar Isi & Sistematika Kajian", "Ringkasan 6 bab utama studi kelayakan bisnis sesuai standar FEB UKRIDA", tag="AGENDA PRESENTASI", slide_idx=2)

    toc_items = [
        ("01", "BAB I: PENDAHULUAN & PROFIL", "Perjalanan usaha 1999 Taman Aries -> 2016 Duri Kosambi, basis pelanggan 2019, & peluang ekspansi 2027/2028."),
        ("02", "BAB II: ASPEK PASAR & PEMASARAN", "Segmentasi STP, 3 minyak mie, topping ayam putih/babi kecap/casiu, benchmarking, dan bauran 7P."),
        ("03", "BAB III: ASPEK MANAJEMEN", "Struktur organisasi 3 staf karyawan fungsional, job description terstruktur, dan total alokasi gaji Rp 10 Jt/bln."),
        ("04", "BAB IV: ASPEK TEKNIS & OPERASI", "Tata letak ruko 1 lantai (22 kursi), modal awal 2016 Rp 62 Jt, open kitchen gerobak, & alur produksi mie mingguan."),
        ("05", "BAB V: ASPEK KEUANGAN TERPADU", f"Initial outlay Rp {met['initial_outlay']/1e6:.1f} jt, proyeksi 5 tahun arus kas (CIF/COF), dan depresiasi garis lurus 20%."),
        ("06", "BAB VI: KELAYAKAN INVESTASI", f"Hasil pengujian 4 kriteria investasi modal: NPV Rp {met['npv']/1e9:.2f} M, IRR {met['irr']*100:.2f}%, PP {met['pp_months']:.2f} Bulan, dan PI {met['pi']:.2f}.")
    ]

    for idx, (num, title, desc) in enumerate(toc_items):
        col = idx % 3
        row = idx // 3
        x = Inches(0.8 + col * 3.98)
        y = Inches(1.85 + row * 2.55)
        w = Inches(3.78)
        h = Inches(2.35)

        add_card(s2, x, y, w, h, bg_color=CARD_BG, border_color=MUTED_BORDER)
        
        # Inner number badge
        add_card(s2, x + Inches(0.25), y + Inches(0.25), Inches(0.8), Inches(0.45), bg_color=DARK_BG, border_color=GOLD_ACCENT)
        tb_n = s2.shapes.add_textbox(x + Inches(0.25), y + Inches(0.25), Inches(0.8), Inches(0.45))
        tb_n.text_frame.margin_left = tb_n.text_frame.margin_right = tb_n.text_frame.margin_top = tb_n.text_frame.margin_bottom = 0
        pn = tb_n.text_frame.paragraphs[0]
        pn.alignment = PP_ALIGN.CENTER
        pn.text = num
        pn.font.name = "Times New Roman"
        pn.font.size = Pt(16)
        pn.font.bold = True
        pn.font.color.rgb = GOLD_ACCENT

        # Title & Desc Box
        tb_t = s2.shapes.add_textbox(x + Inches(0.25), y + Inches(0.8), w - Inches(0.5), h - Inches(0.9))
        tf_t = tb_t.text_frame
        tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0

        pt = tf_t.paragraphs[0]
        pt.text = title
        pt.font.name = "Times New Roman"
        pt.font.size = Pt(13.5)
        pt.font.bold = True
        pt.font.color.rgb = WHITE

        pd = tf_t.add_paragraph()
        pd.text = desc
        pd.font.name = "Times New Roman"
        pd.font.size = Pt(11)
        pd.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 3: LATAR BELAKANG & EVOLUSI (Type 4: Content - Timeline/Story)
    # =================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_background(s3)
    add_header(s3, "Latar Belakang & Filosofi Usaha", "Perjalanan kuliner legendaris sejak 1999 dan peluang ekspansi masa depan", tag="BAB I: PENDAHULUAN", slide_idx=3)

    story_steps = [
        ("1999", "Awal di Taman Aries", "Usaha kuliner bakmi keluarga pertama kali didirikan pada tahun 1999 di Taman Aries (Meruya) dengan resep autentik turun-temurun."),
        ("2016", "Relokasi ke Duri Kosambi", "Tahun 2016 kedai berpindah tempat ke ruko di Jl. Angsoka Hijau IV Duri Kosambi, Cengkareng untuk melayani pemukiman sekitar."),
        ("2019", "Basis Pelanggan Solid", "Sekitar tahun 2019 basis pelanggan setia terbentuk mantap, menjadi destinasi sarapan & makan siang favorit keluarga dan civitas Kalam Kudus."),
        ("2027/28", "Peluang Ekspansi Cabang", "Tingginya permintaan membuka peluang ekspansi sekitar tahun 2027/2028 untuk mencari tempat baru yang lebih luas atau membuka cabang baru.")
    ]

    for idx, (yr, ttl, dsc) in enumerate(story_steps):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(1.9)
        w = Inches(2.78)
        h = Inches(4.0)

        add_card(s3, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 3 else MUTED_BORDER)
        
        # Step header box
        tb_step = s3.shapes.add_textbox(x + Inches(0.25), y + Inches(0.3), w - Inches(0.5), Inches(0.8))
        tf_step = tb_step.text_frame
        tf_step.margin_left = tf_step.margin_right = tf_step.margin_top = tf_step.margin_bottom = 0
        
        p_yr = tf_step.paragraphs[0]
        p_yr.text = yr
        p_yr.font.name = "Times New Roman"
        p_yr.font.size = Pt(22)
        p_yr.font.bold = True
        p_yr.font.color.rgb = GOLD_ACCENT

        p_ttl = tf_step.add_paragraph()
        p_ttl.text = ttl
        p_ttl.font.name = "Times New Roman"
        p_ttl.font.size = Pt(13)
        p_ttl.font.bold = True
        p_ttl.font.color.rgb = WHITE

        # Content text
        tb_dsc = s3.shapes.add_textbox(x + Inches(0.25), y + Inches(1.3), w - Inches(0.5), h - Inches(1.5))
        tf_dsc = tb_dsc.text_frame
        tf_dsc.margin_left = tf_dsc.margin_right = tf_dsc.margin_top = tf_dsc.margin_bottom = 0
        p_dsc = tf_dsc.paragraphs[0]
        p_dsc.text = dsc
        p_dsc.font.name = "Times New Roman"
        p_dsc.font.size = Pt(11)
        p_dsc.font.color.rgb = GRAY_TEXT

    # Bottom Callout: Filosofi Adonan Mie Mingguan
    add_card(s3, Inches(0.8), Inches(6.05), Inches(11.7), Inches(0.9), bg_color=CARD_BG, border_color=GOLD_LIGHT)
    tb_ph = s3.shapes.add_textbox(Inches(1.1), Inches(6.15), Inches(11.1), Inches(0.7))
    tf_ph = tb_ph.text_frame
    tf_ph.margin_left = tf_ph.margin_right = tf_ph.margin_top = tf_ph.margin_bottom = 0
    pph1 = tf_ph.paragraphs[0]
    pph1.text = "FILOSOFI MUTU PRODUK: ADONAN MIE FRESH SETIAP MINGGU (BATCH MINGGUAN TANPA PENGAWET)"
    pph1.font.name = "Times New Roman"
    pph1.font.size = Pt(11.5)
    pph1.font.bold = True
    pph1.font.color.rgb = GOLD_ACCENT
    pph2 = tf_ph.add_paragraph()
    pph2.text = "Adonan mie selalu dibuat mandiri setiap minggu menghasilkan tekstur kenyal alami berkilau (shining), bebas dari pengawet kimia berbahaya demi menjaga kesehatan dan keaslian cita rasa."
    pph2.font.name = "Times New Roman"
    pph2.font.size = Pt(10.5)
    pph2.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 4: 5 PILAR KEUNGGULAN KOMPETITIF (Type 4: 5-Cards Grid)
    # =================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_background(s4)
    add_header(s4, "5 Pilar Keunggulan Kompetitif", "Faktor diferensiasi penentu daya saing Bakmi Mimu Carina Sayang", tag="BAB I: PENDAHULUAN", slide_idx=4)

    pillars = [
        ("01", "Taste (Cita Rasa)", "Mie kenyal buatan sendiri dipadu 3 pilihan minyak mie (babi, ayam, sayur+wijen), kaldu asli, & 3 topping murni: ayam putih, babi kecap, dan casiu madu (tanpa bebek/jamur)."),
        ("02", "Location (Lokasi)", "Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi, Cengkareng (persis di depan Taman TK, SD, SMP Kalam Kudus), akses sangat strategis & parkir motor aman."),
        ("03", "Service (Open Kitchen)", "Dapur terbuka (open kitchen) gerobak di teras depan ruko, perebusan & peracikan higienis terlihat jelas, serta dilengkapi saus botolan terpercaya (Cap Belibis & Mangga Besar)."),
        ("04", "Size (Variasi Porsi)", "Pilihan fleksibel: Porsi Kecil (27k), Reguler (29k-39k), hingga Porsi Jumbo (+100% mie dari porsi standar / 2x lipat porsi biasa 49k-59k) untuk segala tingkatan selera."),
        ("05", "Pelengkap & Minuman", "Pilihan pelengkap swikiaw rebus, pangsit rebus, baso sapi & ikan kuah (@4.5k-5.5k/pcs), baso goreng babi-udang 8k, badak sarsaparilla, susu kacang & liang teh 12k.")
    ]

    for idx, (num, tit, dsc) in enumerate(pillars):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(1.9)
        w = Inches(2.2)
        h = Inches(4.9)

        add_card(s4, x, y, w, h, bg_color=CARD_BG, border_color=MUTED_BORDER)
        
        # Pill indicator
        add_card(s4, x + Inches(0.2), y + Inches(0.25), Inches(0.6), Inches(0.35), bg_color=DARK_BG, border_color=GOLD_ACCENT)
        tb_p = s4.shapes.add_textbox(x + Inches(0.2), y + Inches(0.25), Inches(0.6), Inches(0.35))
        tb_p.text_frame.margin_left = tb_p.text_frame.margin_right = tb_p.text_frame.margin_top = tb_p.text_frame.margin_bottom = 0
        pp = tb_p.text_frame.paragraphs[0]
        pp.alignment = PP_ALIGN.CENTER
        pp.text = num
        pp.font.name = "Times New Roman"
        pp.font.size = Pt(13)
        pp.font.bold = True
        pp.font.color.rgb = GOLD_ACCENT

        tb_ct = s4.shapes.add_textbox(x + Inches(0.2), y + Inches(0.75), w - Inches(0.4), h - Inches(0.9))
        tf_ct = tb_ct.text_frame
        tf_ct.margin_left = tf_ct.margin_right = tf_ct.margin_top = tf_ct.margin_bottom = 0
        
        pct1 = tf_ct.paragraphs[0]
        pct1.text = tit
        pct1.font.name = "Times New Roman"
        pct1.font.size = Pt(13)
        pct1.font.bold = True
        pct1.font.color.rgb = WHITE

        pct2 = tf_ct.add_paragraph()
        pct2.text = f"\n{dsc}"
        pct2.font.name = "Times New Roman"
        pct2.font.size = Pt(10.5)
        pct2.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 5: DAFTAR MENU & PENETAPAN HARGA (Type 4: Data/Table Showcase)
    # =================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_background(s5)
    add_header(s5, "Daftar Menu & Struktur Harga Jual Resmi", "Struktur harga bersaing (Value-Based Pricing) sesuai daftar menu riil kedai", tag="BAB I: PENDAHULUAN", slide_idx=5)

    # Left Card: Menu Utama
    add_card(s5, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    tb_m1 = s5.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_m1 = tb_m1.text_frame
    tf_m1.margin_left = tf_m1.margin_right = tf_m1.margin_top = tf_m1.margin_bottom = 0

    p_m1_t = tf_m1.paragraphs[0]
    p_m1_t.text = "HIDANGAN BAKMI & NASI TIM (TOPPING MURNI)"
    p_m1_t.font.name = "Times New Roman"
    p_m1_t.font.size = Pt(13)
    p_m1_t.font.bold = True
    p_m1_t.font.color.rgb = GOLD_ACCENT

    p_note = tf_m1.add_paragraph()
    p_note.text = "*Topping: Ayam Putih, Babi Kecap, Casiu (Gak Ada Bebek & Jamur)"
    p_note.font.name = "Times New Roman"
    p_note.font.size = Pt(9.5)
    p_note.font.italic = True
    p_note.font.color.rgb = GOLD_LIGHT

    m1_items = [
        ("Mie Spesial 3 Isi (Ayam Putih, Babi Kecap, Casiu)", "Rp 39.000", "Porsi reguler dengan 3 topping lengkap"),
        ("Mie Daging Casiu (Spesial Chasiu Madu)", "Rp 31.000", "Topping irisan babi casiu madu panggang gurih"),
        ("Mie Campur (Ayam Putih + Babi Kecap)", "Rp 29.000", "Kombinasi gurih ayam putih & babi kecap"),
        ("Mie Ayam Putih saja / Babi Kecap saja", "Rp 29.000", "Pilihan topping tunggal gurih autentik"),
        ("Porsi Jumbo (+100% Mie dari Standar)", "Rp 49.000 - 59.000", "Campur 49K | Casiu 54K | Spesial 3 Isi 59K"),
        ("Porsi Kecil (Sarapan Hemat / Anak-anak)", "Rp 27.000 - 29.000", "Campur/Ayam/Babi 27K | Casiu 29K")
    ]
    for n, h, d in m1_items:
        p_row = tf_m1.add_paragraph()
        p_row.text = f"• {n}  ➔  {h}\n  {d}"
        p_row.font.name = "Times New Roman"
        p_row.font.size = Pt(10)
        p_row.font.color.rgb = WHITE

    # Right Card: Pelengkap & Minuman
    add_card(s5, Inches(6.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=BLUE_ACCENT)
    tb_m2 = s5.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_m2 = tb_m2.text_frame
    tf_m2.margin_left = tf_m2.margin_right = tf_m2.margin_top = tf_m2.margin_bottom = 0

    p_m2_t = tf_m2.paragraphs[0]
    p_m2_t.text = "MENU PELENGKAP, KUAH & MINUMAN"
    p_m2_t.font.name = "Times New Roman"
    p_m2_t.font.size = Pt(13)
    p_m2_t.font.bold = True
    p_m2_t.font.color.rgb = BLUE_ACCENT

    p_note2 = tf_m2.add_paragraph()
    p_note2.text = "*Saus Meja Standar Botol: Cap Belibis & Mangga Besar (Tanpa Jeruk)"
    p_note2.font.name = "Times New Roman"
    p_note2.font.size = Pt(9.5)
    p_note2.font.italic = True
    p_note2.font.color.rgb = BLUE_ACCENT

    m2_items = [
        ("Pangsit Rebus Kuah (5 pcs)", "Rp 22.500", "@Rp 4.500/pcs semangkuk kuah kaldu gurih"),
        ("Swikiaw Rebus Kuah (5 pcs)", "Rp 27.500", "@Rp 5.500/pcs padat isian udang & babi"),
        ("Baso Sapi & Baso Ikan Kuah (5 pcs)", "Rp 22.500", "Sama persis harga pangsit rebus (@Rp 4.500/pcs)"),
        ("Pangsit Goreng & Baso Goreng", "Rp 5.000 - 8.000", "Pangsit grg 5K/pcs | Baso grg babi-udang 8K/pcs"),
        ("Badak Sarsaparilla", "Rp 12.000 - 14.000", "Soda legendaris Siantar (Botol 12K / +Es Batu 14K)"),
        ("Aneka Minuman Tradisional (1K - 15K)", "Rp 1.000 - 12.000", "Susu Kacang 12K | Liang Teh 12K | Botol 6K | Teh 1-5K")
    ]
    for n, h, d in m2_items:
        p_row2 = tf_m2.add_paragraph()
        p_row2.text = f"• {n}  ➔  {h}\n  {d}"
        p_row2.font.name = "Times New Roman"
        p_row2.font.size = Pt(10)
        p_row2.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 6: ANALISIS STP PASAR (Type 4: 3-Column Deep Dive)
    # =================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_background(s6)
    add_header(s6, "Analisis Pasar: STP Framework", "Segmenting, Targeting, dan Positioning pasar sasaran Bakmi Mimu", tag="BAB II: ASPEK PASAR", slide_idx=6)

    stp_cols = [
        ("SEGMENTASI (SEGMENTING)", [
            ("Geografis", "Radius 0 - 5 km meliputi Perumahan Duri Kosambi, Semanan, Taman Semanan Indah, Cengkareng, dan Puri Kembangan."),
            ("Demografis", "Pria & wanita usia 7 - 65 tahun, siswa-siswi, guru, orang tua murid Kalam Kudus, karyawan kantor, SES Menengah & Menengah ke Atas."),
            ("Perilaku (Behavioral)", "Pencinta kuliner mie oriental yang mengutamakan tekstur kenyal, cita rasa bumbu otentik, serta higienitas ruang makan.")
        ]),
        ("TARGET PASAR (TARGETING)", [
            ("Prioritas Utama", "Keluarga residensial perumahan sekitar Duri Kosambi yang mencari santapan pagi dan makan siang lezat."),
            ("Civitas Akademika", "Siswa-siswi, guru, staf, dan orang tua murid Sekolah Kristen Kalam Kudus Kosambi tepat di depan gerai."),
            ("Komunitas & Online", "Jemaat gereja hari Sabtu-Minggu serta pelanggan pesan antar digital (GoFood/GrabFood/ShopeeFood).")
        ]),
        ("POSISI PASAR (POSITIONING)", [
            ("Core Identity", "“Kedai Bakmi Otentik Legendaris dengan Kenyamanan Ruko Modern & Standar Kebersihan Tinggi.”"),
            ("Value Proposition", "Memberikan kepuasan cita rasa resep 1999 dengan porsi mengenyangkan dan harga sangat sebanding (value-for-money)."),
            ("Brand Perception", "Bukan sekadar warung kaki lima, melainkan restoran bakmi keluarga yang bersih, ramah, dan terpercaya.")
        ])
    ]

    for idx, (title, points) in enumerate(stp_cols):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(1.9)
        w = Inches(3.78)
        h = Inches(4.9)

        add_card(s6, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 2 else MUTED_BORDER)
        
        tb_stp = s6.shapes.add_textbox(x + Inches(0.25), y + Inches(0.25), w - Inches(0.5), h - Inches(0.5))
        tf_stp = tb_stp.text_frame
        tf_stp.margin_left = tf_stp.margin_right = tf_stp.margin_top = tf_stp.margin_bottom = 0
        
        pt1 = tf_stp.paragraphs[0]
        pt1.text = title
        pt1.font.name = "Times New Roman"
        pt1.font.size = Pt(13)
        pt1.font.bold = True
        pt1.font.color.rgb = GOLD_ACCENT

        for sub, txt in points:
            p_sub = tf_stp.add_paragraph()
            p_sub.text = f"\n• {sub}:"
            p_sub.font.name = "Times New Roman"
            p_sub.font.size = Pt(11)
            p_sub.font.bold = True
            p_sub.font.color.rgb = WHITE

            p_txt = tf_stp.add_paragraph()
            p_txt.text = txt
            p_txt.font.name = "Times New Roman"
            p_txt.font.size = Pt(10.5)
            p_txt.font.color.rgb = GRAY_TEXT

    # =================================================================
    # =================================================================
    # SLIDE 7: PETA PERSAINGAN KOMPETITOR & KARAKTERISTIK KEDAI (Type 4: Comparison Matrix)
    # =================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_background(s7)
    add_header(s7, "Analisis Persaingan & Peta Kompetitor", "Matriks perbandingan daya saing Bakmi Mimu terhadap pelaku usaha sejenis", tag="BAB II: ASPEK PASAR", slide_idx=7)

    comp_data = [
        ("BAKMI MIMU CARINA SAYANG", "Brand Kita (Unggulan)", GOLD_ACCENT, [
            ("Rasa & Resep", "Resep keluarga sejak 1999 di Taman Aries, 3 racikan minyak khas (babi, ayam, sayur+wijen), adonan fresh mingguan."),
            ("Fasilitas Kedai", "Ruko 1 lantai, sirkulasi 6 unit kipas angin, suasana hangat padat (maks 22 orang), open kitchen gerobak (tanpa AC/Wi-Fi)."),
            ("Varian Topping", "3 Topping murni: Ayam putih gurih, babi kecap manis-gurih, casiu madu panggang (tanpa bebek/jamur), plus pangsit/swikiaw."),
            ("Kemasan & Layanan", "Penyajian cepat < 5 menit, pemesanan nota manual, saus botol berkualitas (Cap Belibis & Mangga Besar)."),
            ("Harga & Value", "Rp 24.000 - Rp 59.000 (Sangat terjangkau dengan varian porsi kecil, reguler, hingga jumbo +100% mie).")
        ]),
        ("BAKMI ALOK / BRAND BESAR", "Kompetitor Merek Terkenal", GRAY_TEXT, [
            ("Rasa & Resep", "Cita rasa bakmi ayam rebus gurih, brand equity kuat di Jakarta Barat."),
            ("Fasilitas Kedai", "Restoran permanen, kapasitas besar ber-AC, antrean panjang saat jam makan siang."),
            ("Varian Topping", "Fokus dominan ayam kampung rebus, tidak menyediakan olahan babi casiu madu."),
            ("Kemasan & Layanan", "Standar restoran waralaba terstruktur, kemasan branded formal."),
            ("Harga & Value", "Rp 45.000 - Rp 65.000 (Segmen premium, relatif lebih mahal untuk porsi harian).")
        ]),
        ("WARUNG BAKMI LOKAL KOSAMBI", "Kompetitor Tradisional Sekitar", GRAY_TEXT, [
            ("Rasa & Resep", "Cita rasa standar warung tenda gerobak kaki lima, bumbu penyedap dominan."),
            ("Fasilitas Kedai", "Kios semi terbuka/tenda pinggir jalan, sirkulasi udara terbatas, parkir sempit."),
            ("Varian Topping", "Topping terbatas ayam cincang biasa, jarang menyediakan menu swikiaw rebus."),
            ("Kemasan & Layanan", "Kemasan plastik mika konvensional, rentan tumpah dan cepat dingin."),
            ("Harga & Value", "Rp 20.000 - Rp 25.000 (Harga murah namun kebersihan & fasilitas sangat minim).")
        ])
    ]

    for idx, (b_name, b_sub, b_col, pts) in enumerate(comp_data):
        x = Inches(0.8 + idx * 3.98)
        y = Inches(1.9)
        w = Inches(3.78)
        h = Inches(4.9)

        add_card(s7, x, y, w, h, bg_color=CARD_BG, border_color=b_col, border_width=1.5 if idx == 0 else 1.0)
        
        tb_cp = s7.shapes.add_textbox(x + Inches(0.25), y + Inches(0.2), w - Inches(0.5), h - Inches(0.4))
        tf_cp = tb_cp.text_frame
        tf_cp.margin_left = tf_cp.margin_right = tf_cp.margin_top = tf_cp.margin_bottom = 0
        
        pt1 = tf_cp.paragraphs[0]
        pt1.text = b_name
        pt1.font.name = "Times New Roman"
        pt1.font.size = Pt(13)
        pt1.font.bold = True
        pt1.font.color.rgb = b_col

        pt2 = tf_cp.add_paragraph()
        pt2.text = b_sub
        pt2.font.name = "Times New Roman"
        pt2.font.size = Pt(10.5)
        pt2.font.color.rgb = GRAY_TEXT

        for p_name, p_desc in pts:
            p_row_t = tf_cp.add_paragraph()
            p_row_t.text = f"\n▸ {p_name}:"
            p_row_t.font.name = "Times New Roman"
            p_row_t.font.size = Pt(10.5)
            p_row_t.font.bold = True
            p_row_t.font.color.rgb = WHITE

            p_row_d = tf_cp.add_paragraph()
            p_row_d.text = p_desc
            p_row_d.font.name = "Times New Roman"
            p_row_d.font.size = Pt(10)
            p_row_d.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 8: BAURAN PEMASARAN 7P - BAGIAN 1 (Type 4: 2x2 Grid 4P)
    # =================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_background(s8)
    add_header(s8, "Strategi Bauran Pemasaran (4P Tradisional)", "Strategi eksekusi taktis: Product, Price, Place, dan Promotion", tag="BAB II: ASPEK PASAR", slide_idx=8)

    p4_cards = [
        ("PRODUCT (PRODUK)", GOLD_ACCENT, [
            "Resep autentik 1999 (Taman Aries), adonan mie fresh dibuat mandiri setiap minggu tanpa pengawet.",
            "3 Topping murni: Ayam Putih, Babi Kecap, & Casiu Madu gurih (tanpa daging bebek & jamur).",
            "3 Pilihan racikan minyak mie: Minyak Babi, Minyak Ayam, & Minyak Sayur + Wijen.",
            "Variasi porsi fleksibel: Porsi Kecil (27k), Reguler (29k-39k), hingga Porsi Jumbo (+100% mie, 49k-59k)."
        ]),
        ("PRICE (HARGA)", GREEN_ACCENT, [
            "Menerapkan Value-Based Pricing bersaing di kawasan residensial Duri Kosambi.",
            "Range harga bakmi: Rp 24.000 s.d. Rp 59.000 (belum termasuk menu pelengkap tambahan).",
            "Menu kuah: Pangsit rebus Rp 22.500 (5 pcs), Swikiaw Rp 27.500 (5 pcs), Baso sapi/ikan Rp 22.500.",
            "Gorengan Rp 5.000 - 8.000, minuman segar Rp 1.000 - 15.000 (Badak, Liang Teh, Susu Kacang, tanpa jeruk)."
        ]),
        ("PLACE (DISTRIBUSI / LOKASI)", BLUE_ACCENT, [
            "Lokasi ruko 1 lantai di Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi, Cengkareng, Jakarta Barat.",
            "Posisi emas persis di depan Taman TK, SD, dan SMP Sekolah Kristen Kalam Kudus.",
            "Fasilitas ruang makan berdaya tampung maksimal 22 orang padat (4 meja pendek + meja bar dinding).",
            "Fokus saluran penjualan langsung: Makan di tempat (Dine-In) dan pesanan bawa pulang (Take-Away)."
        ]),
        ("PROMOTION (PROMOSI)", GOLD_LIGHT, [
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
        y = Inches(1.9 + r * 2.5)
        w = Inches(5.7)
        h = Inches(2.3)

        add_card(s8, x, y, w, h, bg_color=CARD_BG, border_color=col)
        tb = s8.shapes.add_textbox(x + Inches(0.25), y + Inches(0.2), w - Inches(0.5), h - Inches(0.4))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = col

        for b in bullets:
            p_b = tf.add_paragraph()
            p_b.text = f"• {b}"
            p_b.font.name = "Times New Roman"
            p_b.font.size = Pt(10.5)
            p_b.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 9: BAURAN PEMASARAN 7P - BAGIAN 2 (Type 4: 3 Horizontal Cards 3P)
    # =================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_background(s9)
    add_header(s9, "Strategi Bauran Layanan (3P Extended)", "Pilar keunggulan operasional: People, Process, dan Physical Evidence", tag="BAB II: ASPEK PASAR", slide_idx=9)

    p3_cards = [
        ("PEOPLE (SUMBER DAYA MANUSIA)", GOLD_ACCENT, [
            "Tim operasional 3 karyawan fungsional: Koki Utama, Asisten Koki, dan Kasir-Pramusaji.",
            "Sistem remunerasi berbasis gaji pokok bulanan tetap (total alokasi gaji Rp 10 Jt/bulan).",
            "Tidak menerapkan bonus insentif penjualan, menjaga fokus tim pada konsistensi racikan rasa.",
            "Karyawan berseragam bersih, ramah menyambut tamu, dan menguasai varian menu topping."
        ]),
        ("PROCESS (PROSES OPERASIONAL)", BLUE_ACCENT, [
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
        y = Inches(1.9)
        w = Inches(3.78)
        h = Inches(4.9)

        add_card(s9, x, y, w, h, bg_color=CARD_BG, border_color=col)
        tb = s9.shapes.add_textbox(x + Inches(0.25), y + Inches(0.25), w - Inches(0.5), h - Inches(0.5))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = col

        for b in bullets:
            pb = tf.add_paragraph()
            pb.text = f"\n▸ {b}"
            pb.font.name = "Times New Roman"
            pb.font.size = Pt(10.5)
            pb.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 10: STRUKTUR ORGANISASI & PEMBAGIAN TUGAS (Type 4: Hierarchy Flow)
    # =================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_background(s10)
    add_header(s10, "Struktur Organisasi & Pembagian Tugas", "Desain tata kelola tim operasional gerai Bakmi Mimu Carina Sayang", tag="BAB III: ASPEK MANAJEMEN", slide_idx=10)

    # Top Manager Card
    add_card(s10, Inches(3.8), Inches(1.9), Inches(5.7), Inches(1.3), bg_color=CARD_BG, border_color=GOLD_ACCENT, border_width=1.5)
    tb_mng = s10.shapes.add_textbox(Inches(4.0), Inches(2.0), Inches(5.3), Inches(1.1))
    tf_mng = tb_mng.text_frame
    tf_mng.margin_left = tf_mng.margin_right = tf_mng.margin_top = tf_mng.margin_bottom = 0
    p_m1 = tf_mng.paragraphs[0]
    p_m1.text = "OWNER / PENGELOLA UTAMA KEDAI (1 ORANG)"
    p_m1.font.name = "Times New Roman"
    p_m1.font.size = Pt(13)
    p_m1.font.bold = True
    p_m1.font.color.rgb = GOLD_ACCENT
    p_m2 = tf_mng.add_paragraph()
    p_m2.text = "Tanggung Jawab: Penetapan strategi usaha, pengawasan keuangan harian, kontrol cita rasa SOP resep 1999, pengadaan bahan baku daging segar, dan perencanaan ekspansi 2027/2028."
    p_m2.font.name = "Times New Roman"
    p_m2.font.size = Pt(10.5)
    p_m2.font.color.rgb = WHITE

    # 3 Subordinate Cards (3 Karyawan Tetap)
    sub_roles = [
        ("KOKI UTAMA (1 ORANG)", BLUE_ACCENT, [
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
        ("KASIR & PRAMUSAJI (1 ORANG)", GOLD_LIGHT, [
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
        h = Inches(3.35)

        add_card(s10, x, y, w, h, bg_color=CARD_BG, border_color=r_col)
        tb_sub = s10.shapes.add_textbox(x + Inches(0.2), y + Inches(0.2), w - Inches(0.4), h - Inches(0.4))
        tf_sub = tb_sub.text_frame
        tf_sub.margin_left = tf_sub.margin_right = tf_sub.margin_top = tf_sub.margin_bottom = 0

        p1 = tf_sub.paragraphs[0]
        p1.text = r_title
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12.5)
        p1.font.bold = True
        p1.font.color.rgb = r_col

        for t in r_tasks:
            pt = tf_sub.add_paragraph()
            pt.text = f"• {t}"
            pt.font.name = "Times New Roman"
            pt.font.size = Pt(10.5)
            pt.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 11: TENAGA KERJA & SKEMA KOMPENSASI (Type 4: Stat + Breakdown)
    # =================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_background(s11)
    add_header(s11, "Tenaga Kerja & Kebijakan Kompensasi", "Kebutuhan SDM, alokasi upah bulanan, dan kebijakan ketenagakerjaan", tag="BAB III: ASPEK MANAJEMEN", slide_idx=11)

    # Top Stat Callouts
    add_card(s11, Inches(0.8), Inches(1.9), Inches(5.7), Inches(1.3), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    tb_s1 = s11.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(1.1))
    tf_s1 = tb_s1.text_frame
    tf_s1.margin_left = tf_s1.margin_right = tf_s1.margin_top = tf_s1.margin_bottom = 0
    p1 = tf_s1.paragraphs[0]
    p1.text = "TOTAL TENAGA KERJA: 3 STAF KARYAWAN"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT
    p2 = tf_s1.add_paragraph()
    p2.text = "Formasi Fungsional: 1 Koki Utama Dapur, 1 Asisten Koki Dapur, 1 Kasir & Pramusaji."
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(11)
    p2.font.color.rgb = WHITE

    add_card(s11, Inches(6.8), Inches(1.9), Inches(5.7), Inches(1.3), bg_color=CARD_BG, border_color=GREEN_ACCENT)
    tb_s2 = s11.shapes.add_textbox(Inches(7.0), Inches(2.0), Inches(5.3), Inches(1.1))
    tf_s2 = tb_s2.text_frame
    tf_s2.margin_left = tf_s2.margin_right = tf_s2.margin_top = tf_s2.margin_bottom = 0
    p3 = tf_s2.paragraphs[0]
    p3.text = "TOTAL ANGGARAN GAJI: RP 120.000.000 / TAHUN"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(15)
    p3.font.bold = True
    p3.font.color.rgb = GREEN_ACCENT
    p4 = tf_s2.add_paragraph()
    p4.text = "Alokasi anggaran gaji tetap stabil Rp 10.000.000 per bulan (tanpa skema bonus insentif)."
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(11)
    p4.font.color.rgb = WHITE

    # 4 Compensation Cards Grid
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
        y = Inches(3.45 + r * 1.7)
        w = Inches(5.7)
        h = Inches(1.5)

        add_card(s11, x, y, w, h, bg_color=CARD_BG, border_color=MUTED_BORDER)
        tb = s11.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), w - Inches(0.4), h - Inches(0.3))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = f"• {c_tit}"
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf.add_paragraph()
        p2.text = c_txt
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 12: TATA LETAK RUKO & REKAPITULASI MODAL 2016 (Type 4: Layout & Capital)
    # =================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_background(s12)
    add_header(s12, "Tata Letak Ruko 1 Lantai & Modal Awal 2016", "Denah penataan kedai 22 kursi padat dan rincian realisasi modal awal relokasi 2016", tag="BAB IV: TEKNIS & OPERASI", slide_idx=12)

    # Left Container: Tata Letak Ruko 1 Lantai
    add_card(s12, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    tb_f1 = s12.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_f1 = tb_f1.text_frame
    tf_f1.margin_left = tf_f1.margin_right = tf_f1.margin_top = tf_f1.margin_bottom = 0

    p_f1 = tf_f1.paragraphs[0]
    p_f1.text = "TATA LETAK KEDAI (RUKO 1 LANTAI - 22 KURSI PADAT)"
    p_f1.font.name = "Times New Roman"
    p_f1.font.size = Pt(12.5)
    p_f1.font.bold = True
    p_f1.font.color.rgb = GOLD_ACCENT

    f1_details = [
        ("Area Depan (Teras Ruko)", "Gerobak etalase open kitchen (stasiun rebus mie, dandang kaldu panas, kompor gas, meja bumbu & 3 jenis minyak mie)."),
        ("Area Makan (Tengah)", "Kapasitas 22 kursi padat: 4 meja pendek kayu (16 kursi) + 1 meja panjang menempel dinding (6 kursi), sirkulasi 6 kipas angin."),
        ("Area Kasir & Minuman", "Meja kasir nota manual, etalase minuman botol dingin (Badak, Liang Teh, Susu Kacang, Aqua, Teh Botol/Pucuk)."),
        ("Area Belakang (Dapur Cuci)", "Dapur persiapan & stok, 1 unit kulkas, 1 unit freezer daging beku, stasiun cuci piring mangkok, dan 1 toilet ruko.")
    ]
    for n, d in f1_details:
        p = tf_f1.add_paragraph()
        p.text = f"\n▸ {n}:\n  {d}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(10.5)
        p.font.color.rgb = WHITE

    # Right Container: Realisasi Modal Awal 2016
    add_card(s12, Inches(6.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=BLUE_ACCENT)
    tb_f2 = s12.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_f2 = tb_f2.text_frame
    tf_f2.margin_left = tf_f2.margin_right = tf_f2.margin_top = tf_f2.margin_bottom = 0

    p_f2 = tf_f2.paragraphs[0]
    p_f2.text = "REALISASI MODAL AWAL RELOKASI 2016 (OUTLAY RP 62 JT)"
    p_f2.font.name = "Times New Roman"
    p_f2.font.size = Pt(12.5)
    p_f2.font.bold = True
    p_f2.font.color.rgb = BLUE_ACCENT

    f2_details = [
        ("Gerobak Etalase Utama", "Rp 8.000.000 (Gerobak alumunium kaca stasiun open kitchen depan)"),
        ("Meja Makan & Kursi", "Rp 8.300.000 (4 Meja pendek 2,8 Jt, Meja dinding 1,8 Jt, Kursi 37 pcs 3,7 Jt)"),
        ("Kipas Angin & Pendingin", "Rp 6.900.000 (6 Kipas angin @150k = 900k, Kulkas 2,5 Jt, Freezer 3,5 Jt)"),
        ("Alat Dapur, Panci & Stok Mangkok", "Rp 6.500.000 (Panci masak, dandang, mangkok makan, sendok, sumpit)"),
        ("Renovasi & Signage Ruko", "Rp 11.000.000 (Cat ruko 1 lantai, kelistrikan/air, spanduk/plang nama)"),
        ("Sewa Tempat Ruko 1 Tahun", "Rp 20.000.000 (Alokasi biaya sewa ruko tahun pertama 2016)"),
        ("Biaya Pra-Operasi & Perizinan", "Rp 1.300.000 (AMDAL/kebersihan 300k, survey 500k, promosi awal 500k)")
    ]
    for n, d in f2_details:
        p = tf_f2.add_paragraph()
        p.text = f"• {n} ➔ {d}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(10)
        p.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 13: ALUR PRODUKSI & SIKLUS MINGGUAN ADONAN (Type 4: 5 Steps Process Flow)
    # =================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_background(s13)
    add_header(s13, "Alur Proses Produksi & Siklus Mingguan Adonan", "Tahapan pengolahan higienis dari adonan mingguan hingga siap disajikan", tag="BAB IV: TEKNIS & OPERASI", slide_idx=13)

    steps_prod = [
        ("LANGKAH 01", "Adonan Fresh Mingguan", "Setiap Minggu Pagi", "Adonan mie dibuat mandiri secara konsisten setiap minggu tanpa bahan pengawet kimia."),
        ("LANGKAH 02", "Pengolahan 3 Topping Murni", "05.30 - 07.30 WIB", "Memanggang babi casiu madu, menumis babi kecap gurih, dan merebus potongan ayam putih."),
        ("LANGKAH 03", "Racikan 3 Minyak & Kaldu", "06.00 - 08.00 WIB", "Menyiapkan minyak babi, minyak ayam, dan minyak sayur+wijen serta rebusan kaldu tulang gurih."),
        ("LANGKAH 04", "Perebusan Mie Seketika", "Saat Order Masuk", "Mie direbus 45 detik dalam air mendidih saat nota pesanan diterima dari tamu/kasir."),
        ("LANGKAH 05", "Plating & Penyajian Cepat", "< 5 Menit Total", "Pencampuran minyak pilihan, penataan topping daging melimpah, daun bawang, & saus meja.")
    ]

    for idx, (st, tit, tm, dsc) in enumerate(steps_prod):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(1.9)
        w = Inches(2.2)
        h = Inches(4.9)

        add_card(s13, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 3 else MUTED_BORDER)
        
        # Step Badge
        add_card(s13, x + Inches(0.15), y + Inches(0.2), w - Inches(0.3), Inches(0.4), bg_color=DARK_BG, border_color=GOLD_ACCENT)
        tb_st = s13.shapes.add_textbox(x + Inches(0.15), y + Inches(0.2), w - Inches(0.3), Inches(0.4))
        tb_st.text_frame.margin_left = tb_st.text_frame.margin_right = tb_st.text_frame.margin_top = tb_st.text_frame.margin_bottom = 0
        pst = tb_st.text_frame.paragraphs[0]
        pst.alignment = PP_ALIGN.CENTER
        pst.text = st
        pst.font.name = "Times New Roman"
        pst.font.size = Pt(11)
        pst.font.bold = True
        pst.font.color.rgb = GOLD_ACCENT

        tb_c = s13.shapes.add_textbox(x + Inches(0.15), y + Inches(0.75), w - Inches(0.3), h - Inches(0.9))
        tf_c = tb_c.text_frame
        tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = 0

        p1 = tf_c.paragraphs[0]
        p1.text = tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = WHITE

        p2 = tf_c.add_paragraph()
        p2.text = f"Waktu: {tm}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10)
        p2.font.color.rgb = BLUE_ACCENT

        p3 = tf_c.add_paragraph()
        p3.text = f"\n{dsc}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10.5)
        p3.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 14: JADWAL PELAKSANAAN PRA-OPERASI RELOKASI 2016 (Type 4: Timeline/Gantt)
    # =================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_background(s14)
    add_header(s14, "Jadwal Pelaksanaan Pra-Operasi Relokasi 2016", "Rencana timeline 12 minggu persiapan operasional gerai Duri Kosambi", tag="BAB IV: TEKNIS & OPERASI", slide_idx=14)

    phases = [
        ("FASE 1: SURVEY & SEWA RUKO", "Minggu 1 - 4", GOLD_ACCENT, [
            "A. Observasi lokasi Jl. Angsoka Hijau IV depan Kalam Kudus.",
            "B. Negosiasi dan pembayaran sewa ruko 1 lantai Rp 20.000.000.",
            "C. Pengurusan izin lingkungan RT/RW dan retribusi kebersihan.",
            "D. Analisis potensi pasar sarapan murid dan orang tua sekolah."
        ]),
        ("FASE 2: RENOVASI & FIT-OUT KEDAI", "Minggu 5 - 8", BLUE_ACCENT, [
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
        ("FASE 4: TRIAL RUN & PEMBUKAAN", "Minggu 11 - 12", GOLD_LIGHT, [
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
        y = Inches(1.9 + r * 2.5)
        w = Inches(5.7)
        h = Inches(2.3)

        add_card(s14, x, y, w, h, bg_color=CARD_BG, border_color=p_col)
        tb = s14.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), w - Inches(0.4), h - Inches(0.3))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = p_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = p_col

        p2 = tf.add_paragraph()
        p2.text = f"Jadwal: {p_time}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(10)
        p2.font.color.rgb = GRAY_TEXT

        for itm in p_items:
            pit = tf.add_paragraph()
            pit.text = f"• {itm}"
            pit.font.name = "Times New Roman"
            pit.font.size = Pt(10)
            pit.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 15: RENCANA INVESTASI AWAL (INITIAL OUTLAY) (Type 4: Stat + Breakdown)
    # =================================================================
    s15 = prs.slides.add_slide(blank_layout)
    add_background(s15)
    add_header(s15, "Rencana Modal Investasi Awal (Initial Outlay)", "Alokasi anggaran modal awal relokasi 2016 untuk pembukaan kedai Duri Kosambi", tag="BAB V: ASPEK KEUANGAN", slide_idx=15)

    # Hero Stat Callout Box
    add_card(s15, Inches(0.8), Inches(1.9), Inches(11.7), Inches(1.4), bg_color=CARD_BG, border_color=GOLD_ACCENT, border_width=1.5)
    tb_out = s15.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(11.1), Inches(1.2))
    tf_out = tb_out.text_frame
    tf_out.margin_left = tf_out.margin_right = tf_out.margin_top = tf_out.margin_bottom = 0

    p_out1 = tf_out.paragraphs[0]
    p_out1.text = "TOTAL INITIAL OUTLAY KEBUTUHAN MODAL: RP 62.000.000"
    p_out1.font.name = "Times New Roman"
    p_out1.font.size = Pt(22)
    p_out1.font.bold = True
    p_out1.font.color.rgb = GOLD_ACCENT

    p_out2 = tf_out.add_paragraph()
    p_out2.text = "Investasi awal didanai 100% modal sendiri tanpa utang bank. Seluruh alokasi modal difokuskan untuk pengadaan aktiva tetap gerobak & fasilitas kedai, sewa tempat ruko 1 tahun, dan pra-operasi."
    p_out2.font.name = "Times New Roman"
    p_out2.font.size = Pt(11.5)
    p_out2.font.color.rgb = WHITE

    # 5 Components Cards
    outlay_items = [
        ("01. Aktiva Tetap (Capex)", "Rp 40.700.000", "65,6%", "Pengadaan gerobak etalase, meja, kursi 37 pcs, kipas angin, kulkas, freezer, panci, & mangkok."),
        ("02. Sewa Tempat Ruko", "Rp 20.000.000", "32,3%", "Alokasi biaya sewa ruko 1 lantai Duri Kosambi untuk 1 tahun masa operasional."),
        ("03. Biaya Perizinan AMDAL", "Rp 300.000", "0,5%", "Retribusi kebersihan lingkungan RT/RW dan pengadaan tempat sampah kedai higienis."),
        ("04. Biaya Survey Pasar", "Rp 500.000", "0,8%", "Observasi demografi perumahan Kosambi dan potensi kantin sekolah Kalam Kudus."),
        ("05. Biaya Promosi Awal", "Rp 500.000", "0,8%", "Pembuatan spanduk pembukaan dan plang penunjuk arah kedai Bakmi Mimu.")
    ]

    for idx, (c_name, c_val, c_pct, c_dsc) in enumerate(outlay_items):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(3.55)
        w = Inches(2.2)
        h = Inches(3.2)

        add_card(s15, x, y, w, h, bg_color=CARD_BG, border_color=MUTED_BORDER)
        tb = s15.shapes.add_textbox(x + Inches(0.15), y + Inches(0.2), w - Inches(0.3), h - Inches(0.4))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = c_name
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = WHITE

        p2 = tf.add_paragraph()
        p2.text = c_val
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = GOLD_ACCENT

        p3 = tf.add_paragraph()
        p3.text = f"Porsi: {c_pct}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.color.rgb = GREEN_ACCENT

        p4 = tf.add_paragraph()
        p4.text = f"\n{c_dsc}"
        p4.font.name = "Times New Roman"
        p4.font.size = Pt(10)
        p4.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 16: RINCIAN AKTIVA TETAP & DEPRESIASI (Type 4: Table & Stat)
    # =================================================================
    s16 = prs.slides.add_slide(blank_layout)
    add_background(s16)
    add_header(s16, "Aktiva Tetap & Beban Depresiasi", "Inventarisasi 12 item aktiva tetap dan penyusutan garis lurus 20% per tahun", tag="BAB V: ASPEK KEUANGAN", slide_idx=16)

    # Stat box left
    add_card(s16, Inches(0.8), Inches(1.9), Inches(3.5), Inches(4.9), bg_color=CARD_BG, border_color=GOLD_ACCENT)
    tb_st = s16.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(3.1), Inches(4.5))
    tf_st = tb_st.text_frame
    tf_st.margin_left = tf_st.margin_right = tf_st.margin_top = tf_st.margin_bottom = 0

    ps1 = tf_st.paragraphs[0]
    ps1.text = "PARAMETER DEPRESIASI"
    ps1.font.name = "Times New Roman"
    ps1.font.size = Pt(13)
    ps1.font.bold = True
    ps1.font.color.rgb = GOLD_ACCENT

    ps2 = tf_st.add_paragraph()
    ps2.text = f"\nTotal Nilai Perolehan:\nRp {met['capex_total']:,.0f}".replace(",", ".")
    ps2.font.name = "Times New Roman"
    ps2.font.size = Pt(15)
    ps2.font.bold = True
    ps2.font.color.rgb = WHITE

    ps3 = tf_st.add_paragraph()
    ps3.text = "\nMetode Penyusutan:\nMetode Garis Lurus (Straight-Line)"
    ps3.font.name = "Times New Roman"
    ps3.font.size = Pt(11)
    ps3.font.color.rgb = GRAY_TEXT

    ps4 = tf_st.add_paragraph()
    ps4.text = "\nTarif Depresiasi Tahunan:\n20,00% (Masa Manfaat 5 Tahun)"
    ps4.font.name = "Times New Roman"
    ps4.font.size = Pt(11)
    ps4.font.color.rgb = GRAY_TEXT

    ps5 = tf_st.add_paragraph()
    ps5.text = f"\nBeban Depresiasi / Tahun:\nRp {met['depresiasi_per_year']:,.0f}".replace(",", ".")
    ps5.font.name = "Times New Roman"
    ps5.font.size = Pt(15)
    ps5.font.bold = True
    ps5.font.color.rgb = GREEN_ACCENT

    # Right Card: 6 Kelompok Aktiva Tetap
    add_card(s16, Inches(4.6), Inches(1.9), Inches(7.9), Inches(4.9), bg_color=CARD_BG, border_color=MUTED_BORDER)
    tb_rt = s16.shapes.add_textbox(Inches(4.9), Inches(2.1), Inches(7.3), Inches(4.5))
    tf_rt = tb_rt.text_frame
    tf_rt.margin_left = tf_rt.margin_right = tf_rt.margin_top = tf_rt.margin_bottom = 0

    pr1 = tf_rt.paragraphs[0]
    pr1.text = "REKAPITULASI KELOMPOK AKTIVA TETAP (12 ITEM CAPEX)"
    pr1.font.name = "Times New Roman"
    pr1.font.size = Pt(13)
    pr1.font.bold = True
    pr1.font.color.rgb = GOLD_ACCENT

    capex_groups = [
        ("1. Gerobak & Etalase Kaca Depan", "Rp 8.000.000", "Stasiun masak open kitchen teras ruko, dandang mie, kompor gas, dan kaca display."),
        ("2. Perabotan Dine-In (Meja & Kursi)", "Rp 8.300.000", "4 Meja pendek kayu (2,8 Jt), Meja panjang dinding (1,8 Jt), Kursi plastik 37 pcs (3,7 Jt)."),
        ("3. Sarana Sirkulasi Udara (6 Kipas Angin)", "Rp 900.000", "6 Unit kipas angin dinding/plafon sejuk (@Rp 150.000) untuk ruang makan 22 kursi."),
        ("4. Mesin Pendingin Dapur (Kulkas & Freezer)", "Rp 6.000.000", "1 Unit kulkas pendingin sayur & saus (2,5 Jt) dan 1 unit chest freezer daging beku (3,5 Jt)."),
        ("5. Panci Masak, Mangkok & Alat Makan", "Rp 6.500.000", "Stok mangkok keramik, sendok, sumpit (2,5 Jt), panci masak kaldu (3 Jt), botol saus (1 Jt)."),
        ("6. Renovasi Sederhana Ruko & Signage", "Rp 11.000.000", "Pengecatan interior ruko 1 lantai, instalasi air/listrik (9 Jt), dan plang nama kedai (2 Jt).")
    ]
    for grp, val, dtl in capex_groups:
        p = tf_rt.add_paragraph()
        p.text = f"\n• {grp} ➔ {val}\n   {dtl}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(10.5)
        p.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 17: PROYEKSI PENDAPATAN (CASH INFLOW 5 TAHUN) (Type 4: Trend Visual)
    # =================================================================
    s17 = prs.slides.add_slide(blank_layout)
    add_background(s17)
    add_header(s17, "Proyeksi Pendapatan (Cash Inflow 5 Tahun)", "Estimasi penerimaan kas berbasis 300 hari operasi tahunan dengan eskalasi volume riil", tag="BAB V: ASPEK KEUANGAN", slide_idx=17)

    cif_years = [
        ("Tahun 1 (2016/17)", met['cif_years'][0], "Basis Awal", "Omzet riil 300 hari operasi: rata-rata Rp 4,94 Jt/hari (menu utama + pelengkap)"),
        ("Tahun 2 (2017/18)", met['cif_years'][1], f"+{(met['cif_years'][1]/met['cif_years'][0]-1)*100:.1f}%", "Peningkatan loyalitas pelanggan sekolah Kalam Kudus & warga sekitar Kosambi"),
        ("Tahun 3 (2018/19)", met['cif_years'][2], f"+{(met['cif_years'][2]/met['cif_years'][1]-1)*100:.1f}%", "Basis pelanggan mantap terbentuk (2019), pesanan bungkus take-away meningkat"),
        ("Tahun 4 (2019/20)", met['cif_years'][3], f"+{(met['cif_years'][3]/met['cif_years'][2]-1)*100:.1f}%", "Penjualan stabil dengan pelanggan tetap keluarga residensial perumahan"),
        ("Tahun 5 (2020/21)", met['cif_years'][4], f"+{(met['cif_years'][4]/met['cif_years'][3]-1)*100:.1f}%", "Kapasitas kedai ruko optimal & persiapan rencana ekspansi cabang 2027/28")
    ]

    for idx, (yr, val, grw, dsc) in enumerate(cif_years):
        x = Inches(0.8 + idx * 2.38)
        y = Inches(1.9)
        w = Inches(2.2)
        h = Inches(4.9)

        add_card(s17, x, y, w, h, bg_color=CARD_BG, border_color=GOLD_ACCENT if idx == 0 else MUTED_BORDER)
        tb = s17.shapes.add_textbox(x + Inches(0.15), y + Inches(0.25), w - Inches(0.3), h - Inches(0.5))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = yr
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = WHITE

        p2 = tf.add_paragraph()
        p2.text = f"Rp {val:,.0f}".replace(",", ".")
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = GOLD_ACCENT

        p3 = tf.add_paragraph()
        p3.text = f"Pertumbuhan: {grw}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.color.rgb = GREEN_ACCENT

        p4 = tf.add_paragraph()
        p4.text = f"\n{dsc}"
        p4.font.name = "Times New Roman"
        p4.font.size = Pt(10)
        p4.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 18: PROYEKSI BIAYA & LABA BERSIH (NCF/EAT) (Type 4: Comparative Cards)
    # =================================================================
    s18 = prs.slides.add_slide(blank_layout)
    add_background(s18)
    add_header(s18, "Proyeksi Biaya & Laba Bersih (NCF / EAT)", "Analisis arus kas keluar operasional dan laba bersih selama 5 tahun", tag="BAB V: ASPEK KEUANGAN", slide_idx=18)

    y1_metrics = [
        ("CASH INFLOW (CIF)", f"Rp {met['cif_years'][0]:,.0f}".replace(",", "."), BLUE_ACCENT, "Total omzet penjualan bakmi, menu kuah swikiaw/pangsit, & aneka minuman."),
        ("CASH OUTFLOW (COF)", f"Rp {met['cof_years'][0]:,.0f}".replace(",", "."), GOLD_LIGHT, "Total HPP bahan baku, gaji 3 staf (Rp 120 Jt), sewa ruko (Rp 20 Jt), gas & galon."),
        ("NET CASH FLOW (EAT)", f"Rp {met['ncf_years'][0]:,.0f}".replace(",", "."), GREEN_ACCENT, "Laba bersih operasional (CIF dikurangi COF), margin keuntungan 61,97%."),
        ("PROCEED (NCF - DEPR/ADJ)", f"Rp {met['proceed_years'][0]:,.0f}".replace(",", "."), GOLD_ACCENT, "Arus kas bersih riil yang memperhitungkan penyusutan aktiva tetap.")
    ]

    for idx, (m_tit, m_val, m_col, m_dsc) in enumerate(y1_metrics):
        x = Inches(0.8 + idx * 2.98)
        y = Inches(1.9)
        w = Inches(2.78)
        h = Inches(2.4)

        add_card(s18, x, y, w, h, bg_color=CARD_BG, border_color=m_col)
        tb = s18.shapes.add_textbox(x + Inches(0.2), y + Inches(0.2), w - Inches(0.4), h - Inches(0.4))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = m_tit
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = m_col

        p2 = tf.add_paragraph()
        p2.text = m_val
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(15.5)
        p2.font.bold = True
        p2.font.color.rgb = WHITE

        p3 = tf.add_paragraph()
        p3.text = f"\n{m_dsc}"
        p3.font.name = "Times New Roman"
        p3.font.size = Pt(10)
        p3.font.color.rgb = GRAY_TEXT

    # Bottom summary card: 5 Year Trend
    add_card(s18, Inches(0.8), Inches(4.5), Inches(11.7), Inches(2.3), bg_color=CARD_BG, border_color=MUTED_BORDER)
    tb_btm = s18.shapes.add_textbox(Inches(1.1), Inches(4.7), Inches(11.1), Inches(1.9))
    tf_btm = tb_btm.text_frame
    tf_btm.margin_left = tf_btm.margin_right = tf_btm.margin_top = tf_btm.margin_bottom = 0

    pb1 = tf_btm.paragraphs[0]
    pb1.text = "RINGKASAN TREN KINERJA KEUANGAN 5 TAHUN KEDAI DURI KOSAMBI"
    pb1.font.name = "Times New Roman"
    pb1.font.size = Pt(12.5)
    pb1.font.bold = True
    pb1.font.color.rgb = GOLD_ACCENT

    pb2 = tf_btm.add_paragraph()
    pb2.text = f"• Tahun 1: NCF Rp {met['ncf_years'][0]/1e6:,.2f} Jt  |  Proceed Rp {met['proceed_years'][0]/1e6:,.2f} Jt  (Margin Operasional: 61,97%)\n".replace(",", ".") + \
               f"• Tahun 2: NCF Rp {met['ncf_years'][1]/1e6:,.2f} Jt  |  Proceed Rp {met['proceed_years'][1]/1e6:,.2f} Jt  (Pertumbuhan Omzet Solid)\n".replace(",", ".") + \
               f"• Tahun 3: NCF Rp {met['ncf_years'][2]/1e6:,.2f} Jt  |  Proceed Rp {met['proceed_years'][2]/1e6:,.2f} Jt  (Basis Pelanggan Solid Terbentuk)\n".replace(",", ".") + \
               f"• Tahun 4: NCF Rp {met['ncf_years'][3]/1e6:,.2f} Jt  |  Proceed Rp {met['proceed_years'][3]/1e6:,.2f} Jt  (Arus Kas Operasional Kuat)\n".replace(",", ".") + \
               f"• Tahun 5: NCF Rp {met['ncf_years'][4]/1e6:,.2f} Jt  |  Proceed Rp {met['proceed_years'][4]/1e6:,.2f} Jt  (Akumulasi Kas untuk Ekspansi 2027/28)".replace(",", ".")
    pb2.font.name = "Times New Roman"
    pb2.font.size = Pt(10.5)
    pb2.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 19: EVALUASI MODAL: NET PRESENT VALUE (NPV) (Type 4: Hero Stat Callout)
    # =================================================================
    s19 = prs.slides.add_slide(blank_layout)
    add_background(s19)
    add_header(s19, "Evaluasi Modal: Net Present Value (NPV)", "Hasil perhitungan nilai sekarang bersih kas masuk dengan discount rate 20%", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=19)

    # Giant Stat Box Left
    add_card(s19, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=GREEN_ACCENT, border_width=2.0)
    tb_npv = s19.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.1), Inches(4.3))
    tf_npv = tb_npv.text_frame
    tf_npv.margin_left = tf_npv.margin_right = tf_npv.margin_top = tf_npv.margin_bottom = 0

    pn1 = tf_npv.paragraphs[0]
    pn1.text = "NET PRESENT VALUE (NPV)"
    pn1.font.name = "Times New Roman"
    pn1.font.size = Pt(15)
    pn1.font.bold = True
    pn1.font.color.rgb = GOLD_ACCENT

    pn2 = tf_npv.add_paragraph()
    pn2.text = f"Rp {met['npv']:,.0f}".replace(",", ".")
    pn2.font.name = "Times New Roman"
    pn2.font.size = Pt(28)
    pn2.font.bold = True
    pn2.font.color.rgb = GREEN_ACCENT

    pn3 = tf_npv.add_paragraph()
    pn3.text = "\nSTATUS KELAYAKAN: SANGAT LAYAK (POSITIVE)"
    pn3.font.name = "Times New Roman"
    pn3.font.size = Pt(13)
    pn3.font.bold = True
    pn3.font.color.rgb = WHITE

    pn4 = tf_npv.add_paragraph()
    pn4.text = f"\nBerdasarkan kriteria standar FEB UKRIDA, rencana investasi dinyatakan LAYAK jika NPV > 0. Nilai NPV sebesar Rp {met['npv']/1e9:.2f} Milyar membuktikan bahwa proyek Bakmi Mimu Carina Sayang mampu menutupi seluruh modal investasi awal Rp 62 Juta dan menghasilkan surplus kekayaan bersih yang sangat luar biasa."
    pn4.font.name = "Times New Roman"
    pn4.font.size = Pt(11)
    pn4.font.color.rgb = GRAY_TEXT

    # Right Table Card: Discounted PV
    add_card(s19, Inches(6.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=MUTED_BORDER)
    tb_pvt = s19.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_pvt = tb_pvt.text_frame
    tf_pvt.margin_left = tf_pvt.margin_right = tf_pvt.margin_top = tf_pvt.margin_bottom = 0

    ppv1 = tf_pvt.paragraphs[0]
    ppv1.text = "RINCIAN DISCOUNTED PRESENT VALUE (RATE 20%)"
    ppv1.font.name = "Times New Roman"
    ppv1.font.size = Pt(13)
    ppv1.font.bold = True
    ppv1.font.color.rgb = GOLD_ACCENT

    pv_rows = [
        ("Tahun 0 (Initial Outlay)", f"-Rp {met['initial_outlay']:,.0f}".replace(",", "."), "Investasi Awal"),
        ("Tahun 1 (Proceed Thn 1)", f"Rp {met['pv_years'][0]:,.0f}".replace(",", "."), "DF (20%): 0.8333"),
        ("Tahun 2 (Proceed Thn 2)", f"Rp {met['pv_years'][1]:,.0f}".replace(",", "."), "DF (20%): 0.6944"),
        ("Tahun 3 (Proceed Thn 3)", f"Rp {met['pv_years'][2]:,.0f}".replace(",", "."), "DF (20%): 0.5787"),
        ("Tahun 4 (Proceed Thn 4)", f"Rp {met['pv_years'][3]:,.0f}".replace(",", "."), "DF (20%): 0.4823"),
        ("Tahun 5 (Proceed Thn 5)", f"Rp {met['pv_years'][4]:,.0f}".replace(",", "."), "DF (20%): 0.4019")
    ]
    for yr, val, df in pv_rows:
        p = tf_pvt.add_paragraph()
        p.text = f"\n• {yr} ➔ {val} ({df})"
        p.font.name = "Times New Roman"
        p.font.size = Pt(10.5)
        p.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 20: EVALUASI MODAL: IRR & PAYBACK PERIOD (Type 4: Dual Hero Stat)
    # =================================================================
    s20 = prs.slides.add_slide(blank_layout)
    add_background(s20)
    add_header(s20, "Evaluasi Modal: IRR & Payback Period", "Tingkat pengembalian internal dan kecepatan pemulihan modal investasi", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=20)

    # Hero Stat Left: IRR
    add_card(s20, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=GREEN_ACCENT, border_width=1.5)
    tb_irr = s20.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.1), Inches(4.3))
    tf_irr = tb_irr.text_frame
    tf_irr.margin_left = tf_irr.margin_right = tf_irr.margin_top = tf_irr.margin_bottom = 0

    pi1 = tf_irr.paragraphs[0]
    pi1.text = "INTERNAL RATE OF RETURN (IRR)"
    pi1.font.name = "Times New Roman"
    pi1.font.size = Pt(15)
    pi1.font.bold = True
    pi1.font.color.rgb = GOLD_ACCENT

    pi2 = tf_irr.add_paragraph()
    pi2.text = "> 20,00%"
    pi2.font.name = "Times New Roman"
    pi2.font.size = Pt(38)
    pi2.font.bold = True
    pi2.font.color.rgb = GREEN_ACCENT

    pi3 = tf_irr.add_paragraph()
    pi3.text = "\nTARGET / SYARAT: > 20,00% (Opportunity Cost)"
    pi3.font.name = "Times New Roman"
    pi3.font.size = Pt(12.5)
    pi3.font.bold = True
    pi3.font.color.rgb = WHITE

    pi4 = tf_irr.add_paragraph()
    pi4.text = "\nSTATUS: SANGAT LAYAK (Jauh Melampaui Standar)"
    pi4.font.name = "Times New Roman"
    pi4.font.size = Pt(12)
    pi4.font.bold = True
    pi4.font.color.rgb = GREEN_ACCENT

    pi5 = tf_irr.add_paragraph()
    pi5.text = f"\nNilai IRR yang melampaui 20,00% membuktikan efisiensi operasional kedai yang sangat prima. Imbal hasil riil proyek jauh melampaui tingkat diskonto modal dan suku bunga perbankan."
    pi5.font.name = "Times New Roman"
    pi5.font.size = Pt(10.5)
    pi5.font.color.rgb = GRAY_TEXT

    # Hero Stat Right: Payback Period
    add_card(s20, Inches(6.8), Inches(1.9), Inches(5.7), Inches(4.9), bg_color=CARD_BG, border_color=BLUE_ACCENT, border_width=1.5)
    tb_pp = s20.shapes.add_textbox(Inches(7.1), Inches(2.2), Inches(5.1), Inches(4.3))
    tf_pp = tb_pp.text_frame
    tf_pp.margin_left = tf_pp.margin_right = tf_pp.margin_top = tf_pp.margin_bottom = 0

    ppp1 = tf_pp.paragraphs[0]
    ppp1.text = "PAYBACK PERIOD (PP)"
    ppp1.font.name = "Times New Roman"
    ppp1.font.size = Pt(15)
    ppp1.font.bold = True
    ppp1.font.color.rgb = BLUE_ACCENT

    ppp2 = tf_pp.add_paragraph()
    ppp2.text = f"{met['pp_months']:.2f} BULAN"
    ppp2.font.name = "Times New Roman"
    ppp2.font.size = Pt(38)
    ppp2.font.bold = True
    ppp2.font.color.rgb = BLUE_ACCENT

    ppp3 = tf_pp.add_paragraph()
    ppp3.text = f"\nSETARA: {met['pp_years']:.2f} TAHUN (KURANG DARI 1 BULAN OPERASI)"
    ppp3.font.name = "Times New Roman"
    ppp3.font.size = Pt(12)
    ppp3.font.bold = True
    ppp3.font.color.rgb = WHITE

    ppp4 = tf_pp.add_paragraph()
    ppp4.text = "\nTARGET / SYARAT: < 3,00 TAHUN (36 BULAN) ➔ SANGAT LAYAK"
    ppp4.font.name = "Times New Roman"
    ppp4.font.size = Pt(12)
    ppp4.font.bold = True
    ppp4.font.color.rgb = GREEN_ACCENT

    ppp5 = tf_pp.add_paragraph()
    ppp5.text = f"\nModal awal sebesar Rp 62.000.000 telah pulih kembali sepenuhnya hanya dalam tempo {met['pp_months']:.2f} bulan (sekitar 25 hari kerja pada bulan pertama). Seluruh arus kas berikutnya murni menjadi keuntungan likuid."
    ppp5.font.name = "Times New Roman"
    ppp5.font.size = Pt(10.5)
    ppp5.font.color.rgb = GRAY_TEXT

    # =================================================================
    # SLIDE 21: EVALUASI MODAL: PROFITABILITY INDEX & SENSITIVITAS (Type 4: Scoreboard)
    # =================================================================
    s21 = prs.slides.add_slide(blank_layout)
    add_background(s21)
    add_header(s21, "Profitability Index & Matriks Kelayakan", "Rangkuman skor evaluasi investasi dan ketahanan terhadap risiko pasar", tag="BAB VI: KELAYAKAN INVESTASI", slide_idx=21)

    # 4 Criteria Cards
    crit_list = [
        ("NET PRESENT VALUE (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "> Rp 0 (Positif)", "SANGAT LAYAK", GREEN_ACCENT),
        ("INTERNAL RATE OF RETURN (IRR)", "> 20,00%", "> 20,00%", "SANGAT LAYAK", GREEN_ACCENT),
        ("PAYBACK PERIOD (PP)", f"{met['pp_months']:.2f} Bulan", "< 36 Bulan (3 Thn)", "SANGAT LAYAK", GREEN_ACCENT),
        ("PROFITABILITY INDEX (PI)", f"{met['pi']:.2f}", "> 1,20", "SANGAT LAYAK", GREEN_ACCENT)
    ]

    for idx, (c_name, c_val, c_req, c_st, c_col) in enumerate(crit_list):
        c = idx % 2
        r = idx // 2
        x = Inches(0.8 + c * 6.0)
        y = Inches(1.9 + r * 1.6)
        w = Inches(5.7)
        h = Inches(1.4)

        add_card(s21, x, y, w, h, bg_color=CARD_BG, border_color=c_col)
        tb = s21.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), w - Inches(0.4), h - Inches(0.3))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = c_name
        p1.font.name = "Times New Roman"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = GOLD_ACCENT

        p2 = tf.add_paragraph()
        p2.text = f"{c_val}   [Standar: {c_req}]   ➔  {c_st}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = WHITE

    # Bottom card: Analisis Sensitivitas
    add_card(s21, Inches(0.8), Inches(5.3), Inches(11.7), Inches(1.5), bg_color=CARD_BG, border_color=BLUE_ACCENT)
    tb_sens = s21.shapes.add_textbox(Inches(1.1), Inches(5.45), Inches(11.1), Inches(1.2))
    tf_sens = tb_sens.text_frame
    tf_sens.margin_left = tf_sens.margin_right = tf_sens.margin_top = tf_sens.margin_bottom = 0

    ps1 = tf_sens.paragraphs[0]
    ps1.text = "ANALISIS SENSITIVITAS & KETAHANAN TERHADAP RISIKO"
    ps1.font.name = "Times New Roman"
    ps1.font.size = Pt(12)
    ps1.font.bold = True
    ps1.font.color.rgb = BLUE_ACCENT

    ps2 = tf_sens.add_paragraph()
    ps2.text = "• Skenario Kenaikan Harga Bahan Baku Daging (+10%): Nilai NPV tetap positif di atas Rp 5,2 Milyar dengan Payback Period < 1 bulan.\n" \
               "• Skenario Penurunan Volume Penjualan (-15%): Kedai tetap menghasilkan arus kas operasional surplus yang sangat sehat untuk menutup gaji dan sewa."
    ps2.font.name = "Times New Roman"
    ps2.font.size = Pt(10.5)
    ps2.font.color.rgb = WHITE

    # =================================================================
    # SLIDE 22: KESIMPULAN AKHIR & PENUTUP (Type 5: Executive Outro)
    # =================================================================
    s22 = prs.slides.add_slide(blank_layout)
    add_background(s22)

    # Grand Box Container
    add_card(s22, Inches(1.0), Inches(0.5), Inches(11.333), Inches(6.5), bg_color=CARD_BG, border_color=GREEN_ACCENT, border_width=2.0)

    # UKRIDA Logo at top center
    if os.path.exists(logo_path):
        s22.shapes.add_picture(logo_path, Inches(6.066), Inches(0.75), Inches(1.2), Inches(1.2))

    tb_end = s22.shapes.add_textbox(Inches(1.3), Inches(2.05), Inches(10.733), Inches(4.8))
    tf_end = tb_end.text_frame
    tf_end.margin_left = tf_end.margin_right = tf_end.margin_top = tf_end.margin_bottom = 0

    pe1 = tf_end.paragraphs[0]
    pe1.alignment = PP_ALIGN.CENTER
    pe1.text = "KESIMPULAN AKHIR KELAYAKAN INVESTASI"
    pe1.font.name = "Times New Roman"
    pe1.font.size = Pt(16)
    pe1.font.bold = True
    pe1.font.color.rgb = GOLD_ACCENT

    pe2 = tf_end.add_paragraph()
    pe2.alignment = PP_ALIGN.CENTER
    pe2.text = "“ BISNIS DINYATAKAN SANGAT LAYAK ”"
    pe2.font.name = "Times New Roman"
    pe2.font.size = Pt(32)
    pe2.font.bold = True
    pe2.font.color.rgb = GREEN_ACCENT

    pe3 = tf_end.add_paragraph()
    pe3.alignment = PP_ALIGN.CENTER
    pe3.text = f"Berdasarkan hasil analisis komprehensif aspek pasar, manajemen, teknis operasi, dan finansial,\n" \
               f"proyek Bakmi Mimu Carina Sayang menghasilkan NPV Rp {met['npv']:,.0f}".replace(",", ".") + \
               f", IRR > 20,00%, Payback Period {met['pp_months']:.2f} bulan (hanya 0,07 tahun), dan PI {met['pi']:.2f}."
    pe3.font.name = "Times New Roman"
    pe3.font.size = Pt(12)
    pe3.font.color.rgb = WHITE

    pe4 = tf_end.add_paragraph()
    pe4.alignment = PP_ALIGN.CENTER
    pe4.text = "\nDisusun Oleh Tim Mahasiswa:\n" \
               "Arthur Reezan (312023002)   •   Jennese Putra Alamsyah Sukadi (312023033)   •   Valendrik Dwiputra Wirawan (312023013)\n" \
               "Affandy (312023075)   •   Steven Putra Tjhin (312023015)"
    pe4.font.name = "Times New Roman"
    pe4.font.size = Pt(11)
    pe4.font.bold = True
    pe4.font.color.rgb = GOLD_LIGHT

    pe5 = tf_end.add_paragraph()
    pe5.alignment = PP_ALIGN.CENTER
    pe5.text = "\nSekian & Terima Kasih  |  Sesi Tanya Jawab Dibuka\n" \
               "Program Studi Manajemen  |  Fakultas Ekonomi & Bisnis  |  Universitas Kristen Krida Wacana"
    pe5.font.name = "Times New Roman"
    pe5.font.size = Pt(10.5)
    pe5.font.italic = True
    pe5.font.color.rgb = GRAY_TEXT

    prs.save(out_pptx)
    print(f"[SUCCESS] Presentasi Eksekutif Modern berhasil dibuat di: {out_pptx}")

if __name__ == "__main__":
    create_presentation()
