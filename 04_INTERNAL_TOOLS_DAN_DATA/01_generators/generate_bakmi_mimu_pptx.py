import sys
import os
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    with open(r"d:\Perkuliahan\Kelass\SKB\config_bakmi_mimu.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)
    with open(r"d:\Perkuliahan\Kelass\SKB\metrics_bakmi_mimu.json", "r", encoding="utf-8") as f:
        met = json.load(f)

    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Theme Colors
    DARK_BG = RGBColor(20, 29, 44)       # Deep Navy/Slate
    CARD_BG = RGBColor(32, 45, 68)       # Lighter Navy Card
    GOLD_ACCENT = RGBColor(230, 168, 64) # Warm Gold
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(190, 200, 215)
    GREEN_ACCENT = RGBColor(46, 204, 113)

    def add_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BG
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, subtitle_text=None, tag="STUDI KELAYAKAN BISNIS"):
        # Tag
        tb_tag = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.4))
        p_tag = tb_tag.text_frame.paragraphs[0]
        p_tag.text = tag.upper()
        p_tag.font.name = "Times New Roman"
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = GOLD_ACCENT

        # Title
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        p.text = title_text
        p.font.name = "Times New Roman"
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = WHITE

        if subtitle_text:
            p2 = tb.text_frame.add_paragraph()
            p2.text = subtitle_text
            p2.font.name = "Times New Roman"
            p2.font.size = Pt(13)
            p2.font.color.rgb = GRAY_TEXT

    # -----------------------------------------------------------------
    # SLIDE 1: COVER
    # -----------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    add_background(s1)

    # Accent decorative box
    dec = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1))
    dec.fill.solid()
    dec.fill.fore_color.rgb = CARD_BG
    dec.line.color.rgb = GOLD_ACCENT
    dec.line.width = Pt(1.5)

    tb1 = s1.shapes.add_textbox(Inches(1.8), Inches(1.8), Inches(9.7), Inches(3.8))
    tf1 = tb1.text_frame
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = "STUDI KELAYAKAN BISNIS (SKB)"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(18)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT

    p2 = tf1.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = f"“ {cfg['business_name'].upper()} ”"
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = WHITE

    p3 = tf1.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.text = f"- {cfg['tagline']} -"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(16)
    p3.font.italic = True
    p3.font.color.rgb = GRAY_TEXT

    p4 = tf1.add_paragraph()
    p4.alignment = PP_ALIGN.CENTER
    p4.text = f"\n\nDisusun Oleh: {cfg['author']['name']} (NIM: {cfg['author']['nim']})\n{cfg['author']['institution']} | {cfg['author']['faculty']} | {cfg['author']['year']}"
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(13)
    p4.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 2: AGENDA / TABLE OF CONTENTS
    # -----------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_background(s2)
    add_header(s2, "Daftar Isi & Struktur Laporan", "Gambaran komprehensif kajian studi kelayakan bisnis")

    agendas = [
        ("01", "PENDAHULUAN & PROFIL USAHA", "Latar belakang kuliner legendaris sejak 1999, peluang pasar perumahan, dan konsep modernisasi gerai."),
        ("02", "ASPEK PASAR & PEMASARAN", "Segmentasi demografis/geografis, 5 pilar diferensiasi, aspek lokasi strategis, dan bauran 7P."),
        ("03", "MANAJEMEN & TEKNIS OPERASI", "Struktur organisasi, job description, layout ruko, kebutuhan aktiva tetap, dan network planning A-L."),
        ("04", "ASPEK KEUANGAN & KELAYAKAN", "Initial outlay, proyeksi 5 tahun arus kas, depresiasi 20%, serta evaluasi NPV, IRR, PP, dan PI.")
    ]
    for idx, (num, title, desc) in enumerate(agendas):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(2.0 + (idx // 2) * 2.5)
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT
        
        tb = s2.shapes.add_textbox(x + Inches(0.3), y + Inches(0.2), Inches(5.1), Inches(1.8))
        tf = tb.text_frame
        p_n = tf.paragraphs[0]
        p_n.text = f"BAGIAN {num}"
        p_n.font.name = "Times New Roman"
        p_n.font.size = Pt(11)
        p_n.font.bold = True
        p_n.font.color.rgb = GOLD_ACCENT

        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = WHITE

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = GRAY_TEXT

    # -----------------------------------------------------------------
    # SLIDE 3: FILOSOFI & LATAR BELAKANG
    # -----------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    add_background(s3)
    add_header(s3, "Latar Belakang & Filosofi Usaha", "Eksistensi kuliner bakmi legendaris sejak 1999 di Jakarta Barat")

    pts3 = [
        ("Sejarah Tradisi Bakmi", "Bakmi merupakan hidangan akulturasi budaya Tionghoa dan Indonesia yang telah menjadi makanan pokok alternatif paling digemari lintas generasi."),
        ("Otentisitas Sejak 1999", "Bakmi Mimu Carina Sayang telah melayani pelanggan selama lebih dari 25 tahun di kawasan Duri Kosambi dengan resep keluarga turun-temurun."),
        ("Kualitas Tanpa Pengawet", "Menggunakan adonan mie segar bertekstur kenyal berkilau alami (shining) yang diproduksi higienis setiap pagi tanpa bahan kimia pengawet."),
        ("Peluang Modernisasi Gerai", "Ekspansi ruko baru memberikan kenyamanan ruang dine-in ber-AC, standar sanitasi open-kitchen, serta integrasi pemesanan digital GoFood/GrabFood.")
    ]
    for idx, (title, desc) in enumerate(pts3):
        x = Inches(0.8 + idx * 2.95)
        y = Inches(2.2)
        card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(2.8), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s3.shapes.add_textbox(x + Inches(0.2), y + Inches(0.3), Inches(2.4), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(15)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{desc}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 4: KARAKTERISTIK PRODUK & KEUNGGULAN
    # -----------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    add_background(s4)
    add_header(s4, "Karakteristik & Varian Produk", "Kombinasi tekstur mie kenyal, topping gurih, dan pelengkap homemade")

    prods_info = [
        ("Bakmi Spesial Campur", "Signature menu dengan kombinasi topping ayam jamur cincang, babi kecap, dan casiu merah manis gurih."),
        ("Swikiauw Homemade", "Olahan daging cincang dan udang segar bertekstur juicy dengan bumbu khas rebusan kaldu wangi."),
        ("Pangsit Goreng / Rebus", "Kulit pangsit renyah dengan isian daging gurih yang melengkapi kenikmatan semangkuk bakmi."),
        ("Minuman Tradisional", "Pilihan penyegar otentik: Es Jeruk Sonkit Kietna, Es Liang Teh Herbal Medan, dan Badak Sarsaparilla.")
    ]
    for idx, (t, d) in enumerate(prods_info):
        x = Inches(0.8 + idx * 2.95)
        card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(2.8), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s4.shapes.add_textbox(x + Inches(0.2), Inches(2.5), Inches(2.4), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 5: DAFTAR MENU & STRUKTUR HARGA (Tabel 1.1)
    # -----------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    add_background(s5)
    add_header(s5, "Daftar Menu & Penetapan Harga", "Struktur harga terjangkau dan bernilai tinggi (Value-Based Pricing)")

    # Add Table
    rows = 7
    cols = 4
    left = Inches(0.8)
    top = Inches(2.0)
    width = Inches(11.7)
    height = Inches(4.6)
    table_shape = s5.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    table_headers = ["Kategori", "Nama Menu Hidangan", "Deskripsi / Porsi", "Harga Satuan"]
    for c_idx, h in enumerate(table_headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = GOLD_ACCENT
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Times New Roman"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = DARK_BG

    menu_samples = [
        ("Bakmi Utama", "Bakmi Spesial Campur", "Topping Ayam Jamur, Babi Kecap & Casiu", "Rp 37.000"),
        ("Bakmi Utama", "Bakmi Daging Babi Casiu", "Mie kenyal dengan casiu merah manis gurih", "Rp 34.000"),
        ("Bakmi Utama", "Bakmi Daging Ayam Jamur", "Mie kenyal dengan olahan ayam jamur gurih", "Rp 30.000"),
        ("Pelengkap", "Swikiauw Homemade (5 pcs)", "Isian daging babi dan udang segar kuah kaldu", "Rp 32.000"),
        ("Pelengkap", "Pangsit Rebus / Goreng (5 pcs)", "Olahan pangsit gurih renyah homemade", "Rp 25.000"),
        ("Minuman", "Es Jeruk Sonkit / Liang Teh", "Minuman segar tradisional pendamping makan", "Rp 12.000 - 15.000")
    ]
    for r_idx, row_data in enumerate(menu_samples, start=1):
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Times New Roman"
            p.font.size = Pt(12)
            p.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 6: 5 PILAR DIFERENSIASI USAHA
    # -----------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    add_background(s6)
    add_header(s6, "5 Pilar Diferensiasi Usaha", "Keunggulan kompetitif unik dalam memenangkan pasar")

    diffs = [
        ("1. RASA", "Tekstur mie kenyal berkilau alami tanpa pengawet dipadu topping babi casiu, babi kecap, dan ayam gurih kaldu asli."),
        ("2. LOKASI", "Berada di pusat keramaian perumahan Duri Kosambi persis di depan Sekolah Kristen Kalam Kudus."),
        ("3. LAYANAN", "Sistem open-kitchen higienis, staf berseragam rapi, waktu penyajian cepat, dan delivery box penjaga suhu."),
        ("4. UKURAN", "Pilihan porsi fleksibel (Regular, Spesial Campur, Jumbo) serta aneka jenis mie (halus, lebar, kwetiau, bihun)."),
        ("5. PENYAJIAN", "Mangkok keramik oriental estetik, kuah kaldu terpisah, sambal khas, dan kemasan takeaway ramah lingkungan.")
    ]
    for idx, (t, d) in enumerate(diffs):
        x = Inches(0.8 + idx * 2.36)
        card = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(2.25), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s6.shapes.add_textbox(x + Inches(0.15), Inches(2.4), Inches(1.95), Inches(3.8))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 7: SEGMENTASI & TARGET PASAR (STP)
    # -----------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    add_background(s7)
    add_header(s7, "Segmenting, Targeting, Positioning (STP)", "Fokus pada pangsa pasar strategis kawasan Duri Kosambi")

    stp_data = [
        ("Segmentasi Demografis", "Masyarakat keluarga, orang tua murid, guru Kalam Kudus, pelajar, dan pekerja usia 7-65 tahun kelas menengah ke bawah hingga atas."),
        ("Segmentasi Geografis", "Wilayah perumahan Duri Kosambi, Taman Semanan Indah, Cengkareng, Puri Indah, dan sekitarnya di Jakarta Barat."),
        ("Target Pemasaran", "Fokus sarapan dan makan siang keluarga serta civitas sekolah Kalam Kudus dengan daya beli stabil dan repeat order tinggi."),
        ("Positioning", "Sebagai kedai bakmi legendaris terpercaya sejak 1999 yang menyajikan bakmi otentik paling lezat, higienis, dan nyaman.")
    ]
    for idx, (t, d) in enumerate(stp_data):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(2.0 + (idx // 2) * 2.5)
        card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s7.shapes.add_textbox(x + Inches(0.3), y + Inches(0.2), Inches(5.1), Inches(1.8))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 8: BAURAN PEMASARAN 7P (Bagian 1: Product, Price, Place, Promotion)
    # -----------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    add_background(s8)
    add_header(s8, "Bauran Pemasaran (7P) - Bagian 1", "Strategi produk, harga, distribusi, dan promosi agresif")

    p4_data = [
        ("PRODUCT", "Resep keluarga otentik sejak 1999 tanpa bahan pengawet, varian topping babi casiu, babi kecap, dan ayam jamur, serta pelengkap swikiauw homemade."),
        ("PRICE", "Rentang harga terjangkau Rp 15.000 - Rp 45.000 dengan porsi mengenyangkan dan value-for-money yang sangat tinggi."),
        ("PLACE", "Ruko 2 lantai komersial tepat di depan Sekolah Kristen Kalam Kudus Duri Kosambi dengan fasilitas dine-in ber-AC dan parkir nyaman."),
        ("PROMOTION", "Promosi media sosial Instagram/TikTok kuliner, voucher potongan harga pembukaan, kerja sama komunitas, dan layanan GoFood/GrabFood.")
    ]
    for idx, (t, d) in enumerate(p4_data):
        x = Inches(0.8 + idx * 2.95)
        card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(2.8), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s8.shapes.add_textbox(x + Inches(0.2), Inches(2.5), Inches(2.4), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 9: BAURAN PEMASARAN 7P (Bagian 2: People, Process, Physical Evidence)
    # -----------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    add_background(s9)
    add_header(s9, "Bauran Pemasaran (7P) - Bagian 2", "Standarisasi sumber daya manusia, proses produksi, dan bukti fisik")

    p3_data = [
        ("PEOPLE (SDM)", "Tenaga koki ahli peracik resep keluarga, kasir yang teliti, pramusaji ramah berseragam rapi, dan kurir pengantaran sigap."),
        ("PROCESS (PROSES)", "Sistem operasional open-kitchen higienis, peracikan mie berstandar waktu cepat (< 7 menit), dan SOP pengemasan rapi anti-tumpah."),
        ("PHYSICAL EVIDENCE", "Desain interior oriental kontemporer ber-AC, kebersihan meja makan, etalase kaca display transparan, dan kemasan berlogo resmi.")
    ]
    for idx, (t, d) in enumerate(p3_data):
        x = Inches(0.8 + idx * 3.95)
        card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(3.75), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s9.shapes.add_textbox(x + Inches(0.25), Inches(2.5), Inches(3.25), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 10: ANALISIS ASPEK LOKASI
    # -----------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    add_background(s10)
    add_header(s10, "Aspek Lokasi & Keunggulan Aksesibilitas", "Kawasan komersial prima di Jalan Angsoka Hijau IV, Duri Kosambi")

    loc_cards = [
        ("Titik Lokasi Prima", "Jl. Angsoka Hijau IV No. 17 E6, Perumahan Duri Kosambi, Jakarta Barat. Tepat berhadapan langsung dengan gerbang Sekolah Kristen Kalam Kudus."),
        ("Tingkat Kepadatan", "Dikelilingi kawasan perumahan padat keluarga menengah-atas dengan tingkat mobilitas harian tinggi pada jam sarapan dan makan siang."),
        ("Fasilitas Penunjang", "Area parkir motor dan mobil, akses jalan 2 lajur yang mudah dijangkau, dan kedekatan dengan jalur utama Semanan-Cengkareng."),
        ("Konektivitas Online", "Radius jangkauan pengiriman cepat mencakup Semanan, Cengkareng, Puri Kembangan, hingga perbatasan Cipondoh.")
    ]
    for idx, (t, d) in enumerate(loc_cards):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(2.0 + (idx // 2) * 2.5)
        card = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s10.shapes.add_textbox(x + Inches(0.3), y + Inches(0.2), Inches(5.1), Inches(1.8))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 11: STRUKTUR ORGANISASI & MANAJEMEN
    # -----------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    add_background(s11)
    add_header(s11, "Aspek Manajemen & Struktur Organisasi", "Struktur fungsional untuk mengontrol kualitas mutu dan operasional gerai")

    roles_data = [
        ("Manajer Kedai", "1 Orang", "Penanggung jawab operasional harian, laporan keuangan, pengawasan SOP mutu."),
        ("Kepala Dapur (Chef)", "1 Orang", "Pengawasan racikan bumbu otentik, perebusan mie presisi, dan konsistensi rasa."),
        ("Asisten Koki", "1 Orang", "Persiapan topping daging, adonan swikiauw & pangsit homemade, dan plating hidangan."),
        ("Kasir (Cashier)", "1 Orang", "Transaksi tunai & non-tunai (POS/QRIS), pencatatan pesanan GoFood/GrabFood."),
        ("Pramusaji (Server)", "2 Orang", "Penyajian makanan ke meja tamu, kebersihan ruang makan dine-in, hospitality."),
        ("Stock & Delivery", "1 Orang", "Pengelolaan stok bahan baku, kontrol suhu chiller, dan pengantaran lokal.")
    ]
    for idx, (role, count, desc) in enumerate(roles_data):
        x = Inches(0.8 + (idx % 3) * 3.95)
        y = Inches(2.0 + (idx // 3) * 2.5)
        card = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.75), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s11.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), Inches(3.35), Inches(1.9))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = f"{role} ({count})"
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{desc}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 12: AKTIVA TETAP & KEBUTUHAN PERALATAN (CAPEX)
    # -----------------------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    add_background(s12)
    add_header(s12, "Aktiva Tetap (Capex) Peralatan & Renovasi", f"Total investasi aktiva tetap mencapai Rp {met['capex_total']:,.0f}".replace(",", "."))

    capex_samples = [
        ("Renovasi Interior & Partisi Kaca Ruko", "1 Paket", "Rp 28.000.000"),
        ("2 Unit Motor Operasional Delivery", "2 Unit", "Rp 42.000.000"),
        ("Dandang Rebus Mie 4 Lubang Stainless", "2 Unit", "Rp 9.000.000"),
        ("Chiller Undercounter & Chest Freezer 400L", "2 Unit", "Rp 15.500.000"),
        ("Mesin Press / Adonan Mie Komersial", "1 Unit", "Rp 8.500.000"),
        ("Meja Kursi Kayu Resto (10 Meja, 40 Kursi)", "50 Item", "Rp 22.000.000"),
        ("AC Inverter 2 PK & Kipas Exhaust Ducting", "3 Unit", "Rp 20.500.000"),
        ("Mesin Kasir POS, CCTV 8 Cam, TV LED & Signage", "1 Paket", "Rp 16.200.000")
    ]
    for idx, (item, qty, tot) in enumerate(capex_samples):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(2.0 + (idx // 2) * 1.25)
        card = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(1.05))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s12.shapes.add_textbox(x + Inches(0.2), y + Inches(0.1), Inches(5.3), Inches(0.85))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = f"{item} ({qty})"
        p.font.name = "Times New Roman"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p2 = tf.add_paragraph()
        p2.text = f"Nilai Investasi: {tot}"
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(12)
        p2.font.color.rgb = GOLD_ACCENT

    # -----------------------------------------------------------------
    # SLIDE 13: LAYOUT DENAH TATA LETAK GERAI
    # -----------------------------------------------------------------
    s13 = prs.slides.add_slide(blank_layout)
    add_background(s13)
    add_header(s13, "Layout Tata Letak Ruangan Gerai", "Desain fungsional untuk mengoptimalkan kenyamanan dan efisiensi alur kerja")

    layout_rooms = [
        ("Dining Area (Lantai 1)", "Kapasitas 10 meja makan (40 kursi) dengan pendingin udara AC 2 PK, Wi-Fi gratis, dan penerangan nyaman."),
        ("Open-Kitchen Stainless", "Dapur bersih dengan partisi kaca transparan agar konsumen dapat melihat proses perebusan mie yang higienis."),
        ("Cashier Counter & POS", "Meja kasir terintegrasi mesin POS touchscreen, printer struk dapur, dan display menu transparan."),
        ("Waiting Area Ojek Online", "Ruang tunggu khusus driver GoFood/GrabFood agar tidak mengganggu kenyamanan tamu dine-in."),
        ("Chiller & Dry Storage", "Gudang penyimpanan mie segar harian, chiller daging, dan stok bumbu terpisah bebas hama."),
        ("Staff Room & Fasilitas", "Ruang istirahat karyawan di lantai 2, toilet bersih pengunjung, dan area parkir motor di depan ruko.")
    ]
    for idx, (rm, desc) in enumerate(layout_rooms):
        x = Inches(0.8 + (idx % 3) * 3.95)
        y = Inches(2.0 + (idx // 3) * 2.5)
        card = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.75), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s13.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), Inches(3.35), Inches(1.9))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = rm
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{desc}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 14: NETWORK PLANNING (A-L) & GANTT CHART
    # -----------------------------------------------------------------
    s14 = prs.slides.add_slide(blank_layout)
    add_background(s14)
    add_header(s14, "Network Planning (A s/d L) & Timeline", "Jadwal tahapan pra-operasional selama 3 bulan intensif")

    steps_summary = [
        ("Bulan 1: Perencanaan & Legalitas", "A. Penyusunan proposal SKB\nB. Survei ruko Duri Kosambi\nC. Studi kelayakan lingkungan\nD. Penandatanganan akad sewa"),
        ("Bulan 2: Renovasi & Pengadaan", "E. Desain arsitektur interior\nF. Renovasi fisik & instalasi dapur\nG. Pengadaan mesin rebus & chiller\nH. Setting meja kursi & tata letak"),
        ("Bulan 3: Persiapan & Launching", "I. Uji coba fungsi dapur & sanitasi\nJ. Perekrutan 7 staf operasional\nK. Pelatihan SOP resep & kasir\nL. Grand Opening & promo diskon")
    ]
    for idx, (m_title, m_desc) in enumerate(steps_summary):
        x = Inches(0.8 + idx * 3.95)
        card = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(3.75), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s14.shapes.add_textbox(x + Inches(0.25), Inches(2.5), Inches(3.25), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = m_title
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{m_desc}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 15: RINCIAN INITIAL OUTLAY (MODAL AWAL)
    # -----------------------------------------------------------------
    s15 = prs.slides.add_slide(blank_layout)
    add_background(s15)
    add_header(s15, "Rincian Investasi Awal (Initial Outlay)", f"Total modal awal yang dibutuhkan: Rp {met['initial_outlay']:,.0f}".replace(",", "."))

    outlays = [
        ("Aktiva Tetap (Capex)", f"Rp {met['capex_total']:,.0f}".replace(",", "."), "Pengadaan peralatan dapur rebus mie, chiller, motor delivery, meja kursi, renovasi ruko."),
        ("Biaya Sewa Ruko", f"Rp {cfg['operational_costs']['rent_per_year']:,.0f}".replace(",", "."), "Alokasi sewa ruko 2 lantai komersial selama 1 tahun pertama."),
        ("Biaya AMDAL & Sanitasi", f"Rp {cfg['operational_costs']['amdal_initial']:,.0f}".replace(",", "."), "Penyediaan fasilitas tempat sampah pilah, grease trap lemak, dan sanitasi."),
        ("Biaya Survei & Perizinan", f"Rp {cfg['operational_costs']['survey_initial']:,.0f}".replace(",", "."), "Biaya transportasi, uji pasar lingkungan, dan perizinan kelayakan usaha."),
        ("Biaya Promosi Awal", f"Rp {cfg['operational_costs']['promosi_initial']:,.0f}".replace(",", "."), "Pemasangan neon box, spanduk grand opening, iklan media sosial, dan voucher promo.")
    ]
    for idx, (item, val, desc) in enumerate(outlays):
        y = Inches(2.0 + idx * 0.95)
        card = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(11.7), Inches(0.85))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s15.shapes.add_textbox(Inches(1.0), y + Inches(0.1), Inches(11.3), Inches(0.65))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = f"{item} : {val}"
        p.font.name = "Times New Roman"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = "Times New Roman"
        p2.font.size = Pt(11)
        p2.font.color.rgb = GRAY_TEXT

    # -----------------------------------------------------------------
    # SLIDE 16: PROYEKSI CASH INFLOW 5 TAHUN
    # -----------------------------------------------------------------
    s16 = prs.slides.add_slide(blank_layout)
    add_background(s16)
    add_header(s16, "Proyeksi Penerimaan Kas (Cash Inflow 5 Tahun)", "Asumsi 25 hari kerja/bulan (300 hari operasi/tahun) dengan pertumbuhan penjualan")

    for idx, cif in enumerate(met['cif_years'], start=1):
        x = Inches(0.8 + (idx - 1) * 2.36)
        card = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(2.25), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s16.shapes.add_textbox(x + Inches(0.15), Inches(2.5), Inches(1.95), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = f"TAHUN {idx}"
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_v = tf.add_paragraph()
        p_v.text = f"\nRp {cif:,.0f}".replace(",", ".")
        p_v.font.name = "Times New Roman"
        p_v.font.size = Pt(15)
        p_v.font.bold = True
        p_v.font.color.rgb = WHITE

        p_d = tf.add_paragraph()
        p_d.text = f"\nProyeksi penjualan bakmi dan minuman segar dengan volume penjualan harian terstandarisasi."
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = GRAY_TEXT

    # -----------------------------------------------------------------
    # SLIDE 17: PROYEKSI CASH OUTFLOW 5 TAHUN
    # -----------------------------------------------------------------
    s17 = prs.slides.add_slide(blank_layout)
    add_background(s17)
    add_header(s17, "Proyeksi Pengeluaran Kas (Cash Outflow 5 Tahun)", "Beban gaji karyawan, bahan baku daging & mie segar, serta utilitas komersial")

    for idx, cof in enumerate(met['cof_years'], start=1):
        x = Inches(0.8 + (idx - 1) * 2.36)
        card = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.2), Inches(2.25), Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s17.shapes.add_textbox(x + Inches(0.15), Inches(2.5), Inches(1.95), Inches(3.7))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = f"TAHUN {idx}"
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_v = tf.add_paragraph()
        p_v.text = f"\nRp {cof:,.0f}".replace(",", ".")
        p_v.font.name = "Times New Roman"
        p_v.font.size = Pt(15)
        p_v.font.bold = True
        p_v.font.color.rgb = WHITE

        p_d = tf.add_paragraph()
        p_d.text = f"\nMencakup gaji 7 staf, bahan baku daging babi/ayam/mie, listrik PLN B-1, gas LPG, dan air PAM."
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = GRAY_TEXT

    # -----------------------------------------------------------------
    # SLIDE 18: TABEL NCF & PROCEED 5 TAHUN (Tabel 5.1)
    # -----------------------------------------------------------------
    s18 = prs.slides.add_slide(blank_layout)
    add_background(s18)
    add_header(s18, "Tabel Arus Kas Bersih (NCF) & Proceed", "Penyusutan aktiva tetap metode garis lurus sebesar 20% per tahun")

    table_shape18 = s18.shapes.add_table(7, 6, Inches(0.8), Inches(2.0), Inches(11.7), Inches(4.6))
    t18 = table_shape18.table
    hdrs18 = ["Komponen Arus Kas", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]
    for c_i, h in enumerate(hdrs18):
        c = t18.cell(0, c_i)
        c.fill.solid()
        c.fill.fore_color.rgb = GOLD_ACCENT
        p = c.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Times New Roman"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = DARK_BG

    ncf_table_data = [
        ["Cash Inflow (CIF)"] + [f"Rp {x:,.0f}".replace(",", ".") for x in met['cif_years']],
        ["Cash Outflow (COF)"] + [f"Rp {x:,.0f}".replace(",", ".") for x in met['cof_years']],
        ["NCF = EAT (CIF - COF)"] + [f"Rp {x:,.0f}".replace(",", ".") for x in met['ncf_years']],
        ["Depresiasi Garis Lurus (20%)"] + [f"Rp {met['depresiasi_per_year']:,.0f}".replace(",", ".")] * 5,
        ["Proceed (NCF - Depresiasi)"] + [f"Rp {x:,.0f}".replace(",", ".") for x in met['proceed_years']],
        ["Present Value PV (20%)"] + [f"Rp {x:,.0f}".replace(",", ".") for x in met['pv_years']]
    ]
    for r_i, row in enumerate(ncf_table_data, start=1):
        for c_i, val in enumerate(row):
            cell = t18.cell(r_i, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Times New Roman"
            p.font.size = Pt(11)
            p.font.color.rgb = WHITE
            if c_i > 0:
                p.alignment = PP_ALIGN.RIGHT

    # -----------------------------------------------------------------
    # SLIDE 19: HASIL UJI KRITERIA KELAYAKAN FINANSIAL
    # -----------------------------------------------------------------
    s19 = prs.slides.add_slide(blank_layout)
    add_background(s19)
    add_header(s19, "Hasil Evaluasi Kelayakan Finansial", "Seluruh indikator melampaui standar kelayakan investasi FEB")

    metrics_cards = [
        ("NET PRESENT VALUE (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "Standar: > 0 (Positif)", "STATUS: SANGAT LAYAK"),
        ("INTERNAL RATE OF RETURN (IRR)", f"{met['irr']*100:.2f}%", "Standar: > 20.00% (Opportunity Cost)", "STATUS: SANGAT LAYAK"),
        ("PAYBACK PERIOD (PP)", f"{met['pp_months']:.2f} Bulan ({met['pp_years']:.2f} Thn)", "Standar: < 3.0 Tahun (36 Bulan)", "STATUS: SANGAT LAYAK"),
        ("PROFITABILITY INDEX (PI)", f"{met['pi']:.2f}", "Standar: > 1.20", "STATUS: SANGAT LAYAK")
    ]
    for idx, (title, val, std, stat) in enumerate(metrics_cards):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(2.0 + (idx // 2) * 2.5)
        card = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GREEN_ACCENT
        card.line.width = Pt(1.5)

        tb = s19.shapes.add_textbox(x + Inches(0.3), y + Inches(0.15), Inches(5.1), Inches(1.9))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_v = tf.add_paragraph()
        p_v.text = val
        p_v.font.name = "Times New Roman"
        p_v.font.size = Pt(22)
        p_v.font.bold = True
        p_v.font.color.rgb = WHITE

        p_s = tf.add_paragraph()
        p_s.text = f"{std}\n{stat}"
        p_s.font.name = "Times New Roman"
        p_s.font.size = Pt(12)
        p_s.font.bold = True
        p_s.font.color.rgb = GREEN_ACCENT

    # -----------------------------------------------------------------
    # SLIDE 20: TABEL RINGKASAN KELAYAKAN (Tabel 6.1)
    # -----------------------------------------------------------------
    s20 = prs.slides.add_slide(blank_layout)
    add_background(s20)
    add_header(s20, "Tabel 6.1 Kesimpulan Kelayakan Investasi", "Rekapitulasi kualifikasi kelayakan finansial Bakmi Mimu Carina Sayang")

    t_s20 = s20.shapes.add_table(5, 4, Inches(0.8), Inches(2.0), Inches(11.7), Inches(4.3)).table
    hdrs20 = ["Kriteria Evaluasi", "Hasil Perhitungan", "Standar Kelayakan", "Status Kelayakan"]
    for c_i, h in enumerate(hdrs20):
        cell = t_s20.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = GOLD_ACCENT
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Times New Roman"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = DARK_BG
        p.alignment = PP_ALIGN.CENTER

    t20_rows = [
        ("Net Present Value (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "> 0 (Bernilai Positif)", "LAYAK"),
        ("Internal Rate of Return (IRR)", f"{met['irr']*100:.2f}%", "> 20.00% (Opportunity Cost)", "LAYAK"),
        ("Payback Period (PP)", f"{met['pp_months']:.2f} Bulan (< 1 Tahun)", "< 3.0 Tahun (36 Bulan)", "LAYAK"),
        ("Profitability Index (PI)", f"{met['pi']:.2f}", "> 1.20", "LAYAK")
    ]
    for r_i, r_vals in enumerate(t20_rows, start=1):
        for c_i, val in enumerate(r_vals):
            cell = t_s20.cell(r_i, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Times New Roman"
            p.font.size = Pt(12)
            p.font.bold = (c_i == 3)
            p.font.color.rgb = GREEN_ACCENT if c_i == 3 else WHITE
            p.alignment = PP_ALIGN.CENTER

    # -----------------------------------------------------------------
    # SLIDE 21: ANALISIS RISIKO & MITIGASI
    # -----------------------------------------------------------------
    s21 = prs.slides.add_slide(blank_layout)
    add_background(s21)
    add_header(s21, "Analisis Risiko & Strategi Mitigasi", "Langkah antisipatif menjamin keberlangsungan operasional gerai")

    risks = [
        ("Fluktuasi Harga Daging", "Mitigasi: Mengikat kontrak kemitraan dengan supplier pasar induk grosir dan menjaga stok beku di chest freezer 400L."),
        ("Persaingan Kedai Bakmi", "Mitigasi: Mempertahankan keaslian resep turun-temurun sejak 1999, higienitas dapur terbuka, dan konsistensi rasa mie kenyal berkilau."),
        ("Kenaikan Biaya Sewa Ruko", "Mitigasi: Menandatangani kontrak sewa jangka panjang (5 tahun) dengan nilai tetap guna menghindari lonjakan biaya tempat."),
        ("Ketergantungan Koki", "Mitigasi: Standardisasi resep bumbu minyak dalam takaran baku (SOP terpusat) sehingga proses peracikan dapat diduplikasi asisten koki.")
    ]
    for idx, (t, d) in enumerate(risks):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(2.0 + (idx // 2) * 2.5)
        card = s21.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = GOLD_ACCENT

        tb = s21.shapes.add_textbox(x + Inches(0.3), y + Inches(0.2), Inches(5.1), Inches(1.8))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        p_t.text = t
        p_t.font.name = "Times New Roman"
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = GOLD_ACCENT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{d}"
        p_d.font.name = "Times New Roman"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = WHITE

    # -----------------------------------------------------------------
    # SLIDE 22: KESIMPULAN AKHIR & PENUTUP
    # -----------------------------------------------------------------
    s22 = prs.slides.add_slide(blank_layout)
    add_background(s22)

    dec22 = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1))
    dec22.fill.solid()
    dec22.fill.fore_color.rgb = CARD_BG
    dec22.line.color.rgb = GREEN_ACCENT
    dec22.line.width = Pt(2.0)

    tb22 = s22.shapes.add_textbox(Inches(1.8), Inches(1.6), Inches(9.7), Inches(4.3))
    tf22 = tb22.text_frame
    p1 = tf22.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = "KESIMPULAN AKHIR KELAYAKAN BISNIS"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(20)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT

    p2 = tf22.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = "“ BISNIS DINYATAKAN SANGAT LAYAK ”"
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(32)
    p2.font.bold = True
    p2.font.color.rgb = GREEN_ACCENT

    p3 = tf22.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.text = f"\nBerdasarkan kajian menyeluruh aspek pasar, manajemen, teknis operasi, serta finansial, usaha Bakmi Mimu Carina Sayang membuktikan prospek laba yang sangat solid dengan pengembalian investasi hanya dalam tempo {met['pp_months']:.2f} bulan."
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(14)
    p3.font.color.rgb = WHITE

    p4 = tf22.add_paragraph()
    p4.alignment = PP_ALIGN.CENTER
    p4.text = "\nSekian & Terima Kasih\nFakultas Ekonomi & Bisnis | Universitas Kristen Krida Wacana"
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(13)
    p4.font.italic = True
    p4.font.color.rgb = GRAY_TEXT

    out_pptx = r"d:\Perkuliahan\Kelass\SKB\Bakmi_Mimu_Presentasi_SKB.pptx"
    prs.save(out_pptx)
    print(f"[SUCCESS] Presentasi PowerPoint berhasil dibuat di: {out_pptx}")

if __name__ == "__main__":
    create_presentation()
