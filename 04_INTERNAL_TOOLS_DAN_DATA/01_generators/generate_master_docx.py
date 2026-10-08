import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import os

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn(f'w:{key}'), str(edge_data[key]))

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, line_spacing=1.5, space_after=Pt(6), space_before=Pt(0)):
    p.alignment = align
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = space_after
    p.paragraph_format.space_before = space_before

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    if level == 1:
        apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.5, space_after=Pt(12), space_before=Pt(18))
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
    elif level == 2:
        apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.LEFT, line_spacing=1.5, space_after=Pt(6), space_before=Pt(12))
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
    elif level == 3:
        apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.LEFT, line_spacing=1.5, space_after=Pt(4), space_before=Pt(6))
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
    return p

def add_p(doc, text, bold_prefix="", indent=0):
    p = doc.add_paragraph()
    apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, line_spacing=1.5, space_after=Pt(6))
    if indent > 0:
        p.paragraph_format.left_indent = Inches(indent * 0.25)
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.font.name = "Times New Roman"
        rb.font.size = Pt(12)
        rb.bold = True
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    return p

def add_tbl_title(doc, text):
    p = doc.add_paragraph()
    apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.15, space_after=Pt(4), space_before=Pt(10))
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.bold = True
    return p

def add_fig_cap(doc, text):
    p = doc.add_paragraph()
    apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.15, space_after=Pt(10), space_before=Pt(4))
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.bold = True
    return p

def add_fig_image(doc, img_path, width_in=5.8):
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        apply_p_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.0, space_after=Pt(4), space_before=Pt(8))
        r = p.add_run()
        r.add_picture(img_path, width=Inches(width_in))
        return p
    return None

def style_table(t, col_widths=None):
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    border_kwargs = {
        'top': {'sz': 4, 'val': 'single', 'color': '808080'},
        'bottom': {'sz': 4, 'val': 'single', 'color': '808080'},
        'left': {'sz': 4, 'val': 'single', 'color': '808080'},
        'right': {'sz': 4, 'val': 'single', 'color': '808080'},
        'insideH': {'sz': 4, 'val': 'single', 'color': 'D0D0D0'},
        'insideV': {'sz': 4, 'val': 'single', 'color': 'D0D0D0'},
    }
    for row_idx, row in enumerate(t.rows):
        is_header = (row_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            set_cell_border(cell, **border_kwargs)
            if is_header:
                set_cell_shading(cell, "D9E1F2")
            if col_widths and c_idx < len(col_widths):
                cell.width = col_widths[c_idx]
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)
                txt = p.text.strip()
                if not is_header and (txt.startswith("Rp") or txt.startswith("-Rp") or (txt.replace(".", "").replace(",", "").replace("%", "").isdigit() and len(txt) > 2)):
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                elif is_header or (txt.isdigit() and len(txt) <= 2):
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(10)
                    if is_header:
                        r.bold = True

def build_master_docx():
    base = r"d:\Perkuliahan\Kelass\SKB"
    data_dir = os.path.join(base, "04_INTERNAL_TOOLS_DAN_DATA", "03_data_dan_aset")
    cfg_path = os.path.join(data_dir, "config_bakmi_mimu.json")
    met_path = os.path.join(data_dir, "metrics_bakmi_mimu.json")
    out_docx = os.path.join(base, "01_TUGAS_FINAL_BAKMI_MIMU", "Laporan_SKB_Bakmi_Mimu_Carina_Sayang.docx")
    logo_path = os.path.join(data_dir, "logo_ukrida.png")

    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    with open(met_path, 'r', encoding='utf-8') as f:
        met = json.load(f)

    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)

    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(12)

    # -----------------------------------------------------------------
    # COVER PAGE
    # -----------------------------------------------------------------
    p_c1 = doc.add_paragraph()
    apply_p_format(p_c1, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=Pt(20), space_after=Pt(4))
    r = p_c1.add_run("Studi Kelayakan Bisnis")
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)
    r.bold = True

    p_c2 = doc.add_paragraph()
    apply_p_format(p_c2, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(4))
    r = p_c2.add_run(f"“ {cfg['business_name'].upper()} ”")
    r.font.name = "Times New Roman"
    r.font.size = Pt(18)
    r.bold = True

    p_c2b = doc.add_paragraph()
    apply_p_format(p_c2b, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(16))
    r = p_c2b.add_run(f"- {cfg['tagline']} -")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.italic = True

    # Logo UKRIDA
    p_logo = doc.add_paragraph()
    apply_p_format(p_logo, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(20))
    if os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Inches(2.0))

    # Authors / Team Members
    p_c3 = doc.add_paragraph()
    apply_p_format(p_c3, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(4))
    r = p_c3.add_run("Disusun Oleh (Kelompok):")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.bold = True

    members = [
        "Arthur Reezan (312023002)",
        "Jennese Putra Alamsyah Sukadi (312023033)",
        "Valendrik Dwiputra Wirawan (312023013)",
        "Affandy (312023075)",
        "Steven Putra Tjhin (312023015)"
    ]
    p_m = doc.add_paragraph()
    apply_p_format(p_m, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(30))
    for m in members:
        rm = p_m.add_run(f"{m}\n")
        rm.font.name = "Times New Roman"
        rm.font.size = Pt(11)

    # Institution
    p_c4 = doc.add_paragraph()
    apply_p_format(p_c4, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=Pt(20), space_after=Pt(4))
    r = p_c4.add_run("Program Studi Manajemen\nFakultas Ekonomi & Bisnis\nUniversitas Kristen Krida Wacana\n2023/2024")
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.bold = True

    doc.add_page_break()

    # -----------------------------------------------------------------
    # BAB I: PENDAHULUAN
    # -----------------------------------------------------------------
    add_heading(doc, "BAB 1\nPENDAHULUAN", level=1)
    add_p(doc, "Latar Belakang")
    add_p(doc, cfg['background']['history'])
    add_p(doc, cfg['background']['opportunity'])
    add_p(doc, cfg['background']['challenges'])
    add_p(doc, cfg['background']['concept'])

    add_p(doc, f"Berikut merupakan rincian produk dan daftar menu lengkap {cfg['business_name']} berdasarkan daftar menu resmi kedai:")
    add_tbl_title(doc, f"Tabel 1.1. Jenis Produk dan Daftar Menu {cfg['business_name']}")

    t1 = doc.add_table(rows=1, cols=3)
    t1.rows[0].cells[0].text = "No."
    t1.rows[0].cells[1].text = f"Kategori & Varian Menu {cfg['business_name']}"
    t1.rows[0].cells[2].text = "Harga Jual"
    
    item_num = 1
    menu_catalog = cfg.get("menu_catalog", [])
    if menu_catalog:
        for cat in menu_catalog:
            cat_row = t1.add_row()
            cat_row.cells[0].text = "-"
            cat_row.cells[1].text = f"[{cat['category']}]"
            cat_row.cells[2].text = ""
            set_cell_shading(cat_row.cells[0], "F2F4F8")
            set_cell_shading(cat_row.cells[1], "F2F4F8")
            set_cell_shading(cat_row.cells[2], "F2F4F8")
            for it in cat.get("items", []):
                row = t1.add_row()
                row.cells[0].text = str(item_num)
                row.cells[1].text = it["name"]
                row.cells[2].text = f"Rp {it['price']:,}".replace(",", ".")
                item_num += 1
    else:
        for i, prod in enumerate(cfg.get("products", []), start=1):
            row = t1.add_row()
            row.cells[0].text = str(i)
            row.cells[1].text = prod["name"]
            row.cells[2].text = f"Rp {prod['price_year_1']:,}".replace(",", ".")

    style_table(t1, [Inches(0.6), Inches(4.2), Inches(1.8)])
    
    # Bold the category header rows
    for r in t1.rows:
        if r.cells[1].text.startswith("["):
            for p in r.cells[1].paragraphs:
                for run in p.runs:
                    run.bold = True
                    run.font.name = "Times New Roman"

    add_p(doc, f"Catatan Informasi Operasional: Kedai beroperasi setiap hari kerja pukul {cfg['marketing'].get('opening_hours', '07.30 - 13.30 WIB')} melayani santap di tempat (Dine-In), bawa pulang (Takeaway), pemesanan WhatsApp ({cfg['marketing'].get('phone_wa', '0859 3984 7536')}), serta terintegrasi penuh pada platform digital {cfg['marketing'].get('delivery_services', 'GoFood & GrabFood')}.")
    add_p(doc, f"\nBerikut merupakan gambaran visual dari produk {cfg['business_name']}:")
    add_fig_image(doc, os.path.join(data_dir, "produk_bakmi_mimu.png"), width_in=5.8)
    add_fig_cap(doc, f"Gambar 1.1 Gambaran Produk {cfg['business_name']}")

    # -----------------------------------------------------------------
    # BAB II: ASPEK PEMASARAN DAN PASAR
    # -----------------------------------------------------------------
    doc.add_page_break()
    add_heading(doc, "BAB II\nASPEK PEMASARAN DAN PASAR", level=1)
    add_heading(doc, "2.1 Segmentasi Pasar", level=2)
    add_p(doc, f"Aspek pemasaran merupakan pilar fundamental yang perlu diperhatikan dalam ekspansi usaha kuliner {cfg['business_name']}. Dengan memahami karakteristik pasar secara mendalam, manajemen dapat memetakan potensi konsumen, merancang saluran promosi yang efisien, dan mencapai target volume penjualan yang diharapkan. Segmentasi pasar membagi pasar heterogen ke dalam kelompok konsumen homogen yang memiliki kebutuhan dan preferensi serupa.")
    add_p(doc, f"Secara demografis, segmen pasar dari {cfg['business_name']} yaitu {cfg['marketing']['segmentation']['demographic']}")
    add_p(doc, f"Secara geografis, segmen pasar utama berpusat pada {cfg['marketing']['segmentation']['geographic']}")
    add_p(doc, f"Secara perilaku dan psikografis, target pasar adalah {cfg['marketing']['segmentation']['behavioral']}")

    add_heading(doc, "2.2 Diferensiasi Usaha", level=2)
    add_p(doc, "Untuk memenangkan persaingan di industri kuliner bakmi Jakarta Barat, usaha ini memiliki 5 pilar diferensiasi unggulan:")
    add_p(doc, cfg['marketing']['differentiation']['taste'], bold_prefix="- Rasa: ")
    add_p(doc, cfg['marketing']['differentiation']['location'], bold_prefix="- Lokasi: ")
    add_p(doc, cfg['marketing']['differentiation']['service'], bold_prefix="- Layanan: ")
    add_p(doc, cfg['marketing']['differentiation']['size'], bold_prefix="- Ukuran: ")
    add_p(doc, cfg['marketing']['differentiation']['presentation'], bold_prefix="- Penyajian: ")

    add_heading(doc, "2.3 Aspek Lokasi", level=2)
    add_p(doc, f"{cfg['business_name']} berlokasi di titik yang sangat strategis: {cfg['marketing']['location']}. Lokasi ini berada persis di depan gerbang utama institusi pendidikan dan perumahan padat penduduk yang menjamin tingginya lalu lintas pengunjung alami (foot traffic), akses parkir yang memadai, serta visibilitas gerai yang sangat optimal.")
    add_fig_image(doc, os.path.join(data_dir, "peta_lokasi_bakmi_mimu.png"), width_in=5.8)
    add_fig_cap(doc, f"Gambar 2.1 Peta Lokasi Kedai {cfg['business_name']}")

    add_heading(doc, "2.4 Target Pemasaran", level=2)
    add_p(doc, f"Target pemasaran utama dari {cfg['business_name']} difokuskan kepada para orang tua murid dan guru sekolah Kalam Kudus saat jam antar-jemput, warga perumahan Duri Kosambi dan Semanan yang mencari sarapan dan makan siang keluarga, serta komunitas pecinta bakmi otentik di area Jakarta Barat dengan kisaran harga yang kompetitif ({cfg['marketing']['price_range']}).")

    add_heading(doc, "2.5 Bauran Pemasaran (7P)", level=2)
    add_p(doc, "Bauran pemasaran 7P diterapkan secara komprehensif sebagai berikut:")
    p7 = [
        ("Product: ", "Menyajikan bakmi bertekstur kenyal berkilau alami buatan sendiri setiap minggu tanpa bahan pengawet dengan pilihan racikan 3 jenis minyak mie (minyak babi, minyak ayam, dan minyak sayur campur wijen), serta 3 varian topping daging murni: ayam putih gurih, babi cincang kecap, dan babi casiu madu panggang merah (tanpa olahan daging bebek maupun jamur). Dilengkapi menu pelengkap homemade swikiaw rebus, pangsit rebus, baso sapi kuah, baso ikan kuah, pangsit goreng, dan baso goreng babi-udang, serta saus botolan standar meja (Cap Belibis dan Mangga Besar)."),
        ("Price: ", f"Penetapan harga berbasis nilai (value-based pricing) yang sangat bersaing: range harga bakmi Rp 24.000 s.d. Rp 59.000 per porsi (tersedia porsi kecil Rp 27.000, porsi standar Rp 29.000 - Rp 39.000, dan porsi jumbo +100% mie Rp 49.000 - Rp 59.000). Menu pelengkap kuah Rp 22.500 s.d. Rp 27.500, gorengan Rp 5.000 - Rp 8.000, serta aneka minuman tradisional berkisar Rp 1.000 s.d. Rp 15.000 (tanpa minuman berbahan jeruk)."),
        ("Place: ", "Lokasi kedai berupa ruko satu lantai yang strategis dan bersih di Jl. Angsoka Hijau IV Blok E6 No. 17 Duri Kosambi, persis di depan Taman TK, SD, SMP Kalam Kudus dengan akses parkir motor yang memadai di halaman depan serta terdaftar pada layanan pesan-antar online."),
        ("Promotion: ", "Kedai tidak mengadakan program diskon atau kartu loyalitas formal; strategi pemasaran murni mengandalkan kekuatan penjualan lokal melalui kepuasan pelanggan setia dan rekomendasi getok tular (word-of-mouth), serta sesekali mendapatkan ulasan organik sukarela dari food reviewer di platform media sosial seperti TikTok dan Instagram."),
        ("People: ", "Dikelola oleh tim kerja yang solid terdiri dari 3 orang karyawan tetap: Koki Utama (pengolah adonan dan peracik bumbu), Asisten Koki (persiapan swikiaw, pangsit, dan kaldu), serta Kasir dan Pramusaji. Seluruh staf menerima gaji pokok tetap bulanan (total alokasi sekitar Rp 10.000.000 per bulan) tanpa skema bonus insentif penjualan."),
        ("Process: ", "Mengusung sistem open kitchen gerobak di teras depan ruko sehingga proses perebusan dan peracikan mie berlangsung higienis dan transparan di hadapan pelanggan. Durasi penyajian terstandarisasi di bawah 5-7 menit per porsi. Sistem pencatatan pesanan masih menggunakan metode manual dengan nota kertas fisik."),
        ("Physical Evidence: ", "Desain interior standar ruko satu lantai yang sederhana, bersih, dan fungsional. Area bersantap berkapasitas maksimal 22 orang (padat) dilengkapi 6 unit kipas angin dinding/plafon yang menjamin sirkulasi udara bebas gerah (tanpa AC). Kedai sengaja tidak menyediakan fasilitas Wi-Fi agar pengunjung fokus bersantap dan menjaga tingkat perputaran meja (table turnover) yang sehat.")
    ]
    for b_pre, b_txt in p7:
        add_p(doc, b_txt, bold_prefix=b_pre, indent=1)

    # -----------------------------------------------------------------
    # BAB III: ASPEK MANAJEMEN
    # -----------------------------------------------------------------
    doc.add_page_break()
    add_heading(doc, "BAB III\nASPEK MANAJEMEN", level=1)
    add_p(doc, "Struktur Organisasi")
    add_fig_image(doc, os.path.join(data_dir, "struktur_organisasi_bakmi_mimu.png"), width_in=5.8)
    add_fig_cap(doc, f"Gambar 3.1 Struktur Organisasi {cfg['business_name']}")
    add_p(doc, f"Struktur organisasi {cfg['business_name']} dirancang ramping, fungsional, dan efektif dengan total 3 orang tenaga kerja inti. Total alokasi anggaran gaji karyawan ditetapkan pada kisaran sekitar Rp 10.000.000 per bulan (Rp 120.000.000 per tahun) tanpa sistem insentif penjualan. Berikut rincian uraian pekerjaan (Job Description) untuk setiap posisi:")
    for role in cfg['management']['organization_roles']:
        add_p(doc, role['job_desc'], bold_prefix=f"{role['title']}:\n", indent=0)

    # -----------------------------------------------------------------
    # BAB IV: ASPEK TEKNIS OPERASI
    # -----------------------------------------------------------------
    doc.add_page_break()
    add_heading(doc, "BAB IV\nASPEK TEKNIS OPERASI", level=1)
    add_heading(doc, "4.1 Aktiva Tetap & Bahan Baku", level=2)
    add_heading(doc, "4.1.1 Gedung dan Bangunan", level=3)
    add_p(doc, f"{cfg['business_name']} memilih sistem sewa tempat usaha: {cfg['operations']['facility_status']}. Bangunan ruko direnovasi untuk mengakomodasi seluruh fungsi operasional gerai, mulai dari ruang makan utama, dapur bersih, gudang penyimpanan dingin, kasir, hingga area tunggu pengemudi ojek online.")

    add_heading(doc, "4.1.2 Peralatan dan Perlengkapan", level=3)
    add_p(doc, f"Peralatan dan perlengkapan merupakan aktiva tetap (capex) operasional yang digunakan langsung untuk kegiatan produksi dan pelayanan {cfg['business_name']}. Berikut rincian aktiva tetap yang dibutuhkan secara lengkap:")

    add_tbl_title(doc, f"Tabel 4.1 Peralatan dan Perlengkapan {cfg['business_name']}")
    t41 = doc.add_table(rows=1, cols=5)
    for idx, h in enumerate(["No.", "Keterangan", "Jumlah", "Harga Satuan", "Total"]):
        t41.rows[0].cells[idx].text = h

    tot_capex_calc = 0
    for idx, item in enumerate(cfg.get("capex_items", []), start=1):
        tot_val = item["qty"] * item["unit_price"]
        tot_capex_calc += tot_val
        row = t41.add_row()
        row.cells[0].text = str(idx)
        row.cells[1].text = item["name"]
        row.cells[2].text = str(item["qty"])
        row.cells[3].text = f"Rp {item['unit_price']:,}".replace(",", ".")
        row.cells[4].text = f"Rp {tot_val:,}".replace(",", ".")

    r_tot = t41.add_row()
    r_tot.cells[1].text = "Total Biaya Aktiva Tetap"
    r_tot.cells[4].text = f"Rp {tot_capex_calc:,}".replace(",", ".")
    style_table(t41, [Inches(0.5), Inches(2.6), Inches(0.8), Inches(1.3), Inches(1.4)])

    add_heading(doc, "4.2 Layout Usaha", level=2)
    layout_path = os.path.join(data_dir, "denah_tata_letak_ruko.png")
    add_fig_image(doc, layout_path, width_in=5.8)
    add_fig_cap(doc, f"Gambar 4.1 Denah Tata Letak Kedai Ruko 1 Lantai {cfg['business_name']}")
    add_p(doc, f"Tata letak (layout) ruangan {cfg['business_name']} dirancang dengan prinsip ergonomi alur kerja (workflow ergonomics) dan efisiensi ruang ruko satu lantai berdimensi memanjang. Penataan ruang menerapkan sistem arus satu arah (straight-line flow) dari akses masuk depan menuju ruang santap utama dan area servis belakang, guna meminimalisasi tabrakan arus antara pengunjung dan staf operasional:")
    add_p(doc, "1. Area Depan & Teras (Dapur Semi-Terbuka / Open Kitchen & Dine-In Luar): Terletak di bagian depan sebelum rolling door sebagai titik kontak pertama pelanggan. Dilengkapi etalase kaca gerobak utama untuk meracik mie, stasiun kompor gas perebusan bertemperatur tinggi, meja kuah kaldu, dispenser tempat minum, kulkas penyimpanan sayuran dan bahan segar harian, serta meja persiapan sayur dan pangsit. Pada sisi seberang teras disediakan 1 meja makan dengan 4 kursi santap luar (juga difungsikan sebagai waiting area pesanan take-away), 1 unit freezer pembeku stok daging, serta rak perlengkapan sanitasi dan kipas angin teras.")
    add_p(doc, "2. Rolling Door (Partisi Fleksibel Ruang): Berfungsi sebagai sekat pembatas fungsional antara area dapur depan dengan ruang makan dalam. Rolling door dibuka penuh pada jam sibuk dan dapat ditutup sebagian saat malam hari untuk menjaga kestabilan sirkulasi udara serta mencegah debu jalan masuk ke ruang makan utama.")
    add_p(doc, "3. Area Makan Utama (Indoor Dining Room - 19 Kursi Santap): Ruang makan utama yang nyaman dan terlindung, dilengkapi 3 unit meja makan reguler berkapasitas 4 orang per meja (total 12 kursi) di sisi atas serta 1 unit meja makan dinding memanjang dengan 7 kursi sejajar di sisi bawah yang sangat ramah bagi pelanggan individu (solo diner). Sirkulasi udara ditunjang 6 unit kipas angin dinding untuk menjaga kenyamanan pengunjung tanpa membebani biaya listrik pendingin AC berlebihan.")
    add_p(doc, "4. Area Servis & Fasilitas Sanitasi Belakang: Terletak di ujung belakang ruko secara terpisah demi menjamin higienitas pangan. Terdiri dari meja barang untuk penataan inventaris alat makan bersih dan kemasan take-away, wastafel cuci tangan higienis bagi pelanggan, serta 1 unit kamar mandi/toilet bersih tertutup.")
    add_p(doc, f"Total kapasitas tempat duduk pengunjung di seluruh area kedai {cfg['business_name']} mencapai 23 kursi aktif (4 kursi di teras depan dan 19 kursi di ruang makan dalam), menjamin pemenuhan target perputaran meja (table turnover rate) harian secara optimal.")

    add_heading(doc, "4.3 Network Planning", level=2)
    add_fig_image(doc, os.path.join(data_dir, "network_planning_bakmi_mimu.png"), width_in=6.0)
    add_fig_cap(doc, f"Gambar 4.2 Network Planning {cfg['business_name']}")
    add_p(doc, "Keterangan urutan aktivitas pra-operasional (Critical Path Activity A s/d L):")
    steps = [
        "A → Perencanaan Usaha (Penyusunan Proposal SKB Komprehensif)",
        "B → Pencarian & Seleksi Lokasi Ruko Strategis",
        "C → Survey Tempat, Aksesibilitas & Potensi Pasar Lingkungan",
        "D → Negosiasi & Penandatanganan Akad Kontrak Sewa Ruko",
        "E → Perancangan Desain Arsitektur Interior & Tata Letak Dapur Stainless",
        "F → Pelaksanaan Renovasi Fisik & Pemasangan Instalasi Gas/Listrik/Air",
        "G → Pembelian & Pengadaan Mesin Rebus Mie, Chiller, POS & Peralatan Saji",
        "H → Pemasangan & Setting Tata Letak Ruang Dining & Dapur",
        "I → Uji Coba Fungsi Fasilitas, Sanitasi & Simulasi Alur Masak (Dry Run)",
        "J → Perekrutan Tenaga Kerja & Barista/Koki",
        "K → Pelatihan Standarisasi Resep, SOP Layanan & Pelatihan Kasir POS",
        "L → Pembukaan Resmi Usaha (Grand Opening & Peluncuran Promo)"
    ]
    for s in steps:
        add_p(doc, s, indent=1)

    add_heading(doc, "4.4 Gantt Chart Pelaksanaan", level=2)
    add_tbl_title(doc, f"Tabel 4.2 Gantt Chart {cfg['business_name']}")
    tg = doc.add_table(rows=1, cols=7)
    g_hdrs = ["Kegiatan", "Bln 1 (W1-2)", "Bln 1 (W3-4)", "Bln 2 (W1-2)", "Bln 2 (W3-4)", "Bln 3 (W1-2)", "Bln 3 (W3-4)"]
    for i, gh in enumerate(g_hdrs):
        tg.rows[0].cells[i].text = gh

    g_data = [
        ("A. Perencanaan Proposal", "X", "", "", "", "", ""),
        ("B & C. Lokasi & Survey", "X", "X", "", "", "", ""),
        ("D. Sewa Tempat", "", "X", "", "", "", ""),
        ("E & F. Renovasi Interior", "", "", "X", "X", "", ""),
        ("G & H. Peralatan & Setting", "", "", "", "X", "X", ""),
        ("I. Final Inspection", "", "", "", "", "X", ""),
        ("J & K. Rekrutmen & Training", "", "", "", "", "X", "X"),
        ("L. Grand Opening", "", "", "", "", "", "X")
    ]
    for act, w1, w2, w3, w4, w5, w6 in g_data:
        rw = tg.add_row()
        rw.cells[0].text = act
        for idx, val in enumerate([w1, w2, w3, w4, w5, w6], start=1):
            rw.cells[idx].text = val
    style_table(tg)

    # -----------------------------------------------------------------
    # BAB V: ASPEK KEUANGAN (LENGKAP 26 TABEL ASLI WORD!)
    # -----------------------------------------------------------------
    doc.add_page_break()
    add_heading(doc, "BAB V\nASPEK KEUANGAN", level=1)
    add_p(doc, f"Aspek keuangan merupakan pilar penentu utama dalam menilai kelayakan investasi bisnis {cfg['business_name']}. Seluruh asumsi arus kas disusun secara terukur selama 5 tahun ke depan dengan mengintegrasikan initial outlay, proyeksi cash inflow berbasis 25 hari kerja efektif bulanan (300 hari operasi tahunan), eskalasi harga dan biaya, hingga pengujian kriteria kelayakan modal.")

    # Formatted figures helpers
    cif_f = [f"Rp {x:,.0f}".replace(",", ".") for x in met['cif_years']]
    cof_f = [f"Rp {x:,.0f}".replace(",", ".") for x in met['cof_years']]
    ncf_f = [f"Rp {x:,.0f}".replace(",", ".") for x in met['ncf_years']]
    dep_f = f"Rp {met['depresiasi_per_year']:,.0f}".replace(",", ".")
    pro_f = [f"Rp {x:,.0f}".replace(",", ".") for x in met['proceed_years']]
    outlay_f = f"Rp {met['initial_outlay']:,.0f}".replace(",", ".")
    capex_f = f"Rp {met['capex_total']:,.0f}".replace(",", ".")

    # Tabel 5.1 NCF & Initial Outlay
    add_tbl_title(doc, f"Tabel 5.1. NCF & Initial Outlay {cfg['business_name']}")
    t51 = doc.add_table(rows=7, cols=7)
    for c_i, h in enumerate(["Keterangan", "Tahun 0", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t51.rows[0].cells[c_i].text = h
    t51_rows = [
        ("Investasi Awal", outlay_f, "-", "-", "-", "-", "-"),
        ("Cash Inflow / CIF", "-", cif_f[0], cif_f[1], cif_f[2], cif_f[3], cif_f[4]),
        ("Cash Outflow / COF (-)", "-", cof_f[0], cof_f[1], cof_f[2], cof_f[3], cof_f[4]),
        ("NCF = EAT", "-", ncf_f[0], ncf_f[1], ncf_f[2], ncf_f[3], ncf_f[4]),
        ("Depresiasi (+)", "-", dep_f, dep_f, dep_f, dep_f, dep_f),
        ("Proceed", "-", pro_f[0], pro_f[1], pro_f[2], pro_f[3], pro_f[4])
    ]
    for r_i, r_data in enumerate(t51_rows, start=1):
        for c_i, val in enumerate(r_data):
            t51.rows[r_i].cells[c_i].text = val
    style_table(t51)

    # Tabel 5.2 Aktiva Tetap
    add_tbl_title(doc, f"Tabel 5.2. Aktiva Tetap {cfg['business_name']}")
    t52 = doc.add_table(rows=1, cols=5)
    for idx, h in enumerate(["No.", "Keterangan", "Jumlah", "Harga Satuan", "Total"]):
        t52.rows[0].cells[idx].text = h
    for idx, item in enumerate(cfg.get("capex_items", []), start=1):
        tot_val = item["qty"] * item["unit_price"]
        row = t52.add_row()
        row.cells[0].text = str(idx)
        row.cells[1].text = item["name"]
        row.cells[2].text = str(item["qty"])
        row.cells[3].text = f"Rp {item['unit_price']:,}".replace(",", ".")
        row.cells[4].text = f"Rp {tot_val:,}".replace(",", ".")
    r_t52 = t52.add_row()
    r_t52.cells[1].text = "Total Biaya Aktiva Tetap"
    r_t52.cells[4].text = capex_f
    style_table(t52, [Inches(0.5), Inches(2.6), Inches(0.8), Inches(1.3), Inches(1.4)])

    # Tabel 5.3 Biaya Sewa
    add_tbl_title(doc, f"Tabel 5.3. Biaya Sewa {cfg['business_name']}")
    t53 = doc.add_table(rows=2, cols=5)
    for idx, h in enumerate(["No", "Keterangan", "Jangka Waktu (/tahun)", "Harga Sewa per tahun", "Harga Sewa per bulan"]):
        t53.rows[0].cells[idx].text = h
    t53.rows[1].cells[0].text = "1"
    t53.rows[1].cells[1].text = "Biaya Sewa Tempat Ruko 1 Lantai"
    t53.rows[1].cells[2].text = "1 Tahun"
    t53.rows[1].cells[3].text = f"Rp {cfg['operational_costs']['rent_per_year']:,}".replace(",", ".")
    t53.rows[1].cells[4].text = f"Rp {cfg['operational_costs']['rent_per_year']//12:,}".replace(",", ".")
    style_table(t53, [Inches(0.5), Inches(2.5), Inches(1.2), Inches(1.5), Inches(1.5)])

    # Tabel 5.4 Biaya AMDAL
    add_tbl_title(doc, f"Tabel 5.4. Biaya AMDAL {cfg['business_name']}")
    t54 = doc.add_table(rows=2, cols=5)
    for idx, h in enumerate(["No", "Keterangan", "Jumlah", "Harga satuan", "Total"]):
        t54.rows[0].cells[idx].text = h
    t54.rows[1].cells[0].text = "1"
    t54.rows[1].cells[1].text = "Tempat Sampah & Kebersihan Lingkungan"
    t54.rows[1].cells[2].text = "1"
    t54.rows[1].cells[3].text = f"Rp {cfg['operational_costs']['amdal_initial']:,}".replace(",", ".")
    t54.rows[1].cells[4].text = f"Rp {cfg['operational_costs']['amdal_initial']:,}".replace(",", ".")
    style_table(t54, [Inches(0.5), Inches(2.5), Inches(1.0), Inches(1.5), Inches(1.5)])

    # Tabel 5.5 Biaya Survey
    add_tbl_title(doc, f"Tabel 5.5. Biaya Survey {cfg['business_name']}")
    t55 = doc.add_table(rows=3, cols=5)
    for idx, h in enumerate(["No", "Keterangan", "Jumlah", "Harga satuan", "Total"]):
        t55.rows[0].cells[idx].text = h
    t55.rows[1].cells[0].text = "1"
    t55.rows[1].cells[1].text = "Biaya Konsumsi Tim Survei"
    t55.rows[1].cells[2].text = "1"
    t55.rows[1].cells[3].text = "Rp 250.000"
    t55.rows[1].cells[4].text = "Rp 250.000"
    t55.rows[2].cells[0].text = "2"
    t55.rows[2].cells[1].text = "Biaya Transportasi & Administrasi"
    t55.rows[2].cells[2].text = "1"
    t55.rows[2].cells[3].text = "Rp 250.000"
    t55.rows[2].cells[4].text = "Rp 250.000"
    style_table(t55, [Inches(0.5), Inches(2.5), Inches(1.0), Inches(1.5), Inches(1.5)])

    # Tabel 5.6 Biaya Promosi
    add_tbl_title(doc, f"Tabel 5.6. Biaya Promosi {cfg['business_name']}")
    t56 = doc.add_table(rows=2, cols=5)
    for idx, h in enumerate(["No", "Keterangan", "Jumlah", "Harga Iklan", "Total"]):
        t56.rows[0].cells[idx].text = h
    t56.rows[1].cells[0].text = "1"
    t56.rows[1].cells[1].text = "Spanduk & Banner Grand Opening"
    t56.rows[1].cells[2].text = "1"
    t56.rows[1].cells[3].text = f"Rp {cfg['operational_costs']['promosi_initial']:,}".replace(",", ".")
    t56.rows[1].cells[4].text = f"Rp {cfg['operational_costs']['promosi_initial']:,}".replace(",", ".")
    style_table(t56, [Inches(0.5), Inches(2.5), Inches(1.0), Inches(1.5), Inches(1.5)])

    # Tabel 5.7 Cash Inflow
    add_tbl_title(doc, f"Tabel 5.7. Cash Inflow {cfg['business_name']}")
    t57 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Kategori", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t57.rows[0].cells[idx].text = h
    t57.rows[1].cells[0].text = "Total Cash Inflow"
    for i in range(5):
        t57.rows[1].cells[i+1].text = cif_f[i]
    style_table(t57)

    # Tabel 5.8 Cash Outflow
    add_tbl_title(doc, f"Tabel 5.8. Cash Outflow {cfg['business_name']}")
    t58 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Kategori", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t58.rows[0].cells[idx].text = h
    t58.rows[1].cells[0].text = "Total Cash Outflow"
    for i in range(5):
        t58.rows[1].cells[i+1].text = cof_f[i]
    style_table(t58)

    add_p(doc, "\nPenjelasan pada bagian aspek keuangan:")
    add_heading(doc, "6.1 Initial Outlay", level=2)
    add_p(doc, f"Initial outlay adalah total modal awal yang dikeluarkan untuk mendanai pembukaan gerai baru {cfg['business_name']}, yaitu sebesar {outlay_f}. Modal awal ini dialokasikan untuk pengadaan aktiva tetap sebesar {capex_f}, biaya sewa ruko 1 tahun sebesar Rp {cfg['operational_costs']['rent_per_year']:,}".replace(",", ".") + f", biaya AMDAL/kebersihan sebesar Rp {cfg['operational_costs']['amdal_initial']:,}".replace(",", ".") + f", biaya survey tempat Rp {cfg['operational_costs']['survey_initial']:,}".replace(",", ".") + f", dan biaya promosi awal Rp {cfg['operational_costs']['promosi_initial']:,}".replace(",", ".") + ".")

    add_heading(doc, "6.2 Aktiva Tetap", level=2)
    add_p(doc, f"Nilai aktiva tetap usaha {cfg['business_name']} mencapai {capex_f}. Aktiva tetap diasumsikan memiliki umur ekonomis 5 tahun dengan tingkat penyusutan metode garis lurus sebesar 20% per tahun ({dep_f} per tahun).")

    add_heading(doc, "6.3 Biaya-Biaya", level=2)
    add_p(doc, "Biaya operasional mencakup beban sewa bangunan ruko, retribusi AMDAL kebersihan, survey perizinan, promosi pemasaran, beban gaji karyawan terstandarisasi, pembelian bahan baku mie dan daging segar, serta utilitas daya listrik, air PAM, dan bahan bakar gas kompor.")

    add_heading(doc, "6.4 Cash Inflow", level=2)
    add_p(doc, f"Cash inflow bersumber dari penjualan aneka hidangan bakmi otentik, swikiauw, pangsit, dan minuman segar dengan asumsi 25 hari kerja efektif per bulan (300 hari operasional per tahun). Pada Tahun 1 diproyeksikan total penerimaan mencapai {cif_f[0]} dan terus mengalami pertumbuhan volume dan eskalasi harga bertahap hingga mencapai {cif_f[4]} pada Tahun 5.")

    add_heading(doc, "6.5 Cash Outflow", level=2)
    add_p(doc, f"Cash outflow merupakan pengeluaran kas operasional rutin tahunan yang pada Tahun 1 sebesar {cof_f[0]} dan meningkat seiring inflasi bahan baku serta kenaikan gaji hingga mencapai {cof_f[4]} pada Tahun 5.")

    # Tabel 5.9 Perincian Gaji Karyawan
    add_tbl_title(doc, f"Tabel 5.9. Perincian Gaji Karyawan {cfg['business_name']}")
    t59 = doc.add_table(rows=1, cols=4)
    for idx, h in enumerate(["Jabatan", "Jumlah", "Gaji Bulanan", "Total Gaji Tahunan"]):
        t59.rows[0].cells[idx].text = h
    tot_gaji_thn = 0
    for emp in cfg['operational_costs']['employees']:
        tot_thn = emp['count'] * emp['monthly_salary'] * 12
        tot_gaji_thn += tot_thn
        rw = t59.add_row()
        rw.cells[0].text = emp['role']
        rw.cells[1].text = str(emp['count'])
        rw.cells[2].text = f"Rp {emp['monthly_salary']:,}".replace(",", ".")
        rw.cells[3].text = f"Rp {tot_thn:,}".replace(",", ".")
    rw_totg = t59.add_row()
    rw_totg.cells[0].text = "Total Gaji Karyawan"
    rw_totg.cells[3].text = f"Rp {tot_gaji_thn:,}".replace(",", ".")
    style_table(t59, [Inches(2.5), Inches(0.8), Inches(1.8), Inches(1.8)])

    # Tabel 5.10 Kenaikan Gaji Karyawan
    add_tbl_title(doc, f"Tabel 5.10. Perincian Kenaikan Gaji Karyawan {cfg['business_name']}")
    t510 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Keterangan", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t510.rows[0].cells[idx].text = h
    t510.rows[1].cells[0].text = "Total Gaji Tahunan"
    cur_g = tot_gaji_thn
    for i in range(5):
        if i > 0:
            cur_g = cur_g * (1 + 0.10 + i * 0.02)
        t510.rows[1].cells[i+1].text = f"Rp {cur_g:,.0f}".replace(",", ".")
    style_table(t510)

    # Tabel 5.11 Perincian Bahan Baku Bulanan
    add_tbl_title(doc, f"Tabel 5.11. Perincian Bahan Baku Bulanan {cfg['business_name']}")
    t511 = doc.add_table(rows=1, cols=5)
    for idx, h in enumerate(["Bahan Pokok", "Jumlah", "Satuan", "Harga Satuan", "Total Bulanan"]):
        t511.rows[0].cells[idx].text = h
    raw_samples = [
        ("Tepung Terigu & Mie Segar Basah", "250", "Kg", "Rp 22.000", "Rp 5.500.000"),
        ("Daging Babi Casiu & Daging Babi Kecap", "160", "Kg", "Rp 95.000", "Rp 15.200.000"),
        ("Daging Ayam Fillet Putih", "100", "Kg", "Rp 38.000", "Rp 3.800.000"),
        ("Udang Segar Isian Swikiaw", "40", "Kg", "Rp 85.000", "Rp 3.400.000"),
        ("Kulit Pangsit & Swikiaw Homemade", "60", "Pack", "Rp 25.000", "Rp 1.500.000"),
        ("Minyak Babi, Minyak Ayam & Minyak Wijen", "40", "Liter", "Rp 45.000", "Rp 1.800.000"),
        ("Sayuran Sawi Hijau, Daun Bawang & Tongcai", "100", "Kg", "Rp 15.000", "Rp 1.500.000"),
        ("Saus Belibis, Mangga Besar & Bumbu Racik", "1", "Paket", "Rp 1.800.000", "Rp 1.800.000")
    ]
    for r_mat in raw_samples:
        rw = t511.add_row()
        for c_i, v in enumerate(r_mat):
            rw.cells[c_i].text = v
    rw_totbb = t511.add_row()
    rw_totbb.cells[0].text = "Total Bahan Baku Bulanan"
    rw_totbb.cells[4].text = f"Rp {cfg['operational_costs']['raw_materials_monthly_base']:,}".replace(",", ".")
    style_table(t511, [Inches(2.5), Inches(0.8), Inches(0.8), Inches(1.4), Inches(1.5)])

    # Tabel 5.12 Kenaikan Bahan Baku 5 Tahun
    add_tbl_title(doc, f"Tabel 5.12. Perincian Kenaikan Bahan Baku {cfg['business_name']}")
    t512 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Keterangan", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t512.rows[0].cells[idx].text = h
    t512.rows[1].cells[0].text = "Bahan Baku Tahunan"
    cur_bb = cfg['operational_costs']['raw_materials_monthly_base'] * 12
    for i in range(5):
        if i > 0:
            cur_bb = cur_bb * (1 + 0.12 + i * 0.03)
        t512.rows[1].cells[i+1].text = f"Rp {cur_bb:,.0f}".replace(",", ".")
    style_table(t512)

    # Tabel 5.13 & 5.14 Listrik
    add_tbl_title(doc, f"Tabel 5.13. Perincian Biaya Listrik {cfg['business_name']}")
    t513 = doc.add_table(rows=3, cols=2)
    t513.rows[0].cells[0].text = "Komponen Listrik"
    t513.rows[0].cells[1].text = "Nominal"
    t513.rows[1].cells[0].text = "Penggunaan Listrik Bulanan (Kipas Angin, Kulkas, Lampu)"
    t513.rows[1].cells[1].text = f"Rp {cfg['operational_costs']['electricity_monthly_base']:,}".replace(",", ".")
    t513.rows[2].cells[0].text = "Penggunaan Listrik Tahunan"
    t513.rows[2].cells[1].text = f"Rp {cfg['operational_costs']['electricity_monthly_base']*12:,}".replace(",", ".")
    style_table(t513, [Inches(4.0), Inches(2.5)])

    add_tbl_title(doc, f"Tabel 5.14. Kenaikan Biaya Listrik {cfg['business_name']}")
    t514 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Keterangan", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t514.rows[0].cells[idx].text = h
    t514.rows[1].cells[0].text = "Listrik Tahunan"
    cur_l = cfg['operational_costs']['electricity_monthly_base'] * 12
    for i in range(5):
        if i > 0:
            cur_l = cur_l * (1 + 0.10 + i * 0.02)
        t514.rows[1].cells[i+1].text = f"Rp {cur_l:,.0f}".replace(",", ".")
    style_table(t514)

    # Tabel 5.15 & 5.16 Air PAM
    add_tbl_title(doc, f"Tabel 5.15. Perincian Biaya Air {cfg['business_name']}")
    t515 = doc.add_table(rows=3, cols=2)
    t515.rows[0].cells[0].text = "Komponen Air PAM"
    t515.rows[0].cells[1].text = "Nominal"
    t515.rows[1].cells[0].text = "Penggunaan Air Bulanan"
    t515.rows[1].cells[1].text = f"Rp {cfg['operational_costs']['water_monthly_base']:,}".replace(",", ".")
    t515.rows[2].cells[0].text = "Penggunaan Air Tahunan"
    t515.rows[2].cells[1].text = f"Rp {cfg['operational_costs']['water_monthly_base']*12:,}".replace(",", ".")
    style_table(t515, [Inches(3.5), Inches(3.0)])

    add_tbl_title(doc, f"Tabel 5.16. Kenaikan Biaya Air {cfg['business_name']}")
    t516 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Keterangan", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t516.rows[0].cells[idx].text = h
    t516.rows[1].cells[0].text = "Air PAM Tahunan"
    cur_a = cfg['operational_costs']['water_monthly_base'] * 12
    for i in range(5):
        if i > 0:
            cur_a = cur_a * (1 + 0.08 + i * 0.02)
        t516.rows[1].cells[i+1].text = f"Rp {cur_a:,.0f}".replace(",", ".")
    style_table(t516)

    # Tabel 5.17 & 5.18 Air Minum Galon
    add_tbl_title(doc, f"Tabel 5.17. Perincian Air Minum Galon {cfg['business_name']}")
    t517 = doc.add_table(rows=3, cols=4)
    for idx, h in enumerate(["Keterangan", "Total Harga", "Jumlah", "Harga Satuan"]):
        t517.rows[0].cells[idx].text = h
    t517.rows[1].cells[0].text = "Air Minum Bulanan"
    t517.rows[1].cells[1].text = "Rp 400.000"
    t517.rows[1].cells[2].text = "20"
    t517.rows[1].cells[3].text = "Rp 20.000"
    t517.rows[2].cells[0].text = "Air Minum Tahunan"
    t517.rows[2].cells[1].text = "Rp 4.800.000"
    t517.rows[2].cells[2].text = "240"
    t517.rows[2].cells[3].text = "Rp 20.000"
    style_table(t517, [Inches(2.5), Inches(1.5), Inches(1.0), Inches(1.5)])

    add_tbl_title(doc, f"Tabel 5.18. Kenaikan Air Minum Galon {cfg['business_name']}")
    t518 = doc.add_table(rows=2, cols=6)
    for idx, h in enumerate(["Keterangan", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]):
        t518.rows[0].cells[idx].text = h
    t518.rows[1].cells[0].text = "Total Tagihan Air Galon"
    for i in range(5):
        val = 4800000 + (i * 5 * 20000)
        t518.rows[1].cells[i+1].text = f"Rp {val:,}".replace(",", ".")
    style_table(t518)

    # Tabel 5.19 Initial Outlay Rekap
    add_tbl_title(doc, f"Tabel 5.19. Initial Outlay {cfg['business_name']}")
    t519 = doc.add_table(rows=7, cols=2)
    t519.rows[0].cells[0].text = "Komponen Investasi"
    t519.rows[0].cells[1].text = "Nominal"
    t519.rows[1].cells[0].text = "Aktiva Tetap"
    t519.rows[1].cells[1].text = capex_f
    t519.rows[2].cells[0].text = "Biaya Sewa Ruko"
    t519.rows[2].cells[1].text = f"Rp {cfg['operational_costs']['rent_per_year']:,}".replace(",", ".")
    t519.rows[3].cells[0].text = "Biaya AMDAL"
    t519.rows[3].cells[1].text = f"Rp {cfg['operational_costs']['amdal_initial']:,}".replace(",", ".")
    t519.rows[4].cells[0].text = "Biaya Survey"
    t519.rows[4].cells[1].text = f"Rp {cfg['operational_costs']['survey_initial']:,}".replace(",", ".")
    t519.rows[5].cells[0].text = "Biaya Promosi"
    t519.rows[5].cells[1].text = f"Rp {cfg['operational_costs']['promosi_initial']:,}".replace(",", ".")
    t519.rows[6].cells[0].text = "Total Initial Outlay"
    t519.rows[6].cells[1].text = outlay_f
    style_table(t519, [Inches(3.5), Inches(3.0)])

    # Tabel 5.20 Net Cash Flow
    add_tbl_title(doc, f"Tabel 5.20. Net Cash Flow {cfg['business_name']}")
    t520 = doc.add_table(rows=6, cols=2)
    t520.rows[0].cells[0].text = "Tahun"
    t520.rows[0].cells[1].text = "Net Cash Flow (EAT)"
    for i in range(5):
        t520.rows[i+1].cells[0].text = f"Tahun {i+1}"
        t520.rows[i+1].cells[1].text = ncf_f[i]
    style_table(t520, [Inches(3.0), Inches(3.5)])

    # Tabel 5.21 Opportunity Cost
    add_tbl_title(doc, f"Tabel 5.21. Opportunity Cost {cfg['business_name']}")
    t521 = doc.add_table(rows=2, cols=2)
    t521.rows[0].cells[0].text = "Parameter"
    t521.rows[0].cells[1].text = "Nilai"
    t521.rows[1].cells[0].text = "Opportunity Cost (Tingkat Diskonto WACC)"
    t521.rows[1].cells[1].text = "20%"
    style_table(t521, [Inches(4.5), Inches(2.0)])

    # Tabel 5.22 Investasi
    add_tbl_title(doc, f"Tabel 5.22. Standar Kriteria Investasi {cfg['business_name']}")
    t522 = doc.add_table(rows=5, cols=2)
    t522.rows[0].cells[0].text = "Indikator Kelayakan"
    t522.rows[0].cells[1].text = "Kriteria Kelayakan"
    t522.rows[1].cells[0].text = "Net Present Value (NPV)"
    t522.rows[1].cells[1].text = "> 0 (Positif)"
    t522.rows[2].cells[0].text = "Internal Rate of Return (IRR)"
    t522.rows[2].cells[1].text = "> 20% (Opportunity Cost)"
    t522.rows[3].cells[0].text = "Payback Period (PP)"
    t522.rows[3].cells[1].text = "< 3.0 Tahun"
    t522.rows[4].cells[0].text = "Profitability Index (PI)"
    t522.rows[4].cells[1].text = "> 1.20"
    style_table(t522, [Inches(3.5), Inches(3.0)])

    # Tabel 5.23 Net Present Value (NPV)
    add_tbl_title(doc, f"Tabel 5.23. Net Present Value {cfg['business_name']}")
    t523 = doc.add_table(rows=8, cols=3)
    t523.rows[0].cells[0].text = "Tahun"
    t523.rows[0].cells[1].text = "Cashflow"
    t523.rows[0].cells[2].text = "Present Value (PV 20%)"
    t523.rows[1].cells[0].text = "0"
    t523.rows[1].cells[1].text = f"-{outlay_f}"
    t523.rows[1].cells[2].text = f"-{outlay_f}"
    for i in range(5):
        t523.rows[i+2].cells[0].text = str(i+1)
        t523.rows[i+2].cells[1].text = pro_f[i]
        t523.rows[i+2].cells[2].text = f"Rp {met['pv_years'][i]:,.0f}".replace(",", ".")
    t523.rows[7].cells[0].text = "Net Present Value (NPV)"
    t523.rows[7].cells[2].text = f"Rp {met['npv']:,.0f}".replace(",", ".")
    style_table(t523, [Inches(2.0), Inches(2.3), Inches(2.4)])

    # Tabel 5.24 Internal Rate of Return (IRR)
    add_tbl_title(doc, f"Tabel 5.24. Internal Rate of Return {cfg['business_name']}")
    t524 = doc.add_table(rows=8, cols=2)
    t524.rows[0].cells[0].text = "Tahun"
    t524.rows[0].cells[1].text = "Cashflow"
    t524.rows[1].cells[0].text = "0"
    t524.rows[1].cells[1].text = f"-{outlay_f}"
    for i in range(5):
        t524.rows[i+2].cells[0].text = str(i+1)
        t524.rows[i+2].cells[1].text = pro_f[i]
    t524.rows[7].cells[0].text = "Internal Rate of Return (IRR)"
    t524.rows[7].cells[1].text = f"{met['irr']*100:.2f}%"
    style_table(t524, [Inches(3.5), Inches(3.0)])

    # Tabel 5.25 Payback Period (PP)
    add_tbl_title(doc, f"Tabel 5.25. Payback Period {cfg['business_name']}")
    t525 = doc.add_table(rows=8, cols=4)
    for idx, h in enumerate(["Tahun", "Cashflow", "Akumulasi", "Payback Period"]):
        t525.rows[0].cells[idx].text = h
    t525.rows[1].cells[0].text = "0"
    t525.rows[1].cells[1].text = f"-{outlay_f}"
    t525.rows[1].cells[2].text = f"-{outlay_f}"
    t525.rows[1].cells[3].text = "-"
    cum = -met['initial_outlay']
    for i in range(5):
        cum += met['proceed_years'][i]
        t525.rows[i+2].cells[0].text = str(i+1)
        t525.rows[i+2].cells[1].text = pro_f[i]
        t525.rows[i+2].cells[2].text = f"Rp {cum:,.0f}".replace(",", ".")
        if i == 0:
            t525.rows[i+2].cells[3].text = f"{met['pp_months']:.2f} Bulan"
        else:
            t525.rows[i+2].cells[3].text = ""
    t525.rows[7].cells[0].text = "Kesimpulan Payback Period"
    t525.rows[7].cells[3].text = f"{met['pp_months']:.2f} Bulan ({met['pp_years']:.2f} Thn)"
    style_table(t525, [Inches(1.0), Inches(1.8), Inches(1.8), Inches(2.1)])

    # Tabel 5.26 Profitability Index (PI)
    add_tbl_title(doc, f"Tabel 5.26. Profitability Index {cfg['business_name']}")
    t526 = doc.add_table(rows=2, cols=2)
    t526.rows[0].cells[0].text = "Parameter"
    t526.rows[0].cells[1].text = "Hasil Rasio"
    t526.rows[1].cells[0].text = "Profitability Index (PI)"
    t526.rows[1].cells[1].text = f"{met['pi']:.2f}"
    style_table(t526, [Inches(4.5), Inches(2.0)])

    # -----------------------------------------------------------------
    # BAB VI: KESIMPULAN
    # -----------------------------------------------------------------
    doc.add_page_break()
    add_heading(doc, "BAB VI\nKESIMPULAN", level=1)
    add_p(doc, f"Demikian proposal Studi Kelayakan Bisnis {cfg['business_name']} disusun secara komprehensif dan mendalam. Berdasarkan hasil analisis aspek pasar, manajemen, teknis operasional, serta pemodelan finansial diperoleh kesimpulan kelayakan sebagai berikut:")

    add_tbl_title(doc, f"Tabel 6.1 Kesimpulan Kelayakan {cfg['business_name']}")
    t61 = doc.add_table(rows=5, cols=4)
    for idx, h in enumerate(["Kriteria", "Hasil", "Target", "Kelayakan"]):
        t61.rows[0].cells[idx].text = h
    eval_matrix = [
        ("Net Present Value (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "Positif (> 0)", "LAYAK"),
        ("Internal Rate of Return (IRR)", f"{met['irr']*100:.2f}%", "> 20.00% (Opportunity Cost)", "LAYAK"),
        ("Payback Period (PP)", f"{met['pp_months']:.2f} Bulan (< 1 Thn)", "< 3.0 Tahun", "LAYAK"),
        ("Profitability Index (PI)", f"{met['pi']:.2f}", "> 1.20", "LAYAK")
    ]
    for r_i, r_vals in enumerate(eval_matrix, start=1):
        for c_i, val in enumerate(r_vals):
            t61.rows[r_i].cells[c_i].text = val
    style_table(t61, [Inches(2.5), Inches(1.8), Inches(1.8), Inches(1.0)])

    add_p(doc, f"Berdasarkan seluruh kriteria evaluasi investasi di atas, bisnis {cfg['business_name']} telah melampaui standar kelayakan yang disyaratkan oleh Fakultas Ekonomi & Bisnis UKRIDA. Maka, usaha {cfg['business_name']} dinyatakan SANGAT LAYAK untuk direalisasikan.")

    add_heading(doc, "7.1 Kesimpulan dan Saran", level=2)
    add_p(doc, f"Kajian studi kelayakan membuktikan bahwa {cfg['business_name']} memiliki kekuatan diferensiasi cita rasa autentik yang telah teruji sejak 1999 di kawasan Duri Kosambi, basis pelanggan setia yang solid dari civitas Sekolah Kristen Kalam Kudus dan perumahan sekitar, serta kelayakan finansial yang sangat kuat. Usaha ini menghasilkan nilai NPV positif sebesar Rp {met['npv']:,.0f}".replace(",", ".") + f", tingkat pengembalian IRR mencapai {met['irr']*100:.2f}%, dan periode pengembalian modal modal awal (Payback Period) yang sangat singkat yakni hanya {met['pp_months']:.2f} bulan.")
    add_p(doc, "Untuk memaksimalkan keberhasilan dan menjaga keberlangsungan usaha, disarankan beberapa langkah strategis:")
    add_p(doc, "1. Menjaga konsistensi keaslian rasa mie dan topping melalui SOP takaran bahan baku yang terstandarisasi.", indent=1)
    add_p(doc, "2. Mengoptimalkan saluran pemasaran digital ojek online (GoFood, GrabFood, ShopeeFood) serta program promosi jam makan siang.", indent=1)
    add_p(doc, "3. Menjalin kontrak pasokan jangka panjang dengan distributor daging segar dan tepung terigu untuk mengunci harga bahan baku dan menghindari fluktuasi harga pasar.", indent=1)

    out_docx_2 = os.path.join(base, "naskah utama", "Laporan_SKB_Bakmi_Mimu_Carina_Sayang.docx")
    doc.save(out_docx)
    doc.save(out_docx_2)
    print(f"[SUCCESS] Laporan Master SKB Bakmi Mimu berhasil dibuat:")
    print(f"  - {out_docx}")
    print(f"  - {out_docx_2}")

if __name__ == "__main__":
    build_master_docx()
