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

def apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, line_spacing=1.5, space_after=Pt(6), space_before=Pt(0)):
    p.alignment = align
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = space_after
    p.paragraph_format.space_before = space_before

def add_styled_heading(doc, text, level=1):
    p = doc.add_paragraph()
    if level == 1:
        apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.5, space_after=Pt(12), space_before=Pt(18))
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.bold = True
    elif level == 2:
        apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.LEFT, line_spacing=1.5, space_after=Pt(6), space_before=Pt(12))
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.bold = True
    elif level == 3:
        apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.LEFT, line_spacing=1.5, space_after=Pt(4), space_before=Pt(6))
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.bold = True
    return p

def add_body_p(doc, text, bold_prefix="", indent=0):
    p = doc.add_paragraph()
    apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, line_spacing=1.5, space_after=Pt(6))
    if indent > 0:
        p.paragraph_format.left_indent = Inches(indent * 0.25)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(12)
        r_pre.bold = True
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    return p

def add_table_title(doc, title_text):
    p = doc.add_paragraph()
    apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.15, space_after=Pt(4), space_before=Pt(10))
    run = p.add_run(title_text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    run.bold = True
    return p

def add_figure_caption(doc, caption_text):
    p = doc.add_paragraph()
    apply_paragraph_format(p, align=WD_PARAGRAPH_ALIGNMENT.CENTER, line_spacing=1.15, space_after=Pt(10), space_before=Pt(4))
    run = p.add_run(caption_text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    run.bold = True
    return p

def format_table(table, col_widths=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    border_kwargs = {
        'top': {'sz': 4, 'val': 'single', 'color': '808080'},
        'bottom': {'sz': 4, 'val': 'single', 'color': '808080'},
        'left': {'sz': 4, 'val': 'single', 'color': '808080'},
        'right': {'sz': 4, 'val': 'single', 'color': '808080'},
        'insideH': {'sz': 4, 'val': 'single', 'color': 'D0D0D0'},
        'insideV': {'sz': 4, 'val': 'single', 'color': 'D0D0D0'},
    }
    for row_idx, row in enumerate(table.rows):
        is_header = (row_idx == 0)
        for col_idx, cell in enumerate(row.cells):
            set_cell_border(cell, **border_kwargs)
            if is_header:
                set_cell_shading(cell, "D9E1F2")
            if col_widths and col_idx < len(col_widths):
                cell.width = col_widths[col_idx]
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(10)
                    if is_header:
                        r.bold = True

def generate_bakmi_mimu_docx():
    with open(r"d:\Perkuliahan\Kelass\SKB\config_bakmi_mimu.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)
    with open(r"d:\Perkuliahan\Kelass\SKB\metrics_bakmi_mimu.json", "r", encoding="utf-8") as f:
        met = json.load(f)

    doc = docx.Document()

    # Set Margins (2.54 cm all around)
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)

    # Base Normal Style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)

    # -------------------------------------------------------------
    # COVER PAGE
    # -------------------------------------------------------------
    p_cov_1 = doc.add_paragraph()
    apply_paragraph_format(p_cov_1, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=Pt(36), space_after=Pt(6))
    r = p_cov_1.add_run("Studi Kelayakan Bisnis")
    r.bold = True
    r.font.size = Pt(14)

    p_cov_2 = doc.add_paragraph()
    apply_paragraph_format(p_cov_2, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(120))
    r = p_cov_2.add_run(f"“ {cfg['business_name']} ”\n{cfg['tagline']}")
    r.bold = True
    r.font.size = Pt(17)

    p_cov_3 = doc.add_paragraph()
    apply_paragraph_format(p_cov_3, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(6))
    r = p_cov_3.add_run(f"Disusun Oleh:\n{cfg['author']['name']}\nNIM: {cfg['author']['nim']}")
    r.font.size = Pt(12)

    p_cov_4 = doc.add_paragraph()
    apply_paragraph_format(p_cov_4, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=Pt(90), space_after=Pt(4))
    r = p_cov_4.add_run(f"{cfg['author']['institution']}\n{cfg['author']['faculty']}\n{cfg['author']['year']}")
    r.bold = True
    r.font.size = Pt(13)

    doc.add_page_break()

    # -------------------------------------------------------------
    # BAB I: PENDAHULUAN
    # -------------------------------------------------------------
    add_styled_heading(doc, "BAB 1\nPENDAHULUAN", level=1)
    
    add_body_p(doc, "Latar Belakang", bold_prefix="")
    add_body_p(doc, cfg['background']['history'])
    add_body_p(doc, cfg['background']['opportunity'])
    add_body_p(doc, cfg['background']['challenges'])
    add_body_p(doc, cfg['background']['concept'])

    add_body_p(doc, f"Berikut merupakan rincian produk {cfg['business_name']}:")
    add_table_title(doc, f"Tabel 1.1. Jenis Produk {cfg['business_name']}")

    t1 = doc.add_table(rows=1, cols=3)
    t1.rows[0].cells[0].text = "No."
    t1.rows[0].cells[1].text = f"Produk {cfg['business_name']}"
    t1.rows[0].cells[2].text = "Harga"
    for i, prod in enumerate(cfg.get("products", []), start=1):
        row = t1.add_row()
        row.cells[0].text = str(i)
        row.cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row.cells[1].text = prod["name"]
        row.cells[2].text = f"Rp {prod['price_year_1']:,}".replace(",", ".")
        row.cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    format_table(t1, [Inches(0.6), Inches(4.2), Inches(1.8)])

    add_body_p(doc, f"\nBerikut merupakan gambaran visual dari produk {cfg['business_name']}:")
    add_figure_caption(doc, f"Gambar 1.1 Gambaran Produk {cfg['business_name']}")

    # -------------------------------------------------------------
    # BAB II: ASPEK PEMASARAN DAN PASAR
    # -------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB II\nASPEK PEMASARAN DAN PASAR", level=1)

    add_styled_heading(doc, "2.1 Segmentasi Pasar", level=2)
    add_body_p(doc, f"Aspek pemasaran merupakan pilar fundamental yang perlu diperhatikan dalam ekspansi usaha kuliner {cfg['business_name']}. Dengan memahami karakteristik pasar secara mendalam, manajemen dapat memetakan potensi konsumen, merancang saluran promosi yang efisien, dan mencapai target volume penjualan yang diharapkan. Segmentasi pasar membagi pasar heterogen ke dalam kelompok konsumen homogen yang memiliki kebutuhan dan preferensi serupa.")
    add_body_p(doc, f"Secara demografis, segmen pasar dari {cfg['business_name']} yaitu {cfg['marketing']['segmentation']['demographic']}")
    add_body_p(doc, f"Secara geografis, segmen pasar utama berpusat pada {cfg['marketing']['segmentation']['geographic']}")
    add_body_p(doc, f"Secara perilaku dan psikografis, target pasar adalah {cfg['marketing']['segmentation']['behavioral']}")

    add_styled_heading(doc, "2.2 Diferensiasi Usaha", level=2)
    add_body_p(doc, "Untuk memenangkan persaingan di industri kuliner bakmi Jakarta Barat, usaha ini memiliki 5 pilar diferensiasi unggulan:")
    add_body_p(doc, cfg['marketing']['differentiation']['taste'], bold_prefix="- Rasa: ")
    add_body_p(doc, cfg['marketing']['differentiation']['location'], bold_prefix="- Lokasi: ")
    add_body_p(doc, cfg['marketing']['differentiation']['service'], bold_prefix="- Layanan: ")
    add_body_p(doc, cfg['marketing']['differentiation']['size'], bold_prefix="- Ukuran: ")
    add_body_p(doc, cfg['marketing']['differentiation']['presentation'], bold_prefix="- Penyajian: ")

    add_styled_heading(doc, "2.3 Aspek Lokasi", level=2)
    add_body_p(doc, f"{cfg['business_name']} berlokasi di titik yang sangat strategis: {cfg['marketing']['location']}. Lokasi ini berada persis di depan gerbang utama institusi pendidikan dan perumahan padat penduduk yang menjamin tingginya lalu lintas pengunjung alami (foot traffic), akses parkir yang memadai, serta visibilitas gerai yang sangat optimal.")
    add_figure_caption(doc, f"Gambar 2.1 Lokasi {cfg['business_name']}")

    add_styled_heading(doc, "2.4 Target Pemasaran", level=2)
    add_body_p(doc, f"Target pemasaran utama dari {cfg['business_name']} difokuskan kepada para orang tua murid dan guru sekolah Kalam Kudus saat jam antar-jemput, warga perumahan Duri Kosambi dan Semanan yang mencari sarapan dan makan siang keluarga, serta komunitas pecinta bakmi otentik di area Jakarta Barat dengan kisaran harga yang kompetitif ({cfg['marketing']['price_range']}).")

    add_styled_heading(doc, "2.5 Bauran Pemasaran (7P)", level=2)
    add_body_p(doc, "Bauran pemasaran 7P diterapkan secara komprehensif sebagai berikut:")
    p7 = [
        ("Product: ", "Menyajikan bakmi bertekstur kenyal berkilau alami tanpa bahan pengawet dengan pilihan topping ayam jamur gurih, babi kecap, dan casiu manis lezat, serta pelengkap swikiauw udang-babi dan pangsit homemade."),
        ("Price: ", f"Penetapan harga berbasis nilai (value-based pricing) yang sangat kompetitif mulai dari {cfg['marketing']['price_range']} yang terjangkau bagi kantong keluarga maupun pelajar."),
        ("Place: ", "Lokasi gerai ruko komersial yang bersih, ber-AC, nyaman, mudah diakses kendaraan roda dua dan roda empat, serta terdaftar resmi pada layanan pesan-antar online."),
        ("Promotion: ", "Promosi melalui media sosial (Instagram, TikTok Kuliner), ulasan food vlogger lokal, pembagian kupon potongan harga saat grand opening, dan ulasan Google Maps bintang lima."),
        ("People: ", "Koki berpengalaman resep keluarga, kasir yang teliti dan ramah, pramusaji yang sigap menyajikan pesanan, serta staf dapur yang mengenakan seragam, masker, dan celemek bersih."),
        ("Process: ", "Sistem peracikan pesanan terbuka (open-kitchen) berstandar sanitasi tinggi, durasi pembuatan di bawah 7 menit per mangkok, dan standar kemasan takeaway kedap udara."),
        ("Physical Evidence: ", "Desain interior ruko bertema oriental modern yang bersih, penerangan terang dan hangat, tata meja kursi nyaman, serta kemasan takeaway berlogo resmi.")
    ]
    for b_pre, b_txt in p7:
        add_body_p(doc, b_txt, bold_prefix=b_pre, indent=1)

    # -------------------------------------------------------------
    # BAB III: ASPEK MANAJEMEN
    # -------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB III\nASPEK MANAJEMEN", level=1)

    add_body_p(doc, "Struktur Organisasi", bold_prefix="")
    add_figure_caption(doc, f"Gambar 3.1 Struktur Organisasi {cfg['business_name']}")

    add_body_p(doc, "Struktur organisasi dirancang terstruktur dan fungsional agar operasional harian gerai berjalan produktif dan efisien. Berikut rincian uraian pekerjaan (Job Description) untuk setiap posisi:")
    for role in cfg['management']['organization_roles']:
        add_body_p(doc, role['job_desc'], bold_prefix=f"{role['title']}:\n", indent=0)

    # -------------------------------------------------------------
    # BAB IV: ASPEK TEKNIS OPERASI
    # -------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB IV\nASPEK TEKNIS OPERASI", level=1)

    add_styled_heading(doc, "4.1 Aktiva Tetap & Bahan Baku", level=2)
    add_styled_heading(doc, "4.1.1 Gedung dan Bangunan", level=3)
    add_body_p(doc, f"{cfg['business_name']} memilih sistem sewa tempat usaha: {cfg['operations']['facility_status']}. Bangunan ruko direnovasi untuk mengakomodasi seluruh fungsi operasional gerai, mulai dari ruang makan utama, dapur bersih, gudang penyimpanan dingin, kasir, hingga area tunggu pengemudi ojek online.")

    add_styled_heading(doc, "4.1.2 Peralatan dan Perlengkapan", level=3)
    add_body_p(doc, f"Peralatan dan perlengkapan merupakan aktiva tetap (capex) operasional yang digunakan langsung untuk kegiatan produksi dan pelayanan {cfg['business_name']}. Berikut rincian aktiva tetap yang dibutuhkan secara lengkap:")

    add_table_title(doc, f"Tabel 4.1 Peralatan dan Perlengkapan {cfg['business_name']}")
    t_capex = doc.add_table(rows=1, cols=5)
    for idx, h in enumerate(["No.", "Keterangan", "Jumlah", "Harga Satuan", "Total"]):
        t_capex.rows[0].cells[idx].text = h

    tot_capex_calc = 0
    for idx, item in enumerate(cfg.get("capex_items", []), start=1):
        tot_val = item["qty"] * item["unit_price"]
        tot_capex_calc += tot_val
        row = t_capex.add_row()
        row.cells[0].text = str(idx)
        row.cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row.cells[1].text = item["name"]
        row.cells[2].text = str(item["qty"])
        row.cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row.cells[3].text = f"Rp {item['unit_price']:,}".replace(",", ".")
        row.cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        row.cells[4].text = f"Rp {tot_val:,}".replace(",", ".")
        row.cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    r_tot = t_capex.add_row()
    r_tot.cells[1].text = "Total Biaya Aktiva Tetap"
    r_tot.cells[4].text = f"Rp {tot_capex_calc:,}".replace(",", ".")
    r_tot.cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    format_table(t_capex, [Inches(0.5), Inches(2.6), Inches(0.8), Inches(1.3), Inches(1.4)])

    add_styled_heading(doc, "4.2 Layout Usaha", level=2)
    add_figure_caption(doc, f"Gambar 4.1 Layout {cfg['business_name']}")
    add_body_p(doc, f"Tata letak ruangan {cfg['business_name']} dirancang mengutamakan kebersihan, kelancaran arus staf dapur, dan kenyamanan pengunjung bersantap:")
    for rm in cfg['operations']['rooms']:
        add_body_p(doc, rm, bold_prefix="- ", indent=1)

    add_styled_heading(doc, "4.3 Network Planning", level=2)
    add_figure_caption(doc, f"Gambar 4.2 Network Planning {cfg['business_name']}")
    add_body_p(doc, "Keterangan urutan aktivitas pra-operasional (Critical Path Activity A s/d L):")
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
        add_body_p(doc, s, indent=1)

    add_styled_heading(doc, "4.4 Gantt Chart Pelaksanaan", level=2)
    add_table_title(doc, f"Tabel 4.2 Gantt Chart {cfg['business_name']}")
    t_gantt = doc.add_table(rows=1, cols=7)
    gantt_hdrs = ["Kegiatan", "Bln 1 (W1-2)", "Bln 1 (W3-4)", "Bln 2 (W1-2)", "Bln 2 (W3-4)", "Bln 3 (W1-2)", "Bln 3 (W3-4)"]
    for i, gh in enumerate(gantt_hdrs):
        t_gantt.rows[0].cells[i].text = gh

    gantt_data = [
        ("A. Perencanaan Proposal", "X", "", "", "", "", ""),
        ("B & C. Lokasi & Survey", "X", "X", "", "", "", ""),
        ("D. Sewa Tempat", "", "X", "", "", "", ""),
        ("E & F. Renovasi Interior", "", "", "X", "X", "", ""),
        ("G & H. Peralatan & Setting", "", "", "", "X", "X", ""),
        ("I. Final Inspection", "", "", "", "", "X", ""),
        ("J & K. Rekrutmen & Training", "", "", "", "", "X", "X"),
        ("L. Grand Opening", "", "", "", "", "", "X")
    ]
    for act, w1, w2, w3, w4, w5, w6 in gantt_data:
        rw = t_gantt.add_row()
        rw.cells[0].text = act
        for idx, val in enumerate([w1, w2, w3, w4, w5, w6], start=1):
            rw.cells[idx].text = val
            rw.cells[idx].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    format_table(t_gantt)

    # -------------------------------------------------------------
    # BAB V: ASPEK KEUANGAN
    # -------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB V\nASPEK KEUANGAN", level=1)

    add_body_p(doc, f"Aspek keuangan merupakan pilar penentu utama dalam menilai kelayakan investasi bisnis {cfg['business_name']}. Seluruh asumsi arus kas disusun secara terukur selama 5 tahun ke depan dengan mengintegrasikan initial outlay, proyeksi cash inflow berbasis 25 hari kerja efektif bulanan (300 hari operasi tahunan), eskalasi harga dan biaya, hingga pengujian kriteria kelayakan modal.")

    # Table 5.1 NCF & Initial Outlay
    add_table_title(doc, f"Tabel 5.1. NCF & Initial Outlay {cfg['business_name']}")
    t_ncf = doc.add_table(rows=7, cols=7)
    ncf_hdrs = ["Keterangan", "Tahun 0", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]
    for c_i, h in enumerate(ncf_hdrs):
        t_ncf.rows[0].cells[c_i].text = h

    cif_fmt = [f"Rp {x:,.0f}".replace(",", ".") for x in met['cif_years']]
    cof_fmt = [f"Rp {x:,.0f}".replace(",", ".") for x in met['cof_years']]
    ncf_fmt = [f"Rp {x:,.0f}".replace(",", ".") for x in met['ncf_years']]
    dep_fmt = [f"Rp {met['depresiasi_per_year']:,.0f}".replace(",", ".")] * 5
    pro_fmt = [f"Rp {x:,.0f}".replace(",", ".") for x in met['proceed_years']]
    outlay_fmt = f"Rp {met['initial_outlay']:,.0f}".replace(",", ".")

    ncf_rows = [
        ("Investasi Awal", outlay_fmt, "-", "-", "-", "-", "-"),
        ("Cash Inflow / CIF", "-", cif_fmt[0], cif_fmt[1], cif_fmt[2], cif_fmt[3], cif_fmt[4]),
        ("Cash Outflow / COF (-)", "-", cof_fmt[0], cof_fmt[1], cof_fmt[2], cof_fmt[3], cof_fmt[4]),
        ("NCF = EAT", "-", ncf_fmt[0], ncf_fmt[1], ncf_fmt[2], ncf_fmt[3], ncf_fmt[4]),
        ("Depresiasi (+)", "-", dep_fmt[0], dep_fmt[1], dep_fmt[2], dep_fmt[3], dep_fmt[4]),
        ("Proceed", "-", pro_fmt[0], pro_fmt[1], pro_fmt[2], pro_fmt[3], pro_fmt[4])
    ]
    for r_idx, r_data in enumerate(ncf_rows, start=1):
        for c_idx, val in enumerate(r_data):
            t_ncf.rows[r_idx].cells[c_idx].text = val
            if c_idx > 0:
                t_ncf.rows[r_idx].cells[c_idx].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    format_table(t_ncf)

    add_body_p(doc, "\nPenjelasan pada bagian aspek keuangan:")
    add_styled_heading(doc, "6.1 Initial Outlay", level=2)
    add_body_p(doc, f"Initial outlay adalah total modal awal yang dikeluarkan untuk mendanai pembukaan gerai baru {cfg['business_name']}, yaitu sebesar {outlay_fmt}. Modal awal ini dialokasikan untuk pengadaan aktiva tetap sebesar Rp {met['capex_total']:,.0f}".replace(",", ".") + f", biaya sewa ruko 1 tahun sebesar Rp {cfg['operational_costs']['rent_per_year']:,.0f}".replace(",", ".") + f", biaya AMDAL/kebersihan sebesar Rp {cfg['operational_costs']['amdal_initial']:,.0f}".replace(",", ".") + f", biaya survey tempat Rp {cfg['operational_costs']['survey_initial']:,.0f}".replace(",", ".") + f", dan biaya promosi awal Rp {cfg['operational_costs']['promosi_initial']:,.0f}".replace(",", ".") + ".")

    add_styled_heading(doc, "6.2 Aktiva Tetap", level=2)
    add_body_p(doc, f"Nilai aktiva tetap usaha {cfg['business_name']} mencapai Rp {met['capex_total']:,.0f}".replace(",", ".") + f". Aktiva tetap diasumsikan memiliki umur ekonomis 5 tahun dengan tingkat penyusutan metode garis lurus sebesar 20% per tahun (Rp {met['depresiasi_per_year']:,.0f}".replace(",", ".") + " per tahun).")

    add_styled_heading(doc, "6.3 Biaya-Biaya", level=2)
    add_body_p(doc, "Biaya operasional mencakup beban sewa bangunan ruko, retribusi AMDAL kebersihan, survey perizinan, promosi pemasaran, beban gaji karyawan terstandarisasi, pembelian bahan baku mie dan daging segar, serta utilitas daya listrik, air PAM, dan bahan bakar gas kompor.")

    add_styled_heading(doc, "6.4 Cash Inflow", level=2)
    add_body_p(doc, f"Cash inflow bersumber dari penjualan aneka hidangan bakmi otentik, swikiauw, pangsit, dan minuman segar dengan asumsi 25 hari kerja efektif per bulan (300 hari operasional per tahun). Pada Tahun 1 diproyeksikan total penerimaan mencapai {cif_fmt[0]} dan terus mengalami pertumbuhan volume dan eskalasi harga bertahap hingga mencapai {cif_fmt[4]} pada Tahun 5.")

    add_styled_heading(doc, "6.5 Cash Outflow", level=2)
    add_body_p(doc, f"Cash outflow merupakan pengeluaran kas operasional rutin tahunan yang pada Tahun 1 sebesar {cof_fmt[0]} dan meningkat seiring inflasi bahan baku serta kenaikan gaji hingga mencapai {cof_fmt[4]} pada Tahun 5.")

    # Evaluasi Finansial Table
    add_body_p(doc, f"\nBerikut ringkasan hasil perhitungan kriteria kelayakan investasi {cfg['business_name']} selama 5 tahun ke depan:")
    add_table_title(doc, f"Tabel 5.23. Hasil Uji Kelayakan Investasi {cfg['business_name']}")
    t_eval = doc.add_table(rows=5, cols=4)
    for c_i, h in enumerate(["Kriteria Evaluasi", "Hasil Perhitungan", "Standar Kelayakan", "Status"]):
        t_eval.rows[0].cells[c_i].text = h

    eval_rows = [
        ("Net Present Value (NPV)", f"Rp {met['npv']:,.0f}".replace(",", "."), "> 0 (Bernilai Positif)", "LAYAK"),
        ("Internal Rate of Return (IRR)", f"{met['irr']*100:.2f}%", "> 20.00% (Opportunity Cost)", "LAYAK"),
        ("Payback Period (PP)", f"{met['pp_months']:.2f} Bulan ({met['pp_years']:.2f} Thn)", "< 3.0 Tahun", "LAYAK"),
        ("Profitability Index (PI)", f"{met['pi']:.2f}", "> 1.20", "LAYAK")
    ]
    for r_i, r_vals in enumerate(eval_rows, start=1):
        for c_i, val in enumerate(r_vals):
            t_eval.rows[r_i].cells[c_i].text = val
            t_eval.rows[r_i].cells[c_i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    format_table(t_eval)

    # -------------------------------------------------------------
    # BAB VI: KESIMPULAN
    # -------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB VI\nKESIMPULAN", level=1)

    add_body_p(doc, f"Demikian proposal Studi Kelayakan Bisnis {cfg['business_name']} disusun dengan cermat dan mendalam. Berdasarkan hasil analisis komprehensif pada aspek pasar, manajemen, teknis operasi, serta aspek keuangan diperoleh kesimpulan kelayakan sebagai berikut:")

    add_table_title(doc, f"Tabel 6.1 Kesimpulan Kelayakan {cfg['business_name']}")
    t_sum = doc.add_table(rows=5, cols=4)
    for c_i, h in enumerate(["Kriteria", "Hasil", "Target", "Kelayakan"]):
        t_sum.rows[0].cells[c_i].text = h
    for r_i, r_vals in enumerate(eval_rows, start=1):
        for c_i, val in enumerate(r_vals):
            t_sum.rows[r_i].cells[c_i].text = val
            t_sum.rows[r_i].cells[c_i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    format_table(t_sum)

    add_body_p(doc, f"Berdasarkan matriks kriteria evaluasi di atas, seluruh parameter telah memenuhi dan melampaui standar kelayakan yang ditetapkan. Maka, usaha {cfg['business_name']} dinyatakan SANGAT LAYAK untuk direalisasikan.")

    add_styled_heading(doc, "7.1 Kesimpulan dan Saran", level=2)
    add_body_p(doc, f"Kajian studi kelayakan membuktikan bahwa {cfg['business_name']} memiliki kekuatan diferensiasi rasa otentik yang telah teruji sejak 1999, lokasi yang sangat strategis di depan gerbang Sekolah Kristen Kalam Kudus Duri Kosambi, dan basis konsumen loyal yang sangat solid. Secara finansial, usaha ini menghasilkan nilai NPV yang sangat positif (Rp {met['npv']:,.0f}".replace(",", ".") + f"), tingkat pengembalian IRR tinggi ({met['irr']*100:.2f}%), dan waktu balik modal yang sangat cepat yaitu hanya dalam tempo {met['pp_months']:.2f} bulan.")
    add_body_p(doc, "Untuk menjamin keberlanjutan dan keunggulan bersaing gerai, diajukan beberapa saran strategis:")
    add_body_p(doc, "1. Menjaga konsistensi cita rasa autentik dan kebersihan dapur terbuka melalui SOP mutu yang ketat.", indent=1)
    add_body_p(doc, "2. Memaksimalkan penetrasi pemasaran digital melalui kemitraan ojek online dan program promosi berkala.", indent=1)
    add_body_p(doc, "3. Menjalankan manajemen kas dan persediaan secara tertib guna mengantisipasi fluktuasi harga bahan baku daging segar.", indent=1)

    out_docx = r"d:\Perkuliahan\Kelass\SKB\Bakmi_Mimu_Laporan_SKB.docx"
    doc.save(out_docx)
    print(f"[SUCCESS] Laporan SKB Bakmi Mimu berhasil dibuat di: {out_docx}")

if __name__ == "__main__":
    generate_bakmi_mimu_docx()
