import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import sys
import os

def set_cell_border(cell, **kwargs):
    """
    kwargs: top, bottom, left, right
    values: dict(sz=12, val='single', color='A6A6A6')
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

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

def build_skb_document(config_path, output_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    doc = docx.Document()

    # Section & Page Setup (1 inch = 2.54 cm margins all around)
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)

    # Set base Normal style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)

    # ------------------------------------------------------------------
    # COVER PAGE
    # ------------------------------------------------------------------
    p_cov_1 = doc.add_paragraph()
    apply_paragraph_format(p_cov_1, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=Pt(36), space_after=Pt(6))
    r = p_cov_1.add_run("Studi Kelayakan Bisnis")
    r.bold = True
    r.font.size = Pt(14)

    p_cov_2 = doc.add_paragraph()
    apply_paragraph_format(p_cov_2, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(120))
    r = p_cov_2.add_run(f"“ {cfg['business_name']} ”")
    r.bold = True
    r.font.size = Pt(18)

    p_cov_3 = doc.add_paragraph()
    apply_paragraph_format(p_cov_3, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=Pt(6))
    r = p_cov_3.add_run(f"Disusun Oleh:\n{cfg['author']['name']}\nNIM: {cfg['author']['nim']}")
    r.font.size = Pt(12)

    p_cov_4 = doc.add_paragraph()
    apply_paragraph_format(p_cov_4, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=Pt(100), space_after=Pt(4))
    r = p_cov_4.add_run(f"{cfg['author']['institution']}\n{cfg['author']['faculty']}\n{cfg['author']['year']}")
    r.bold = True
    r.font.size = Pt(13)

    doc.add_page_break()

    # ------------------------------------------------------------------
    # BAB I: PENDAHULUAN
    # ------------------------------------------------------------------
    add_styled_heading(doc, "BAB 1\nPENDAHULUAN", level=1)
    
    add_body_p(doc, "Latar Belakang", bold_prefix="")
    add_body_p(doc, cfg['background']['history'])
    add_body_p(doc, cfg['background']['opportunity'])
    add_body_p(doc, cfg['background']['challenges'])
    add_body_p(doc, cfg['background']['concept'])
    
    add_body_p(doc, f"Berikut merupakan rincian produk {cfg['business_name']}:")
    add_table_title(doc, f"Tabel 1.1. Jenis Produk {cfg['business_name']}")
    
    # Table 1.1
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
    format_table(t1, [Inches(0.6), Inches(4.0), Inches(1.8)])

    add_body_p(doc, f"\nBerikut merupakan gambaran visual dari produk {cfg['business_name']}:")
    add_figure_caption(doc, f"Gambar 1.1 Gambaran Produk {cfg['business_name']}")

    # ------------------------------------------------------------------
    # BAB II: ASPEK PEMASARAN DAN PASAR
    # ------------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB II\nASPEK PEMASARAN DAN PASAR", level=1)
    
    add_styled_heading(doc, "2.1 Segmentasi Pasar", level=2)
    add_body_p(doc, "Aspek pemasaran merupakan hal penting yang perlu diperhatikan oleh pelaku usaha. Dengan memahami aspek-aspek tersebut, pelaku usaha dapat memasarkan produknya dengan efektif dan mencapai target pasar yang diinginkan. Segmentasi pasar merupakan proses membagi pasar menjadi kelompok-kelompok konsumen yang memiliki kebutuhan dan karakteristik serupa.")
    add_body_p(doc, f"Secara demografis, segmen pasar dari {cfg['business_name']} yaitu {cfg['marketing']['segmentation']['demographic']}")
    add_body_p(doc, f"Secara geografis, segmen pasar utama berpusat pada {cfg['marketing']['segmentation']['geographic']}")
    add_body_p(doc, f"Secara perilaku dan psikografis, target konsumen adalah {cfg['marketing']['segmentation']['behavioral']}")

    add_styled_heading(doc, "2.2 Diferensiasi Usaha", level=2)
    add_body_p(doc, cfg['marketing']['differentiation']['taste'], bold_prefix="- Rasa: ")
    add_body_p(doc, cfg['marketing']['differentiation']['location'], bold_prefix="- Lokasi: ")
    add_body_p(doc, cfg['marketing']['differentiation']['service'], bold_prefix="- Layanan: ")
    add_body_p(doc, cfg['marketing']['differentiation']['size'], bold_prefix="- Ukuran: ")
    add_body_p(doc, cfg['marketing']['differentiation']['presentation'], bold_prefix="- Penyajian: ")

    add_styled_heading(doc, "2.3 Aspek Lokasi", level=2)
    add_body_p(doc, f"{cfg['business_name']} berlokasi di daerah yang sangat strategis: {cfg['marketing']['location']}. Lokasi ini dipilih karena memiliki visibilitas tinggi, aksesibilitas yang sangat mudah, serta berada di pusat aktivitas civitas akademika dan pemukiman warga.")
    add_figure_caption(doc, f"Gambar 2.1 Lokasi {cfg['business_name']}")

    add_styled_heading(doc, "2.4 Target Pemasaran", level=2)
    add_body_p(doc, f"Target pemasaran utama dari {cfg['business_name']} difokuskan kepada pelajar, mahasiswa, dan masyarakat umum di sekitar lokasi dengan rentang harga terjangkau ({cfg['marketing']['price_range']}).")

    add_styled_heading(doc, "2.5 Bauran Pemasaran (7P)", level=2)
    add_body_p(doc, "Bauran pemasaran terdiri dari 7P sebagai fondasi keunggulan bersaing:")
    p7 = [
        ("Product: ", f"Produk minuman dan hidangan dibuat dengan bahan baku terstandarisasi, higienis, halal, dan memiliki cita rasa konsisten."),
        ("Price: ", f"Harga kompetitif mulai dari {cfg['marketing']['price_range']} yang sesuai dengan daya beli target pasar."),
        ("Place: ", f"Lokasi strategis di dekat kampus, fasilitas nyaman, bersih, dan mudah diakses kendaraan roda dua maupun mobil."),
        ("Promotion: ", "Promosi melalui media sosial (Instagram, TikTok, Google Maps Review), kemitraan komunitas kampus, dan promo grand opening."),
        ("People: ", "Barista dan staf terlatih dengan standar keramahan tinggi (hospitality), berpenampilan rapi, dan berseragam kerja khusus."),
        ("Process: ", "Sistem operasional terstandarisasi mulai dari pemesanan kasir, peracikan pesanan terbuka (open bar), hingga penyajian cepat."),
        ("Physical Evidence: ", "Desain interior modern, pencahayaan hangat, ketersediaan stopkontak di setiap meja, dan kemasan estetik berlogo.")
    ]
    for b_pre, b_txt in p7:
        add_body_p(doc, b_txt, bold_prefix=b_pre, indent=1)

    # ------------------------------------------------------------------
    # BAB III: ASPEK MANAJEMEN
    # ------------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB III\nASPEK MANAJEMEN", level=1)
    
    add_body_p(doc, "Struktur Organisasi", bold_prefix="")
    add_figure_caption(doc, f"Gambar 3.1 Struktur Organisasi {cfg['business_name']}")
    
    add_body_p(doc, "Berikut merupakan rincian pembagian tugas dan tanggung jawab (Job Description) pada masing-masing posisi:")
    for role in cfg['management']['organization_roles']:
        add_body_p(doc, role['job_desc'], bold_prefix=f"{role['title']}:\n", indent=0)

    # ------------------------------------------------------------------
    # BAB IV: ASPEK TEKNIS OPERASI
    # ------------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB IV\nASPEK TEKNIS OPERASI", level=1)
    
    add_styled_heading(doc, "4.1 Aktiva Tetap & Bahan Baku", level=2)
    add_styled_heading(doc, "4.1.1 Gedung dan Bangunan", level=3)
    add_body_p(doc, f"{cfg['business_name']} memilih sistem sewa tempat usaha: {cfg['operations']['facility_status']}. Bangunan direnovasi untuk mengakomodasi seluruh fungsi operasional dan kenyamanan konsumen.")
    
    add_styled_heading(doc, "4.1.2 Peralatan dan Perlengkapan", level=3)
    add_body_p(doc, f"Peralatan dan perlengkapan merupakan aset tetap (aktiva tetap) yang digunakan langsung untuk kegiatan operasional {cfg['business_name']}. Berikut rincian aktiva tetap yang dibutuhkan:")
    
    add_table_title(doc, f"Tabel 4.1 Peralatan dan Perlengkapan {cfg['business_name']}")
    t_capex = doc.add_table(rows=1, cols=5)
    for idx, h in enumerate(["No.", "Keterangan", "Jumlah", "Harga Satuan", "Total"]):
        t_capex.rows[0].cells[idx].text = h
    
    total_capex = 0
    for idx, item in enumerate(cfg.get("capex_items", []), start=1):
        tot = item["qty"] * item["unit_price"]
        total_capex += tot
        row = t_capex.add_row()
        row.cells[0].text = str(idx)
        row.cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row.cells[1].text = item["name"]
        row.cells[2].text = str(item["qty"])
        row.cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row.cells[3].text = f"Rp {item['unit_price']:,}".replace(",", ".")
        row.cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        row.cells[4].text = f"Rp {tot:,}".replace(",", ".")
        row.cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    r_tot = t_capex.add_row()
    r_tot.cells[1].text = "Total Biaya Aktiva Tetap"
    r_tot.cells[4].text = f"Rp {total_capex:,}".replace(",", ".")
    r_tot.cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    format_table(t_capex, [Inches(0.5), Inches(2.6), Inches(0.8), Inches(1.3), Inches(1.4)])

    add_styled_heading(doc, "4.2 Layout Usaha", level=2)
    add_figure_caption(doc, f"Gambar 4.1 Layout {cfg['business_name']}")
    add_body_p(doc, f"Tata letak ruangan {cfg['business_name']} dirancang optimal untuk menjamin kelancaran arus kerja dan kenyamanan pengunjung:")
    for rm in cfg['operations']['rooms']:
        add_body_p(doc, rm, bold_prefix="- ", indent=1)

    add_styled_heading(doc, "4.3 Network Planning", level=2)
    add_figure_caption(doc, f"Gambar 4.2 Network Planning {cfg['business_name']}")
    add_body_p(doc, "Keterangan alur aktivitas pra-operasional:")
    steps = [
        "A → Perencanaan Usaha (Penyusunan Proposal SKB)",
        "B → Pencarian Lokasi Usaha Strategis",
        "C → Survey Tempat & Lingkungan",
        "D → Negosiasi & Perjanjian Sewa Ruko",
        "E → Perancangan Desain Interior & Layout Toko",
        "F → Proses Renovasi & Pembangunan Fasilitas",
        "G → Pengadaan Peralatan, Mesin & Perlengkapan Bar",
        "H → Pemasangan & Setting Tata Letak Ruangan",
        "I → Final Inspection & Uji Fungsi Fasilitas",
        "J → Perekrutan Tenaga Kerja",
        "K → Pelatihan (Training) Barista & Staf",
        "L → Pembukaan Resmi Usaha (Grand Opening)"
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

    # ------------------------------------------------------------------
    # BAB V: ASPEK KEUANGAN
    # ------------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB V\nASPEK KEUANGAN", level=1)
    
    add_body_p(doc, f"Aspek keuangan merupakan salah satu aspek paling fundamental dalam studi kelayakan bisnis {cfg['business_name']}. Perencanaan keuangan yang matang disusun untuk memproyeksikan kelayakan investasi selama 5 tahun ke depan.")

    # Table 5.1 NCF & Initial Outlay
    add_table_title(doc, f"Tabel 5.1. NCF & Initial Outlay {cfg['business_name']}")
    t_ncf = doc.add_table(rows=7, cols=7)
    ncf_hdrs = ["Keterangan", "Tahun 0", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]
    for c_i, h in enumerate(ncf_hdrs):
        t_ncf.rows[0].cells[c_i].text = h
    
    ncf_rows = [
        ("Investasi Awal", "Rp 190.585.000", "-", "-", "-", "-", "-"),
        ("Cash Inflow / CIF", "-", "Rp 913.500.000", "Rp 1.263.600.000", "Rp 1.680.600.000", "Rp 2.115.300.000", "Rp 2.702.100.000"),
        ("Cash Outflow / COF (-)", "-", "Rp 416.386.000", "Rp 472.371.400", "Rp 550.098.352", "Rp 651.125.946", "Rp 799.839.300"),
        ("NCF = EAT", "-", "Rp 497.114.000", "Rp 791.228.600", "Rp 1.130.501.648", "Rp 1.464.174.054", "Rp 1.902.260.700"),
        ("Depresiasi (+)", "-", "Rp 29.702.000", "Rp 29.702.000", "Rp 29.702.000", "Rp 29.702.000", "Rp 29.702.000"),
        ("Proceed", "-", "Rp 467.412.000", "Rp 761.526.600", "Rp 1.100.799.648", "Rp 1.434.472.054", "Rp 1.872.558.700")
    ]
    for r_idx, r_data in enumerate(ncf_rows, start=1):
        for c_idx, val in enumerate(r_data):
            t_ncf.rows[r_idx].cells[c_idx].text = val
            if c_idx > 0:
                t_ncf.rows[r_idx].cells[c_idx].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    format_table(t_ncf)

    add_body_p(doc, "\nPenjelasan pada bagian aspek keuangan:")
    add_styled_heading(doc, "6.1 Initial Outlay", level=2)
    add_body_p(doc, f"Initial outlay adalah total modal awal yang dikeluarkan untuk memulai usaha {cfg['business_name']}. Modal ini mencakup aktiva tetap, sewa tempat, AMDAL, survey pasar, dan promosi awal.")

    add_styled_heading(doc, "6.2 Aktiva Tetap", level=2)
    add_body_p(doc, "Aktiva tetap memiliki umur ekonomis 5 tahun dengan tingkat penyusutan metode garis lurus sebesar 20% per tahun.")

    add_styled_heading(doc, "6.3 Biaya-Biaya", level=2)
    add_body_p(doc, "Biaya operasional bulanan dan tahunan mencakup biaya sewa ruko, AMDAL kebersihan, survey, promosi, gaji tenaga kerja, serta pembelian bahan baku dan utilitas daya listrik dan air.")

    add_styled_heading(doc, "6.4 Cash Inflow", level=2)
    add_body_p(doc, "Cash inflow bersumber dari penjualan menu harian dengan asumsi 25 hari kerja efektif per bulan (300 hari operasi per tahun), dengan proyeksi peningkatan kuantitas dan penyesuaian harga tahunan.")

    add_styled_heading(doc, "6.5 Cash Outflow", level=2)
    add_body_p(doc, "Cash outflow merupakan beban operasional tahunan yang terdiri dari beban gaji karyawan, konsumsi bahan baku, listrik PLN, air PAM, air galon, dan gas.")

    # NPV, IRR, Payback Period, PI summary
    add_body_p(doc, f"\nBerikut ringkasan hasil uji kelayakan investasi {cfg['business_name']} selama 5 tahun:")
    add_table_title(doc, f"Tabel 5.23. Hasil Uji Kelayakan Investasi {cfg['business_name']}")
    t_eval = doc.add_table(rows=5, cols=4)
    for c_i, h in enumerate(["Kriteria Evaluasi", "Hasil Perhitungan", "Standar Kelayakan", "Status"]):
        t_eval.rows[0].cells[c_i].text = h
    eval_rows = [
        ("Net Present Value (NPV)", "Rp 2.809.117.669", "> 0 (Bernilai Positif)", "LAYAK"),
        ("Internal Rate of Return (IRR)", "297.98%", "> 20.00% (Opportunity Cost)", "LAYAK"),
        ("Payback Period (PP)", "4.9 Bulan (< 1 Tahun)", "< 3.0 Tahun", "LAYAK"),
        ("Profitability Index (PI)", "29.58", "> 1.20", "LAYAK")
    ]
    for r_i, r_vals in enumerate(eval_rows, start=1):
        for c_i, val in enumerate(r_vals):
            t_eval.rows[r_i].cells[c_i].text = val
            if c_i in [1, 2]:
                t_eval.rows[r_i].cells[c_i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            elif c_i == 3:
                t_eval.rows[r_i].cells[c_i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    format_table(t_eval)

    # ------------------------------------------------------------------
    # BAB VI: KESIMPULAN
    # ------------------------------------------------------------------
    doc.add_page_break()
    add_styled_heading(doc, "BAB VI\nKESIMPULAN", level=1)
    
    add_body_p(doc, f"Demikian laporan Studi Kelayakan Bisnis {cfg['business_name']} disusun secara komprehensif. Berdasarkan analisis aspek pasar, manajemen, teknis operasi, dan keuangan, diperoleh kesimpulan kelayakan sebagai berikut:")
    
    add_table_title(doc, f"Tabel 6.1 Kesimpulan Kelayakan {cfg['business_name']}")
    t_sum = doc.add_table(rows=5, cols=4)
    for c_i, h in enumerate(["Kriteria", "Hasil", "Target", "Kelayakan"]):
        t_sum.rows[0].cells[c_i].text = h
    for r_i, r_vals in enumerate(eval_rows, start=1):
        for c_i, val in enumerate(r_vals):
            t_sum.rows[r_i].cells[c_i].text = val
            t_sum.rows[r_i].cells[c_i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    format_table(t_sum)

    add_body_p(doc, f"Maka, usaha {cfg['business_name']} dinyatakan SANGAT LAYAK untuk dijalankan karena telah melampaui seluruh kualifikasi standar kelayakan investasi.")

    add_styled_heading(doc, "7.1 Kesimpulan dan Saran", level=2)
    add_body_p(doc, f"Berdasarkan hasil analisis dan kajian yang mendalam, bisnis {cfg['business_name']} memiliki potensi pasar yang sangat menjanjikan dengan dukungan lokasi strategis di sekitar area kampus dan pemukiman mahasiswa. Kekuatan diferensiasi produk dan manajemen operasional yang efisien memberikan margin keuntungan yang kuat serta kemampuan pengembalian investasi yang sangat cepat.")
    add_body_p(doc, "Untuk memaksimalkan keberhasilan jangka panjang, disarankan beberapa langkah strategis:")
    add_body_p(doc, "1. Menjaga konsistensi kualitas rasa produk dan higienitas melalui SOP yang ketat.", indent=1)
    add_body_p(doc, "2. Mengoptimalkan strategi promosi digital dan program loyalitas pelanggan bagi mahasiswa.", indent=1)
    add_body_p(doc, "3. Melakukan pengelolaan arus kas secara disiplin dan menyiapkan dana cadangan operasional.", indent=1)

    doc.save(output_path)
    print(f"[SUCCESS] Laporan SKB (.docx) berhasil dibuat di: {output_path}")

if __name__ == "__main__":
    cfg_file = sys.argv[1] if len(sys.argv) > 1 else r"d:\Perkuliahan\Kelass\SKB\.agents\skills\skb-format-writer\templates\sample_business_config.json"
    out_file = sys.argv[2] if len(sys.argv) > 2 else r"d:\Perkuliahan\Kelass\SKB\Laporan_SKB_Template.docx"
    build_skb_document(cfg_file, out_file)
