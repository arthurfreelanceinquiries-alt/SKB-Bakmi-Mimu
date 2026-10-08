import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Circle, FancyArrowPatch
import numpy as np
from PIL import Image, ImageOps, ImageDraw, ImageFont

base_dir = r"d:\Perkuliahan\Kelass\SKB"
out_dir = os.path.join(base_dir, "04_INTERNAL_TOOLS_DAN_DATA", "03_data_dan_aset")
os.makedirs(out_dir, exist_ok=True)

# Set Global Font Family to Times New Roman if available
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif', 'Liberation Serif', 'serif']
plt.rcParams['mathtext.fontset'] = 'stix'

# Palette
NAVY_DEEP = "#0F294D"
NAVY_LIGHT = "#1E3A8A"
GOLD_AMBER = "#C27803"
GOLD_LIGHT = "#FEF3C7"
GREEN_EMERALD = "#059669"
GREEN_LIGHT = "#ECFDF5"
SLATE_DARK = "#0F172A"
SLATE_MUTED = "#475569"
CARD_BG = "#FFFFFF"
BORDER_GRAY = "#CBD5E1"
BG_LIGHT = "#FAFBFD"

# =============================================================================
# 1. GAMBAR 3.1: STRUKTUR ORGANISASI BAKMI MIMU CARINA SAYANG (ASPECT EQUAL)
# =============================================================================
def generate_struktur_organisasi():
    fig, ax = plt.subplots(figsize=(14, 7.5), dpi=300)
    fig.patch.set_facecolor(BG_LIGHT)
    ax.set_facecolor(BG_LIGHT)
    ax.set_xlim(0, 140)
    ax.set_ylim(0, 75)
    ax.set_aspect('equal')
    ax.axis('off')

    # Title & Subtitle
    ax.text(70, 71, "STRUKTUR ORGANISASI KEDAI BAKMI MIMU CARINA SAYANG", 
            ha='center', va='center', fontsize=15, fontweight='bold', color=NAVY_DEEP)
    ax.text(70, 67, "Model Organisasi Fungsional Berbasis Efisiensi Operasional & Pengendalian Resep 1999", 
            ha='center', va='center', fontsize=10.5, style='italic', color=SLATE_MUTED)

    # Level 1: Owner Box
    owner_w, owner_h = 48, 16
    owner_x = 70 - owner_w / 2
    owner_y = 47
    rect_owner = FancyBboxPatch((owner_x, owner_y), owner_w, owner_h,
                                boxstyle="round,pad=0.3,rounding_size=1.0",
                                facecolor=CARD_BG, edgecolor=GOLD_AMBER, linewidth=2.2)
    ax.add_patch(rect_owner)
    
    # Header strip inside Owner Box
    strip_owner = FancyBboxPatch((owner_x, owner_y + owner_h - 4.0), owner_w, 4.0,
                                 boxstyle="round,pad=0.2,rounding_size=0.8",
                                 facecolor=GOLD_LIGHT, edgecolor=GOLD_AMBER, linewidth=1.0)
    ax.add_patch(strip_owner)
    ax.text(70, owner_y + owner_h - 2.0, "OWNER / PENGELOLA UTAMA", 
            ha='center', va='center', fontsize=11.5, fontweight='bold', color=GOLD_AMBER)
    
    owner_desc = (
        "• Pengambil Kebijakan Strategis & Modal Investasi\n"
        "• Pengawasan Keuangan & Pembukuan Kas Harian\n"
        "• Kontrol SOP Standar Cita Rasa Resep 1999\n"
        "• Pengadaan Bahan Baku Daging Segar & Hubungan RT/RW\n"
        "• Perencanaan Peluang Ekspansi Cabang 2027/2028"
    )
    ax.text(owner_x + 2.5, owner_y + 5.8, owner_desc, ha='left', va='center', fontsize=8.6, color=SLATE_DARK)

    # Connector lines
    # Vertical line down from Owner
    ax.plot([70, 70], [owner_y, 40], color=NAVY_DEEP, linewidth=2.0)
    # Horizontal distributor bar
    ax.plot([24, 116], [40, 40], color=NAVY_DEEP, linewidth=2.0)
    # Drop lines to each employee
    ax.plot([24, 24], [40, 34], color=NAVY_DEEP, linewidth=2.0)
    ax.plot([70, 70], [40, 34], color=NAVY_DEEP, linewidth=2.0)
    ax.plot([116, 116], [40, 34], color=NAVY_DEEP, linewidth=2.0)

    # Level 2: 3 Subordinates
    sub_w, sub_h = 38, 25
    subordinates = [
        (24 - sub_w / 2, "KOKI UTAMA (1 ORANG)", "Gaji: Rp 3.800.000 / Bulan", NAVY_LIGHT, [
            "Tanggung Jawab Teknis:",
            "• Pembuatan adonan mie fresh mingguan",
            "• Menjaga tekstur kenyal tanpa pengawet",
            "• Memasak 3 topping: ayam putih gurih,",
            "  babi kecap manis, & casiu madu panggang",
            "• Menyiapkan 3 racikan minyak & kaldu",
            "• Memimpin perebusan mie kilat 45 detik",
            "  di gerobak stasiun open kitchen teras"
        ]),
        (70 - sub_w / 2, "ASISTEN KOKI (1 ORANG)", "Gaji: Rp 3.200.000 / Bulan", GREEN_EMERALD, [
            "Tanggung Jawab Pendukung:",
            "• Menyiapkan bahan mentah & bumbu dapur",
            "• Melipat kulit pangsit & swikiaw babi-udang",
            "• Merebus baso sapi, baso ikan, & sayuran",
            "• Pengisian ulang air dispenser & kaldu",
            "• Mencuci mangkok, piring, sumpit & panci",
            "• Menjaga sanitasi higienis teras kedai",
            "• Standar operasional kebersihan tertutup"
        ]),
        (116 - sub_w / 2, "KASIR & PRAMUSAJI (1 ORANG)", "Gaji: Rp 3.000.000 / Bulan", GOLD_AMBER, [
            "Tanggung Jawab Layanan:",
            "• Mencatat pesanan tamu dengan nota fisik",
            "• Melayani transaksi kas tunai & QRIS",
            "• Menyajikan mangkok mie ke 23 kursi",
            "• Menyajikan minuman segar & saus meja",
            "• Membersihkan & menata ulang meja santap",
            "• Melayani pemesanan bawa pulang (take-away)",
            "• Rekonsiliasi fisik uang kas di akhir shift"
        ])
    ]

    for sx, title, sal, color, tasks in subordinates:
        sy = 9
        card = FancyBboxPatch((sx, sy), sub_w, sub_h,
                              boxstyle="round,pad=0.3,rounding_size=1.0",
                              facecolor=CARD_BG, edgecolor=color, linewidth=1.8)
        ax.add_patch(card)

        # Header strip
        strip = FancyBboxPatch((sx, sy + sub_h - 3.8), sub_w, 3.8,
                               boxstyle="round,pad=0.2,rounding_size=0.8",
                               facecolor=CARD_BG, edgecolor=color, linewidth=1.0)
        ax.add_patch(strip)
        ax.text(sx + sub_w / 2, sy + sub_h - 1.9, title, 
                ha='center', va='center', fontsize=9.8, fontweight='bold', color=color)
        
        # Salary badge
        ax.text(sx + sub_w / 2, sy + sub_h - 5.4, sal, 
                ha='center', va='center', fontsize=8.6, fontweight='bold', color=SLATE_MUTED)

        # Divider line inside card
        ax.plot([sx + 1.5, sx + sub_w - 1.5], [sy + sub_h - 7.0, sy + sub_h - 7.0], color=BORDER_GRAY, linewidth=0.8)

        # Tasks
        task_text = "\n".join(tasks)
        ax.text(sx + 2.0, sy + 9.5, task_text, ha='left', va='center', fontsize=8.0, color=SLATE_DARK)

    # Bottom summary footnote
    footnote = (
        "Total Formasi SDM: 1 Owner (Pengelola) + 3 Staf Karyawan Tetap  |  Total Alokasi Upah: Rp 10.000.000 / Bulan (Rp 120.000.000 / Tahun)\n"
        "Kebijakan Kompensasi: Gaji Pokok Bulanan Pasti, Fasilitas Makan Harian Kedai, & Tunjangan Hari Raya (THR) Tahunan"
    )
    ax.text(70, 3.5, footnote, ha='center', va='center', fontsize=8.8, fontweight='bold', color=NAVY_DEEP)

    out_file = os.path.join(out_dir, "struktur_organisasi_bakmi_mimu.png")
    plt.tight_layout()
    plt.savefig(out_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[GENERATED] Gambar 3.1: {out_file}")

# =============================================================================
# 2. GAMBAR 4.2: NETWORK PLANNING & CRITICAL PATH METHOD (ASPECT EQUAL, 100% BULAT)
# =============================================================================
def generate_network_planning():
    fig, ax = plt.subplots(figsize=(14, 7.5), dpi=300)
    fig.patch.set_facecolor(BG_LIGHT)
    ax.set_facecolor(BG_LIGHT)
    ax.set_xlim(0, 140)
    ax.set_ylim(0, 75)
    ax.set_aspect('equal')
    ax.axis('off')

    # Header
    ax.text(70, 71, "NETWORK PLANNING & CRITICAL PATH METHOD (CPM) RELOKASI 2016", 
            ha='center', va='center', fontsize=15, fontweight='bold', color=NAVY_DEEP)
    ax.text(70, 67, "Diagram Alur Ketergantungan Aktivitas & Jalur Kritis Pelaksanaan Proyek (Total Durasi: 12 Minggu)", 
            ha='center', va='center', fontsize=10.5, style='italic', color=SLATE_MUTED)

    # Nodes Definitions: (id, x, y, ES, LF)
    nodes = {
        1: (12, 48, 0, 0),
        2: (26, 48, 2, 2),
        3: (42, 58, 4, 4),
        4: (42, 38, 4, 6),
        5: (58, 58, 5, 5),
        6: (74, 58, 7, 7),
        7: (74, 38, 6, 7),
        8: (90, 58, 9, 9),
        9: (90, 38, 8, 9),
        10: (106, 58, 10, 10),
        11: (124, 48, 12, 12)
    }

    # Edges Definitions: (from_node, to_node, code, name, dur, is_critical)
    edges = [
        (1, 2, "A", "Survei Lokasi", 2, True),
        (2, 3, "B", "Sewa Ruko 1 Thn", 2, True),
        (2, 4, "C", "Izin RT/RW & Retribusi", 2, False),
        (4, 7, "D", "Desain Tata Letak", 1, False),
        (3, 5, "E", "Pengecatan & Saluran Air", 1, True),
        (5, 6, "G", "Pembuatan Gerobak Open Kitchen", 2, True),
        (5, 7, "F", "Pasang 6 Kipas Angin", 1, False),
        (7, 8, "Dummy", "", 0, False),
        (6, 8, "I", "Pengadaan Meja & Kursi (37 pcs)", 2, True),
        (6, 9, "H", "Pasang Plang Signage Nama", 1, False),
        (9, 10, "J", "Beli Kulkas, Freezer & Panci", 1, False),
        (8, 10, "K", "Setting Tata Letak 23 Kursi", 1, True),
        (10, 11, "M-P", "Rekrutmen, SOP, Trial & Buka", 2, True),
        (4, 11, "L", "Stok Saus & Bahan Mentah", 1, False)
    ]

    r = 4.2  # EXACT CIRCLE RADIUS (1.0000 ASPECT)

    # Draw Edges (Arrows)
    for u, v, code, name, dur, is_crit in edges:
        x1, y1, _, _ = nodes[u]
        x2, y2, _, _ = nodes[v]
        
        dx = x2 - x1
        dy = y2 - y1
        dist = np.hypot(dx, dy)
        if dist == 0: continue
        
        # Stop exactly at circle boundary
        sx = x1 + (dx / dist) * r
        sy = y1 + (dy / dist) * r
        ex = x2 - (dx / dist) * r
        ey = y2 - (dy / dist) * r

        color = GOLD_AMBER if is_crit else NAVY_LIGHT
        lwidth = 2.4 if is_crit else 1.3
        linestyle = '-' if dur > 0 else '--'

        arrow = FancyArrowPatch((sx, sy), (ex, ey),
                                arrowstyle='-|>',
                                mutation_scale=12 if is_crit else 9,
                                color=color,
                                linewidth=lwidth,
                                linestyle=linestyle,
                                zorder=3)
        ax.add_patch(arrow)

        # Label position
        if dur > 0:
            mx = (sx + ex) / 2
            my = (sy + ey) / 2
            if dy == 0:
                ly = my + (2.6 if is_crit else -2.6)
                lx = mx
            elif dy > 0:
                lx = mx - 2.5
                ly = my + 1.8
            else:
                lx = mx - 2.5
                ly = my - 1.8
            
            lbl_txt = f"{code} ({dur} mg)"
            ax.text(lx, ly, lbl_txt, ha='center', va='center', 
                    fontsize=8.0, fontweight='bold', color=color, zorder=5)

    # Draw Nodes (TRUE Circles with 3 sections: Node ID, ES, LF)
    for nid, (nx, ny, es, lf) in nodes.items():
        # Outer Node Circle - 100% Round
        circ = patches.Circle((nx, ny), r, edgecolor=NAVY_DEEP, facecolor=CARD_BG, linewidth=1.8, zorder=6)
        ax.add_patch(circ)

        # Horizontal divider inside node
        ax.plot([nx - r, nx + r], [ny, ny], color=NAVY_DEEP, linewidth=1.0, zorder=7)
        # Vertical divider on bottom half
        ax.plot([nx, nx], [ny - r, ny], color=NAVY_DEEP, linewidth=1.0, zorder=7)

        # Node ID top center
        ax.text(nx, ny + 1.8, str(nid), ha='center', va='center', fontsize=9.5, fontweight='bold', color=NAVY_DEEP, zorder=8)
        # ES bottom left
        ax.text(nx - 2.0, ny - 2.0, str(es), ha='center', va='center', fontsize=7.2, color=SLATE_MUTED, zorder=8)
        # LF bottom right
        ax.text(nx + 2.0, ny - 2.0, str(lf), ha='center', va='center', fontsize=7.2, color=GOLD_AMBER if es == lf else SLATE_MUTED, fontweight='bold' if es == lf else 'normal', zorder=8)

    # Critical Path Indicator Box
    crit_box = FancyBboxPatch((10, 10), 120, 16, boxstyle="round,pad=0.3,rounding_size=0.8",
                              facecolor=GOLD_LIGHT, edgecolor=GOLD_AMBER, linewidth=1.5, zorder=4)
    ax.add_patch(crit_box)

    ax.text(12, 23.5, "KETERANGAN SIMPUL & JALUR KRITIS (CRITICAL PATH METHOD):", fontsize=9.2, fontweight='bold', color=GOLD_AMBER)
    c_desc = (
        "• Jalur Kritis (Tebal Emas): 1 -> 2 -> 3 -> 5 -> 6 -> 8 -> 10 -> 11  |  Total Durasi Waktu Proyek: 12 Minggu\n"
        "• Urutan Aktivitas Kritis: A (Survei 2 mg) -> B (Sewa Ruko 2 mg) -> E (Renovasi Cat/Pipa 1 mg) -> G (Gerobak Teras 2 mg) ->\n"
        "  I (Meja/Kursi 2 mg) -> K (Setting 23 Kursi 1 mg) -> M-P (Rekrutmen 3 Staf, SOP Adonan, Trial Run & Grand Opening 2 mg)\n"
        "• Keterangan Kotak Node: Bagian Atas = Nomor Simpul; Kiri Bawah = Earliest Start (ES); Kanan Bawah = Latest Finish (LF)"
    )
    ax.text(12, 16.5, c_desc, fontsize=8.2, color=SLATE_DARK, va='center')

    # Bottom Footnote
    ax.text(70, 4.5, "Diagram Network Planning 100% Selaras dengan Gantt Chart 12 Minggu pada BAB IV Subbab 4.3 Dokumen Laporan",
            ha='center', va='center', fontsize=8.8, fontweight='bold', color=NAVY_DEEP)

    out_file = os.path.join(out_dir, "network_planning_bakmi_mimu.png")
    plt.tight_layout()
    plt.savefig(out_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[GENERATED] Gambar 4.2: {out_file}")

# =============================================================================
# 3. GAMBAR 1.1: KOLEKSI FOTO PRODUK ASLI (KOMPOSIT TRIPTYCH)
# =============================================================================
def generate_gambaran_produk_real():
    f_dine = os.path.join(out_dir, 'foto_mie_dinein_swikiaw.jpg')
    f_take = os.path.join(out_dir, 'foto_mie_spesial_3topping.jpg')
    f_sumpit = os.path.join(out_dir, 'foto_tekstur_mie_lebar.jpg')

    W_PROD, H_PROD = 2400, 1200
    img_prod = Image.new('RGB', (W_PROD, H_PROD), color=BG_LIGHT)
    draw = ImageDraw.Draw(img_prod)

    try:
        font_title = ImageFont.truetype('timesbd.ttf', 44)
        font_sub = ImageFont.truetype('timesi.ttf', 26)
        font_lbl = ImageFont.truetype('timesbd.ttf', 26)
        font_sublbl = ImageFont.truetype('times.ttf', 22)
    except Exception:
        font_title = font_sub = font_lbl = font_sublbl = ImageFont.load_default()

    draw.text((W_PROD//2, 45), 'DOKUMENTASI PRODUK OTENTIK BAKMI MIMU CARINA SAYANG', fill='#0F294D', font=font_title, anchor='mm')
    draw.text((W_PROD//2, 90), 'Dokumentasi Menu Lapangan: Mie Spesial 3 Topping, Penyajian Dine-In, & Tekstur Adonan Mie Lebar 1999', fill='#475569', font=font_sub, anchor='mm')

    panel_w, panel_h = 720, 960
    top_y = 140
    gap = 40
    start_x = (W_PROD - (3 * panel_w + 2 * gap)) // 2

    panels_data = [
        (f_take, '(a) Mie Spesial 3 Topping Murni', 'Ayam Putih, Babi Cincang Kecap, & Casiu Madu'),
        (f_dine, '(b) Penyajian Santap Dine-In', 'Mie Campur, Swikiaw Kuah, & Baso Goreng'),
        (f_sumpit, '(c) Kualitas Adonan Mie Lebar', 'Tekstur Kenyal Autentik Fresh Mingguan')
    ]

    for idx, (path, title_p, desc_p) in enumerate(panels_data):
        px = start_x + idx * (panel_w + gap)
        py = top_y
        
        draw.rectangle([px, py, px + panel_w, py + panel_h], fill='#FFFFFF', outline='#CBD5E1', width=2)
        
        raw_im = Image.open(path)
        target_img_w = panel_w - 20
        target_img_h = 800
        
        im_fitted = ImageOps.fit(raw_im, (target_img_w, target_img_h), method=Image.Resampling.LANCZOS)
        img_prod.paste(im_fitted, (px + 10, py + 10))
        draw.rectangle([px + 10, py + 10, px + 10 + target_img_w, py + 10 + target_img_h], outline='#E2E8F0', width=1)
        
        draw.text((px + panel_w//2, py + 855), title_p, fill='#0F294D', font=font_lbl, anchor='mm')
        draw.text((px + panel_w//2, py + 900), desc_p, fill='#64748B', font=font_sublbl, anchor='mm')

    out_prod = os.path.join(out_dir, 'produk_bakmi_mimu.png')
    img_prod.save(out_prod, 'PNG', quality=95)
    print(f"[GENERATED] Gambar 1.1 Real: {out_prod}")

# =============================================================================
# 4. GAMBAR 2.1: LOKASI & FASAD KEDAI ASLI (KOMPOSIT GOOGLE MAPS + FASAD RUKO)
# =============================================================================
def generate_peta_lokasi_real():
    f_map = os.path.join(out_dir, 'foto_peta_google_maps.jpg')
    f_ruko = os.path.join(out_dir, 'foto_fasad_kedai_ruko.jpg')

    W_MAP, H_MAP = 2200, 1100
    img_map = Image.new('RGB', (W_MAP, H_MAP), color=BG_LIGHT)
    draw_m = ImageDraw.Draw(img_map)

    try:
        font_title = ImageFont.truetype('timesbd.ttf', 44)
        font_sub = ImageFont.truetype('timesi.ttf', 26)
        font_lbl = ImageFont.truetype('timesbd.ttf', 26)
        font_sublbl = ImageFont.truetype('times.ttf', 22)
    except Exception:
        font_title = font_sub = font_lbl = font_sublbl = ImageFont.load_default()

    draw_m.text((W_MAP//2, 45), 'DOKUMENTASI LOKASI & KEDAI FISIK BAKMI MIMU CARINA SAYANG', fill='#0F294D', font=font_title, anchor='mm')
    draw_m.text((W_MAP//2, 90), 'Peta Navigasi Google Maps & Fasad Ruko Kedai di Jl. Angsoka Hijau IV Duri Kosambi (Depan Kalam Kudus)', fill='#475569', font=font_sub, anchor='mm')

    m_panel_w1 = 1120
    m_panel_w2 = 880
    m_panel_h = 860
    m_top_y = 140
    m_gap = 40
    m_start_x = (W_MAP - (m_panel_w1 + m_panel_w2 + m_gap)) // 2

    # Panel 1: Map
    raw_map = Image.open(f_map)
    fit_map = ImageOps.fit(raw_map, (m_panel_w1 - 20, m_panel_h - 100), method=Image.Resampling.LANCZOS)
    draw_m.rectangle([m_start_x, m_top_y, m_start_x + m_panel_w1, m_top_y + m_panel_h], fill='#FFFFFF', outline='#CBD5E1', width=2)
    img_map.paste(fit_map, (m_start_x + 10, m_top_y + 10))
    draw_m.rectangle([m_start_x + 10, m_top_y + 10, m_start_x + m_panel_w1 - 10, m_top_y + m_panel_h - 90], outline='#E2E8F0', width=1)
    draw_m.text((m_start_x + m_panel_w1//2, m_top_y + m_panel_h - 55), '(a) Peta Navigasi Google Maps Titik Kedai Bakmi Mimu', fill='#0F294D', font=font_lbl, anchor='mm')
    draw_m.text((m_start_x + m_panel_w1//2, m_top_y + m_panel_h - 25), 'Jl. Angsoka Hijau IV Blok E6 No. 17, Duri Kosambi (Catchment Area Sekolah & Residensial)', fill='#64748B', font=font_sublbl, anchor='mm')

    # Panel 2: Fasad Ruko
    p2_x = m_start_x + m_panel_w1 + m_gap
    raw_ruko = Image.open(f_ruko)
    fit_ruko = ImageOps.fit(raw_ruko, (m_panel_w2 - 20, m_panel_h - 100), method=Image.Resampling.LANCZOS)
    draw_m.rectangle([p2_x, m_top_y, p2_x + m_panel_w2, m_top_y + m_panel_h], fill='#FFFFFF', outline='#CBD5E1', width=2)
    img_map.paste(fit_ruko, (p2_x + 10, m_top_y + 10))
    draw_m.rectangle([p2_x + 10, m_top_y + 10, p2_x + m_panel_w2 - 10, m_top_y + m_panel_h - 90], outline='#E2E8F0', width=1)
    draw_m.text((p2_x + m_panel_w2//2, m_top_y + m_panel_h - 55), '(b) Fasad Kedai & Dapur Teras (Open Kitchen)', fill='#0F294D', font=font_lbl, anchor='mm')
    draw_m.text((p2_x + m_panel_w2//2, m_top_y + m_panel_h - 25), 'Spanduk Menu, Gerobak Etalase Racik, & Antrean Pengunjung', fill='#64748B', font=font_sublbl, anchor='mm')

    out_map = os.path.join(out_dir, 'peta_lokasi_bakmi_mimu.png')
    img_map.save(out_map, 'PNG', quality=95)
    print(f"[GENERATED] Gambar 2.1 Real: {out_map}")

if __name__ == "__main__":
    print("=" * 60)
    print("MENGHASILKAN ASET VISUAL DENGAN PROPORSI ISOMETRIK & FOTO ASLI")
    print("=" * 60)
    generate_struktur_organisasi()
    generate_network_planning()
    generate_gambaran_produk_real()
    generate_peta_lokasi_real()
    print("=" * 60)
    print("SELURUH 4 ASET VISUAL BERHASIL DIGENERATE TANPA DISTORSI (GEPENG)")
    print("=" * 60)
