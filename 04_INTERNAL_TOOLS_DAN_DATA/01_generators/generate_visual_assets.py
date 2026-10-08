import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Circle, Arrow, FancyArrowPatch
import numpy as np

base_dir = r"d:\Perkuliahan\Kelass\SKB"
out_dir = os.path.join(base_dir, "04_INTERNAL_TOOLS_DAN_DATA", "03_data_dan_aset")
os.makedirs(out_dir, exist_ok=True)

# Set Global Font Family to Times New Roman if available, with serif fallback
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
# 1. GAMBAR 3.1: STRUKTUR ORGANISASI BAKMI MIMU CARINA SAYANG
# =============================================================================
def generate_struktur_organisasi():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    fig.patch.set_facecolor(BG_LIGHT)
    ax.set_facecolor(BG_LIGHT)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Title & Subtitle
    ax.text(50, 94, "STRUKTUR ORGANISASI KEDAI BAKMI MIMU CARINA SAYANG", 
            ha='center', va='center', fontsize=16, fontweight='bold', color=NAVY_DEEP)
    ax.text(50, 90, "Model Organisasi Fungsional Berbasis Efisiensi Operasional & Pengendalian Resep 1999", 
            ha='center', va='center', fontsize=11, style='italic', color=SLATE_MUTED)

    # Level 1: Owner Box
    owner_w, owner_h = 36, 18
    owner_x = 50 - owner_w / 2
    owner_y = 66
    rect_owner = FancyBboxPatch((owner_x, owner_y), owner_w, owner_h,
                                boxstyle="round,pad=0.5,rounding_size=1.5",
                                facecolor=CARD_BG, edgecolor=GOLD_AMBER, linewidth=2.5)
    ax.add_patch(rect_owner)
    
    # Header strip inside Owner Box
    strip_owner = FancyBboxPatch((owner_x, owner_y + owner_h - 4.5), owner_w, 4.5,
                                 boxstyle="round,pad=0.2,rounding_size=1.0",
                                 facecolor=GOLD_LIGHT, edgecolor=GOLD_AMBER, linewidth=1.0)
    ax.add_patch(strip_owner)
    ax.text(50, owner_y + owner_h - 2.3, "OWNER / PENGELOLA UTAMA", 
            ha='center', va='center', fontsize=12, fontweight='bold', color=GOLD_AMBER)
    
    owner_desc = (
        "• Pengambil Kebijakan Strategis & Modal Investasi\n"
        "• Pengawasan Keuangan & Pembukuan Kas Harian\n"
        "• Kontrol SOP Standar Cita Rasa Resep 1999\n"
        "• Pengadaan Bahan Baku Daging Segar & Hubungan RT/RW\n"
        "• Perencanaan Peluang Ekspansi Cabang 2027/2028"
    )
    ax.text(owner_x + 2, owner_y + 6.5, owner_desc, ha='left', va='center', fontsize=9.2, color=SLATE_DARK)

    # Connector lines
    # Vertical line down from Owner
    ax.plot([50, 50], [owner_y, 49], color=NAVY_DEEP, linewidth=2.2)
    # Horizontal distributor bar
    ax.plot([19, 81], [49, 49], color=NAVY_DEEP, linewidth=2.2)
    # Drop lines to each employee
    ax.plot([19, 19], [49, 42], color=NAVY_DEEP, linewidth=2.2)
    ax.plot([50, 50], [49, 42], color=NAVY_DEEP, linewidth=2.2)
    ax.plot([81, 81], [49, 42], color=NAVY_DEEP, linewidth=2.2)

    # Level 2: 3 Subordinates
    sub_w, sub_h = 28, 30
    subordinates = [
        (19 - sub_w / 2, "KOKI UTAMA (1 ORANG)", "Gaji: Rp 3.800.000 / Bulan", NAVY_LIGHT, [
            "Tanggung Jawab Teknis:",
            "• Pembuatan adonan mie fresh mingguan",
            "• Menjaga tekstur kenyal tanpa pengawet",
            "• Memasak 3 topping: ayam putih gurih,",
            "  babi kecap manis, & casiu madu panggang",
            "• Menyiapkan 3 racikan minyak & kaldu",
            "• Memimpin perebusan mie kilat 45 detik",
            "  di gerobak stasiun open kitchen teras"
        ]),
        (50 - sub_w / 2, "ASISTEN KOKI (1 ORANG)", "Gaji: Rp 3.200.000 / Bulan", GREEN_EMERALD, [
            "Tanggung Jawab Pendukung:",
            "• Menyiapkan bahan mentah & bumbu dapur",
            "• Melipat kulit pangsit & swikiaw babi-udang",
            "• Merebus baso sapi, baso ikan, & sayuran",
            "• Pengisian ulang air dispenser & kaldu",
            "• Mencuci mangkok, piring, sumpit & panci",
            "• Menjaga sanitasi higienis teras kedai",
            "• Standar operasional kebersihan tertutup"
        ]),
        (81 - sub_w / 2, "KASIR & PRAMUSAJI (1 ORANG)", "Gaji: Rp 3.000.000 / Bulan", GOLD_AMBER, [
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
        sy = 12
        card = FancyBboxPatch((sx, sy), sub_w, sub_h,
                              boxstyle="round,pad=0.5,rounding_size=1.5",
                              facecolor=CARD_BG, edgecolor=color, linewidth=2.0)
        ax.add_patch(card)

        # Header strip
        strip = FancyBboxPatch((sx, sy + sub_h - 4.5), sub_w, 4.5,
                               boxstyle="round,pad=0.2,rounding_size=1.0",
                               facecolor=CARD_BG, edgecolor=color, linewidth=1.0)
        ax.add_patch(strip)
        ax.text(sx + sub_w / 2, sy + sub_h - 2.3, title, 
                ha='center', va='center', fontsize=10.5, fontweight='bold', color=color)
        
        # Salary badge
        ax.text(sx + sub_w / 2, sy + sub_h - 6.2, sal, 
                ha='center', va='center', fontsize=9.2, fontweight='bold', color=SLATE_MUTED)

        # Divider line inside card
        ax.plot([sx + 1.5, sx + sub_w - 1.5], [sy + sub_h - 8, sy + sub_h - 8], color=BORDER_GRAY, linewidth=1.0)

        # Tasks
        task_text = "\n".join(tasks)
        ax.text(sx + 1.8, sy + 11.5, task_text, ha='left', va='center', fontsize=8.6, color=SLATE_DARK)

    # Bottom summary footnote
    footnote = (
        "Total Formasi SDM: 1 Owner (Pengelola) + 3 Staf Karyawan Tetap  |  Total Alokasi Upah: Rp 10.000.000 / Bulan (Rp 120.000.000 / Tahun)\n"
        "Kebijakan Kompensasi: Gaji Pokok Bulanan Pasti, Fasilitas Makan Harian Kedai, & Tunjangan Hari Raya (THR) Tahunan"
    )
    ax.text(50, 4, footnote, ha='center', va='center', fontsize=9.5, fontweight='bold', color=NAVY_DEEP)

    out_file = os.path.join(out_dir, "struktur_organisasi_bakmi_mimu.png")
    plt.tight_layout()
    plt.savefig(out_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[GENERATED] Gambar 3.1: {out_file}")

# =============================================================================
# 2. GAMBAR 2.1: PETA LOKASI KEDAI & ZONASI PASAR BAKMI MIMU
# =============================================================================
def generate_peta_lokasi():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    fig.patch.set_facecolor(BG_LIGHT)
    ax.set_facecolor("#F8FAFC")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Title & Header
    ax.text(50, 95, "PETA LOKASI & ANALISIS TANGKAPAN PASAR BAKMI MIMU CARINA SAYANG", 
            ha='center', va='center', fontsize=15, fontweight='bold', color=NAVY_DEEP)
    ax.text(50, 91.5, "Jl. Angsoka Hijau IV Blok E6 No. 17, Duri Kosambi, Cengkareng, Jakarta Barat (Persis Depan Sekolah Kalam Kudus)", 
            ha='center', va='center', fontsize=10.5, style='italic', color=SLATE_MUTED)

    # Draw Map Canvas Container
    map_box = FancyBboxPatch((4, 8), 92, 80, boxstyle="round,pad=0.5,rounding_size=1.0",
                             facecolor="#F1F5F9", edgecolor=BORDER_GRAY, linewidth=1.5)
    ax.add_patch(map_box)

    # Draw Road Network (Major & Local Roads)
    # 1. Jl. Lingkar Luar Barat (Highway on East / Right)
    ax.plot([82, 82], [10, 86], color="#CBD5E1", linewidth=12, zorder=2)
    ax.plot([82, 82], [10, 86], color="#94A3B8", linewidth=1.5, linestyle="--", zorder=3)
    ax.text(84.5, 50, "Jl. Lingkar Luar Barat (JORR W1)", rotation=90, va='center', fontsize=8.5, fontweight='bold', color=SLATE_MUTED)

    # 2. Jl. Duri Kosambi Raya (Main Arterial West to East)
    ax.plot([6, 82], [42, 42], color="#CBD5E1", linewidth=14, zorder=2)
    ax.text(32, 44, "Jl. Duri Kosambi Raya (Arteri Utama Kawasan Kosambi)", va='bottom', fontsize=9, fontweight='bold', color=SLATE_MUTED)

    # 3. Jl. Kresek Raya (Diagonal North to West)
    ax.plot([6, 45], [75, 42], color="#E2E8F0", linewidth=9, zorder=2)
    ax.text(18, 62, "Jl. Kresek Raya", rotation=-35, va='center', fontsize=8, color=SLATE_MUTED)

    # 4. Jl. Semanan Raya (North Axis)
    ax.plot([45, 45], [42, 86], color="#E2E8F0", linewidth=9, zorder=2)
    ax.text(46.5, 78, "Jl. Raya Semanan", rotation=90, va='center', fontsize=8, color=SLATE_MUTED)

    # 5. Jl. Angsoka Hijau Raya & Jl. Angsoka Hijau IV (Local Street Grid)
    ax.plot([30, 70], [30, 30], color="#FFFFFF", linewidth=8, zorder=3)
    ax.plot([52, 52], [12, 42], color="#FFFFFF", linewidth=9, zorder=3)
    ax.text(53.5, 36, "Jl. Angsoka Hijau Raya", rotation=90, va='center', fontsize=8, color=NAVY_DEEP)

    ax.plot([38, 66], [22, 22], color="#FFFFFF", linewidth=11, zorder=3)
    ax.text(48, 20, "Jl. Angsoka Hijau IV", ha='center', va='top', fontsize=8.5, fontweight='bold', color=NAVY_DEEP)

    # Landmark 1: Perumahan Duri Kosambi & Taman Semanan Indah (Polygons)
    res_1 = FancyBboxPatch((8, 14), 22, 22, boxstyle="round,pad=0.2", facecolor="#E0F2FE", edgecolor="#BAE6FD", linewidth=1.5, zorder=3)
    ax.add_patch(res_1)
    ax.text(19, 25, "PERUMAHAN\nDURI KOSAMBI\n(Residensial)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="#0369A1", zorder=4)

    res_2 = FancyBboxPatch((8, 50), 24, 20, boxstyle="round,pad=0.2", facecolor="#E0F2FE", edgecolor="#BAE6FD", linewidth=1.5, zorder=3)
    ax.add_patch(res_2)
    ax.text(20, 60, "KOMPLEKS PERUMAHAN\nTAMAN SEMANAN INDAH (TSI)\n(Potensi Pasar Siang/Malam)", ha='center', va='center', fontsize=8, fontweight='bold', color="#0369A1", zorder=4)

    # Landmark 2: Sekolah Kristen Kalam Kudus Kompleks
    kk_poly = FancyBboxPatch((40, 24), 24, 6, boxstyle="round,pad=0.2", facecolor="#DCFCE7", edgecolor=GREEN_EMERALD, linewidth=2.0, zorder=4)
    ax.add_patch(kk_poly)
    ax.text(52, 27, "TAMAN & SEKOLAH KRISTEN KALAM KUDUS (TK - SD - SMP)\n[Persis Berhadapan / Di Depan Kedai Bakmi Mimu]", 
            ha='center', va='center', fontsize=8.2, fontweight='bold', color="#166534", zorder=5)

    # Landmark 3: Titik Kedai Bakmi Mimu Carina Sayang (HERO PIN)
    mimu_poly = FancyBboxPatch((43, 14), 18, 6.5, boxstyle="round,pad=0.3", facecolor=GOLD_LIGHT, edgecolor=GOLD_AMBER, linewidth=2.5, zorder=6)
    ax.add_patch(mimu_poly)
    
    # Pin marker icon
    ax.plot(52, 18, marker='o', markersize=10, color=GOLD_AMBER, zorder=7)
    ax.plot(52, 18, marker='*', markersize=6, color="#FFFFFF", zorder=8)
    ax.text(52, 15.5, "[ KEDAI BAKMI MIMU CARINA SAYANG ]\n(Jl. Angsoka Hijau IV Blok E6 No. 17)", 
            ha='center', va='center', fontsize=8.2, fontweight='bold', color=NAVY_DEEP, zorder=9)

    # Landmark 4: Stasiun Rawa Buaya & Fasilitas Umum
    st_poly = FancyBboxPatch((64, 70), 16, 12, boxstyle="round,pad=0.2", facecolor="#FEF2F2", edgecolor="#FECACA", linewidth=1.5, zorder=3)
    ax.add_patch(st_poly)
    ax.text(72, 76, "STASIUN\nRAWA BUAYA\n(Akses Komuter)", ha='center', va='center', fontsize=8, fontweight='bold', color="#991B1B", zorder=4)

    # Market Catchment Radius Rings (Dashed Circles around Kedai)
    circle_1km = patches.Circle((52, 18), 12, edgecolor=GOLD_AMBER, facecolor="none", linestyle=":", linewidth=1.8, zorder=2)
    circle_3km = patches.Circle((52, 18), 24, edgecolor=NAVY_LIGHT, facecolor="none", linestyle=":", linewidth=1.5, zorder=2)
    circle_5km = patches.Circle((52, 18), 38, edgecolor=BORDER_GRAY, facecolor="none", linestyle=":", linewidth=1.2, zorder=2)
    ax.add_patch(circle_1km)
    ax.add_patch(circle_3km)
    ax.add_patch(circle_5km)

    # Radius labels
    ax.text(64.5, 18, "Radius 1 km (Pelanggan Harian)", fontsize=7.5, color=GOLD_AMBER, fontweight='bold', zorder=4)
    ax.text(76.5, 18, "Radius 3 km", fontsize=7.5, color=NAVY_LIGHT, fontweight='bold', zorder=4)
    ax.text(90.5, 18, "Radius 5 km", fontsize=7.5, color=SLATE_MUTED, zorder=4)

    # Map Controls: Compass / North Arrow
    ax.plot([90, 90], [78, 84], color=NAVY_DEEP, linewidth=2, zorder=5)
    ax.text(90, 85.5, "U", ha='center', fontsize=11, fontweight='bold', color=NAVY_DEEP, zorder=5)
    ax.text(90, 76.5, "UTARA", ha='center', fontsize=7.5, fontweight='bold', color=NAVY_DEEP, zorder=5)

    # Legend Card (Bottom Left)
    leg_box = FancyBboxPatch((6, 9.5), 32, 14, boxstyle="round,pad=0.3", facecolor=CARD_BG, edgecolor=BORDER_GRAY, linewidth=1.2, zorder=6)
    ax.add_patch(leg_box)
    ax.text(7, 21.5, "LEGENDA PETA & ZONASI PASAR:", fontsize=8.2, fontweight='bold', color=NAVY_DEEP, zorder=7)
    ax.plot(8.5, 19, marker='o', markersize=6, color=GOLD_AMBER, zorder=7)
    ax.text(10.5, 19, "Lokasi Kedai Bakmi Mimu Carina Sayang", va='center', fontsize=7.5, zorder=7)
    ax.plot(8.5, 16.5, marker='s', markersize=6, color=GREEN_EMERALD, zorder=7)
    ax.text(10.5, 16.5, "Sekolah Kristen Kalam Kudus (TK/SD/SMP)", va='center', fontsize=7.5, zorder=7)
    ax.plot(8.5, 14, marker='s', markersize=6, color="#0369A1", zorder=7)
    ax.text(10.5, 14, "Zona Permukiman Residensial (Pelanggan Utama)", va='center', fontsize=7.5, zorder=7)
    ax.plot([7, 10], [11.5, 11.5], color=NAVY_DEEP, linewidth=2, zorder=7)
    ax.text(10.5, 11.5, "Jalan Utama & Akses Kawasan Kosambi", va='center', fontsize=7.5, zorder=7)

    out_file = os.path.join(out_dir, "peta_lokasi_bakmi_mimu.png")
    plt.tight_layout()
    plt.savefig(out_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[GENERATED] Gambar 2.1: {out_file}")

# =============================================================================
# 3. GAMBAR 4.2: NETWORK PLANNING & CRITICAL PATH METHOD (CPM) RELOKASI
# =============================================================================
def generate_network_planning():
    fig, ax = plt.subplots(figsize=(13.5, 7.5), dpi=300)
    fig.patch.set_facecolor(BG_LIGHT)
    ax.set_facecolor(BG_LIGHT)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Header
    ax.text(50, 95, "NETWORK PLANNING & CRITICAL PATH METHOD (CPM) RELOKASI 2016", 
            ha='center', va='center', fontsize=15, fontweight='bold', color=NAVY_DEEP)
    ax.text(50, 91.5, "Diagram Alur Ketergantungan Aktivitas & Jalur Kritis Pelaksanaan Proyek (Total Durasi: 12 Minggu)", 
            ha='center', va='center', fontsize=10.5, style='italic', color=SLATE_MUTED)

    # Nodes Definitions: (id, x, y, ES, LF)
    nodes = {
        1: (8, 50, 0, 0),
        2: (20, 50, 2, 2),
        3: (32, 64, 4, 4),
        4: (32, 36, 4, 6),
        5: (44, 64, 5, 5),
        6: (56, 64, 7, 7),
        7: (56, 36, 6, 7),
        8: (68, 64, 9, 9),
        9: (68, 36, 8, 9),
        10: (80, 64, 10, 10),
        11: (92, 50, 12, 12)
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

    # Draw Edges (Arrows)
    for u, v, code, name, dur, is_crit in edges:
        x1, y1, _, _ = nodes[u]
        x2, y2, _, _ = nodes[v]
        
        # Calculate angle and offset so arrow stops at node circle border (radius ~ 3.2)
        dx = x2 - x1
        dy = y2 - y1
        dist = np.hypot(dx, dy)
        if dist == 0: continue
        r = 3.6
        sx = x1 + (dx / dist) * r
        sy = y1 + (dy / dist) * r
        ex = x2 - (dx / dist) * r
        ey = y2 - (dy / dist) * r

        color = GOLD_AMBER if is_crit else NAVY_LIGHT
        lwidth = 2.8 if is_crit else 1.4
        linestyle = '-' if dur > 0 else '--'

        arrow = FancyArrowPatch((sx, sy), (ex, ey),
                                arrowstyle='-|>',
                                mutation_scale=14 if is_crit else 10,
                                color=color,
                                linewidth=lwidth,
                                linestyle=linestyle,
                                zorder=3)
        ax.add_patch(arrow)

        # Label position
        if dur > 0:
            mx = (sx + ex) / 2
            my = (sy + ey) / 2
            # Offset label
            if dy == 0:
                ly = my + (2.8 if is_crit else -2.8)
                lx = mx
            elif dy > 0:
                lx = mx - 2.5
                ly = my + 1.8
            else:
                lx = mx - 2.5
                ly = my - 1.8
            
            lbl_txt = f"{code} ({dur} mg)"
            ax.text(lx, ly, lbl_txt, ha='center', va='center', 
                    fontsize=8.5, fontweight='bold', color=color, zorder=5)

    # Draw Nodes (Circles with 3 sections: Node ID, ES, LF)
    for nid, (nx, ny, es, lf) in nodes.items():
        # Outer Node Circle
        circ = patches.Circle((nx, ny), 3.4, edgecolor=NAVY_DEEP, facecolor=CARD_BG, linewidth=2.0, zorder=6)
        ax.add_patch(circ)

        # Horizontal divider inside node
        ax.plot([nx - 3.4, nx + 3.4], [ny, ny], color=NAVY_DEEP, linewidth=1.2, zorder=7)
        # Vertical divider on bottom half
        ax.plot([nx, nx], [ny - 3.4, ny], color=NAVY_DEEP, linewidth=1.2, zorder=7)

        # Texts inside node
        # Node ID top center
        ax.text(nx, ny + 1.4, str(nid), ha='center', va='center', fontsize=10, fontweight='bold', color=NAVY_DEEP, zorder=8)
        # ES bottom left
        ax.text(nx - 1.6, ny - 1.6, str(es), ha='center', va='center', fontsize=7.5, color=SLATE_MUTED, zorder=8)
        # LF bottom right
        ax.text(nx + 1.6, ny - 1.6, str(lf), ha='center', va='center', fontsize=7.5, color=GOLD_AMBER if es == lf else SLATE_MUTED, fontweight='bold' if es == lf else 'normal', zorder=8)

    # Critical Path Indicator Box
    crit_box = FancyBboxPatch((8, 16), 84, 15, boxstyle="round,pad=0.4",
                              facecolor=GOLD_LIGHT, edgecolor=GOLD_AMBER, linewidth=1.5, zorder=4)
    ax.add_patch(crit_box)

    ax.text(10, 27.5, "KETERANGAN SIMPUL & JALUR KRITIS (CRITICAL PATH METHOD):", fontsize=9.5, fontweight='bold', color=GOLD_AMBER)
    c_desc = (
        "• Jalur Kritis (Tebal Emas): 1 -> 2 -> 3 -> 5 -> 6 -> 8 -> 10 -> 11  |  Total Durasi Waktu Proyek: 12 Minggu\n"
        "• Urutan Aktivitas Kritis: A (Survei 2 mg) -> B (Sewa Ruko 2 mg) -> E (Renovasi Cat/Pipa 1 mg) -> G (Gerobak Teras 2 mg) ->\n"
        "  I (Meja/Kursi 2 mg) -> K (Setting 23 Kursi 1 mg) -> M-P (Rekrutmen 3 Staf, SOP Adonan, Trial Run & Grand Opening 2 mg)\n"
        "• Keterangan Kotak Node: Bagian Atas = Nomor Simpul; Kiri Bawah = Earliest Start (ES); Kanan Bawah = Latest Finish (LF)"
    )
    ax.text(10, 20.5, c_desc, fontsize=8.6, color=SLATE_DARK, va='center')

    # Bottom Footnote
    ax.text(50, 7, "Diagram Network Planning 100% Selaras dengan Gantt Chart 12 Minggu pada BAB IV Subbab 4.3 Dokumen Laporan",
            ha='center', va='center', fontsize=9.2, fontweight='bold', color=NAVY_DEEP)

    out_file = os.path.join(out_dir, "network_planning_bakmi_mimu.png")
    plt.tight_layout()
    plt.savefig(out_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[GENERATED] Gambar 4.2: {out_file}")

# =============================================================================
# 4. GAMBAR 1.1: GAMBARAN PRODUK & MENU BAKMI MIMU CARINA SAYANG
# =============================================================================
def generate_gambaran_produk():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    fig.patch.set_facecolor(BG_LIGHT)
    ax.set_facecolor(BG_LIGHT)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Title & Subtitle
    ax.text(50, 95, "GAMBARAN PRODUK & BAURAN MENU BAKMI MIMU CARINA SAYANG", 
            ha='center', va='center', fontsize=15, fontweight='bold', color=NAVY_DEEP)
    ax.text(50, 91.5, "Resep Keluarga Sejak 1999 • 100% Daging Murni Pilihan (Tanpa Bebek & Jamur) • 3 Pilihan Racikan Minyak Khas", 
            ha='center', va='center', fontsize=10.5, style='italic', color=SLATE_MUTED)

    # 4 Product Category Showcase Panels
    prod_cards = [
        (6, 48, 42, 38, "1. BAKMI SPESIAL 3 ISI & CAMPUR", GOLD_AMBER, [
            ("Mie Spesial 3 Isi Lengkap", "Rp 39.000", "Ayam Putih + Babi Kecap + Casiu Madu Gurih"),
            ("Mie Campur (Paling Laris)", "Rp 29.000", "Topping Ayam Putih Gurih & Babi Kecap Manis"),
            ("Mie Ayam Putih / Babi Saja", "Rp 29.000", "Topping Tunggal Murni Sesuai Selera Tamu"),
            ("Mie Daging Casiu Madu", "Rp 31.000", "Casiu Panggang Merah Manis Gurih Renyah"),
            ("Karakteristik Produk:", "", "Tekstur kenyal shining, adonan fresh mingguan, dipadu 3 pilihan minyak (babi, ayam, sayur+wijen).")
        ]),
        (52, 48, 42, 38, "2. VARIASI UKURAN PORSI FLEKSIBEL", NAVY_LIGHT, [
            ("Porsi Jumbo (+100% Mie)", "Rp 49.000 - 59.000", "2x Lipat Porsi Reguler (+100% takaran mie adonan)"),
            (" - Jumbo Daging Campur", "Rp 49.000", "Porsi ekstra kenyang daging campur ayam-babi"),
            (" - Jumbo Daging Casiu", "Rp 54.000", "Porsi jumbo topping melimpah casiu panggang"),
            (" - Jumbo Spesial 3 Isi", "Rp 59.000", "Porsi jumbo lengkap 3 varian daging murni"),
            ("Porsi Kecil (Sarapan Pagi)", "Rp 27.000 - 29.000", "Porsi ringan anak sekolah / sarapan pagi hemat")
        ]),
        (6, 6, 42, 38, "3. MENU PELENGKAP & KUAH GURIH", GREEN_EMERALD, [
            ("Pangsit Rebus Kuah (5 pcs)", "Rp 22.500", "Daging babi cincang gurih (@Rp 4.500 / pcs)"),
            ("Swikiaw Rebus Kuah (5 pcs)", "Rp 27.500", "Isian babi + udang segar (@Rp 5.500 / pcs)"),
            ("Baso Sapi Kuah (5 pcs)", "Rp 22.500", "Baso sapi kenyal kuah kaldu tulang gurih"),
            ("Baso Ikan Kuah (5 pcs)", "Rp 22.500", "Baso ikan putih lembut higienis segar"),
            ("Baso Goreng Babi-Udang", "Rp 8.000 / pcs", "Garing renyah di luar, empuk gurih di dalam"),
            ("Pangsit Goreng Renyah", "Rp 5.000 / pcs", "Kulit pangsit garing renyah pendamping mie")
        ]),
        (52, 6, 42, 38, "4. MINUMAN SEGAR & SAUS MEJA", NAVY_DEEP, [
            ("Badak Sarsaparilla", "Rp 12.000 - 14.000", "Minuman soda legendaris Siantar aroma sarsaparilla"),
            ("Susu Kacang Kedelai", "Rp 10.000 - 12.000", "Susu kedelai kental manis dingin menyegarkan"),
            ("Liang Teh Tradisional", "Rp 10.000 - 12.000", "Minuman herbal pereda panas dalam alami"),
            ("Es Teh Manis / Tawar", "Rp 1.000 - 5.000", "Teh seduh wangi melati segar"),
            ("Saus Meja Resmi:", "", "Botol resmi Cap Belibis & Mangga Besar di setiap meja"),
            ("Higienitas Sajian:", "", "Penyajian mangkok keramik higienis standar resto")
        ])
    ]

    for cx, cy, cw, ch, header, col, items in prod_cards:
        card = FancyBboxPatch((cx, cy), cw, ch, boxstyle="round,pad=0.4",
                              facecolor=CARD_BG, edgecolor=col, linewidth=2.0)
        ax.add_patch(card)

        # Header bar
        hbar = FancyBboxPatch((cx, cy + ch - 4.5), cw, 4.5, boxstyle="round,pad=0.2",
                              facecolor=col, edgecolor=col, linewidth=1.0)
        ax.add_patch(hbar)
        ax.text(cx + cw / 2, cy + ch - 2.3, header, 
                ha='center', va='center', fontsize=10.5, fontweight='bold', color=CARD_BG)

        # Items
        start_y = cy + ch - 7.5
        for it_name, it_price, it_desc in items:
            if it_price:
                ax.text(cx + 1.8, start_y, f"• {it_name}", ha='left', va='center', fontsize=9.2, fontweight='bold', color=NAVY_DEEP)
                ax.text(cx + cw - 1.8, start_y, it_price, ha='right', va='center', fontsize=9.2, fontweight='bold', color=GOLD_AMBER)
                ax.text(cx + 3.5, start_y - 2.5, it_desc, ha='left', va='center', fontsize=8.0, color=SLATE_MUTED)
                start_y -= 5.5
            else:
                ax.text(cx + 1.8, start_y, f"• {it_name}", ha='left', va='center', fontsize=8.8, fontweight='bold', color=col)
                ax.text(cx + 3.5, start_y - 2.3, it_desc, ha='left', va='center', fontsize=8.0, color=SLATE_DARK)
                start_y -= 5.0

    out_file = os.path.join(out_dir, "produk_bakmi_mimu.png")
    plt.tight_layout()
    plt.savefig(out_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[GENERATED] Gambar 1.1: {out_file}")

if __name__ == "__main__":
    print("Mulai pembuatan seluruh aset visual pendukung tugas SKB FEB UKRIDA...")
    generate_struktur_organisasi()
    generate_peta_lokasi()
    generate_network_planning()
    generate_gambaran_produk()
    print("\n[SUKSES] Seluruh aset visual berhasil dibuat 100% presisi.")
