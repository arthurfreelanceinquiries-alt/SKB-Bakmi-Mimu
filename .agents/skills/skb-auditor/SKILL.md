---
name: skb-auditor
description: Skill audit komprehensif untuk memeriksa kepatuhan format, konsistensi data narasi Word vs model Excel, validitas rumus keuangan (NPV, IRR, PP, PI), dan kelayakan investasi tugas SKB FEB UKRIDA.
---

# SKB Comprehensive Auditor (`skb-auditor`)

Skill ini bertindak sebagai auditor independen dan asisten penelaah mutu tugas **Studi Kelayakan Bisnis (SKB)** berstandar FEB UKRIDA. Tugasnya adalah memastikan bahwa naskah Word (`.docx`) dan lembar kerja Excel (`.xlsx`) bebas dari kesalahan format, bebas dari kontradiksi angka antar-bab, dan memiliki perhitungan finansial yang 100% akurat secara matematis.

---

## 1. Rubrik Audit 4 Pilar Mutu

```
                                  +-------------------------------------+
                                  |     SKB QUALITY AUDIT SUITE         |
                                  +-------------------------------------+
                                                     |
         +-------------------+-----------------------+-----------------------+-------------------+
         |                   |                       |                       |                   |
         v                   v                       v                       v                   v
   [PILAR 1: FORMAT]   [PILAR 2: STRUKTUR]     [PILAR 3: SINKRONISASI]  [PILAR 4: MATEMATIKA]  [PILAR 5: KELAYAKAN]
   - Times New Roman   - BAB I - BAB VI        - Teks Docx vs Tabel     - Formula CIF & COF    - NPV > 0
   - Spasi 1.5         - Subbab 1.1 s/d 7.1    - Docx vs Sheet 1,2,3    - Depresiasi 20%       - IRR > 20%
   - Margin 2.54 cm    - Urutan Tabel & Gambar - Nilai Modal Awal       - Rumus PV & NPV       - PP < 3 Tahun
   - Justified text    - Denah & Network Plan  - Nilai Proceed          - Rumus IRR & PI       - PI > 1.2
```

---

## 2. Rincian Checklist Pengujian

### Pilar 1: Format & Tipografi
- [ ] **Font Naskah**: Seluruh teks menggunakan font *Times New Roman* ukuran 12 pt.
- [ ] **Spasi Baris**: Naskah isi berjarak 1.5 spasi (bukan 1.0 atau 1.15).
- [ ] **Margin Halaman**: Keempat sisi (Atas, Bawah, Kiri, Kanan) berukuran tepat 2.54 cm (1 inci).
- [ ] **Perataan Teks**: Paragraf isi rata kiri-kanan (*Justified*). Judul bab dan nomor gambar/tabel di tengah (*Center*).
- [ ] **Format Keterangan**:
  - Judul tabel diletakkan di **atas** tabel dengan format: `Tabel [Bab].[No]. [Nama Tabel]`.
  - Judul gambar diletakkan di **bawah** gambar dengan format: `Gambar [Bab].[No] [Nama Gambar]`.

### Pilar 2: Kelengkapan Struktur Naskah
- [ ] **BAB I**: Memuat sejarah kuliner/komoditas, peluang, profil usaha, `Tabel 1.1`, dan `Gambar 1.1`.
- [ ] **BAB II**: Memuat Segmentasi (Demografis, Geografis, Perilaku), Diferensiasi 5 pilar, Lokasi (`Gambar 2.1`), Target, dan Bauran 7P.
- [ ] **BAB III**: Memuat Bagan Struktur Organisasi (`Gambar 3.1`) dan Uraian Pekerjaan (*Job Description*) lengkap per peran.
- [ ] **BAB IV**: Memuat Sewa Gedung, Capex Peralatan (`Tabel 4.1`), Layout Denah (`Gambar 4.1`), Network Planning A-L (`Gambar 4.2`), dan Gantt Chart (`Tabel 4.2`).
- [ ] **BAB V**: Memuat Penjelasan Naratif (6.1 s/d 6.5) dan 25 tabel keuangan.
- [ ] **BAB VI**: Memuat `Tabel 6.1 Kesimpulan Kelayakan`, pernyataan kelayakan, serta subbab `7.1 Kesimpulan dan Saran`.

### Pilar 3: Sinkronisasi Silang Data (Docx vs Xlsx)
- [ ] **Investasi Awal (Initial Outlay)**: Angka di teks Bab V.6.1 = Total Tabel 5.19 Docx = Cell `Sheet1!C19` = Cell `Sheet3!C9`.
- [ ] **Aktiva Tetap (Capex)**: Total Tabel 4.1 & 5.2 Docx = Cell `Sheet1!F66` = Cell `Sheet3!C4`.
- [ ] **Cash Inflow (CIF)**: Angka CIF Tahun 1 s/d 5 di Tabel 5.1 Docx = Cell `Sheet2!C20:G20`.
- [ ] **Cash Outflow (COF)**: Angka COF Tahun 1 s/d 5 di Tabel 5.1 Docx = Cell `Sheet2!C173:G173`.
- [ ] **Proceed**: Angka Proceed Tahun 1 s/d 5 di Tabel 5.1 Docx = Cell `Sheet1!D10:H10` = Cell `Sheet3!C13:C17`.

### Pilar 4: Integritas Matematis Finansial
- [ ] **Asumsi Hari Kerja**: Penjualan bulanan dihitung dengan pengali 25 hari kerja/bulan ($\times$ 12 bulan = 300 hari operasi/tahun).
- [ ] **Depresiasi Garis Lurus**: Depresiasi per tahun = $20\% \times \text{Total Aktiva Tetap}$ (Cell `Sheet1!D68`).
- [ ] **Net Cash Flow (NCF / EAT)**: $\text{NCF} = \text{CIF} - \text{COF}$.
- [ ] **Proceed**: $\text{Proceed} = \text{NCF} - \text{Depresiasi}$.
- [ ] **Present Value (PV)**: Formula Excel `=PV(Discount_Rate, Tahun, , -Proceed)`.
- [ ] **Net Present Value (NPV)**: $\sum \text{PV} - \text{Initial Outlay}$.
- [ ] **Internal Rate of Return (IRR)**: Formula `=IRR(Cashflow_Tahun_0_sd_5)`.
- [ ] **Payback Period (PP)**: Waktu pemulihan modal dalam bulan/tahun.
- [ ] **Profitability Index (PI)**: $\frac{\sum \text{Proceed}}{\text{Initial Outlay}}$.

### Pilar 5: Kriteria Kelayakan Akademik FEB
- [ ] **NPV**: Bernilai **Positif ($> 0$)** $\rightarrow$ Status: **LAYAK**.
- [ ] **IRR**: Nilai persentase **$> 20.0\%$ (Opportunity Cost)** $\rightarrow$ Status: **LAYAK**.
- [ ] **Payback Period**: Waktu kembali modal **$< 3.0$ Tahun** $\rightarrow$ Status: **LAYAK**.
- [ ] **Profitability Index**: Nilai rasio **$> 1.20$** $\rightarrow$ Status: **LAYAK**.

---

## 3. Cara Menjalankan Audit Otomatis

Jalankan script audit bawaan skill:
```bash
python .agents/skills/skb-auditor/scripts/audit_skb.py --docx "Laporan_SKB.docx" --xlsx "Model_Finansial_SKB.xlsx"
```

Script akan memindai kedua file secara serentak dan mencetak laporan audit kelulusan (**PASS / WARNING / FAIL**) beserta rekomendasi revisi spesifik.
