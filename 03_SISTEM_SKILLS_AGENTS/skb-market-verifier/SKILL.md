---
name: skb-market-verifier
description: Protokol dan helper riset pasar untuk melakukan fact-checking dan verifikasi data pasar ke Google (harga bahan baku riil, standar UMR wilayah, biaya sewa ruko, tarif utilitas PLN/PDAM, dan benchmarking kompetitor) untuk tugas SKB.
---

# SKB Market & Fact-Checking Verifier (`skb-market-verifier`)

Skill ini digunakan untuk memvalidasi dan memverifikasi kebenaran parameter angka dalam laporan **Studi Kelayakan Bisnis (SKB)** ke dunia nyata melalui penelusuran **Google Search**, database resmi (BPS, Pemprov, PLN, PDAM), dan marketplace/properti komersial.

Dengan skill ini, seluruh angka yang dimasukkan dalam Bab II (Pasar & Pemasaran), Bab IV (Teknis Operasi), dan Bab V (Keuangan) memiliki dasar fakta yang kuat dan dapat dipertanggungjawabkan di hadapan dosen penguji.

---

## 1. Protokol Verifikasi 5 Dimensi Pasar

| Dimensi Parameter | Objek Verifikasi | Sumber Data Referensi Utama | Rumus / Standar Validasi |
| :--- | :--- | :--- | :--- |
| **1. Standar Gaji & UMR** | UMR/UMK wilayah operasional usaha | Keputusan Gubernur (Kepgub) terkini / Disnakertrans | Gaji staf tidak boleh di bawah UMR kecuali ada kompensasi makan/tunjangan operasional yang jelas |
| **2. Harga Bahan Baku** | Biji kopi, tepung, daging, susu, minyak, topping | E-Commerce (Tokopedia, Shopee), Indonetwork, Pasar Induk | Gunakan median harga satuan (kg/liter/pack) grosir, bukan harga eceran minimarket |
| **3. Harga Peralatan (Capex)**| Mesin espresso, loyang martabak, chiller, etalase, POS | Marketplace resmi distributor F&B (Restomart, Tokopedia Official) | Sertakan biaya ongkos kirim dan instalasi awal |
| **4. Sewa Tempat & Renovasi** | Harga sewa ruko 1-2 lantai di area target | Lamudi, OLX Properti, Rumah123 | Bandingkan harga per meter persegi di jalan arteri vs jalan lingkungan kampus |
| **5. Utilitas (Listrik & Air)** | Tarif listrik komersial B-1 / B-2, air PAM | Website resmi PLN (Tarif Tenaga Listrik B-1/TR) & PAM Jaya | Hitung daya (VA) vs jam operasional harian (12 jam/hari $\times$ 30 hari) |

---

## 2. Template Query Google Search yang Efektif

Gunakan formula pencarian berikut saat memverifikasi data untuk ide bisnis baru:

### A. Verifikasi UMR / Standar Upah Kerja
```text
"UMR Jakarta Barat [tahun]" OR "UMK DKI Jakarta [tahun]" keputusan gubernur
"standar gaji barista Jakarta [tahun]" jobstreet OR glints
"gaji kasir restoran Jakarta [tahun]" indeed OR karir
```

### B. Verifikasi Harga Sewa Ruko & Lokasi Usaha
```text
"sewa ruko tanjung duren" site:lamudi.co.id OR site:olx.co.id OR site:rumah123.com
"sewa kios kampus ukrida" OR "sewa tempat usaha grogol petamburan"
```

### C. Verifikasi Harga Bahan Baku Utama
```text
"harga biji kopi espresso blend 1kg" site:tokopedia.com
"harga tepung terigu cakra kembar 25kg sak" grosir
"harga keju cheddar prochiz 2kg" distributor
"harga daging sapi sengkel sirloin 1kg" pasar jaya
```

### D. Verifikasi Tarif Utilitas Resmi
```text
"tarif dasar listrik PLN golongan B-1 bisnis" [tahun]
"tarif air pdam jaya komersial per m3" [tahun]
"harga gas elpiji 12kg isi ulang agen resmi"
```

---

## 3. Logika Validasi Kapasitas & Permintaan Riil

Saat memeriksa Bab II & Bab V, pastikan proyeksi volume penjualan masuk akal (*sanity check*):

1. **Rasio Kapasitas Duduk vs Target Harian**:
   - Jika toko memiliki 6 meja $\times$ 4 kursi = 24 kapasitas duduk (*dine-in*).
   - Waktu operasi: 10.00 – 22.00 (12 jam).
   - Jika rata-rata waktu makan konsumen = 45 menit, maka perputaran meja (*table turnover*) maksimum $\approx 16 \times 24 = 384$ transaksi/hari.
   - Jika target penjualan harian di Bab V adalah 70–120 porsi/hari, maka utilisasi kapasitas adalah:
     $$\text{Utilisasi} = \frac{80}{384} \approx 20.8\% \quad (\text{Sangat Realistis})$$
   - *Catatan Audit*: Jika target harian melampaui kapasitas dapur atau kapasitas meja tanpa layanan *takeaway/delivery* yang memadai, tandai sebagai **REVISI**.

2. **Asumsi 300 Hari Operasi Tahunan**:
   - Standar acuan SKB FEB menggunakan **25 hari kerja/bulan** $\times$ 12 bulan = **300 hari/tahun**.
   - Hari libur atau hari perawatan/maintenance dialokasikan 5-6 hari per bulan.

---

## 4. Helper Script (`research_helper.py`)

Gunakan skrip pembantu di `.agents/skills/skb-market-verifier/scripts/research_helper.py` untuk menghasilkan daftar kata kunci verifikasi dan tabel referensi kutipan data.
