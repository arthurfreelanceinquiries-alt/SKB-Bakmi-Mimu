import docx
import openpyxl
import re

doc = docx.Document(r"d:\Perkuliahan\Kelass\SKB\skb debby MARTABAK HATI\MARTABAK HATI\Martabak Hati_312020064_DebbyDelicia.docx")
wb = openpyxl.load_workbook(r"d:\Perkuliahan\Kelass\SKB\skb debby MARTABAK HATI\MARTABAK HATI\Martabak Hati _312020064_Debby Delicia.xlsx", data_only=True)

print("=== CHECKING TABLE 1.1 vs TEXT vs EXCEL ===")
# Table 0 is Tabel 1.1 Jenis Produk
t0 = doc.tables[0]
t0_data = []
for row in t0.rows:
    t0_data.append([c.text.strip() for c in row.cells])

print("Tabel 1.1 Content:")
for r in t0_data:
    print(" ", r)

# Check paragraph 139 (Cash Inflow text)
print("\nParagraf 139 (Text Penjelasan 6.4 Cash Inflow):")
for p in doc.paragraphs:
    if "Yang dimana usaha martabak hati menetapkan target" in p.text:
        print(" ", p.text)

# Check Excel Sheet 2 prices
ws2 = wb["Sheet2"]
print("\nExcel Sheet 2 (Harga Martabak Tahun 1):")
for r in range(43, 115):
    val_b = ws2[f'B{r}'].value
    val_c = ws2[f'C{r}'].value
    if val_b and any(k in str(val_b) for k in ["Martabak", "Penjualan", "Harga"]):
        print(f"  Row {r}: B='{val_b}', C='{val_c}'")

# Check Raw Materials prices Table 5.11 / Table 5.12 vs Excel
print("\nCheck Table 5.11 vs Excel:")
for t_idx, t in enumerate(doc.tables):
    txt_header = [c.text.strip() for c in t.rows[0].cells]
    if any("Bahan" in h for h in txt_header) or any("Gaji" in h for h in txt_header) or any("Harga" in h for h in txt_header):
        print(f"Table {t_idx} header: {txt_header}")
        for row in t.rows[:5]:
            print("  ", [c.text.strip() for c in row.cells])
