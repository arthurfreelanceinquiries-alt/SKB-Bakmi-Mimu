import docx
import openpyxl

doc = docx.Document(r"d:\Perkuliahan\Kelass\SKB\skb debby MARTABAK HATI\MARTABAK HATI\Martabak Hati_312020064_DebbyDelicia.docx")
wb = openpyxl.load_workbook(r"d:\Perkuliahan\Kelass\SKB\skb debby MARTABAK HATI\MARTABAK HATI\Martabak Hati _312020064_Debby Delicia.xlsx", data_only=True)

print("=== CHECKING ALL TABLES FOR PRICING / DATA ISSUES ===")
for i, t in enumerate(doc.tables):
    title = f"Table {i}"
    # find preceding paragraph
    rows_data = []
    for r in t.rows:
        rows_data.append([c.text.replace("\n", " ").strip() for c in r.cells])
    print(f"\n--- TABLE {i} (Rows: {len(t.rows)}, Cols: {len(t.columns)}) ---")
    for r in rows_data[:4]:
        print("  ", r)
    if len(rows_data) > 4:
        print("   ... (total rows:", len(rows_data), ")")
        print("   LAST ROW:", rows_data[-1])

