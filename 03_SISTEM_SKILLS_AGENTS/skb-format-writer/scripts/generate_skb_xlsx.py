import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import sys
import os

def build_skb_excel(config_path, output_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    wb = openpyxl.Workbook()
    # Default sheet
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws2 = wb.create_sheet(title="Sheet2")
    ws3 = wb.create_sheet(title="Sheet3")

    # Styling helpers
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_normal = Font(name="Times New Roman", size=11)
    font_header = Font(name="Times New Roman", size=11, bold=True, color="000000")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    accent_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )

    # -------------------------------------------------------------
    # SHEET 1: NCF, Initial Outlay, Aktiva Tetap, Sewa, AMDAL, Survey, Promosi
    # -------------------------------------------------------------
    ws1['B3'] = "NCF"
    ws1['B3'].font = Font(name="Times New Roman", size=12, bold=True)

    headers_ncf = ["Keterangan", "Tahun 0", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]
    for col_idx, h in enumerate(headers_ncf, start=2):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    ws1['B5'] = "Investasi Awal"
    ws1['C5'] = "=C19"
    ws1['B6'] = "Cash Inflow/CIF"
    ws1['D6'] = "=Sheet2!C20"
    ws1['E6'] = "=Sheet2!D20"
    ws1['F6'] = "=Sheet2!E20"
    ws1['G6'] = "=Sheet2!F20"
    ws1['H6'] = "=Sheet2!G20"

    ws1['B7'] = "Cash Outflow/COF(-)"
    ws1['D7'] = "=Sheet2!C173"
    ws1['E7'] = "=Sheet2!D173"
    ws1['F7'] = "=Sheet2!E173"
    ws1['G7'] = "=Sheet2!F173"
    ws1['H7'] = "=Sheet2!G173"

    ws1['B8'] = "NCF = EAT"
    ws1['D8'] = "=D6-D7"
    ws1['E8'] = "=E6-E7"
    ws1['F8'] = "=F6-F7"
    ws1['G8'] = "=G6-G7"
    ws1['H8'] = "=H6-H7"

    ws1['B9'] = "Depresiasi (+)"
    ws1['D9'] = "=$D$68"
    ws1['E9'] = "=$D$68"
    ws1['F9'] = "=$D$68"
    ws1['G9'] = "=$D$68"
    ws1['H9'] = "=$D$68"

    ws1['B10'] = "Proceed"
    ws1['D10'] = "=D8-D9"
    ws1['E10'] = "=E8-E9"
    ws1['F10'] = "=F8-F9"
    ws1['G10'] = "=G8-G9"
    ws1['H10'] = "=H8-H9"

    for r in range(5, 11):
        for c in range(2, 9):
            cell = ws1.cell(row=r, column=c)
            cell.font = font_bold if r in [5, 8, 10] else font_normal
            cell.border = thin_border
            if c >= 3 and cell.value:
                cell.number_format = '#,##0'

    # Initial Outlay Block
    ws1['B13'] = "Initial Outlay"
    ws1['B13'].font = Font(name="Times New Roman", size=12, bold=True)
    outlay_items = [
        ("Aktiva Tetap", "=F66"),
        ("Biaya Sewa", "=E72"),
        ("Biaya Amdal", "=F78"),
        ("Biaya Survey", "=F85"),
        ("Biaya Promosi", "=F90")
    ]
    for i, (label, formula) in enumerate(outlay_items, start=14):
        ws1.cell(row=i, column=2, value=label).font = font_normal
        ws1.cell(row=i, column=2).border = thin_border
        c_val = ws1.cell(row=i, column=3, value=formula)
        c_val.font = font_normal
        c_val.number_format = '#,##0'
        c_val.border = thin_border

    ws1['B19'] = "Total"
    ws1['B19'].font = font_bold
    ws1['B19'].border = thin_border
    c_tot = ws1.cell(row=19, column=3, value="=SUM(C14:C18)")
    c_tot.font = font_bold
    c_tot.number_format = '#,##0'
    c_tot.border = thin_border

    # AKTIVA TETAP (Capex)
    ws1['B22'] = "AKTIVA TETAP"
    ws1['B22'].font = Font(name="Times New Roman", size=12, bold=True)
    capex_hdrs = ["No.", "Keterangan", "Jumlah", "Harga Satuan", "Total"]
    for col_idx, h in enumerate(capex_hdrs, start=2):
        cell = ws1.cell(row=23, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    capex_list = cfg.get("capex_items", [])
    start_row = 24
    for idx, item in enumerate(capex_list, start=1):
        r = start_row + idx - 1
        ws1.cell(row=r, column=2, value=idx).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=3, value=item["name"])
        ws1.cell(row=r, column=4, value=item["qty"]).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=5, value=item["unit_price"]).number_format = '#,##0'
        ws1.cell(row=r, column=6, value=f"=D{r}*E{r}").number_format = '#,##0'
        for col_idx in range(2, 7):
            ws1.cell(row=r, column=col_idx).font = font_normal
            ws1.cell(row=r, column=col_idx).border = thin_border

    last_capex_row = start_row + len(capex_list) - 1
    total_capex_row = 66
    ws1.cell(row=total_capex_row, column=2, value="Total Biaya").font = font_bold
    ws1.cell(row=total_capex_row, column=2).border = thin_border
    tot_val = ws1.cell(row=total_capex_row, column=6, value=f"=SUM(F{start_row}:F{last_capex_row})")
    tot_val.font = font_bold
    tot_val.number_format = '#,##0'
    tot_val.border = thin_border

    # Depresiasi 20%
    ws1.cell(row=68, column=2, value="Depresiasi (20%)").font = font_bold
    ws1.cell(row=68, column=2).border = thin_border
    dep_val = ws1.cell(row=68, column=4, value=f"=F{total_capex_row}*20%")
    dep_val.font = font_bold
    dep_val.number_format = '#,##0'
    dep_val.border = thin_border

    # Biaya Sewa
    ws1['B70'] = "BIAYA SEWA"
    ws1['B70'].font = Font(name="Times New Roman", size=12, bold=True)
    sewa_hdrs = ["No", "Keterangan", "Jangka Waktu (/tahun)", "Harga Sewa per tahun", "Harga Sewa per bulan"]
    for c_i, h in enumerate(sewa_hdrs, start=2):
        cell = ws1.cell(row=71, column=c_i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.border = thin_border
    ws1.cell(row=72, column=2, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=72, column=3, value="Biaya Sewa Tempat Ruko 1 Lantai")
    ws1.cell(row=72, column=4, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=72, column=5, value=cfg["operational_costs"].get("rent_per_year", 20000000)).number_format = '#,##0'
    ws1.cell(row=72, column=6, value="=(D72*E72)/12").number_format = '#,##0'
    for c_i in range(2, 7):
        ws1.cell(row=72, column=c_i).font = font_normal
        ws1.cell(row=72, column=c_i).border = thin_border

    ws1.cell(row=73, column=2, value="Total Biaya").font = font_bold
    ws1.cell(row=73, column=2).border = thin_border
    ws1.cell(row=73, column=6, value="=F72").font = font_bold
    ws1.cell(row=73, column=6).number_format = '#,##0'
    ws1.cell(row=73, column=6).border = thin_border

    # AMDAL
    ws1['B75'] = "BIAYA AMDAL"
    ws1['B75'].font = Font(name="Times New Roman", size=12, bold=True)
    amdal_hdrs = ["No", "Keterangan", "Jumlah", "Harga satuan", "Total"]
    for c_i, h in enumerate(amdal_hdrs, start=2):
        cell = ws1.cell(row=76, column=c_i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.border = thin_border
    ws1.cell(row=77, column=2, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=77, column=3, value="Tempat Sampah & Kebersihan Lingkungan")
    ws1.cell(row=77, column=4, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=77, column=5, value=cfg["operational_costs"].get("amdal_initial", 300000)).number_format = '#,##0'
    ws1.cell(row=77, column=6, value="=D77*E77").number_format = '#,##0'
    for c_i in range(2, 7):
        ws1.cell(row=77, column=c_i).font = font_normal
        ws1.cell(row=77, column=c_i).border = thin_border
    ws1.cell(row=78, column=2, value="Total Biaya").font = font_bold
    ws1.cell(row=78, column=2).border = thin_border
    ws1.cell(row=78, column=6, value="=SUM(F77:F77)").font = font_bold
    ws1.cell(row=78, column=6).number_format = '#,##0'
    ws1.cell(row=78, column=6).border = thin_border

    # Biaya Survey
    ws1['B81'] = "BIAYA SURVEY"
    ws1['B81'].font = Font(name="Times New Roman", size=12, bold=True)
    for c_i, h in enumerate(amdal_hdrs, start=2):
        cell = ws1.cell(row=82, column=c_i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.border = thin_border
    ws1.cell(row=83, column=2, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=83, column=3, value="Biaya Konsumsi Tim Survey")
    ws1.cell(row=83, column=4, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=83, column=5, value=250000).number_format = '#,##0'
    ws1.cell(row=83, column=6, value="=D83*E83").number_format = '#,##0'
    
    ws1.cell(row=84, column=2, value=2).alignment = Alignment(horizontal="center")
    ws1.cell(row=84, column=3, value="Biaya Transportasi & Administrasi")
    ws1.cell(row=84, column=4, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=84, column=5, value=250000).number_format = '#,##0'
    ws1.cell(row=84, column=6, value="=D84*E84").number_format = '#,##0'
    for r_i in [83, 84]:
        for c_i in range(2, 7):
            ws1.cell(row=r_i, column=c_i).font = font_normal
            ws1.cell(row=r_i, column=c_i).border = thin_border
    ws1.cell(row=85, column=2, value="Total Biaya").font = font_bold
    ws1.cell(row=85, column=2).border = thin_border
    ws1.cell(row=85, column=6, value="=SUM(F83:F84)").font = font_bold
    ws1.cell(row=85, column=6).number_format = '#,##0'
    ws1.cell(row=85, column=6).border = thin_border

    # Biaya Promosi
    ws1['B87'] = "BIAYA PROMOSI"
    ws1['B87'].font = Font(name="Times New Roman", size=12, bold=True)
    promosi_hdrs = ["No", "Keterangan", "Jumlah", "Harga Iklan", "Total"]
    for c_i, h in enumerate(promosi_hdrs, start=2):
        cell = ws1.cell(row=88, column=c_i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.border = thin_border
    ws1.cell(row=89, column=2, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=89, column=3, value="Spanduk & Banner Grand Opening")
    ws1.cell(row=89, column=4, value=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=89, column=5, value=cfg["operational_costs"].get("promosi_initial", 500000)).number_format = '#,##0'
    ws1.cell(row=89, column=6, value="=D89*E89").number_format = '#,##0'
    for c_i in range(2, 7):
        ws1.cell(row=89, column=c_i).font = font_normal
        ws1.cell(row=89, column=c_i).border = thin_border
    ws1.cell(row=90, column=2, value="Total Biaya").font = font_bold
    ws1.cell(row=90, column=2).border = thin_border
    ws1.cell(row=90, column=6, value="=SUM(F89)").font = font_bold
    ws1.cell(row=90, column=6).number_format = '#,##0'
    ws1.cell(row=90, column=6).border = thin_border

    # -------------------------------------------------------------
    # SHEET 2: CASH INFLOW & CASH OUTFLOW
    # -------------------------------------------------------------
    ws2['B2'] = "CASH INFLOW"
    ws2['B2'].font = Font(name="Times New Roman", size=12, bold=True)
    inflow_years = ["Cash Inflow", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"]
    for c_i, h in enumerate(inflow_years, start=2):
        cell = ws2.cell(row=3, column=c_i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.border = thin_border

    products = cfg.get("products", [])
    ws2['B4'] = "Penjualan Produk Utama"
    ws2['B4'].font = font_bold

    # Product Inflow rows (B6 to B19)
    for p_idx, prod in enumerate(products):
        r = 6 + p_idx
        ws2.cell(row=r, column=2, value=prod["name"]).font = font_normal
        ws2.cell(row=r, column=2).border = thin_border
        for yr_idx in range(5):
            c_letter = get_column_letter(3 + yr_idx)
            ref_row = 149 + p_idx
            c_val = ws2.cell(row=r, column=3 + yr_idx, value=f"={c_letter}{ref_row}")
            c_val.font = font_normal
            c_val.number_format = '#,##0'
            c_val.border = thin_border

    ws2['B20'] = "Total"
    ws2['B20'].font = font_bold
    ws2['B20'].border = thin_border
    for yr_idx in range(5):
        c_letter = get_column_letter(3 + yr_idx)
        c_tot = ws2.cell(row=20, column=3 + yr_idx, value=f"=SUM({c_letter}6:{c_letter}19)")
        c_tot.font = font_bold
        c_tot.number_format = '#,##0'
        c_tot.border = thin_border

    # Target Penjualan Harian
    ws2['B23'] = "Target Penjualan Harian"
    ws2['B23'].font = Font(name="Times New Roman", size=12, bold=True)
    for c_i, h in enumerate(["Target Penjualan Harian", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"], start=2):
        ws2.cell(row=23, column=c_i, value=h).font = font_header
        ws2.cell(row=23, column=c_i).fill = header_fill
        ws2.cell(row=23, column=c_i).border = thin_border

    for p_idx, prod in enumerate(products):
        r = 26 + p_idx
        ws2.cell(row=r, column=2, value=prod["name"]).font = font_normal
        ws2.cell(row=r, column=2).border = thin_border
        sales = prod.get("daily_sales", [10, 12, 14, 16, 18])
        for yr_idx in range(5):
            val = sales[yr_idx] if yr_idx < len(sales) else sales[-1]
            c_val = ws2.cell(row=r, column=3 + yr_idx, value=val)
            c_val.font = font_normal
            c_val.alignment = Alignment(horizontal="center")
            c_val.border = thin_border

    # Harga naik per tahunnya (B42 to B100)
    ws2['B42'] = "Harga Produk naik per tahunnya"
    ws2['B42'].font = font_bold
    for p_idx, prod in enumerate(products):
        base_r = 43 + (p_idx * 5)
        ws2.cell(row=base_r, column=2, value=prod["name"]).font = font_bold
        ws2.cell(row=base_r+1, column=2, value="Harga").font = font_normal
        ws2.cell(row=base_r+2, column=2, value="Kenaikan Harga").font = font_normal
        ws2.cell(row=base_r+3, column=2, value="Harga Baru").font = font_normal
        ws2.cell(row=base_r+4, column=2, value="Pembulatan Harga").font = font_bold

        p1 = prod.get("price_year_1", 20000)
        rates = prod.get("price_increase_rates", [0.0, 0.10, 0.12, 0.13, 0.15])
        
        # Year 1
        ws2.cell(row=base_r+1, column=3, value=p1).number_format = '#,##0'
        ws2.cell(row=base_r+4, column=3, value=f"=C{base_r+1}").number_format = '#,##0'

        # Years 2 to 5
        for yr in range(1, 5):
            c_cur = get_column_letter(3 + yr)
            c_prev = get_column_letter(3 + yr - 1)
            ws2.cell(row=base_r+1, column=3 + yr, value=f"={c_prev}{base_r+4}").number_format = '#,##0'
            ws2.cell(row=base_r+2, column=3 + yr, value=rates[yr]).number_format = '0.0%'
            ws2.cell(row=base_r+3, column=3 + yr, value=f"={c_cur}{base_r+1}*(1+{c_cur}{base_r+2})").number_format = '#,##0'
            # Simple rounded to nearest 1000
            ws2.cell(row=base_r+4, column=3 + yr, value=f"=ROUNDUP({c_cur}{base_r+3}, -3)").number_format = '#,##0'

    # Pendapatan Bulanan (B120 to B144) & Tahunan (B146 to B163)
    ws2['B128'] = "Pendapatan Bulanan (25 Hari Kerja)"
    ws2['B128'].font = font_bold
    for p_idx, prod in enumerate(products):
        r_mo = 130 + p_idx
        r_yr = 149 + p_idx
        ws2.cell(row=r_mo, column=2, value=prod["name"]).font = font_normal
        ws2.cell(row=r_yr, column=2, value=prod["name"]).font = font_normal
        for yr in range(5):
            c_l = get_column_letter(3 + yr)
            qty_cell = f"{c_l}{26 + p_idx}"
            prc_cell = f"{c_l}{43 + (p_idx*5) + 4}"
            ws2.cell(row=r_mo, column=3 + yr, value=f"={qty_cell}*{prc_cell}*25").number_format = '#,##0'
            ws2.cell(row=r_yr, column=3 + yr, value=f"={c_l}{r_mo}*12").number_format = '#,##0'

    # CASH OUTFLOW (B165 to B173)
    ws2['B165'] = "CASH OUTFLOW"
    ws2['B165'].font = Font(name="Times New Roman", size=12, bold=True)
    for c_i, h in enumerate(["Cash Outflow", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4", "Tahun 5"], start=2):
        ws2.cell(row=166, column=c_i, value=h).font = font_header
        ws2.cell(row=166, column=c_i).fill = header_fill
        ws2.cell(row=166, column=c_i).border = thin_border

    outflow_cats = [
        ("Gaji Karyawan", "=C199", "=D199", "=E199", "=F199", "=G199"),
        ("Pembelian bahan baku", "=C257", "=D257", "=E257", "=F257", "=G257"),
        ("Listrik", "=C268", "=D268", "=E268", "=F268", "=G268"),
        ("Isi ulang tabung gas", "=30000*30*12", "=30000*30*12", "=30000*30*12", "=30000*30*12", "=30000*30*12"),
        ("Air", "=C279", "=D279", "=E279", "=F279", "=G279"),
        ("Air minum", "=C290", "=D290", "=E290", "=F290", "=G290")
    ]
    for idx, row_data in enumerate(outflow_cats, start=167):
        ws2.cell(row=idx, column=2, value=row_data[0]).font = font_normal
        ws2.cell(row=idx, column=2).border = thin_border
        for yr in range(5):
            c_val = ws2.cell(row=idx, column=3 + yr, value=row_data[1 + yr])
            c_val.font = font_normal
            c_val.number_format = '#,##0'
            c_val.border = thin_border

    ws2['B173'] = "Total"
    ws2['B173'].font = font_bold
    ws2['B173'].border = thin_border
    for yr in range(5):
        c_l = get_column_letter(3 + yr)
        c_tot = ws2.cell(row=173, column=3 + yr, value=f"=SUM({c_l}167:{c_l}172)")
        c_tot.font = font_bold
        c_tot.number_format = '#,##0'
        c_tot.border = thin_border

    # Gaji Detail (Row 176 - 199)
    ws2['B176'] = "Perincian Gaji Karyawan"
    ws2['B176'].font = font_bold
    ws2.cell(row=176, column=3, value="Jumlah Karyawan").font = font_header
    ws2.cell(row=176, column=4, value="Gaji Bulanan").font = font_header
    ws2.cell(row=176, column=5, value="Total Gaji Tahunan").font = font_header
    
    emp_list = cfg["operational_costs"].get("employees", [])
    for e_i, emp in enumerate(emp_list):
        r = 177 + e_i
        ws2.cell(row=r, column=2, value=emp["role"]).font = font_normal
        ws2.cell(row=r, column=3, value=emp["count"]).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=4, value=emp["monthly_salary"]).number_format = '#,##0'
        ws2.cell(row=r, column=5, value=f"=C{r}*D{r}*12").number_format = '#,##0'

    ws2['B183'] = "Total Gaji"
    ws2['B183'].font = font_bold
    ws2.cell(row=183, column=5, value="=SUM(E177:E182)").font = font_bold
    ws2.cell(row=183, column=5).number_format = '#,##0'

    # Kenaikan Gaji Tahunan
    ws2['B186'] = "Gaji Karyawan Tahunan (Kenaikan per tahun)"
    ws2['B186'].font = font_bold
    ws2.cell(row=187, column=2, value="Total Gaji").font = font_normal
    ws2.cell(row=187, column=3, value="=E183").number_format = '#,##0'
    for yr in range(1, 5):
        c_cur = get_column_letter(3 + yr)
        c_prev = get_column_letter(3 + yr - 1)
        rate = 0.10 + (yr * 0.02)
        ws2.cell(row=188, column=3 + yr, value=rate).number_format = '0.0%'
        ws2.cell(row=187, column=3 + yr, value=f"={c_prev}187*(1+{c_cur}188)").number_format = '#,##0'

    ws2.cell(row=199, column=2, value="Gaji Summary").font = font_bold
    for yr in range(5):
        c_l = get_column_letter(3 + yr)
        ws2.cell(row=199, column=3 + yr, value=f"={c_l}187").number_format = '#,##0'

    # Bahan Baku Detail (Row 250 - 257)
    raw_base = cfg["operational_costs"].get("raw_materials_monthly_base", 20000000)
    ws2.cell(row=250, column=2, value="Biaya bahan baku bulanan").font = font_normal
    ws2.cell(row=250, column=3, value=raw_base).number_format = '#,##0'
    ws2.cell(row=251, column=2, value="Biaya bahan baku tahunan").font = font_normal
    ws2.cell(row=251, column=3, value="=C250*12").number_format = '#,##0'

    ws2.cell(row=254, column=2, value="Kenaikan bahan baku per tahun").font = font_bold
    ws2.cell(row=255, column=3, value="=C251").number_format = '#,##0'
    ws2.cell(row=257, column=3, value="=C255").number_format = '#,##0'
    for yr in range(1, 5):
        c_cur = get_column_letter(3 + yr)
        c_prev = get_column_letter(3 + yr - 1)
        rate = 0.12 + (yr * 0.03)
        ws2.cell(row=256, column=3 + yr, value=rate).number_format = '0.0%'
        ws2.cell(row=255, column=3 + yr, value=f"={c_prev}257").number_format = '#,##0'
        ws2.cell(row=257, column=3 + yr, value=f"={c_cur}255*(1+{c_cur}256)").number_format = '#,##0'

    # Listrik (Row 260 - 268)
    elec_base = cfg["operational_costs"].get("electricity_monthly_base", 2500000)
    ws2.cell(row=260, column=2, value="Penggunaan listrik bulanan").font = font_normal
    ws2.cell(row=260, column=3, value=elec_base).number_format = '#,##0'
    ws2.cell(row=261, column=2, value="Penggunaan listrik tahunan").font = font_normal
    ws2.cell(row=261, column=3, value="=C260*12").number_format = '#,##0'
    ws2.cell(row=266, column=3, value="=C261").number_format = '#,##0'
    ws2.cell(row=268, column=3, value="=C266").number_format = '#,##0'
    for yr in range(1, 5):
        c_cur = get_column_letter(3 + yr)
        c_prev = get_column_letter(3 + yr - 1)
        rate = 0.10 + (yr * 0.02)
        ws2.cell(row=267, column=3 + yr, value=rate).number_format = '0.0%'
        ws2.cell(row=266, column=3 + yr, value=f"={c_prev}268").number_format = '#,##0'
        ws2.cell(row=268, column=3 + yr, value=f"={c_cur}266*(1+{c_cur}267)").number_format = '#,##0'

    # Air PAM (Row 271 - 279)
    water_base = cfg["operational_costs"].get("water_monthly_base", 600000)
    ws2.cell(row=271, column=2, value="Penggunaan air bulanan").font = font_normal
    ws2.cell(row=271, column=3, value=water_base).number_format = '#,##0'
    ws2.cell(row=272, column=2, value="Penggunaan air tahunan").font = font_normal
    ws2.cell(row=272, column=3, value="=C271*12").number_format = '#,##0'
    ws2.cell(row=277, column=3, value="=C272").number_format = '#,##0'
    ws2.cell(row=279, column=3, value="=C277").number_format = '#,##0'
    for yr in range(1, 5):
        c_cur = get_column_letter(3 + yr)
        c_prev = get_column_letter(3 + yr - 1)
        rate = 0.10 + (yr * 0.02)
        ws2.cell(row=278, column=3 + yr, value=rate).number_format = '0.0%'
        ws2.cell(row=277, column=3 + yr, value=f"={c_prev}279").number_format = '#,##0'
        ws2.cell(row=279, column=3 + yr, value=f"={c_cur}277*(1+{c_cur}278)").number_format = '#,##0'

    # Air Galon (Row 283 - 290)
    ws2.cell(row=283, column=2, value="Air Minum Galon Bulanan").font = font_normal
    ws2.cell(row=283, column=3, value=400000).number_format = '#,##0'
    ws2.cell(row=284, column=2, value="Air Minum Galon Tahunan").font = font_normal
    ws2.cell(row=284, column=3, value="=C283*12").number_format = '#,##0'
    ws2.cell(row=288, column=3, value=20000).number_format = '#,##0'
    ws2.cell(row=289, column=3, value=240).alignment = Alignment(horizontal="center")
    ws2.cell(row=290, column=3, value="=C288*C289").number_format = '#,##0'
    for yr in range(1, 5):
        c_cur = get_column_letter(3 + yr)
        c_prev = get_column_letter(3 + yr - 1)
        ws2.cell(row=288, column=3 + yr, value=20000).number_format = '#,##0'
        ws2.cell(row=289, column=3 + yr, value=f"={c_prev}289+5").alignment = Alignment(horizontal="center")
        ws2.cell(row=290, column=3 + yr, value=f"={c_cur}288*{c_cur}289").number_format = '#,##0'

    # -------------------------------------------------------------
    # SHEET 3: EVALUASI INVESTASI (NPV, IRR, PP, PI)
    # -------------------------------------------------------------
    ws3['B3'] = "Initial Outlay"
    ws3['B3'].font = Font(name="Times New Roman", size=12, bold=True)
    recap_outlay = [
        ("Aktiva Tetap", "=Sheet1!C14"),
        ("Biaya Sewa", "=Sheet1!C15"),
        ("Biaya Amdal", "=Sheet1!C16"),
        ("Biaya Survey", "=Sheet1!C17"),
        ("Biaya Iklan", "=Sheet1!C18")
    ]
    for i, (l, f_str) in enumerate(recap_outlay, start=4):
        ws3.cell(row=i, column=2, value=l).font = font_normal
        ws3.cell(row=i, column=2).border = thin_border
        c_v = ws3.cell(row=i, column=3, value=f_str)
        c_v.font = font_normal
        c_v.number_format = '#,##0'
        c_v.border = thin_border

    ws3['B9'] = "Total"
    ws3['B9'].font = font_bold
    ws3['B9'].border = thin_border
    c_tot3 = ws3.cell(row=9, column=3, value="=SUM(C4:C8)")
    c_tot3.font = font_bold
    c_tot3.number_format = '#,##0'
    c_tot3.border = thin_border

    ws3['B12'] = "Tahun"
    ws3['C12'] = "Laba / Proceed"
    ws3['B12'].font = font_header
    ws3['C12'].font = font_header
    ws3['B12'].fill = header_fill
    ws3['C12'].fill = header_fill

    sheet1_proceed_cols = ['D', 'E', 'F', 'G', 'H']
    for yr in range(1, 6):
        r = 12 + yr
        ws3.cell(row=r, column=2, value=yr).alignment = Alignment(horizontal="center")
        c_p = ws3.cell(row=r, column=3, value=f"=Sheet1!{sheet1_proceed_cols[yr-1]}10")
        c_p.font = font_normal
        c_p.number_format = '#,##0'
        ws3.cell(row=r, column=2).border = thin_border
        c_p.border = thin_border

    ws3['B19'] = "Opportunity Cost"
    ws3['B19'].font = font_bold
    ws3.cell(row=19, column=3, value=0.20).number_format = '0.0%'

    # NPV Block
    ws3['B26'] = "NPV"
    ws3['B26'].font = Font(name="Times New Roman", size=12, bold=True)
    ws3['B27'] = "Tahun"
    ws3['C27'] = "Cashflow"
    ws3['D27'] = "Present Value (PV)"
    for c_i in [2, 3, 4]:
        ws3.cell(row=27, column=c_i).font = font_header
        ws3.cell(row=27, column=c_i).fill = header_fill
        ws3.cell(row=27, column=c_i).border = thin_border

    # Year 0
    ws3.cell(row=28, column=2, value=0).alignment = Alignment(horizontal="center")
    ws3.cell(row=28, column=3, value="=-C9").number_format = '#,##0'
    ws3.cell(row=28, column=4, value="=PV($C$19,B28,,-C28,)").number_format = '#,##0'
    for c_i in [2, 3, 4]:
        ws3.cell(row=28, column=c_i).font = font_normal
        ws3.cell(row=28, column=c_i).border = thin_border

    # Years 1-5
    for yr in range(1, 6):
        r = 28 + yr
        ws3.cell(row=r, column=2, value=yr).alignment = Alignment(horizontal="center")
        ws3.cell(row=r, column=3, value=f"=C{12+yr}").number_format = '#,##0'
        ws3.cell(row=r, column=4, value=f"=PV($C$19,B{r},,-C{r},)").number_format = '#,##0'
        for c_i in [2, 3, 4]:
            ws3.cell(row=r, column=c_i).font = font_normal
            ws3.cell(row=r, column=c_i).border = thin_border

    ws3['B34'] = "NPV"
    ws3['B34'].font = font_bold
    ws3['B34'].border = thin_border
    c_npv = ws3.cell(row=34, column=4, value="=SUM(D28:D33)")
    c_npv.font = font_bold
    c_npv.number_format = '#,##0'
    c_npv.border = thin_border

    # IRR Block
    ws3['B36'] = "IRR"
    ws3['B36'].font = Font(name="Times New Roman", size=12, bold=True)
    ws3['B37'] = "Tahun"
    ws3['C37'] = "Cashflow"
    ws3['B37'].font = font_header
    ws3['C37'].font = font_header
    ws3['B37'].fill = header_fill
    ws3['C37'].fill = header_fill

    ws3.cell(row=38, column=2, value=0).alignment = Alignment(horizontal="center")
    ws3.cell(row=38, column=3, value="=C28").number_format = '#,##0'
    for yr in range(1, 6):
        r = 38 + yr
        ws3.cell(row=r, column=2, value=yr).alignment = Alignment(horizontal="center")
        ws3.cell(row=r, column=3, value=f"=C{28+yr}").number_format = '#,##0'
        ws3.cell(row=r, column=2).border = thin_border
        ws3.cell(row=r, column=3).border = thin_border

    ws3['B44'] = "IRR"
    ws3['B44'].font = font_bold
    ws3['B44'].border = thin_border
    c_irr = ws3.cell(row=44, column=3, value="=IRR(C38:C43)")
    c_irr.font = font_bold
    c_irr.number_format = '0.0%'
    c_irr.border = thin_border

    # Payback Period Block
    ws3['B46'] = "Payback Period"
    ws3['B46'].font = Font(name="Times New Roman", size=12, bold=True)
    pp_hdrs = ["Tahun", "Cashflow", "Akumulasi", "Payback Period (Bulan)"]
    for c_i, h in enumerate(pp_hdrs, start=2):
        cell = ws3.cell(row=47, column=c_i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.border = thin_border

    ws3.cell(row=48, column=2, value=0).alignment = Alignment(horizontal="center")
    ws3.cell(row=48, column=3, value="=C38").number_format = '#,##0'
    ws3.cell(row=48, column=4, value="=C48").number_format = '#,##0'
    for c_i in range(2, 6):
        ws3.cell(row=48, column=c_i).border = thin_border
        ws3.cell(row=48, column=c_i).font = font_normal

    for yr in range(1, 6):
        r = 48 + yr
        ws3.cell(row=r, column=2, value=yr).alignment = Alignment(horizontal="center")
        ws3.cell(row=r, column=3, value=f"=C{38+yr}").number_format = '#,##0'
        ws3.cell(row=r, column=4, value=f"=D{r-1}+C{r}").number_format = '#,##0'
        if yr == 1:
            ws3.cell(row=r, column=5, value=f"=-D48/(C{r}/12)").number_format = '0.0'
        for c_i in range(2, 6):
            ws3.cell(row=r, column=c_i).border = thin_border
            ws3.cell(row=r, column=c_i).font = font_normal

    # Profitability Index
    ws3['B55'] = "Profitability Index"
    ws3['B55'].font = font_bold
    ws3['B55'].border = thin_border
    c_pi = ws3.cell(row=55, column=3, value="=SUM(C49:C53)/-C48")
    c_pi.font = font_bold
    c_pi.number_format = '0.00'
    c_pi.border = thin_border

    # Tabel Kesimpulan Kelayakan
    ws3['B58'] = "Kesimpulan"
    ws3['C58'] = "Hasil"
    ws3['D58'] = "Target"
    ws3['E58'] = "Kelayakan"
    for c_i in range(2, 6):
        ws3.cell(row=58, column=c_i).font = font_header
        ws3.cell(row=58, column=c_i).fill = header_fill
        ws3.cell(row=58, column=c_i).border = thin_border

    kriteria = [
        ("NPV", "=D34", "Positif (> 0)", "Layak"),
        ("IRR", "=C44", "> 20%", "Layak"),
        ("PP", "=E49", "< 3 Tahun (36 Bulan)", "Layak"),
        ("PI", "=C55", "> 1.2", "Layak")
    ]
    for idx, (crit, form, tgt, stat) in enumerate(kriteria, start=59):
        ws3.cell(row=idx, column=2, value=crit).font = font_bold
        c_h = ws3.cell(row=idx, column=3, value=form)
        c_h.font = font_bold
        if crit == "NPV":
            c_h.number_format = '#,##0'
        elif crit == "IRR":
            c_h.number_format = '0.0%'
        elif crit == "PP":
            c_h.number_format = '0.0'
        elif crit == "PI":
            c_h.number_format = '0.00'

        ws3.cell(row=idx, column=4, value=tgt).font = font_normal
        ws3.cell(row=idx, column=5, value=stat).font = font_bold
        for c_i in range(2, 6):
            ws3.cell(row=idx, column=c_i).border = thin_border

    # Auto-adjust column widths for all sheets
    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_path)
    print(f"[SUCCESS] Model finansial Excel berhasil dibuat di: {output_path}")

if __name__ == "__main__":
    cfg_file = sys.argv[1] if len(sys.argv) > 1 else r"d:\Perkuliahan\Kelass\SKB\.agents\skills\skb-format-writer\templates\sample_business_config.json"
    out_file = sys.argv[2] if len(sys.argv) > 2 else r"d:\Perkuliahan\Kelass\SKB\Model_Finansial_Template.xlsx"
    build_skb_excel(cfg_file, out_file)
