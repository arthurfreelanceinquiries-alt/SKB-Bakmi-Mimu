import os

with open(r'd:\Perkuliahan\Kelass\SKB\generate_bakmi_mimu_pptx.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace Slide 1 Cover to include UKRIDA logo and all 5 authors
old_s1 = """    # Accent decorative box
    dec = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1))
    dec.fill.solid()
    dec.fill.fore_color.rgb = CARD_BG
    dec.line.color.rgb = GOLD_ACCENT
    dec.line.width = Pt(1.5)

    tb1 = s1.shapes.add_textbox(Inches(1.8), Inches(1.8), Inches(9.7), Inches(3.8))
    tf1 = tb1.text_frame
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = "STUDI KELAYAKAN BISNIS (SKB)"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(18)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT

    p2 = tf1.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = f"“ {cfg['business_name'].upper()} ”"
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = WHITE

    p3 = tf1.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.text = f"- {cfg['tagline']} -"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(16)
    p3.font.italic = True
    p3.font.color.rgb = GRAY_TEXT

    p4 = tf1.add_paragraph()
    p4.alignment = PP_ALIGN.CENTER
    p4.text = f"\\n\\nDisusun Oleh: {cfg['author']['name']} (NIM: {cfg['author']['nim']})\\n{cfg['author']['institution']} | {cfg['author']['faculty']} | {cfg['author']['year']}"
    p4.font.name = "Times New Roman"
    p4.font.size = Pt(13)
    p4.font.color.rgb = WHITE"""

new_s1 = """    # Accent decorative box
    dec = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(0.6), Inches(10.933), Inches(6.3))
    dec.fill.solid()
    dec.fill.fore_color.rgb = CARD_BG
    dec.line.color.rgb = GOLD_ACCENT
    dec.line.width = Pt(1.5)

    # UKRIDA Logo
    logo_path = r"d:\\Perkuliahan\\Kelass\\SKB\\logo_ukrida.png"
    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(6.066), Inches(0.85), Inches(1.2), Inches(1.2))

    tb1 = s1.shapes.add_textbox(Inches(1.5), Inches(2.15), Inches(10.333), Inches(4.5))
    tf1 = tb1.text_frame
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = "STUDI KELAYAKAN BISNIS (SKB)"
    p1.font.name = "Times New Roman"
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = GOLD_ACCENT

    p2 = tf1.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = f"“ {cfg['business_name'].upper()} ”"
    p2.font.name = "Times New Roman"
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.color.rgb = WHITE

    p3 = tf1.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.text = f"- {cfg['tagline']} -"
    p3.font.name = "Times New Roman"
    p3.font.size = Pt(13)
    p3.font.italic = True
    p3.font.color.rgb = GRAY_TEXT

    p_lbl = tf1.add_paragraph()
    p_lbl.alignment = PP_ALIGN.CENTER
    p_lbl.text = "\\nDisusun Oleh Kelompok Mahasiswa:"
    p_lbl.font.name = "Times New Roman"
    p_lbl.font.size = Pt(11)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = GOLD_ACCENT

    p_m1 = tf1.add_paragraph()
    p_m1.alignment = PP_ALIGN.CENTER
    p_m1.text = "Arthur Reezan (312023002)   •   Jennese Putra Alamsyah Sukadi (312023033)   •   Valendrik Dwiputra Wirawan (312023013)"
    p_m1.font.name = "Times New Roman"
    p_m1.font.size = Pt(11)
    p_m1.font.color.rgb = WHITE

    p_m2 = tf1.add_paragraph()
    p_m2.alignment = PP_ALIGN.CENTER
    p_m2.text = "Affandy (312023075)   •   Steven Putra Tjhin (312023015)"
    p_m2.font.name = "Times New Roman"
    p_m2.font.size = Pt(11)
    p_m2.font.color.rgb = WHITE

    p_inst = tf1.add_paragraph()
    p_inst.alignment = PP_ALIGN.CENTER
    p_inst.text = "\\nProgram Studi Manajemen   |   Fakultas Ekonomi & Bisnis   |   Universitas Kristen Krida Wacana (UKRIDA)   |   2024"
    p_inst.font.name = "Times New Roman"
    p_inst.font.size = Pt(10.5)
    p_inst.font.italic = True
    p_inst.font.color.rgb = GRAY_TEXT"""

assert old_s1 in content, "old_s1 not found in content"
content = content.replace(old_s1, new_s1)

# Replace Slide 22 Closing to include UKRIDA logo
old_s22 = """    dec22 = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1))
    dec22.fill.solid()
    dec22.fill.fore_color.rgb = CARD_BG
    dec22.line.color.rgb = GREEN_ACCENT
    dec22.line.width = Pt(2.0)

    tb22 = s22.shapes.add_textbox(Inches(1.8), Inches(1.6), Inches(9.7), Inches(4.3))"""

new_s22 = """    dec22 = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(0.6), Inches(10.933), Inches(6.3))
    dec22.fill.solid()
    dec22.fill.fore_color.rgb = CARD_BG
    dec22.line.color.rgb = GREEN_ACCENT
    dec22.line.width = Pt(2.0)

    logo_path = r"d:\\Perkuliahan\\Kelass\\SKB\\logo_ukrida.png"
    if os.path.exists(logo_path):
        s22.shapes.add_picture(logo_path, Inches(6.066), Inches(0.85), Inches(1.2), Inches(1.2))

    tb22 = s22.shapes.add_textbox(Inches(1.5), Inches(2.15), Inches(10.333), Inches(4.5))"""

assert old_s22 in content, "old_s22 not found in content"
content = content.replace(old_s22, new_s22)

# Output path
old_out = 'out_pptx = r"d:\\Perkuliahan\\Kelass\\SKB\\Bakmi_Mimu_Presentasi_SKB.pptx"'
new_out = 'out_pptx = r"d:\\Perkuliahan\\Kelass\\SKB\\01_TUGAS_FINAL_BAKMI_MIMU\\Presentasi_SKB_Bakmi_Mimu_Carina_Sayang.pptx"'
assert old_out in content, "old_out not found in content"
content = content.replace(old_out, new_out)

with open(r'd:\Perkuliahan\Kelass\SKB\04_INTERNAL_TOOLS_DAN_DATA\generate_master_pptx.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("[OK] Master PPTX generator script saved at 04_INTERNAL_TOOLS_DAN_DATA/generate_master_pptx.py")
