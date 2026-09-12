import json
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

REPORTS_DIR = Path(r"c:\Users\fioni\Documents\audioetria yara\Cait panama nuevo\data\reports")
with open(REPORTS_DIR / "pedregal_extracted.json", "r", encoding="utf-8") as f:
    patients = json.load(f)

# Sort by Area then by Name
area_order = ["Pollo Vivo", "Desplume", "Patitas", "Lavado de Canasta", "Evisceración"]
def sort_key(p):
    area = p.get("position", "")
    idx = area_order.index(area) if area in area_order else 99
    return (idx, p.get("clean_name", ""))

patients_sorted = sorted(patients, key=sort_key)

out_pdf = Path(r"c:\Users\fioni\Downloads\resultados\pedregal\LISTADO DE ASISTENCIA PEDREGAL DIGITAL.pdf")
out_att = Path(r"c:\Users\fioni\Documents\audioetria yara\Cait panama nuevo\data\attachments\report_adjuntos\LISTADO DE ASISTENCIA PEDREGAL DIGITAL.pdf")

doc = SimpleDocTemplate(
    str(out_pdf),
    pagesize=letter,
    leftMargin=36,
    rightMargin=36,
    topMargin=36,
    bottomMargin=36
)

styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'DocTitle',
    parent=styles['Heading1'],
    fontName='Helvetica-Bold',
    fontSize=16,
    leading=20,
    textColor=colors.HexColor('#002855'),
    alignment=1
)
subtitle_style = ParagraphStyle(
    'DocSubTitle',
    parent=styles['Normal'],
    fontName='Helvetica',
    fontSize=10,
    leading=14,
    textColor=colors.HexColor('#475569'),
    alignment=1
)
cell_bold = ParagraphStyle(
    'CellBold',
    parent=styles['Normal'],
    fontName='Helvetica-Bold',
    fontSize=8,
    leading=10,
    textColor=colors.HexColor('#1e293b')
)
cell_normal = ParagraphStyle(
    'CellNormal',
    parent=styles['Normal'],
    fontName='Helvetica',
    fontSize=8,
    leading=10,
    textColor=colors.HexColor('#334155')
)
header_cell = ParagraphStyle(
    'HeaderCell',
    parent=styles['Normal'],
    fontName='Helvetica-Bold',
    fontSize=8,
    leading=10,
    textColor=colors.white,
    alignment=1
)

elements = []
elements.append(Paragraph("LISTADO GENERAL DE ASISTENCIA Y EVALUACIÓN", title_style))
elements.append(Spacer(1, 4))
elements.append(Paragraph("PRODUCTOS TOLEDANO S.A. — PLANTA PEDREGAL (PAPSA 2026)<br/>Pruebas de Función Pulmonar / Espirometría Ocupacional", subtitle_style))
elements.append(Spacer(1, 12))

table_data = [[
    Paragraph("N°", header_cell),
    Paragraph("Área / Departamento", header_cell),
    Paragraph("Nombre y Apellidos", header_cell),
    Paragraph("Cédula", header_cell),
    Paragraph("Edad", header_cell),
    Paragraph("Resultado Espirometría", header_cell),
    Paragraph("Estado", header_cell)
]]

for idx, p in enumerate(patients_sorted, 1):
    res_color = "#15803d" if "normal" in p["result"].lower() else "#b91c1c"
    res_style = ParagraphStyle(
        f'Res_{idx}',
        parent=cell_bold,
        textColor=colors.HexColor(res_color)
    )
    table_data.append([
        Paragraph(str(idx), cell_normal),
        Paragraph(p["position"], cell_bold),
        Paragraph(p["clean_name"], cell_normal),
        Paragraph(p["cedula"], cell_bold),
        Paragraph(p["age"], cell_normal),
        Paragraph(p["result"], res_style),
        Paragraph("Firmado / Evaluado", cell_normal)
    ])

t = Table(table_data, colWidths=[24, 95, 145, 80, 30, 110, 56], repeatRows=1)
t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#002855')),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('TOPPADDING', (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')])
]))

elements.append(t)
doc.build(elements)

import shutil
shutil.copy2(out_pdf, out_att)
print("Digital Attendance PDF created!")
