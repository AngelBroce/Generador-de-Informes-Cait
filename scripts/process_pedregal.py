import os
import re
import json
import shutil
import zipfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import openpyxl
import pypdf

# Paths
WORKSPACE_ROOT = Path(r"c:\Users\fioni\Documents\audioetria yara\Cait panama nuevo")
PEDREGAL_ROOT = Path(r"c:\Users\fioni\Downloads\resultados\pedregal")
ARCHIVE_ROOT = PEDREGAL_ROOT / "archive"
LISTA_DIR = ARCHIVE_ROOT / "lista de asistencia"
FIRMADAS_DIR = ARCHIVE_ROOT / "FIRMADAS" / "FIRMADAS"
EXCEL_PATH = ARCHIVE_ROOT / "cuadro pacientes TOLEDANO Pedregal.xlsx"
CALIB_PDF = ARCHIVE_ROOT / "Carta-de-Certificacion de calibracion.pdf"

ATTACHMENTS_DIR = WORKSPACE_ROOT / "data" / "attachments" / "report_adjuntos"
REPORTS_DIR = WORKSPACE_ROOT / "data" / "reports"
DATABASES_DIR = WORKSPACE_ROOT / "data" / "databases"

ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
DATABASES_DIR.mkdir(parents=True, exist_ok=True)

# 1. Load Excel altered data
wb = openpyxl.load_workbook(EXCEL_PATH)
sheet = wb.active

def normalize_text(t):
    return str(t).lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u').strip()

excel_cases = {}
for row in sheet.iter_rows(min_row=4, values_only=True):
    nombre, edad, resultado, comentario = row[0], row[1], row[2], row[3]
    if nombre and resultado and not str(resultado).startswith("Total"):
        clean_name = " ".join(str(nombre).strip().split()).title()
        excel_cases[normalize_text(clean_name)] = {
            "name": clean_name,
            "age": str(edad) if edad else "",
            "result": str(resultado).strip(),
            "comment": str(comentario).strip() if comentario else ""
        }

print(f"[1] Loaded {len(excel_cases)} special clinical cases from Excel.")

# 2. Map Area from attendance sheets
worker_areas = {}
# POLLO VIVO (25)
for n in [
    'Gemino Acosta', 'Ramiro Abrego', 'Ignacio Vasquez', 'Manuel Perez',
    'Anastacio Rodriguez', 'Manuel Degracia', 'Secundino Vanega', 'Isael Gonzalez',
    'Armando Morales', 'Ceferino Mendoza', 'Ariel Gil', 'Javier Abrego',
    'Elvin Cortez', 'Jose Ines Garcia', 'Jesus Bernal', 'Alberto Carpintero',
    'Leonaldo Jaen', 'Indalecio Vasquez', 'Roman Eugenio', 'Noriel Sam',
    'David Quintero', 'Tano Miranda', 'Guillermo Brisonte', 'Meliton Morales', 'Celestino Flores'
]:
    worker_areas[normalize_text(n)] = "Pollo Vivo"

# DESPLUME (3)
for n in ['Luis Palma', 'Aristides Alveo', 'Marvin Mendoza']:
    worker_areas[normalize_text(n)] = "Desplume"

# PATITAS (5)
for n in ['Jennifer Sanjur', 'Edgardo Gonzalez', 'Julio Bejerano', 'Gionela Zapata', 'Yesenia Sevillano']:
    worker_areas[normalize_text(n)] = "Patitas"

# LAVADO DE CANASTA (11)
for n in [
    'Alexis Pitti', 'Francisco Acosta', 'Renato Palacio', 'Josue Rugama', 'Abdiel Ramos', 'Alfredo Baker',
    'Rafael Anguizola', 'Asalon Morales', 'Maria de Jaramillo', 'Evangelista Pineda', 'Roger Zurdo'
]:
    worker_areas[normalize_text(n)] = "Lavado de Canasta"

# EVISCERACION (16)
for n in [
    'Eymiss Arcia', 'Pedro Solis', 'Eloy Smith', 'Shanira Crosdale', 'Tomas Marin',
    'Carlos Martinez', 'Maria del Carmen Gonzalez', 'Isidra Peralta', 'Argelia Ibarra',
    'Vilma Avila', 'Rosa Gonzalez', 'Katherine Saez', 'Ovidio Sanchez', 'Yesica Edwards',
    'Doris Gracia', 'Benedicto Abrego'
]:
    worker_areas[normalize_text(n)] = "Evisceración"

# Bidirectional mappings between attendance list names and PDF filenames
pdf_to_attendance = {
    normalize_text('ignacio vasuez'): normalize_text('ignacio vasquez'),
    normalize_text('manel de gracia'): normalize_text('manuel degracia'),
    normalize_text('arisitdes alveo'): normalize_text('aristides alveo'),
    normalize_text('guillermo brisonte'): normalize_text('guillermo brisonto'),
    normalize_text('luis enrique palma'): normalize_text('luis palma'),
    normalize_text('maria del carmen gonzalez'): normalize_text('maria del carmen gonzalez'),
}
attendance_to_pdf = {v: k for k, v in pdf_to_attendance.items()}
attendance_to_pdf[normalize_text('fracisco acosta')] = normalize_text('francisco acosta')
attendance_to_pdf[normalize_text('yesenia servillano')] = normalize_text('yesenia sevillano')
attendance_to_pdf[normalize_text('maria del c. gonzalez')] = normalize_text('maria del carmen gonzalez')


# 3. Read and extract all 60 PDFs
pdf_files = sorted(list(FIRMADAS_DIR.glob("*.pdf")))
print(f"[2] Found {len(pdf_files)} PDF files in FIRMADAS.")

patients = []
for f in pdf_files:
    reader = pypdf.PdfReader(f)
    text = reader.pages[0].extract_text()
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Cedula
    m_ced = re.search(r'C[oó\?d\.]*\s*paciente[:\s]*([0-9A-Za-z\-]+)', text, re.IGNORECASE)
    ced = m_ced.group(1).strip() if m_ced else ''

    # Age
    m_age = re.search(r'Edad\s*(\d+)', text, re.IGNORECASE)
    age = m_age.group(1).strip() if m_age else ''

    # Raw interpretation
    interp = ''
    for idx, l in enumerate(lines):
        if 'Interpretaci' in l:
            if idx + 1 < len(lines):
                interp = lines[idx+1].strip()
            break

    raw_name = f.stem.replace("Signed", "").strip()
    clean_name = " ".join(raw_name.split()).title()
    norm_n = normalize_text(clean_name)
    mapped_key = pdf_to_attendance.get(norm_n, norm_n)

    # Determine Area
    area = worker_areas.get(mapped_key, "")
    if not area:
        area = worker_areas.get(norm_n, "")
    if not area:
        for k, v in worker_areas.items():
            if mapped_key in k or k in mapped_key or norm_n in k or k in norm_n:
                area = v
                break

    # Determine Result from Excel
    matched_ex = None
    # 1. Direct key match
    for cand in [norm_n, mapped_key]:
        if cand in excel_cases:
            matched_ex = excel_cases[cand]
            break
    # 2. Fuzzy words match
    if not matched_ex:
        for k, v in excel_cases.items():
            if norm_n in k or k in norm_n or mapped_key in k or k in mapped_key:
                matched_ex = v
                break
            nw = set(norm_n.split()) | set(mapped_key.split())
            kw = set(k.split())
            if len(nw & kw) >= 2 or ('alveo' in nw and 'alveo' in kw):
                matched_ex = v
                break

    if matched_ex:
        res_raw = matched_ex["result"]
        if "restric" in res_raw.lower() and "leve" in res_raw.lower():
            result = "Restricción leve"
        elif "obstruc" in res_raw.lower() and "leve" in res_raw.lower():
            result = "Obstrucción leve"
        elif "moderada" in res_raw.lower() and "severa" not in res_raw.lower():
            result = "Restricción moderada"
        elif "severa" in res_raw.lower():
            result = "Restricción moderadamente severa"
        else:
            result = res_raw
        comment = matched_ex["comment"]
    else:
        result = "Espirometría normal"
        comment = ""

    signed_pdf_name = f"{clean_name} Signed.pdf"

    patients.append({
        "original_file": f.name,
        "signed_pdf_name": signed_pdf_name,
        "clean_name": clean_name,
        "cedula": ced,
        "age": age,
        "position": area,
        "result": result,
        "comment": comment
    })

# Sort patients alphabetically by clean_name
patients = sorted(patients, key=lambda p: normalize_text(p["clean_name"]))
print(f"[3] Extracted and mapped {len(patients)} patients.")

# 4. Generate Stamped Attendance PDF
try:
    font = ImageFont.truetype("arialbd.ttf", 24)
except Exception:
    font = ImageFont.load_default()

by_ced = {p["cedula"]: p for p in patients}
by_name_dict = {normalize_text(p["clean_name"]): p for p in patients}
for att_name, pdf_name in attendance_to_pdf.items():
    if pdf_name in by_name_dict:
        by_name_dict[att_name] = by_name_dict[pdf_name]

def stamp_sheet(img_path, stamp_list):
    im = Image.open(img_path).convert('RGB')
    draw = ImageDraw.Draw(im)
    for key, x, y in stamp_list:
        norm_k = normalize_text(key)
        p = by_name_dict.get(norm_k)
        if not p:
            pdf_k = attendance_to_pdf.get(norm_k, norm_k)
            p = by_name_dict.get(pdf_k)
        if not p:
            for k, v in by_name_dict.items():
                if norm_k in k or k in norm_k:
                    p = v
                    break
        if not p:
            print(f"WARN: Could not find patient for '{key}'")
            continue
        ced = p["cedula"]
        bbox = draw.textbbox((x, y), ced, font=font)
        pad_x, pad_y = 6, 2
        rect = [bbox[0] - pad_x, bbox[1] - pad_y, bbox[2] + pad_x, bbox[3] + pad_y]
        draw.rounded_rectangle(rect, radius=4, fill=(245, 248, 255), outline=(0, 60, 140), width=2)
        draw.text((x, y), ced, fill=(0, 35, 100), font=font)
    return im

# Stamping definitions with exact row alignment
sheet1_stamps = [
    ('gemino acosta', 820, 380),
    ('ramiro abrego', 820, 440),
    ('ignacio vasquez', 820, 500),
    ('manuel perez', 820, 560),
    # Agripino Gonzalez is at ~600 (not evaluated)
    ('anastacio rodriguez', 820, 650),
    ('manuel degracia', 820, 695),
    ('secundino vanega', 820, 738),
    ('isael gonzalez', 820, 785),
    ('armando morales', 820, 830),
    ('ceferino mendoza', 820, 875),
    ('ariel gil', 820, 920),
    ('javier abrego', 820, 965),
    ('elvin cortez', 820, 1010),
    ('jose ines garcia', 820, 1055),
    ('jesus bernal', 820, 1100),
    ('alberto carpintero', 820, 1145),
    ('leonaldo jaen', 820, 1190),
    ('indalecio vasquez', 820, 1235),
    # Jose Ines Castillo is at ~1280 (not evaluated)
    ('roman eugenio', 820, 1330),
    ('noriel sam', 820, 1375),
    ('david quintero', 820, 1420),
    # Bottom handwritten
    ('tano miranda', 470, 1485),
    ('guillermo brisonto', 540, 1545),
    ('meliton morales', 950, 1485),
    ('celestino flores', 970, 1545),
]

sheet2_stamps = [
    ('luis palma', 720, 205),
    ('aristides alveo', 720, 260),
    ('marvin mendoza', 720, 350),
    ('jennifer sanjur', 650, 555),
    ('edgardo gonzalez', 650, 675),
    ('julio bejerano', 650, 735),
    ('gionela zapata', 650, 795),
    ('yesenia sevillano', 650, 855),
    ('alexis pitti', 760, 1105),
    ('fracisco acosta', 760, 1185),
    ('renato palacio', 760, 1265),
    ('josue rugama', 760, 1345),
    ('abdiel ramos', 760, 1425),
    ('alfredo baker', 760, 1505),
]

sheet3_stamps = [
    ('rafael anguizola', 800, 230),
    ('asalon morales', 800, 280),
    ('maria de jaramillo', 650, 730),
    ('evangelista pineda', 660, 800),
    ('roger zurdo', 580, 860),
]

sheet4_stamps = [
    ('eymiss arcia', 620, 220),
    ('pedro solis', 620, 260),
    ('eloy smith', 620, 300),
    ('shanira crosdale', 620, 345),
    ('tomas marin', 620, 390),
    ('carlos martinez', 620, 505),
    ('maria del c. gonzalez', 830, 560),
    ('isidra peralta', 620, 640),
    ('argelia ibarra', 620, 810),
    ('vilma avila', 620, 865),
    ('rosa gonzalez', 620, 920),
    ('katherine saez', 620, 975),
    ('ovidio sanchez', 620, 1040),
    ('yesica edwards', 620, 1100),
    ('doris gracia', 620, 1155),
    ('benedicto abrego', 620, 1310),
]

stamped_p1 = stamp_sheet(LISTA_DIR / "listado 4.jpeg", sheet1_stamps)
stamped_p2 = stamp_sheet(LISTA_DIR / "listado2.jpeg", sheet2_stamps)
stamped_p3 = stamp_sheet(LISTA_DIR / "WhatsApp Image 2026-09-08 at 10.44.08 PM.jpeg", sheet3_stamps)
stamped_p4 = stamp_sheet(LISTA_DIR / "listado 3.jpeg", sheet4_stamps)

# Save multi-page PDF
stamped_pdf_pedregal = PEDREGAL_ROOT / "LISTADO DE ASISTENCIA PEDREGAL.pdf"
stamped_pdf_archive = LISTA_DIR / "LISTADO DE ASISTENCIA PEDREGAL CON CEDULAS.pdf"
stamped_pdf_adjuntos = ATTACHMENTS_DIR / "LISTADO DE ASISTENCIA PEDREGAL.pdf"

stamped_p1.save(
    stamped_pdf_pedregal,
    "PDF",
    resolution=150.0,
    save_all=True,
    append_images=[stamped_p2, stamped_p3, stamped_p4]
)
shutil.copy2(stamped_pdf_pedregal, stamped_pdf_archive)
shutil.copy2(stamped_pdf_pedregal, stamped_pdf_adjuntos)
print(f"[4] Multi-page Stamped Attendance PDF created successfully at:")
print(f"    - {stamped_pdf_pedregal}")
print(f"    - {stamped_pdf_archive}")
print(f"    - {stamped_pdf_adjuntos}")

# 5. Build rvpedregal folder & rvpedregal.zip
rv_dir = PEDREGAL_ROOT / "rvpedregal"
rv_dir.mkdir(parents=True, exist_ok=True)

# Copy PDFs to rvpedregal, to attachments, etc.
for p in patients:
    src_file = FIRMADAS_DIR / p["original_file"]
    # Save both with original and Signed name for maximum compatibility
    dst_signed = rv_dir / p["signed_pdf_name"]
    shutil.copy2(src_file, dst_signed)
    
    # In attachments, save as Signed.pdf
    dst_att = ATTACHMENTS_DIR / p["signed_pdf_name"]
    shutil.copy2(src_file, dst_att)

# Copy calibration cert
calib_dst_pedregal = PEDREGAL_ROOT / "Carta-de-Certificacion de calibracion.pdf"
calib_dst_att = ATTACHMENTS_DIR / "Carta-de-Certificacion de calibracion.pdf"
shutil.copy2(CALIB_PDF, calib_dst_pedregal)
shutil.copy2(CALIB_PDF, calib_dst_att)

# Create rvpedregal.zip
zip_path = PEDREGAL_ROOT / "rvpedregal.zip"
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for f in rv_dir.glob("*.pdf"):
        zf.write(f, arcname=f.name)

print(f"[5] rvpedregal folder and rvpedregal.zip created with {len(list(rv_dir.glob('*.pdf')))} files.")

# 6. Build Draft Report JSON
adjuntos_list = []
for p in patients:
    adjuntos_list.append({
        "name": p["signed_pdf_name"],
        "tipo": "Reporte Espirometría"
    })
adjuntos_list.append({
    "name": "Carta-de-Certificacion de calibracion.pdf",
    "tipo": "Certificado Calibración"
})
adjuntos_list.append({
    "name": "LISTADO DE ASISTENCIA PEDREGAL.pdf",
    "tipo": "Listado Asistencia"
})

resultados_espiro = []
for p in patients:
    resultados_espiro.append({
        "name": p["clean_name"],
        "cedula": p["cedula"],
        "age": p["age"],
        "position": p["position"],
        "result": p["result"]
    })

report_data = {
    "company_name": "Toledano Planta Pedregal",
    "report_type": "espirometria",
    "location": "Pedregal",
    "evaluation_date": "2026-08-11",
    "study_date": "2026-08-11",
    "plant": "Planta Pedregal",
    "company_activity": "Procesamiento y distribución de productos avícolas",
    "company_address": "Pedregal, Ciudad de Panamá",
    "counterpart_name": "",
    "counterpart_role": "",
    "evaluator_main": "",
    "evaluator_audio": "",
    "evaluator_spiro": "",
    "conclusion_text": "Se evaluaron un total de 60 colaboradores en la Planta Pedregal mediante pruebas de espirometría ocupacional. De ellos, 42 colaboradores (70%) presentaron parámetros espirométricos normales. 18 colaboradores (30%) presentaron hallazgos de alteración ventilatoria (12 restricción leve, 1 obstrucción leve, 3 restricción moderada y 2 restricción moderadamente severa).",
    "recommendations_text": "1. Mantener el uso estricto de equipo de protección respiratoria en áreas operativas y de procesamiento.\n2. Seguimiento médico periódico y control clínico para los colaboradores con restricción moderada y moderadamente severa.\n3. Realizar control espirométrico anual preventivo.",
    "resultados_audiometria": [],
    "resultados_espirometria": resultados_espiro,
    "adjuntos": adjuntos_list,
    "_draft_name": "toledano pedregal",
    "_version": "2.3.2",
    "_normalized": True,
    "date_mode": "multiple",
    "evaluation_dates": "11, 12, 19 y 21 de agosto de 2026",
    "study_date_mode": "multiple",
    "study_dates_multi": "11, 12, 19 y 21 de agosto de 2026",
    "study_dates": "2026-08-11",
    "country": "Panamá, República de Panamá"
}

# Save report draft with multiple names for easy access
with open(REPORTS_DIR / "toledano pedregal.json", "w", encoding="utf-8") as f:
    json.dump(report_data, f, ensure_ascii=False, indent=4)

report_data_alias = dict(report_data)
report_data_alias["_draft_name"] = "pedregal"
with open(REPORTS_DIR / "pedregal.json", "w", encoding="utf-8") as f:
    json.dump(report_data_alias, f, ensure_ascii=False, indent=4)

with open(DATABASES_DIR / "pedregal_extracted.json", "w", encoding="utf-8") as f:
    json.dump(patients, f, ensure_ascii=False, indent=2)

print(f"[6] Draft reports saved to:")
print(f"    - {REPORTS_DIR / 'toledano pedregal.json'}")
print(f"    - {REPORTS_DIR / 'pedregal.json'}")
print(f"    - {DATABASES_DIR / 'pedregal_extracted.json'}")

# 7. Update persons.json database
persons_file = DATABASES_DIR / "persons.json"
persons = {}
if persons_file.exists():
    try:
        with open(persons_file, "r", encoding="utf-8") as f:
            persons = json.load(f)
    except Exception:
        persons = {}

for p in patients:
    ced = p["cedula"]
    persons[ced] = {
        "identification": ced,
        "name": p["clean_name"],
        "age": p["age"],
        "position": p["position"],
        "last_result_label": p["result"],
        "last_test_type": "espirometria"
    }

with open(persons_file, "w", encoding="utf-8") as f:
    json.dump(persons, f, ensure_ascii=False, indent=2)

print(f"[7] Updated {persons_file} with {len(patients)} patients. Total in DB: {len(persons)}.")

# 8. Create extract_pedregal.py in workspace root
extract_script_content = f'''import json
from pathlib import Path

# Script de extracción y verificación de Pedregal
with open(r"data/reports/pedregal_extracted.json", "r", encoding="utf-8") as f:
    extracted = json.load(f)

print(f"Total colaboradores extraídos de Pedregal: {{len(extracted)}}")
normal_count = sum(1 for x in extracted if "normal" in x["result"].lower())
altered_count = len(extracted) - normal_count
print(f"Espirometrías normales: {{normal_count}}")
print(f"Espirometrías con hallazgos: {{altered_count}}")

print("\\nResumen por áreas:")
areas = {{}}
for x in extracted:
    a = x["position"] or "Sin Área"
    areas[a] = areas.get(a, 0) + 1

for a, count in areas.items():
    print(f"  - {{a}}: {{count}}")
'''

with open(WORKSPACE_ROOT / "extract_pedregal.py", "w", encoding="utf-8") as f:
    f.write(extract_script_content)

print("[8] Created extract_pedregal.py.")
print("\n>>> ALL PEDREGAL PROCESSING COMPLETED SUCCESSFULLY! <<<")
