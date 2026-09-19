import os
import json
import shutil
from pathlib import Path
from PIL import Image

# Rutas del entorno
WORKSPACE_ROOT = Path(r"c:\Users\fioni\Documents\audioetria yara\Cait panama nuevo")
SRC_DIR = Path(r"c:\Users\fioni\Downloads\resultados\rio congo")
PDF_SRC_DIR = Path(r"c:\Users\fioni\Downloads\resultados\pedregal pero audiometria\12-707-523 JIMENEZ, NATANAEL 14-09-2026 08-19 a. m. Signed")

ATTACHMENTS_DIR = WORKSPACE_ROOT / "data" / "attachments" / "report_adjuntos"
REPORTS_DIR = WORKSPACE_ROOT / "data" / "reports"
DATABASES_DIR = WORKSPACE_ROOT / "data" / "databases"

ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
DATABASES_DIR.mkdir(parents=True, exist_ok=True)

# 1. Lista VERIFICADA al 100% de los 40 colaboradores de Río Congo Audiometría
# Cruzada entre las Pruebas Audiométricas Firmadas Oficiales y la Lista de Asistencia (Área)
patients_rio_congo = [
    {"row": 1, "name": "Robert Pinzón", "cedula": "8-877-983", "age": "32", "dob": "13/08/1994", "position": "Mantenimiento", "result": "Normal bilateral", "emp_num": "110432"},
    {"row": 2, "name": "Gregorio Márquez", "cedula": "8-854-1915", "age": "34", "dob": "05/07/1992", "position": "Mantenimiento", "result": "Normal bilateral", "emp_num": ""},
    {"row": 3, "name": "Moisés García", "cedula": "4-806-2427", "age": "26", "dob": "04/04/2000", "position": "Despacho", "result": "Normal bilateral", "emp_num": "109442"},
    {"row": 4, "name": "Teofila Santo", "cedula": "4-784-2210", "age": "39", "dob": "02/02/1987", "position": "Limpieza", "result": "Normal bilateral", "emp_num": "111299"},
    {"row": 5, "name": "Mariángel Reyes", "cedula": "8-1005-559", "age": "23", "dob": "06/09/2003", "position": "Notificación", "result": "Normal bilateral", "emp_num": "110971"},
    {"row": 6, "name": "Albis Martínez", "cedula": "8-385-738", "age": "60", "dob": "28/03/1966", "position": "Lavandería", "result": "Normal bilateral", "emp_num": "6040"},
    {"row": 7, "name": "Jennifer Flores", "cedula": "8-831-2203", "age": "36", "dob": "05/12/1990", "position": "Línea", "result": "Normal bilateral", "emp_num": "111517"},
    {"row": 8, "name": "Eladio Pérez", "cedula": "9-200-569", "age": "53", "dob": "26/05/1973", "position": "Termosellado", "result": "Caída leve unilateral", "emp_num": "109161"},
    {"row": 9, "name": "Elizabeth Bonilla", "cedula": "6-59-705", "age": "60", "dob": "31/03/1966", "position": "Termosellado", "result": "Caída leve bilateral", "emp_num": "109629"},
    {"row": 10, "name": "Silvia De Batista", "cedula": "8-524-2150", "age": "61", "dob": "09/01/1965", "position": "Línea", "result": "Normal bilateral", "emp_num": "10"},
    {"row": 11, "name": "Marianela Sánchez", "cedula": "8-949-213", "age": "26", "dob": "16/06/2000", "position": "Limpieza interna", "result": "Normal bilateral", "emp_num": "111665"},
    {"row": 12, "name": "Yanis Cedeño", "cedula": "8-909-696", "age": "29", "dob": "29/08/1997", "position": "Supervisora", "result": "Normal bilateral", "emp_num": "110791"},
    {"row": 13, "name": "Víctor Man", "cedula": "8-747-2255", "age": "45", "dob": "18/01/1981", "position": "Termosellado", "result": "Normal bilateral", "emp_num": "110620"},
    {"row": 14, "name": "Odalis Martínez", "cedula": "3-747-1101", "age": "25", "dob": "21/04/2001", "position": "Denester", "result": "Normal bilateral", "emp_num": "111099"},
    {"row": 15, "name": "Francisco Murillo", "cedula": "8-1035-2102", "age": "21", "dob": "04/02/2005", "position": "Denester", "result": "Normal bilateral", "emp_num": "113133"},
    {"row": 16, "name": "Julia Fula", "cedula": "8-779-20", "age": "41", "dob": "11/11/1985", "position": "Líneas", "result": "Normal bilateral", "emp_num": "109592"},
    {"row": 17, "name": "Isaí Castillo", "cedula": "8-1009-1703", "age": "22", "dob": "31/08/2004", "position": "Termosellado", "result": "Normal bilateral", "emp_num": "111317"},
    {"row": 18, "name": "Samuel Mendoza", "cedula": "8-905-1038", "age": "30", "dob": "11/01/1996", "position": "Hacen cajita", "result": "Normal bilateral", "emp_num": "110496"},
    {"row": 19, "name": "Emily Duarte", "cedula": "4-786-1641", "age": "29", "dob": "12/10/1997", "position": "Denester", "result": "Normal bilateral", "emp_num": "111820"},
    {"row": 20, "name": "Yanacell Carrera", "cedula": "8-829-2466", "age": "36", "dob": "02/07/1990", "position": "Producción", "result": "Normal bilateral", "emp_num": "109145"},
    {"row": 21, "name": "Bolívar García", "cedula": "9-742-126", "age": "32", "dob": "02/03/1994", "position": "Denester", "result": "Normal bilateral", "emp_num": "110740"},
    {"row": 22, "name": "Maycol Chirú", "cedula": "8-1032-581", "age": "20", "dob": "13/07/2006", "position": "Multicargador", "result": "Normal bilateral", "emp_num": "111589"},
    {"row": 23, "name": "Gladys Rodríguez", "cedula": "8-999-192", "age": "28", "dob": "19/11/1998", "position": "Línea", "result": "Normal bilateral", "emp_num": "111848"},
    {"row": 24, "name": "Mayra Fernández", "cedula": "8-890-1233", "age": "31", "dob": "11/06/1995", "position": "Línea", "result": "Normal bilateral", "emp_num": "111928"},
    {"row": 25, "name": "David Rodríguez", "cedula": "8-299-6", "age": "58", "dob": "05/05/1968", "position": "Recibo", "result": "Normal bilateral", "emp_num": "108879"},
    {"row": 26, "name": "Luis Martínez", "cedula": "8-786-1831", "age": "31", "dob": "09/10/1995", "position": "Recibo", "result": "Normal bilateral", "emp_num": "109825"},
    {"row": 27, "name": "Félix Ruiz", "cedula": "8-1017-792", "age": "21", "dob": "17/09/2005", "position": "Despacho", "result": "Normal bilateral", "emp_num": "111548"},
    {"row": 28, "name": "Anthony Quintero", "cedula": "8-984-474", "age": "23", "dob": "29/11/2002", "position": "Termosellado", "result": "Normal bilateral", "emp_num": "111343"},
    {"row": 29, "name": "Natanael Jiménez", "cedula": "12-707-523", "age": "24", "dob": "10/03/2002", "position": "Estibador", "result": "Normal bilateral", "emp_num": "111001"},
    {"row": 30, "name": "Benjamín Miranda", "cedula": "12-706-620", "age": "28", "dob": "12/12/1998", "position": "Proimat", "result": "Normal bilateral", "emp_num": ""},
    {"row": 31, "name": "Jean Carlos Salinas", "cedula": "8-1032-369", "age": "22", "dob": "18/07/2004", "position": "Línea 9410", "result": "Normal bilateral", "emp_num": "111090"},
    {"row": 32, "name": "Vladimir Bocanegra", "cedula": "8-772-1863", "age": "42", "dob": "21/12/1984", "position": "Termosellado", "result": "Normal bilateral", "emp_num": "113223"},
    {"row": 33, "name": "Edier Quintero", "cedula": "8-1026-836", "age": "21", "dob": "23/05/2005", "position": "Estibador", "result": "Normal bilateral", "emp_num": "113105"},
    {"row": 34, "name": "José Arena", "cedula": "4-800-805", "age": "36", "dob": "13/01/1990", "position": "Operador", "result": "Normal bilateral", "emp_num": "109408"},
    {"row": 35, "name": "Nehemías Gutiérrez", "cedula": "8-886-2361", "age": "32", "dob": "01/10/1994", "position": "Lavado", "result": "Normal bilateral", "emp_num": ""},
    {"row": 36, "name": "José Mariscal", "cedula": "8-993-2405", "age": "24", "dob": "10/01/2002", "position": "Termosellado", "result": "Normal bilateral", "emp_num": "111575"},
    {"row": 37, "name": "Teófilo Castillo", "cedula": "8-1006-1228", "age": "22", "dob": "23/02/2004", "position": "Notificación", "result": "Normal bilateral", "emp_num": "113132"},
    {"row": 38, "name": "Joydeth González", "cedula": "8-893-1992", "age": "31", "dob": "27/04/1995", "position": "Calidad", "result": "Normal bilateral", "emp_num": "112030"},
    {"row": 39, "name": "Iván Valdés", "cedula": "8-984-1911", "age": "24", "dob": "19/08/2002", "position": "Supervisor", "result": "Normal bilateral", "emp_num": "111910"},
    {"row": 40, "name": "Nixon Fogona", "cedula": "11-700-735", "age": "30", "dob": "11/01/1996", "position": "SISO", "result": "Normal bilateral", "emp_num": ""},
]

print(f"[Río Congo] {len(patients_rio_congo)} colaboradores verificados y preparados.")

# 2. Generar PDF multipágina del Listado de Asistencia si aún no existe en attachments
img_files = [
    SRC_DIR / "WhatsApp Image 2026-09-13 at 12.20.15 PM.jpeg",
    SRC_DIR / "WhatsApp Image 2026-09-13 at 12.20.23 PM.jpeg",
    SRC_DIR / "WhatsApp Image 2026-09-13 at 12.20.36 PM.jpeg"
]

pil_imgs = [Image.open(f).convert("RGB") for f in img_files if f.exists()]
if pil_imgs:
    pdf_dest = ATTACHMENTS_DIR / "LISTADO DE ASISTENCIA RIO CONGO AUDIOMETRIA.pdf"
    pil_imgs[0].save(pdf_dest, "PDF", resolution=150.0, save_all=True, append_images=pil_imgs[1:])
    shutil.copy2(pdf_dest, SRC_DIR / "LISTADO DE ASISTENCIA RIO CONGO AUDIOMETRIA.pdf")
    print(f"[Río Congo] PDF de asistencia generado/actualizado: {pdf_dest}")

# 3. Formato para el reporte JSON de CAIT
report_patients = [
    {
        "name": p["name"],
        "cedula": p["cedula"],
        "age": p["age"],
        "position": p["position"],
        "result": p["result"]
    }
    for p in patients_rio_congo
]

report_rio_congo = {
    "company_name": "Toledano Río Congo",
    "report_type": "audiometria",
    "location": "Río Congo",
    "evaluation_date": "2026-08-13",
    "study_date": "2026-08-13",
    "plant": "Planta Río Congo",
    "company_activity": "Procesamiento y distribución de productos avícolas",
    "company_address": "Río Congo, Panamá",
    "counterpart_name": "",
    "counterpart_role": "",
    "evaluator_main": "Licda. Yara Lizeth Pérez A.",
    "evaluator_audio": "Licda. Yara Lizeth Pérez A.",
    "evaluator_spiro": "",
    "conclusion_text": (
        "La empresa Productos Toledano S.A., realizó la toma de audiometrías ocupacionales el mes de agosto "
        "(13 de agosto de 2026) a los colaboradores en la Planta Río Congo.\n\n"
        "Se aplicó la prueba de audiometría laboral a un total de 40 colaboradores convocados, recopilando de forma "
        "confidencial su historia clínica y laboral. Del total evaluado: 38 colaboradores (95.0%) presentaron audición "
        "dentro de los límites normales (Normal bilateral), y 2 colaboradores (5.0%) presentaron alteración auditiva "
        "(1 caso con caída leve unilateral en oído derecho y 1 caso con caída leve bilateral).\n\n"
        "Todos los colaboradores fueron debidamente orientados sobre el cuidado de su audición, técnicas de colocación "
        "de protección auditiva y prevención de fatiga acústica.\n\n"
        "El audiómetro utilizado corresponde al equipo Otopod con certificado de calibración anual ISO, asegurando "
        "la validez de las mediciones."
    ),
    "recommendations_text": (
        "• Mantener el control audiométrico periódico (anual) para todos los colaboradores expuestos a niveles sonoros continuos o de impacto.\n"
        "• Suministrar y exigir el uso continuo y adecuado del equipo de protección auditiva normado (tapones / orejeras) en planta.\n"
        "• Dar seguimiento médico a los 2 colaboradores que presentaron caída leve (Eladio Pérez con caída unilateral y Elizabeth Bonilla con caída bilateral).\n"
        "• Reforzar el programa de vigilancia epidemiológica de conservación auditiva en las instalaciones de Río Congo."
    ),
    "resultados_audiometria": report_patients,
    "resultados_espirometria": [],
    "adjuntos": [
        {
            "name": "LISTADO DE ASISTENCIA RIO CONGO AUDIOMETRIA.pdf",
            "tipo": "Listado Asistencia"
        }
    ],
    "_draft_name": "rio congo",
    "_version": "2.3.5",
    "_normalized": True,
    "date_mode": "single",
    "evaluation_dates": "13 de agosto de 2026",
    "study_date_mode": "single",
    "study_dates_multi": "13 de agosto de 2026",
    "study_dates": "2026-08-13",
    "country": "Panamá, República de Panamá"
}

with open(REPORTS_DIR / "rio congo.json", "w", encoding="utf-8") as f:
    json.dump(report_rio_congo, f, ensure_ascii=False, indent=4)

report_rio_congo_alias = dict(report_rio_congo)
report_rio_congo_alias["_draft_name"] = "toledano rio congo"
with open(REPORTS_DIR / "toledano rio congo.json", "w", encoding="utf-8") as f:
    json.dump(report_rio_congo_alias, f, ensure_ascii=False, indent=4)

print(f"[Río Congo] Borradores guardados: rio congo.json y toledano rio congo.json")

# 4. Actualizar catálogo general persons.json
persons_file = DATABASES_DIR / "persons.json"
persons = {}
if persons_file.exists():
    try:
        with open(persons_file, "r", encoding="utf-8") as f:
            persons = json.load(f)
    except Exception:
        persons = {}

for p in report_patients:
    ced = p["cedula"]
    if not ced:
        continue
    existing_entry = persons.get(ced, {})
    persons[ced] = {
        "identification": ced,
        "name": p["name"],
        "age": p["age"] or existing_entry.get("age", ""),
        "position": p["position"] or existing_entry.get("position", ""),
        "last_result_label": p["result"],
        "last_test_type": "audiometria"
    }

with open(persons_file, "w", encoding="utf-8") as f:
    json.dump(persons, f, ensure_ascii=False, indent=4)

print(f"[Río Congo] Base de datos de personas actualizada con los 40 colaboradores verificados.")
