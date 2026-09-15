import os
import json
import shutil
from pathlib import Path
from PIL import Image

# Rutas del entorno
WORKSPACE_ROOT = Path(r"c:\Users\fioni\Documents\audioetria yara\Cait panama nuevo")
SRC_DIR = Path(r"c:\Users\fioni\Downloads\resultados\pedregal pero audiometria")

ATTACHMENTS_DIR = WORKSPACE_ROOT / "data" / "attachments" / "report_adjuntos"
REPORTS_DIR = WORKSPACE_ROOT / "data" / "reports"
DATABASES_DIR = WORKSPACE_ROOT / "data" / "databases"

ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
DATABASES_DIR.mkdir(parents=True, exist_ok=True)

# 1. Lista de 60 pacientes de Pedregal Audiometría
patients_pedregal = [
    {"name": "Luis Gonzalez", "cedula": "9-709-406", "age": "", "position": "Supervisor Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Gemino Acosta", "cedula": "9-127-447", "age": "61", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Ramiro Abrego", "cedula": "1-742-1439", "age": "47", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Ignacio Vasquez", "cedula": "9-169-694", "age": "57", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Manuel Perez", "cedula": "8-822-467", "age": "38", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Agripino Gonzalez", "cedula": "4-265-307", "age": "", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Anastacio Rodriguez", "cedula": "9-114-2060", "age": "60", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Manuel Degracia", "cedula": "4-718-1245", "age": "46", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Secundino Vanega", "cedula": "3-711-1778", "age": "43", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Isael Gonzalez", "cedula": "2-750-992", "age": "24", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Armando Morales", "cedula": "2-709-871", "age": "45", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Ceferino Mendoza", "cedula": "4-765-706", "age": "46", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Ariel Gil", "cedula": "8-937-841", "age": "27", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Javier Abrego", "cedula": "1-734-351", "age": "32", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Elvin Cortez", "cedula": "4-805-1669", "age": "30", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Jose Ines Garcia", "cedula": "9-205-635", "age": "52", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Jesus Bernal", "cedula": "8-722-1538", "age": "47", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Alberto Carpintero", "cedula": "4-803-677", "age": "43", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Leonaldo Jaen", "cedula": "7-711-1269", "age": "28", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Indalecio Vasquez", "cedula": "9-711-138", "age": "46", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Jose Ines Castillo", "cedula": "8-280-525", "age": "", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Roman Eugenio", "cedula": "1-43-939", "age": "54", "position": "Pollo Vivo", "result": "Normal bilateral"},
    {"name": "Eymiss Arcia", "cedula": "8-718-2101", "age": "48", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Pedro Solis", "cedula": "4-195-916", "age": "61", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Eloy Smith", "cedula": "12-704-1497", "age": "25", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Shanira Crosdale", "cedula": "8-709-2446", "age": "49", "position": "Supervisora Evisceración", "result": "Normal bilateral"},
    {"name": "Tomas Marin", "cedula": "9-201-776", "age": "53", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Carlos Martinez", "cedula": "8-813-1647", "age": "38", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Evangelista Pineda", "cedula": "4-203-194", "age": "57", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Isidra Peralta", "cedula": "9-702-1488", "age": "49", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Yesenia Aparicio", "cedula": "4-273-197", "age": "", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Argelia Ibarra", "cedula": "2-712-585", "age": "43", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Vilma Avila", "cedula": "8-280-850", "age": "59", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Rosa Gonzalez", "cedula": "2-162-780", "age": "50", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Katherine Saez", "cedula": "8-763-1989", "age": "43", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Ovidio Sanchez", "cedula": "2-701-1801", "age": "49", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Yesica Edwards", "cedula": "8-516-2110", "age": "58", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Doris Gracia", "cedula": "8-522-692", "age": "53", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Nilda Valdez", "cedula": "4-737-2171", "age": "", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Benedicto Abrego", "cedula": "1-746-676", "age": "27", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Melania Caraballo", "cedula": "9-149-570", "age": "", "position": "Evisceración", "result": "Normal bilateral"},
    {"name": "Luis Palma", "cedula": "8-372-482", "age": "61", "position": "Desplume", "result": "Normal bilateral"},
    {"name": "Aristides Alveo", "cedula": "2-126-597", "age": "56", "position": "Desplume", "result": "Normal bilateral"},
    {"name": "Martin Ortega", "cedula": "4-712-1342", "age": "", "position": "Desplume", "result": "Normal bilateral"},
    {"name": "Marvin Mendoza", "cedula": "4-796-595", "age": "45", "position": "Desplume", "result": "Normal bilateral"},
    {"name": "Justino Castillo", "cedula": "1-703-2053", "age": "", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Alexis Pitti", "cedula": "8-467-446", "age": "52", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Francisco Acosta", "cedula": "4-256-768", "age": "54", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Renato Palacio", "cedula": "1-705-1277", "age": "46", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Josue Rugama", "cedula": "1-740-623", "age": "30", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Asalon Morales", "cedula": "1-714-1456", "age": "41", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Abdiel Ramos", "cedula": "8-962-1133", "age": "25", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Alfredo Baker", "cedula": "1-737-1462", "age": "33", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Rafael Anguizola", "cedula": "8-1014-565", "age": "22", "position": "Lavado de Canasta", "result": "Normal bilateral"},
    {"name": "Rolando Mora", "cedula": "", "age": "", "position": "Operador Deshuesado", "result": "Normal bilateral"},
    {"name": "Ernesto Gonzalez", "cedula": "8-934-2172", "age": "", "position": "Deshuesado", "result": "Normal bilateral"},
    {"name": "Elmira Guevara", "cedula": "9-732-2314", "age": "", "position": "Deshuesado", "result": "Normal bilateral"},
    {"name": "Eliecer Pinzon", "cedula": "8-874-1041", "age": "", "position": "Deshuesado", "result": "Normal bilateral"},
    {"name": "Ezequiel Rueda", "cedula": "8-863-825", "age": "", "position": "Deshuesado", "result": "Normal bilateral"},
    {"name": "Domingo Cruz", "cedula": "8-704-1688", "age": "", "position": "Deshuesado", "result": "Normal bilateral"},
]

print(f"[Pedregal] 60 pacientes preparados.")

# 2. Generar PDF multipágina del Listado de Asistencia
img_files = [
    SRC_DIR / "WhatsApp Image 2026-09-13 at 12.07.49 PM.jpeg",
    SRC_DIR / "WhatsApp Image 2026-09-13 at 12.07.58 PM.jpeg",
    SRC_DIR / "WhatsApp Image 2026-09-13 at 12.08.12 PM.jpeg"
]

pil_imgs = [Image.open(f).convert("RGB") for f in img_files if f.exists()]
if pil_imgs:
    pdf_dest = ATTACHMENTS_DIR / "LISTADO DE ASISTENCIA PEDREGAL AUDIOMETRIA.pdf"
    pil_imgs[0].save(pdf_dest, "PDF", resolution=150.0, save_all=True, append_images=pil_imgs[1:])
    # Copiar también en la carpeta de origen
    shutil.copy2(pdf_dest, SRC_DIR / "LISTADO DE ASISTENCIA PEDREGAL AUDIOMETRIA.pdf")
    print(f"[Pedregal] PDF de asistencia generado: {pdf_dest}")

# 3. Crear borrador JSON de Pedregal Audiometría
report_audio = {
    "company_name": "Toledano Planta Pedregal",
    "report_type": "audiometria",
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
    "conclusion_text": "Se evaluaron un total de 60 colaboradores en la Planta Pedregal mediante pruebas de audiometría ocupacional.",
    "recommendations_text": "1. Uso obligatorio de protección auditiva adecuada en áreas de exposición a ruido continuo o de impacto.\n2. Seguimiento y control audiométrico periódico anual.\n3. Capacitación continua en conservación de la audición.",
    "resultados_audiometria": patients_pedregal,
    "resultados_espirometria": [],
    "adjuntos": [
        {
            "name": "LISTADO DE ASISTENCIA PEDREGAL AUDIOMETRIA.pdf",
            "tipo": "Listado Asistencia"
        }
    ],
    "_draft_name": "pedregal audiometria",
    "_version": "2.3.5",
    "_normalized": True,
    "date_mode": "multiple",
    "evaluation_dates": "11, 12, 19 y 21 de agosto de 2026",
    "study_date_mode": "multiple",
    "study_dates_multi": "11, 12, 19 y 21 de agosto de 2026",
    "study_dates": "2026-08-11",
    "country": "Panamá, República de Panamá"
}

with open(REPORTS_DIR / "pedregal audiometria.json", "w", encoding="utf-8") as f:
    json.dump(report_audio, f, ensure_ascii=False, indent=4)

report_audio_alias = dict(report_audio)
report_audio_alias["_draft_name"] = "toledano pedregal audiometria"
with open(REPORTS_DIR / "toledano pedregal audiometria.json", "w", encoding="utf-8") as f:
    json.dump(report_audio_alias, f, ensure_ascii=False, indent=4)

print(f"[Pedregal] Borradores guardados: pedregal audiometria.json y toledano pedregal audiometria.json")

# 4. Los borradores de audiometría se mantienen independientes (pedregal audiometria.json)
# para preservar la espirometría original intacta en pedregal.json.

# 5. Actualizar catálogo persons.json
persons_file = DATABASES_DIR / "persons.json"
persons = {}
if persons_file.exists():
    try:
        with open(persons_file, "r", encoding="utf-8") as f:
            persons = json.load(f)
    except Exception:
        persons = {}

for p in patients_pedregal:
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
    json.dump(persons, f, ensure_ascii=False, indent=2)

print(f"[Pedregal] persons.json actualizado. Total personas: {len(persons)}.")
