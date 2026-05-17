import os
import sys
# pyrefly: ignore [missing-import]
import uvicorn
import webview
from threading import Thread
from pathlib import Path
from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
import json
import tempfile
import shutil
from datetime import datetime
import ctypes
from typing import Optional, List, Dict

_console_allocated = False

def set_console_visibility(visible: bool):
    """Controla la visibilidad de la consola de Windows, creándola si es necesario."""
    global _console_allocated
    try:
        kernel32 = ctypes.WinDLL('kernel32')
        user32 = ctypes.WinDLL('user32')
        
        if visible:
            if not _console_allocated:
                # Crear consola si no existe (para modo windowed)
                kernel32.AllocConsole()
                sys.stdout = open('CONOUT$', 'w', encoding='utf-8')
                sys.stderr = open('CONOUT$', 'w', encoding='utf-8')
                _console_allocated = True
                print("--- CAIT INFORMES: MODO DEPUREACIÓN ACTIVADO ---")
            
            hWnd = kernel32.GetConsoleWindow()
            if hWnd:
                user32.ShowWindow(hWnd, 5) # SW_SHOW
        else:
            hWnd = kernel32.GetConsoleWindow()
            if hWnd:
                user32.ShowWindow(hWnd, 0) # SW_HIDE
    except Exception as e:
        print(f"Error toggle console: {e}")

def _resolve_runtime_data_root() -> Path:
    if getattr(sys, "frozen", False):
        local_appdata = Path(os.getenv("LOCALAPPDATA") or Path.home())
        return local_appdata / "CAIT Informes" / "data"
    # En desarrollo: usar la carpeta de datos de la app original
    practica = Path(__file__).resolve().parent.parent / "practica profecional" / "data"
    if practica.exists():
        return practica
    return Path(__file__).resolve().parent / "data"

data_root = _resolve_runtime_data_root()
os.environ["CAIT_DATA_ROOT"] = str(data_root)

# Crear carpetas si no existen
(data_root / "reports").mkdir(parents=True, exist_ok=True)
(data_root / "exports").mkdir(parents=True, exist_ok=True)
(data_root / "attachments").mkdir(parents=True, exist_ok=True)
(data_root / "databases").mkdir(parents=True, exist_ok=True)

REPORT_DATA_FILE = data_root / "reports" / "current_report.json"

print(f"--- INICIALIZANDO CAIT ---")
print(f"Data Root: {data_root}")
print(f"--------------------------")

from src.core.result_schemes import RESULT_SCHEMES

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Servir archivos estaticos del frontend HTML
base_dir = Path(__file__).resolve().parent

# Montar las carpetas del frontend
app.mount("/presentacion", StaticFiles(directory=str(base_dir / "frontend" / "presentacion"), html=True), name="presentacion")
app.mount("/resultados", StaticFiles(directory=str(base_dir / "frontend" / "resultados"), html=True), name="resultados")
app.mount("/conclusiones", StaticFiles(directory=str(base_dir / "frontend" / "conclusiones"), html=True), name="conclusiones")
app.mount("/certificados", StaticFiles(directory=str(base_dir / "frontend" / "certificados"), html=True), name="certificados")
app.mount("/config", StaticFiles(directory=str(base_dir / "frontend" / "config"), html=True), name="config")
app.mount("/static", StaticFiles(directory=str(base_dir / "static")), name="static")
app.mount("/data", StaticFiles(directory=str(data_root)), name="data")

from fastapi.responses import FileResponse

@app.get("/api/download-zip/{filename}")
async def download_zip(filename: str):
    file_path = data_root / "reports" / filename
    if file_path.exists():
        return FileResponse(path=file_path, filename=filename, media_type='application/zip')
    return {"status": "error", "message": "File not found"}

@app.get("/")
def read_root():
    return RedirectResponse(url="/presentacion/")

# ================= API ENDPOINTS =================
# Aqui irian los endpoints que conectan con src.services.*

@app.get("/api/health")
def health_check():
    return {"status": "ok"}

# Inicializacion de servicios
from src.services.evaluators_repository import EvaluatorRepository
from src.services.counterparts_repository import CounterpartRepository
from src.services.persons_repository import PersonsRepository
from pydantic import BaseModel

evaluators_repo = EvaluatorRepository()
counterparts_repo = CounterpartRepository()
persons_repo = PersonsRepository()

@app.get("/api/result-schemes")
def get_result_schemes():
    return RESULT_SCHEMES

@app.get("/api/evaluators")
def get_evaluators():
    try:
        return evaluators_repo.list_all()
    except Exception as e:
        print(f"Error en /api/evaluators: {e}")
        return []

@app.get("/api/counterparts")
def get_counterparts():
    try:
        return counterparts_repo.list_all()
    except Exception as e:
        print(f"Error en /api/counterparts: {e}")
        return []

@app.get("/api/persons")
def get_persons():
    try:
        return persons_repo.list_all()
    except Exception as e:
        print(f"Error en /api/persons: {e}")
        return []

@app.post("/api/persons")
async def save_person(request: Request):
    data = await request.json()
    persons_repo.upsert(data)
    return {"status": "ok"}

REPORT_DATA_FILE = data_root / "reports" / "current_report.json"


@app.post("/api/evaluators")
async def add_evaluator(
    name: str = Form(...),
    profession: str = Form(""),
    registry: str = Form(""),
    applicable_reports: str = Form("[]"),
    file: Optional[UploadFile] = File(None)
):
    try:
        applies = json.loads(applicable_reports)
    except:
        applies = []
        
    evaluator_data = {
        "name": name,
        "profession": profession,
        "registry": registry,
        "applicable_reports": applies
    }
    
    if file:
        cert_dir = data_root / "attachments" / "idoneidad"
        cert_dir.mkdir(parents=True, exist_ok=True)
        # Limpiar nombre de archivo para evitar problemas de ruta
        safe_name = "".join([c for c in name if c.isalnum() or c in (" ", "_", "-")]).strip().replace(" ", "_")
        ext = Path(file.filename).suffix
        file_path = cert_dir / f"{safe_name}{ext}"
        
        with open(file_path, "wb") as f:
            f.write(await file.read())
            
        # Guardar ruta relativa para que sea portable
        evaluator_data["credential_file"] = f"data/attachments/idoneidad/{file_path.name}"
    
    evaluators_repo.add_evaluator(evaluator_data)
    return {"status": "ok"}

@app.delete("/api/evaluators/{ev_id}")
def delete_evaluator(ev_id: str):
    evaluators_repo.remove_evaluator(ev_id)
    return {"status": "deleted"}

@app.post("/api/counterparts")
async def add_counterpart(request: Request):
    data = await request.json()
    counterparts_repo.add_counterpart(data)
    return {"status": "ok"}

@app.delete("/api/counterparts/{cp_id}")
def delete_counterpart(cp_id: str):
    counterparts_repo.remove_counterpart(cp_id)
    return {"status": "deleted"}

from src.services.zip_exporter import ZipExporter
zip_exporter = ZipExporter()

@app.post("/api/upload-attachment")
async def upload_attachment(file: UploadFile = File(...)):
    """Sube un archivo adjunto a la carpeta global de anexos."""
    try:
        att_dir = data_root / "attachments" / "report_adjuntos"
        att_dir.mkdir(parents=True, exist_ok=True)
        
        file_path = att_dir / file.filename
        with open(file_path, "wb") as f:
            f.write(await file.read())
            
        return {"status": "ok", "filename": file.filename, "path": f"data/attachments/report_adjuntos/{file.filename}"}
    except Exception as e:
        print(f"Error en upload_attachment: {e}")
        return {"status": "error", "message": str(e)}

from src.services.pdf_generator import PDFGenerator
pdf_gen = PDFGenerator()

from starlette.concurrency import run_in_threadpool

@app.post("/api/export-zip")
async def export_zip(request: Request):
    try:
        data = await request.json()
    except:
        data = {}
    target_folder = data.get("target_folder")
    
    if not REPORT_DATA_FILE.exists():
        return {"status": "error", "message": "No hay datos para exportar. Guarde el reporte primero."}
        
    with open(REPORT_DATA_FILE, "r", encoding="utf-8") as f:
        report = json.load(f)
    
    company = str(report.get('company_name', 'CAIT')).strip()
    location = str(report.get('location') or report.get('plant') or 'General').strip()
    
    # Extraer año de evaluation_date (formato esperado YYYY-MM-DD)
    eval_date = str(report.get('evaluation_date', ''))
    try:
        if '-' in eval_date:
            year = eval_date.split('-')[0]
        else:
            year = datetime.now().year
    except:
        year = datetime.now().year
        
    report_name = f"{company} - {year} - {location}"
    
    # Sanitizar nombre para evitar errores de sistema de archivos
    for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        report_name = report_name.replace(char, '_')
    
    # Limpiar espacios múltiples y finales
    report_name = " ".join(report_name.split()).strip()
    
    # 1. Preparar directorio de exportación (donde se creará la carpeta descomprimida)
    # Si no hay carpeta elegida, usar la carpeta de exports del sistema
    export_base = Path(target_folder) if target_folder and os.path.exists(target_folder) else (data_root / "exports")
    export_base.mkdir(parents=True, exist_ok=True)
    
    dest_dir = export_base / report_name
    if dest_dir.exists(): shutil.rmtree(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    
    # 2. Enriquecer datos del reporte (resolver IDs a nombres/perfiles)
    try:
        # Resolver evaluadores
        eval_ids = {
            "main": report.get("evaluator_main"),
            "audio": report.get("evaluator_audio"),
            "spiro": report.get("evaluator_spiro")
        }
        
        # El principal es el que sale en el encabezado
        if eval_ids["main"]:
            ev = evaluators_repo.get_by_id(eval_ids["main"])
            if ev:
                report["evaluator_name"] = ev.get("name")
                report["evaluator_profile"] = ev
        elif eval_ids["audio"]:
            ev = evaluators_repo.get_by_id(eval_ids["audio"])
            if ev:
                report["evaluator_name"] = ev.get("name")
                report["evaluator_profile"] = ev
        elif eval_ids["spiro"]:
            ev = evaluators_repo.get_by_id(eval_ids["spiro"])
            if ev:
                report["evaluator_name"] = ev.get("name")
                report["evaluator_profile"] = ev
        
        # Resolver Contraparte Técnica
        cp_id = report.get("counterpart_id") or report.get("counterpart")
        if cp_id:
            cp = counterparts_repo.get_by_id(cp_id)
            if cp:
                report["counterpart_profile"] = cp
                report["counterpart_name"] = cp.get("name")
        
        # Equipo técnico (si es combinado, incluir ambos)
        tech_team = []
        rpt_type_str = str(report.get("report_type") or report.get("type") or "").lower()
        is_combined = "audiometr" in rpt_type_str and "espiro" in rpt_type_str
        
        for role in ["audio", "spiro"]:
            eid = eval_ids[role]
            if eid:
                ev = evaluators_repo.get_by_id(eid)
                if ev:
                    member = {
                        "name": ev.get("name"),
                        "details": ev.get("technical_details", [])
                    }
                    # Si es combinado, queremos ambos bloques aunque sea la misma persona
                    # para que la portada respete el diseño de dos columnas.
                    if is_combined:
                        tech_team.append(member)
                    elif not any(m["name"] == member["name"] for m in tech_team):
                        tech_team.append(member)
        
        if tech_team:
            report["technical_team"] = tech_team
            
    except Exception as e:
        print(f"Error enriqueciendo datos para PDF: {e}")

    # 3. Generar PDF real directamente en la carpeta de destino
    pdf_name = f"Informe_{report_name}.pdf"
    final_pdf_path = dest_dir / pdf_name
    logo_path = str(base_dir / "static" / "logo.png")
    
    try:
        await run_in_threadpool(pdf_gen.generate, report, str(final_pdf_path), logo_path=logo_path)
    except Exception as e:
        print(f"Error generando PDF: {e}")
        if not final_pdf_path.exists():
            with open(final_pdf_path, "w") as f: f.write("%PDF-1.4 DUMMY")
    
    # 3. Recolectar y copiar adjuntos categorizados
    adjuntos_root = dest_dir / "Adjuntos"
    adjuntos_root.mkdir(parents=True, exist_ok=True)
    
    # Categorías mapeadas (compatibilidad con keys antiguas)
    attachments_map = {
        "Calibracion": report.get("calibration_files", []),
        "Audiometrias": (report.get("result_attachments", {}) or {}).get("audiometria", []),
        "Espirometrias": (report.get("result_attachments", {}) or {}).get("espirometria", []),
        "Asistencia": report.get("attendance_files", []) or [],
    }
    
    # Integrar desde la lista unificada 'adjuntos' del nuevo frontend
    adjuntos_list = report.get("adjuntos", [])
    if isinstance(adjuntos_list, list):
        for a in adjuntos_list:
            if not isinstance(a, dict): continue
            tipo = a.get("tipo")
            folder = None
            if tipo == "Certificado Calibración": folder = "Calibracion"
            elif tipo == "Audiograma" or tipo == "Protocolo Resultados": folder = "Audiometrias"
            elif tipo == "Reporte Espirometría": folder = "Espirometrias"
            elif tipo == "Listado Asistencia": folder = "Asistencia"
            
            if folder:
                path = a.get("path")
                if not path:
                    name = a.get("name")
                    if name: path = f"data/attachments/report_adjuntos/{name}"
                if path and path not in attachments_map.setdefault(folder, []):
                    attachments_map[folder].append(path)

    # Añadir Idoneidades
    evaluators_path = data_root / "databases" / "evaluators.json"
    if evaluators_path.exists():
        with open(evaluators_path, "r", encoding="utf-8") as f:
            evaluators = json.load(f)
        eval_ids = set([report.get("evaluator_main"), report.get("evaluator_audio"), report.get("evaluator_spiro")])
        idoneidades = []
        for eid in eval_ids:
            if not eid: continue
            ev = next((e for e in evaluators if e.get("id") == eid), None)
            if ev and ev.get("credential_file"):
                # Credential file suele estar relativo a la carpeta raíz de datos
                cpath = data_root.parent / ev["credential_file"]
                if cpath.exists(): idoneidades.append(str(cpath))
        attachments_map["Idoneidades"] = idoneidades

    # Copiar archivos físicamente a las subcarpetas
    def resolve_p(p):
        if not p: return None
        if os.path.exists(p): return p
        # Resolver data/ -> data_root
        if str(p).startswith("data/"):
            full = data_root / p[5:]
            if full.exists(): return str(full)
        if str(p).startswith("data\\"):
            full = data_root / p[5:]
            if full.exists(): return str(full)
        return None

    for cat_name, files in attachments_map.items():
        if not files: continue
        cat_dir = adjuntos_root / cat_name
        cat_dir.mkdir(parents=True, exist_ok=True)
        for f in files:
            path = f if isinstance(f, str) else (f.get("file") if isinstance(f, dict) else None)
            resolved = resolve_p(path)
            if resolved:
                shutil.copy2(resolved, cat_dir)

    # 4. Crear el ZIP a partir de la carpeta que ya llenamos
    zip_base_name = str(export_base / report_name)
    zip_path = await run_in_threadpool(
        shutil.make_archive,
        zip_base_name,
        'zip',
        root_dir=str(export_base),
        base_dir=report_name
    )
    
    filename = f"{report_name}.zip"
    return {
        "status": "ok", 
        "filename": filename, 
        "zip_path": zip_path,
        "folder_path": str(dest_dir),
        "message": "Se han creado tanto la carpeta como el archivo ZIP en el destino."
    }

@app.get("/api/drafts")
def list_drafts():
    drafts_dir = data_root / "reports"
    files = list(drafts_dir.glob("*.json"))
    return [
        {
            "name": f.name,
            "modified": os.path.getmtime(f),
            "size": os.path.getsize(f)
        }
        for f in files if f.name != "current_report.json"
    ]

@app.post("/api/report")
async def save_report(request: Request):
    data = await request.json()
    data["_version"] = "2.2.8"
    name = data.get("_draft_name", "current_report.json")
    if not name.endswith(".json"): name += ".json"
    
    target = data_root / "reports" / name
    target.parent.mkdir(parents=True, exist_ok=True)
    
    # Si es el guardado automático o el actual, también actualizar current_report.json
    with open(target, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        
    if name != "current_report.json":
        # También guardar como actual para que se mantenga al recargar
        with open(data_root / "reports" / "current_report.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
    return {"status": "saved", "filename": name}

def normalize_report(data: dict) -> dict:
    """Normaliza los campos de un reporte para asegurar compatibilidad con versiones anteriores."""
    if not data: return {}
    
    # 0. Desenvolver si viene envuelto en un objeto "report"
    if "report" in data and isinstance(data["report"], dict) and len(data) == 1:
        data = data["report"]

    data["_normalized"] = True
    data["_version"] = "2.2.2-DEBUG"
    
    # 1. Normalizar tipo de informe (Clave para otros procesos)
    rt = str(data.get("report_type") or data.get("type") or "").lower()
    if "audiometr" in rt and "espirom" in rt:
        data["report_type"] = "audiometria_espirometria"
    elif "audiometr" in rt:
        data["report_type"] = "audiometria"
    elif "espirom" in rt:
        data["report_type"] = "espirometria"

    # 2. Mapeo de campos de cabecera antiguos a nuevos
    if "company" in data and not data.get("company_name"):
        data["company_name"] = data["company"]
    if "company_counterpart" in data and not data.get("counterpart_name"):
        data["counterpart_name"] = data["company_counterpart"]
    if "plant" in data and not data.get("location"):
        data["location"] = data["plant"]
    if "activity" in data and not data.get("company_activity"):
        data["company_activity"] = data["activity"]
    if "counterpart" in data and not data.get("counterpart_name"):
        data["counterpart_name"] = data["counterpart"]
    
    # Normalizar fechas (intentar convertir DD/MM/YYYY a YYYY-MM-DD)
    def normalize_date(d):
        if not d or not isinstance(d, str): return d
        if "/" in d:
            parts = d.split("/")
            if len(parts) == 3:
                # Asumir DD/MM/YYYY
                day, month, year = parts
                if len(year) == 4: return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                if len(day) == 4: return f"{day}-{month.zfill(2)}-{year.zfill(2)}"
        return d

    if "date" in data and not data.get("evaluation_date"):
        data["evaluation_date"] = normalize_date(data["date"])
    if data.get("evaluation_date"):
        data["evaluation_date"] = normalize_date(data["evaluation_date"])
    if "study_date" in data:
        data["study_date"] = normalize_date(data["study_date"])
        
    # Evaluadores
    for k in ["evaluador", "evaluator", "evaluator_name"]:
        if k in data and not data.get("evaluator_main"):
            data["evaluator_main"] = data[k]

    # 3. Normalizar conclusiones y recomendaciones
    if "conclusion" in data and not data.get("conclusion_text"):
        data["conclusion_text"] = data["conclusion"]
    if "recommendations" in data and not data.get("recommendations_text"):
        data["recommendations_text"] = data["recommendations"]

    # 4. Migrar listas de resultados (evaluated, results, evaluaciones, evaluated_entries)
    ee = data.get("evaluated_entries")
    if isinstance(ee, dict):
        if "audiometria" in ee and not data.get("resultados_audiometria"):
            data["resultados_audiometria"] = ee["audiometria"]
        if "espirometria" in ee and not data.get("resultados_espirometria"):
            data["resultados_espirometria"] = ee["espirometria"]

    legacy_list = data.get("evaluated") or data.get("results") or data.get("evaluaciones")
    if legacy_list and isinstance(legacy_list, list):
        target_key = "resultados_audiometria"
        if data.get("report_type") == "espirometria":
            target_key = "resultados_espirometria"
        
        if not data.get(target_key):
            data[target_key] = legacy_list

    # Normalizar campos internos de los resultados (español -> inglés)
    for key in ["resultados_audiometria", "resultados_espirometria"]:
        if key in data and isinstance(data[key], list):
            for item in data[key]:
                if not isinstance(item, dict): continue
                # Español -> Inglés
                if "nombre" in item and "name" not in item: item["name"] = item["nombre"]
                if "identificacion" in item and "cedula" not in item: item["cedula"] = item["identificacion"]
                if "puesto" in item and "position" not in item: item["position"] = item["puesto"]
                if "edad" in item and "age" not in item: item["age"] = item["edad"]
                if "resultado" in item and "result" not in item: item["result"] = item["resultado"]
                
                # Antiguo Inglés -> Nuevo Inglés
                if "identification" in item and "cedula" not in item: item["cedula"] = item["identification"]
                if "result_label" in item and "result" not in item: item["result"] = item["result_label"]

    # 5. Normalizar adjuntos (migrar de listas planas a lista de objetos)
    if "adjuntos" not in data:
        data["adjuntos"] = []
    
    # Mapear archivos de calibración antiguos
    for k in ["calibration_files", "calibraciones"]:
        if k in data and isinstance(data[k], list):
            for f in data[k]:
                if not isinstance(f, str): continue
                name = os.path.basename(f)
                if not any(a.get("name") == name for a in data["adjuntos"]):
                    data["adjuntos"].append({"name": name, "tipo": "Certificado Calibración"})
                
    # Mapear archivos de asistencia antiguos
    for k in ["attendance_files", "asistencia"]:
        if k in data and isinstance(data[k], list):
            for f in data[k]:
                if not isinstance(f, str): continue
                name = os.path.basename(f)
                if not any(a.get("name") == name for a in data["adjuntos"]):
                    data["adjuntos"].append({"name": name, "tipo": "Listado Asistencia"})

    # Mapear adjuntos de resultados antiguos (result_attachments o listas directas)
    ra = data.get("result_attachments") or data.get("test_attachment_files") or {}
    if isinstance(ra, dict):
        for subkey, type_label in [("audiometria", "Audiograma"), ("espirometria", "Reporte Espirometría")]:
            if subkey in ra and isinstance(ra[subkey], list):
                for f in ra[subkey]:
                    if not isinstance(f, str): continue
                    name = os.path.basename(f)
                    if not any(a.get("name") == name for a in data["adjuntos"]):
                        data["adjuntos"].append({"name": name, "tipo": type_label})
                
    return data

@app.get("/api/report")
def get_report(name: str = "current_report.json"):
    target = data_root / "reports" / name
    if target.exists():
        try:
            with open(target, "r", encoding="utf-8") as f:
                data = json.load(f)
                return normalize_report(data)
        except Exception as e:
            print(f"Error cargando reporte {name}: {e}")
            return {}
    return {}

@app.delete("/api/drafts/{name}")
def delete_draft(name: str):
    target = data_root / "reports" / name
    if target.exists() and name != "current_report.json":
        target.unlink()
    return {"status": "ok"}

# ================= TEMPLATES API =================

@app.get("/api/templates/default")
def get_default_templates(report_type: str = "audiometria"):
    """Retorna las plantillas predefinidas del sistema original."""
    # Normalizar el tipo que viene del frontend
    nt = report_type.lower()
    is_audio = "audiometria" in nt
    is_spiro = "espirometria" in nt
    
    # Obtener datos del reporte actual para rellenar variables básicas
    report = {}
    if REPORT_DATA_FILE.exists():
        try:
            with open(REPORT_DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                report = normalize_report(data)
        except Exception as e:
            print(f"Error cargando current_report para plantilla: {e}")
            
    company = report.get("company_name", "[Nombre de Empresa]")
    location = report.get("location", "[Ubicación]")
    activity = report.get("company_activity", "[Actividad]")
    dates = report.get("evaluation_dates") or report.get("evaluation_date", "[Fechas]")
    
    # Reconstrucción de plantillas de conclusión (basado en src/ui/app.py)
    conclusion = ""
    if is_audio and is_spiro:
        conclusion = (
            f"La empresa {company}, coordinó jornadas de audiometrías y espirometrías durante {dates}, "
            f"para los colaboradores en el área {location}.\n\n"
            "Se aplicaron ambas pruebas al personal convocado, recopilando de forma confidencial la historia clínica y laboral de cada participante. "
            "Todos recibieron orientación sobre la protección auditiva y respiratoria que deben aplicar.\n\n"
            "Los resultados individuales se incluyen dentro del informe, junto con resúmenes diagnósticos y gráficos porcentuales.\n\n"
            "Los equipos utilizados cuentan con certificados de calibración anual, lo que respalda la confiabilidad de las mediciones."
        )
    elif is_spiro:
        conclusion = (
            f"La empresa {company}, realizó la toma de las pruebas de espirometrías durante {dates}, "
            f"a los colaboradores expuestos en el área de {location}.\n\n"
            "Se aplicó la prueba a los colaboradores, quienes suministraron su historia clínica de salud y laboral confidencialmente. "
            "Fueron orientados sobre sus resultados y la protección laboral que deben aplicar.\n\n"
            "El resultado de las pruebas se encuentra documentado en este informe, junto con un resumen diagnóstico.\n\n"
            "El Espirómetro utilizado posee su certificado de calibración anual, lo que garantiza la confiabilidad de los resultados."
        )
    else:
        conclusion = (
            f"La empresa {company}, realizó las audiometrías durante {dates}, "
            f"a los colaboradores en el área de {location}.\n\n"
            "Se aplicó la prueba de audiometría laboral, cada colaborador suministró su historia clínica y laboral de forma confidencial. "
            "Recibió orientación sobre los cuidados de su audición y la protección laboral que debe aplicar.\n\n"
            "El resultado de cada evaluación se encuentra documentado dentro de este informe.\n\n"
            "El audiómetro utilizado corresponde al equipo Otopod, con certificado de calibración anual ISO."
        )

    # Recomendaciones (basado en src/ui/app.py)
    recs = []
    if is_audio:
        recs += [
            "• Se recomienda la realización periódica (anual) de la evaluación auditiva (audiometría laboral) para mantener el registro y control.",
            "• Continuar suministrándoles los equipos de seguridad a los colaboradores para disminuir el ruido laboral.",
            "• Seguimiento a casos de vigilancia ocupacional y recomendación de visita a Otorrinolaringología si hay molestias persistentes."
        ]
    if is_spiro:
        recs += [
            "• Se recomienda la realización periódica (anual) de las pruebas de espirometría laboral.",
            "• Los colaboradores no deben exponerse a contaminantes (humo de cigarrillo, químicos) fuera de horas laborales.",
            "• Utilización de protección respiratoria adecuada en áreas de exposición a partículas.",
            "• Preservar e incentivar el uso correcto de implementos de seguridad durante toda la jornada."
        ]
    
    return {
        "conclusion": conclusion,
        "recommendations": "\n\n".join(recs)
    }

TEMPLATE_FILE = data_root / "databases" / "conclusion_templates.json"
RECS_TEMPLATE_FILE = data_root / "databases" / "recommendation_templates.json"

@app.get("/api/conclusion-templates")
def get_user_templates():
    if TEMPLATE_FILE.exists():
        try:
            with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content: return []
                data = json.loads(content)
                if isinstance(data, dict):
                    # Migración: de viejo {name: text} a nuevo [{name, text}]
                    return [{"name": k, "text": v} for k, v in data.items()]
                return data
        except Exception as e:
            print(f"Error cargando plantillas de usuario: {e}")
            return []
    return []

@app.post("/api/conclusion-templates")
async def save_user_template(request: Request):
    data = await request.json()
    TEMPLATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    templates = get_user_templates()
    templates.append(data)
    with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=4)
    return {"status": "ok"}

@app.delete("/api/conclusion-templates/{index}")
def delete_user_template(index: int):
    templates = get_user_templates()
    if 0 <= index < len(templates):
        templates.pop(index)
        with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
            json.dump(templates, f, ensure_ascii=False, indent=4)
    return {"status": "ok"}

@app.get("/api/recommendation-templates")
def get_recs_templates():
    if RECS_TEMPLATE_FILE.exists():
        try:
            with open(RECS_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content: return []
                data = json.loads(content)
                if isinstance(data, dict):
                    # Migración: de viejo {name: text} a nuevo [{name, text}]
                    return [{"name": k, "text": v} for k, v in data.items()]
                return data
        except Exception as e:
            print(f"Error cargando plantillas de recomendaciones: {e}")
            return []
    return []

@app.post("/api/recommendation-templates")
async def save_recs_template(request: Request):
    data = await request.json()
    templates = get_recs_templates()
    templates.append(data)
    with open(RECS_TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=4)
    return {"status": "ok"}

@app.delete("/api/recommendation-templates/{index}")
def delete_recs_template(index: int):
    templates = get_recs_templates()
    if 0 <= index < len(templates):
        templates.pop(index)
        with open(RECS_TEMPLATE_FILE, "w", encoding="utf-8") as f:
            json.dump(templates, f, ensure_ascii=False, indent=4)
    return {"status": "ok"}

@app.post("/api/debug/console")
async def toggle_console(request: Request):
    data = await request.json()
    if data.get("password") == "@utp.2898":
        set_console_visibility(True)
        return {"status": "ok", "message": "Consola habilitada"}
    return {"status": "error", "message": "Contraseña incorrecta"}

# Endpoint para generar PDF, etc.
# ...

def run_server():
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="error")

class Api:
    def __init__(self):
        self._window = None

    def set_window(self, window):
        self._window = window

    def select_folder(self):
        if not self._window: return None
        result = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        return result[0] if result else None

api = Api()

def wait_for_server(port, timeout=10):
    """Espera a que el puerto local esté listo antes de abrir la UI."""
    import socket
    import time
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return True
        except (socket.timeout, ConnectionRefusedError):
            time.sleep(0.2)
    return False

if __name__ == "__main__":
    # Iniciar servidor en hilo aparte
    t = Thread(target=run_server)
    t.daemon = True
    t.start()
    
    # Esperar a que el servidor esté listo (máximo 10 segundos)
    if not wait_for_server(8000):
        print("Error: El servidor no inició a tiempo.")
        sys.exit(1)
    
    # Iniciar ventana principal
    icon_path = str(Path(__file__).parent / 'logo-apli-removebg-preview.ico')
    window = webview.create_window(
        'CAIT Panamá - Generador de Informes v2.2.8', 
        'http://127.0.0.1:8000', 
        width=1360, 
        height=900, 
        text_select=True, 
        background_color='#fcf9f8',
        js_api=api
    )
    api.set_window(window)
    webview.start(icon=icon_path)
