from pathlib import Path
from datetime import datetime
import sys

from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import landscape, letter

project_root = Path(__file__).resolve().parents[1]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.services.pdf_generator import PDFGenerator
from src.services.evaluators_repository import EvaluatorRepository, build_technical_details

logo = project_root / "src" / "assets" / "logo_cait.png"
repo = EvaluatorRepository()
report_type = "espirometría"
report_code = "espirometria"
primary_evaluator = repo.get_primary_for_report(report_code) or {}
technical_team = [
    {
        "name": member.get("name", ""),
        "details": build_technical_details(member),
    }
    for member in repo.get_team_for_report(report_type)
]
if not technical_team and primary_evaluator:
    technical_team.append({
        "name": primary_evaluator.get("name", ""),
        "details": build_technical_details(primary_evaluator),
    })

entries = []
for idx in range(1, 70):
    mod = idx % 4
    if mod == 0:
        code = "espiro_normal"
        label = "Normal"
    elif mod == 1:
        code = "restriccion_leve"
        label = "Restricción leve"
    elif mod == 2:
        code = "obstruccion_moderada"
        label = "Obstrucción moderada"
    else:
        code = "restriccion_grave"
        label = "Restricción grave"

    entries.append(
        {
            "name": f"Colaborador {idx:02d}",
            "identification": f"ID-{idx:03d}",
            "age": str(24 + (idx % 18)),
            "position": "Operario",
            "result_label": label,
            "result_code": code,
            "test_type": "espirometria",
        }
    )

def _create_demo_certificate(path: Path, title: str, certificate_number: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    page_size = landscape(letter)
    canvas = rl_canvas.Canvas(str(path), pagesize=page_size)
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawString(72, page_size[1] - 72, title)
    canvas.setFont("Helvetica", 12)
    canvas.drawString(72, page_size[1] - 100, f"Certificado: {certificate_number}")
    canvas.drawString(72, page_size[1] - 120, f"Emitido: {datetime.now().strftime('%d/%m/%Y')}")
    canvas.drawString(72, page_size[1] - 140, "Documento de demostración para anexar al informe.")
    canvas.save()
    return str(path)


certificates_dir = project_root / "data" / "attachments" / "demo_certificados"
calibration_files = [
    _create_demo_certificate(certificates_dir / "calibracion_espirometro.pdf", "Certificado Espirómetro EasyOne", "CAL-2026-001"),
    _create_demo_certificate(certificates_dir / "calibracion_nebulizador.pdf", "Certificado Nebulizador", "CAL-2026-017"),
]


def _create_demo_result_attachment(path: Path, title: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    page_size = landscape(letter)
    canvas = rl_canvas.Canvas(str(path), pagesize=page_size)
    canvas.setFont("Helvetica-Bold", 18)
    canvas.drawString(72, page_size[1] - 72, title)
    canvas.setFont("Helvetica", 12)
    canvas.drawString(72, page_size[1] - 110, "Documento de referencia para anexos del informe.")
    canvas.drawString(72, page_size[1] - 130, f"Generado: {datetime.now().strftime('%d/%m/%Y')}")
    canvas.save()
    return str(path)


results_dir = project_root / "data" / "attachments" / "demo_resultados"
spirometry_files = [
    _create_demo_result_attachment(results_dir / "reporte_espirometria.pdf", "Reporte gráfico de espirometría"),
]

report = {
    "id": "REP_DEMO_EXTENSO",
    "type": report_type,
    "company": "WDA",
    "location": "DWD",
    "evaluator": primary_evaluator.get("name", "Licda. Stephanie María Thorne."),
    "date": datetime.now().strftime("%d/%m/%Y"),
    "plant": "PLANTA DEMO",
    "activity": "Producción",
    "company_counterpart": "Ing. Rivera",
    "counterpart_role": "Supervisor",
    "country": "Panamá",
    "study_dates": "01/02/2026 - 04/02/2026",
    "evaluated": entries,
    "calibration_certificates": calibration_files,
    "calibration_files": calibration_files,
    "spirometry_files": spirometry_files,
    "attachments": calibration_files + spirometry_files,
    "evaluator_profile": primary_evaluator,
    "technical_team": technical_team,
    "conclusion": "Se realizaron pruebas espirométricas a 69 colaboradores expuestos a partículas y los resultados se resumen en las tablas adjuntas.",
    "recommendations": "• Mantener el control anual de espirometrías.\n• Garantizar el uso de protección respiratoria en las áreas críticas.",
    "logo_path": str(logo),
}

output = project_root / "data" / "exports" / "informe_demo_extenso.pdf"
PDFGenerator().generate(report, str(output), str(logo))
print(f"PDF extenso generado en {output}")
