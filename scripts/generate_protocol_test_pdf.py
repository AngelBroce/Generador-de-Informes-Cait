from __future__ import annotations

from datetime import datetime
from pathlib import Path
import sys

from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas as rl_canvas

project_root = Path(__file__).resolve().parents[1]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.services.pdf_generator import PDFGenerator  # noqa: E402


def _create_demo_attachment(path: Path, title: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    page_size = landscape(letter)
    canvas = rl_canvas.Canvas(str(path), pagesize=page_size)
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawString(72, page_size[1] - 72, title)
    canvas.setFont("Helvetica", 12)
    canvas.drawString(72, page_size[1] - 100, "Demo attachment for protocol ordering test.")
    canvas.drawString(72, page_size[1] - 120, f"Generated: {datetime.now().strftime('%d/%m/%Y')}")
    canvas.save()
    return str(path)


def _build_entries(total: int = 18) -> list[dict]:
    entries = []
    codes = [
        ("espiro_normal", "Espirometria normal"),
        ("restriccion_leve", "Restriccion leve"),
        ("obstruccion_moderada", "Obstruccion moderada"),
    ]
    for idx in range(1, total + 1):
        code, label = codes[idx % len(codes)]
        entries.append(
            {
                "name": f"Colaborador {idx:02d}",
                "identification": f"ID-{idx:03d}",
                "age": str(20 + (idx % 15)),
                "position": "Operario",
                "result_label": label,
                "result_code": code,
                "test_type": "espirometria",
            }
        )
    return entries


def main() -> None:
    exports_dir = project_root / "data" / "exports"
    exports_dir.mkdir(parents=True, exist_ok=True)

    attachments_dir = project_root / "data" / "attachments" / "demo_protocolo"
    calibration_files = [
        _create_demo_attachment(attachments_dir / "calibracion_equipo.pdf", "Certificado de calibracion"),
    ]
    spirometry_files = [
        _create_demo_attachment(attachments_dir / "reporte_espirometria_a.pdf", "Reporte de espirometria - Turno A"),
        _create_demo_attachment(attachments_dir / "reporte_espirometria_b.pdf", "Reporte de espirometria - Turno B"),
    ]
    attendance_files = [
        _create_demo_attachment(attachments_dir / "listado_asistencia.pdf", "Listado de asistencia"),
    ]

    logo_path = project_root / "src" / "assets" / "logo_cait.png"
    logo = str(logo_path) if logo_path.exists() else None

    report = {
        "id": f"REP_PROTOCOLO_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "type": "Espirometria",
        "company": "Empresa Demo",
        "location": "Ciudad de Panama",
        "evaluator": "Licda. Yara Lizeth Perez A.",
        "date": datetime.now().strftime("%d/%m/%Y"),
        "plant": "Planta Demo",
        "activity": "Produccion",
        "company_counterpart": "Ing. Rivera",
        "counterpart_role": "Supervisor",
        "country": "Panama",
        "study_dates": "01/02/2026 - 04/02/2026",
        "evaluated": _build_entries(),
        "conclusion": "Conclusiones de demostracion para el protocolo.",
        "recommendations": "• Ejemplo de recomendacion 1.\n• Ejemplo de recomendacion 2.",
        "calibration_files": calibration_files,
        "spirometry_files": spirometry_files,
        "attendance_files": attendance_files,
        "attachments": calibration_files + spirometry_files + attendance_files,
    }

    output_path = exports_dir / "informe_prueba_protocolo_espirometria.pdf"
    ok = PDFGenerator().generate(report, str(output_path), logo)
    if not ok:
        raise RuntimeError("No se pudo generar el PDF de prueba")

    print(f"PDF de prueba generado en: {output_path}")


if __name__ == "__main__":
    main()
