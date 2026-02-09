from __future__ import annotations

import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas as rl_canvas

project_root = Path(__file__).resolve().parents[1]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.services.pdf_generator import PDFGenerator  # noqa: E402


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


def _build_demo_entries(total: int = 24) -> list[dict]:
    entries = []
    for idx in range(1, total + 1):
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
                "age": str(22 + (idx % 17)),
                "position": "Operario",
                "result_label": label,
                "result_code": code,
                "test_type": "espirometria",
            }
        )
    return entries


def _sanitize_filename(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*]', "_", value)
    sanitized = re.sub(r"\s+", " ", sanitized).strip()
    return sanitized or "Informe"


def _extract_year(*values: str) -> str:
    for value in values:
        if not value:
            continue
        match = re.search(r"(19|20)\d{2}", value)
        if match:
            return match.group(0)
    return datetime.now().strftime("%Y")


def _copy_report_attachments(package_dir: Path, report: dict) -> int:
    attachments_root = package_dir / "Adjuntos"
    groups = [
        ("Calibracion", report.get("calibration_files") or []),
        ("Audiometrias", report.get("audiogram_files") or []),
        ("Espirometrias", report.get("spirometry_files") or []),
        ("Asistencia", report.get("attendance_files") or []),
    ]

    copied = 0
    seen_paths: set[str] = set()

    def ensure_root():
        attachments_root.mkdir(parents=True, exist_ok=True)

    def normalize(path_obj: Path) -> str:
        try:
            return str(path_obj.resolve())
        except OSError:
            return str(path_obj)

    for folder_name, file_list in groups:
        valid_files: list[Path] = []
        for file_path in file_list:
            candidate = Path(file_path)
            if not candidate.exists():
                continue
            normalized = normalize(candidate)
            if normalized in seen_paths:
                continue
            seen_paths.add(normalized)
            valid_files.append(candidate)

        if not valid_files:
            continue

        ensure_root()
        target_folder = attachments_root / folder_name
        target_folder.mkdir(parents=True, exist_ok=True)
        for source in valid_files:
            _copy_file_with_unique_name(source, target_folder)
            copied += 1

    extra_files = []
    for file_path in report.get("attachments") or []:
        candidate = Path(file_path)
        if not candidate.exists():
            continue
        normalized = normalize(candidate)
        if normalized in seen_paths:
            continue
        seen_paths.add(normalized)
        extra_files.append(candidate)

    if extra_files:
        ensure_root()
        target_folder = attachments_root / "Otros"
        target_folder.mkdir(parents=True, exist_ok=True)
        for source in extra_files:
            _copy_file_with_unique_name(source, target_folder)
            copied += 1

    return copied


def _copy_file_with_unique_name(source: Path, target_dir: Path) -> None:
    destination = target_dir / source.name
    if not destination.exists():
        shutil.copy2(source, destination)
        return

    counter = 1
    while True:
        candidate = target_dir / f"{source.stem}_{counter}{source.suffix}"
        if not candidate.exists():
            shutil.copy2(source, candidate)
            break
        counter += 1


def main():
    exports_dir = project_root / "data" / "exports"
    exports_dir.mkdir(parents=True, exist_ok=True)

    attachments_dir = project_root / "data" / "attachments"
    certificates_dir = attachments_dir / "demo_zip_certificados"
    results_dir = attachments_dir / "demo_zip_resultados"

    calibration_files = [
        _create_demo_certificate(
            certificates_dir / "calibracion_espirometro.pdf",
            "Certificado Espirómetro",
            "CAL-2026-001",
        ),
        _create_demo_certificate(
            certificates_dir / "calibracion_nebulizador.pdf",
            "Certificado Nebulizador",
            "CAL-2026-014",
        ),
    ]

    spirometry_files = [
        _create_demo_result_attachment(
            results_dir / "reporte_espirometria_turno_a.pdf",
            "Reporte de espirometría - Turno A",
        ),
        _create_demo_result_attachment(
            results_dir / "reporte_espirometria_turno_b.pdf",
            "Reporte de espirometría - Turno B",
        ),
    ]

    audiogram_files = [
        _create_demo_result_attachment(
            results_dir / "audiograma_turno_a.pdf",
            "Audiograma - Turno A",
        )
    ]

    attendance_files = [
        _create_demo_result_attachment(
            results_dir / "listado_asistencia_general.pdf",
            "Listado de asistencia - Jornada",
        )
    ]

    company = "Industrias La Esperanza"
    report_type = "Audiometría y Espirometría"
    study_dates = "01/02/2026 - 04/02/2026"
    year = _extract_year(study_dates)
    base_name = _sanitize_filename(f"Informe {company} {year} {report_type}")

    package_dir = exports_dir / base_name
    if package_dir.exists():
        shutil.rmtree(package_dir)
    package_dir.mkdir(parents=True, exist_ok=True)

    logo_path = project_root / "src" / "assets" / "logo_cait.png"
    logo = str(logo_path) if logo_path.exists() else None

    report = {
        "id": f"REP_DEMO_ZIP_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "type": report_type,
        "company": company,
        "location": "Ciudad de Panamá",
        "evaluator": "Licda. Stephanie María Thorne",
        "date": datetime.now().strftime("%d/%m/%Y"),
        "plant": "Planta Central",
        "activity": "Manufactura",
        "company_counterpart": "Ing. Rivera",
        "counterpart_role": "Jefe de Seguridad",
        "country": "Panamá",
        "study_dates": study_dates,
        "evaluated": _build_demo_entries(),
        "conclusion": "Conclusiones de demostración para el paquete ZIP.",
        "recommendations": "• Ejemplo de recomendación 1.\n• Ejemplo de recomendación 2.",
        "calibration_certificates": [{"file": path} for path in calibration_files],
        "calibration_files": calibration_files,
        "audiogram_files": audiogram_files,
        "spirometry_files": spirometry_files,
        "attendance_files": attendance_files,
        "attachments": calibration_files + audiogram_files + spirometry_files + attendance_files,
        "evaluator_profile": {
            "name": "Licda. Stephanie María Thorne",
            "profession": "Terapeuta Respiratoria",
            "registry": "Reg. Idóneo 124",
            "technical_details": [
                "Terapeuta Respiratoria,",
                "Registro Idóneo 124.",
            ],
        },
    }

    pdf_path = package_dir / f"{base_name}.pdf"
    if not PDFGenerator().generate(report, str(pdf_path), logo):
        raise RuntimeError("No se pudo generar el PDF de demostración")

    attachments_copied = _copy_report_attachments(package_dir, report)

    zip_base = exports_dir / base_name
    zip_path = zip_base.with_suffix(".zip")
    if zip_path.exists():
        zip_path.unlink()

    shutil.make_archive(str(zip_base), "zip", root_dir=exports_dir, base_dir=base_name)

    print("Paquete ZIP de demostración creado:")
    print(f" - Carpeta: {package_dir}")
    print(f" - PDF: {pdf_path.name}")
    print(f" - Adjuntos copiados: {attachments_copied}")
    print(f" - ZIP final: {zip_path}")


if __name__ == "__main__":
    main()
