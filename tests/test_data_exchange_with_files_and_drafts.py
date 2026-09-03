import io
import json
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from fastapi.testclient import TestClient
from api_main import app, data_root, persons_repo, evaluators_repo, counterparts_repo
from src.services.data_exchange_service import DataExchangeService

client = TestClient(app)
exchange_service = DataExchangeService()

def run_test():
    print("=== INICIANDO PRUEBA: EXPORTACIÓN/IMPORTACIÓN CON PDFS Y BORRADORES ===")

    # 1. Crear un archivo adjunto PDF real en data/attachments/report_adjuntos
    att_dir = data_root / "attachments" / "report_adjuntos"
    att_dir.mkdir(parents=True, exist_ok=True)
    test_pdf_name = "test_adjunto_clinico.pdf"
    test_pdf_path = att_dir / test_pdf_name
    test_pdf_content = b"%PDF-1.4 TEST CLINICAL ATTACHMENT FOR CAIT DATA PORTABILITY"
    test_pdf_path.write_bytes(test_pdf_content)
    print(f"1. PDF de prueba creado en: {test_pdf_path}")

    # 2. Crear un borrador adicional en data/reports
    reports_dir = data_root / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    sample_draft_name = "Borrador_Test_Empresa_Beta.json"
    sample_draft_path = reports_dir / sample_draft_name
    sample_draft_data = {
        "company_name": "Empresa Beta S.A.",
        "report_type": "audiometria",
        "evaluation_date": "2026-09-01",
        "adjuntos": [{"name": test_pdf_name, "tipo": "Certificado Calibración"}],
        "resultados_audiometria": [
            {"name": "Carlos Gomez", "cedula": "8-999-888", "age": "30", "position": "Operador", "result": "Normal bilateral"}
        ]
    }
    sample_draft_path.write_text(json.dumps(sample_draft_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"2. Borrador de prueba guardado en: {sample_draft_path}")

    # 3. Guardar el reporte actual mediante POST /api/report
    current_report_data = {
        "company_name": "Empresa Alfa Panamá",
        "report_type": "audiometria_espirometria",
        "evaluation_date": "2026-09-02",
        "adjuntos": [{"name": test_pdf_name, "tipo": "Certificado Calibración"}],
        "resultados_audiometria": [
            {"name": "Ana Rodriguez", "cedula": "8-111-222", "age": "28", "position": "Administrativa", "result": "Normal bilateral"}
        ],
        "resultados_espirometria": [
            {"name": "Ana Rodriguez", "cedula": "8-111-222", "age": "28", "position": "Administrativa", "result": "Espirometría normal"}
        ]
    }
    res_save = client.post("/api/report", json=current_report_data)
    assert res_save.status_code == 200, f"Error al guardar reporte: {res_save.text}"
    print("3. Reporte actual guardado en el servidor.")

    # 4. Probar exportación de .cait (JSON con PDFs y borradores)
    res_cait = client.get("/api/export/cait")
    assert res_cait.status_code == 200, f"Error exportando .cait: {res_cait.text}"
    cait_json = res_cait.json()
    assert "embedded_files" in cait_json, "embedded_files no encontrado en paquete .cait"
    assert len(cait_json["embedded_files"]) > 0, "No se embebió ningún archivo en .cait"
    embedded_names = [f["name"] for f in cait_json["embedded_files"]]
    assert test_pdf_name in embedded_names, f"{test_pdf_name} no está en embedded_files"
    assert "saved_drafts" in cait_json, "saved_drafts no encontrado en paquete .cait"
    draft_names = [d["name"] for d in cait_json["saved_drafts"]]
    assert sample_draft_name in draft_names, f"{sample_draft_name} no está en saved_drafts"
    print(f"4. Exportación .cait exitosa con {len(cait_json['embedded_files'])} archivos y {len(cait_json['saved_drafts'])} borradores embebidos.")

    # 5. Probar exportación de .caitpkg (ZIP con PDFs físicos y borradores en carpeta drafts/)
    res_pkg = client.get("/api/export/caitpkg")
    assert res_pkg.status_code == 200, f"Error exportando .caitpkg: {res_pkg.text}"
    pkg_bytes = res_pkg.content
    with zipfile.ZipFile(io.BytesIO(pkg_bytes)) as zf:
        pkg_namelist = zf.namelist()
        assert "manifest.json" in pkg_namelist, "manifest.json no encontrado en .caitpkg"
        assert any(n.endswith(test_pdf_name) for n in pkg_namelist), f"{test_pdf_name} no incluido en .caitpkg"
        assert f"drafts/{sample_draft_name}" in pkg_namelist, f"drafts/{sample_draft_name} no incluido en .caitpkg"
    print(f"5. Exportación .caitpkg exitosa con adjunto físico y borrador en carpeta drafts/.")

    # 6. Probar exportación de .caitbackup (Copia de seguridad completa del sistema)
    res_backup = client.get("/api/system/backup")
    assert res_backup.status_code == 200, f"Error exportando .caitbackup: {res_backup.text}"
    backup_bytes = res_backup.content
    with zipfile.ZipFile(io.BytesIO(backup_bytes)) as zf:
        backup_namelist = zf.namelist()
        assert "backup_manifest.json" in backup_namelist, "backup_manifest.json no encontrado en .caitbackup"
        assert any(n.endswith(test_pdf_name) for n in backup_namelist), f"{test_pdf_name} no incluido en backup"
        assert any(n.endswith(sample_draft_name) for n in backup_namelist), f"{sample_draft_name} no incluido en backup"
    print("6. Exportación de Copia de Seguridad (.caitbackup) exitosa con todos los borradores y adjuntos.")

    # 7. Probar importación y restauración en un directorio temporal aislado
    temp_isolated_dir = Path(tempfile.mkdtemp(prefix="cait_restore_test_"))
    try:
        # A) Restaurar backup en directorio temporal
        restore_result = exchange_service.restore_system_backup(backup_bytes, temp_isolated_dir, mode="merge")
        assert restore_result["status"] == "ok", f"Error en restore_system_backup: {restore_result}"
        assert (temp_isolated_dir / "reports" / sample_draft_name).exists(), "El borrador no fue restaurado por restore_system_backup"
        assert (temp_isolated_dir / "attachments" / "report_adjuntos" / test_pdf_name).exists(), "El PDF adjunto no fue restaurado por restore_system_backup"
        restored_content = (temp_isolated_dir / "attachments" / "report_adjuntos" / test_pdf_name).read_bytes()
        assert restored_content == test_pdf_content, "El contenido del PDF restaurado no coincide"
        print(f"7A. Restauración de backup en entorno aislado exitosa: {restore_result['message']}")

        # B) Importar paquete .caitpkg en otro directorio temporal aislado
        temp_pkg_dir = Path(tempfile.mkdtemp(prefix="cait_pkg_import_test_"))
        try:
            from api_main import normalize_report
            import_pkg_res = exchange_service.import_report_package(
                file_bytes=pkg_bytes,
                filename="test.caitpkg",
                data_root=temp_pkg_dir,
                persons_repo=persons_repo,
                evaluators_repo=evaluators_repo,
                counterparts_repo=counterparts_repo,
                normalize_func=normalize_report
            )
            assert import_pkg_res["status"] == "ok", f"Error importando .caitpkg: {import_pkg_res}"
            assert (temp_pkg_dir / "attachments" / "report_adjuntos" / test_pdf_name).exists(), "El PDF no se restauró desde .caitpkg"
            assert (temp_pkg_dir / "reports" / sample_draft_name).exists(), "El borrador no se restauró desde .caitpkg"
            print(f"7B. Importación de paquete .caitpkg exitosa: {import_pkg_res['message']}")
        finally:
            shutil.rmtree(temp_pkg_dir, ignore_errors=True)

        # C) Importar archivo .cait (JSON) en otro directorio temporal aislado
        temp_cait_dir = Path(tempfile.mkdtemp(prefix="cait_json_import_test_"))
        try:
            cait_bytes = json.dumps(cait_json).encode("utf-8")
            import_cait_res = exchange_service.import_report_package(
                file_bytes=cait_bytes,
                filename="test.cait",
                data_root=temp_cait_dir,
                persons_repo=persons_repo,
                evaluators_repo=evaluators_repo,
                counterparts_repo=counterparts_repo,
                normalize_func=normalize_report
            )
            assert import_cait_res["status"] == "ok", f"Error importando .cait: {import_cait_res}"
            assert (temp_cait_dir / "attachments" / "report_adjuntos" / test_pdf_name).exists(), "El PDF no se recreó desde el base64 de .cait"
            assert (temp_cait_dir / "reports" / sample_draft_name).exists(), "El borrador no se recreó desde .cait"
            print(f"7C. Importación de archivo .cait (JSON) exitosa con reconstrucción de PDFs y borradores.")
        finally:
            shutil.rmtree(temp_cait_dir, ignore_errors=True)

    finally:
        shutil.rmtree(temp_isolated_dir, ignore_errors=True)

    print("\n>>> ¡TODAS LAS PRUEBAS DE EXPORTACIÓN, IMPORTACIÓN Y COPIA PASARON EXITOSAMENTE! <<<")

if __name__ == "__main__":
    run_test()
