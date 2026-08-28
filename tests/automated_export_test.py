import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from fastapi.testclient import TestClient
from api_main import app


client = TestClient(app)

def test_full_export_flow():
    print("\n--- INICIANDO PRUEBA AUTOMATIZADA DE EXPORTACION ---")
    
    # 1. Simular el llenado de datos (Presentacion, Resultados, Conclusion)
    sample_report = {
        "report_type": "audiometria_espirometria",
        "company_name": "Empresa de Prueba S.A.",
        "location": "Ciudad de Panama",
        "evaluation_date": "2026-05-13",
        "evaluation_dates": "12 y 13 de Mayo, 2026",
        "evaluator_main": "eval-001",
        "conclusion_text": "Esta es una conclusion de prueba generada automaticamente para validar el flujo del sistema.",
        "recommendations_text": "* Recomendacion 1\n* Recomendacion 2",
        "evaluated": [
            {"name": "Juan Perez", "id": "8-000-000", "result": "Normal"},
            {"name": "Maria Lopez", "id": "4-123-456", "result": "Vigilancia"}
        ],
        "adjuntos": [
            {"name": "test_adjunto.pdf", "tipo": "Certificado Calibracion"}
        ]
    }
    
    print("Step 1: Guardando datos del reporte...")
    res_save = client.post("/api/report", json=sample_report)
    assert res_save.status_code == 200
    print("OK: Reporte guardado.")

    # 2. Simular la subida de un archivo adjunto real
    print("Step 2: Simulando subida de archivo adjunto...")
    test_file_path = Path("test_adjunto.pdf")
    with open(test_file_path, "w") as f: f.write("%PDF-1.4 TEST FILE")
    
    with open(test_file_path, "rb") as f:
        res_upload = client.post("/api/upload-attachment", files={"file": ("test_adjunto.pdf", f, "application/pdf")})
    
    assert res_upload.status_code == 200
    print("OK: Archivo adjunto subido.")

    # 3. Disparar la creacion del ZIP (que ahora genera el PDF real)
    print("Step 3: Disparando generacion de ZIP y PDF...")
    res_export = client.post("/api/export-zip")
    assert res_export.status_code == 200
    data = res_export.json()
    assert data["status"] == "ok"
    zip_path = data["zip_path"]
    print(f"OK: ZIP generado en: {zip_path}")

    # 4. Verificar existencia de archivos
    print("Step 4: Verificando integridad de archivos...")
    assert os.path.exists(zip_path)
    
    # El PDF se genera dentro de la carpeta de exportación
    folder_path = Path(data.get("folder_path", ""))
    pdf_files = list(folder_path.glob("*.pdf")) if folder_path.exists() else []
    assert len(pdf_files) > 0, f"No se encontró el PDF en {folder_path}"
    pdf_path = pdf_files[0]
    assert os.path.exists(pdf_path)
    print(f"OK: PDF verificado: {pdf_path}")

    
    print("\n--- PRUEBA COMPLETADA CON EXITO ---")
    print(f"Resultado final: ZIP y PDF creados correctamente para '{sample_report['company_name']}'")

if __name__ == "__main__":
    try:
        test_full_export_flow()
    except Exception as e:
        print(f"\nERROR EN LA PRUEBA: {e}")
    finally:
        # Limpieza
        if os.path.exists("test_adjunto.pdf"): os.remove("test_adjunto.pdf")
