import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import json
from fastapi.testclient import TestClient
from api_main import app

client = TestClient(app)

def test_draft_persistence():
    print("=== INICIANDO PRUEBA DE PERSISTENCIA DE BORRADORES Y NAVEGACIÓN ===")
    
    # 1. Test drafts list
    res_drafts = client.get('/api/drafts')
    assert res_drafts.status_code == 200
    drafts = res_drafts.json()
    pedregal_drafts = [d['name'] for d in drafts if 'pedregal' in d['name'].lower()]
    print(f"1. Borradores con pedregal en lista: {pedregal_drafts}")
    assert len(pedregal_drafts) == 1, f"Debe haber exactamente 1 borrador de pedregal, se encontraron: {pedregal_drafts}"

    # 2. Test load draft
    res_load = client.post('/api/drafts/load', json={'name': 'pedregal'})
    assert res_load.status_code == 200, res_load.text
    rep_data = res_load.json()['report']
    print(f"2. Cargado borrador: {rep_data.get('_draft_name')}, Tipo: {rep_data.get('report_type')}, Total espiro: {len(rep_data.get('resultados_espirometria', []))}")
    assert len(rep_data.get('resultados_espirometria', [])) == 60, "Debe tener 60 resultados"

    # 3. Simulate page change: autoSave sends partial data from an uninitialized page
    res_save_partial = client.post('/api/report', json={
        'report_type': 'audiometria',
        'company_name': '',
        'resultados_espirometria': []
    })
    assert res_save_partial.status_code == 200

    # 4. Check GET report
    res_get = client.get('/api/report')
    rep_get = res_get.json()
    print(f"3. Tras guardado parcial -> Tipo: {rep_get.get('report_type')}, Empresa: {rep_get.get('company_name')}, Espiro: {len(rep_get.get('resultados_espirometria', []))}")
    
    assert rep_get.get('report_type') == 'espirometria', f"Tipo de reporte incorrecto: {rep_get.get('report_type')}"
    assert len(rep_get.get('resultados_espirometria', [])) == 60, f"Resultados perdidos: {len(rep_get.get('resultados_espirometria', []))}"
    assert rep_get.get('company_name') == 'Toledano Planta Pedregal', f"Empresa perdida: {rep_get.get('company_name')}"
    assert rep_get.get('_draft_name') == 'pedregal.json', f"Draft name perdido: {rep_get.get('_draft_name')}"

    # Restaurar estado original
    with open(os.path.join(os.path.dirname(__file__), "..", "data", "databases", "toledano pedregal.json"), "r", encoding="utf-8") as f:
        clean_data = json.load(f)
    for target_name in ["pedregal.json", "current_report.json"]:
        with open(os.path.join(os.path.dirname(__file__), "..", "data", "reports", target_name), "w", encoding="utf-8") as f:
            json.dump(clean_data, f, ensure_ascii=False, indent=4)

    print("=== TODAS LAS PRUEBAS DE PROTECCIÓN PASARON CON ÉXITO ===")

if __name__ == "__main__":
    test_draft_persistence()
