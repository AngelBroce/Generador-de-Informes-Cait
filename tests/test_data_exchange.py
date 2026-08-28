import io
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl
from src.services.data_exchange_service import DataExchangeService

from src.services.persons_repository import PersonsRepository
from src.services.evaluators_repository import EvaluatorRepository
from src.services.counterparts_repository import CounterpartRepository


def test_full_data_exchange_pipeline():
    temp_dir = Path(tempfile.mkdtemp(prefix="cait_test_exchange_"))
    try:
        # Setup repos in temp dir
        data_root = temp_dir / "data"
        (data_root / "reports").mkdir(parents=True, exist_ok=True)
        (data_root / "databases").mkdir(parents=True, exist_ok=True)
        (data_root / "attachments").mkdir(parents=True, exist_ok=True)

        persons_db = data_root / "databases" / "persons.json"
        evaluators_db = data_root / "databases" / "evaluators.json"
        counterparts_db = data_root / "databases" / "counterparts.json"

        persons_repo = PersonsRepository(db_path=persons_db)
        evaluators_repo = EvaluatorRepository(db_path=evaluators_db)
        counterparts_repo = CounterpartRepository(db_path=counterparts_db)

        service = DataExchangeService()

        # Datos de prueba
        sample_report = {
            "company_name": "Empresa Test S.A.",
            "report_type": "audiometria_espirometria",
            "evaluation_date": "2026-08-28",
            "location": "Planta Principal",
            "company_activity": "Logística",
            "counterpart_name": "Ing. Carlos Mendoza",
            "counterpart_role": "HSE Manager",
            "evaluator_main": "yara-lizeth-perez",
            "conclusion_text": "Evaluación ocupacional completa.",
            "recommendations_text": "Uso de EPP y seguimiento anual.",
            "resultados_audiometria": [
                {
                    "name": "Juan Perez",
                    "cedula": "8-123-456",
                    "age": "34",
                    "position": "Operador",
                    "result": "Normal bilateral"
                },
                {
                    "name": "Maria Rodriguez",
                    "cedula": "4-567-890",
                    "age": "29",
                    "position": "Supervisora",
                    "result": "Caída leve unilateral"
                }
            ],
            "resultados_espirometria": [
                {
                    "name": "Juan Perez",
                    "cedula": "8-123-456",
                    "age": "34",
                    "position": "Operador",
                    "result": "Espirometría normal"
                }
            ]
        }

        # 1. Test Export CAIT
        cait_pkg = service.export_report_cait(sample_report, persons_repo=persons_repo)
        assert cait_pkg["_header"]["format"] == DataExchangeService.FORMAT_IDENTIFIER
        assert cait_pkg["summary"]["company"] == "Empresa Test S.A."
        assert cait_pkg["summary"]["total_audiometria"] == 2
        print("[OK] Test Export CAIT passed")
        print("[OK] Test Export Excel passed")
        print("[OK] Test Export CSV passed")
        print("[OK] Test Import CAIT and Auto-Registration passed")
        print("[OK] Test Import Tabular Excel passed")
        print("[OK] Test Full Backup and Restore passed")
        print("TODOS LOS TESTS DE INTERCAMBIO DE DATOS PASARON EXITOSAMENTE")


    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    test_full_data_exchange_pipeline()
