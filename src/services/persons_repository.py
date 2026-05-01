"""Repositorio local de personas evaluadas, indexado por cédula."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Dict, List, Optional


def _get_default_db_path() -> Path:
    runtime_data_root = os.getenv("CAIT_DATA_ROOT")
    if runtime_data_root:
        return Path(runtime_data_root) / "databases" / "persons.json"
    return Path(__file__).resolve().parents[2] / "data" / "databases" / "persons.json"


class PersonsRepository:
    """Almacena datos de personas evaluadas indexados por cédula.

    Estructura de cada registro::

        {
            "identification": "8-1234-56789",
            "name": "Juan Pérez",
            "age": "35",
            "position": "Operario",
            "last_result_label": "Normal",
            "last_test_type": "audiometria",
        }
    """

    def __init__(self, db_path: Optional[Path] = None):
        resolved = Path(db_path) if db_path else _get_default_db_path()
        self.db_path = resolved
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.db_path.exists():
            self.save_all({})

    # ------------------------------------------------------------------
    # Storage helpers
    # ------------------------------------------------------------------
    def load_all(self) -> Dict[str, dict]:
        """Devuelve el dict completo { cedula: record }."""
        try:
            data = json.loads(self.db_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (json.JSONDecodeError, OSError):
            pass
        return {}

    def save_all(self, records: Dict[str, dict]) -> None:
        with self.db_path.open("w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

    # ------------------------------------------------------------------
    # Queries
    # ------------------------------------------------------------------
    def get_by_id(self, identification: str) -> Optional[dict]:
        """Busca un registro por cédula exacta."""
        key = (identification or "").strip().upper()
        if not key:
            return None
        return self.load_all().get(key)

    def list_all(self) -> List[dict]:
        """Devuelve todos los registros como lista, ordenados por nombre."""
        records = self.load_all()
        return sorted(records.values(), key=lambda r: r.get("name", "").lower())

    def search(self, query: str) -> List[dict]:
        """Búsqueda parcial por cédula o nombre."""
        q = (query or "").strip().lower()
        if not q:
            return self.list_all()
        return [
            r for r in self.list_all()
            if q in r.get("identification", "").lower()
            or q in r.get("name", "").lower()
        ]

    # ------------------------------------------------------------------
    # Mutations
    # ------------------------------------------------------------------
    def upsert(self, record: dict) -> dict:
        """Inserta o actualiza un registro por cédula. Devuelve el registro guardado."""
        identification = (record.get("identification") or "").strip().upper()
        if not identification:
            raise ValueError("La cédula es obligatoria para guardar una persona.")

        records = self.load_all()
        existing = records.get(identification, {})
        merged = {**existing, **{k: v for k, v in record.items() if v is not None}}
        merged["identification"] = identification
        records[identification] = merged
        self.save_all(records)
        return merged

    def remove(self, identification: str) -> bool:
        """Elimina un registro. Devuelve True si existía."""
        key = (identification or "").strip().upper()
        records = self.load_all()
        if key not in records:
            return False
        del records[key]
        self.save_all(records)
        return True
