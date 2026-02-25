"""Helper utilities to manage technical counterpart metadata stored on disk."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, List, Optional

DEFAULT_DB_PATH = Path(__file__).resolve().parents[2] / "data" / "databases" / "counterparts.json"


n_slug_invalid = re.compile(r"[^a-z0-9]+")


class CounterpartRepository:
    """Administra el catalogo persistente de contrapartes tecnicas."""

    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = Path(db_path) if db_path else DEFAULT_DB_PATH
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_storage()

    def _ensure_storage(self) -> None:
        if not self.db_path.exists():
            self.save_all([])

    def load_all(self) -> List[Dict]:
        try:
            data = json.loads(self.db_path.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
        self.save_all([])
        return []

    def save_all(self, entries: List[Dict]) -> None:
        with self.db_path.open("w", encoding="utf-8") as handler:
            json.dump(entries, handler, ensure_ascii=False, indent=2)

    def list_all(self) -> List[Dict]:
        entries = self.load_all()
        return sorted(entries, key=lambda item: (int(item.get("priority", 999)), item.get("name", "")))

    def add_counterpart(self, payload: Dict) -> Dict:
        entries = self.load_all()
        entry = self._build_entry(payload, entries)
        entries.append(entry)
        self.save_all(entries)
        return entry

    def remove_counterpart(self, counterpart_id: str) -> bool:
        if not counterpart_id:
            return False

        entries = self.load_all()
        remaining = [entry for entry in entries if entry.get("id") != counterpart_id]
        if len(remaining) == len(entries):
            return False

        self.save_all(remaining)
        return True

    def _build_entry(self, payload: Dict, existing: List[Dict]) -> Dict:
        name = (payload.get("name") or "").strip()
        if not name:
            raise ValueError("El nombre de la contraparte tecnica es obligatorio.")

        role = (payload.get("role") or "").strip()
        candidate_id = (payload.get("id") or "").strip() or self._generate_unique_id(name, existing)
        priority = int(payload.get("priority") or (len(existing) + 1))

        return {
            "id": candidate_id,
            "name": name,
            "role": role,
            "priority": priority,
        }

    def _generate_unique_id(self, name: str, existing: List[Dict]) -> str:
        base = n_slug_invalid.sub("-", self._strip_accents(name.lower())).strip("-")
        base = base or "contraparte"
        suffix = 1
        unique_id = base
        existing_ids = {entry.get("id") for entry in existing}
        while unique_id in existing_ids:
            suffix += 1
            unique_id = f"{base}-{suffix}"
        return unique_id

    @staticmethod
    def _strip_accents(value: str) -> str:
        return (
            value.replace("á", "a")
            .replace("é", "e")
            .replace("í", "i")
            .replace("ó", "o")
            .replace("ú", "u")
            .replace("ñ", "n")
        )
