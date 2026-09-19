"""Helper utilities to manage clinical evaluator metadata stored on disk."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Dict, List, Optional


def _get_default_db_path() -> Path:
    runtime_data_root = os.getenv("CAIT_DATA_ROOT")
    if runtime_data_root:
        return Path(runtime_data_root) / "databases" / "evaluators.json"
    return Path(__file__).resolve().parents[2] / "data" / "databases" / "evaluators.json"


DEFAULT_EVALUATORS: List[Dict] = [
    {
        "id": "yara-lizeth-perez",
        "name": "Licda. Yara Lizeth Pérez A.",
        "title_label": "Licda.",
        "header_label": "Licenciada",
        "profession": "Fonoaudióloga",
        "registry": "Registro 180.",
        "credential_file": "data/attachments/idoneidad/idoneidad_yara.pdf",
        "technical_details": [
            "Fonoaudióloga,",
            "Registro 180.",
        ],
        "applicable_reports": ["audiometria"],
        "priority": 1,
    },
    {
        "id": "licdo-franklin-solano",
        "name": "Licdo. Franklin Solano",
        "title_label": "Licdo.",
        "header_label": "Licenciado",
        "profession": "Terapeuta Respiratorio",
        "registry": "Registro Profesional",
        "credential_file": "data/attachments/idoneidad/Licdo_Franklin_Solano.pdf",
        "technical_details": [
            "Terapeuta Respiratorio,",
            "Registro Profesional.",
        ],
        "applicable_reports": ["espirometria"],
        "priority": 2,
    },
    {
        "id": "licda-zenety-yauricel-campos-vega",
        "name": "Licda. Zenety Yauricel Campos Vega",
        "title_label": "Licda.",
        "header_label": "Licenciada",
        "profession": "Fonoaudióloga",
        "registry": "Registro Profesional",
        "credential_file": "data/attachments/idoneidad/Licda.Zenety Yauricel Campos Vega.pdf",
        "technical_details": [
            "Fonoaudióloga,",
            "Registro Profesional.",
        ],
        "applicable_reports": ["audiometria", "espirometria"],
        "priority": 3,
    },
    {
        "id": "stephanie-maria-thorne",
        "name": "Licda. Stephanie María Thorne.",
        "title_label": "Licda.",
        "header_label": "Licenciada",
        "profession": "Terapeuta Respiratoria",
        "registry": "Registro 124.",
        "credential_file": "data/attachments/idoneidad/idoneidad_stephanie.pdf",
        "technical_details": [
            "Terapeuta Respiratoria,",
            "Registro 124.",
        ],
        "applicable_reports": ["espirometria"],
        "priority": 4,
    },
]


def normalize_report_type(report_type: str) -> str:
    """Normaliza etiquetas de reportes para filtros consistentes."""

    if not report_type:
        return "audiometria"

    normalized = report_type.lower()
    has_audio = "audiometr" in normalized
    has_espiro = "espirom" in normalized

    if has_audio and has_espiro:
        return "combinado"
    if has_espiro:
        return "espirometria"
    return "audiometria"


def build_technical_details(profile: Dict) -> List[str]:
    """Devuelve las líneas de detalle mostradas en Equipo Técnico."""

    details = []
    raw_details = profile.get("technical_details")
    if isinstance(raw_details, list):
        details = [line.strip() for line in raw_details if isinstance(line, str) and line.strip()]

    if details:
        return details

    profession = (profile.get("profession") or "").strip()
    registry = (profile.get("registry") or "").strip()

    computed = []
    if profession:
        suffix = "" if profession.endswith((",", ".")) else ","
        computed.append(f"{profession}{suffix}")
    if registry:
        suffix = "" if registry.endswith((",", ".")) else "."
        computed.append(f"{registry}{suffix}")

    if computed:
        return computed

    name = (profile.get("name") or "").strip()
    return [name] if name else []


n_slug_invalid = re.compile(r"[^a-z0-9]+")


class EvaluatorRepository:
    """Administra el catálogo persistente de evaluadores."""

    def __init__(self, db_path: Optional[Path] = None):
        resolved_path = Path(db_path) if db_path else _get_default_db_path()
        self.db_path = resolved_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_storage()

    # ------------------------------------------------------------------
    # Storage helpers
    # ------------------------------------------------------------------
    def _ensure_storage(self) -> None:
        if not self.db_path.exists():
            self.save_all(DEFAULT_EVALUATORS)
            return
        entries = self.load_all()
        modified = False
        # Asegurar que los evaluadores base estén disponibles si no existen por id ni nombre
        for def_ev in DEFAULT_EVALUATORS:
            def_id = def_ev.get("id")
            def_name = (def_ev.get("name") or "").strip().lower()
            exists = any(
                e.get("id") == def_id or (e.get("name") or "").strip().lower() == def_name
                for e in entries
            )
            if not exists:
                entries.append(def_ev)
                modified = True
        if modified:
            self.save_all(entries)

    def load_all(self) -> List[Dict]:
        try:
            data = json.loads(self.db_path.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
        self.save_all(DEFAULT_EVALUATORS)
        return list(DEFAULT_EVALUATORS)

    def save_all(self, entries: List[Dict]) -> None:
        with self.db_path.open("w", encoding="utf-8") as handler:
            json.dump(entries, handler, ensure_ascii=False, indent=2)

    # ------------------------------------------------------------------
    # Queries
    # ------------------------------------------------------------------
    def list_all(self) -> List[Dict]:
        entries = self.load_all()
        return sorted(entries, key=lambda item: (int(item.get("priority", 999)), item.get("name", "")))

    def get_by_id(self, evaluator_id: str) -> Optional[Dict]:
        if not evaluator_id:
            return None
        evaluator_id_str = str(evaluator_id).strip()
        for entry in self.load_all():
            if entry.get("id") == evaluator_id_str:
                return entry
        return None

    def get_by_id_or_name(self, query: str) -> Optional[Dict]:
        """Busca un evaluador por ID exacto, nombre completo o slug normalizado."""
        if not query or not isinstance(query, str):
            return None
        q = query.strip()
        q_norm = self._strip_accents(q.lower())
        q_slug = n_slug_invalid.sub("-", q_norm).strip("-")
        
        for entry in self.load_all():
            e_id = entry.get("id", "")
            e_name = entry.get("name", "")
            e_name_norm = self._strip_accents(e_name.lower())
            e_name_slug = n_slug_invalid.sub("-", e_name_norm).strip("-")
            
            if e_id == q or e_name.lower() == q.lower():
                return entry
            if e_id == q_slug or e_name_slug == q_slug:
                return entry
            if q_norm and (q_norm in e_name_norm or e_name_norm in q_norm):
                return entry
        return None

    def get_primary_for_report(self, report_code: str) -> Optional[Dict]:
        code = report_code or ""
        candidates = [
            entry
            for entry in self.list_all()
            if code in (entry.get("applicable_reports") or [])
        ]
        return candidates[0] if candidates else None

    def get_team_for_report(self, report_type: str) -> List[Dict]:
        normalized = normalize_report_type(report_type)
        if normalized == "combinado":
            codes = ["audiometria", "espirometria"]
        else:
            codes = [normalized]

        team: List[Dict] = []
        seen_ids = set()
        for code in codes:
            entry = self.get_primary_for_report(code)
            if entry and entry.get("id") not in seen_ids:
                team.append(entry)
                seen_ids.add(entry["id"])
        return team

    # ------------------------------------------------------------------
    # Mutations
    # ------------------------------------------------------------------
    def add_evaluator(self, payload: Dict) -> Dict:
        entries = self.load_all()
        entry = self._build_entry(payload, entries)
        entries.append(entry)
        self.save_all(entries)
        return entry

    def upsert_evaluator(self, payload: Dict) -> Dict:
        """Inserta o actualiza un evaluador preservando su ID o creándolo si no existe."""
        if not payload or not isinstance(payload, dict):
            raise ValueError("Payload de evaluador inválido.")

        target_id = (payload.get("id") or "").strip()
        target_name = (payload.get("name") or "").strip()

        if not target_id and not target_name:
            raise ValueError("El evaluador requiere al menos ID o nombre.")

        existing_entry = None
        if target_id:
            existing_entry = self.get_by_id(target_id)
        if not existing_entry and target_name:
            existing_entry = self.get_by_id_or_name(target_name)

        if existing_entry:
            eval_id = existing_entry["id"]
            updates = dict(payload)
            # No sobreescribir con valores vacíos
            clean_updates = {k: v for k, v in updates.items() if v is not None and v != "" and v != []}
            updated = self.update_evaluator(eval_id, clean_updates)
            return updated or existing_entry
        else:
            return self.add_evaluator(payload)

    def update_evaluator(self, evaluator_id: str, updates: Dict) -> Optional[Dict]:
        """Actualiza un evaluador existente y devuelve el registro."""

        if not evaluator_id:
            return None

        entries = self.load_all()
        updated_entry = None
        for idx, entry in enumerate(entries):
            if entry.get("id") == evaluator_id:
                merged = dict(entry)
                merged.update({key: value for key, value in updates.items() if value is not None})
                entries[idx] = merged
                updated_entry = merged
                break

        if updated_entry:
            self.save_all(entries)
        return updated_entry

    def remove_evaluator(self, evaluator_id: str) -> bool:
        """Elimina un evaluador del catálogo. La Licda. Yara no puede ser eliminada."""

        if not evaluator_id:
            return False

        # Protección permanente: la Licda. Yara es evaluadora principal
        if evaluator_id == "yara-lizeth-perez":
            return False

        entries = self.load_all()
        remaining = [entry for entry in entries if entry.get("id") != evaluator_id]
        if len(remaining) == len(entries):
            return False

        self.save_all(remaining)
        return True

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _build_entry(self, payload: Dict, existing: List[Dict]) -> Dict:
        name = (payload.get("name") or "").strip()
        if not name:
            raise ValueError("El nombre del evaluador es obligatorio.")

        candidate_id = (payload.get("id") or "").strip() or self._generate_unique_id(name, existing)
        title_label = (payload.get("title_label") or "Licda.").strip()
        header_label = (payload.get("header_label") or "Licenciada").strip()
        profession = (payload.get("profession") or "").strip()
        registry = (payload.get("registry") or "").strip()
        technical_details = payload.get("technical_details")
        credential_file = (payload.get("credential_file") or "").strip()
        applicable_reports = self._build_applicable_reports(payload.get("applicable_reports"))
        priority = int(payload.get("priority") or (len(existing) + 1))

        entry = {
            "id": candidate_id,
            "name": name,
            "title_label": title_label,
            "header_label": header_label,
            "profession": profession,
            "registry": registry,
            "credential_file": credential_file,
            "technical_details": technical_details or build_technical_details(
                {
                    "profession": profession,
                    "registry": registry,
                    "name": name,
                }
            ),
            "applicable_reports": applicable_reports,
            "priority": priority,
        }
        return entry

    def _build_applicable_reports(self, raw: Optional[List[str]]) -> List[str]:
        if not raw:
            return ["audiometria"]
        normalized = []
        for item in raw:
            code = normalize_report_type(item)
            if code == "combinado":
                normalized.extend(["audiometria", "espirometria"])
            else:
                normalized.append(code)
        cleaned = []
        for code in normalized:
            if code not in cleaned:
                cleaned.append(code)
        return cleaned or ["audiometria"]

    def _generate_unique_id(self, name: str, existing: List[Dict]) -> str:
        base = n_slug_invalid.sub("-", self._strip_accents(name.lower())).strip("-")
        base = base or "evaluador"
        suffix = 1
        unique_id = base
        existing_ids = {entry.get("id") for entry in existing}
        while unique_id in existing_ids:
            suffix += 1
            unique_id = f"{base}-{suffix}"
        return unique_id

    @staticmethod
    def _strip_accents(value: str) -> str:
        replacements = (
            ("á", "a"),
            ("é", "e"),
            ("í", "i"),
            ("ó", "o"),
            ("ú", "u"),
            ("ñ", "n"),
        )
        normalized = value
        for accented, plain in replacements:
            normalized = normalized.replace(accented, plain)
        return normalized
