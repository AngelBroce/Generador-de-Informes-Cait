"""Utilities to build the static content outline for reports."""

from typing import List


def _normalize_report_type(report_type: str) -> str:
    if not report_type:
        return "espirometria"
    text = report_type.lower()
    if "audiometr" in text and "espirom" in text:
        return "combinado"
    if "audiometr" in text:
        return "audiometria"
    return "espirometria"


def get_content_outline(report_type: str) -> List[str]:
    """Returns the static outline shown on the second page and in the UI preview."""

    normalized = _normalize_report_type(report_type)

    if normalized == "audiometria":
        tests_label = "AUDIOMETRÍAS"
    elif normalized == "combinado":
        tests_label = "AUDIOMETRÍAS Y ESPIROMETRÍAS"
    else:
        tests_label = "ESPIROMETRÍAS"

    results_title = f"CUADRO DE RESULTADOS DE LAS PRUEBAS DE {tests_label}."
    protocol_title = f"PROTOCOLO DE LAS PRUEBAS DE {tests_label}."
    specific_section = f"{tests_label}."

    return [
        "DATOS DE LA EMPRESA.",
        results_title,
        "ESTADISTICA DE RESULTADOS.",
        "CONCLUSIÓN.",
        "RECOMENDACIÓN.",
        "EQUIPO TÉCNICO",
        "CERTIFICADOS DE CALIBRACIÓN.",
        specific_section,
        "LISTADOS DE AISTENCIA.",
        protocol_title,
    ]
