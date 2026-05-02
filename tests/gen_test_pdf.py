"""Script de prueba: genera un PDF de muestra para verificar los cambios."""
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.services.pdf_generator import PDFGenerator

# Datos de prueba realistas
report_data = {
    "company": "Empresa Ejemplo S.A.",
    "location": "Ciudad de Panamá",
    "plant": "Área de Producción",
    "activity": "Manufactura",
    "country": "Panamá",
    "date": "02/05/2026",
    "study_dates": "30 de abril y 1 de mayo del 2026",
    "type": "Audiometría y Espirometría",
    "evaluator_name": "Lic. María López",
    "evaluator_title": "Fonoaudióloga",
    "evaluator_license": "12345",
    "calibration_files": [],
    "calibration_certificates": [],
    "audiogram_files": [],
    "test_attachment_files": {"audiometria": [], "espirometria": []},
    "attendance_files": [],
    "evaluated": [
        {"name": "Juan Pérez", "identification": "8-123-4567", "age": "35", "position": "Operario",
         "result_label": "Normal bilateral", "result_code": "normal", "test_type": "audiometria"},
        {"name": "Ana García", "identification": "4-567-8901", "age": "28", "position": "Supervisora",
         "result_label": "Caída leve bilateral", "result_code": "caida_leve_bilateral", "test_type": "audiometria"},
        {"name": "Carlos Ruiz", "identification": "8-234-5678", "age": "42", "position": "Mecánico",
         "result_label": "Caída moderada a severa bilateral", "result_code": "caida_moderada_a_severa_bilateral", "test_type": "audiometria"},
        {"name": "Luis Mendoza", "identification": "8-345-6789", "age": "50", "position": "Técnico",
         "result_label": "Espirometría normal", "result_code": "espiro_normal", "test_type": "espirometria"},
        {"name": "Rosa Torres", "identification": "8-456-7890", "age": "38", "position": "Operaria",
         "result_label": "Restricción leve", "result_code": "restriccion_leve", "test_type": "espirometria"},
    ],
    "conclusion_text": (
        "La empresa Ejemplo S.A. realizó las audiometrías y espirometrías durante el 30 de abril y 1 de mayo del 2026.\n\n"
        "Se evaluaron 5 colaboradores en el Área de Producción.\n\n"
        "Los resultados fueron satisfactorios en su mayoría, con algunos casos de vigilancia ocupacional.\n\n"
        "Se recomienda seguimiento a los colaboradores con caídas auditivas."
    ),
    "recommendations_text": (
        "• Se recomienda la realización periódica de audiometrías y espirometrías laborales.\n"
        "• Continuar suministrando equipos de protección auditiva y respiratoria.\n"
        "• Los colaboradores con resultados alterados deben ser referidos a especialistas."
    ),
    "link_mode": "absolute",
}

output_path = os.path.join(os.path.dirname(__file__), "test_output.pdf")
logo_path = os.path.abspath("logo-apli-removebg-preview.ico")

gen = PDFGenerator()
gen.generate(report_data, output_path, logo_path=logo_path if os.path.exists(logo_path) else None)
print(f"✅ PDF generado: {output_path}")
os.startfile(output_path)
