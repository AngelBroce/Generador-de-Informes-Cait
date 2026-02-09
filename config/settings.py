"""
Configuración general de la aplicación
"""

import os
from pathlib import Path

# Rutas
BASE_DIR = Path(__file__).parent.parent
SRC_DIR = BASE_DIR / "src"
DATA_DIR = BASE_DIR / "data"
REPORTS_DIR = DATA_DIR / "reports"
ATTACHMENTS_DIR = DATA_DIR / "attachments"
EXPORTS_DIR = DATA_DIR / "exports"
LOGS_DIR = BASE_DIR / "logs"
TEMPLATES_DIR = SRC_DIR / "templates"
ASSETS_DIR = SRC_DIR / "assets"

# Configuración de PDF
PDF_CONFIG = {
    "page_width": 8.5,  # pulgadas
    "page_height": 11,
    "left_margin": 0.5,
    "right_margin": 0.5,
    "top_margin": 0.75,
    "bottom_margin": 0.75,
    "default_font": "Helvetica",
    "title_font_size": 14,
    "heading_font_size": 12,
    "body_font_size": 10,
}

# Configuración de aplicación
APP_CONFIG = {
    "app_name": "Generador de Informes Clínicos",
    "version": "1.0.0",
    "organization": "Centro de Atención Integral Terapéutico Panamá",
    "debug": False,
}

# Tipos de informes soportados
REPORT_TYPES = [
    "audiometría",
    "espirometría"
]

# Categorías de resultados
RESULT_CATEGORIES = {
    "normal": "Normal",
    "leve": "Leve",
    "moderado": "Moderado",
    "severo": "Severo"
}

# Crear directorios si no existen
for directory in [REPORTS_DIR, ATTACHMENTS_DIR, EXPORTS_DIR, LOGS_DIR]:
    directory.mkdir(parents=True, exist_ok=True)
