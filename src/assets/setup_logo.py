"""
Script para generar el logo como imagen PNG
Esto se ejecuta una sola vez durante la instalación
"""

import base64
from pathlib import Path

# Logo CAIT Panamá (pequeño, para demostración)
# En producción, la imagen real del logo se coloca en esta carpeta

logo_base64 = """
iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==
"""

# Decodificar y guardar
logo_data = base64.b64decode(logo_base64.strip())
logo_path = Path(__file__).parent / "logo_cait.png"

with open(logo_path, 'wb') as f:
    f.write(logo_data)

print(f"Logo guardado en: {logo_path}")
