"""
Aplicación de Generación de Informes Clínicos
Centro de Atención Integral Terapéutico Panamá

Módulo principal de la aplicación
"""

import sys
import os

# Agregar la ruta del proyecto al path de Python
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.ui.app import MainApplication


def main():
    """Punto de entrada principal de la aplicación"""
    try:
        app = MainApplication()
        app.run()
    except Exception as e:
        print(f"Error al iniciar la aplicación: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
