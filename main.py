"""
Aplicación de Generación de Informes Clínicos
Centro de Atención Integral Terapéutico Panamá

Módulo principal de la aplicación
"""

import sys
import os
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

# Agregar la ruta del proyecto al path de Python
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.ui.app import MainApplication


def _resolve_log_dir() -> Path:
    """Devuelve la carpeta de logs según modo instalado o desarrollo."""

    if getattr(sys, "frozen", False):
        return Path(os.getenv("LOCALAPPDATA") or Path.home()) / "CAIT Informes" / "logs"
    return Path(__file__).resolve().parent / "logs"


def _persist_startup_error(exc: Exception) -> Path | None:
    """Guarda el traceback del error de inicio y devuelve su ruta."""

    try:
        import traceback

        log_dir = _resolve_log_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "startup_error.log"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        details = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        content = (
            f"[{timestamp}] Error al iniciar CAIT Informes\n"
            f"Python: {sys.version}\n"
            f"Frozen: {getattr(sys, 'frozen', False)}\n"
            f"\n{details}\n"
        )
        log_path.write_text(content, encoding="utf-8")
        return log_path
    except Exception:
        return None


def _show_startup_error_dialog(exc: Exception, log_path: Path | None) -> None:
    """Muestra al usuario un mensaje claro cuando ocurre un fallo de arranque."""

    try:
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()

        msg = [
            "No se pudo iniciar CAIT Informes.",
            "",
            f"Detalle: {exc}",
        ]
        if log_path is not None:
            msg.extend(["", f"Se generó un log en: {log_path}"])

        messagebox.showerror("Error al iniciar", "\n".join(msg), parent=root)
        root.destroy()
    except Exception:
        pass


def main():
    """Punto de entrada principal de la aplicación"""
    try:
        app = MainApplication()
        app.run()
    except Exception as e:
        print(f"Error al iniciar la aplicación: {e}")
        log_path = _persist_startup_error(e)
        _show_startup_error_dialog(e, log_path)
        sys.exit(1)


if __name__ == "__main__":
    main()
