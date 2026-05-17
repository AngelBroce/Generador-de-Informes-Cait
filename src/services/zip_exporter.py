"""
Servicio de exportación a ZIP
Crea paquetes comprimidos con informe y anexos
"""

import zipfile
import os
from typing import List
from pathlib import Path


class ZipExporter:
    """Exportador de informes a formato ZIP"""
    
    def __init__(self):
        """Inicializa el exportador"""
        pass
    
    def create_package(self, report_name: str, pdf_path: str, 
                      attachments: dict, output_dir: str) -> str:
        """
        Crea un paquete ZIP con el informe y anexos organizados por carpetas
        
        Args:
            report_name: Nombre del informe
            pdf_path: Ruta del PDF principal
            attachments: Diccionario { 'Carpeta': [lista_de_rutas] }
            output_dir: Directorio de salida
        """
        try:
            import shutil
            # Crear carpeta temporal
            temp_root = os.path.join(output_dir, report_name)
            if os.path.exists(temp_root):
                shutil.rmtree(temp_root)
            os.makedirs(temp_root, exist_ok=True)
            
            # 1. Copiar PDF principal
            if os.path.exists(pdf_path):
                shutil.copy2(pdf_path, os.path.join(temp_root, os.path.basename(pdf_path)))
            
            # 2. Copiar anexos categorizados
            adjuntos_root = os.path.join(temp_root, "Adjuntos")
            os.makedirs(adjuntos_root, exist_ok=True)
            
            for category, files in attachments.items():
                if not files: continue
                cat_dir = os.path.join(adjuntos_root, category)
                os.makedirs(cat_dir, exist_ok=True)
                
                for f_path in files:
                    if f_path and os.path.exists(f_path):
                        shutil.copy2(f_path, cat_dir)
            
            # 3. Crear ZIP
            zip_path = os.path.join(output_dir, f"{report_name}.zip")
            if os.path.exists(zip_path):
                os.remove(zip_path)
                
            shutil.make_archive(os.path.join(output_dir, report_name), 'zip', 
                               root_dir=output_dir, base_dir=report_name)
            
            # Limpiar carpeta temporal
            shutil.rmtree(temp_root)
            
            return f"{zip_path}"
            
        except Exception as e:
            print(f"Error al crear paquete ZIP: {e}")
            return ""
    
    def validate_zip(self, zip_path: str) -> bool:
        """
        Valida que un archivo ZIP sea válido
        
        Args:
            zip_path: Ruta del archivo ZIP
            
        Returns:
            True si el ZIP es válido
        """
        try:
            with zipfile.ZipFile(zip_path, 'r') as zipf:
                return zipf.testzip() is None
        except:
            return False
