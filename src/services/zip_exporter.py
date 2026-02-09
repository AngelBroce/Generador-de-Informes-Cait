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
                      attachments: List[str], output_dir: str) -> str:
        """
        Crea un paquete ZIP con el informe y anexos
        
        Args:
            report_name: Nombre del informe
            pdf_path: Ruta del PDF principal
            attachments: Lista de rutas de archivos adjuntos
            output_dir: Directorio de salida
            
        Returns:
            Ruta del archivo ZIP creado
        """
        try:
            # Crear carpeta temporal
            temp_dir = os.path.join(output_dir, report_name)
            os.makedirs(temp_dir, exist_ok=True)
            
            # Copiar PDF principal
            if os.path.exists(pdf_path):
                import shutil
                shutil.copy(pdf_path, os.path.join(temp_dir, os.path.basename(pdf_path)))
            
            # Crear carpeta de anexos
            if attachments:
                attachments_dir = os.path.join(temp_dir, "Anexos")
                os.makedirs(attachments_dir, exist_ok=True)
                
                for attachment in attachments:
                    if os.path.exists(attachment):
                        import shutil
                        shutil.copy(attachment, attachments_dir)
            
            # Crear ZIP
            zip_path = os.path.join(output_dir, f"{report_name}.zip")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, output_dir)
                        zipf.write(file_path, arcname)
            
            # Limpiar carpeta temporal
            import shutil
            shutil.rmtree(temp_dir)
            
            return zip_path
            
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
