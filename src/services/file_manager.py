"""
Servicio de gestión de archivos
Maneja operaciones con archivos del sistema
"""

import os
import shutil
from typing import List
from pathlib import Path


class FileManager:
    """Gestor de archivos del sistema"""
    
    def __init__(self, base_path: str = None):
        """
        Inicializa el gestor de archivos
        
        Args:
            base_path: Ruta base del proyecto
        """
        self.base_path = base_path or os.getcwd()
        self.reports_dir = os.path.join(self.base_path, "data", "reports")
        self.attachments_dir = os.path.join(self.base_path, "data", "attachments")
        self.exports_dir = os.path.join(self.base_path, "data", "exports")
    
    def validate_pdf(self, file_path: str) -> bool:
        """
        Valida que un archivo sea PDF válido
        
        Args:
            file_path: Ruta del archivo
            
        Returns:
            True si es un PDF válido
        """
        if not os.path.exists(file_path):
            return False
        
        if not file_path.lower().endswith('.pdf'):
            return False
        
        # Verificar que sea un PDF válido
        try:
            with open(file_path, 'rb') as f:
                header = f.read(4)
                return header == b'%PDF'
        except:
            return False
    
    def copy_attachment(self, source_path: str, report_id: str) -> str:
        """
        Copia un archivo adjunto a la carpeta de almacenamiento
        
        Args:
            source_path: Ruta del archivo original
            report_id: ID del informe
            
        Returns:
            Ruta del archivo copiado o vacío si falla
        """
        if not self.validate_pdf(source_path):
            return ""
        
        try:
            report_attachments_dir = os.path.join(self.attachments_dir, report_id)
            os.makedirs(report_attachments_dir, exist_ok=True)
            
            filename = os.path.basename(source_path)
            destination = os.path.join(report_attachments_dir, filename)
            
            shutil.copy2(source_path, destination)
            return destination
        except Exception as e:
            print(f"Error al copiar archivo: {e}")
            return ""
    
    def delete_attachment(self, file_path: str) -> bool:
        """
        Elimina un archivo adjunto
        
        Args:
            file_path: Ruta del archivo
            
        Returns:
            True si se eliminó exitosamente
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception as e:
            print(f"Error al eliminar archivo: {e}")
            return False
    
    def list_attachments(self, report_id: str) -> List[str]:
        """
        Lista los archivos adjuntos de un informe
        
        Args:
            report_id: ID del informe
            
        Returns:
            Lista de rutas de archivos
        """
        report_attachments_dir = os.path.join(self.attachments_dir, report_id)
        
        if not os.path.exists(report_attachments_dir):
            return []
        
        files = []
        for filename in os.listdir(report_attachments_dir):
            file_path = os.path.join(report_attachments_dir, filename)
            if os.path.isfile(file_path):
                files.append(file_path)
        
        return files
    
    def get_file_size(self, file_path: str) -> int:
        """
        Obtiene el tamaño de un archivo en bytes
        
        Args:
            file_path: Ruta del archivo
            
        Returns:
            Tamaño en bytes
        """
        try:
            return os.path.getsize(file_path)
        except:
            return 0
