"""
Modelo de datos para archivos adjuntos
"""

from dataclasses import dataclass
from datetime import datetime
import os


@dataclass
class Attachment:
    """Modelo de archivo adjunto"""
    
    file_path: str
    file_name: str
    file_type: str = "pdf"
    added_date: datetime = None
    file_size: int = 0
    
    def __post_init__(self):
        """Inicialización post-creación"""
        if self.added_date is None:
            self.added_date = datetime.now()
        
        if os.path.exists(self.file_path):
            self.file_size = os.path.getsize(self.file_path)
    
    def is_valid(self) -> bool:
        """Valida que el archivo exista y sea PDF"""
        return (os.path.exists(self.file_path) and 
                self.file_path.lower().endswith('.pdf'))
    
    def to_dict(self) -> dict:
        """Convierte a diccionario"""
        return {
            "file_name": self.file_name,
            "file_path": self.file_path,
            "file_type": self.file_type,
            "added_date": self.added_date.isoformat(),
            "file_size": self.file_size
        }
