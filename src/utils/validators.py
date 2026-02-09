"""
Validadores de datos
"""

import re
from typing import Any


class Validator:
    """Clase para validación de datos"""
    
    @staticmethod
    def is_valid_email(email: str) -> bool:
        """Valida formato de email"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    @staticmethod
    def is_valid_identification(identification: str) -> bool:
        """Valida formato de cédula/identificación"""
        # Eliminar espacios y guiones
        clean_id = identification.replace(" ", "").replace("-", "")
        return clean_id.isalnum() and len(clean_id) >= 8
    
    @staticmethod
    def is_valid_name(name: str) -> bool:
        """Valida que el nombre tenga caracteres válidos"""
        return len(name.strip()) >= 3 and all(
            c.isalpha() or c.isspace() for c in name
        )
    
    @staticmethod
    def is_non_empty(value: Any) -> bool:
        """Valida que no esté vacío"""
        if isinstance(value, str):
            return len(value.strip()) > 0
        return value is not None
    
    @staticmethod
    def is_valid_phone(phone: str) -> bool:
        """Valida formato de teléfono"""
        clean_phone = phone.replace(" ", "").replace("-", "").replace("+", "")
        return clean_phone.isdigit() and len(clean_phone) >= 7
