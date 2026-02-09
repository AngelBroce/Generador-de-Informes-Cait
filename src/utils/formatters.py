"""
Formateadores de datos
"""

from datetime import datetime
from typing import Any


class Formatter:
    """Clase para formatear datos"""
    
    @staticmethod
    def format_date(date_obj: datetime, format_str: str = "%d/%m/%Y") -> str:
        """Formatea una fecha"""
        try:
            return date_obj.strftime(format_str)
        except:
            return str(date_obj)
    
    @staticmethod
    def format_datetime(date_obj: datetime, format_str: str = "%d/%m/%Y %H:%M") -> str:
        """Formatea una fecha y hora"""
        try:
            return date_obj.strftime(format_str)
        except:
            return str(date_obj)
    
    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """Convierte bytes a formato legible"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"
    
    @staticmethod
    def format_currency(amount: float, currency: str = "USD") -> str:
        """Formatea una cantidad como moneda"""
        return f"{currency} {amount:,.2f}"
    
    @staticmethod
    def format_percentage(value: float, decimals: int = 2) -> str:
        """Formatea un porcentaje"""
        return f"{value:.{decimals}f}%"
