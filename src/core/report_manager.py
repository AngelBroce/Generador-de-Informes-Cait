"""
Gestor central de informes clínicos
Maneja la creación, edición y exportación de informes
"""

import os
from datetime import datetime
from typing import Optional, Dict, List


class ReportManager:
    """Gestor principal de informes"""
    
    def __init__(self):
        """Inicializa el gestor de informes"""
        self.current_report = None
        self.report_history = []
        
    def create_report(self, report_type: str, company: str, evaluator: str) -> Dict:
        """
        Crea un nuevo informe
        
        Args:
            report_type: Tipo de informe (audiometría o espirometría)
            company: Empresa asociada
            evaluator: Responsable de la evaluación
            
        Returns:
            Diccionario con datos del informe
        """
        report = {
            "id": self._generate_report_id(),
            "type": report_type,
            "company": company,
            "evaluator": evaluator,
            "date": datetime.now().isoformat(),
            "evaluated": [],
            "attachments": [],
            "conclusions": "",
            "recommendations": ""
        }
        self.current_report = report
        return report
    
    def add_evaluated_person(self, name: str, identification: str, 
                            position: str, results: Dict) -> bool:
        """
        Agrega una persona evaluada al informe
        
        Args:
            name: Nombre completo
            identification: Cédula o identificación
            position: Cargo o área
            results: Resultados de la evaluación
            
        Returns:
            True si se agregó exitosamente
        """
        if not self.current_report:
            return False
            
        evaluated = {
            "name": name,
            "identification": identification,
            "position": position,
            "results": results
        }
        self.current_report["evaluated"].append(evaluated)
        return True
    
    def add_attachment(self, file_path: str) -> bool:
        """
        Agrega un archivo adjunto al informe
        
        Args:
            file_path: Ruta del archivo PDF
            
        Returns:
            True si se agregó exitosamente
        """
        if not self.current_report:
            return False
            
        if not os.path.exists(file_path):
            return False
            
        if not file_path.lower().endswith('.pdf'):
            return False
            
        self.current_report["attachments"].append(file_path)
        return True
    
    def _generate_report_id(self) -> str:
        """Genera un ID único para el informe"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        return f"REP_{timestamp}"
    
    def get_current_report(self) -> Optional[Dict]:
        """Obtiene el informe actual"""
        return self.current_report
    
    def save_report(self) -> bool:
        """Guarda el informe actual"""
        if not self.current_report:
            return False
        self.report_history.append(self.current_report.copy())
        return True
