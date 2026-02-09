"""
Modelo de datos para informes clínicos
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


@dataclass
class Report:
    """Modelo de informe clínico"""
    
    report_id: str
    report_type: str  # audiometría o espirometría
    company: str
    location: str
    evaluation_date: datetime
    evaluator: str
    evaluated_persons: List['EvaluatedPerson'] = field(default_factory=list)
    attachments: List[str] = field(default_factory=list)
    conclusions: str = ""
    recommendations: str = ""
    statistical_summary: dict = field(default_factory=dict)
    
    def add_person(self, person: 'EvaluatedPerson') -> None:
        """Agrega una persona evaluada"""
        self.evaluated_persons.append(person)
    
    def add_attachment(self, file_path: str) -> None:
        """Agrega un archivo adjunto"""
        if file_path not in self.attachments:
            self.attachments.append(file_path)
    
    def to_dict(self) -> dict:
        """Convierte el informe a diccionario"""
        return {
            "report_id": self.report_id,
            "report_type": self.report_type,
            "company": self.company,
            "location": self.location,
            "evaluation_date": self.evaluation_date.isoformat(),
            "evaluator": self.evaluator,
            "evaluated_count": len(self.evaluated_persons),
            "attachments_count": len(self.attachments),
            "conclusions": self.conclusions,
            "recommendations": self.recommendations
        }
