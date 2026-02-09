"""
Modelo de datos para persona evaluada
"""

from dataclasses import dataclass, field
from typing import Dict, Any


@dataclass
class EvaluatedPerson:
    """Modelo de persona evaluada"""
    
    full_name: str
    identification: str
    position: str
    department: str = ""
    test_results: Dict[str, Any] = field(default_factory=dict)
    observations: str = ""
    
    def to_dict(self) -> dict:
        """Convierte a diccionario"""
        return {
            "full_name": self.full_name,
            "identification": self.identification,
            "position": self.position,
            "department": self.department,
            "test_results": self.test_results,
            "observations": self.observations
        }
