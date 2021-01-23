from dataclasses import dataclass
from enum import Enum
from typing import Dict, Optional

from .data import DeclaredType


class ComparisonMode(Enum):
    BINARY = "Binary"
    TEXT = "Text"


class LowerBound(Enum):
    ZERO = 0
    ONE = 1


class VariableDeclarationMode(Enum):
    IMPLICIT = "Implicit"
    EXPLICIT = "Explicit"


class ClassModuleAccessibility(Enum):
    PRIVATE = "Private"
    PUBLIC_CREATABLE = "Public Creatable"
    PUBLIC_NON_CREATABLE = "Public Non Creatable"


class ProceduralModuleAccessibility(Enum):
    PRIVATE = "Private"
    PUBLIC = "Public"


@dataclass(frozen=True)
class LetterSpec:
    lower_bound: str
    upper_bound: Optional[str]

    def is_universal_range(self):
        return (self.lower_bound, self.upper_bound) == ("A", "Z")


ImplicitTypeRules = Dict[LetterSpec, DeclaredType]
