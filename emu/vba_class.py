from dataclasses import dataclass
from typing import List

from .reference import ClassModule


@dataclass
class Class:
    variables: List[str]
    class_reference: ClassModule
