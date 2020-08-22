"""Deobfuscation tools."""

from enum import Enum

from .abstract_syntax_tree import *
from .compiler import Program
from .error import DeobfuscationError


class Deobfuscator:
    """Class performing AST deobfuscation."""
    class EvaluationLevel(Enum):
        """
        Used to specify what expressions should be evaluated : none of them,
        literal expressions or pure functions.
        """
        NONE = 0
        LITERAL = 1
        FUNCTION = 2

    evaluation_level: "Deobfuscator.EvaluationLevel"
    rename_symbols: bool

    def __init__(
        self,
        program: Program,
        evaluation_level: "Deobfuscator.EvaluationLevel" = 1,
        rename_symbols: bool = False,
    ) -> None:
        self.evaluation_level = evaluation_level
        self.rename_symbols = rename_symbols

    def deobfuscate(self, ast: AST) -> AST:
        """
        Return the de-obfuscated version of an AST.
        """
        return ast
