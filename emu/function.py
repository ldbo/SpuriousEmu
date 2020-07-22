from abc import abstractmethod
from dataclasses import dataclass
from typing import Any, Callable, List

from .value import Value
from .abstract_syntax_tree import Block


@dataclass
class Function:
    """An abstract function, be it implemented in VBA or Python."""
    name: str
    arguments_names: List[str]

    @abstractmethod
    def call(self, interpreter: "Interpreter", arguments_values: List[Value]) \
            -> None:
        pass


@dataclass
class ExternalFunction(Function):
    """Python function."""
    external_function: Callable[[Any], Any]

    def call(self, interpreter: "Interpreter", arguments_values: List[Value]) \
            -> None:
        self.external_function(interpreter, arguments_values)


@dataclass
class InternalFunction(Function):
    """VBA function."""
    body: Block

    def call(self, interpreter: "Interpreter", arguments_values: List[Value]) \
            -> None:
        interpreter.call_function(self, arguments_values)
