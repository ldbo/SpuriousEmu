from abc import abstractmethod
from dataclasses import dataclass
from typing import Callable, List

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
    """
    Python function wrapper, the function will be called with two arguments:
      - a reference to the interpreter, allowing to have side effects
      - a list of Value objects, the actual arguments of the function
    It has to return a Value object.
    """
    external_function: Callable[["Interpreter", List[Value]], Value]

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
