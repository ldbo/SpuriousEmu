from abc import abstractmethod
from dataclasses import dataclass
from inspect import getfullargspec
from typing import Callable, List, Optional

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
    Signature = Callable[["Interpreter", List[Value]], Value]
    external_function: "ExternalFunction.Signature"

    def call(self, interpreter: "Interpreter", arguments_values: List[Value]) \
            -> None:
        self.external_function(interpreter, arguments_values)

    @staticmethod
    def from_function(
            python_function: Optional["ExternalFunction.Signature"] = None,
            name: Optional[str] = None) -> "ExternalFunction.Signature":
        """
        Create an ExternalFunction from a Python function, optionally changing
        its name.
        """
        def decorator(python_function: "ExternalFunction.Signature") \
                -> "ExternalFunction":
            new_name = python_function.__name__ if name is None else name
            arguments_names = list(getfullargspec(python_function)[0])
            return ExternalFunction(new_name, arguments_names, python_function)

        if python_function is None:
            return decorator
        else:
            return decorator(python_function)


@dataclass
class InternalFunction(Function):
    """VBA function."""
    body: Block

    def call(self, interpreter: "Interpreter", arguments_values: List[Value]) \
            -> None:
        interpreter.call_function(self, arguments_values)
