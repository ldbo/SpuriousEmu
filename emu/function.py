from dataclasses import dataclass
from inspect import getfullargspec
from typing import Callable, List, Optional, Union

from .abstract_syntax_tree import Block
from .reference import FunctionReference
from .value import Value
from .type import Type


@dataclass
class Function(Value):
    """An abstract function, be it implemented in VBA or Python."""
    base_type = Type.Function

    name: str
    arguments_names: List[str]

    def __init__(self, name: str, arguments_names: List[str],
                 function: Union["ExternalFunction.Signature", Block]) -> None:
        self.name = name
        self.arguments_names = arguments_names
        self.value = function

    def convert_to_different_type(self, to_type: Type) -> Optional[Value]:
        return None


class ExternalFunction(Function):
    """
    Python function wrapper, the function will be called with two arguments:
      - a reference to the interpreter, allowing to have side effects
      - a list of Value objects, the actual arguments of the function
    It has to return a Value object.
    """
    Signature = Callable[["Interpreter", List[Value]], Value]

    def __init__(self, name: str, arguments_names: List[str],
                 external_function: "ExternalFunction.Signature") -> None:
        super().__init__(name, arguments_names, external_function)

    @staticmethod
    def from_function(
            python_function: Optional["ExternalFunction.Signature"] = None,
            name: Optional[str] = None) -> \
        Union["ExternalFunction.Signature",
              Callable[["ExternalFunction.Signature"], "ExternalFunction"]]:
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

    @property
    def external_function(self) -> "ExternalFunction.Signature":
        return self.value


class InternalFunction(Function):
    """VBA function."""
    reference: FunctionReference

    def __init__(self, name: str, arguments_names: List[str],
                 body: Block, reference: FunctionReference) -> None:
        super().__init__(name, arguments_names, body)
        self.reference = reference

    @property
    def body(self) -> Block:
        return self.value
