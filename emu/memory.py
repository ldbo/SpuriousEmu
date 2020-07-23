"""The memory is used to store the values of variables and functions."""

from dataclasses import dataclass
from typing import List, Dict

from .function import Function
from .value import Value


@dataclass
class Variable:
    name: str
    value: Value


class Memory:
    """
    Represent the memory used for running a program, including static elements
    like functions and dynamic ones like local variables.
    """
    _local_variables: List[Dict[str, Value]]
    _functions: Dict[str, Function]

    def __init__(self) -> None:
        self._local_variables = []
        self._functions = dict()

    def add_function(self, full_name: str, function: Function) -> None:
        self._functions[full_name] = function

    def get_function(self, full_name: str) -> Function:
        return self._functions[full_name]

    def set_variable(self, local_name: str, value: Value) -> None:
        self._local_variables[-1][local_name] = value

    def get_variable(self, local_name: str) -> Value:
        return self._local_variables[-1][local_name]

    def new_locals(self) -> None:
        self._local_variables.append(dict())
