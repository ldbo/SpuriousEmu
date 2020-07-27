"""The memory is used to store the values of variables and functions."""

from dataclasses import dataclass
from enum import Enum
from time import time
from typing import Any, Callable, Dict, List

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

    def discard_locals(self) -> None:
        self._local_variables.pop()


class OutsideWorld:
    class EventType(Enum):
        STDOUT = "stdout"
        FILE = "file"
        NETWORK = "network"

    @dataclass
    class Event:
        time: float
        data: Any

    __events: Dict["OutsideWorld.EventType", List[Any]]
    __hooks: Dict["OutsideWorld.EventType", Callable[[Any], None]]
    __start_time: float

    def __init__(self) -> None:
        self.__events = {event_type: []
                         for event_type in OutsideWorld.EventType}
        self.__hooks = {event_type: lambda t: None
                        for event_type in OutsideWorld.EventType}
        self.__start_time = time()

    def add_event(self, event_type: "OutsideWorld.EventType", data: Any) \
            -> None:
        self.__hooks[event_type](data)
        t = time() - self.__start_time
        self.__events[event_type].append(OutsideWorld.Event(t, data))

    def to_dict(self) -> Dict[str, List[Any]]:
        return {event_type.value: self.__events[event_type]
                for event_type in OutsideWorld.EventType}

    def add_hook(self, event_type: "OutsideWorld.EventType",
                 hook: Callable[[Any], None]) -> None:
        self.__hooks[event_type] = hook
