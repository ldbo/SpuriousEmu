"""The memory is used to store the values of variables and functions."""

from dataclasses import dataclass
from enum import Enum
from time import time
from typing import Any, Callable, Dict, List

from .function import Function
from .value import Value
from .vba_class import Class


class Memory:
    """
    Represent the memory used for running a program, including static elements
    like functions and dynamic ones like local variables.
    """
    global_variables: Dict[str, Value]
    _local_variables: List[Dict[str, Value]]
    _functions: Dict[str, Function]
    _classes: Dict[str, Class]

    def __init__(self) -> None:
        self.global_variables = dict()
        self._local_variables = []
        self._functions = dict()
        self._classes = dict()

    def add_function(self, full_name: str, function: Function) -> None:
        self._functions[full_name] = function

    def get_function(self, full_name: str) -> Function:
        return self._functions[full_name]

    def set_variable(self, local_name: str, value: Value) -> None:
        self._local_variables[-1][local_name] = value

    def get_variable(self, name: str) -> Value:
        # TODO use full name for variable storage
        # TODO use variable reference instead of name to ensure unicity
        local_name = name.split('.')[-1]
        try:
            return self.locals[local_name]
        except KeyError:
            return self.global_variables[name]

    @property
    def locals(self):
        return self._local_variables[-1]

    @property
    def functions(self):
        return self._functions

    @property
    def classes(self):
        return self._classes

    def new_locals(self) -> None:
        self._local_variables.append(dict())

    def discard_locals(self) -> None:
        self._local_variables.pop()


# TODO test hooks
class OutsideWorld:
    class EventType(Enum):
        STDOUT = "stdout"
        FILE = "file"
        NETWORK = "network"
        EXECUTION = "execution"

    @dataclass
    class Event:
        time: float
        data: Any

        def to_dict(self) -> Dict[str, Any]:
            return self.data

    __events: Dict["OutsideWorld.EventType", List[Any]]
    __hooks: Dict["OutsideWorld.EventType", Callable[[Any], None]]
    __start_time: float
    files: Dict[str, str]

    def __init__(self) -> None:
        self.__events = {event_type: []
                         for event_type in OutsideWorld.EventType}
        self.__hooks = {event_type: lambda t: None
                        for event_type in OutsideWorld.EventType}
        self.__start_time = time()
        self.files = dict()

    def add_event(self, event_type: "OutsideWorld.EventType", data: Any) \
            -> None:
        self.__hooks[event_type](data)
        t = time() - self.__start_time
        self.__events[event_type].append(OutsideWorld.Event(t, data))

        if event_type is OutsideWorld.EventType.FILE:
            if data['type'] == 'Write':
                content = self.files.get(data['path'], "")
                self.files[data['path']] = content + data['data']

    def to_dict(self) -> Dict[str, List[Any]]:
        return {event_type.value: [e.to_dict()
                                   for e in self.__events[event_type]]
                for event_type in OutsideWorld.EventType}

    def add_hook(self, event_type: "OutsideWorld.EventType",
                 hook: Callable[[Any], None]) -> None:
        self.__hooks[event_type] = hook
