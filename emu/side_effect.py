"""The memory is used to store the values of variables and functions."""

from dataclasses import dataclass
from enum import Enum
from itertools import count
from time import time
from typing import Any, Callable, Dict, List, Iterator

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
    """
    Represent the interaction between a running program and an emulated
    exterior environment.
    """
    @dataclass
    class Event:
        """
        Class used to record usefull events. It stores its category, the
        occuring time and a corresponding identifier, its context (the name of
        the symbol that triggered the event), and a data field, which depends
        on the event category.
        """
        class Category(Enum):
            STDOUT = "stdout"
            FILESYSTEM = "filesystem"
            NETWORK = "network"
            EXECUTION = "execution"
            ENVIRONMENT = "environment"

        identifier: int
        time: float
        category: "OutsideWorld.Event.Category"
        context: str
        data: Any

        def to_dict(self, reproducible: bool = False) -> Dict[str, Any]:
            """
            Return the dict of the event attributes.

            :arg reproducible: If True, don't add the time field
            :returns: The event attributes
            """
            d = {
                "identifier": self.identifier,
                "category": self.category.value,
                "context": self.context,
                "data": self.data,
            }

            if not reproducible:
                d["time"] = self.time

            return d

    __start_time: float
    __events_count: Iterator[int]
    __hooks: Dict["OutsideWorld.Event.Category",
                  Callable[["OutsideWorld.Event"], None]]
    events: List["OutsideWorld.Event"]
    files: Dict[str, str]

    def __init__(self) -> None:
        self.__start_time = time()
        self.__events_count = count()
        self.__hooks = {event_category: OutsideWorld.nop
                        for event_category in OutsideWorld.Event.Category}
        self.events = []
        self.files = dict()

    def add_event(self, category: "OutsideWorld.Event.Category", context: str,
                  data: Any) -> None:
        """
        Record a new event, calling the corresponding hook.
        """
        t = time() - self.__start_time
        identifier = next(self.__events_count)
        event = OutsideWorld.Event(identifier, t, category, context, data)

        self.events.append(event)
        self.__hooks[category](event)

        if category is OutsideWorld.Event.Category.FILESYSTEM:
            if data['type'] == 'Write':
                path = data['path']
                content = self.files.get(path, "")
                self.files[path] = content + data['data']

    def to_dict(self, reproducible: bool = False) -> Dict[str, Any]:
        """
        Return the list of events and files in a dict.

        :arg reproducible: If True, discard the time field of the events
        :returns: A dict with keys 'events' and 'files'
        """
        d: Dict[str, Any] = dict()
        d['events'] = [event.to_dict(reproducible)
                       for event in self.events]
        d['files'] = self.files

        return d

    def add_hook(self, category: "OutsideWorld.Event.Category",
                 hook: Callable[["OutsideWorld.Event"], None]) -> None:
        """
        Add a hook that will be called for each recording of a given event
        category.

        Hooks should not be lambda's for OutsideWorld to be used with
        Serializer.
        """
        self.__hooks[category] = hook

    @staticmethod
    def nop(event: "OutsideWorld.Event") -> None:
        """Function that can be used as a hook that does nothing."""
        pass
