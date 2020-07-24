from enum import Enum
from typing import Optional, List, Tuple

from .error import CompilationError, ResolutionError


class Symbol:
    """
    A symbol extracted during compilation, representing a named programming
    element, and its hierarchical children.
    """
    Type = Enum("Type", "Namespace Module Function Variable")

    _name: str
    _symbol_type: Type
    _parent: Optional["Symbol"]
    _children: List["Symbol"]
    _position: Tuple[int]
    __search_nodes: List["Symbol"]
    __full_name: str

    @staticmethod
    def build_root() -> "Symbol":
        """Build a root namespace symbol, with no parent."""
        name = "<Root>"
        symbol_type = Symbol.Type.Namespace
        parent = None
        children: List['Symbol'] = []
        position: Tuple[int, ...] = tuple()
        root = Symbol(name, symbol_type, parent, children, position)
        root.__full_name = "<Root>"
        return root

    def __init__(self, name, symbol_type, parent=None, children=None,
                 position=None) -> None:
        """
        Shall not be called outside of Symbol or a child class. Use add_child
        instead.
        """
        self._name = name
        self._symbol_type = symbol_type
        self._parent = parent
        self._children = children
        self._position = position
        self.__full_name = ""

    @property
    def name(self):
        return self._name

    @property
    def symbol_type(self):
        return self._symbol_type

    @property
    def parent(self):
        return self._parent

    def __add_child(self, child) -> "Symbol":
        """Add an already initialized child."""
        if self.find_child(child._name) is None:
            child._parent = self
            child._position = self._position + (len(self._children), )
            child.__full_name = f"{self.__full_name}.{child._name}"
            self._children.append(child)
            return child
        else:
            msg = f"Trying to add the symbol {child._name} in {self}, which " \
                "already contains a symbol with the same name"
            raise CompilationError(msg)

    def add_child(self, name, child_type) -> "Symbol":
        """Initialize a child, add it and return it."""
        parent = None
        children: List['Symbol'] = []
        position = None
        child = Symbol(
            name, child_type, parent, children, position)
        return self.__add_child(child)

    def find_child(self, name: str) -> Optional["Symbol"]:
        """
        Search for a child with a given name. Return it or None if it is not
        found, and raise a CompilationError if there are two children with the
        same name.
        """

        results = list(
            filter(lambda child: child._name == name, self._children))
        if len(results) == 0:
            return None
        elif len(results) == 1:
            return results[0]
        else:
            msg = f"{self} has multiple children name {name}"
            raise CompilationError(msg)

    def find(self, name: str) -> List["Symbol"]:
        """
        Search for a direct or indirect child with the given name, using
        find_child recursively.
        """
        own_child = self.find_child(name)
        if own_child is None:
            ret = []
        else:
            ret = [own_child]

        return sum((child.find(name) for child in self._children), ret)

    def resolve(self, name: str) -> "Symbol":
        """Resolve a name in the recursive children of the symbol."""
        resolution = self.__resolve(name.split('.'))
        if resolution is None:
            msg = f"Can't resolve {name} from {self.full_name()}"
            raise ResolutionError(msg)
        else:
            return resolution

    def __resolve(self, name: List[str]) -> Optional["Symbol"]:
        if self._symbol_type == Symbol.Type.Function:
            if len(name) == 0:
                return self
            else:
                if self._parent is None:
                    return None
                else:
                    return self._parent.__resolve(name)
        else:
            child = self.find_child(name[0])
            if child is None:
                if self._parent is None:
                    return None
                else:
                    return self._parent.__resolve(name)
            else:
                return child.__resolve(name[1:])

    def is_root(self) -> bool:
        return self._parent is None

    def full_name(self) -> str:
        """Return the absolute name of the symbol, starting from the root"""
        return self.__full_name

    def __str__(self) -> str:
        return f"{self.full_name()}({self._symbol_type.name})"

    def __iter__(self):
        """
        Return an iterator to all the direct and indirect children of the
        symbol.
        """
        self.__search_nodes = [self]
        return self

    def __next__(self):
        try:
            last_node = self.__search_nodes.pop()
            self.__search_nodes.extend(last_node._children)
            return last_node
        except IndexError:
            raise StopIteration

    def __contains__(self, name):
        """Check if name is a direct child of the symbol."""
        return name in map(lambda s: s._name, self._children)
