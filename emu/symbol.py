from enum import Enum
from typing import Optional, List, Tuple


class Symbol:
    """
    A symbol extracted during compilation, representing a named programming
    element.
    """
    Type = Enum("Type", "Namespace Module Function Variable")

    _name: str
    _symbol_type: Type
    _parent: Optional["Symbol"]
    _children: List["Symbol"]
    _position: Tuple[int]
    __search_nodes: List["Symbol"]

    @staticmethod
    def build_root() -> "Symbol":
        """Build a root namespace symbol, with no parent."""
        name = "<Root>"
        symbol_type = Symbol.Type.Namespace
        parent = None
        children = []
        position = tuple()
        return Symbol(name, symbol_type, parent, children, position)

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
            self._children.append(child)
            return child
        else:
            msg = f"Trying to add the symbol {child._name} in {self}, which " \
                "already contains a symbol with the same name"
            raise CompilationError(msg)

    def add_child(self, name, child_type) -> "Symbol":
        """Initialize a child, add it and return it."""
        parent = None
        children = []
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
        own_child = self.find_child(name)
        if own_child is None:
            ret = []
        else:
            ret = [own_child]

        return sum((child.find(name) for child in self._children), ret)

    def resolve(self, name: str) -> Optional["Symbol"]:
        """Resolve a name in the recursive children of the symbol."""
        return self.__resolve(name.split('.'))

    def __resolve(self, name: List[str]) -> Optional["Symbol"]:
        if self._symbol_type == Symbol.Type.Function:
            if len(name) == 0:
                return self
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
        """Build the absolute name of the symbol, starting with <Root>."""
        if self.is_root():
            return self._name
        else:
            return self._parent.full_name() + '.' + self._name

    def __str__(self) -> str:
        return f"{self.full_name()}({self._symbol_type.name})"

    def __iter__(self):
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
        return name in map(lambda s: s._name, self._children)