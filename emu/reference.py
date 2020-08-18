"""
Implement the different kind of Reference classes produced by the compilation
"""

from abc import ABC, abstractmethod
from enum import Enum
from typing import Any, Dict, List, Optional

from .error import ResolutionError, CompilationError


class Reference(ABC):
    """
    Semantic representation of an identifier, either structural (for
    namespaces) or computational (for variables and functions). The references
    of a program are seen as a tree denoting its organization, with variables
    as the only kind of leaf.
    """
    Category = Enum("Category", "Structural Computational")

    category: "Reference.Category"

    name: str
    parent: Optional["Reference"]
    children: List["Reference"]
    __full_name: str

    @abstractmethod
    def __init__(self, name: str, parent: "Reference" = None,
                 children: List["Reference"] = None) -> None:
        self.name = name
        self.parent = parent
        self.children = children if children is not None else []
        self.__full_name = self.name

    def get_child(self, name: str) -> "Reference":
        """
        Get a child from the current reference, by its name.

        :throws ResolutionError: If there is no child with the given name
        """
        ret = list(filter(lambda child: child.name == name, self.children))
        assert(len(ret) <= 1)

        if len(ret) == 1:
            return ret[0]
        else:
            msg = f"Reference {self} doesn't have child called {name}"
            raise ResolutionError(msg)

    def add_child(self, child: "Reference") -> None:
        """
        Add a child to the current reference, updating its parent field.

        :throws CompilationError: If there is already a child with the same
        name
        """
        try:
            self.get_child(child.name)
            msg = f"Reference {self} already has a child called {child.name}"
            raise CompilationError(msg)
        except ResolutionError:
            self.children.append(child)
            child.parent = self
            child.__full_name = f"{self.__full_name}.{child.name}"

    def build_child(self, child_type, *args, **kwargs) -> "Reference":
        """
        Build a reference, add it as a child and return it.

        :param child_type: Type of the child to create
        :param args: Arguments to pass to the child constructor
        :param kwargs: Keyword arguments to pass to the child constructor
        :returns: The built child
        """
        child = child_type(*args, **kwargs)
        self.add_child(child)
        return child

    def search_child(self, name: str,
                     exclude_child: Optional["Reference"] = None) \
            -> "Reference":
        try:
            return self.get_child(name)
        except ResolutionError:
            for child in self.children:
                if child is exclude_child:
                    pass

                try:
                    return child.search_child(name)
                except ResolutionError:
                    pass

        msg = f"Can't find a reference named {name} in {self} children"
        raise ResolutionError(msg)

    def to_dict(self) -> Dict[str, Any]:
        d = {'name': self.name}
        for child in self.children:
            child_type = type(child).__name__
            similar_childs = d.get(child_type, [])
            d[child_type] = similar_childs + [child.to_dict()]

        return d

    def __str__(self) -> str:
        """Return the absolute name of the reference"""
        return self.__full_name


class Environment(Reference):
    category = Reference.Category.Structural

    DEFAULT_NAME = "VBAEnv"

    def __init__(self, name: Optional[str] = None) -> None:
        if name is None:
            name = Environment.DEFAULT_NAME

        super().__init__(name=name)


class Project(Reference):
    category = Reference.Category.Structural

    def __init__(self, name: str) -> None:
        super().__init__(name=name)


class Module(Reference):
    category = Reference.Category.Structural

    @abstractmethod
    def __init__(self, name: str) -> None:
        super().__init__(name=name)

    @property
    def functions(self):
        return list(filter(lambda child: type(child) is FunctionReference,
                           self.children))

    @property
    def variables(self):
        return list(filter(lambda child: type(child) is Variable,
                           self.children))

    def get_function(self, name: str) -> "FunctionReference":
        function = self.get_child(name)
        assert(type(function) is FunctionReference)
        return function

    def get_variable(self, name: str) -> "Variable":
        variable = self.get_child(name)
        assert(type(variable) is Variable)
        return variable


class ClassModule(Module):
    category = Reference.Category.Structural

    def __init__(self, name: str) -> None:
        super().__init__(name=name)


class ProceduralModule(Module):
    category = Reference.Category.Structural

    def __init__(self, name: str) -> None:
        super().__init__(name=name)


class Variable(Reference):
    class Extent(Enum):
        Program = 1
        Module = 2
        Procedure = 3
        Object = 4
        Aggregate = 5

    category = Reference.Category.Computational

    extent: "Variable.Extent"

    def __init__(self, name: str, extent: "Variable.Extent") -> None:
        super().__init__(name=name)
        self.extent = extent

    def add_child(self, child: Reference) -> None:
        msg = f"Can't add a child to Variable {self}"
        raise CompilationError(msg)

    def get_child(self, name: str) -> Reference:
        msg = f"Variable reference {self} can't have a child"
        raise ResolutionError(msg)


class FunctionReference(Reference):
    category = Reference.Category.Computational

    def __init__(self, name: str) -> None:
        super().__init__(name=name)
