"""
Provide abstract base classes to implement the Visitor design pattern.

A Visitable object has nothing more to do than inheriting Visitable. All the
algorithm must be only defined in the Visitor object. To handle a special
Visitable type SubVisitable, a Visitor must implement visit_SubVisitable, and
can then call visit(SubVisitable).
"""


from abc import abstractmethod
from typing import Any, Generic, TypeVar


class Visitable:
    """Base class for Visitable classes, must simply be inherited."""

    @abstractmethod
    def __init__(self):
        pass

    def accept(self, visitor: "Visitor") -> Any:
        assert isinstance(visitor, Visitor)

        visiter = "visit_" + type(self).__qualname__.replace(".", "_")
        try:
            visiter_function = getattr(visitor, visiter)
        except AttributeError:
            msg = (
                f"Visitor {type(visitor).__qualname__} doesn't handle "
                + f"Visitable type {type(self).__qualname__}"
            )
            raise NotImplementedError(msg)

        return visiter_function(self)


T = TypeVar("T")


class Visitor(Generic[T]):
    """
    Base class for Visitor classes. For a subclass to be able to handle the
    Visitable type SubVisitable, it must implement a method
    visit_SubVisitable(self, visitable : Visitable) -> T. Then, to apply the
    algorithm to a Visitable object visitable, simply call
    visitor.visit(visitable). Use the class argument T to specify the output
    type of the visit_ methods.
    """

    @abstractmethod
    def __init__(self):
        pass

    def visit(self, visitable: Visitable) -> T:
        """
        Apply an algorithm to a Visitable object of type T. The algorithm must
        be implemented in visit_T, otherwise a NotImplementedError is thrown.
        """
        assert isinstance(visitable, Visitable)

        return visitable.accept(self)
