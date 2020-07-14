"""Define classes allowing to store object values"""

# See type.py for a list of type still to be implemented
#
# To add a new supported value, you must:
#  - subclass Value
#  - add the Type/Value correspondance in TYPES_MAP

from abc import abstractmethod
from typing import Any, Union

from .abstract_syntax_tree import Literal
from .type import Type


class ConversionError(Exception):
    pass


class Value:
    """Represents the interpretation of an expression."""
    base_type: Type
    value: Any

    @abstractmethod
    def __init__(self):
        pass

    def convert_to(self, to_type: Type) -> "Value":
        """Used to convert values between types, for example to get the
        boolean interpretation of a non Boolean value."""
        if self.base_type == to_type:
            return self
        else:
            converted = self.__convert_to(to_type)
            if converted is not None:
                return converted
            else:
                msg = f"Can't convert {self.base_type} to {to_type}"
                raise ConversionError(msg)

    @abstractmethod
    def __convert_to(self, to_type: Type) -> "Value":
        """
        Function overridden by child classes, to define how to convert the
        value. Don't override convert_to !
        """
        pass

    @staticmethod
    def from_value(value, to_type):
        """
        Use a couple value/type to build a Value subclass object.

        >>> str(Value.from_value(12, Type.Integer))
        'Integer(12)'

        It relies on the TYPES_MAP dictionnary.
        """
        try:
            return TYPES_MAP[to_type](value)
        except KeyError:
            return NotImplementedError(f"Type {to_type} is not already "
                                       "implemented")

    @staticmethod
    def from_literal(literal):
        """
        Build the value corresponding to a literal, using Value.from_value.
        """
        return Value.from_value(literal.value, literal.type)

    def __str__(self):
        return f"{self.base_type.name}({self.value})"


class Integer(Value):
    base_type = Type.Integer

    def __init__(self, integer: Union[str, int]):
        if isinstance(integer, int):
            self.value = integer
        else:
            self.value = int(integer, 10)

    def __convert_to(self, to_type):
        if to_type == Type.Boolean:
            return Boolean(self.value != 0)


class Boolean(Value):
    base_type = Type.Boolean

    def __init__(self, boolean: Union[str, bool]):
        if isinstance(boolean, bool):
            self.value = boolean
        else:
            self.value = boolean == "True"

    def __convert_to(self, to_type):
        if to_type == Type.Integer:
            return Integer(1) if self.value else Integer(0)


class String(Value):
    base_type = Type.String

    def __init__(self, string: str):
        self.value = string

    def __convert_to(self, to_type):
        # TODO
        pass


TYPES_MAP = {
    Type.Integer: Integer,
    Type.Boolean: Boolean,
    Type.String: String
}
