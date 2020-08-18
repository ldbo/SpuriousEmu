"""Define classes allowing to store object values"""

# See type.py for a list of type still to be implemented
#
# To add a new supported value, you must:
#  - subclass Value
#  - add the Type/Value correspondance in TYPES_MAP

from abc import abstractmethod, ABC
from typing import Any, Dict, Optional, Union

from .error import ConversionError
from .reference import ClassModule
from .type import Type


class Value(ABC):
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
            converted = self.convert_to_different_type(to_type)
            if converted is not None:
                return converted
            else:
                msg = f"Can't convert {self.base_type} to {to_type}"
                raise ConversionError(msg)

    @abstractmethod
    def convert_to_different_type(self, to_type: Type) -> Optional["Value"]:
        """
        Function overridden by child classes, to define how to convert the
        value to different types. It should return None for non-supported
        types. Don't override convert_to !
        """
        pass

    @staticmethod
    def from_value(value, to_type: Type) -> "Value":
        """
        Use a couple value/type to build a Value subclass object.

        >>> str(Value.from_value(12, Type.Integer))
        'Integer(12)'

        It relies on the TYPES_MAP dictionnary.
        """
        try:
            return TYPES_MAP[to_type](value)
        except KeyError:
            raise NotImplementedError(f"Type {to_type} is not already "
                                      "implemented")

    @staticmethod
    def from_literal(literal) -> "Value":
        """
        Build the value corresponding to a literal, using Value.from_value.
        """
        return Value.from_value(literal.value, literal.type)

    def __str__(self) -> str:
        return f"{self.base_type.name}({self.value})"

    @staticmethod
    def from_python_base_type(value: Any) -> "Value":
        value_type = type(value)
        if isinstance(value, Value) or value is None:
            return value
        elif value_type is int:
            return Integer(value)
        elif value_type is bool:
            return Boolean(value)
        elif value_type is str:
            return String(value)
        else:
            msg = f"Can't create VBA value for Python type {value_type}"
            raise ConversionError(msg)


class Integer(Value):
    base_type = Type.Integer

    def __init__(self, integer: Union[str, int]) -> None:
        if isinstance(integer, int):
            self.value = integer
        else:
            self.value = int(integer, 10)

    def convert_to_different_type(self, to_type: Type) -> Optional[Value]:
        if to_type == Type.Boolean:
            return Boolean(self.value != 0)
        return None


class Boolean(Value):
    base_type = Type.Boolean

    def __init__(self, boolean: Union[str, bool]) -> None:
        if isinstance(boolean, bool):
            self.value = boolean
        else:
            self.value = boolean == "True"

    def convert_to_different_type(self, to_type: Type) -> Optional[Value]:
        if to_type == Type.Integer:
            return Integer(1) if self.value else Integer(0)
        return None


class String(Value):
    base_type = Type.String

    def __init__(self, string: str) -> None:
        self.value = string

    def convert_to_different_type(self, to_type: Type) -> Optional[Value]:
        # TODO
        return None


class Object(Value):
    base_type = Type.Object

    class_reference: ClassModule
    value: Dict[str, Value]  # Holds the state of the object

    def __init__(self, class_reference=ClassModule,
                 variables: Dict[str, Value] = None) -> None:
        self.class_reference = class_reference
        self.value = variables if variables is not None else dict()

    @property
    def variables(self) -> Dict[str, Value]:
        return self.value

    def convert_to_different_type(self, to_type: Type) -> Optional[Value]:
        # TODO
        return None


TYPES_MAP = {
    Type.Integer: Integer,
    Type.Boolean: Boolean,
    Type.String: String
}
