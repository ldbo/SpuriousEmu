"""
This module defines the VBA computational environment elements, conforming
to [MS-VBA]_ section 2.
"""

from dataclasses import InitVar, dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Type, Union

from .utils import MissingOptional

# Data value ([MS-VBA] section 2.1)


@dataclass(frozen=True)
class Value:
    """
    Base class of internal values. Child classes all have default constructor
    conforming to the initial data value of the corresponding declared data
    type.

    Args:
      value: Contains the internal representation of the value. Its type is
             specified in child classes.

    .. note:: Child classes are :class:`Boolean`, :class:`Byte`,
              :class:`Currency`, :class:`Date`, :class:`Decimal`,
              :class:`Double`, :class:`Integer`, :class:`Long`,
              :class:`LongLong`, :class:`ObjectReference`, :class:`Single`,
              :class:`String`, :class:`Empty`, :class:`Error`, :class:`Missing`,
              :class:`Array` and :class:`UDT`.
    """

    value: Any

    @staticmethod
    def from_declared_type(declared_type: "DeclaredType") -> "Value":
        """
        Returns:
            The default value corresponding to a given declared type

        Args:
          declared_type (:attr:`DeclaredType`): Either a :class:`Value` subclass
                                                or a :class:`PureDeclaredType`
                                                instance

        Raises:
          :exc:`TypeError`: If ``declared_type`` is not a supported declared
                            type
        """
        if isinstance(declared_type, PureDeclaredType):
            if isinstance(declared_type, Variant):
                return Empty()
            elif isinstance(declared_type, StringN):
                return String.from_stringn(declared_type)
            elif isinstance(declared_type, UDTName):
                return UDT.from_udt_name(declared_type)
            elif isinstance(declared_type, (FArray, FVArray)):
                if isinstance(declared_type, FArray):
                    element_type = declared_type.element_type
                else:
                    element_type = Variant()
                bounds = declared_type.bounds
                return Array.from_bounds(bounds, element_type)
            else:  # Resizeable array
                return Array()
        elif issubclass(declared_type, Value):
            try:
                return declared_type()  # type: ignore [call-arg]
            except TypeError:
                msg = f"{declared_type} is not a supported concrete data value"
                raise TypeError(msg)
        else:
            msg = f"{declared_type} is not a supported declared type"
            raise TypeError(msg)


@dataclass(frozen=True)
class Boolean(Value):
    value: bool = False


@dataclass(frozen=True)
class Byte(Value):
    value: int = 0

    def __post_init__(self) -> None:
        if not 0 <= self.value <= 255:
            msg = "Byte value must be in [0; 255]"
            raise ValueError(msg)


@dataclass(frozen=True)
class Currency(Value):
    value: int = 0

    def __post_init__(self) -> None:
        if not -9223372036854775808 <= self.value <= 9223372036854775807:
            msg = (
                "Currency value must be in "
                "[922,337,203,685,477.5808; 922,337,203,685,477.5808]"
            )
            raise ValueError(msg)


@dataclass(frozen=True)
class Date(Value):
    value: float = 0.0


@dataclass(frozen=True)
class Decimal(Value):
    value: float = 0.0


@dataclass(frozen=True)
class Double(Value):
    value: float = 0.0


@dataclass(frozen=True)
class Integer(Value):
    value: int = 0

    def __post_init__(self) -> None:
        if not -32768 <= self.value <= 32767:
            msg = "Integer value must be in [-32,768; 32,767]"
            raise ValueError(msg)


@dataclass(frozen=True)
class Long(Value):
    value: int = 0

    def __post_init__(self) -> None:
        if not -2147483648 <= self.value <= 2147483647:
            msg = "Long value must be in [-2,147,483,648; 2,147,483,647]"
            raise ValueError(msg)


@dataclass(frozen=True)
class LongLong(Value):
    value: int = 0

    def __post_init__(self) -> None:
        if not -9223372036854775808 <= self.value <= 9223372036854775807:
            msg = (
                "LongLong value must be in"
                "[-9,223,372,036,854,775,808; 9,223,372,036,854,775,807]"
            )
            raise ValueError(msg)


LongPtr = LongLong


@dataclass(frozen=True)
class ObjectReference(Value):
    value: Optional["Object"] = None  # Empty means nothing


@dataclass(frozen=True)
class Single(Value):
    value: float = 0.0


@dataclass(frozen=True)
class String(Value):
    value: bytes = b""

    @staticmethod
    def from_str(string: str) -> "String":
        return String(string.encode("utf-16"))

    @staticmethod
    def from_stringn(stringn: "StringN") -> "String":
        assert isinstance(stringn.n, int)
        return String(b"\x00\x00" * stringn.n)


@dataclass(frozen=True)
class Empty(Value):
    # Check if no one can access implementation specific byte pattern from VBA
    value: None = field(init=False, default=None)


@dataclass(frozen=True)
class Error(Value):
    value: int

    def __post_init__(self) -> None:
        if not 0 <= self.value <= 4294967295:
            msg = "Error value must be in [0; 4294967295]"
            raise ValueError(msg)


@dataclass(frozen=True)
class Null(Value):
    # Check if no one can access implementation specific byte pattern
    value: None = field(init=False, default=None)


@dataclass(frozen=True)
class Missing(Value):
    # Check if no one can access implementation specific byte pattern
    value: None = field(init=False, default=None)


Bounds = Tuple[Tuple[int, int], ...]


@dataclass(frozen=True)
class Array(Value):
    """
    Multi-dimensional heterogeneous array

    Attributes:
      value: Multi-dimensional list of :class:`Value`, must be a hypercube
      bounds: Lower and upper bounds of the different dimensions, must be
               sorted.

    Raises:
      :py:exc:`ValueError`: If some indices are not ordered, the ``value`` is
                            not a hypercube, or if the length of ``bounds``
                            does not match the dimensions of ``value``.
    """

    value: List[Any] = field(default_factory=list)
    bounds: Bounds = field(default_factory=tuple)

    def __post_init__(self) -> None:
        current_array = self.value

        if current_array == [] and self.bounds == ():
            return

        for lower_index, higher_index in self.bounds:
            size = higher_index - lower_index + 1
            if size <= 0:
                raise ValueError("Array dimensions must be non-negative")

            if size != len(current_array):
                raise ValueError("Dimension mismatch")

            if not isinstance(current_array, list):
                msg = "Array value must be a multi-dimensional list"
                raise ValueError(msg)

            new_array = current_array[0]
            new_list = type(new_array) == list

            if (
                any(  # Any has a type mismatch (e.g. list and Variant)
                    map(
                        lambda dim: (type(dim) == list) != new_list,
                        current_array,
                    )
                )
                or new_list
                and any(  # Two have different length
                    map(lambda dim: len(dim) != len(new_array), current_array)
                )
            ):
                msg = "The array value must be an hypercube enclosed in its "
                msg += "bounds"
                raise ValueError(msg)

            current_array = new_array

        if isinstance(current_array, list):
            msg = "An array must have bounds for each of its dimensions"
            raise ValueError(msg)

    def __getitem__(self, indices: List[int]) -> Value:
        """
        Returns:
          A value from the multi-dimensional array

        Args:
          indices: The multi-index of the value

        Raises:
          :exc:`IndexError`: If the multi-index has not the same dimension as
                             the array, or is outside of its bounds
        """
        if len(indices) != len(self.bounds):
            msg = "The multi-index must match with the array dimension"
            raise IndexError(msg)

        array = self.value
        for index, bounds in zip(indices, self.bounds):
            if bounds[0] <= index <= bounds[1]:
                array = array[index - bounds[0]]
            else:
                raise IndexError("Array index must be in bounds")

        return array  # type: ignore [return-value]

    @staticmethod
    def from_bounds(bounds: Bounds, declared_type: "DeclaredType") -> "Array":
        """
        Returns:
          An :class:`Array` filled with default-initialized values

        Args:
          bounds: Bounds of the array
          declared_type: Type of the array elements
        """
        if len(bounds) == 0:
            return Array([], ())

        def helper(index: int = 0) -> List[Any]:
            new_bounds = bounds[index]
            new_size = new_bounds[1] - new_bounds[0] + 1

            nd: List[Any]
            if index == len(bounds) - 1:
                nd = [
                    Value.from_declared_type(declared_type)
                    for _ in range(new_size)
                ]
            else:
                nd = [helper(index + 1) for _ in range(new_size)]

            return nd

        return Array(helper(), bounds)


@dataclass(frozen=True)
class UDT(Value):
    value: Dict[str, Value] = field(default_factory=dict)
    name: Optional[str] = None  # Not sure if is needed

    @staticmethod
    def from_udt_name(udt_name: "UDTName") -> "UDT":
        name = udt_name.name
        value = {
            field: Value.from_declared_type(element_type)
            for field, element_type in udt_name.element_types.items()
        }

        return UDT(value, name)


@dataclass(frozen=True)
class Object(Value):
    # TODO: implement
    pass


# Declared types ([MS-VBA] section 2.2)


@dataclass(frozen=True)
class PureDeclaredType:
    """
    Abstract class representing a declared type that is not directly a value
    type (Variant for example).

    Warning:
      PureDeclaredType `instances` must be used for value typing, not the class
      itself !

    Note:
      See [MS-VBA]_ section 2.2
    """

    pass


@dataclass(frozen=True)
class Variant(PureDeclaredType):
    pass


@dataclass(frozen=True)
class StringN(PureDeclaredType):
    n: Union[int, str]

    def __post_init__(self) -> None:
        if isinstance(self.n, int) and not 1 <= self.n <= 65535:
            raise ValueError("String length must be between 1 and 65535")


@dataclass(frozen=True)
class ClassName(PureDeclaredType):
    name: str = "Nothing"


@dataclass(frozen=True)
class UDTName(PureDeclaredType):
    element_types: Dict[str, "DeclaredType"] = field(default_factory=dict)
    name: Optional[str] = None  # TODO Not sure if needed


@dataclass(frozen=True)
class FArray(PureDeclaredType):
    """Fixed-size multi-dimensional array"""

    element_type: Union[PureDeclaredType, StringN, ClassName, UDTName]
    bounds: Bounds

    def __post_init__(self) -> None:
        if len(self.bounds) <= 60:
            if all(dim[0] <= dim[1] for dim in self.bounds):
                return

            raise ValueError("An array's bound should must be ordered")

        raise ValueError("An array can't have more than 60 dimensions")


@dataclass(frozen=True)
class FVArray(PureDeclaredType):
    """Fixed-size multi-dimensional Variant array"""

    bounds: Bounds

    def __post_init__(self) -> None:
        if len(self.bounds) <= 60:
            if all(dim[0] <= dim[1] for dim in self.bounds):
                return

            raise ValueError("An array's bound should must be ordered")

        raise ValueError("An array can't have more than 60 dimensions")


@dataclass(frozen=True)
class RArray(PureDeclaredType):
    """Resizeable multi-dimensional array"""

    element_type: Union[PureDeclaredType, StringN, ClassName, UDTName]


@dataclass(frozen=True)
class RVArray(PureDeclaredType):
    """Resizeable multi-dimensional Variant array"""

    pass


DeclaredType = Union[Type[Value], PureDeclaredType]
"""
Declared type of a variable, used for typing: either one of the :class:`Value`
child classes, or a :class:`PureDeclaredType` object.
"""


# Variable ([MS-VBA] section 2.3)


@dataclass
class Variable:
    """
    Interface to manipulating VBA values.

    Parameters:
      declared_type (:attr:`DeclaredType`): Declared type
      init_value (Optional[:class:`Value`]): Optional initialization value. If
                                             ``None``, the variable uses the
                                             default value of its declared type.
    """

    declared_type: DeclaredType  #: Declared type of the variable
    value: Value = field(init=False)
    """Internal immutable value of the variable"""
    init_value: InitVar[Optional[Value]] = MissingOptional

    def __post_init__(self, optional_value: Optional[Value]) -> None:
        if optional_value is MissingOptional:
            self.value = Value.from_declared_type(self.declared_type)
        else:
            self.value = optional_value
