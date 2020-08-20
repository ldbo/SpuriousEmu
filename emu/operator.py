"""Defines the logic of built-in VBA operators"""

# To implement support for a new operator, you must:
#  - ensure the types it works with are implemented, see type.py
#  - add a new entry to OPERATORS_MAP
#  - add it to the list of operator tokens

# TODO:
#  - Take Let-coercion into account, improve conversion process
#  - Extends supported types
#  - Implement Like and Is

from abc import ABC, abstractmethod
from dataclasses import InitVar, dataclass
from inspect import signature
from typing import (Any, Callable, Dict, Generic, List, Optional, Tuple,
                    TypeVar, Union)
from typing import Type as pType

from pyparsing import Or, opAssoc, ParseExpression

from .error import OperatorError
from .type import Type
from .value import ConversionError, Value


class Operation(ABC):
    pass


@dataclass
class UnaryOperation(Operation):
    """Represent the underlying function of a unary operator"""
    function: InitVar[Callable[[Any], Any]]
    arg_type: Type
    return_type: Type

    # Workaround for mypy not supporting class-level hints for callbacks
    def __post_init__(self, function: Callable[[Any], Any]) -> None:
        self.function = function


@dataclass
class BinaryOperation(Operation):
    """Represent the underlying function of a binary operator"""
    function: InitVar[Callable[[Any, Any], Any]]
    left_type: Type
    right_type: Type
    return_type: Type

    # Workaround for mypy not supporting class-level hints for callbacks
    def __post_init__(self, function: Callable[[Any, Any], Any]) -> None:
        self.function = function


T = TypeVar('T', UnaryOperation, BinaryOperation)


class Operator(Generic[T], ABC):
    """
    Abstract operator, be it unary or binary.
    """
    symbol: str
    operations: List[T]

    @abstractmethod
    def __init__(self, symbol: str, operations: List[T]) -> None:
        self.symbol = symbol
        self.operations = operations

    @staticmethod
    def from_symbol(symbol: str) -> "Operator":
        """Build an operator from its symbol. It uses OPERATOR_MAP."""
        try:
            return OPERATORS_MAP[symbol]
        except KeyError:
            msg = f"Operator {symbol} is not supported yet"
            raise NotImplementedError(msg)

    def __str__(self) -> str:
        return self.symbol

    @staticmethod
    def from_operation(symbol: str, operation: Operation) -> "Operator":
        """Build an operator from a symbol and a single."""
        operator_type: Union[pType[BinaryOperator], pType[UnaryOperator]]
        if isinstance(operation, UnaryOperation):
            return UnaryOperator(symbol, [operation])
        elif isinstance(operation, BinaryOperation):
            return BinaryOperator(symbol, [operation])

        msg = f"Can't build Operator from type {type(operation)}"
        raise RuntimeError(msg)


class UnaryOperator(Operator[UnaryOperation]):
    """
    Class used to evaluate unary operations with operators that may accept
    different types.
    """

    def __init__(self, symbol: str, operations: List[UnaryOperation]) -> None:
        super().__init__(symbol, operations)

    def operate(self, value: Value) -> Value:
        """
        Find the first UnaryOperation in self.operations that can be called
        with the given value, and return its result. Raise a ConversionError if
        there is no compatible operator.
        """
        for operation in self.operations:
            try:
                converted_arg = value.convert_to(operation.arg_type)
                value = operation.function(converted_arg)
                value_type = operation.return_type
                return Value.from_value(value, value_type)
            except ConversionError:
                pass

        msg = f"Type {value.base_type} does not match with operator " \
              f"{self.symbol}"
        raise OperatorError(msg)


class BinaryOperator(Operator[BinaryOperation]):
    """
    Class used to evaluate binary operations with operators that may accept
    different types.
    """

    def __init__(self, symbol: str, operations: List[BinaryOperation]) -> None:
        super().__init__(symbol, operations)

    def operate(self, left_value: Value, right_value: Value) -> Value:
        """
        Find the first BinaryOperation in self.operations that can be called
        with the two given values, and return its result. Raise a
        ConversionError if there is no compatible operator.
        """
        for operation in self.operations:
            try:
                converted_left = left_value.convert_to(operation.left_type)
                converted_right = right_value.convert_to(operation.right_type)
                value = operation.function(converted_left, converted_right)
                value_type = operation.return_type
                return Value.from_value(value, value_type)
            except ConversionError:
                pass

        msg = f"Types {left_value.base_type}, {right_value.base_type} do " \
              f"not match with operator {self.symbol}"
        raise OperatorError(msg)


class OperatorsMap:
    """
    Main class used to add operators to the parser. A single global instance is
    used, OPERATORS_MAP.

    To implement support for a new operator, use add_operation or the <<
    operator. Precedence is managed by start_precedence_group and
    end_precedence_group. Finally, you can use get_precedence_list with
    pyparsing.infixNotation.
    """
    __operators: Dict[str, Operator]
    __ordered_operators: List[List[str]]
    __growing_precedence: bool

    def __init__(self) -> None:
        self.__operators = dict()
        self.__ordered_operators = []
        self.__growing_precedence = True

    def start_precedence_group(self) -> None:
        """
        Start a precedence group: all the following added operations will have
        the same precedence.
        """
        self.__ordered_operators.append([])
        self.__growing_precedence = False

    def end_precedence_group(self) -> None:
        """
        Stop the current precedence group: all the added operations will have
        increasing precedence.
        """
        self.__growing_precedence = True

    def get_precedence_list(self, parse_unary, parse_binary):
        """Build a precedence list to be used with pyparsing.infixNotation."""
        precedence_list = []
        for symbol_list in self.__ordered_operators:
            expression = Or(symbol_list)
            if type(self.__operators[symbol_list[0]]) is UnaryOperator:
                arity = 1
                parsing_function = parse_unary
                associativity = opAssoc.RIGHT
            else:
                arity = 2
                parsing_function = parse_binary
                associativity = opAssoc.LEFT

            entry = (expression, arity, associativity, parsing_function)
            precedence_list.append(entry)

        return precedence_list

    def __insert_operation(self, symbol: str, operation: Operation) -> None:
        """
        Insert an operation associated with a given symbol, updating
        __operators and __ordered_operators
        """
        if symbol in self.__operators:
            assert(isinstance(operation,
                              type(self.__operators[symbol].operations[-1])))
            self.__operators[symbol].operations.append(operation)
        else:
            if self.__growing_precedence:
                self.__ordered_operators.append([symbol])
            else:
                self.__ordered_operators[-1].append(symbol)

            operator = Operator.from_operation(symbol, operation)
            self.__operators[symbol] = operator

    def add_unary_operation(self, symbol: str, function: Callable[[Any], Any],
                            arg_type: Type, rtype: Optional[Type] = None) \
            -> None:
        """
        Add a unary operation to the map. Be carefull of the call order of this
        method: operators are added with decreasing precedence.
        """
        assert(len(signature(function).parameters) == 1)

        if rtype is None:
            rtype = arg_type

        operation = UnaryOperation(function, arg_type, rtype)
        self.__insert_operation(symbol, operation)

    def add_binary_operation(self, symbol: str,
                             function: Callable[[Any, Any], Any],
                             ltype: Type, rtype: Optional[Type] = None,
                             return_type: Optional[Type] = None) -> None:
        """
        Add a binary operation to the map. Be carefull of the call order of
        this method: operators are added with decreasing precedence.
        """
        assert(len(signature(function).parameters) == 2)

        if return_type is None:
            return_type = ltype

        if rtype is None:
            rtype = ltype

        operation = BinaryOperation(function, ltype, rtype, return_type)
        self.__insert_operation(symbol, operation)

    def __getitem__(self, symbol: str) -> Operator:
        """Return the BinaryOperator corresponding to the given symbol."""
        return self.__operators[symbol]

    def __lshift__(self, operation_tuple: Tuple[Any, ...]) -> None:
        """
        Helper function to call add_operation. Pass it a tuple with the same
        arguments you would pass to add_operation, in the same order. There
        is one exception : you can specify only the return type and left
        argument type of a binary operator, with (symbol, function, left,
        return).
        """
        assert(len(operation_tuple) >= 3)
        symbol = operation_tuple[0]
        function = operation_tuple[1]
        ltype = operation_tuple[2]

        arity = len(signature(function).parameters)
        assert(arity in (1, 2))

        if arity == 1:
            assert(len(operation_tuple) in (3, 4))
            self.add_unary_operation(*operation_tuple)
        elif arity == 2:
            assert(len(operation_tuple) in (3, 4, 5))

            if len(operation_tuple) == 4:
                return_type = operation_tuple[3]
                self.add_binary_operation(symbol, function, ltype,
                                          return_type=return_type)
            else:
                self.add_binary_operation(*operation_tuple)


OPERATORS_MAP = OperatorsMap()

# Arithmetic operators
OPERATORS_MAP << ('^', lambda l, r: l.value ** r.value, Type.Integer)

OPERATORS_MAP.start_precedence_group()
OPERATORS_MAP << ('*', lambda l, r: l.value * r.value, Type.Integer)
OPERATORS_MAP << ('/', lambda l, r: l.value / r.value, Type.Integer)
OPERATORS_MAP.end_precedence_group()

OPERATORS_MAP << ('\\', lambda l, r: l.value // r.value, Type.Integer)
OPERATORS_MAP << ('Mod', lambda l, r: l.value % r.value, Type.Integer)

OPERATORS_MAP.start_precedence_group()
OPERATORS_MAP << ('+', lambda l, r: l.value + r.value, Type.Integer)
OPERATORS_MAP << ('+', lambda l, r: l.value + r.value, Type.String)
OPERATORS_MAP << ('-', lambda l, r: l.value - r.value, Type.Integer)
OPERATORS_MAP.end_precedence_group()

# Concatenation operator
OPERATORS_MAP << ('&', lambda l, r: l.value + r.value, Type.String)

# Relational operators
OPERATORS_MAP.start_precedence_group()
OPERATORS_MAP << ('=', lambda l, r: l.value == r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP << ('=', lambda l, r: l.value == r.value,
                  Type.String, Type.Boolean)
OPERATORS_MAP << ('<>', lambda l, r: l.value != r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP << ('<>', lambda l, r: l.value != r.value,
                  Type.String, Type.Boolean)
OPERATORS_MAP << ('><', lambda l, r: l.value != r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP << ('><', lambda l, r: l.value != r.value,
                  Type.String, Type.Boolean)
OPERATORS_MAP << ('<', lambda l, r: l.value < r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP << ('>', lambda l, r: l.value > r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP << ('<=', lambda l, r: l.value <= r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP << ('>=', lambda l, r: l.value >= r.value,
                  Type.Integer, Type.Boolean)
OPERATORS_MAP.end_precedence_group()


# Logical and bitwise operators
OPERATORS_MAP << ('Not', lambda a: not a.value, Type.Boolean)
OPERATORS_MAP << ('And', lambda l, r: l.value and r.value, Type.Boolean)
OPERATORS_MAP << ('Or', lambda l, r: l.value or r.value, Type.Boolean)
OPERATORS_MAP << ('Xor', lambda l, r: l.value != r.value, Type.Boolean)
OPERATORS_MAP << ('Eqv', lambda l, r: l.value == r.value, Type.Boolean)
OPERATORS_MAP << ('Imp', lambda l, r: (not l.value) or r.value, Type.Boolean)
