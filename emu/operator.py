"""Defines the logic of built-in VBA operators"""

# To implement support for a new operator, you must:
#  - ensure the types it works with are implemented, see type.py
#  - add a new entry to OPERATORS_MAP
#  - add it to the list of operator tokens

# TODO:
#  - Take Let-coercion into account, improve conversion process
#  - Extends supported types
#  - Implement Like and Is

from dataclasses import InitVar, dataclass
from inspect import signature
from typing import Callable, Any, List, Dict, Optional, Tuple

from pyparsing import Or, opAssoc

from .error import OperatorError
from .type import Type
from .value import ConversionError, Value


@dataclass
class BinaryOperation:
    function: InitVar[Callable[[Any, Any], Any]]
    left_type: Type
    right_type: Type
    return_type: Type

    # Workaround for mypy not supporting class-level hints for callbacks
    def __post_init__(self, function: Callable[[Any, Any], Any]) -> None:
        self.function = function


class BinaryOperator:
    """
    Class used to evaluate binary operations with operators that may accept
    different types.
    """
    operations: List[BinaryOperation]

    def __init__(self, symbol: str, operations: List[BinaryOperation]):
        self.symbol = symbol
        self.operations = operations

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

        error_message = "Types {left_value.type}, {right_value.type} do not" \
            " match with operator {self.symbol}"
        raise OperatorError(error_message)

    @staticmethod
    def build_operator(symbol: str) -> "BinaryOperator":
        """Build an operator from its symbol. It uses OPERATOR_MAP."""
        try:
            return OPERATORS_MAP[symbol]
        except KeyError:
            msg = f"Operator {symbol} is not supported yet"
            raise NotImplementedError(msg)

    def __str__(self):
        return self.symbol


class OperatorsMap:
    """
    Main class used to add operators to the parser. A single global instance is
    used, OPERATORS_MAP.

    To implement support for a new operator, use add_operation or the <<
    operator. Precedence is managed by start_precedence_group and
    end_precedence_group. Finally, you can use get_precedence_list with
    pyparsing.infixNotation.
    """
    __operators: Dict[str, BinaryOperator]
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

    def get_precedence_list(self, parsing_function):
        """Build a precedence list to be used with pyparsing.infixNotation."""
        precedence_list = []
        for symbol_list in self.__ordered_operators:
            expression = Or(symbol_list)
            entry = (expression, 2, opAssoc.LEFT, parsing_function)
            precedence_list.append(entry)

        return precedence_list

    def add_operation(self, symbol: str, function: Callable[[Any, Any], Any],
                      ltype: Type, rtype: Optional[Type] = None,
                      return_type: Optional[Type] = None) -> None:
        """
        Add an operation to the map. Be carefull of the call order of this
        method: operators are added with decreasing precedence.
        """
        # Extract the signature of the function
        op_signature = signature(function)
        parameters_number = len(op_signature.parameters)
        assert(parameters_number in (1, 2))

        if return_type is None:
            return_type = ltype

        # Update __ordered_operators
        if symbol not in self.__operators:
            if self.__growing_precedence:
                self.__ordered_operators.append([symbol])
            else:
                self.__ordered_operators[-1].append(symbol)

        # Update __operators
        if parameters_number == 1:
            raise NotImplemented("Una... what operators ?")
        elif parameters_number == 2:
            if rtype is None:
                rtype = ltype

            operation = BinaryOperation(function, ltype, rtype, return_type)
            if symbol not in self.__operators:
                self.__operators[symbol] = BinaryOperator(symbol, [operation])
            else:
                self.__operators[symbol].operations.append(operation)

    def __getitem__(self, symbol: str) -> BinaryOperator:
        """Return the BinaryOperator corresponding to the given symbol."""
        return self.__operators[symbol]

    def __lshift__(self, operation_tuple: Tuple[Any, ...]) -> None:
        """
        Helper function to call add_operation. Pass it a tuple with the same
        arguments you would pass to add_operation, in the same order.
        """
        assert(len(operation_tuple) in (3, 4, 5))

        symbol = operation_tuple[0]
        function = operation_tuple[1]
        ltype = operation_tuple[2]

        if len(operation_tuple) == 3:
            self.add_operation(symbol, function, ltype)
        elif len(operation_tuple) == 4:
            return_type = operation_tuple[3]
            self.add_operation(symbol, function, ltype,
                               return_type=return_type)
        elif len(operation_tuple) == 5:
            rtype = operation_tuple[3]
            return_type = operation_tuple[4]
            self.add_operation(symbol, function, ltype, rtype, return_type)


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
OPERATORS_MAP << ('And', lambda l, r: l.value and r.value, Type.Boolean)
OPERATORS_MAP << ('Or', lambda l, r: l.value or r.value, Type.Boolean)
OPERATORS_MAP << ('Xor', lambda l, r: l.value != r.value, Type.Boolean)
OPERATORS_MAP << ('Eqv', lambda l, r: l.value == r.value, Type.Boolean)
OPERATORS_MAP << ('Imp', lambda l, r: (not l.value) or r.value, Type.Boolean)
