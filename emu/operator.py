"""Defines the logic of built-in VBA operators"""

# To implement support for a new operator, you must:
#  - ensure the types it works with are implemented, see type.py
#  - add a new entry to OPERATORS_MAP, with the symbol of the operator as key,
#    and build the corresponding BinaryOperator as value.

from dataclasses import InitVar, dataclass
from typing import Callable, Any, List

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
            raise NotImplementedError(f"Op {symbol} not here lad")

    def __str__(self):
        return self.symbol


OPERATORS_MAP = {
    '^': BinaryOperator('^', [
        BinaryOperation(lambda l, r: l.value ** r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    '*': BinaryOperator('*', [
        BinaryOperation(lambda l, r: l.value * r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    '/': BinaryOperator('/', [
        BinaryOperation(lambda l, r: l.value / r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    '\\': BinaryOperator('\\', [
        BinaryOperation(lambda l, r: l.value // r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    'Mod': BinaryOperator('Mod', [
        BinaryOperation(lambda l, r: l.value % r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    "+": BinaryOperator('+', [
        BinaryOperation(lambda l, r: l.value + r.value,
                        Type.Integer, Type.Integer, Type.Integer),
        BinaryOperation(lambda l, r: l.value + r.value,
                        Type.String, Type.String, Type.String)
    ]),
    '-': BinaryOperator('-', [
        BinaryOperation(lambda l, r: l.value - r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    '&': BinaryOperator('&', [
        BinaryOperation(lambda l, r: l.value + r.value,
                        Type.String, Type.String, Type.String)
    ]),
    'And': BinaryOperator('And', [
        BinaryOperation(lambda l, r: l.value and r.value,
                        Type.Boolean, Type.Boolean, Type.Boolean)]),
    'Or': BinaryOperator('Or', [
        BinaryOperation(lambda l, r: l.value or r.value,
                        Type.Boolean, Type.Boolean, Type.Boolean)]),
    'Xor': BinaryOperator('Xor', [
        BinaryOperation(lambda l, r: l.value != r.value,
                        Type.Boolean, Type.Boolean, Type.Boolean)])
}
