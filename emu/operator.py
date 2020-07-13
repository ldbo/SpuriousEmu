from dataclasses import dataclass
from typing import Callable, Any, List

from .type import Type
from .value import ConversionError, Value


class OperatorError(Exception):
    pass


@dataclass
class BinaryOperation:
    function: Callable[[Any], Any]
    left_type: Type
    right_type: Type
    return_type: Type


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
    def build_operator(symbol: str) : "BinaryOperator":
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
                        Type.Integer, Type.Integer, Type.Integer)
    ]),
    '-': BinaryOperator('-', [
        BinaryOperation(lambda l, r: l.value - r.value,
                        Type.Integer, Type.Integer, Type.Integer)
    ])
}
