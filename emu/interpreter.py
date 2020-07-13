from abc import abstractmethod

from .type import Type
from .value import Value
from .abstract_syntax_tree import *

class Interpreter:
    def evaluate_expression(self, expression):
        if isinstance(expression, Literal):
            return Value.from_literal(expression)
        elif isinstance(expression, Identifier):
            raise NotImplementedError("Don't use variables, you fool")
        elif isinstance(expression, FunCall):
            raise NotImplementedError("Don't call functions, you dumb ass")
        elif isinstance(expression, BinOp):
            left = self.evaluate_expression(expression.left)
            right = self.evaluate_expression(expression.right)
            op = expression.operator
            return op.operate(left, right)
        elif isinstance(expression, UnOp):
            raise NotImplementedError("Don't use unary operators, you stupid s**t")