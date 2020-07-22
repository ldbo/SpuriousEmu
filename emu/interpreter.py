from typing import List, Dict, Any

from .abstract_syntax_tree import *
from .function import Function
from .memory import Memory
from .symbol import Symbol
from .value import Value


class Interpreter:
    _local_vars: List[Dict[str, Any]]
    _symbols: Symbol
    _memory: Dict[str, Function]

    def __init__(self, symbols=None, memory=None):
        self._local_vars = []
        self._symbols = symbols if symbols is not None else Symbol.build_root()
        self._memory = memory if memory is not None else Memory()

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
            raise NotImplementedError(
                "Don't use unary operators, you stupid s**t")

    def call_procedure(self, procedure: Function) -> None:
        self._local_vars.append(dict())
        for statement in procedure.body:
            self.interprete_statement(statement)

    def call_function():
        # TODODO
        pass

    def interprete_statement(self, statement: AST) -> None:
        t = type(statement)
        if t is VarAssign:
            value = self.evaluate_expression(statement.value)
            self._local_vars[-1][statement.variable.name] = value
        elif t is FunCall:
            short_name = statement.function.name
            function_name = self._symbols.resolve(short_name).full_name()
            if function_name is None:
                raise RuntimeError(f"Can't resolve symbol {short_name}")

            function = self._memory.get_function(function_name)
            arg_list = statement.arguments.args
            print(f"Calling {function_name} with arguments {arg_list}")
            function.call(self, list(map(self.evaluate_expression, arg_list)))
