from typing import List, Dict

from .function import Function, InternalFunction, ExternalFunction
from .abstract_syntax_tree import *
from .memory import Memory
from .symbol import Symbol
from .value import Value


class Interpreter:
    _symbols: Symbol
    _current_symbol: Symbol
    _memory: Dict[str, Function]

    def __init__(self, symbols=None, memory=None):
        self._symbols = symbols if symbols is not None else Symbol.build_root()
        self._current_symbol = self._symbols
        self._memory = memory if memory is not None else Memory()

    def evaluate_expression(self, expression: Expression) -> Value:
        if isinstance(expression, Literal):
            return Value.from_literal(expression)
        elif isinstance(expression, Identifier):
            return self._memory.get_variable(expression.name)
        elif isinstance(expression, FunCall):
            return self.evaluate_function_call(expression)
        elif isinstance(expression, BinOp):
            left = self.evaluate_expression(expression.left)
            right = self.evaluate_expression(expression.right)
            op = expression.operator
            return op.operate(left, right)
        elif isinstance(expression, UnOp):
            raise NotImplementedError(
                "Don't use unary operators, you stupid s**t")

    def evaluate_function_call(self, call: FunCall) -> Value:
        # Resolution
        function_symbol = self._current_symbol.resolve(call.function.name)
        function_name = function_symbol.full_name()
        function_object = self._memory.get_function(function_name)

        # Arguments handling
        arg_list = call.arguments.args
        arg_values = list(map(self.evaluate_expression, arg_list))

        # Move to the function
        previous_symbol = self._current_symbol
        self._current_symbol = function_symbol
        return_value = self.call_function(function_object, arg_values)
        self._current_symbol = previous_symbol

        return return_value

    def execute_block(self, block: List[Statement]) -> None:
        for statement in block:
            self.interprete_statement(statement)

    def call_function(self, function: Function,
                      arguments_values: List[Value]) -> Optional[Value]:
        if type(function) is ExternalFunction:
            return function.external_function(self, arguments_values)
        elif type(function) is InternalFunction:
            self._memory.new_locals()
            for name, value in zip(function.arguments_names, arguments_values):
                self._memory.set_variable(name, value)

            self.execute_block(function.body)

            try:
                return_value = self._memory.get_variable(function.name)
            except KeyError:
                return_value = None

            return return_value

    def interprete_statement(self, statement: AST) -> None:
        t = type(statement)
        if t is VarAssign:
            value = self.evaluate_expression(statement.value)
            self._memory.set_variable(statement.variable.name, value)
        elif t is FunCall:
            self.evaluate_function_call(statement)

    def run(self, function_name: str, args: Optional[List[Value]] = None) \
            -> Value:
        if args is None:
            args = []

        functions = self._symbols.find(function_name)
        for function in functions:
            self._current_symbol = function
            self.call_function(self._memory.get_function(function.full_name()),
                               args)
