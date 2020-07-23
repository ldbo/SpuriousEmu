"""Provide the Interpreter class, that is used to execute VBA programs."""

from typing import List

from .function import Function, InternalFunction, ExternalFunction
from .abstract_syntax_tree import *
from .memory import Memory
from .operator import OPERATORS_MAP
from .symbol import Symbol
from .value import Value


class Interpreter:
    """Class used to interprete a program which has already been compiled."""
    _symbols: Symbol
    _current_symbol: Symbol
    _memory: Memory

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
            value = self.evaluate_function_call(expression)
            if value is None:
                raise RuntimeError("Procedure called inside an expression")
            else:
                return value
        elif isinstance(expression, BinOp):
            left = self.evaluate_expression(expression.left)
            right = self.evaluate_expression(expression.right)
            op = expression.operator
            return op.operate(left, right)
        elif isinstance(expression, UnOp):
            raise NotImplementedError(
                "Don't use unary operators, you stupid s**t")
        else:
            raise NotImplemented("Don't know what you've thrown at me dude")

    def evaluate_function_call(self, call: FunCall) -> Optional[Value]:
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

            return_value: Optional[Value]
            try:
                return_value = self._memory.get_variable(function.name)
            except KeyError:
                return_value = None

            return return_value
        else:
            msg = f"{type(function)} is not handled yet by the interpreter"
            raise NotImplemented(msg)

    def execute_for(self, loop: For) -> None:
        # Evaluation of parameters
        # TODO check if Let-coercible to Double, raise E13 if not
        start_value = self.evaluate_expression(loop.start)
        end_value = self.evaluate_expression(loop.end)
        if loop.step is None:
            step_value = Integer(1)
        else:
            step_value = self.evaluate_expression(loop.step)
        step_literal = Literal.from_value(step_value)
        counter_name = loop.counter.name

        # Comparison function
        if step_value.value >= 0:
            def keep_going():
                counter_value = self._memory.get_variable(counter_name)
                return counter_value.value <= end_value.value
        else:
            def keep_going():
                counter_value = self._memory.get_variable(counter_name)
                return counter_value.value > end_value.value

        # Loop
        self._memory.set_variable(counter_name, start_value)
        while keep_going():
            self.execute_block(loop.body)
            add_expression = BinOp(OPERATORS_MAP['+'],
                                   loop.counter, step_literal)
            new_counter_value = self.evaluate_expression(add_expression)
            self._memory.set_variable(counter_name, new_counter_value)

    def interprete_statement(self, statement: AST) -> None:
        t = type(statement)
        if t is VarAssign:
            value = self.evaluate_expression(statement.value)
            self._memory.set_variable(statement.variable.name, value)
        elif t is FunCall:
            self.evaluate_function_call(statement)
        elif t is For:
            self.execute_for(statement)

    def run(self, function_name: str, args: Optional[List[Value]] = None) \
            -> None:
        """
        Search for a function entry point in the program symbol, and execute
        it. You should use it most of the time.
        """
        if args is None:
            args = []

        functions = self._symbols.find(function_name)
        for function in functions:
            print(f"Executing {function.full_name()}...")
            print("-------")
            self._current_symbol = function
            self.call_function(self._memory.get_function(function.full_name()),
                               args)
            print("-------")
            print()
