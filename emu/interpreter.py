"""Provide the Interpreter class, that is used to execute VBA programs."""

from typing import List, Optional

from .function import Function, InternalFunction, ExternalFunction
from .abstract_syntax_tree import *
from .compiler import compile_files
from .side_effect import Memory, OutsideWorld
from .operator import OPERATORS_MAP
from .symbol import Symbol
from .value import Value, Integer
from .visitor import Visitor


class Interpreter(Visitor):
    """Class used to interprete a program which has already been compiled."""
    _symbols: Symbol
    _current_symbol: Symbol
    _memory: Memory
    _outside_world: OutsideWorld
    _evaluation: Optional[Value]

    # Main API
    def __init__(self, symbols=None, memory=None):
        self._symbols = symbols if symbols is not None else Symbol.build_root()
        self._current_symbol = self._symbols
        self._memory = memory if memory is not None else Memory()
        self._outside_world = OutsideWorld()
        self._evaluation = None

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
            self._current_symbol = function
            self.call_function(self._memory.get_function(function.full_name()),
                               args)

    def print_stdout(self, state: bool) -> None:
        """Enable or disable stdout forwarding"""
        if state:
            hook = print
        else:
            def hook(*args, **kwargs):
                pass

        self._outside_world.add_hook(OutsideWorld.EventType.STDOUT, hook)

    # Internal interface

    def evaluate(self, expression: Expression) -> Value:
        """Return the value of an expression"""
        assert(isinstance(expression, Expression))

        self.visit(expression)
        value = self._evaluation
        self._evaluation = None
        return value

    def visit_Block(self, block: Block) -> None:
        for statement in block.body:
            self.visit(statement)

    def visit_VarDec(self, var_dec: VarDec) -> None:
        # TODO
        pass

    def visit_VarAssign(self, var_assign: VarAssign) -> None:
        value = self.evaluate(var_assign.value)
        self._memory.set_variable(var_assign.variable.name, value)

    def visit_Literal(self, literal: Literal) -> None:
        self._evaluation = Value.from_literal(literal)

    def visit_Identifier(self, identifier: Identifier) -> None:
        self._evaluation = self._memory.get_variable(identifier.name)

    def visit_BinOp(self, bin_op: BinOp) -> None:
        left_value = self.evaluate(bin_op.left)
        right_value = self.evaluate(bin_op.right)

        op = bin_op.operator
        self._evaluation = op.operate(left_value, right_value)

    def visit_FunCall(self, fun_call: FunCall) -> None:
        # Resolution
        function_symbol = self._current_symbol.resolve(fun_call.function.name)
        function_name = function_symbol.full_name()
        function_object = self._memory.get_function(function_name)

        # Arguments handling
        arg_list = fun_call.arguments.args
        arg_values = list(map(self.evaluate, arg_list))

        # Move to the function
        previous_symbol = self._current_symbol
        self._current_symbol = function_symbol
        return_value = self.call_function(function_object, arg_values)
        self._current_symbol = previous_symbol

        self._evaluation = return_value

    def visit_If(self, conditional: If) -> None:
        done = False

        # If
        condition_value = self.evaluate(conditional.condition)
        if condition_value.convert_to(Type.Boolean).value:
            self.visit_Block(conditional)
            done = True

        # Else ifs
        if not done:
            for elseif in conditional.elsifs:
                condition_value = elseif.evaluate(elseif.condition)
                if condition_value.convert_to(Type.Boolean).value:
                    self.visit_Block(elseif)
                    done = True

        # Else
        if not done:
            self.visit_Block(conditional.else_block)

    def visit_For(self, loop: For) -> None:
        # Evaluation of parameters
        # TODO check if Let-coercible to Double, raise E13 if not
        start_value = self.evaluate(loop.start)
        end_value = self.evaluate(loop.end)
        if loop.step is None:
            step_value = Integer(1)
        else:
            step_value = self.evaluate(loop.step)
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
            self.visit_Block(loop)
            add_expression = BinOp(OPERATORS_MAP['+'],
                                   loop.counter, step_literal)
            new_counter_value = self.evaluate(add_expression)
            self._memory.set_variable(counter_name, new_counter_value)

    def visit_OnError(self, on_error: OnError) -> None:
        # TODO
        pass

    def visit_Resume(self, resume: Resume) -> None:
        # TODO
        pass

    def visit_ErrorStatement(self, error: ErrorStatement) -> None:
        # TODO
        pass

    def call_function(self, function: Function,
                      arguments_values: List[Value]) -> Optional[Value]:
        if type(function) is ExternalFunction:
            return function.external_function(self, arguments_values)
        elif type(function) is InternalFunction:
            self._memory.new_locals()
            for name, value in zip(function.arguments_names, arguments_values):
                self._memory.set_variable(name, value)

            self.visit_Block(function)

            return_value: Optional[Value]
            try:
                return_value = self._memory.get_variable(function.name)
            except KeyError:
                return_value = None

            self._memory.discard_locals()

            return return_value
        else:
            msg = f"{type(function)} is not handled yet by the interpreter"
            raise NotImplemented(msg)

    # External functions interface
    def add_stdout(self, content: str) -> None:
        self._outside_world.add_event(OutsideWorld.EventType.STDOUT, content)

    def add_file_event(self, *args, **kwargs) -> None:
        # TODO
        pass

    def add_network_event(self, *args, **kwargs) -> None:
        # TODO
        pass


def run_program(file_paths: List[str]) -> None:
    program = compile_files(file_paths)
    interpreter = Interpreter(program.symbols, program.memory)
    interpreter.run('Main')


def run_single_file(file_path: str) -> None:
    run_program([file_path])
